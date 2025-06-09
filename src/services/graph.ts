import { DeviceCodeCredential, useIdentityPlugin } from "@azure/identity";
import { cachePersistencePlugin } from "@azure/identity-cache-persistence";
import { Client } from "@microsoft/microsoft-graph-client";

// Ensure plugin is loaded only in non-test environments
if (process.env.NODE_ENV !== "test" && !process.env.VITEST) {
  useIdentityPlugin(cachePersistencePlugin);
}

export interface AuthStatus {
  isAuthenticated: boolean;
  userPrincipalName?: string | undefined;
  displayName?: string | undefined;
  expiresAt?: string | undefined;
}

const CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const SCOPES = [
  "User.Read",
  "User.ReadBasic.All",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "ChannelMessage.Send",
  "TeamMember.Read.All",
  "Chat.ReadBasic",
  "Chat.ReadWrite",
];

export class GraphService {
  private static instance: GraphService;
  private client: Client | undefined;
  private credential: DeviceCodeCredential | undefined;
  private silentCredential: DeviceCodeCredential | undefined;
  private isInitialized = false;

  static getInstance(): GraphService {
    if (!GraphService.instance) {
      GraphService.instance = new GraphService();
    }
    return GraphService.instance;
  }

  private async initializeClient(): Promise<void> {
    if (this.isInitialized) return;

    try {
      // Regular credential for normal operations (allows interactive auth)
      this.credential = new DeviceCodeCredential({
        clientId: CLIENT_ID,
        tenantId: "common",
        tokenCachePersistenceOptions: {
          enabled: true,
        },
      });

      // Silent credential for status checks (no interactive auth)
      this.silentCredential = new DeviceCodeCredential({
        clientId: CLIENT_ID,
        tenantId: "common",
        tokenCachePersistenceOptions: {
          enabled: true,
        },
        disableAutomaticAuthentication: true,
      });

      // Verify authentication by attempting to get a token silently
      const authToken = await this.silentCredential.getToken(["User.Read"]);
      if (!authToken?.token) {
        console.error("No valid authentication found in cache. Please authenticate first.");
        return; // Don't mark as initialized if no valid auth
      }

      // Verify the token works by making a test call to Graph API
      const testResponse = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: {
          Authorization: `Bearer ${authToken.token}`,
          "Content-Type": "application/json",
        },
      });

      if (!testResponse.ok) {
        console.error(`Authentication validation failed: HTTP ${testResponse.status}`);
        return; // Don't mark as initialized if auth validation fails
      }

      // Create Graph client with Azure Identity credential
      this.client = Client.initWithMiddleware({
        authProvider: {
          getAccessToken: async () => {
            if (!this.credential) {
              throw new Error("Credential not initialized");
            }
            const token = await this.credential.getToken(SCOPES);
            if (!token?.token) {
              throw new Error("Failed to obtain access token");
            }
            return token.token;
          },
        },
      });

      this.isInitialized = true;
      console.log("Graph service initialized successfully with valid authentication");
    } catch (error) {
      console.error("Failed to initialize Graph client:", error);
      // Reset credentials on failure
      this.credential = undefined;
      this.silentCredential = undefined;
      this.client = undefined;
    }
  }

  async getAuthStatus(): Promise<AuthStatus> {
    await this.initializeClient();

    // If initialization failed, we're not authenticated
    if (!this.isInitialized || !this.silentCredential) {
      return { isAuthenticated: false };
    }

    try {
      // Get token silently and fetch user info
      const token = await this.silentCredential.getToken(["User.Read"]);
      if (!token) {
        return { isAuthenticated: false };
      }

      // Use a direct fetch call with the silent token for this status check
      const me = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: {
          Authorization: `Bearer ${token.token}`,
          "Content-Type": "application/json",
        },
      }).then((response) => {
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        return response.json();
      });

      return {
        isAuthenticated: true,
        userPrincipalName: me?.userPrincipalName ?? undefined,
        displayName: me?.displayName ?? undefined,
        expiresAt: new Date(token.expiresOnTimestamp).toISOString(),
      };
    } catch (error) {
      console.error("Error getting user info:", error);
      return { isAuthenticated: false };
    }
  }

  async getClient(): Promise<Client> {
    await this.initializeClient();

    if (!this.client || !this.isInitialized) {
      throw new Error(
        "Authentication required. Please run: npx @floriscornel/teams-mcp@latest authenticate"
      );
    }
    return this.client;
  }

  isAuthenticated(): boolean {
    return !!this.client && this.isInitialized;
  }
}

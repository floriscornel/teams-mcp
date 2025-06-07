import { DeviceCodeCredential, useIdentityPlugin } from "@azure/identity";
import { cachePersistencePlugin } from "@azure/identity-cache-persistence";
import { Client } from "@microsoft/microsoft-graph-client";

// Ensure plugin is loaded
useIdentityPlugin(cachePersistencePlugin);

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
      this.credential = new DeviceCodeCredential({
        clientId: CLIENT_ID,
        tenantId: "common",
        tokenCachePersistenceOptions: {
          enabled: true,
        },
      });

      // Create Graph client with Azure Identity credential
      this.client = Client.initWithMiddleware({
        authProvider: {
          getAccessToken: async () => {
            if (!this.credential) {
              throw new Error("Credential not initialized");
            }
            const token = await this.credential.getToken(SCOPES);
            return token?.token || "";
          },
        },
      });

      this.isInitialized = true;
    } catch (error) {
      console.error("Failed to initialize Graph client:", error);
    }
  }

  async getAuthStatus(): Promise<AuthStatus> {
    await this.initializeClient();

    if (!this.client || !this.credential) {
      return { isAuthenticated: false };
    }

    try {
      // Try to get a token silently first
      const token = await this.credential.getToken(["User.Read"]);
      if (!token) {
        return { isAuthenticated: false };
      }

      const me = await this.client.api("/me").get();
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

    if (!this.client) {
      throw new Error(
        "Not authenticated. Please run the authentication CLI tool first: npx @floriscornel/teams-mcp@latest authenticate"
      );
    }
    return this.client;
  }

  isAuthenticated(): boolean {
    return !!this.client && this.isInitialized;
  }
}

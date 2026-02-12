import { type AccountInfo, PublicClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { cachePlugin } from "../msal-cache.js";

const CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const AUTHORITY = "https://login.microsoftonline.com/common";

// Read-only mode detection from environment variable
function isReadOnlyMode(): boolean {
  const value = process.env.TEAMS_MCP_READ_ONLY?.toLowerCase()?.trim();
  return value === "true" || value === "1" || value === "yes";
}

const FULL_ACCESS_SCOPES = [
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

const READ_ONLY_SCOPES = [
  "User.Read",
  "User.ReadBasic.All",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "TeamMember.Read.All",
  "Chat.ReadBasic",
  "Chat.Read",
];

function getDelegatedScopes(readOnly: boolean): string[] {
  return readOnly ? READ_ONLY_SCOPES : FULL_ACCESS_SCOPES;
}

export interface AuthStatus {
  isAuthenticated: boolean;
  userPrincipalName?: string | undefined;
  displayName?: string | undefined;
  expiresAt?: string | undefined;
}

export class GraphService {
  private static instance: GraphService;
  private client: Client | undefined;
  private isInitialized = false;
  private tokenExpiresAt: Date | undefined;
  private msalApp: PublicClientApplication | undefined;
  private msalAccount: AccountInfo | undefined;

  static getInstance(): GraphService {
    if (!GraphService.instance) {
      GraphService.instance = new GraphService();
    }
    return GraphService.instance;
  }

  private async initializeClient(): Promise<void> {
    if (this.isInitialized) return;

    try {
      // Priority 1: AUTH_TOKEN environment variable (direct token injection)
      const envToken = process.env.AUTH_TOKEN;
      if (envToken) {
        const validatedToken = this.validateToken(envToken);
        if (validatedToken) {
          this.client = Client.initWithMiddleware({
            authProvider: {
              getAccessToken: async () => validatedToken,
            },
          });
          this.isInitialized = true;
        }
        return;
      }

      // Priority 2: MSAL with cached refresh token for automatic token renewal
      this.msalApp = new PublicClientApplication({
        auth: {
          clientId: CLIENT_ID,
          authority: AUTHORITY,
        },
        cache: {
          cachePlugin,
        },
      });

      const accounts = await this.msalApp.getTokenCache().getAllAccounts();
      if (accounts.length === 0) {
        return;
      }

      this.msalAccount = accounts[0];

      // Verify we can acquire a token
      const result = await this.msalApp.acquireTokenSilent({
        scopes: getDelegatedScopes(isReadOnlyMode()),
        account: this.msalAccount,
      });

      if (!result) {
        return;
      }

      this.tokenExpiresAt = result.expiresOn ?? undefined;

      // Create Graph client with MSAL-backed auth provider for automatic token refresh
      this.client = Client.initWithMiddleware({
        authProvider: {
          getAccessToken: () => this.acquireToken(),
        },
      });

      this.isInitialized = true;
    } catch (error) {
      console.error("Failed to initialize Graph client:", error);
    }
  }

  private async acquireToken(): Promise<string> {
    if (!this.msalApp || !this.msalAccount) {
      throw new Error("MSAL not initialized");
    }

    const result = await this.msalApp.acquireTokenSilent({
      scopes: getDelegatedScopes(isReadOnlyMode()),
      account: this.msalAccount,
    });

    if (!result) {
      throw new Error(
        "Failed to acquire access token. Please re-authenticate: npx @floriscornel/teams-mcp@latest authenticate"
      );
    }

    this.tokenExpiresAt = result.expiresOn ?? undefined;
    return result.accessToken;
  }

  async getAuthStatus(): Promise<AuthStatus> {
    await this.initializeClient();

    if (!this.client) {
      return { isAuthenticated: false };
    }

    try {
      const me = await this.client.api("/me").get();
      return {
        isAuthenticated: true,
        userPrincipalName: me?.userPrincipalName ?? undefined,
        displayName: me?.displayName ?? undefined,
        expiresAt: this.tokenExpiresAt?.toISOString(),
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

  validateToken(token: string): string | undefined {
    const tokenSplits = token.split(".");
    if (tokenSplits.length !== 3) {
      console.error("Invalid JWT token: missing claims");
      return undefined;
    }

    try {
      const payload = JSON.parse(atob(tokenSplits[1]));
      const audiences = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
      if (!audiences.includes("https://graph.microsoft.com")) {
        console.error("Invalid JWT token: Not a valid Microsoft Graph token");
        return undefined;
      }
    } catch (error) {
      console.error("Invalid JWT token: Failed to parse payload", error);
      return undefined;
    }

    return token;
  }
}

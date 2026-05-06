import { randomUUID } from "node:crypto";
import { ConfidentialClientApplication } from "@azure/msal-node";
import type { OAuthRegisteredClientsStore } from "@modelcontextprotocol/sdk/server/auth/clients.js";
import type {
  AuthorizationParams,
  OAuthServerProvider,
} from "@modelcontextprotocol/sdk/server/auth/provider.js";
import type { AuthInfo } from "@modelcontextprotocol/sdk/server/auth/types.js";
import type {
  OAuthClientInformationFull,
  OAuthTokenRevocationRequest,
  OAuthTokens,
} from "@modelcontextprotocol/sdk/shared/auth.js";
import type { Request, Response } from "express";
import { FULL_SCOPES, READ_ONLY_SCOPES } from "../services/graph.js";
import { InMemoryClientsStore } from "./clients-store.js";
import { SessionStore } from "./session-store.js";

interface PendingAuthorization {
  mcpCodeChallenge: string;
  mcpRedirectUri: string;
  mcpClientId: string;
  mcpState?: string | undefined;
  entraState: string;
  createdAt: number;
}

const PENDING_AUTH_TTL_MS = 10 * 60 * 1000; // 10 minutes
const PENDING_AUTH_CLEANUP_INTERVAL_MS = 60 * 1000; // 1 minute

export interface EntraOAuthProviderOptions {
  clientId: string;
  clientSecret: string;
  authority: string;
  baseUrl: string;
  readOnly: boolean;
}

/**
 * OAuth server provider that federates MCP OAuth to Azure AD (Entra ID).
 *
 * Flow:
 * 1. MCP client calls /authorize -> we redirect to Entra's /authorize
 * 2. Entra calls back to /oauth/callback -> we exchange the code for Graph tokens
 * 3. We generate an MCP auth code and redirect back to the MCP client
 * 4. MCP client exchanges the MCP auth code for an opaque MCP session token
 * 5. MCP requests carry the session token; we look up the Graph token server-side
 */
export class EntraOAuthProvider implements OAuthServerProvider {
  private _clientsStore = new InMemoryClientsStore();
  private _sessionStore = new SessionStore();
  private _msalApp: ConfidentialClientApplication;
  private _pendingAuthorizations = new Map<string, PendingAuthorization>();
  private _options: EntraOAuthProviderOptions;

  private _pendingAuthCleanupTimer: ReturnType<typeof setInterval>;

  constructor(options: EntraOAuthProviderOptions) {
    this._options = options;
    this._msalApp = new ConfidentialClientApplication({
      auth: {
        clientId: options.clientId,
        clientSecret: options.clientSecret,
        authority: options.authority,
      },
    });
    this._pendingAuthCleanupTimer = setInterval(
      () => this.cleanupPendingAuthorizations(),
      PENDING_AUTH_CLEANUP_INTERVAL_MS
    );
    this._pendingAuthCleanupTimer.unref();
  }

  private cleanupPendingAuthorizations(): void {
    const now = Date.now();
    for (const [key, pending] of this._pendingAuthorizations) {
      if (now - pending.createdAt > PENDING_AUTH_TTL_MS) {
        this._pendingAuthorizations.delete(key);
      }
    }
  }

  get clientsStore(): OAuthRegisteredClientsStore {
    return this._clientsStore;
  }

  get sessionStore(): SessionStore {
    return this._sessionStore;
  }

  private get graphScopes(): string[] {
    return this._options.readOnly ? READ_ONLY_SCOPES : FULL_SCOPES;
  }

  /**
   * MCP client initiates authorization. We save the MCP PKCE challenge/redirect
   * and redirect the user to Entra's authorize endpoint.
   */
  async authorize(
    client: OAuthClientInformationFull,
    params: AuthorizationParams,
    res: Response
  ): Promise<void> {
    const entraState = randomUUID();

    this._pendingAuthorizations.set(entraState, {
      mcpCodeChallenge: params.codeChallenge,
      mcpRedirectUri: params.redirectUri,
      mcpClientId: client.client_id,
      mcpState: params.state,
      entraState,
      createdAt: Date.now(),
    });

    const redirectUri = `${this._options.baseUrl}/oauth/callback`;

    const authCodeUrl = await this._msalApp.getAuthCodeUrl({
      scopes: this.graphScopes,
      redirectUri,
      state: entraState,
      prompt: "select_account",
    });

    res.redirect(authCodeUrl);
  }

  /**
   * Handle the callback from Entra ID after the user authenticates.
   * Exchanges the Entra auth code for Graph tokens, generates an MCP auth code,
   * and redirects back to the MCP client.
   */
  async handleEntraCallback(req: Request, res: Response): Promise<void> {
    const { code, state, error, error_description } = req.query as Record<string, string>;

    if (error) {
      console.error("Entra OAuth error:", error, error_description);
      res.status(400).json({ error, error_description });
      return;
    }

    if (!state || !code) {
      res.status(400).json({ error: "missing_params", error_description: "Missing code or state" });
      return;
    }

    const pending = this._pendingAuthorizations.get(state);
    if (!pending || Date.now() - pending.createdAt > PENDING_AUTH_TTL_MS) {
      if (pending) this._pendingAuthorizations.delete(state);
      res
        .status(400)
        .json({ error: "invalid_state", error_description: "Unknown or expired state" });
      return;
    }
    this._pendingAuthorizations.delete(state);

    try {
      const redirectUri = `${this._options.baseUrl}/oauth/callback`;
      const result = await this._msalApp.acquireTokenByCode({
        code,
        scopes: this.graphScopes,
        redirectUri,
      });

      if (!result) {
        res.status(500).json({ error: "token_exchange_failed" });
        return;
      }

      const mcpAuthCode = randomUUID();
      const graphRefreshToken = this.extractRefreshToken(result.account?.homeAccountId ?? "");

      this._sessionStore.storeAuthCode({
        mcpAuthCode,
        codeChallenge: pending.mcpCodeChallenge,
        redirectUri: pending.mcpRedirectUri,
        mcpClientId: pending.mcpClientId,
        graphAccessToken: result.accessToken,
        graphRefreshToken,
        entraAccountId: result.account?.homeAccountId ?? "",
        state: pending.mcpState,
        createdAt: Date.now(),
      });

      const redirectUrl = new URL(pending.mcpRedirectUri);
      redirectUrl.searchParams.set("code", mcpAuthCode);
      if (pending.mcpState) {
        redirectUrl.searchParams.set("state", pending.mcpState);
      }

      res.redirect(redirectUrl.toString());
    } catch (err) {
      console.error("Failed to exchange Entra code:", err);
      res.status(500).json({ error: "token_exchange_failed" });
    }
  }

  /**
   * Best-effort extraction of the refresh token from MSAL's in-memory cache.
   * MSAL doesn't expose the refresh token directly on the AuthenticationResult,
   * but stores it internally. For the confidential client flow we rely on MSAL's
   * internal cache for silent token renewal.
   */
  private extractRefreshToken(homeAccountId: string): string {
    try {
      const cache = this._msalApp.getTokenCache().serialize();
      const parsed = JSON.parse(cache);
      const refreshTokens = parsed?.RefreshToken ?? {};
      for (const entry of Object.values<{ home_account_id?: string; secret?: string }>(
        refreshTokens
      )) {
        if (entry?.home_account_id === homeAccountId) {
          return entry.secret ?? "";
        }
      }
    } catch {
      // Fall through
    }
    return "";
  }

  /** Returns the PKCE code challenge associated with the given authorization code. */
  async challengeForAuthorizationCode(
    _client: OAuthClientInformationFull,
    authorizationCode: string
  ): Promise<string> {
    const pending = this._sessionStore.getAuthCode(authorizationCode);
    if (!pending) {
      throw new Error("Unknown or expired authorization code");
    }
    return pending.codeChallenge;
  }

  /** Exchanges an MCP authorization code for session tokens backed by the stored Graph token. */
  async exchangeAuthorizationCode(
    _client: OAuthClientInformationFull,
    authorizationCode: string,
    _codeVerifier?: string,
    _redirectUri?: string,
    _resource?: URL
  ): Promise<OAuthTokens> {
    const pending = this._sessionStore.consumeAuthCode(authorizationCode);
    if (!pending) {
      throw new Error("Unknown, expired, or already-used authorization code");
    }

    const session = this._sessionStore.createSession({
      graphAccessToken: pending.graphAccessToken,
      graphRefreshToken: pending.graphRefreshToken,
      entraAccountId: pending.entraAccountId,
      clientId: pending.mcpClientId,
      scopes: this.graphScopes,
      expiresInSeconds: 3600,
    });

    return {
      access_token: session.mcpAccessToken,
      token_type: "Bearer",
      expires_in: 3600,
      refresh_token: session.mcpRefreshToken,
    };
  }

  /** Rotates an MCP session by silently refreshing the underlying Graph token via MSAL. */
  async exchangeRefreshToken(
    _client: OAuthClientInformationFull,
    refreshToken: string,
    _scopes?: string[],
    _resource?: URL
  ): Promise<OAuthTokens> {
    const oldSession = this._sessionStore.getSessionByRefreshToken(refreshToken);
    if (!oldSession) {
      throw new Error("Invalid or expired refresh token");
    }

    let newGraphAccessToken: string | undefined;

    if (oldSession.entraAccountId) {
      try {
        const accounts = await this._msalApp.getTokenCache().getAllAccounts();
        const account = accounts.find((a) => a.homeAccountId === oldSession.entraAccountId);
        if (account) {
          const result = await this._msalApp.acquireTokenSilent({
            scopes: this.graphScopes,
            account,
          });
          newGraphAccessToken = result?.accessToken;
        }
      } catch (err) {
        console.error("MSAL silent refresh failed:", err);
      }
    }

    if (!newGraphAccessToken) {
      this._sessionStore.deleteSession(oldSession.mcpAccessToken);
      throw new Error("Graph token refresh failed; client must re-authenticate");
    }

    this._sessionStore.deleteSession(oldSession.mcpAccessToken);

    const newSession = this._sessionStore.createSession({
      graphAccessToken: newGraphAccessToken,
      graphRefreshToken: oldSession.graphRefreshToken,
      entraAccountId: oldSession.entraAccountId,
      clientId: oldSession.clientId,
      scopes: this.graphScopes,
      expiresInSeconds: 3600,
    });

    return {
      access_token: newSession.mcpAccessToken,
      token_type: "Bearer",
      expires_in: 3600,
      refresh_token: newSession.mcpRefreshToken,
    };
  }

  /** Validates an MCP access token and returns the associated auth info including the Graph token. */
  async verifyAccessToken(token: string): Promise<AuthInfo> {
    const session = this._sessionStore.getSession(token);
    if (!session) {
      throw new Error("Invalid or expired access token");
    }

    return {
      token: session.mcpAccessToken,
      clientId: session.clientId,
      scopes: session.scopes,
      expiresAt: Math.floor(session.expiresAt / 1000),
      extra: {
        graphToken: session.graphAccessToken,
      },
    };
  }

  /** Revokes an MCP access or refresh token by deleting the associated session. */
  async revokeToken(
    _client: OAuthClientInformationFull,
    request: OAuthTokenRevocationRequest
  ): Promise<void> {
    // Try as access token first
    const session = this._sessionStore.getSession(request.token);
    if (session) {
      this._sessionStore.deleteSession(request.token);
      return;
    }

    // Try as refresh token
    const sessionByRefresh = this._sessionStore.getSessionByRefreshToken(request.token);
    if (sessionByRefresh) {
      this._sessionStore.deleteSession(sessionByRefresh.mcpAccessToken);
    }
  }
}

import { randomUUID } from "node:crypto";

export interface Session {
  mcpAccessToken: string;
  mcpRefreshToken: string;
  graphAccessToken: string;
  graphRefreshToken: string;
  entraAccountId: string;
  clientId: string;
  scopes: string[];
  createdAt: number;
  expiresAt: number;
}

export interface PendingAuthCode {
  mcpAuthCode: string;
  codeChallenge: string;
  redirectUri: string;
  mcpClientId: string;
  graphAccessToken: string;
  graphRefreshToken: string;
  entraAccountId: string;
  state?: string | undefined;
  createdAt: number;
}

const AUTH_CODE_TTL_MS = 5 * 60 * 1000; // 5 minutes
const SESSION_TTL_MS = 60 * 60 * 1000; // 1 hour
const CLEANUP_INTERVAL_MS = 60 * 1000; // run cleanup every minute

export class SessionStore {
  private sessions = new Map<string, Session>();
  private pendingCodes = new Map<string, PendingAuthCode>();
  private cleanupTimer: ReturnType<typeof setInterval>;

  constructor() {
    this.cleanupTimer = setInterval(() => this.cleanup(), CLEANUP_INTERVAL_MS);
    this.cleanupTimer.unref();
  }

  createSession(params: {
    graphAccessToken: string;
    graphRefreshToken: string;
    entraAccountId: string;
    clientId: string;
    scopes: string[];
    expiresInSeconds?: number;
  }): Session {
    const session: Session = {
      mcpAccessToken: randomUUID(),
      mcpRefreshToken: randomUUID(),
      graphAccessToken: params.graphAccessToken,
      graphRefreshToken: params.graphRefreshToken,
      entraAccountId: params.entraAccountId,
      clientId: params.clientId,
      scopes: params.scopes,
      createdAt: Date.now(),
      expiresAt: Date.now() + (params.expiresInSeconds ?? SESSION_TTL_MS / 1000) * 1000,
    };
    this.sessions.set(session.mcpAccessToken, session);
    return session;
  }

  getSession(mcpAccessToken: string): Session | undefined {
    const session = this.sessions.get(mcpAccessToken);
    if (!session) return undefined;
    if (Date.now() > session.expiresAt) {
      this.sessions.delete(mcpAccessToken);
      return undefined;
    }
    return session;
  }

  getSessionByRefreshToken(mcpRefreshToken: string): Session | undefined {
    for (const session of this.sessions.values()) {
      if (session.mcpRefreshToken === mcpRefreshToken) {
        if (Date.now() > session.expiresAt) {
          this.sessions.delete(session.mcpAccessToken);
          return undefined;
        }
        return session;
      }
    }
    return undefined;
  }

  deleteSession(mcpAccessToken: string): void {
    this.sessions.delete(mcpAccessToken);
  }

  storeAuthCode(pending: PendingAuthCode): void {
    this.pendingCodes.set(pending.mcpAuthCode, pending);
  }

  consumeAuthCode(mcpAuthCode: string): PendingAuthCode | undefined {
    const pending = this.pendingCodes.get(mcpAuthCode);
    if (!pending) return undefined;
    this.pendingCodes.delete(mcpAuthCode);
    if (Date.now() - pending.createdAt > AUTH_CODE_TTL_MS) {
      return undefined;
    }
    return pending;
  }

  getAuthCode(mcpAuthCode: string): PendingAuthCode | undefined {
    const pending = this.pendingCodes.get(mcpAuthCode);
    if (!pending) return undefined;
    if (Date.now() - pending.createdAt > AUTH_CODE_TTL_MS) {
      this.pendingCodes.delete(mcpAuthCode);
      return undefined;
    }
    return pending;
  }

  private cleanup(): void {
    const now = Date.now();
    for (const [key, session] of this.sessions) {
      if (now > session.expiresAt) {
        this.sessions.delete(key);
      }
    }
    for (const [key, pending] of this.pendingCodes) {
      if (now - pending.createdAt > AUTH_CODE_TTL_MS) {
        this.pendingCodes.delete(key);
      }
    }
  }

  destroy(): void {
    clearInterval(this.cleanupTimer);
  }
}

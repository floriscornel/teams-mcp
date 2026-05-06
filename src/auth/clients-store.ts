import { randomUUID } from "node:crypto";
import type { OAuthRegisteredClientsStore } from "@modelcontextprotocol/sdk/server/auth/clients.js";
import type { OAuthClientInformationFull } from "@modelcontextprotocol/sdk/shared/auth.js";

/**
 * In-memory implementation of OAuthRegisteredClientsStore for RFC 7591
 * dynamic client registration. MCP hosts (Cursor, Claude, etc.) self-register
 * when they first connect.
 */
export class InMemoryClientsStore implements OAuthRegisteredClientsStore {
  private clients = new Map<string, OAuthClientInformationFull>();

  /** Retrieves a previously registered client by its ID, or undefined if not found. */
  getClient(clientId: string): OAuthClientInformationFull | undefined {
    return this.clients.get(clientId);
  }

  /** Registers a new OAuth client, assigning it a unique ID and issued-at timestamp. */
  registerClient(
    client: Omit<OAuthClientInformationFull, "client_id" | "client_id_issued_at">
  ): OAuthClientInformationFull {
    const fullClient: OAuthClientInformationFull = {
      ...client,
      client_id: randomUUID(),
      client_id_issued_at: Math.floor(Date.now() / 1000),
    };
    this.clients.set(fullClient.client_id, fullClient);
    return fullClient;
  }
}

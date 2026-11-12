import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { GraphService } from "./services/graph.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerChatTools } from "./tools/chats.js";
import { registerSearchTools } from "./tools/search.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerUsersTools } from "./tools/users.js";

/** Path where the CLI persists auth metadata. */
export const AUTH_INFO_PATH = join(homedir(), ".msgraph-mcp-auth.json");

/** Read the persisted auth info file (best-effort). */
async function readAuthInfo(): Promise<Record<string, unknown> | undefined> {
  try {
    const data = await fs.readFile(AUTH_INFO_PATH, "utf8");
    return JSON.parse(data) as Record<string, unknown>;
  } catch {
    return undefined;
  }
}

/**
 * Create the MCP server with all tools registered.
 * The returned server is not connected to any transport yet.
 */
export async function createMcpServer(readOnly: boolean): Promise<McpServer> {
  const server = new McpServer({
    name: "teams-mcp",
    version: "1.0.0",
  });

  // Initialize Graph service (singleton)
  const graphService = GraphService.getInstance();
  graphService.readOnlyMode = readOnly;

  // Detect scope mismatch: warn when switching from read-only → full mode
  if (!readOnly && !process.env.AUTH_TOKEN) {
    const authInfo = await readAuthInfo();
    if (authInfo) {
      const grantedScopes = authInfo.grantedScopes as string[] | undefined;
      const hasWriteScopes = grantedScopes?.some(
        (s: string) =>
          s === "ChannelMessage.Send" ||
          s === "ChannelMessage.ReadWrite" ||
          s === "Chat.ReadWrite" ||
          s === "Files.ReadWrite.All"
      );
      if (grantedScopes && !hasWriteScopes) {
        console.error(
          "⚠️  Warning: You authenticated with read-only scopes but the server is running in full mode."
        );
        console.error("   Write operations may fail. Re-authenticate without --read-only:");
        console.error("   npx @floriscornel/teams-mcp@latest authenticate");
      } else if (!grantedScopes) {
        console.error(
          "⚠️  Warning: Could not determine granted scopes. If you experience permission errors,"
        );
        console.error("   re-authenticate: npx @floriscornel/teams-mcp@latest authenticate");
      }
    }
  }

  // Register all tools (write tools are skipped when readOnly is true)
  registerAuthTools(server, graphService, readOnly);
  registerUsersTools(server, graphService, readOnly);
  registerTeamsTools(server, graphService, readOnly);
  registerChatTools(server, graphService, readOnly);
  registerSearchTools(server, graphService, readOnly);

  return server;
}

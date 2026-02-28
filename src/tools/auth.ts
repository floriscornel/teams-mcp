import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../services/graph.js";

export function registerAuthTools(
  server: McpServer,
  graphService: GraphService,
  _readOnly: boolean
) {
  // Authentication status tool
  server.tool(
    "auth_status",
    "Check the authentication status of the Microsoft Graph connection. Returns whether the user is authenticated and shows their basic profile information.",
    {},
    async () => {
      const status = await graphService.getAuthStatus();
      return {
        content: [
          {
            type: "text",
            text: status.isAuthenticated
              ? `✅ Authenticated as ${status.displayName || "Unknown User"} (${status.userPrincipalName || "No email available"})`
              : "❌ Not authenticated. Please run: npx @floriscornel/teams-mcp@latest authenticate",
          },
        ],
      };
    }
  );
}

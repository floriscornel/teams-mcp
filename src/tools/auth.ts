import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../services/graph.js";

export function registerAuthTools(server: McpServer, graphService: GraphService, readOnly = false) {
  // Authentication status tool
  server.tool(
    "auth_status",
    "Check the authentication status of the Microsoft Graph connection. Returns whether the user is authenticated and shows their basic profile information.",
    {},
    async () => {
      const status = await graphService.getAuthStatus();
      const modeIndicator = readOnly ? " [Read-Only Mode]" : "";
      return {
        content: [
          {
            type: "text",
            text: status.isAuthenticated
              ? `✅ Authenticated as ${status.displayName || "Unknown User"} (${status.userPrincipalName || "No email available"})${modeIndicator}`
              : "❌ Not authenticated. Please run: npx @floriscornel/teams-mcp@latest authenticate",
          },
        ],
      };
    }
  );
}

import { randomUUID } from "node:crypto";
import { requireBearerAuth } from "@modelcontextprotocol/sdk/server/auth/middleware/bearerAuth.js";
import { mcpAuthRouter } from "@modelcontextprotocol/sdk/server/auth/router.js";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import express from "express";
import { EntraOAuthProvider } from "./auth/provider.js";
import { GraphService } from "./services/graph.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerChatTools } from "./tools/chats.js";
import { registerSearchTools } from "./tools/search.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerUsersTools } from "./tools/users.js";

function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

function createSessionServer(graphToken: string, readOnly: boolean): McpServer {
  const server = new McpServer({
    name: "teams-mcp",
    version: "1.0.0",
  });

  const graphService = GraphService.createForToken(graphToken, readOnly);

  registerAuthTools(server, graphService, readOnly);
  registerUsersTools(server, graphService, readOnly);
  registerTeamsTools(server, graphService, readOnly);
  registerChatTools(server, graphService, readOnly);
  registerSearchTools(server, graphService, readOnly);

  return server;
}

export async function startHttpServer(readOnly: boolean): Promise<void> {
  const baseUrl = requireEnv("TEAMS_MCP_BASE_URL");
  const clientId = requireEnv("TEAMS_MCP_CLIENT_ID");
  const clientSecret = requireEnv("TEAMS_MCP_CLIENT_SECRET");
  const authority = requireEnv("TEAMS_MCP_AUTHORITY");
  const port = Number.parseInt(process.env.TEAMS_MCP_PORT || "3000", 10);

  const provider = new EntraOAuthProvider({
    clientId,
    clientSecret,
    authority,
    baseUrl,
    readOnly,
  });

  const app = express();

  // Mount MCP OAuth endpoints at root (/.well-known/*, /authorize, /token, /register)
  app.use(
    mcpAuthRouter({
      provider,
      issuerUrl: new URL(baseUrl),
      baseUrl: new URL(baseUrl),
      serviceDocumentationUrl: new URL("https://github.com/floriscornel/teams-mcp#readme"),
    })
  );

  // Entra ID callback route
  app.get("/oauth/callback", async (req, res) => {
    try {
      await provider.handleEntraCallback(req, res);
    } catch (err) {
      console.error("OAuth callback error:", err);
      res.status(500).json({ error: "callback_failed" });
    }
  });

  // Per-session transport map
  const sessions = new Map<string, StreamableHTTPServerTransport>();

  const bearerAuth = requireBearerAuth({ verifier: provider });

  // MCP endpoint — handles POST (messages), GET (SSE), DELETE (session close)
  app.all("/mcp", bearerAuth, async (req, res) => {
    const sessionId = req.headers["mcp-session-id"] as string | undefined;

    if (sessionId) {
      const transport = sessions.get(sessionId);
      if (!transport) {
        res.status(404).json({
          error: "invalid_session",
          error_description: "Unknown session. Start a new session without mcp-session-id.",
        });
        return;
      }
      await transport.handleRequest(req, res);
      return;
    }

    // Only POST can initialize a new session
    if (req.method !== "POST") {
      res.status(400).json({
        error: "bad_request",
        error_description: "No valid session. Send an initialization request first.",
      });
      return;
    }

    const authInfo = req.auth;
    if (!authInfo?.extra?.graphToken) {
      res.status(401).json({ error: "unauthorized" });
      return;
    }

    const graphToken = authInfo.extra.graphToken as string;

    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => randomUUID(),
      onsessioninitialized: (sid) => {
        sessions.set(sid, transport);
      },
    });

    transport.onclose = () => {
      const sid = transport.sessionId;
      if (sid) sessions.delete(sid);
    };

    const server = createSessionServer(graphToken, readOnly);
    // @ts-expect-error StreamableHTTPServerTransport satisfies Transport at runtime; TS strictness mismatch on optional onclose
    await server.connect(transport);

    await transport.handleRequest(req, res);
  });

  app.listen(port, () => {
    console.error(`Microsoft Graph MCP Server (HTTP) listening on port ${port}`);
    console.error(`Base URL: ${baseUrl}`);
    console.error(`Read-only mode: ${readOnly}`);
  });
}

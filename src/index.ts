#!/usr/bin/env node

import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import {
  type AuthenticationResult,
  type Configuration,
  PublicClientApplication,
} from "@azure/msal-node";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { getAuthority, getClientId } from "./config.js";
import { cachePlugin } from "./msal-cache.js";
import { FULL_SCOPES, GraphService, READ_ONLY_SCOPES } from "./services/graph.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerChatTools } from "./tools/chats.js";
import { registerSearchTools } from "./tools/search.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerUsersTools } from "./tools/users.js";

const AUTH_INFO_PATH = join(homedir(), ".msgraph-mcp-auth.json");

/** Check whether CLI args contain --read-only. */
function hasReadOnlyFlag(args: string[]): boolean {
  return args.includes("--read-only");
}

/** Check whether CLI args contain --interactive. */
function hasInteractiveFlag(args: string[]): boolean {
  return args.includes("--interactive");
}

/** Read the persisted auth info file (best-effort). */
async function readAuthInfo(): Promise<Record<string, unknown> | undefined> {
  try {
    const data = await fs.readFile(AUTH_INFO_PATH, "utf8");
    return JSON.parse(data) as Record<string, unknown>;
  } catch {
    return undefined;
  }
}

/** Persist auth result metadata (account, scopes, expiry) with owner-only perms. */
async function saveAuthInfo(result: AuthenticationResult, clientId: string): Promise<void> {
  const authInfo = {
    clientId,
    authenticated: true,
    timestamp: new Date().toISOString(),
    expiresAt: result.expiresOn?.toISOString(),
    account: result.account?.username,
    grantedScopes: result.scopes,
  };
  // 0o600 = owner-only; the file contains username + granted scopes (not the token itself,
  // which lives in the MSAL cache file with its own 0o600 perms).
  await fs.writeFile(AUTH_INFO_PATH, JSON.stringify(authInfo, null, 2), {
    encoding: "utf8",
    mode: 0o600,
  });
  try {
    await fs.chmod(AUTH_INFO_PATH, 0o600);
  } catch {
    // chmod unsupported (Windows) — writeFile mode was best-effort
  }
}

/** Print uniform success banner after either auth flow completes. */
function reportAuthSuccess(
  result: AuthenticationResult,
  modeLabel: string,
  flowLabel: string
): void {
  console.log("\n✅ Authentication successful!");
  console.log(`👤 Signed in as: ${result.account?.username || "Unknown"}`);
  console.log(`🔒 Mode: ${modeLabel}`);
  console.log(`🔁 Flow: ${flowLabel}`);
  console.log(`💾 Credentials saved to: ${AUTH_INFO_PATH}`);
  console.log("🔄 Refresh token cached for automatic renewal");
}

/** Map common AADSTS error codes to actionable messages, then exit non-zero. */
function reportAuthError(error: unknown): never {
  const errorMessage = error instanceof Error ? error.message : String(error);
  if (errorMessage.includes("AADSTS50020")) {
    console.error("\n❌ Authentication failed: User account not in tenant");
  } else if (errorMessage.includes("AADSTS65001")) {
    console.error("\n❌ Authentication failed: Admin consent required");
    console.error("   Grant admin consent for the required permissions in Azure Portal");
  } else if (errorMessage.includes("AADSTS530032") || errorMessage.includes("AADSTS530003")) {
    console.error("\n❌ Authentication failed: Blocked by Conditional Access");
    console.error("   Try --interactive (browser-based) flow instead of device code:");
    console.error("   teams-mcp authenticate --interactive");
  } else {
    console.error("\n❌ Authentication failed:", errorMessage);
  }
  process.exit(1);
}

/** Device-code flow — prints user code + URL, polls until completion. */
async function authenticate(readOnly: boolean) {
  const scopes = readOnly ? READ_ONLY_SCOPES : FULL_SCOPES;
  const modeLabel = readOnly ? "read-only" : "full access";
  const clientId = getClientId();

  console.log("🔐 Microsoft Graph Authentication for MCP Server");
  console.log("=".repeat(50));
  console.log(`Client ID: ${clientId}`);
  console.log(`Authority: ${getAuthority()}`);
  console.log(`Mode: ${modeLabel}`);

  try {
    console.log("\n📱 Using device code flow...");

    const msalConfig: Configuration = {
      auth: {
        clientId,
        authority: getAuthority(),
      },
      cache: {
        cachePlugin, // Use our custom file-based cache for refresh tokens
      },
    };

    const client = new PublicClientApplication(msalConfig);

    const result: AuthenticationResult | null = await client.acquireTokenByDeviceCode({
      scopes,
      deviceCodeCallback: (response) => {
        console.log("\n📱 Please complete authentication:");
        console.log(`🌐 Visit: ${response.verificationUri}`);
        console.log(`🔑 Enter code: ${response.userCode}`);
        console.log("\n⏳ Waiting for you to complete authentication...");
      },
    });

    if (result) {
      await saveAuthInfo(result, clientId);
      reportAuthSuccess(result, modeLabel, "device code");
    }
  } catch (error) {
    reportAuthError(error);
  }
}

/** Interactive auth-code + PKCE flow with loopback redirect; opens browser. */
async function authenticateInteractive(readOnly: boolean) {
  const scopes = readOnly ? READ_ONLY_SCOPES : FULL_SCOPES;
  const modeLabel = readOnly ? "read-only" : "full access";
  const clientId = getClientId();

  console.log("🔐 Microsoft Graph Authentication for MCP Server");
  console.log("=".repeat(50));
  console.log(`Client ID: ${clientId}`);
  console.log(`Authority: ${getAuthority()}`);
  console.log(`Mode: ${modeLabel}`);

  try {
    console.log("\n🌐 Using interactive auth-code flow (browser + loopback)...");

    const msalConfig: Configuration = {
      auth: {
        clientId,
        authority: getAuthority(),
      },
      cache: {
        cachePlugin,
      },
    };

    const client = new PublicClientApplication(msalConfig);

    const result = await client.acquireTokenInteractive({
      scopes,
      openBrowser: async (url) => {
        console.log(`\n🌐 Opening browser to: ${url}`);
        const { default: open } = await import("open");
        await open(url);
      },
      successTemplate:
        "<!DOCTYPE html><html><head><meta charset='utf-8'><title>Authentication successful</title></head><body style='font-family:sans-serif;text-align:center;padding:2em;'><h1 style='color:#2e7d32;'>Authentication successful</h1><p>You can close this tab and return to the terminal.</p></body></html>",
      errorTemplate:
        "<!DOCTYPE html><html><head><meta charset='utf-8'><title>Authentication failed</title></head><body style='font-family:sans-serif;text-align:center;padding:2em;'><h1 style='color:#c62828;'>Authentication failed</h1><p>Check the terminal for details.</p></body></html>",
    });

    if (result) {
      await saveAuthInfo(result, clientId);
      reportAuthSuccess(result, modeLabel, "interactive (auth code + PKCE)");
    }
  } catch (error) {
    reportAuthError(error);
  }
}

/** Report persisted auth state (account, scope mode, token expiry) to stdout. */
async function checkAuth() {
  try {
    const data = await fs.readFile(AUTH_INFO_PATH, "utf8");
    const authInfo = JSON.parse(data);

    if (authInfo.authenticated && authInfo.clientId) {
      console.log("✅ Authentication found");
      console.log(`👤 Account: ${authInfo.account || "Unknown"}`);
      console.log(`📅 Authenticated on: ${authInfo.timestamp}`);

      // Show granted scope mode
      const grantedScopes = authInfo.grantedScopes as string[] | undefined;
      if (grantedScopes) {
        const hasWriteScopes = grantedScopes.some(
          (s: string) =>
            s === "ChannelMessage.Send" ||
            s === "ChannelMessage.ReadWrite" ||
            s === "Chat.ReadWrite" ||
            s === "Files.ReadWrite.All"
        );
        console.log(`🔒 Scope mode: ${hasWriteScopes ? "full access" : "read-only"}`);
      } else {
        console.log("⚠️  Scope mode: unknown (authenticated before read-only support)");
      }

      // Check if we have expiration info
      if (authInfo.expiresAt) {
        const expiresAt = new Date(authInfo.expiresAt);
        const now = new Date();

        if (expiresAt > now) {
          console.log(`⏰ Access token expires: ${expiresAt.toLocaleString()}`);
          console.log("🔄 Refresh token will automatically renew access");
          console.log("🎯 Ready to use with MCP server!");
        } else {
          console.log("⏰ Access token expired - will use refresh token");
          console.log("🎯 Ready to use with MCP server!");
        }
      } else {
        console.log("🎯 Ready to use with MCP server!");
      }
      return true;
    }
  } catch (_error) {
    console.log("❌ No authentication found");
    return false;
  }
  return false;
}

/** Remove persisted auth info + MSAL token cache files. */
async function logout() {
  const CACHE_PATH = join(homedir(), ".teams-mcp-token-cache.json");

  try {
    await fs.unlink(AUTH_INFO_PATH);
  } catch (_error) {
    // Ignore if file doesn't exist
  }

  try {
    await fs.unlink(CACHE_PATH);
  } catch (_error) {
    // Ignore if file doesn't exist
  }

  console.log("✅ Successfully logged out");
  console.log("🔄 Run 'teams-mcp authenticate' (or '--interactive') to re-authenticate");
}

/** Boot MCP server over stdio, registering tool groups; warns on scope mismatch. */
async function startMcpServer(readOnly: boolean) {
  // Create MCP server
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

  // Start server
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(`Microsoft Graph MCP Server started${readOnly ? " (read-only mode)" : ""}`);
}

/** CLI entry point — dispatches subcommands or starts the MCP server. */
async function main() {
  const args = process.argv.slice(2);
  const command = args.find((arg) => arg !== "--read-only" && arg !== "--interactive");

  const readOnly = hasReadOnlyFlag(args) || process.env.TEAMS_MCP_READ_ONLY === "true";
  const interactive = hasInteractiveFlag(args);

  // CLI commands
  switch (command) {
    case "authenticate":
    case "auth":
      if (interactive) {
        await authenticateInteractive(readOnly);
      } else {
        await authenticate(readOnly);
      }
      return;
    case "check":
      await checkAuth();
      return;
    case "logout":
      await logout();
      return;
    case "help":
    case "--help":
    case "-h":
      console.log("Microsoft Graph MCP Server");
      console.log("");
      console.log("Usage:");
      console.log(
        "  teams-mcp authenticate                 # Device-code flow (default)"
      );
      console.log(
        "  teams-mcp authenticate --interactive   # Browser auth-code flow (PKCE, loopback)"
      );
      console.log(
        "  teams-mcp authenticate --read-only     # Authenticate with read-only scopes"
      );
      console.log("  teams-mcp check                        # Check authentication status");
      console.log("  teams-mcp logout                       # Clear authentication");
      console.log("  teams-mcp                              # Start MCP server (default)");
      console.log("");
      console.log("Environment variables:");
      console.log("  TEAMS_MCP_READ_ONLY=true      # Start MCP server in read-only mode");
      console.log("  TEAMS_MCP_CLIENT_ID=<guid>    # Override default Entra app (client) ID");
      console.log("  TEAMS_MCP_AUTHORITY=<url>     # Override default authority");
      console.log("  AUTH_TOKEN=<jwt>              # Use a pre-existing access token");
      return;
    case undefined:
      // No command = start MCP server
      await startMcpServer(readOnly);
      return;
    default:
      console.error(`Unknown command: ${command}`);
      console.error("Use --help to see available commands");
      process.exit(1);
  }
}

// Handle uncaught errors
process.on("uncaughtException", (error) => {
  console.error("Uncaught exception:", error);
  process.exit(1);
});

process.on("unhandledRejection", (reason, promise) => {
  console.error("Unhandled rejection at:", promise, "reason:", reason);
  process.exit(1);
});

main().catch((error) => {
  console.error("Failed to start:", error);
  process.exit(1);
});

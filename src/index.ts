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
import { clearTokenCache, createCachePlugin } from "./msal-cache.js";
import {
  deleteAuthInfoSecure,
  readAuthInfoSecure,
  writeAuthInfoSecure,
} from "./secure-storage.js";
import { FULL_SCOPES, GraphService, READ_ONLY_SCOPES } from "./services/graph.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerChatTools } from "./tools/chats.js";
import { registerSearchTools } from "./tools/search.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerUsersTools } from "./tools/users.js";

// Microsoft Graph CLI app ID (default public client)
const CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const AUTHORITY = "https://login.microsoftonline.com/common";

// Legacy path for migration from plaintext auth file to secure storage
const LEGACY_AUTH_INFO_PATH = join(homedir(), ".msgraph-mcp-auth.json");

/** Check whether CLI args contain --read-only. */
function hasReadOnlyFlag(args: string[]): boolean {
  return args.includes("--read-only");
}

/** Read auth info from OS secure storage; migrates from legacy plaintext file if present. */
async function readAuthInfo(): Promise<Record<string, unknown> | undefined> {
  let data: string | undefined = await readAuthInfoSecure();
  if (data !== undefined) {
    try {
      return JSON.parse(data) as Record<string, unknown>;
    } catch {
      return undefined;
    }
  }
  // Migrate from legacy plaintext file
  try {
    data = await fs.readFile(LEGACY_AUTH_INFO_PATH, "utf8");
    const parsed = JSON.parse(data) as Record<string, unknown>;
    await writeAuthInfoSecure(JSON.stringify(parsed, null, 2));
    await fs.unlink(LEGACY_AUTH_INFO_PATH).catch((err) => {
      console.warn(
        `Failed to remove legacy auth file ${LEGACY_AUTH_INFO_PATH}: ${err}`
      );
      throw err;
    });
    return parsed;
  } catch {
    return undefined;
  }
}

// Authentication functions
async function authenticate(readOnly: boolean) {
  const scopes = readOnly ? READ_ONLY_SCOPES : FULL_SCOPES;
  const modeLabel = readOnly ? "read-only" : "full access";

  console.log("🔐 Microsoft Graph Authentication for MCP Server");
  console.log("=".repeat(50));
  console.log(`Using Microsoft Graph CLI app (${modeLabel})`);

  try {
    console.log("\n📱 Using device code flow...");

    const cachePlugin = await createCachePlugin();
    const msalConfig: Configuration = {
      auth: {
        clientId: CLIENT_ID,
        authority: AUTHORITY,
      },
      cache: {
        cachePlugin,
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
      // Save authentication info (for quick status checks via CLI)
      const authInfo = {
        clientId: CLIENT_ID,
        authenticated: true,
        timestamp: new Date().toISOString(),
        expiresAt: result.expiresOn?.toISOString(),
        account: result.account?.username,
        grantedScopes: result.scopes,
      };

      await writeAuthInfoSecure(JSON.stringify(authInfo, null, 2));

      console.log("\n✅ Authentication successful!");
      console.log(`👤 Signed in as: ${result.account?.username || "Unknown"}`);
      console.log(`🔒 Mode: ${modeLabel}`);
      console.log("💾 Credentials saved to OS secure storage (Keychain / Windows Credential Manager / libsecret)");
      console.log("🔄 Refresh token cached for automatic renewal");
      console.log("\n🚀 You can now use the MCP server in Cursor!");
      console.log("   The server will automatically use these credentials.");
    }
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);

    // Provide helpful error messages for common issues
    if (errorMessage.includes("AADSTS50020")) {
      console.error("\n❌ Authentication failed: User account not in tenant");
    } else if (errorMessage.includes("AADSTS65001")) {
      console.error("\n❌ Authentication failed: Admin consent required");
      console.error("   Grant admin consent for the required permissions in Azure Portal");
    } else {
      console.error("\n❌ Authentication failed:", errorMessage);
    }
    process.exit(1);
  }
}

async function checkAuth() {
  try {
    const authInfo = await readAuthInfo();
    if (!authInfo?.authenticated || !authInfo?.clientId) {
      console.log("❌ No authentication found");
      return false;
    }

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
      const expiresAtVal = authInfo.expiresAt;
      if (expiresAtVal != null && (typeof expiresAtVal === "string" || typeof expiresAtVal === "number" || expiresAtVal instanceof Date)) {
        const expiresAt = new Date(expiresAtVal);
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

async function logout() {
  await deleteAuthInfoSecure();
  await clearTokenCache();

  // Remove legacy plaintext auth file if it still exists
  try {
    await fs.unlink(LEGACY_AUTH_INFO_PATH);
  } catch {
    // Ignore if file doesn't exist
  }

  console.log("✅ Successfully logged out");
  console.log("🔄 Run 'npx @floriscornel/teams-mcp@latest authenticate' to re-authenticate");
}

// MCP Server setup
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

// Main function to handle both CLI and MCP server modes
async function main() {
  const args = process.argv.slice(2);
  const command = args.find((arg) => arg !== "--read-only");

  const readOnly = hasReadOnlyFlag(args) || process.env.TEAMS_MCP_READ_ONLY === "true";

  // CLI commands
  switch (command) {
    case "authenticate":
    case "auth":
      await authenticate(readOnly);
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
        "  npx @floriscornel/teams-mcp@latest authenticate              # Authenticate with full scopes"
      );
      console.log(
        "  npx @floriscornel/teams-mcp@latest authenticate --read-only  # Authenticate with read-only scopes"
      );
      console.log(
        "  npx @floriscornel/teams-mcp@latest check                     # Check authentication status"
      );
      console.log(
        "  npx @floriscornel/teams-mcp@latest logout                    # Clear authentication"
      );
      console.log(
        "  npx @floriscornel/teams-mcp@latest                           # Start MCP server (default)"
      );
      console.log("");
      console.log("Environment variables:");
      console.log("  TEAMS_MCP_READ_ONLY=true  # Start MCP server in read-only mode");
      console.log("  AUTH_TOKEN=<jwt>          # Use a pre-existing access token");
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

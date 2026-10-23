#!/usr/bin/env node

import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import {
  type AuthenticationResult,
  type Configuration,
  PublicClientApplication,
} from "@azure/msal-node";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { cachePlugin, writeOwnerOnly } from "./msal-cache.js";
import { AUTH_INFO_PATH, createMcpServer } from "./server.js";
import { resolveScopes } from "./services/graph.js";

// Microsoft Graph CLI app ID (default public client)
// Override with your own app registration via TEAMS_MCP_CLIENT_ID / TEAMS_MCP_TENANT_ID
const CLIENT_ID = process.env.TEAMS_MCP_CLIENT_ID || "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const AUTHORITY = `https://login.microsoftonline.com/${process.env.TEAMS_MCP_TENANT_ID || "common"}`;

/** Check whether CLI args contain --read-only. */
function hasReadOnlyFlag(args: string[]): boolean {
  return args.includes("--read-only");
}

/** Check whether CLI args contain --interactive. */
function hasInteractiveFlag(args: string[]): boolean {
  return args.includes("--interactive");
}

/** Print the auth banner and build an MSAL client backed by the file token cache. */
function createAuthClient(readOnly: boolean) {
  const scopes = resolveScopes(readOnly);
  const modeLabel = readOnly ? "read-only" : "full access";

  console.log("🔐 Microsoft Graph Authentication for MCP Server");
  console.log("=".repeat(50));
  console.log(`Client ID: ${CLIENT_ID}`);
  console.log(`Authority: ${AUTHORITY}`);
  console.log(`Mode: ${modeLabel}`);

  const msalConfig: Configuration = {
    auth: {
      clientId: CLIENT_ID,
      authority: AUTHORITY,
    },
    cache: {
      cachePlugin, // Use our custom file-based cache for refresh tokens
    },
  };

  return { client: new PublicClientApplication(msalConfig), scopes, modeLabel };
}

/** Persist auth result metadata (for quick status checks via CLI) with owner-only perms. */
async function saveAuthInfo(result: AuthenticationResult): Promise<void> {
  const authInfo = {
    clientId: CLIENT_ID,
    authenticated: true,
    timestamp: new Date().toISOString(),
    expiresAt: result.expiresOn?.toISOString(),
    account: result.account?.username,
    grantedScopes: result.scopes,
  };

  await writeOwnerOnly(AUTH_INFO_PATH, JSON.stringify(authInfo, null, 2));
}

/** Print the success banner shared by both auth flows. */
function reportAuthSuccess(result: AuthenticationResult, modeLabel: string, flowLabel: string) {
  console.log("\n✅ Authentication successful!");
  console.log(`👤 Signed in as: ${result.account?.username || "Unknown"}`);
  console.log(`🔒 Mode: ${modeLabel}`);
  console.log(`🔁 Flow: ${flowLabel}`);
  console.log(`💾 Credentials saved to: ${AUTH_INFO_PATH}`);
  console.log("🔄 Refresh token cached for automatic renewal");
  console.log("\n🚀 You can now use the MCP server in Cursor!");
  console.log("   The server will automatically use these credentials.");
}

/** Map common AADSTS error codes to actionable messages, then exit non-zero. */
function reportAuthError(error: unknown, interactive: boolean): never {
  const errorMessage = error instanceof Error ? error.message : String(error);

  if (errorMessage.includes("AADSTS50020")) {
    console.error("\n❌ Authentication failed: User account not in tenant");
  } else if (errorMessage.includes("AADSTS65001")) {
    console.error("\n❌ Authentication failed: Admin consent required");
    console.error("   Grant admin consent for the required permissions in Azure Portal");
  } else if (errorMessage.includes("AADSTS530032") || errorMessage.includes("AADSTS530003")) {
    console.error("\n❌ Authentication failed: Blocked by Conditional Access");
    if (interactive) {
      // The browser flow is already the fallback; the tenant's policy is the blocker
      console.error("   Check the Entra sign-in logs or ask your administrator which policy");
      console.error("   applies (for example a device compliance requirement)");
    } else {
      console.error("   Try the browser-based flow instead of device code:");
      console.error("   npx @floriscornel/teams-mcp@latest authenticate --interactive");
    }
  } else {
    console.error("\n❌ Authentication failed:", errorMessage);
  }
  process.exit(1);
}

// Authentication functions
async function authenticate(readOnly: boolean) {
  const { client, scopes, modeLabel } = createAuthClient(readOnly);

  try {
    console.log("\n📱 Using device code flow...");

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
      await saveAuthInfo(result);
      reportAuthSuccess(result, modeLabel, "device code");
    }
  } catch (error) {
    reportAuthError(error, false);
  }
}

const SUCCESS_PAGE =
  "<!DOCTYPE html><html><head><meta charset='utf-8'><title>Authentication successful</title></head><body style='font-family:sans-serif;text-align:center;padding:2em;'><h1 style='color:#2e7d32;'>Authentication successful</h1><p>You can close this tab and return to the terminal.</p></body></html>";
const ERROR_PAGE =
  "<!DOCTYPE html><html><head><meta charset='utf-8'><title>Authentication failed</title></head><body style='font-family:sans-serif;text-align:center;padding:2em;'><h1 style='color:#c62828;'>Authentication failed</h1><p>Check the terminal for details.</p></body></html>";

/**
 * Auth-code + PKCE flow with a loopback redirect. Goes through the browser, so
 * Conditional Access can evaluate the device — works on tenants that block device code.
 */
async function authenticateInteractive(readOnly: boolean) {
  const { client, scopes, modeLabel } = createAuthClient(readOnly);

  try {
    console.log("\n🌐 Using interactive auth-code flow (browser + loopback)...");

    const result = await client.acquireTokenInteractive({
      scopes,
      openBrowser: async (url) => {
        console.log(`\n🌐 Opening browser to: ${url}`);
        const { default: open } = await import("open");
        await open(url);
      },
      successTemplate: SUCCESS_PAGE,
      errorTemplate: ERROR_PAGE,
    });

    await saveAuthInfo(result);
    reportAuthSuccess(result, modeLabel, "interactive (auth code + PKCE)");
  } catch (error) {
    reportAuthError(error, true);
  }
}

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
  console.log(
    "🔄 Run 'npx @floriscornel/teams-mcp@latest authenticate [--interactive]' to re-authenticate"
  );
}

// MCP Server setup
async function startMcpServer(readOnly: boolean) {
  const server = await createMcpServer(readOnly);

  // Start server
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(`Microsoft Graph MCP Server started${readOnly ? " (read-only mode)" : ""}`);
}

// Main function to handle both CLI and MCP server modes
async function main() {
  const args = process.argv.slice(2);
  const command = args.find((arg) => arg !== "--read-only" && arg !== "--interactive");

  const readOnly = hasReadOnlyFlag(args) || process.env.TEAMS_MCP_READ_ONLY === "true";

  // CLI commands
  switch (command) {
    case "authenticate":
    case "auth":
      if (hasInteractiveFlag(args)) {
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
        "  npx @floriscornel/teams-mcp@latest authenticate              # Authenticate with full scopes"
      );
      console.log(
        "  npx @floriscornel/teams-mcp@latest authenticate --read-only  # Authenticate with read-only scopes"
      );
      console.log(
        "  npx @floriscornel/teams-mcp@latest authenticate --interactive # Authenticate via browser (auth code + PKCE)"
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

// Exposed for tests so they can await full CLI startup deterministically.
export const mainPromise = main().catch((error) => {
  console.error("Failed to start:", error);
  process.exit(1);
});

#!/usr/bin/env node

import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { DeviceCodeCredential, useIdentityPlugin } from "@azure/identity";
import { cachePersistencePlugin } from "@azure/identity-cache-persistence";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { GraphService } from "./services/graph.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerChatTools } from "./tools/chats.js";
import { registerSearchTools } from "./tools/search.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerUsersTools } from "./tools/users.js";

// Enable persistent token caching
useIdentityPlugin(cachePersistencePlugin);

const CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const TOKEN_PATH = join(homedir(), ".msgraph-mcp-auth.json");

const SCOPES = [
  "User.Read",
  "User.ReadBasic.All",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "ChannelMessage.Send",
  "TeamMember.Read.All",
  "Chat.ReadBasic",
  "Chat.ReadWrite",
];

// Authentication functions
async function authenticate() {
  console.log("🔐 Microsoft Graph Authentication for MCP Server");
  console.log("=".repeat(50));

  try {
    const credential = new DeviceCodeCredential({
      clientId: CLIENT_ID,
      tenantId: "common",
      tokenCachePersistenceOptions: {
        enabled: true,
      },
      userPromptCallback: (info) => {
        console.log("\n📱 Please complete authentication:");
        console.log(`🌐 Visit: ${info.verificationUri}`);
        console.log(`🔑 Enter code: ${info.userCode}`);
        console.log("\n⏳ Waiting for you to complete authentication...");
      },
    });

    // Test the credential - this will trigger auth flow if needed
    const token = await credential.getToken(SCOPES);

    if (token) {
      console.log("\n✅ Authentication successful!");
      console.log("💾 Credentials securely cached by Azure Identity");
      console.log(`⏰ Token expires: ${new Date(token.expiresOnTimestamp).toLocaleString()}`);
      console.log("\n🚀 You can now use the MCP server in Cursor!");
      console.log("   The server will automatically use cached credentials.");

      // Clean up old auth file if it exists
      try {
        await fs.unlink(TOKEN_PATH);
        console.log("🧹 Cleaned up old authentication file");
      } catch {
        // File doesn't exist, which is fine
      }
    }
  } catch (error) {
    console.error(
      "\n❌ Authentication failed:",
      error instanceof Error ? error.message : String(error)
    );
    process.exit(1);
  }
}

async function checkAuth() {
  try {
    const credential = new DeviceCodeCredential({
      clientId: CLIENT_ID,
      tenantId: "common",
      tokenCachePersistenceOptions: {
        enabled: true,
      },
    });

    // Try to get a token silently (from cache)
    const token = await credential.getToken(["User.Read"]);

    if (token) {
      console.log("✅ Authentication found in secure cache");
      console.log(`⏰ Token expires: ${new Date(token.expiresOnTimestamp).toLocaleString()}`);
      console.log("🎯 Ready to use with MCP server!");
      return true;
    }

    // Check for old auth file and suggest migration
    try {
      await fs.access(TOKEN_PATH);
      console.log("⚠️  Found old authentication file.");
      console.log("🔄 Please re-authenticate to use secure token caching:");
      console.log("   npx @floriscornel/teams-mcp@latest authenticate");
      return false;
    } catch {
      // No old file either
    }

    console.log("❌ No authentication found");
    return false;
  } catch (error) {
    console.log(
      "❌ Authentication check failed:",
      error instanceof Error ? error.message : String(error)
    );
    return false;
  }
}

async function logout() {
  try {
    // Clean up old auth file if it exists
    await fs.unlink(TOKEN_PATH);
    console.log("🧹 Cleaned up old authentication file");
  } catch {
    // File doesn't exist
  }

  console.log("ℹ️  Azure Identity cached credentials will expire automatically");
  console.log("🔄 Run 'npx @floriscornel/teams-mcp@latest authenticate' to re-authenticate");
}

// MCP Server setup
async function startMcpServer() {
  // Create MCP server
  const server = new McpServer({
    name: "teams-mcp",
    version: "0.3.1",
  });

  // Initialize Graph service (singleton)
  const graphService = GraphService.getInstance();

  // Register all tools
  registerAuthTools(server, graphService);
  registerUsersTools(server, graphService);
  registerTeamsTools(server, graphService);
  registerChatTools(server, graphService);
  registerSearchTools(server, graphService);

  // Start server
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Microsoft Graph MCP Server started");
}

// Main function to handle both CLI and MCP server modes
async function main() {
  const args = process.argv.slice(2);
  const command = args[0];

  // CLI commands
  switch (command) {
    case "authenticate":
    case "auth":
      await authenticate();
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
        "  npx @floriscornel/teams-mcp@latest authenticate # Authenticate with Microsoft"
      );
      console.log(
        "  npx @floriscornel/teams-mcp@latest check        # Check authentication status"
      );
      console.log("  npx @floriscornel/teams-mcp@latest logout       # Clear authentication");
      console.log("  npx @floriscornel/teams-mcp@latest              # Start MCP server (default)");
      return;
    case undefined:
      // No command = start MCP server
      await startMcpServer();
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

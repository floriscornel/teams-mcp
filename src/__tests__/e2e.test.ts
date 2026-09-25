import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { beforeAll, describe, expect, it } from "vitest";
import { createMcpServer } from "../server.js";

// JWT accepted by GraphService.validateToken (aud must target Microsoft Graph).
// MSW intercepts all Graph HTTP traffic, so the token never needs to be real.
const AUTH_TOKEN = (() => {
  const payload = Buffer.from(
    JSON.stringify({ aud: "https://graph.microsoft.com", sub: "e2e-test" })
  ).toString("base64");
  return `e30.${payload}.not-a-real-signature`;
})();

/** All tool names registered in full mode. */
const ALL_TOOL_NAMES = [
  "auth_status",
  "create_chat",
  "delete_channel_message",
  "delete_chat_message",
  "download_chat_hosted_content",
  "download_message_hosted_content",
  "get_channel_message_replies",
  "get_channel_messages",
  "get_chat_messages",
  "get_current_user",
  "get_my_mentions",
  "get_user",
  "list_channels",
  "list_chats",
  "list_team_members",
  "list_teams",
  "reply_to_channel_message",
  "search_messages",
  "search_users",
  "search_users_for_mentions",
  "send_channel_message",
  "send_chat_message",
  "send_file_to_channel",
  "send_file_to_chat",
  "set_channel_message_reaction",
  "set_chat_message_reaction",
  "unset_channel_message_reaction",
  "unset_chat_message_reaction",
  "update_channel_message",
  "update_chat_message",
];

/** Tools available in read-only mode (write tools excluded). */
const READ_ONLY_TOOL_NAMES = ALL_TOOL_NAMES.filter(
  (name) =>
    ![
      "send_channel_message",
      "reply_to_channel_message",
      "update_channel_message",
      "delete_channel_message",
      "send_file_to_channel",
      "send_chat_message",
      "create_chat",
      "update_chat_message",
      "delete_chat_message",
      "send_file_to_chat",
      "set_channel_message_reaction",
      "set_chat_message_reaction",
      "unset_channel_message_reaction",
      "unset_chat_message_reaction",
    ].includes(name)
);

interface TextResult {
  content?: Array<{ type: string; text?: string }>;
  isError?: boolean;
}

function resultText(result: TextResult): string {
  return (result.content ?? [])
    .map((item) => (item.type === "text" ? (item.text ?? "") : ""))
    .join("");
}

/** Connect an MCP client to a fresh server instance over an in-memory transport pair. */
async function connectClient(readOnly: boolean): Promise<{ server: McpServer; client: Client }> {
  const server = await createMcpServer(readOnly);
  const client = new Client({ name: "e2e-test-client", version: "1.0.0" });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);
  return { server, client };
}

describe("E2E: MCP server over InMemoryTransport (full mode)", () => {
  let client: Client;

  beforeAll(async () => {
    process.env.AUTH_TOKEN = AUTH_TOKEN;
    ({ client } = await connectClient(false));
  });

  describe("protocol-level tool discovery", () => {
    it("exposes exactly the expected set of tools via tools/list", async () => {
      const { tools } = await client.listTools();
      expect(tools.map((t) => t.name).sort()).toEqual(ALL_TOOL_NAMES);
    });

    it("returns tool metadata (title, description, input schema, annotations)", async () => {
      const { tools } = await client.listTools();
      const authStatus = tools.find((t) => t.name === "auth_status");
      expect(authStatus).toBeDefined();
      expect(authStatus?.title).toBe("Auth Status");
      expect(authStatus?.description).toContain("authentication");
      expect(authStatus?.inputSchema).toHaveProperty("properties");
      expect(authStatus?.annotations?.readOnlyHint).toBe(true);
    });

    it("declares write tools as non-read-only via annotations", async () => {
      const { tools } = await client.listTools();
      const send = tools.find((t) => t.name === "send_chat_message");
      expect(send?.annotations?.readOnlyHint).toBe(false);
      const del = tools.find((t) => t.name === "delete_chat_message");
      expect(del?.annotations?.destructiveHint).toBe(true);
    });
  });

  describe("tool invocation end-to-end", () => {
    it("auth_status reports an authenticated session", async () => {
      const result = (await client.callTool({ name: "auth_status", arguments: {} })) as TextResult;
      expect(result.isError).toBeFalsy();
      expect(resultText(result)).toContain("Authenticated as Test User");
      expect(resultText(result)).toContain("test.user@example.com");
    });

    it("get_current_user returns the mock user profile", async () => {
      const result = (await client.callTool({
        name: "get_current_user",
        arguments: {},
      })) as TextResult;
      const user = JSON.parse(resultText(result));
      expect(user).toMatchObject({
        displayName: "Test User",
        userPrincipalName: "test.user@example.com",
        id: "test-user-id",
      });
    });

    it("search_users returns matching users for a query", async () => {
      const result = (await client.callTool({
        name: "search_users",
        arguments: { query: "test" },
      })) as TextResult;
      const users = JSON.parse(resultText(result));
      expect(users).toHaveLength(1);
      expect(users[0].displayName).toBe("Test User");
    });

    it("list_teams returns the user's joined teams", async () => {
      const result = (await client.callTool({ name: "list_teams", arguments: {} })) as TextResult;
      expect(resultText(result)).toContain("Test Team");
    });

    it("list_channels returns channels for a team", async () => {
      const result = (await client.callTool({
        name: "list_channels",
        arguments: { teamId: "test-team-id" },
      })) as TextResult;
      expect(resultText(result)).toContain("General");
    });

    it("get_chat_messages returns parsed messages with markdown content", async () => {
      const result = (await client.callTool({
        name: "get_chat_messages",
        arguments: { chatId: "test-chat-id", limit: 5 },
      })) as TextResult;
      const payload = JSON.parse(resultText(result));
      expect(payload.totalReturned).toBe(1);
      expect(payload.messages[0]).toMatchObject({
        content: "Test message content",
        from: "Test User",
      });
    });

    it("send_chat_message posts through to the Graph API and reports success", async () => {
      const result = (await client.callTool({
        name: "send_chat_message",
        arguments: { chatId: "test-chat-id", message: "**Hello** from E2E", format: "markdown" },
      })) as TextResult;
      expect(result.isError).toBeFalsy();
      expect(resultText(result)).toContain("Message sent successfully");
      expect(resultText(result)).toContain("new-chat-message-id");
    });

    it("create_chat resolves members and creates the conversation", async () => {
      const result = (await client.callTool({
        name: "create_chat",
        arguments: { userEmails: ["test.user@example.com"] },
      })) as TextResult;
      expect(resultText(result)).toContain("Chat created successfully");
      expect(resultText(result)).toContain("new-chat-id");
    });

    it("returns error text when Graph API responds with 404", async () => {
      const result = (await client.callTool({
        name: "get_user",
        arguments: { userId: "nonexistent-user" },
      })) as TextResult;
      expect(resultText(result)).toContain("❌ Error:");
    });

    it("returns a protocol error result for invalid tool arguments", async () => {
      const result = (await client.callTool({
        name: "get_chat_messages",
        arguments: {},
      })) as TextResult;
      expect(result.isError).toBe(true);
    });
  });
});

describe("E2E: MCP server over InMemoryTransport (read-only mode)", () => {
  let client: Client;

  beforeAll(async () => {
    process.env.AUTH_TOKEN = AUTH_TOKEN;
    ({ client } = await connectClient(true));
  });

  it("only registers read-only tools", async () => {
    const { tools } = await client.listTools();
    expect(tools.map((t) => t.name).sort()).toEqual(READ_ONLY_TOOL_NAMES);
  });

  it("exposes read tools that still work end-to-end", async () => {
    const result = (await client.callTool({
      name: "list_chats",
      arguments: {},
    })) as TextResult;
    expect(result.isError).toBeFalsy();
    expect(resultText(result)).toContain("Test Chat");
  });

  it("returns a tool-not-found error for unregistered write tools", async () => {
    const result = (await client.callTool({
      name: "send_chat_message",
      arguments: { chatId: "x", message: "hi" },
    })) as TextResult;
    expect(result.isError).toBe(true);
    expect(resultText(result)).toContain("not found");
  });
});

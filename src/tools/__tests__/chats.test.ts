import type { Client } from "@microsoft/microsoft-graph-client";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { GraphService } from "../../services/graph.js";
import { registerChatTools } from "../chats.js";

// Mock the Graph service
const mockGraphService = {
  getClient: vi.fn(),
} as unknown as GraphService;

// Mock the MCP server
const mockServer = {
  tool: vi.fn(),
} as unknown as McpServer;

// Mock client responses
const mockClient = {
  api: vi.fn(),
} as unknown as Client;

describe("Chat Tools", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockGraphService.getClient = vi.fn().mockResolvedValue(mockClient);
  });

  describe("registerChatTools", () => {
    it("should register all chat tools", () => {
      registerChatTools(mockServer, mockGraphService);

      expect(mockServer.tool).toHaveBeenCalledTimes(4);
      expect(mockServer.tool).toHaveBeenCalledWith(
        "list_chats",
        expect.any(String),
        {},
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "get_chat_messages",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "send_chat_message",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "create_chat",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
    });

    it("should register all chat tools when readOnly is false (backward compatibility)", () => {
      registerChatTools(mockServer, mockGraphService, false);

      expect(mockServer.tool).toHaveBeenCalledTimes(4);
      const toolNames = vi.mocked(mockServer.tool).mock.calls.map(([name]: [string]) => name);
      expect(toolNames).toContain("list_chats");
      expect(toolNames).toContain("get_chat_messages");
      expect(toolNames).toContain("send_chat_message");
      expect(toolNames).toContain("create_chat");
    });

    it("should register only read tools when readOnly is true", () => {
      registerChatTools(mockServer, mockGraphService, true);

      expect(mockServer.tool).toHaveBeenCalledTimes(2);
      const toolNames = vi.mocked(mockServer.tool).mock.calls.map(([name]: [string]) => name);
      expect(toolNames).toContain("list_chats");
      expect(toolNames).toContain("get_chat_messages");
      expect(toolNames).not.toContain("send_chat_message");
      expect(toolNames).not.toContain("create_chat");
    });
  });

  describe("list_chats", () => {
    let listChatsHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService);
      const call = vi.mocked(mockServer.tool).mock.calls.find(([name]) => name === "list_chats");
      listChatsHandler = call?.[3] as unknown as (args: any) => Promise<any>;
    });

    it("should return chat list successfully", async () => {
      const mockChats = [
        {
          id: "chat1",
          topic: "Test Chat 1",
          chatType: "group",
          members: [{ displayName: "user1" }, { displayName: "user2" }],
        },
        {
          id: "chat2",
          topic: null,
          chatType: "oneOnOne",
          members: [{ displayName: "user1" }],
        },
      ];

      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: mockChats }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await listChatsHandler();

      expect(mockClient.api).toHaveBeenCalledWith("/me/chats?$expand=members");
      expect(result.content[0].type).toBe("text");

      const parsedText = JSON.parse(result.content[0].text);
      expect(parsedText).toHaveLength(2);
      expect(parsedText[0]).toEqual({
        id: "chat1",
        topic: "Test Chat 1",
        chatType: "group",
        members: "user1, user2",
      });
      expect(parsedText[1]).toEqual({
        id: "chat2",
        topic: "No topic",
        chatType: "oneOnOne",
        members: "user1",
      });
    });

    it("should handle no chats found", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: [] }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await listChatsHandler();

      expect(result.content[0].text).toBe("No chats found.");
    });

    it("should handle null response", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue(null),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await listChatsHandler();

      expect(result.content[0].text).toBe("No chats found.");
    });

    it("should handle errors gracefully", async () => {
      const mockApiChain = {
        get: vi.fn().mockRejectedValue(new Error("API Error")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await listChatsHandler();

      expect(result.content[0].text).toBe("❌ Error: API Error");
    });

    it("should handle unknown errors", async () => {
      const mockApiChain = {
        get: vi.fn().mockRejectedValue("Unknown error"),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await listChatsHandler();

      expect(result.content[0].text).toBe("❌ Error: Unknown error occurred");
    });
  });

  describe("get_chat_messages", () => {
    let getChatMessagesHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "get_chat_messages");
      getChatMessagesHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should get chat messages with default parameters", async () => {
      const mockMessages = [
        {
          id: "msg1",
          body: { content: "Hello world" },
          from: { user: { displayName: "John Doe" } },
          createdDateTime: "2023-01-01T10:00:00Z",
        },
        {
          id: "msg2",
          body: { content: "How are you?" },
          from: { user: { displayName: "Jane Smith" } },
          createdDateTime: "2023-01-01T11:00:00Z",
        },
      ];

      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: mockMessages }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getChatMessagesHandler({
        chatId: "chat123",
      });

      expect(mockClient.api).toHaveBeenCalledWith(
        "/me/chats/chat123/messages?$top=20&$orderby=createdDateTime desc"
      );

      const parsedResponse = JSON.parse(result.content[0].text);
      expect(parsedResponse.messages).toHaveLength(2);
      expect(parsedResponse.filteringMethod).toBe("server-side");
      expect(parsedResponse.totalReturned).toBe(2);
    });

    it("should apply all filtering options", async () => {
      const mockMessages = [
        {
          id: "msg1",
          body: { content: "Hello" },
          from: { user: { displayName: "John" } },
          createdDateTime: "2023-01-01T10:00:00Z",
        },
      ];

      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: mockMessages }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const _result = await getChatMessagesHandler({
        chatId: "chat123",
        limit: 10,
        fromUser: "user123",
        orderBy: "lastModifiedDateTime",
        descending: true, // Changed to true since ascending is not supported
      });

      expect(mockClient.api).toHaveBeenCalledWith(
        "/me/chats/chat123/messages?$top=10&$orderby=lastModifiedDateTime desc&$filter=from/user/id eq 'user123'"
      );
    });

    it("should reject ascending order for datetime fields", async () => {
      const result = await getChatMessagesHandler({
        chatId: "chat123",
        orderBy: "lastModifiedDateTime",
        descending: false,
      });

      expect(result.content[0].text).toBe(
        "❌ Error: QueryOptions to order by 'LastModifiedDateTime' in 'Ascending' direction is not supported."
      );
    });

    it("should apply client-side date filtering", async () => {
      const mockMessages = [
        {
          id: "msg1",
          body: { content: "Old message" },
          from: { user: { displayName: "John" } },
          createdDateTime: "2023-01-01T08:00:00Z", // Should be filtered out
        },
        {
          id: "msg2",
          body: { content: "New message" },
          from: { user: { displayName: "Jane" } },
          createdDateTime: "2023-01-01T12:00:00Z", // Should be included
        },
      ];

      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: mockMessages }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getChatMessagesHandler({
        chatId: "chat123",
        since: "2023-01-01T10:00:00Z",
        until: "2023-01-01T15:00:00Z",
      });

      const parsedResponse = JSON.parse(result.content[0].text);
      expect(parsedResponse.messages).toHaveLength(1);
      expect(parsedResponse.messages[0].content).toBe("New message");
      expect(parsedResponse.filteringMethod).toBe("client-side");
    });

    it("should handle messages without createdDateTime in date filtering", async () => {
      const mockMessages = [
        {
          id: "msg1",
          body: { content: "Message without date" },
          from: { user: { displayName: "John" } },
          createdDateTime: null,
        },
      ];

      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: mockMessages }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getChatMessagesHandler({
        chatId: "chat123",
        since: "2023-01-01T10:00:00Z",
      });

      const parsedResponse = JSON.parse(result.content[0].text);
      expect(parsedResponse.messages).toHaveLength(1); // Should be included
    });

    it("should handle no messages found", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: [] }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getChatMessagesHandler({ chatId: "chat123" });

      expect(result.content[0].text).toBe(
        "No messages found in this chat with the specified filters."
      );
    });

    it("should handle errors", async () => {
      const mockApiChain = {
        get: vi.fn().mockRejectedValue(new Error("Chat not found")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getChatMessagesHandler({ chatId: "chat123" });

      expect(result.content[0].text).toBe("❌ Error: Chat not found");
    });

    describe("pagination", () => {
      it("should fetch single page when fetchAll is false", async () => {
        const mockMessages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const mockApiChain = {
          get: vi.fn().mockResolvedValue({
            value: mockMessages,
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/nextPage",
          }),
        };
        mockClient.api = vi.fn().mockReturnValue(mockApiChain);

        const result = await getChatMessagesHandler({
          chatId: "chat123",
          limit: 100,
          fetchAll: false,
        });

        // Should only call API once
        expect(mockClient.api).toHaveBeenCalledTimes(1);

        const parsedResponse = JSON.parse(result.content[0].text);
        expect(parsedResponse.messages).toHaveLength(50);
      });

      it("should fetch multiple pages when fetchAll is true", async () => {
        const page1Messages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const page2Messages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i + 50}`,
          body: { content: `Message ${i + 50}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - (i + 50) * 1000).toISOString(),
        }));

        const page3Messages = Array.from({ length: 30 }, (_, i) => ({
          id: `msg${i + 100}`,
          body: { content: `Message ${i + 100}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - (i + 100) * 1000).toISOString(),
        }));

        const mockApiChain1 = {
          get: vi.fn().mockResolvedValue({
            value: page1Messages,
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/nextPage2",
          }),
        };

        const mockApiChain2 = {
          get: vi.fn().mockResolvedValue({
            value: page2Messages,
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/nextPage3",
          }),
        };

        const mockApiChain3 = {
          get: vi.fn().mockResolvedValue({
            value: page3Messages,
            "@odata.nextLink": undefined, // No more pages
          }),
        };

        mockClient.api = vi
          .fn()
          .mockReturnValueOnce(mockApiChain1)
          .mockReturnValueOnce(mockApiChain2)
          .mockReturnValueOnce(mockApiChain3);

        const result = await getChatMessagesHandler({
          chatId: "chat123",
          limit: 200,
          fetchAll: true,
        });

        // Should call API three times (initial + 2 pagination calls)
        expect(mockClient.api).toHaveBeenCalledTimes(3);

        const parsedResponse = JSON.parse(result.content[0].text);
        expect(parsedResponse.messages).toHaveLength(130); // 50 + 50 + 30
      });

      it("should stop fetching when limit is reached", async () => {
        const page1Messages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const page2Messages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i + 50}`,
          body: { content: `Message ${i + 50}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - (i + 50) * 1000).toISOString(),
        }));

        const mockApiChain1 = {
          get: vi.fn().mockResolvedValue({
            value: page1Messages,
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/nextPage2",
          }),
        };

        const mockApiChain2 = {
          get: vi.fn().mockResolvedValue({
            value: page2Messages,
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/nextPage3",
          }),
        };

        mockClient.api = vi
          .fn()
          .mockReturnValueOnce(mockApiChain1)
          .mockReturnValueOnce(mockApiChain2);

        const result = await getChatMessagesHandler({
          chatId: "chat123",
          limit: 75,
          fetchAll: true,
        });

        // Should only call API twice because limit is reached
        expect(mockClient.api).toHaveBeenCalledTimes(2);

        const parsedResponse = JSON.parse(result.content[0].text);
        // Should be limited to 75 messages even though 100 were fetched
        expect(parsedResponse.messages).toHaveLength(75);
      });

      it("should stop pagination when no nextLink is present", async () => {
        const mockMessages = Array.from({ length: 30 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const mockApiChain = {
          get: vi.fn().mockResolvedValue({
            value: mockMessages,
            "@odata.nextLink": undefined, // No more pages
          }),
        };
        mockClient.api = vi.fn().mockReturnValue(mockApiChain);

        const result = await getChatMessagesHandler({
          chatId: "chat123",
          limit: 100,
          fetchAll: true,
        });

        // Should only call API once since there's no nextLink
        expect(mockClient.api).toHaveBeenCalledTimes(1);

        const parsedResponse = JSON.parse(result.content[0].text);
        expect(parsedResponse.messages).toHaveLength(30);
      });

      it("should handle pagination errors gracefully", async () => {
        const page1Messages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const mockApiChain1 = {
          get: vi.fn().mockResolvedValue({
            value: page1Messages,
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/nextPage2",
          }),
        };

        const mockApiChain2 = {
          get: vi.fn().mockRejectedValue(new Error("Network error")),
        };

        mockClient.api = vi
          .fn()
          .mockReturnValueOnce(mockApiChain1)
          .mockReturnValueOnce(mockApiChain2);

        const result = await getChatMessagesHandler({
          chatId: "chat123",
          limit: 100,
          fetchAll: true,
        });

        // Should return the messages from the first page even though second page failed
        const parsedResponse = JSON.parse(result.content[0].text);
        expect(parsedResponse.messages).toHaveLength(50);
      });

      it("should use smaller page size when fetchAll is true", async () => {
        const mockMessages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const mockApiChain = {
          get: vi.fn().mockResolvedValue({
            value: mockMessages,
          }),
        };
        mockClient.api = vi.fn().mockReturnValue(mockApiChain);

        await getChatMessagesHandler({
          chatId: "chat123",
          limit: 100,
          fetchAll: true,
        });

        // Should use page size of 50 instead of the full limit
        expect(mockClient.api).toHaveBeenCalledWith(
          "/me/chats/chat123/messages?$top=50&$orderby=createdDateTime desc"
        );
      });

      it("should use limit as page size when fetchAll is false and limit is small", async () => {
        const mockMessages = Array.from({ length: 10 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const mockApiChain = {
          get: vi.fn().mockResolvedValue({
            value: mockMessages,
          }),
        };
        mockClient.api = vi.fn().mockReturnValue(mockApiChain);

        await getChatMessagesHandler({
          chatId: "chat123",
          limit: 10,
          fetchAll: false,
        });

        // Should use limit (10) as page size since it's smaller than 50
        expect(mockClient.api).toHaveBeenCalledWith(
          "/me/chats/chat123/messages?$top=10&$orderby=createdDateTime desc"
        );
      });
    });
  });

  describe("send_chat_message", () => {
    let sendChatMessageHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "send_chat_message");
      sendChatMessageHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should send message with default importance", async () => {
      const mockResponse = { id: "newmsg123" };

      const mockApiChain = {
        post: vi.fn().mockResolvedValue(mockResponse),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendChatMessageHandler({
        chatId: "chat123",
        message: "Hello world!",
      });

      expect(mockClient.api).toHaveBeenCalledWith("/me/chats/chat123/messages");
      expect(mockApiChain.post).toHaveBeenCalledWith({
        body: {
          content: "Hello world!",
          contentType: "text",
        },
        importance: "normal",
      });
      expect(result.content[0].text).toBe("✅ Message sent successfully. Message ID: newmsg123");
    });

    it("should send message with custom importance", async () => {
      const mockResponse = { id: "newmsg456" };

      const mockApiChain = {
        post: vi.fn().mockResolvedValue(mockResponse),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendChatMessageHandler({
        chatId: "chat123",
        message: "Urgent message",
        importance: "urgent",
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({
        body: {
          content: "Urgent message",
          contentType: "text",
        },
        importance: "urgent",
      });
      expect(result.content[0].text).toBe("✅ Message sent successfully. Message ID: newmsg456");
    });

    it("should handle send errors", async () => {
      const mockApiChain = {
        post: vi.fn().mockRejectedValue(new Error("Permission denied")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendChatMessageHandler({
        chatId: "chat123",
        message: "Test message",
      });

      expect(result.content[0].text).toBe("❌ Failed to send message: Permission denied");
    });

    it("should send message with markdown format", async () => {
      const mockResponse = { id: "mdmsg123" };
      const mockApiChain = { post: vi.fn().mockResolvedValue(mockResponse) };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendChatMessageHandler({
        chatId: "chat123",
        message: "**Bold** _Italic_",
        format: "markdown",
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({
        body: {
          content: expect.stringContaining("<strong>Bold</strong>"),
          contentType: "html",
        },
        importance: "normal",
      });
      expect(result.content[0].text).toBe("✅ Message sent successfully. Message ID: mdmsg123");
    });

    it("should send message with text format (default)", async () => {
      const mockResponse = { id: "txtmsg123" };
      const mockApiChain = { post: vi.fn().mockResolvedValue(mockResponse) };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendChatMessageHandler({
        chatId: "chat123",
        message: "Plain text message",
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({
        body: {
          content: "Plain text message",
          contentType: "text",
        },
        importance: "normal",
      });
      expect(result.content[0].text).toBe("✅ Message sent successfully. Message ID: txtmsg123");
    });

    it("should fallback to text for invalid format", async () => {
      const mockResponse = { id: "fallbackmsg123" };
      const mockApiChain = { post: vi.fn().mockResolvedValue(mockResponse) };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendChatMessageHandler({
        chatId: "chat123",
        message: "Fallback message",
        format: "invalid-format",
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({
        body: {
          content: "Fallback message",
          contentType: "text",
        },
        importance: "normal",
      });
      expect(result.content[0].text).toBe(
        "✅ Message sent successfully. Message ID: fallbackmsg123"
      );
    });
  });

  describe("create_chat", () => {
    let createChatHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService);
      const call = vi.mocked(mockServer.tool).mock.calls.find(([name]) => name === "create_chat");
      createChatHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should create one-on-one chat", async () => {
      const mockMe = { id: "currentuser123" };
      const mockUser = { id: "otheruser456" };
      const mockNewChat = { id: "newchat789" };

      const mockApiChain = {
        get: vi
          .fn()
          .mockResolvedValueOnce(mockMe) // /me call
          .mockResolvedValueOnce(mockUser), // /users/email call
        post: vi.fn().mockResolvedValue(mockNewChat),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await createChatHandler({
        userEmails: ["other@example.com"],
      });

      expect(mockClient.api).toHaveBeenCalledWith("/me");
      expect(mockClient.api).toHaveBeenCalledWith("/users/other@example.com");
      expect(mockClient.api).toHaveBeenCalledWith("/chats");

      expect(mockApiChain.post).toHaveBeenCalledWith({
        chatType: "oneOnOne",
        members: [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: { id: "currentuser123" },
            roles: ["owner"],
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: { id: "otheruser456" },
            roles: ["member"],
          },
        ],
      });

      expect(result.content[0].text).toBe("✅ Chat created successfully. Chat ID: newchat789");
    });

    it("should create group chat with topic", async () => {
      const mockMe = { id: "currentuser123" };
      const mockUser1 = { id: "user1" };
      const mockUser2 = { id: "user2" };
      const mockNewChat = { id: "groupchat123" };

      const mockApiChain = {
        get: vi
          .fn()
          .mockResolvedValueOnce(mockMe) // /me call
          .mockResolvedValueOnce(mockUser1) // first user
          .mockResolvedValueOnce(mockUser2), // second user
        post: vi.fn().mockResolvedValue(mockNewChat),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const _result = await createChatHandler({
        userEmails: ["user1@example.com", "user2@example.com"],
        topic: "Project Discussion",
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({
        chatType: "group",
        topic: "Project Discussion",
        members: [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: { id: "currentuser123" },
            roles: ["owner"],
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: { id: "user1" },
            roles: ["member"],
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: { id: "user2" },
            roles: ["member"],
          },
        ],
      });
    });

    it("should ignore topic for one-on-one chats", async () => {
      const mockMe = { id: "currentuser123" };
      const mockUser = { id: "otheruser456" };
      const mockNewChat = { id: "newchat789" };

      const mockApiChain = {
        get: vi.fn().mockResolvedValueOnce(mockMe).mockResolvedValueOnce(mockUser),
        post: vi.fn().mockResolvedValue(mockNewChat),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const _result = await createChatHandler({
        userEmails: ["other@example.com"],
        topic: "This should be ignored",
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({
        chatType: "oneOnOne",
        members: expect.any(Array),
      });

      // Should not include topic in the payload
      const postCall = mockApiChain.post.mock.calls[0][0];
      expect(postCall).not.toHaveProperty("topic");
    });

    it("should handle user lookup errors", async () => {
      const mockMe = { id: "currentuser123" };

      const mockApiChain = {
        get: vi
          .fn()
          .mockResolvedValueOnce(mockMe)
          .mockRejectedValueOnce(new Error("User not found")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await createChatHandler({
        userEmails: ["nonexistent@example.com"],
      });

      expect(result.content[0].text).toBe("❌ Error: User not found");
    });

    it("should handle chat creation errors", async () => {
      const mockMe = { id: "currentuser123" };
      const mockUser = { id: "otheruser456" };

      const mockApiChain = {
        get: vi.fn().mockResolvedValueOnce(mockMe).mockResolvedValueOnce(mockUser),
        post: vi.fn().mockRejectedValue(new Error("Failed to create chat")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await createChatHandler({
        userEmails: ["other@example.com"],
      });

      expect(result.content[0].text).toBe("❌ Error: Failed to create chat");
    });
  });
});

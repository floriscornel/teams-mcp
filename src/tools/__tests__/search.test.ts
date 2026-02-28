import type { Client } from "@microsoft/microsoft-graph-client";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { GraphService } from "../../services/graph.js";
import { formatSearchHits, registerSearchTools } from "../search.js";

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

// Helper to build a standard search response with hits
function makeSearchResponse(hits: any[], total = hits.length, moreResultsAvailable = false) {
  return {
    value: [{ hitsContainers: [{ hits, total, moreResultsAvailable }] }],
  };
}

// Reusable hit fixture
function makeHit(overrides: Record<string, any> = {}) {
  return {
    hitId: "hit1",
    rank: 1,
    summary: "Found message",
    resource: {
      "@odata.type": "#microsoft.graph.chatMessage",
      id: "msg1",
      body: { content: "Hello world" },
      from: { user: { displayName: "John Doe", id: "user1" } },
      createdDateTime: "2025-01-01T10:00:00Z",
      chatId: "chat123",
      importance: "normal",
      webLink: "https://teams.microsoft.com/msg1",
      ...overrides,
    },
  };
}

describe("Search Tools", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockGraphService.getClient = vi.fn().mockResolvedValue(mockClient);
  });

  describe("registerSearchTools", () => {
    it("should register search_messages and get_my_mentions", () => {
      registerSearchTools(mockServer, mockGraphService, false);

      expect(mockServer.tool).toHaveBeenCalledTimes(2);
      expect(mockServer.tool).toHaveBeenCalledWith(
        "search_messages",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "get_my_mentions",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
    });
  });

  describe("formatSearchHits", () => {
    it("should map hit fields to a flat shape", () => {
      const hits = [
        makeHit({
          channelIdentity: { teamId: "team1", channelId: "channel1" },
        }),
      ];

      const results = formatSearchHits(hits as any);

      expect(results).toHaveLength(1);
      expect(results[0]).toEqual({
        id: "msg1",
        summary: "Found message",
        rank: 1,
        content: "Hello world",
        from: "John Doe",
        fromUserId: "user1",
        createdDateTime: "2025-01-01T10:00:00Z",
        importance: "normal",
        webLink: "https://teams.microsoft.com/msg1",
        chatId: "chat123",
        teamId: "team1",
        channelId: "channel1",
      });
    });

    it("should handle missing optional fields gracefully", () => {
      const hits = [
        {
          hitId: "hit1",
          rank: 1,
          summary: "",
          resource: {
            "@odata.type": "#microsoft.graph.chatMessage",
            id: "msg1",
          },
        },
      ];

      const results = formatSearchHits(hits as any);

      expect(results[0].content).toBeUndefined();
      expect(results[0].from).toBeUndefined();
      expect(results[0].webLink).toBeUndefined();
    });
  });

  describe("search_messages", () => {
    let handler: (args: any) => Promise<any>;

    beforeEach(() => {
      registerSearchTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "search_messages");
      handler = call?.[3] as unknown as (args: any) => Promise<any>;
    });

    it("should send a single search request with provided parameters", async () => {
      const mockApiChain = {
        post: vi.fn().mockResolvedValue(makeSearchResponse([makeHit()])),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ query: "hello", from: 0, size: 25, enableTopResults: true });

      expect(mockClient.api).toHaveBeenCalledWith("/search/query");
      expect(mockApiChain.post).toHaveBeenCalledWith({
        requests: [
          {
            entityTypes: ["chatMessage"],
            query: { queryString: "hello" },
            from: 0,
            size: 25,
            enableTopResults: true,
          },
        ],
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.query).toBe("hello");
      expect(parsed.total).toBe(1);
      expect(parsed.results).toHaveLength(1);
      expect(parsed.results[0].content).toBe("Hello world");
    });

    it("should pass KQL query strings through unmodified", async () => {
      const mockApiChain = {
        post: vi.fn().mockResolvedValue(makeSearchResponse([makeHit()])),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      await handler({
        query: "from:bob hasAttachment:true sent>=2025-01-01",
      });

      const postCall = mockApiChain.post.mock.calls[0][0];
      expect(postCall.requests[0].query.queryString).toBe(
        "from:bob hasAttachment:true sent>=2025-01-01"
      );
    });

    it("should support pagination with from offset", async () => {
      const mockApiChain = {
        post: vi.fn().mockResolvedValue(makeSearchResponse([makeHit()], 50, true)),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ query: "test", from: 25, size: 25 });

      const postCall = mockApiChain.post.mock.calls[0][0];
      expect(postCall.requests[0].from).toBe(25);
      expect(postCall.requests[0].size).toBe(25);

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.moreResultsAvailable).toBe(true);
      expect(parsed.from).toBe(25);
    });

    it("should return friendly message when no results found", async () => {
      const mockApiChain = {
        post: vi.fn().mockResolvedValue({ value: [] }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ query: "nonexistent" });
      expect(result.content[0].text).toBe("No messages found matching your search criteria.");
    });

    it("should handle empty hitsContainers", async () => {
      const mockApiChain = {
        post: vi.fn().mockResolvedValue({ value: [{ hitsContainers: [] }] }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ query: "nonexistent" });
      expect(result.content[0].text).toBe("No messages found matching your search criteria.");
    });

    it("should return error message on API failure", async () => {
      const mockApiChain = {
        post: vi.fn().mockRejectedValue(new Error("Search API error")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ query: "error" });
      expect(result.content[0].text).toBe("❌ Error searching messages: Search API error");
    });

    it("should handle non-Error thrown values", async () => {
      const mockApiChain = {
        post: vi.fn().mockRejectedValue("string error"),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ query: "error" });
      expect(result.content[0].text).toBe("❌ Error searching messages: Unknown error occurred");
    });
  });

  describe("get_my_mentions", () => {
    let handler: (args: any) => Promise<any>;

    beforeEach(() => {
      registerSearchTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "get_my_mentions");
      handler = call?.[3] as unknown as (args: any) => Promise<any>;
    });

    it("should query with IsMentioned:true and date filter", async () => {
      const mockUser = { id: "currentuser123", displayName: "Current User" };
      const mockApiChain = {
        get: vi.fn().mockResolvedValue(mockUser),
        post: vi.fn().mockResolvedValue(makeSearchResponse([makeHit()])),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ hours: 24, size: 25 });

      expect(mockClient.api).toHaveBeenCalledWith("/me");
      expect(mockClient.api).toHaveBeenCalledWith("/search/query");

      // Verify the KQL query uses IsMentioned:true
      const postCall = mockApiChain.post.mock.calls[0][0];
      expect(postCall.requests[0].query.queryString).toContain("IsMentioned:true");
      expect(postCall.requests[0].query.queryString).toContain("sent>=");
      expect(postCall.requests[0].enableTopResults).toBe(false);

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.mentionedUser).toBe("Current User");
      expect(parsed.mentions).toHaveLength(1);
      expect(parsed.total).toBe(1);
    });

    it("should return friendly message when no mentions found", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ id: "user1" }),
        post: vi.fn().mockResolvedValue({ value: [] }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ hours: 24 });
      expect(result.content[0].text).toBe("No recent mentions found.");
    });

    it("should return friendly message when hits array is empty", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ id: "user1" }),
        post: vi.fn().mockResolvedValue(makeSearchResponse([])),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ hours: 24 });
      expect(result.content[0].text).toBe("No recent mentions found.");
    });

    it("should error when current user ID cannot be resolved", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue({}),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ hours: 24 });
      expect(result.content[0].text).toBe("❌ Error: Could not determine current user ID");
    });

    it("should return error message on API failure", async () => {
      const mockApiChain = {
        get: vi.fn().mockRejectedValue(new Error("User lookup failed")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await handler({ hours: 24 });
      expect(result.content[0].text).toBe("❌ Error getting mentions: User lookup failed");
    });

    it("should respect hours parameter for date calculation", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ id: "user1" }),
        post: vi.fn().mockResolvedValue(makeSearchResponse([makeHit()])),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      await handler({ hours: 168, size: 10 });

      const postCall = mockApiChain.post.mock.calls[0][0];
      // 168 hours = 7 days ago
      const expectedDate = new Date(Date.now() - 168 * 60 * 60 * 1000).toISOString().split("T")[0];
      expect(postCall.requests[0].query.queryString).toContain(`sent>=${expectedDate}`);
      expect(postCall.requests[0].size).toBe(10);
    });
  });
});

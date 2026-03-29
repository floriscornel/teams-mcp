import { HttpResponse, http } from "msw";
import { setupServer } from "msw/node";
import { afterEach, beforeEach, expect, vi } from "vitest";
import type {
  Channel,
  Chat,
  ChatMessage,
  ConversationMember,
  GraphApiResponse,
  Team,
  User,
} from "../types/graph.js";

// Mock data fixtures
export const mockUser: User = {
  id: "test-user-id",
  displayName: "Test User",
  userPrincipalName: "test.user@example.com",
  mail: "test.user@example.com",
  jobTitle: "Test Engineer",
  department: "Engineering",
  officeLocation: "Remote",
};

export const mockTeam: Team = {
  id: "test-team-id",
  displayName: "Test Team",
  description: "A test team for unit tests",
  isArchived: false,
};

export const mockChannel: Channel = {
  id: "test-channel-id",
  displayName: "General",
  description: "General discussion channel",
  membershipType: "standard",
};

export const mockChat: Chat = {
  id: "test-chat-id",
  topic: "Test Chat",
  chatType: "group",
};

export const mockChatMessage: ChatMessage = {
  id: "test-message-id",
  createdDateTime: "2024-01-01T12:00:00Z",
  body: {
    content: "Test message content",
    contentType: "text",
  },
  from: {
    user: {
      id: "test-user-id",
      displayName: "Test User",
    },
  },
  importance: "normal",
};

export const mockConversationMember: ConversationMember = {
  id: "test-member-id",
  displayName: "Test Member",
  roles: ["owner"],
};

// Microsoft Graph API mock handlers
export const graphApiHandlers = [
  // User endpoints
  http.get("https://graph.microsoft.com/v1.0/me", () => {
    return HttpResponse.json(mockUser);
  }),

  http.get("https://graph.microsoft.com/v1.0/users", ({ request }) => {
    const url = new URL(request.url);
    const filter = url.searchParams.get("$filter");

    // Simulate search functionality
    const response: GraphApiResponse<User> = {
      value: filter?.includes("test") ? [mockUser] : [],
    };
    return HttpResponse.json(response);
  }),

  http.get("https://graph.microsoft.com/v1.0/users/:userId", ({ params }) => {
    if (params.userId === "test-user-id" || params.userId === "test.user@example.com") {
      return HttpResponse.json(mockUser);
    }
    return new HttpResponse(null, { status: 404 });
  }),

  // Teams endpoints
  http.get("https://graph.microsoft.com/v1.0/me/joinedTeams", () => {
    const response: GraphApiResponse<Team> = {
      value: [mockTeam],
    };
    return HttpResponse.json(response);
  }),

  http.get("https://graph.microsoft.com/v1.0/teams/:teamId/channels", ({ params }) => {
    if (params.teamId === "test-team-id") {
      const response: GraphApiResponse<Channel> = {
        value: [mockChannel],
      };
      return HttpResponse.json(response);
    }
    return new HttpResponse(null, { status: 404 });
  }),

  http.get("https://graph.microsoft.com/v1.0/teams/:teamId/members", ({ params }) => {
    if (params.teamId === "test-team-id") {
      const response: GraphApiResponse<ConversationMember> = {
        value: [mockConversationMember],
      };
      return HttpResponse.json(response);
    }
    return new HttpResponse(null, { status: 404 });
  }),

  // Channel messages
  http.get(
    "https://graph.microsoft.com/v1.0/teams/:teamId/channels/:channelId/messages",
    ({ params }) => {
      if (params.teamId === "test-team-id" && params.channelId === "test-channel-id") {
        const response: GraphApiResponse<ChatMessage> = {
          value: [mockChatMessage],
        };
        return HttpResponse.json(response);
      }
      return new HttpResponse(null, { status: 404 });
    }
  ),

  http.post(
    "https://graph.microsoft.com/v1.0/teams/:teamId/channels/:channelId/messages",
    async ({ params, request }) => {
      if (params.teamId === "test-team-id" && params.channelId === "test-channel-id") {
        const body = (await request.json()) as any;
        const response = {
          ...mockChatMessage,
          id: "new-message-id",
          body: body.body,
          createdDateTime: new Date().toISOString(),
        };
        return HttpResponse.json(response);
      }
      return new HttpResponse(null, { status: 404 });
    }
  ),

  // Chats endpoints
  http.get("https://graph.microsoft.com/v1.0/me/chats", () => {
    const response: GraphApiResponse<Chat> = {
      value: [mockChat],
    };
    return HttpResponse.json(response);
  }),

  http.get("https://graph.microsoft.com/v1.0/me/chats/:chatId/messages", ({ params, request }) => {
    if (params.chatId === "test-chat-id") {
      const url = new URL(request.url);
      const fromUser = url.searchParams.get("$filter")?.includes("from/user/id");

      const response: GraphApiResponse<ChatMessage> = {
        value: fromUser ? [] : [mockChatMessage],
      };
      return HttpResponse.json(response);
    }
    return new HttpResponse(null, { status: 404 });
  }),

  http.post(
    "https://graph.microsoft.com/v1.0/me/chats/:chatId/messages",
    async ({ params, request }) => {
      if (params.chatId === "test-chat-id") {
        const body = (await request.json()) as any;
        const response = {
          ...mockChatMessage,
          id: "new-chat-message-id",
          body: body.body,
          createdDateTime: new Date().toISOString(),
        };
        return HttpResponse.json(response);
      }
      return new HttpResponse(null, { status: 404 });
    }
  ),

  http.post("https://graph.microsoft.com/v1.0/me/chats", async ({ request }) => {
    const body = (await request.json()) as any;
    const response = {
      ...mockChat,
      id: "new-chat-id",
      topic: body.topic,
      chatType: body.chatType,
    };
    return HttpResponse.json(response);
  }),

  // Search endpoints
  http.post("https://graph.microsoft.com/v1.0/search/query", async ({ request }) => {
    const body = (await request.json()) as any;
    const searchRequest = body.requests[0];

    const response = {
      value: [
        {
          searchTerms: [searchRequest.query.queryString],
          hitsContainers: [
            {
              hits: [
                {
                  hitId: "search-hit-1",
                  rank: 1,
                  summary: "Test message found in search",
                  resource: {
                    "@odata.type": "#microsoft.graph.chatMessage",
                    id: "search-message-id",
                    createdDateTime: "2024-01-01T12:00:00Z",
                    from: {
                      user: {
                        displayName: "Test User",
                        id: "test-user-id",
                      },
                    },
                    body: {
                      content: "Test search result message",
                      contentType: "text",
                    },
                    chatId: "test-chat-id",
                  },
                },
              ],
              total: 1,
              moreResultsAvailable: false,
            },
          ],
        },
      ],
    };
    return HttpResponse.json(response);
  }),

  // Error scenarios for testing
  http.get("https://graph.microsoft.com/v1.0/error/401", () => {
    return HttpResponse.json(
      {
        error: {
          code: "InvalidAuthenticationToken",
          message: "Access token is empty.",
        },
      },
      { status: 401 }
    );
  }),

  http.get("https://graph.microsoft.com/v1.0/error/403", () => {
    return HttpResponse.json(
      {
        error: {
          code: "Forbidden",
          message: "Insufficient privileges to complete the operation.",
        },
      },
      { status: 403 }
    );
  }),

  http.get("https://graph.microsoft.com/v1.0/error/429", () => {
    return HttpResponse.json(
      {
        error: {
          code: "TooManyRequests",
          message: "Too many requests",
        },
      },
      { status: 429, headers: { "Retry-After": "30" } }
    );
  }),
];

// Setup MSW server
export const server = setupServer(...graphApiHandlers);

// Global test setup
beforeEach(() => {
  // Reset all mocks before each test
  vi.clearAllMocks();

  // Mock file system operations for token storage
  vi.mock("node:fs", async () => {
    const actual = (await vi.importActual("node:fs")) as any;
    return {
      ...actual,
      promises: {
        ...(actual.promises || {}),
        readFile: vi.fn(),
        writeFile: vi.fn(),
        unlink: vi.fn(),
        access: vi.fn(),
      },
    };
  });

  // Mock Azure identity
  vi.mock("@azure/identity", () => ({
    DeviceCodeCredential: vi.fn(),
  }));
});

afterEach(() => {
  // Clean up after each test
  vi.resetAllMocks();
});

// Helper function to create mock authenticated GraphService
export function createMockGraphService() {
  return {
    getInstance: vi.fn().mockReturnThis(),
    getAuthStatus: vi.fn().mockResolvedValue({
      isAuthenticated: true,
      userPrincipalName: mockUser.userPrincipalName,
      displayName: mockUser.displayName,
      expiresAt: new Date(Date.now() + 3600000).toISOString(),
    }),
    getClient: vi.fn().mockResolvedValue({
      api: vi.fn().mockReturnValue({
        get: vi.fn(),
        post: vi.fn(),
        filter: vi.fn().mockReturnThis(),
      }),
    }),
    isAuthenticated: vi.fn().mockReturnValue(true),
  };
}

// Helper function to create mock unauthenticated GraphService
export function createMockUnauthenticatedGraphService() {
  return {
    getInstance: vi.fn().mockReturnThis(),
    getAuthStatus: vi.fn().mockResolvedValue({
      isAuthenticated: false,
    }),
    getClient: vi.fn().mockRejectedValue(new Error("Not authenticated")),
    isAuthenticated: vi.fn().mockReturnValue(false),
  };
}

// Helper function to create mock MCP server
export function createMockMcpServer() {
  const tools = new Map();

  return {
    tool: vi.fn().mockImplementation((name, description, schema, handler) => {
      tools.set(name, { description, schema, handler });
    }),
    registerTool: vi.fn().mockImplementation((name, config, handler) => {
      tools.set(name, {
        description: config.description,
        schema: config.inputSchema,
        handler,
        title: config.title,
        annotations: config.annotations,
      });
    }),
    connect: vi.fn(),
    getTool: (name: string) => tools.get(name),
    getAllTools: () => Array.from(tools.keys()),
  };
}

// Helper function to test MCP tool execution
export async function testMcpTool(
  toolName: string,
  parameters: any,
  mockServer: any,
  expectedResult?: any
) {
  const tool = mockServer.getTool(toolName);
  if (!tool) {
    throw new Error(`Tool ${toolName} not found`);
  }

  const result = await tool.handler(parameters);

  if (expectedResult) {
    expect(result).toEqual(expectedResult);
  }

  return result;
}

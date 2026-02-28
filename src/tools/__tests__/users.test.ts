import { beforeEach, describe, expect, it, vi } from "vitest";
import { createMockMcpServer, mockUser } from "../../test-utils/setup.js";
import type { GraphApiResponse, User } from "../../types/graph.js";
import { registerUsersTools } from "../users.js";

describe("Users Tools", () => {
  let mockServer: any;
  let mockGraphService: any;
  let mockClient: any;

  beforeEach(() => {
    mockServer = createMockMcpServer();
    mockClient = {
      api: vi.fn().mockReturnValue({
        get: vi.fn(),
        filter: vi.fn().mockReturnThis(),
      }),
    };

    mockGraphService = {
      getClient: vi.fn().mockResolvedValue(mockClient),
    };

    vi.clearAllMocks();
  });

  describe("get_current_user tool", () => {
    it("should register get_current_user tool correctly", () => {
      registerUsersTools(mockServer, mockGraphService, false);

      expect(mockServer.tool).toHaveBeenCalledWith(
        "get_current_user",
        "Get the current authenticated user's profile information including display name, email, job title, and department.",
        {},
        expect.any(Function)
      );
    });

    it("should return current user information", async () => {
      mockClient.api().get.mockResolvedValue(mockUser);
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("get_current_user");
      const result = await tool.handler();

      expect(mockClient.api).toHaveBeenCalledWith("/me");
      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                displayName: mockUser.displayName,
                userPrincipalName: mockUser.userPrincipalName,
                mail: mockUser.mail,
                id: mockUser.id,
                jobTitle: mockUser.jobTitle,
                department: mockUser.department,
              },
              null,
              2
            ),
          },
        ],
      });
    });

    it("should handle API errors gracefully", async () => {
      mockClient.api().get.mockRejectedValue(new Error("API Error"));
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("get_current_user");
      const result = await tool.handler();

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "❌ Error: API Error",
          },
        ],
      });
    });

    it("should handle unknown errors", async () => {
      mockClient.api().get.mockRejectedValue("String error");
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("get_current_user");
      const result = await tool.handler();

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "❌ Error: Unknown error occurred",
          },
        ],
      });
    });
  });

  describe("search_users tool", () => {
    it("should register search_users tool with correct schema", () => {
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("search_users");
      expect(tool).toBeDefined();
      expect(tool.schema.query).toBeDefined();
    });

    it("should search users successfully", async () => {
      const searchResponse: GraphApiResponse<User> = {
        value: [mockUser],
      };

      mockClient.api().get.mockResolvedValue(searchResponse);
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("search_users");
      const result = await tool.handler({ query: "test" });

      expect(mockClient.api).toHaveBeenCalledWith("/users");
      expect(mockClient.api().filter).toHaveBeenCalledWith(
        "startswith(displayName,'test') or startswith(mail,'test') or startswith(userPrincipalName,'test')"
      );

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: JSON.stringify(
              [
                {
                  displayName: mockUser.displayName,
                  userPrincipalName: mockUser.userPrincipalName,
                  mail: mockUser.mail,
                  id: mockUser.id,
                },
              ],
              null,
              2
            ),
          },
        ],
      });
    });

    it("should handle empty search results", async () => {
      const emptyResponse: GraphApiResponse<User> = {
        value: [],
      };

      mockClient.api().get.mockResolvedValue(emptyResponse);
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("search_users");
      const result = await tool.handler({ query: "nonexistent" });

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "No users found matching your search.",
          },
        ],
      });
    });

    it("should handle undefined value in response", async () => {
      const undefinedResponse: GraphApiResponse<User> = {};

      mockClient.api().get.mockResolvedValue(undefinedResponse);
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("search_users");
      const result = await tool.handler({ query: "test" });

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "No users found matching your search.",
          },
        ],
      });
    });

    it("should handle search API errors", async () => {
      mockClient.api().get.mockRejectedValue(new Error("Search failed"));
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("search_users");
      const result = await tool.handler({ query: "test" });

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "❌ Error: Search failed",
          },
        ],
      });
    });

    it("should properly escape special characters in search query", async () => {
      const searchResponse: GraphApiResponse<User> = { value: [] };
      mockClient.api().get.mockResolvedValue(searchResponse);
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("search_users");
      await tool.handler({ query: "test@example.com" });

      expect(mockClient.api().filter).toHaveBeenCalledWith(
        "startswith(displayName,'test@example.com') or startswith(mail,'test@example.com') or startswith(userPrincipalName,'test@example.com')"
      );
    });
  });

  describe("get_user tool", () => {
    it("should register get_user tool with correct schema", () => {
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("get_user");
      expect(tool).toBeDefined();
      expect(tool.schema.userId).toBeDefined();
    });

    it("should get user by ID successfully", async () => {
      mockClient.api().get.mockResolvedValue(mockUser);
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("get_user");
      const result = await tool.handler({ userId: "test-user-id" });

      expect(mockClient.api).toHaveBeenCalledWith("/users/test-user-id");
      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                displayName: mockUser.displayName,
                userPrincipalName: mockUser.userPrincipalName,
                mail: mockUser.mail,
                id: mockUser.id,
                jobTitle: mockUser.jobTitle,
                department: mockUser.department,
                officeLocation: mockUser.officeLocation,
              },
              null,
              2
            ),
          },
        ],
      });
    });

    it("should get user by email successfully", async () => {
      mockClient.api().get.mockResolvedValue(mockUser);
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("get_user");
      const result = await tool.handler({ userId: "test.user@example.com" });

      expect(mockClient.api).toHaveBeenCalledWith("/users/test.user@example.com");
      expect(result.content[0].text).toContain(mockUser.displayName);
    });

    it("should handle user not found error", async () => {
      const notFoundError = new Error("User not found");
      mockClient.api().get.mockRejectedValue(notFoundError);
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("get_user");
      const result = await tool.handler({ userId: "nonexistent-user" });

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "❌ Error: User not found",
          },
        ],
      });
    });

    it("should handle partial user data", async () => {
      const partialUser: User = {
        id: "test-id",
        displayName: "Test User",
        // Missing other fields
      };

      mockClient.api().get.mockResolvedValue(partialUser);
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("get_user");
      const result = await tool.handler({ userId: "test-id" });

      const expectedUserSummary = {
        displayName: "Test User",
        userPrincipalName: undefined,
        mail: undefined,
        id: "test-id",
        jobTitle: undefined,
        department: undefined,
        officeLocation: undefined,
      };

      expect(result.content[0].text).toBe(JSON.stringify(expectedUserSummary, null, 2));
    });
  });

  describe("authentication errors", () => {
    it("should handle authentication errors in all tools", async () => {
      const authError = new Error("Not authenticated");
      mockGraphService.getClient.mockRejectedValue(authError);
      registerUsersTools(mockServer, mockGraphService, false);

      const tools = ["get_current_user", "search_users", "get_user"];

      for (const toolName of tools) {
        const tool = mockServer.getTool(toolName);
        const params =
          toolName === "search_users"
            ? { query: "test" }
            : toolName === "get_user"
              ? { userId: "test" }
              : {};

        const result = await tool.handler(params);
        expect(result.content[0].text).toContain("❌ Error: Not authenticated");
      }
    });
  });

  describe("input validation", () => {
    it("should handle empty search query", async () => {
      mockClient.api().get.mockResolvedValue({ value: [] });
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("search_users");
      const result = await tool.handler({ query: "" });

      expect(mockClient.api().filter).toHaveBeenCalledWith(
        "startswith(displayName,'') or startswith(mail,'') or startswith(userPrincipalName,'')"
      );
      expect(result.content[0].text).toBe("No users found matching your search.");
    });

    it("should handle empty userId", async () => {
      registerUsersTools(mockServer, mockGraphService, false);

      const tool = mockServer.getTool("get_user");
      const _result = await tool.handler({ userId: "" });

      expect(mockClient.api).toHaveBeenCalledWith("/users/");
    });
  });
});

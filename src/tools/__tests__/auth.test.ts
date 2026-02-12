import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  createMockGraphService,
  createMockMcpServer,
  createMockUnauthenticatedGraphService,
} from "../../test-utils/setup.js";
import { registerAuthTools } from "../auth.js";

describe("Authentication Tools", () => {
  let mockServer: any;
  let mockGraphService: any;

  beforeEach(() => {
    mockServer = createMockMcpServer();
    vi.clearAllMocks();
  });

  describe("auth_status tool", () => {
    it("should register auth_status tool correctly", () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService);

      expect(mockServer.tool).toHaveBeenCalledWith(
        "auth_status",
        "Check the authentication status of the Microsoft Graph connection. Returns whether the user is authenticated and shows their basic profile information.",
        {},
        expect.any(Function)
      );
    });

    it("should return authenticated status when user is authenticated", async () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService);

      const authTool = mockServer.getTool("auth_status");
      const result = await authTool.handler();

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "✅ Authenticated as Test User (test.user@example.com)",
          },
        ],
      });

      expect(mockGraphService.getAuthStatus).toHaveBeenCalledTimes(1);
    });

    it("should return unauthenticated status when user is not authenticated", async () => {
      mockGraphService = createMockUnauthenticatedGraphService();
      registerAuthTools(mockServer, mockGraphService);

      const authTool = mockServer.getTool("auth_status");
      const result = await authTool.handler();

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "❌ Not authenticated. Please run: npx @floriscornel/teams-mcp@latest authenticate",
          },
        ],
      });

      expect(mockGraphService.getAuthStatus).toHaveBeenCalledTimes(1);
    });

    it("should handle partial authentication data gracefully", async () => {
      const partialMockGraphService = {
        getAuthStatus: vi.fn().mockResolvedValue({
          isAuthenticated: true,
          displayName: "Test User",
          // Missing userPrincipalName
        }),
      } as any;

      registerAuthTools(mockServer, partialMockGraphService);

      const authTool = mockServer.getTool("auth_status");
      const result = await authTool.handler();

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "✅ Authenticated as Test User (No email available)",
          },
        ],
      });
    });

    it("should handle authentication status errors", async () => {
      const errorMockGraphService = {
        getAuthStatus: vi.fn().mockRejectedValue(new Error("Auth check failed")),
      } as any;

      registerAuthTools(mockServer, errorMockGraphService);

      const authTool = mockServer.getTool("auth_status");

      // Should throw the error since it's not caught in the tool
      await expect(authTool.handler()).rejects.toThrow("Auth check failed");
    });

    it("should handle null/undefined user data", async () => {
      const nullDataMockGraphService = {
        getAuthStatus: vi.fn().mockResolvedValue({
          isAuthenticated: true,
          displayName: null,
          userPrincipalName: null,
        }),
      } as any;

      registerAuthTools(mockServer, nullDataMockGraphService);

      const authTool = mockServer.getTool("auth_status");
      const result = await authTool.handler();

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "✅ Authenticated as Unknown User (No email available)",
          },
        ],
      });
    });
  });

  describe("tool registration", () => {
    it("should register all expected authentication tools", () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService);

      const registeredTools = mockServer.getAllTools();
      expect(registeredTools).toContain("auth_status");
      expect(registeredTools).toHaveLength(1);
    });

    it("should handle GraphService being undefined", () => {
      expect(() => {
        registerAuthTools(mockServer, undefined as any);
      }).not.toThrow();

      // Tool should still be registered
      expect(mockServer.tool).toHaveBeenCalledWith(
        "auth_status",
        expect.any(String),
        {},
        expect.any(Function)
      );
    });
  });

  describe("read-only mode indicator", () => {
    it("should include [Read-Only Mode] when readOnly is true", async () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService, true);

      const authTool = mockServer.getTool("auth_status");
      const result = await authTool.handler();

      expect(result.content[0].text).toContain("[Read-Only Mode]");
      expect(result.content[0].text).toContain("✅ Authenticated as Test User");
    });

    it("should NOT include [Read-Only Mode] when readOnly is false", async () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService, false);

      const authTool = mockServer.getTool("auth_status");
      const result = await authTool.handler();

      expect(result.content[0].text).not.toContain("[Read-Only Mode]");
      expect(result.content[0].text).toContain("✅ Authenticated as Test User");
    });

    it("should NOT include [Read-Only Mode] when readOnly is not passed (default)", async () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService);

      const authTool = mockServer.getTool("auth_status");
      const result = await authTool.handler();

      expect(result.content[0].text).not.toContain("[Read-Only Mode]");
    });
  });

  describe("authentication state changes", () => {
    it("should reflect real-time authentication status changes", async () => {
      // Start with unauthenticated state
      let isAuthenticated = false;
      const dynamicMockGraphService = {
        getAuthStatus: vi.fn().mockImplementation(() => {
          return Promise.resolve({
            isAuthenticated,
            displayName: isAuthenticated ? "Test User" : undefined,
            userPrincipalName: isAuthenticated ? "test.user@example.com" : undefined,
          });
        }),
      } as any;

      registerAuthTools(mockServer, dynamicMockGraphService);
      const authTool = mockServer.getTool("auth_status");

      // Check unauthenticated status
      let result = await authTool.handler();
      expect(result.content[0].text).toContain("❌ Not authenticated");

      // Simulate authentication
      isAuthenticated = true;

      // Check authenticated status
      result = await authTool.handler();
      expect(result.content[0].text).toContain("✅ Authenticated as Test User");
    });
  });
});

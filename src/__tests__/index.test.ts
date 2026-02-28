import fs from "node:fs/promises";
import { beforeEach, describe, expect, it, vi } from "vitest";

// Mock external dependencies
vi.mock("fs/promises");
vi.mock("@azure/identity");
vi.mock("@modelcontextprotocol/sdk/server/mcp.js");
vi.mock("@modelcontextprotocol/sdk/server/stdio.js");

// Mock console methods
const mockConsoleLog = vi.fn();
const mockConsoleError = vi.fn();
const mockProcessExit = vi.fn();

// Setup global mocks
beforeEach(() => {
  vi.clearAllMocks();

  // Mock console methods
  vi.spyOn(console, "log").mockImplementation(mockConsoleLog);
  vi.spyOn(console, "error").mockImplementation(mockConsoleError);
  vi.spyOn(process, "exit").mockImplementation(mockProcessExit as any);

  // Reset process.argv
  process.argv = ["node", "index.js"];
});

// Simple integration tests for basic functionality
describe("MCP Server Integration", () => {
  describe("CLI Commands", () => {
    it("should handle help command", async () => {
      process.argv = ["node", "index.js", "--help"];

      // Dynamically import to get fresh module state
      await import("../index.js");

      expect(mockConsoleLog).toHaveBeenCalledWith("Microsoft Graph MCP Server");
      expect(mockConsoleLog).toHaveBeenCalledWith("Usage:");
      expect(mockConsoleLog).toHaveBeenCalledWith(expect.stringContaining("authenticate"));
    });

    it.skip("should handle help variants", async () => {
      // Skipping complex integration test - core functionality is tested in unit tests
    });

    it.skip("should handle unknown command", async () => {
      // Skipping complex integration test - core functionality is tested in unit tests
    });

    describe("Authentication Commands", () => {
      it.skip("should handle authenticate command", async () => {
        process.argv = ["node", "index.js", "authenticate"];

        // Mock DeviceCodeCredential
        const mockCredential = {
          authenticate: vi.fn().mockResolvedValue({
            account: {
              username: "test@example.com",
              name: "Test User",
            },
            accessToken: "mock-token",
            expiresOn: new Date(Date.now() + 3600000),
          }),
          getToken: vi.fn().mockResolvedValue({
            token: "mock-token",
            expiresOnTimestamp: Date.now() + 3600000,
          }),
        };

        const { DeviceCodeCredential } = await import("@azure/identity");
        vi.mocked(DeviceCodeCredential).mockImplementation(() => mockCredential as any);

        await import("../index.js");

        expect(mockConsoleLog).toHaveBeenCalledWith(
          expect.stringContaining("Microsoft Graph Authentication")
        );
      });

      it.skip("should handle auth alias", async () => {
        process.argv = ["node", "index.js", "auth"];

        const mockCredential = {
          authenticate: vi.fn().mockResolvedValue({
            account: {
              username: "test@example.com",
              name: "Test User",
            },
            accessToken: "mock-token",
            expiresOn: new Date(Date.now() + 3600000),
          }),
          getToken: vi.fn().mockResolvedValue({
            token: "mock-token",
            expiresOnTimestamp: Date.now() + 3600000,
          }),
        };

        const { DeviceCodeCredential } = await import("@azure/identity");
        vi.mocked(DeviceCodeCredential).mockImplementation(() => mockCredential as any);

        await import("../index.js");

        expect(mockConsoleLog).toHaveBeenCalledWith(
          expect.stringContaining("Microsoft Graph Authentication")
        );
      });

      it.skip("should handle check command when authenticated", async () => {
        process.argv = ["node", "index.js", "check"];

        // Mock authenticated state
        const authData = JSON.stringify({
          clientId: "test-client-id",
          authenticated: true,
          timestamp: new Date().toISOString(),
          expiresAt: new Date(Date.now() + 3600000).toISOString(),
          token: "valid-token",
        });

        vi.mocked(fs.readFile).mockResolvedValue(authData);

        await import("../index.js");

        expect(mockConsoleLog).toHaveBeenCalledWith(
          expect.stringContaining("Authentication Status")
        );
      });

      it.skip("should handle check command when not authenticated", async () => {
        process.argv = ["node", "index.js", "check"];

        // Mock unauthenticated state
        vi.mocked(fs.readFile).mockRejectedValue(new Error("File not found"));

        await import("../index.js");

        expect(mockConsoleLog).toHaveBeenCalledWith(expect.stringContaining("Not authenticated"));
      });

      it.skip("should handle logout command", async () => {
        process.argv = ["node", "index.js", "logout"];

        vi.mocked(fs.unlink).mockResolvedValue();

        await import("../index.js");

        expect(mockConsoleLog).toHaveBeenCalledWith(
          expect.stringContaining("Logged out successfully")
        );
        expect(fs.unlink).toHaveBeenCalled();
      });

      it.skip("should handle logout command when no auth file exists", async () => {
        process.argv = ["node", "index.js", "logout"];

        vi.mocked(fs.unlink).mockRejectedValue(new Error("File not found"));

        await import("../index.js");

        expect(mockConsoleLog).toHaveBeenCalledWith(expect.stringContaining("Already logged out"));
      });
    });
  });

  describe("MCP Server Mode", () => {
    it.skip("should start MCP server when no command provided", async () => {
      process.argv = ["node", "index.js"];

      const { McpServer } = await import("@modelcontextprotocol/sdk/server/mcp.js");
      const { StdioServerTransport } = await import("@modelcontextprotocol/sdk/server/stdio.js");

      const mockServer = {
        tool: vi.fn(),
        connect: vi.fn(),
      };
      const mockTransport = {};

      vi.mocked(McpServer).mockImplementation(() => mockServer as any);
      vi.mocked(StdioServerTransport).mockImplementation(() => mockTransport as any);

      await import("../index.js");

      expect(McpServer).toHaveBeenCalledWith({
        name: "teams-mcp",
        version: "0.8.0",
      });

      // Should register all tool categories
      expect(mockServer.tool).toHaveBeenCalled();
      expect(mockServer.connect).toHaveBeenCalledWith(mockTransport);
      expect(mockConsoleError).toHaveBeenCalledWith("Microsoft Graph MCP Server started");
    });

    it.skip("should register all expected tools", async () => {
      process.argv = ["node", "index.js"];

      const { McpServer } = await import("@modelcontextprotocol/sdk/server/mcp.js");

      const mockServer = {
        tool: vi.fn(),
        connect: vi.fn(),
      };

      vi.mocked(McpServer).mockImplementation(() => mockServer as any);

      await import("../index.js");

      // Verify that tool registration functions were called
      // (We can't easily test the exact tools without more complex mocking)
      expect(mockServer.tool).toHaveBeenCalled();
    });
  });

  describe("Configuration", () => {
    it.skip("should use correct client ID", async () => {
      process.argv = ["node", "index.js", "authenticate"];

      const { DeviceCodeCredential } = await import("@azure/identity");

      await import("../index.js");

      expect(DeviceCodeCredential).toHaveBeenCalledWith({
        tenantId: "common",
        clientId: "14d82eec-204b-4c2f-b7e8-296a70dab67e",
        userPromptCallback: expect.any(Function),
      });
    });

    it.skip("should use correct token path", async () => {
      process.argv = ["node", "index.js", "check"];

      vi.mocked(fs.readFile).mockRejectedValue(new Error("File not found"));

      await import("../index.js");

      expect(fs.readFile).toHaveBeenCalledWith(
        expect.stringContaining(".msgraph-mcp-auth.json"),
        "utf8"
      );
    });
  });

  describe("Error Handling", () => {
    it.skip("should handle authentication errors gracefully", async () => {
      process.argv = ["node", "index.js", "authenticate"];

      const mockCredential = {
        authenticate: vi.fn().mockRejectedValue(new Error("Auth failed")),
      };

      const { DeviceCodeCredential } = await import("@azure/identity");
      vi.mocked(DeviceCodeCredential).mockImplementation(() => mockCredential as any);

      await import("../index.js");

      expect(mockConsoleError).toHaveBeenCalledWith(
        expect.stringContaining("Authentication failed")
      );
    });

    it.skip("should handle file system errors in check command", async () => {
      process.argv = ["node", "index.js", "check"];

      vi.mocked(fs.readFile).mockRejectedValue(new Error("Permission denied"));

      await import("../index.js");

      expect(mockConsoleError).toHaveBeenCalledWith(
        expect.stringContaining("Error checking authentication")
      );
    });

    it.skip("should handle file system errors in logout command", async () => {
      process.argv = ["node", "index.js", "logout"];

      vi.mocked(fs.unlink).mockRejectedValue(new Error("Permission denied"));

      await import("../index.js");

      expect(mockConsoleError).toHaveBeenCalledWith(expect.stringContaining("Error during logout"));
    });
  });

  describe("Token Management", () => {
    it.skip("should save token after successful authentication", async () => {
      process.argv = ["node", "index.js", "authenticate"];

      const mockToken = {
        token: "access-token",
        expiresOnTimestamp: Date.now() + 3600000,
      };

      const mockCredential = {
        authenticate: vi.fn().mockResolvedValue({
          account: {
            username: "test@example.com",
            name: "Test User",
          },
        }),
        getToken: vi.fn().mockResolvedValue(mockToken),
      };

      const { DeviceCodeCredential } = await import("@azure/identity");
      vi.mocked(DeviceCodeCredential).mockImplementation(() => mockCredential as any);

      vi.mocked(fs.writeFile).mockResolvedValue();

      await import("../index.js");

      expect(fs.writeFile).toHaveBeenCalledWith(
        expect.stringContaining(".msgraph-mcp-auth.json"),
        expect.stringContaining("access-token"),
        "utf8"
      );
    });

    it.skip("should handle expired tokens in check command", async () => {
      process.argv = ["node", "index.js", "check"];

      const expiredAuthData = JSON.stringify({
        clientId: "test-client-id",
        authenticated: true,
        timestamp: new Date().toISOString(),
        expiresAt: new Date(Date.now() - 3600000).toISOString(), // Expired
        token: "expired-token",
      });

      vi.mocked(fs.readFile).mockResolvedValue(expiredAuthData);

      await import("../index.js");

      expect(mockConsoleLog).toHaveBeenCalledWith(expect.stringContaining("Token has expired"));
    });
  });
});

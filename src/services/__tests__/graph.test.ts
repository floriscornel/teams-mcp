import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

// Create a mock credential instance
const mockCredentialInstance = {
  getToken: vi.fn().mockResolvedValue({
    token: "mock-token",
    expiresOnTimestamp: Date.now() + 3600000,
  }),
};

// Mock Azure Identity before any imports
vi.mock("@azure/identity", () => ({
  useIdentityPlugin: vi.fn(),
  DeviceCodeCredential: vi.fn().mockImplementation(() => mockCredentialInstance),
}));

vi.mock("@azure/identity-cache-persistence", () => ({
  cachePersistencePlugin: {},
}));

import { mockUser, server } from "../../test-utils/setup.js";
import { type AuthStatus, GraphService } from "../graph.js";

// Mock @microsoft/microsoft-graph-client
vi.mock("@microsoft/microsoft-graph-client", () => ({
  Client: {
    initWithMiddleware: vi.fn(),
  },
}));

describe("GraphService", () => {
  let graphService: GraphService;

  beforeEach(() => {
    // Start MSW server
    server.listen({ onUnhandledRequest: "error" });

    // Reset GraphService singleton
    (GraphService as any).instance = undefined;
    graphService = GraphService.getInstance();

    // Clear all mocks
    vi.clearAllMocks();
    
    // Reset the mock credential behavior
    mockCredentialInstance.getToken.mockResolvedValue({
      token: "mock-token",
      expiresOnTimestamp: Date.now() + 3600000,
    });
  });

  afterEach(() => {
    server.resetHandlers();
    server.close();
  });

  describe("getInstance", () => {
    it("should return singleton instance", () => {
      const instance1 = GraphService.getInstance();
      const instance2 = GraphService.getInstance();

      expect(instance1).toBe(instance2);
    });
  });

  describe("getAuthStatus", () => {
    it("should return authenticated status with valid setup", async () => {
      // Mock the Graph Client
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      const { Client } = await import("@microsoft/microsoft-graph-client");
      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const status = await graphService.getAuthStatus();

      expect(status.isAuthenticated).toBe(true);
      expect(status.userPrincipalName).toBe(mockUser.userPrincipalName);
      expect(status.displayName).toBe(mockUser.displayName);
      expect(status.expiresAt).toBeDefined();
    });

    it("should handle Graph API errors gracefully", async () => {
      // Mock the Graph Client to throw an error
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockRejectedValue(new Error("API Error")),
        }),
      };

      const { Client } = await import("@microsoft/microsoft-graph-client");
      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({
        isAuthenticated: false,
      });
    });

    it("should handle token acquisition errors gracefully", async () => {
      // Mock token acquisition failure
      mockCredentialInstance.getToken.mockRejectedValue(new Error("Token acquisition failed"));

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({
        isAuthenticated: false,
      });
    });

    it("should handle null token response", async () => {
      // Mock null token response
      mockCredentialInstance.getToken.mockResolvedValue(null);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({
        isAuthenticated: false,
      });
    });
  });

  describe("getClient", () => {
    it("should return client when service is initialized", async () => {
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      const { Client } = await import("@microsoft/microsoft-graph-client");
      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const client = await graphService.getClient();

      expect(client).toBe(mockClient);
    });

    it("should throw error when no token available", async () => {
      // Mock no token available
      mockCredentialInstance.getToken.mockResolvedValue(null);

      await expect(graphService.getClient()).rejects.toThrow(
        "Not authenticated. Please run the authentication CLI tool first"
      );
    });

    it("should return same client on multiple calls", async () => {
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      const { Client } = await import("@microsoft/microsoft-graph-client");
      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const client1 = await graphService.getClient();
      const client2 = await graphService.getClient();

      expect(client1).toBe(client2);
      expect(Client.initWithMiddleware).toHaveBeenCalledTimes(1);
    });
  });

  describe("isAuthenticated", () => {
    it("should return false when not initialized", () => {
      expect(graphService.isAuthenticated()).toBe(false);
    });

    it("should return true when client is initialized", async () => {
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      const { Client } = await import("@microsoft/microsoft-graph-client");
      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      // Initialize the client
      await graphService.getAuthStatus();

      expect(graphService.isAuthenticated()).toBe(true);
    });
  });

  describe("concurrent initialization", () => {
    it("should handle concurrent calls to initializeClient", async () => {
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      const { Client } = await import("@microsoft/microsoft-graph-client");
      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      // Make multiple concurrent calls
      const promises = [
        graphService.getAuthStatus(),
        graphService.getAuthStatus(),
        graphService.getAuthStatus(),
      ];

      const results = await Promise.all(promises);

      // All should return the same authenticated status
      for (const result of results) {
        expect(result.isAuthenticated).toBe(true);
      }

      // Client should only be initialized once
      expect(Client.initWithMiddleware).toHaveBeenCalledTimes(1);
    });
  });
});

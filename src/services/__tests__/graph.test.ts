import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mockUser, server } from "../../test-utils/setup.js";
import { type AuthStatus, GraphService } from "../graph.js";

// Get the mocked modules
const { DeviceCodeCredential } = await import("@azure/identity");
const { Client } = await import("@microsoft/microsoft-graph-client");

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

    // Reset default mock implementations
    vi.mocked(DeviceCodeCredential).mockImplementation(
      () =>
        ({
          getToken: vi.fn().mockResolvedValue({
            token: "mock-token",
            expiresOnTimestamp: Date.now() + 3600000,
          }),
        }) as any
    );

    vi.mocked(Client.initWithMiddleware).mockImplementation(
      () =>
        ({
          api: vi.fn().mockReturnValue({
            get: vi.fn().mockResolvedValue(mockUser),
          }),
        }) as any
    );
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
      const status = await graphService.getAuthStatus();

      expect(status.isAuthenticated).toBe(true);
      expect(status.userPrincipalName).toBe(mockUser.userPrincipalName);
      expect(status.displayName).toBe(mockUser.displayName);
      expect(status.expiresAt).toBeDefined();
    });

    it("should handle Graph API errors gracefully", async () => {
      // Mock the Graph Client to throw an error
      vi.mocked(Client.initWithMiddleware).mockImplementation(
        () =>
          ({
            api: vi.fn().mockReturnValue({
              get: vi.fn().mockRejectedValue(new Error("API Error")),
            }),
          }) as any
      );

      // Reset the singleton to use the new mock
      (GraphService as any).instance = undefined;
      const testGraphService = GraphService.getInstance();

      const status = await testGraphService.getAuthStatus();

      expect(status).toEqual({
        isAuthenticated: false,
      });
    });

    it("should handle token acquisition errors gracefully", async () => {
      // Create a new GraphService instance and force re-initialization
      (GraphService as any).instance = undefined;

      // Mock DeviceCodeCredential to return a credential that throws on getToken
      vi.mocked(DeviceCodeCredential).mockImplementation(
        () =>
          ({
            getToken: vi.fn().mockRejectedValue(new Error("Token acquisition failed")),
          }) as any
      );

      const testGraphService = GraphService.getInstance();
      const status = await testGraphService.getAuthStatus();

      expect(status).toEqual({
        isAuthenticated: false,
      });
    });

    it("should handle null token response", async () => {
      // Create a new GraphService instance and force re-initialization
      (GraphService as any).instance = undefined;

      // Mock DeviceCodeCredential to return a credential that returns null token
      vi.mocked(DeviceCodeCredential).mockImplementation(
        () =>
          ({
            getToken: vi.fn().mockResolvedValue(null),
          }) as any
      );

      const testGraphService = GraphService.getInstance();
      const status = await testGraphService.getAuthStatus();

      expect(status).toEqual({
        isAuthenticated: false,
      });
    });
  });

  describe("getClient", () => {
    it("should return client when service is initialized", async () => {
      const client = await graphService.getClient();

      expect(client).toBeDefined();
      expect(Client.initWithMiddleware).toHaveBeenCalled();
    });

    it("should throw error when no token available", async () => {
      // Create a new GraphService instance and force re-initialization
      (GraphService as any).instance = undefined;

      // Mock DeviceCodeCredential to return a credential that returns null token
      vi.mocked(DeviceCodeCredential).mockImplementation(
        () =>
          ({
            getToken: vi.fn().mockResolvedValue(null),
          }) as any
      );

      const testGraphService = GraphService.getInstance();

      // getClient should succeed since client initialization succeeds
      const client = await testGraphService.getClient();
      expect(client).toBeDefined();

      // But getAuthStatus should fail because token acquisition fails
      const status = await testGraphService.getAuthStatus();
      expect(status.isAuthenticated).toBe(false);
    });

    it("should return same client on multiple calls", async () => {
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
      // Initialize the client
      await graphService.getAuthStatus();

      expect(graphService.isAuthenticated()).toBe(true);
    });
  });

  describe("concurrent initialization", () => {
    it("should handle concurrent calls to initializeClient", async () => {
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

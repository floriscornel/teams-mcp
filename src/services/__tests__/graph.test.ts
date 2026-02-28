import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mockUser, server } from "../../test-utils/setup.js";

// Mock the msal-cache plugin
vi.mock("../../msal-cache.js", () => ({
  cachePlugin: {
    beforeCacheAccess: vi.fn(),
    afterCacheAccess: vi.fn(),
  },
  CACHE_PATH: "/mock/cache/path",
}));

// Mock @azure/msal-node
vi.mock("@azure/msal-node", () => ({
  PublicClientApplication: vi.fn(),
}));

// Mock @microsoft/microsoft-graph-client
vi.mock("@microsoft/microsoft-graph-client", () => ({
  Client: {
    initWithMiddleware: vi.fn(),
  },
}));

// Import after mocks are set up
import { PublicClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { FULL_SCOPES, GraphService, READ_ONLY_SCOPES } from "../graph.js";

/** Set up the default MSAL mock: one account, acquireTokenSilent succeeds */
function setupDefaultMsalMock() {
  vi.mocked(PublicClientApplication).mockImplementation(function () {
    return {
      getTokenCache: vi.fn().mockReturnValue({
        getAllAccounts: vi.fn().mockResolvedValue([{ username: "test@example.com" }]),
      }),
      acquireTokenSilent: vi.fn().mockResolvedValue({
        accessToken: "mock-access-token",
        expiresOn: new Date(Date.now() + 3600000),
      }),
    };
  } as any);
}

describe("GraphService", () => {
  let graphService: GraphService;

  beforeEach(() => {
    server.listen({ onUnhandledRequest: "error" });
    vi.clearAllMocks();
    setupDefaultMsalMock();

    // Reset GraphService singleton
    (GraphService as any).instance = undefined;
    graphService = GraphService.getInstance();
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
    it("should return unauthenticated when no MSAL accounts exist", async () => {
      vi.mocked(PublicClientApplication).mockImplementationOnce(function () {
        return {
          getTokenCache: vi.fn().mockReturnValue({
            getAllAccounts: vi.fn().mockResolvedValue([]),
          }),
          acquireTokenSilent: vi.fn(),
        };
      } as any);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({ isAuthenticated: false });
    });

    it("should return unauthenticated when acquireTokenSilent fails", async () => {
      vi.mocked(PublicClientApplication).mockImplementationOnce(function () {
        return {
          getTokenCache: vi.fn().mockReturnValue({
            getAllAccounts: vi.fn().mockResolvedValue([{ username: "test@example.com" }]),
          }),
          acquireTokenSilent: vi.fn().mockRejectedValue(new Error("InteractionRequiredAuthError")),
        };
      } as any);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({ isAuthenticated: false });
    });

    it("should return unauthenticated when acquireTokenSilent returns null", async () => {
      vi.mocked(PublicClientApplication).mockImplementationOnce(function () {
        return {
          getTokenCache: vi.fn().mockReturnValue({
            getAllAccounts: vi.fn().mockResolvedValue([{ username: "test@example.com" }]),
          }),
          acquireTokenSilent: vi.fn().mockResolvedValue(null),
        };
      } as any);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({ isAuthenticated: false });
    });

    it("should return authenticated status with valid MSAL token", async () => {
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({
        isAuthenticated: true,
        userPrincipalName: mockUser.userPrincipalName,
        displayName: mockUser.displayName,
        expiresAt: expect.any(String),
      });
    });

    it("should handle Graph API errors gracefully", async () => {
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockRejectedValue(new Error("API Error")),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({ isAuthenticated: false });
    });
  });

  describe("getClient", () => {
    it("should throw error when not authenticated", async () => {
      vi.mocked(PublicClientApplication).mockImplementationOnce(function () {
        return {
          getTokenCache: vi.fn().mockReturnValue({
            getAllAccounts: vi.fn().mockResolvedValue([]),
          }),
          acquireTokenSilent: vi.fn(),
        };
      } as any);

      await expect(graphService.getClient()).rejects.toThrow(
        "Not authenticated. Please run the authentication CLI tool first"
      );
    });

    it("should return client when authenticated", async () => {
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const client = await graphService.getClient();

      expect(client).toBeDefined();
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

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      await graphService.getAuthStatus();

      expect(graphService.isAuthenticated()).toBe(true);
    });
  });

  describe("MSAL token refresh", () => {
    it("should use acquireTokenSilent for auth provider", async () => {
      const mockAcquireTokenSilent = vi.fn().mockResolvedValue({
        accessToken: "mock-access-token",
        expiresOn: new Date(Date.now() + 3600000),
      });

      vi.mocked(PublicClientApplication).mockImplementationOnce(function () {
        return {
          getTokenCache: vi.fn().mockReturnValue({
            getAllAccounts: vi.fn().mockResolvedValue([{ username: "test@example.com" }]),
          }),
          acquireTokenSilent: mockAcquireTokenSilent,
        };
      } as any);

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      await graphService.getAuthStatus();

      // Verify MSAL PCA was created with correct config
      expect(PublicClientApplication).toHaveBeenCalledWith(
        expect.objectContaining({
          auth: expect.objectContaining({
            clientId: "14d82eec-204b-4c2f-b7e8-296a70dab67e",
            authority: "https://login.microsoftonline.com/common",
          }),
          cache: expect.objectContaining({
            cachePlugin: expect.any(Object),
          }),
        })
      );

      // Verify acquireTokenSilent was called during initialization
      expect(mockAcquireTokenSilent).toHaveBeenCalled();
    });

    it("should pass auth provider that calls acquireTokenSilent", async () => {
      const mockAcquireTokenSilent = vi.fn().mockResolvedValue({
        accessToken: "mock-access-token",
        expiresOn: new Date(Date.now() + 3600000),
      });

      vi.mocked(PublicClientApplication).mockImplementationOnce(function () {
        return {
          getTokenCache: vi.fn().mockReturnValue({
            getAllAccounts: vi.fn().mockResolvedValue([{ username: "test@example.com" }]),
          }),
          acquireTokenSilent: mockAcquireTokenSilent,
        };
      } as any);

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      await graphService.getClient();

      // Extract the auth provider passed to Client.initWithMiddleware
      const initCall = vi.mocked(Client.initWithMiddleware).mock.calls[0];
      const authProvider = (initCall[0] as any).authProvider;

      // Call getAccessToken to verify it uses acquireTokenSilent
      const token = await authProvider.getAccessToken();
      expect(token).toBe("mock-access-token");

      // acquireTokenSilent should have been called (once during init + once via authProvider)
      expect(mockAcquireTokenSilent).toHaveBeenCalledTimes(2);
    });
  });

  describe("concurrent initialization", () => {
    it("should handle concurrent calls to getAuthStatus", async () => {
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const promises = [
        graphService.getAuthStatus(),
        graphService.getAuthStatus(),
        graphService.getAuthStatus(),
      ];

      const results = await Promise.all(promises);

      for (const result of results) {
        expect(result.isAuthenticated).toBe(true);
      }
    });
  });

  describe("AUTH_TOKEN environment variable", () => {
    const originalEnv = process.env.AUTH_TOKEN;

    afterEach(() => {
      if (originalEnv === undefined) {
        delete process.env.AUTH_TOKEN;
      } else {
        process.env.AUTH_TOKEN = originalEnv;
      }
    });

    it("should use AUTH_TOKEN from environment when provided", async () => {
      const mockPayload = btoa(JSON.stringify({ aud: "https://graph.microsoft.com" }));
      const validToken = `header.${mockPayload}.signature`;
      process.env.AUTH_TOKEN = validToken;

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const status = await graphService.getAuthStatus();

      expect(status.isAuthenticated).toBe(true);
      // MSAL should NOT be used when AUTH_TOKEN is set
      expect(PublicClientApplication).not.toHaveBeenCalled();
    });

    it("should reject invalid JWT format from AUTH_TOKEN", async () => {
      process.env.AUTH_TOKEN = "invalid-token";

      const status = await graphService.getAuthStatus();

      expect(status.isAuthenticated).toBe(false);
    });

    it("should reject JWT without Graph audience from AUTH_TOKEN", async () => {
      const mockPayload = btoa(JSON.stringify({ aud: "https://other-service.com" }));
      const invalidToken = `header.${mockPayload}.signature`;
      process.env.AUTH_TOKEN = invalidToken;

      const status = await graphService.getAuthStatus();

      expect(status.isAuthenticated).toBe(false);
    });

    it("should handle JWT with audience as array from AUTH_TOKEN", async () => {
      const mockPayload = btoa(
        JSON.stringify({ aud: ["https://graph.microsoft.com", "https://other.com"] })
      );
      const validToken = `header.${mockPayload}.signature`;
      process.env.AUTH_TOKEN = validToken;

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const status = await graphService.getAuthStatus();

      expect(status.isAuthenticated).toBe(true);
    });

    it("should prefer AUTH_TOKEN over MSAL-based auth", async () => {
      const mockPayload = btoa(JSON.stringify({ aud: "https://graph.microsoft.com" }));
      const validToken = `header.${mockPayload}.signature`;
      process.env.AUTH_TOKEN = validToken;

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      await graphService.getAuthStatus();

      // MSAL should not be used when AUTH_TOKEN is present
      expect(PublicClientApplication).not.toHaveBeenCalled();
    });
  });

  describe("readOnlyMode", () => {
    it("should default to false", () => {
      expect(graphService.readOnlyMode).toBe(false);
    });

    it("should return FULL_SCOPES when readOnlyMode is false", () => {
      graphService.readOnlyMode = false;
      expect(graphService.scopes).toEqual(FULL_SCOPES);
    });

    it("should return READ_ONLY_SCOPES when readOnlyMode is true", () => {
      graphService.readOnlyMode = true;
      expect(graphService.scopes).toEqual(READ_ONLY_SCOPES);
    });

    it("READ_ONLY_SCOPES should not contain write scopes", () => {
      expect(READ_ONLY_SCOPES).not.toContain("ChannelMessage.Send");
      expect(READ_ONLY_SCOPES).not.toContain("Chat.ReadWrite");
      expect(READ_ONLY_SCOPES).not.toContain("Files.ReadWrite.All");
    });

    it("FULL_SCOPES should contain all READ_ONLY_SCOPES plus write scopes", () => {
      for (const scope of READ_ONLY_SCOPES) {
        expect(FULL_SCOPES).toContain(scope);
      }
      expect(FULL_SCOPES).toContain("ChannelMessage.Send");
      expect(FULL_SCOPES).toContain("Chat.ReadWrite");
      expect(FULL_SCOPES).toContain("Files.ReadWrite.All");
    });
  });
});

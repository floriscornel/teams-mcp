import { promises as fs } from "node:fs";
import type { TokenCacheContext } from "@azure/msal-node";
import { beforeEach, describe, expect, it, vi } from "vitest";

// Mock the filesystem
vi.mock("node:fs", () => ({
  promises: {
    readFile: vi.fn(),
    writeFile: vi.fn(),
  },
}));

// Import after mocks are set up
import { CACHE_PATH, cachePlugin } from "../msal-cache.js";

describe("MSAL Cache Plugin", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("beforeCacheAccess", () => {
    it("should deserialize cache data from file when it exists", async () => {
      const mockCacheData = '{"test": "data"}';
      vi.mocked(fs.readFile).mockResolvedValue(mockCacheData);

      const deserializeMock = vi.fn();
      const cacheContext = {
        tokenCache: {
          deserialize: deserializeMock,
        },
      } as unknown as TokenCacheContext;

      await cachePlugin.beforeCacheAccess(cacheContext);

      expect(fs.readFile).toHaveBeenCalledWith(CACHE_PATH, "utf8");
      expect(deserializeMock).toHaveBeenCalledWith(mockCacheData);
    });

    it("should handle missing cache file (ENOENT) silently", async () => {
      const error = new Error("File not found") as NodeJS.ErrnoException;
      error.code = "ENOENT";
      vi.mocked(fs.readFile).mockRejectedValue(error);

      const deserializeMock = vi.fn();
      const cacheContext = {
        tokenCache: {
          deserialize: deserializeMock,
        },
      } as unknown as TokenCacheContext;

      const consoleErrorSpy = vi.spyOn(console, "error").mockImplementation(() => {
        // Intentionally empty to suppress console output during tests
      });

      await cachePlugin.beforeCacheAccess(cacheContext);

      expect(fs.readFile).toHaveBeenCalledWith(CACHE_PATH, "utf8");
      expect(deserializeMock).not.toHaveBeenCalled();
      expect(consoleErrorSpy).not.toHaveBeenCalled();

      consoleErrorSpy.mockRestore();
    });

    it("should log error for other file read failures", async () => {
      const error = new Error("Permission denied") as NodeJS.ErrnoException;
      error.code = "EACCES";
      vi.mocked(fs.readFile).mockRejectedValue(error);

      const deserializeMock = vi.fn();
      const cacheContext = {
        tokenCache: {
          deserialize: deserializeMock,
        },
      } as unknown as TokenCacheContext;

      const consoleErrorSpy = vi.spyOn(console, "error").mockImplementation(() => {
        // Intentionally empty to suppress console output during tests
      });

      await cachePlugin.beforeCacheAccess(cacheContext);

      expect(fs.readFile).toHaveBeenCalledWith(CACHE_PATH, "utf8");
      expect(deserializeMock).not.toHaveBeenCalled();
      expect(consoleErrorSpy).toHaveBeenCalledWith("Warning: Could not read token cache:", error);

      consoleErrorSpy.mockRestore();
    });
  });

  describe("afterCacheAccess", () => {
    it("should serialize and write cache data when cache has changed", async () => {
      const mockSerializedData = '{"test": "serialized"}';
      const serializeMock = vi.fn().mockReturnValue(mockSerializedData);

      const cacheContext = {
        cacheHasChanged: true,
        tokenCache: {
          serialize: serializeMock,
        },
      } as unknown as TokenCacheContext;

      await cachePlugin.afterCacheAccess(cacheContext);

      expect(serializeMock).toHaveBeenCalled();
      expect(fs.writeFile).toHaveBeenCalledWith(CACHE_PATH, mockSerializedData, "utf8");
    });

    it("should not write cache data when cache has not changed", async () => {
      const serializeMock = vi.fn();

      const cacheContext = {
        cacheHasChanged: false,
        tokenCache: {
          serialize: serializeMock,
        },
      } as unknown as TokenCacheContext;

      await cachePlugin.afterCacheAccess(cacheContext);

      expect(serializeMock).not.toHaveBeenCalled();
      expect(fs.writeFile).not.toHaveBeenCalled();
    });

    it("should log error when cache write fails", async () => {
      const error = new Error("Disk full");
      vi.mocked(fs.writeFile).mockRejectedValue(error);

      const mockSerializedData = '{"test": "serialized"}';
      const serializeMock = vi.fn().mockReturnValue(mockSerializedData);

      const cacheContext = {
        cacheHasChanged: true,
        tokenCache: {
          serialize: serializeMock,
        },
      } as unknown as TokenCacheContext;

      const consoleErrorSpy = vi.spyOn(console, "error").mockImplementation(() => {
        // Intentionally empty to suppress console output during tests
      });

      await cachePlugin.afterCacheAccess(cacheContext);

      expect(serializeMock).toHaveBeenCalled();
      expect(fs.writeFile).toHaveBeenCalledWith(CACHE_PATH, mockSerializedData, "utf8");
      expect(consoleErrorSpy).toHaveBeenCalledWith("Warning: Could not write token cache:", error);

      consoleErrorSpy.mockRestore();
    });
  });

  describe("CACHE_PATH", () => {
    it("should export CACHE_PATH", () => {
      expect(CACHE_PATH).toBeDefined();
      expect(typeof CACHE_PATH).toBe("string");
      expect(CACHE_PATH).toContain(".teams-mcp-token-cache.json");
    });
  });
});

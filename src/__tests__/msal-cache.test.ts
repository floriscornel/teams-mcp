import { promises as fs } from "node:fs";
import type { TokenCacheContext } from "@azure/msal-node";
import { beforeEach, describe, expect, it, vi } from "vitest";

// Mock the filesystem
vi.mock("node:fs", () => ({
  promises: {
    readFile: vi.fn(),
    writeFile: vi.fn(),
    rename: vi.fn(),
    rm: vi.fn(),
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
    it("should write cache data to an owner-only temp file and rename it over the cache", async () => {
      const mockSerializedData = '{"test": "serialized"}';
      const serializeMock = vi.fn().mockReturnValue(mockSerializedData);
      const tmpPath = `${CACHE_PATH}.${process.pid}.tmp`;

      const cacheContext = {
        cacheHasChanged: true,
        tokenCache: {
          serialize: serializeMock,
        },
      } as unknown as TokenCacheContext;

      await cachePlugin.afterCacheAccess(cacheContext);

      expect(serializeMock).toHaveBeenCalled();
      expect(fs.writeFile).toHaveBeenCalledWith(tmpPath, mockSerializedData, {
        encoding: "utf8",
        mode: 0o600,
        flag: "wx",
      });
      expect(fs.rename).toHaveBeenCalledWith(tmpPath, CACHE_PATH);
      // The existing cache file is never written in place
      expect(fs.writeFile).not.toHaveBeenCalledWith(
        CACHE_PATH,
        expect.anything(),
        expect.anything()
      );
    });

    it("should clean up the temp file and warn when the rename fails", async () => {
      const error = new Error("EPERM");
      vi.mocked(fs.rename).mockRejectedValue(error);
      const consoleErrorSpy = vi.spyOn(console, "error").mockImplementation(() => {
        // Intentionally empty to suppress console output during tests
      });

      const cacheContext = {
        cacheHasChanged: true,
        tokenCache: {
          serialize: vi.fn().mockReturnValue("{}"),
        },
      } as unknown as TokenCacheContext;

      await cachePlugin.afterCacheAccess(cacheContext);

      expect(fs.rm).toHaveBeenLastCalledWith(`${CACHE_PATH}.${process.pid}.tmp`, { force: true });
      expect(consoleErrorSpy).toHaveBeenCalledWith("Warning: Could not write token cache:", error);
      consoleErrorSpy.mockRestore();
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
      expect(fs.rename).not.toHaveBeenCalled();
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

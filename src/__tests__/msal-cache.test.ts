import { promises as fs } from "node:fs";
import { basename, join } from "node:path";
import type { TokenCacheContext } from "@azure/msal-node";
import { beforeEach, describe, expect, it, vi } from "vitest";

// Mock the filesystem
vi.mock("node:fs", () => ({
  promises: {
    mkdir: vi.fn(),
    readFile: vi.fn(),
    rename: vi.fn(),
    rm: vi.fn(),
    stat: vi.fn(),
    writeFile: vi.fn(),
  },
}));

// Import after mocks are set up
import { CACHE_PATH, cachePlugin } from "../msal-cache.js";

const CACHE_LOCK_PATH = `${CACHE_PATH}.lock`;
const CACHE_LOCK_OWNER_PATH = join(CACHE_LOCK_PATH, "owner");
const cacheFileNamePattern = escapeRegExp(basename(CACHE_PATH));
const corruptCachePathPattern = new RegExp(`${cacheFileNamePattern}\\.corrupt\\.\\d+\\.\\d+$`);
const tempCachePathPattern = new RegExp(`\\.${cacheFileNamePattern}\\.\\d+\\.\\d+\\.tmp$`);

let currentLockOwner = "";

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function mockCacheRead(cacheData: string): void {
  vi.mocked(fs.readFile).mockImplementation(async (path) => {
    if (path === CACHE_LOCK_OWNER_PATH) {
      return currentLockOwner;
    }
    if (path === CACHE_PATH) {
      return cacheData;
    }
    throw new Error(`Unexpected read path: ${String(path)}`);
  });
}

function mockCacheReadError(error: Error): void {
  vi.mocked(fs.readFile).mockImplementation(async (path) => {
    if (path === CACHE_LOCK_OWNER_PATH) {
      return currentLockOwner;
    }
    if (path === CACHE_PATH) {
      throw error;
    }
    throw new Error(`Unexpected read path: ${String(path)}`);
  });
}

describe("MSAL Cache Plugin", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    currentLockOwner = "";
    vi.mocked(fs.mkdir).mockResolvedValue(undefined);
    vi.mocked(fs.readFile).mockImplementation(async (path) => {
      if (path === CACHE_LOCK_OWNER_PATH) {
        return currentLockOwner;
      }
      throw new Error(`Unexpected read path: ${String(path)}`);
    });
    vi.mocked(fs.rename).mockResolvedValue(undefined);
    vi.mocked(fs.rm).mockResolvedValue(undefined);
    vi.mocked(fs.writeFile).mockImplementation(async (path, data) => {
      if (path === CACHE_LOCK_OWNER_PATH) {
        currentLockOwner = String(data);
      }
    });
  });

  describe("beforeCacheAccess", () => {
    it("should deserialize cache data from file when it exists", async () => {
      const mockCacheData = '{"test": "data"}';
      mockCacheRead(mockCacheData);

      const deserializeMock = vi.fn();
      const cacheContext = {
        tokenCache: {
          deserialize: deserializeMock,
        },
      } as unknown as TokenCacheContext;

      await cachePlugin.beforeCacheAccess(cacheContext);

      expect(fs.readFile).toHaveBeenCalledWith(CACHE_PATH, "utf8");
      expect(deserializeMock).toHaveBeenCalledWith(mockCacheData);
      expect(fs.rm).toHaveBeenCalledWith(CACHE_LOCK_PATH, { recursive: true, force: true });
    });

    it("should handle missing cache file (ENOENT) silently", async () => {
      const error = new Error("File not found") as NodeJS.ErrnoException;
      error.code = "ENOENT";
      mockCacheReadError(error);

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
      mockCacheReadError(error);

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

    it("should quarantine invalid cache data and continue with an empty cache", async () => {
      const error = new SyntaxError("Unexpected non-whitespace character after JSON");
      mockCacheRead("{invalid-json}");

      const deserializeMock = vi.fn().mockImplementation(() => {
        throw error;
      });
      const cacheContext = {
        tokenCache: {
          deserialize: deserializeMock,
        },
      } as unknown as TokenCacheContext;

      const consoleErrorSpy = vi.spyOn(console, "error").mockImplementation(() => {
        // Intentionally empty to suppress console output during tests
      });

      await cachePlugin.beforeCacheAccess(cacheContext);

      expect(deserializeMock).toHaveBeenCalledWith("{invalid-json}");
      expect(fs.rename).toHaveBeenCalledWith(
        CACHE_PATH,
        expect.stringMatching(corruptCachePathPattern)
      );
      expect(consoleErrorSpy).toHaveBeenCalledWith(
        "Warning: Token cache is invalid; moved aside:",
        expect.stringMatching(corruptCachePathPattern)
      );

      consoleErrorSpy.mockRestore();
    });

    it("should not release a lock owned by another process", async () => {
      const mockCacheData = '{"test": "data"}';
      vi.mocked(fs.readFile).mockImplementation(async (path) => {
        if (path === CACHE_LOCK_OWNER_PATH) {
          return currentLockOwner;
        }
        if (path === CACHE_PATH) {
          currentLockOwner = "999999.1.other-owner";
          return mockCacheData;
        }
        throw new Error(`Unexpected read path: ${String(path)}`);
      });

      const cacheContext = {
        tokenCache: {
          deserialize: vi.fn(),
        },
      } as unknown as TokenCacheContext;

      await cachePlugin.beforeCacheAccess(cacheContext);

      expect(fs.rm).not.toHaveBeenCalledWith(CACHE_LOCK_PATH, { recursive: true, force: true });
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
      expect(fs.writeFile).toHaveBeenCalledWith(
        expect.stringMatching(tempCachePathPattern),
        mockSerializedData,
        { encoding: "utf8", mode: 0o600 }
      );
      expect(fs.rename).toHaveBeenCalledWith(
        expect.stringMatching(tempCachePathPattern),
        CACHE_PATH
      );
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
      vi.mocked(fs.writeFile).mockImplementation(async (path, data) => {
        if (path === CACHE_LOCK_OWNER_PATH) {
          currentLockOwner = String(data);
          return;
        }
        throw error;
      });

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
      expect(fs.writeFile).toHaveBeenCalledWith(
        expect.stringMatching(tempCachePathPattern),
        mockSerializedData,
        { encoding: "utf8", mode: 0o600 }
      );
      expect(consoleErrorSpy).toHaveBeenCalledWith("Warning: Could not write token cache:", error);

      consoleErrorSpy.mockRestore();
    });
  });

  describe("CACHE_PATH", () => {
    it("should export CACHE_PATH", () => {
      expect(CACHE_PATH).toBeDefined();
      expect(typeof CACHE_PATH).toBe("string");
      expect(basename(CACHE_PATH).length).toBeGreaterThan(0);
    });
  });
});

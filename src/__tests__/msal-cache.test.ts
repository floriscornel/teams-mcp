import { promises as fs, type Stats } from "node:fs";
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

function errnoError(code: string, message = code): NodeJS.ErrnoException {
  const error = new Error(message) as NodeJS.ErrnoException;
  error.code = code;
  return error;
}

// Make mkdir(CACHE_LOCK_PATH) fail with EEXIST `times` times before succeeding
function mockLockHeld(times: number): void {
  let remaining = times;
  vi.mocked(fs.mkdir).mockImplementation(async (path) => {
    if (path === CACHE_LOCK_PATH && remaining > 0) {
      remaining--;
      throw errnoError("EEXIST");
    }
    return undefined;
  });
}

function mockLockStat(ageMs: number): void {
  vi.mocked(fs.stat).mockResolvedValue({ mtimeMs: Date.now() - ageMs } as Stats);
}

function mockExistingLockOwner(owner: string | Error): void {
  vi.mocked(fs.readFile).mockImplementation(async (path) => {
    if (path === CACHE_LOCK_OWNER_PATH) {
      // The stale lock's owner until this process writes its own
      if (currentLockOwner) return currentLockOwner;
      if (owner instanceof Error) throw owner;
      return owner;
    }
    if (path === CACHE_PATH) {
      throw errnoError("ENOENT");
    }
    throw new Error(`Unexpected read path: ${String(path)}`);
  });
}

function emptyContext(): TokenCacheContext {
  return { tokenCache: { deserialize: vi.fn() } } as unknown as TokenCacheContext;
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

  describe("cache lock", () => {
    const LOCK_RM_ARGS = [CACHE_LOCK_PATH, { recursive: true, force: true }] as const;

    it("should evict a stale lock whose owner process is gone", async () => {
      mockLockHeld(1);
      mockLockStat(60_000);
      mockExistingLockOwner("2147483646.1.dead-owner");
      vi.spyOn(process, "kill").mockImplementation(() => {
        throw errnoError("ESRCH");
      });

      await cachePlugin.beforeCacheAccess(emptyContext());

      expect(fs.rm).toHaveBeenCalledWith(...LOCK_RM_ARGS);
      expect(fs.writeFile).toHaveBeenCalledWith(CACHE_LOCK_OWNER_PATH, expect.any(String), {
        encoding: "utf8",
        flag: "wx",
        mode: 0o600,
      });
      vi.mocked(process.kill).mockRestore();
    });

    it("should evict a stale lock without an owner file", async () => {
      mockLockHeld(1);
      mockLockStat(60_000);
      mockExistingLockOwner(errnoError("ENOENT"));

      await cachePlugin.beforeCacheAccess(emptyContext());

      expect(fs.rm).toHaveBeenCalledWith(...LOCK_RM_ARGS);
    });

    it("should retry when the lock disappears before it can be inspected", async () => {
      mockLockHeld(1);
      vi.mocked(fs.stat).mockRejectedValue(errnoError("ENOENT"));
      mockExistingLockOwner(errnoError("ENOENT"));

      await cachePlugin.beforeCacheAccess(emptyContext());

      expect(fs.stat).toHaveBeenCalledWith(CACHE_LOCK_PATH);
      expect(fs.mkdir).toHaveBeenCalledWith(CACHE_LOCK_PATH);
    });

    it("should wait for a fresh lock and acquire it once released", async () => {
      mockLockHeld(1);
      mockLockStat(0);
      mockExistingLockOwner("1.1.fresh-owner");

      await cachePlugin.beforeCacheAccess(emptyContext());

      expect(fs.mkdir).toHaveBeenCalledTimes(3);
      expect(fs.rm).toHaveBeenCalledTimes(1);
    });

    it.each([
      ["running", undefined],
      ["owned by another user (EPERM)", errnoError("EPERM")],
    ])("should not evict a stale lock whose owner is %s", async (_label, killError) => {
      mockLockHeld(Number.POSITIVE_INFINITY);
      mockLockStat(60_000);
      mockExistingLockOwner(`${process.pid}.1.live-owner`);
      const killSpy = vi.spyOn(process, "kill").mockImplementation(() => {
        if (killError) throw killError;
        return true;
      });
      const now = Date.now();
      const dateSpy = vi
        .spyOn(Date, "now")
        .mockReturnValueOnce(now)
        .mockReturnValue(now + 60_000);

      await expect(cachePlugin.beforeCacheAccess(emptyContext())).rejects.toThrow(
        "Timed out waiting for token cache lock"
      );
      expect(fs.rm).not.toHaveBeenCalled();

      killSpy.mockRestore();
      dateSpy.mockRestore();
    });

    it("should propagate unexpected errors while inspecting the lock", async () => {
      mockLockHeld(1);
      vi.mocked(fs.stat).mockRejectedValue(errnoError("EACCES"));

      await expect(cachePlugin.beforeCacheAccess(emptyContext())).rejects.toThrow("EACCES");
    });

    it("should propagate unexpected errors while reading the lock owner", async () => {
      mockLockHeld(1);
      mockLockStat(60_000);
      mockExistingLockOwner(errnoError("EACCES"));

      await expect(cachePlugin.beforeCacheAccess(emptyContext())).rejects.toThrow("EACCES");
      expect(fs.rm).not.toHaveBeenCalled();
    });

    it("should propagate unexpected errors while creating the lock", async () => {
      vi.mocked(fs.mkdir).mockImplementation(async (path) => {
        if (path === CACHE_LOCK_PATH) throw errnoError("EACCES");
        return undefined;
      });

      await expect(cachePlugin.beforeCacheAccess(emptyContext())).rejects.toThrow("EACCES");
    });

    it("should remove the lock directory when the owner file cannot be written", async () => {
      vi.mocked(fs.writeFile).mockRejectedValue(errnoError("ENOSPC"));

      await expect(cachePlugin.beforeCacheAccess(emptyContext())).rejects.toThrow("ENOSPC");
      expect(fs.rm).toHaveBeenCalledWith(...LOCK_RM_ARGS);
    });

    it("should skip release when the lock was already removed", async () => {
      vi.mocked(fs.writeFile).mockResolvedValue(undefined);
      mockExistingLockOwner(errnoError("ENOENT"));

      await cachePlugin.beforeCacheAccess(emptyContext());

      expect(fs.rm).not.toHaveBeenCalled();
    });

    it("should propagate unexpected errors while releasing the lock", async () => {
      vi.mocked(fs.writeFile).mockResolvedValue(undefined);
      mockExistingLockOwner(errnoError("EACCES"));

      await expect(cachePlugin.beforeCacheAccess(emptyContext())).rejects.toThrow("EACCES");
    });

    it("should warn when an invalid cache cannot be quarantined", async () => {
      mockCacheRead("{invalid-json}");
      const renameError = errnoError("EBUSY");
      vi.mocked(fs.rename).mockRejectedValue(renameError);
      const consoleErrorSpy = vi.spyOn(console, "error").mockImplementation(() => {
        // Intentionally empty to suppress console output during tests
      });
      const cacheContext = {
        tokenCache: {
          deserialize: vi.fn().mockImplementation(() => {
            throw new SyntaxError("bad");
          }),
        },
      } as unknown as TokenCacheContext;

      await cachePlugin.beforeCacheAccess(cacheContext);

      expect(consoleErrorSpy).toHaveBeenCalledWith(
        "Warning: Could not quarantine invalid token cache:",
        renameError
      );
      consoleErrorSpy.mockRestore();
    });

    it("should stay quiet when the invalid cache disappears before quarantine", async () => {
      mockCacheRead("{invalid-json}");
      vi.mocked(fs.rename).mockRejectedValue(errnoError("ENOENT"));
      const consoleErrorSpy = vi.spyOn(console, "error").mockImplementation(() => {
        // Intentionally empty to suppress console output during tests
      });
      const cacheContext = {
        tokenCache: {
          deserialize: vi.fn().mockImplementation(() => {
            throw new SyntaxError("bad");
          }),
        },
      } as unknown as TokenCacheContext;

      await cachePlugin.beforeCacheAccess(cacheContext);

      expect(consoleErrorSpy).not.toHaveBeenCalled();
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

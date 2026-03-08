import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import type { TokenCacheContext } from "@azure/msal-node";
import { beforeEach, describe, expect, it, vi } from "vitest";

const CACHE_PATH = join(homedir(), ".teams-mcp-token-cache.json");

// Mock the filesystem
vi.mock("node:fs", () => ({
  promises: {
    readFile: vi.fn(),
    writeFile: vi.fn(),
    unlink: vi.fn(),
  },
}));

// Mock msal-node-extensions to use our mocked fs so we can test plugin behavior
vi.mock("@azure/msal-node-extensions", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@azure/msal-node-extensions")>();
  const fsMock = await import("node:fs").then((m) => m.promises);
  const createFileBasedPersistence = (path: string) => ({
    load: () => (fsMock.readFile as ReturnType<typeof vi.fn>)(path, "utf8"),
    save: (contents: string) =>
      (fsMock.writeFile as ReturnType<typeof vi.fn>)(path, contents, "utf8"),
    delete: vi.fn().mockResolvedValue(undefined),
  });
  return {
    ...actual,
    KeychainPersistence: {
      create: (_path: string) => Promise.resolve(createFileBasedPersistence(_path)),
    },
    FilePersistenceWithDataProtection: {
      create: (_path: string) => Promise.resolve(createFileBasedPersistence(_path)),
    },
    LibSecretPersistence: {
      create: (_path: string) => Promise.resolve(createFileBasedPersistence(_path)),
    },
    FilePersistence: {
      create: (path: string) => Promise.resolve(createFileBasedPersistence(path)),
    },
  };
});

// Import after mocks
import { CACHE_PATH as EXPORTED_CACHE_PATH, createCachePlugin, clearTokenCache } from "../msal-cache.js";

describe("MSAL Cache Plugin", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("beforeCacheAccess", () => {
    it("should deserialize cache data from file when it exists", async () => {
      const mockCacheData = '{"test": "data"}';
      vi.mocked(fs.readFile).mockResolvedValue(mockCacheData);

      const cachePlugin = await createCachePlugin();
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

      const cachePlugin = await createCachePlugin();
      const deserializeMock = vi.fn();
      const cacheContext = {
        tokenCache: {
          deserialize: deserializeMock,
        },
      } as unknown as TokenCacheContext;

      const consoleErrorSpy = vi.spyOn(console, "error").mockImplementation(() => {});

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

      const cachePlugin = await createCachePlugin();
      const deserializeMock = vi.fn();
      const cacheContext = {
        tokenCache: {
          deserialize: deserializeMock,
        },
      } as unknown as TokenCacheContext;

      const consoleErrorSpy = vi.spyOn(console, "error").mockImplementation(() => {});

      await cachePlugin.beforeCacheAccess(cacheContext);

      expect(fs.readFile).toHaveBeenCalledWith(CACHE_PATH, "utf8");
      expect(deserializeMock).not.toHaveBeenCalled();
      expect(consoleErrorSpy).toHaveBeenCalled();

      consoleErrorSpy.mockRestore();
    });
  });

  describe("afterCacheAccess", () => {
    it("should serialize and write cache data when cache has changed", async () => {
      const mockSerializedData = '{"test": "serialized"}';
      const serializeMock = vi.fn().mockReturnValue(mockSerializedData);

      const cachePlugin = await createCachePlugin();
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

      const cachePlugin = await createCachePlugin();
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

      const cachePlugin = await createCachePlugin();
      const cacheContext = {
        cacheHasChanged: true,
        tokenCache: {
          serialize: serializeMock,
        },
      } as unknown as TokenCacheContext;

      const consoleErrorSpy = vi.spyOn(console, "error").mockImplementation(() => {});

      await cachePlugin.afterCacheAccess(cacheContext);

      expect(serializeMock).toHaveBeenCalled();
      expect(fs.writeFile).toHaveBeenCalledWith(CACHE_PATH, mockSerializedData, "utf8");
      expect(consoleErrorSpy).toHaveBeenCalled();

      consoleErrorSpy.mockRestore();
    });
  });

  describe("CACHE_PATH", () => {
    it("should export CACHE_PATH", () => {
      expect(EXPORTED_CACHE_PATH).toBeDefined();
      expect(typeof EXPORTED_CACHE_PATH).toBe("string");
      expect(EXPORTED_CACHE_PATH).toContain(".teams-mcp-token-cache.json");
    });
  });

  describe("clearTokenCache", () => {
    it("should resolve without throwing", async () => {
      await expect(clearTokenCache()).resolves.toBeUndefined();
    });
  });
});

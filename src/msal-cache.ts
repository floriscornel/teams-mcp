import { randomUUID } from "node:crypto";
import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import type { ICachePlugin, TokenCacheContext } from "@azure/msal-node";

const CACHE_PATH = join(homedir(), ".teams-mcp-token-cache.json");

/**
 * Write a file that only its owner can read. writeFile's mode only applies when
 * the file is created, so the data goes to a fresh temp file that is renamed over
 * the target: an existing file with looser permissions never receives the data.
 */
export async function writeOwnerOnly(path: string, data: string): Promise<void> {
  // Unique per call, so overlapping writes in one process never share a temp file.
  // "wx" makes sure the file is newly created with the owner-only mode.
  const tmpPath = `${path}.${process.pid}.${randomUUID()}.tmp`;
  try {
    await fs.writeFile(tmpPath, data, { encoding: "utf8", mode: 0o600, flag: "wx" });
    await fs.rename(tmpPath, path);
  } catch (error) {
    try {
      await fs.rm(tmpPath, { force: true });
    } catch {
      // Best effort: the original error is the one worth reporting
    }
    throw error;
  }
}

/**
 * Custom file-based cache plugin for MSAL Node
 * Stores tokens (including refresh tokens) in a JSON file
 */
export const cachePlugin: ICachePlugin = {
  async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    try {
      const data = await fs.readFile(CACHE_PATH, "utf8");
      cacheContext.tokenCache.deserialize(data);
    } catch (error) {
      // File doesn't exist or is invalid - start with empty cache
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") {
        console.error("Warning: Could not read token cache:", error);
      }
    }
  },

  async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    if (cacheContext.cacheHasChanged) {
      try {
        // Owner-only: the cache holds refresh tokens.
        await writeOwnerOnly(CACHE_PATH, cacheContext.tokenCache.serialize());
      } catch (error) {
        console.error("Warning: Could not write token cache:", error);
      }
    }
  },
};

export { CACHE_PATH };

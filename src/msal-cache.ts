import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import type { ICachePlugin, TokenCacheContext } from "@azure/msal-node";

const CACHE_PATH = join(homedir(), ".teams-mcp-token-cache.json");

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
        const data = cacheContext.tokenCache.serialize();
        // Owner-only: the cache holds refresh tokens. mode only applies on create,
        // so chmod tightens caches written by older versions too.
        await fs.writeFile(CACHE_PATH, data, { encoding: "utf8", mode: 0o600 });
        try {
          await fs.chmod(CACHE_PATH, 0o600);
        } catch {
          // chmod is unsupported on some platforms (e.g. Windows) — best effort
        }
      } catch (error) {
        console.error("Warning: Could not write token cache:", error);
      }
    }
  },
};

export { CACHE_PATH };

import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import type { ICachePlugin, TokenCacheContext } from "@azure/msal-node";

export const CACHE_PATH = join(homedir(), ".teams-mcp-token-cache.json");
const ENCRYPTED_CACHE_PATH = join(homedir(), ".teams-mcp-token-cache.dat");

/**
 * Attempts to create an OS-level encrypted cache plugin using msal-node-extensions.
 * - Windows  : DPAPI (tied to your Windows user account)
 * - macOS    : Keychain
 * - Linux    : libsecret
 * Returns null if the extension is unavailable (fallback to plaintext).
 */
async function createEncryptedPlugin(): Promise<ICachePlugin | null> {
  try {
    const ext = await import("@azure/msal-node-extensions");

    const platform = process.platform;
    let persistence: unknown;

    if (platform === "win32") {
      const { FilePersistenceWithDataProtection, DataProtectionScope } = ext;
      persistence = await FilePersistenceWithDataProtection.create(
        ENCRYPTED_CACHE_PATH,
        DataProtectionScope.CurrentUser,
        ""
      );
    } else if (platform === "darwin") {
      const { KeychainPersistence } = ext;
      persistence = await KeychainPersistence.create(
        ENCRYPTED_CACHE_PATH,
        "teams-mcp",
        "token-cache"
      );
    } else {
      const { LibSecretPersistence } = ext;
      persistence = await LibSecretPersistence.create(
        ENCRYPTED_CACHE_PATH,
        "teams-mcp",
        "token-cache"
      );
    }

    const { PersistenceCachePlugin } = ext;
    return new PersistenceCachePlugin(persistence as never);
  } catch {
    return null;
  }
}

// Singleton — resolved once at startup
let encryptedPluginPromise: Promise<ICachePlugin | null> | null = null;

function getEncryptedPlugin(): Promise<ICachePlugin | null> {
  if (!encryptedPluginPromise) {
    encryptedPluginPromise = createEncryptedPlugin();
  }
  return encryptedPluginPromise;
}

/**
 * ICachePlugin exposed to MSAL.
 * Uses OS-encrypted storage when available; falls back to the plaintext
 * JSON file with a one-time warning otherwise.
 */
let warnedAboutPlaintext = false;
export const cachePlugin: ICachePlugin = {
  async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    const plugin = await getEncryptedPlugin();
    if (plugin) {
      return plugin.beforeCacheAccess(cacheContext);
    }

     // Plaintext fallback
    if (!warnedAboutPlaintext) {
      warnedAboutPlaintext = true;
      console.error(
        "⚠️  teams-mcp: `@azure/msal-node-extensions` unavailable — " +
        "token cache stored in plaintext at " + CACHE_PATH
      );
    }
    try {
      const data = await fs.readFile(CACHE_PATH, "utf8");
      cacheContext.tokenCache.deserialize(data);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") {
        console.error("Warning: Could not read token cache:", error);
      }
    }
  },

  async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    const plugin = await getEncryptedPlugin();
    if (plugin) {
      return plugin.afterCacheAccess(cacheContext);
    }

    // Plaintext fallback
    if (cacheContext.cacheHasChanged) {
      try {
        const data = cacheContext.tokenCache.serialize();
        await fs.writeFile(CACHE_PATH, data, "utf8");
      } catch (error) {
        console.error("Warning: Could not write token cache:", error);
      }
    }
  },
};

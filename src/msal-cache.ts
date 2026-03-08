import { homedir } from "node:os";
import { join } from "node:path";
import type { ICachePlugin } from "@azure/msal-node";
import {
  DataProtectionScope,
  FilePersistence,
  FilePersistenceWithDataProtection,
  KeychainPersistence,
  LibSecretPersistence,
  PersistenceCachePlugin,
} from "@azure/msal-node-extensions";

// Secure cache location — same directory convention as before but now encrypted
// - macOS:   stored in Keychain (file path used only as a lock/metadata file)
// - Windows: DPAPI-encrypted file at this path
// - Linux:   libsecret keyring (file path used only as a lock/metadata file)
//            falls back to plaintext file if libsecret is unavailable and
//            TEAMS_MCP_ALLOW_PLAINTEXT_CACHE=true is set
const CACHE_PATH = join(homedir(), ".teams-mcp-token-cache.json");

/**
 * Creates the appropriate OS-native persistence backend and wraps it in
 * PersistenceCachePlugin, which implements the ICachePlugin interface that
 * PublicClientApplication expects.
 *
 * Platform behaviour:
 *   macOS   — Tokens stored in the login Keychain under the service name
 *             "teams-mcp". The cache file on disk is only used as a lock file.
 *   Windows — Tokens written to CACHE_PATH encrypted with Windows DPAPI
 *             (CurrentUser scope). Only decryptable by the same user account.
 *   Linux   — Tokens stored in the system Secret Service (libsecret / GNOME
 *             Keyring / KWallet). Requires libsecret-1-dev to be installed.
 *             Set TEAMS_MCP_ALLOW_PLAINTEXT_CACHE=true to fall back to an
 *             unencrypted file when libsecret is unavailable.
 */
/** Creates the OS-native persistence used for the token cache (shared by createCachePlugin and clearTokenCache). */
async function createCachePersistence(): Promise<
  Awaited<ReturnType<typeof KeychainPersistence.create>> &
    { delete: () => Promise<void> }
> {
  const platform = process.platform;

  if (platform === "darwin") {
    return await KeychainPersistence.create(
      CACHE_PATH,
      "teams-mcp",
      "MSALCache"
    );
  }

  if (platform === "win32") {
    return await FilePersistenceWithDataProtection.create(
      CACHE_PATH,
      DataProtectionScope.CurrentUser
    );
  }

  if (platform === "linux") {
    const allowPlaintext =
      process.env.TEAMS_MCP_ALLOW_PLAINTEXT_CACHE === "true";
    try {
      return await LibSecretPersistence.create(
        CACHE_PATH,
        "teams-mcp",
        "MSALCache"
      );
    } catch (err) {
      if (allowPlaintext) {
        console.error(
          "Warning: libsecret unavailable, falling back to unencrypted " +
            "token cache. Set TEAMS_MCP_ALLOW_PLAINTEXT_CACHE=true to " +
            "suppress this warning in environments without a keyring."
        );
        return await FilePersistence.create(CACHE_PATH);
      }
      throw new Error(
        "Unable to initialise secure token cache on Linux: libsecret is " +
          "not available. Install libsecret-1-dev (Debian/Ubuntu) or " +
          "libsecret-devel (Fedora/RHEL), or set " +
          "TEAMS_MCP_ALLOW_PLAINTEXT_CACHE=true to use an unencrypted file " +
          "instead.\n\nOriginal error: " +
          String(err)
      );
    }
  }

  console.error(
    `Warning: Secure token storage is not supported on platform "${platform}". ` +
      "Tokens will be stored in an unencrypted file at " +
      CACHE_PATH
  );
  return await FilePersistence.create(CACHE_PATH);
}

export async function createCachePlugin(): Promise<ICachePlugin> {
  const persistence = await createCachePersistence();
  return new PersistenceCachePlugin(persistence);
}

/**
 * Clears the token cache from OS-native storage (Keychain, DPAPI, or libsecret).
 * Call this on logout so credentials are fully removed.
 */
export async function clearTokenCache(): Promise<void> {
  try {
    const persistence = await createCachePersistence();
    await persistence.delete();
  } catch (err) {
    // Ignore if nothing was stored (e.g. first run or already logged out)
  }
}

export { CACHE_PATH };

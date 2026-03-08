import { homedir } from "node:os";
import { join } from "node:path";
import {
  DataProtectionScope,
  FilePersistence,
  FilePersistenceWithDataProtection,
  KeychainPersistence,
  LibSecretPersistence,
} from "@azure/msal-node-extensions";

// Auth info is stored in OS secure storage (Keychain / DPAPI / libsecret)
// so we don't keep a plaintext file with account/scopes metadata.
const AUTH_INFO_PATH = join(homedir(), ".teams-mcp-auth.json");

type PersistenceLike = {
  load: () => Promise<string | Buffer>;
  save: (contents: string) => Promise<void>;
  delete: () => Promise<void>;
};

async function createAuthInfoPersistence(): Promise<PersistenceLike> {
  const platform = process.platform;
  const serviceName = "teams-mcp";
  const accountName = "AuthInfo";

  if (platform === "darwin") {
    return await KeychainPersistence.create(
      AUTH_INFO_PATH,
      serviceName,
      accountName
    );
  }

  if (platform === "win32") {
    return await FilePersistenceWithDataProtection.create(
      AUTH_INFO_PATH,
      DataProtectionScope.CurrentUser
    );
  }

  if (platform === "linux") {
    const allowPlaintext =
      process.env.TEAMS_MCP_ALLOW_PLAINTEXT_CACHE === "true";
    try {
      return await LibSecretPersistence.create(
        AUTH_INFO_PATH,
        serviceName,
        accountName
      );
    } catch {
      if (allowPlaintext) {
        return await FilePersistence.create(AUTH_INFO_PATH);
      }
      throw new Error(
        "Unable to use secure storage for auth info on Linux. Install libsecret-1-dev or set TEAMS_MCP_ALLOW_PLAINTEXT_CACHE=true."
      );
    }
  }

  // Unsupported platform: fall back to plaintext file
  return await FilePersistence.create(AUTH_INFO_PATH);
}

function bufferToString(data: string | Buffer): string {
  return typeof data === "string" ? data : data.toString("utf8");
}

/**
 * Reads auth info from OS-native secure storage (Keychain, DPAPI, or libsecret).
 * Returns undefined if nothing is stored or read fails.
 */
export async function readAuthInfoSecure(): Promise<string | undefined> {
  try {
    const persistence = await createAuthInfoPersistence();
    const data = await persistence.load();
    if (data === undefined || data === null) return undefined;
    return bufferToString(data);
  } catch {
    return undefined;
  }
}

/**
 * Writes auth info to OS-native secure storage.
 */
export async function writeAuthInfoSecure(contents: string): Promise<void> {
  const persistence = await createAuthInfoPersistence();
  await persistence.save(contents);
}

/**
 * Deletes auth info from OS-native secure storage (e.g. on logout).
 */
export async function deleteAuthInfoSecure(): Promise<void> {
  try {
    const persistence = await createAuthInfoPersistence();
    await persistence.delete();
  } catch {
    // Ignore if nothing was stored
  }
}

export { AUTH_INFO_PATH };

import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { dirname, join } from "node:path";

// Auth info is stored in OS secure storage (Keychain / DPAPI / libsecret)
// so we don't keep a plaintext file with account/scopes metadata.
const AUTH_INFO_PATH = join(homedir(), ".teams-mcp-auth.json");

type PersistenceLike = {
  load: () => Promise<string | Buffer | null>;
  save: (contents: string) => Promise<void>;
  delete: () => Promise<boolean | void>;
};

/** Plaintext file persistence used when native secure storage is unavailable or disabled. */
async function createPlaintextPersistence(): Promise<PersistenceLike> {
  return {
    async load() {
      try {
        const data = await fs.readFile(AUTH_INFO_PATH, "utf8");
        return data;
      } catch {
        return null;
      }
    },
    async save(contents: string) {
      await fs.mkdir(dirname(AUTH_INFO_PATH), { recursive: true });
      await fs.writeFile(AUTH_INFO_PATH, contents, "utf8");
    },
    async delete() {
      try {
        await fs.unlink(AUTH_INFO_PATH);
        return true;
      } catch {
        return false;
      }
    },
  };
}

async function createAuthInfoPersistence(): Promise<PersistenceLike> {
  const allowPlaintext =
    process.env.TEAMS_MCP_ALLOW_PLAINTEXT_CACHE === "true";
  const platform = process.platform;
  const serviceName = "teams-mcp";
  const accountName = "AuthInfo";

  // When plaintext is explicitly allowed, skip loading native module (avoids libsecret etc. on Linux).
  if (allowPlaintext) {
    return createPlaintextPersistence();
  }

  try {
    const {
      DataProtectionScope,
      FilePersistence,
      FilePersistenceWithDataProtection,
      KeychainPersistence,
      LibSecretPersistence,
    } = await import("@azure/msal-node-extensions");

    if (platform === "darwin") {
      return (await KeychainPersistence.create(
        AUTH_INFO_PATH,
        serviceName,
        accountName
      )) as PersistenceLike;
    }

    if (platform === "win32") {
      return (await FilePersistenceWithDataProtection.create(
        AUTH_INFO_PATH,
        DataProtectionScope.CurrentUser
      )) as PersistenceLike;
    }

    if (platform === "linux") {
      try {
        return (await LibSecretPersistence.create(
          AUTH_INFO_PATH,
          serviceName,
          accountName
        )) as PersistenceLike;
      } catch (err) {
        console.error(
          "Secure storage (libsecret) unavailable, falling back to plaintext file:",
          err
        );
        return createPlaintextPersistence();
      }
    }

    // Unsupported platform: fall back to plaintext file
    return (await FilePersistence.create(AUTH_INFO_PATH)) as PersistenceLike;
  } catch (err) {
    console.error(
      "Could not load or use native secure storage, falling back to plaintext file:",
      err
    );
    return createPlaintextPersistence();
  }
}

/** True if the error indicates a missing entry and can be ignored on delete. */
function isNotFoundError(err: unknown): boolean {
  const code = (err as NodeJS.ErrnoException).code;
  const name = err instanceof Error ? err.name : "";
  const message = err instanceof Error ? err.message : String(err);
  if (code === "ENOENT" || code === "ENOTFOUND") return true;
  if (name === "NotFoundError" || /NotFound/i.test(name)) return true;
  return /not found|no such file|does not exist|no entry|could not find/i.test(
    message
  );
}

function bufferToString(data: string | Buffer | null): string {
  if (data === null) return "";
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
  } catch (err) {
    if (isNotFoundError(err)) {
      return; // Nothing was stored
    }
    throw err;
  }
}

export { AUTH_INFO_PATH };

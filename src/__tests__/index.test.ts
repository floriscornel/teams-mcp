import { promises as fs } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

// Isolated temporary home directory so CLI file operations run for real
// without touching the developer's actual credentials.
const TMP_HOME = join(tmpdir(), "teams-mcp-index-test-home");
const AUTH_INFO_PATH = join(TMP_HOME, ".msgraph-mcp-auth.json");
const TOKEN_CACHE_PATH = join(TMP_HOME, ".teams-mcp-token-cache.json");

vi.mock("node:os", async (importOriginal) => {
  const actual = await importOriginal<typeof import("node:os")>();
  return {
    ...actual,
    homedir: () => TMP_HOME,
  };
});

const msalMocks = vi.hoisted(() => ({
  acquireTokenByDeviceCode: vi.fn(),
  acquireTokenInteractive: vi.fn(),
}));

const openMock = vi.hoisted(() => vi.fn());

vi.mock("open", () => ({ default: openMock }));

vi.mock("@azure/msal-node", () => ({
  PublicClientApplication: vi.fn(),
}));

// The global test setup mocks node:fs with bare spies; these tests exercise
// real CLI file operations in a temporary home, so restore the actual module.
vi.mock("node:fs", async (importOriginal) => importOriginal());

let consoleLog: ReturnType<typeof vi.spyOn>;
let consoleError: ReturnType<typeof vi.spyOn>;

beforeEach(async () => {
  vi.resetModules();
  await fs.mkdir(TMP_HOME, { recursive: true });
  consoleLog = vi.spyOn(console, "log").mockImplementation(() => undefined);
  consoleError = vi.spyOn(console, "error").mockImplementation(() => undefined);

  // The global setup's afterEach calls vi.resetAllMocks(), which strips
  // implementations registered in module factories — reapply it here.
  // A `function` implementation is required so `new PublicClientApplication()` works.
  const { PublicClientApplication } = await import("@azure/msal-node");
  vi.mocked(PublicClientApplication).mockImplementation(function () {
    return {
      acquireTokenByDeviceCode: msalMocks.acquireTokenByDeviceCode,
      acquireTokenInteractive: msalMocks.acquireTokenInteractive,
    } as never;
  });
});

afterEach(async () => {
  consoleLog.mockRestore();
  consoleError.mockRestore();
  await fs.rm(TMP_HOME, { recursive: true, force: true });
});

async function importIndex(argv: string[]) {
  vi.resetModules();
  process.argv = ["node", "index.js", ...argv];
  const mod = await import("../index.js");
  await mod.mainPromise;
}

function authInfoFixture(overrides: Record<string, unknown> = {}) {
  return JSON.stringify({
    clientId: "test-client-id",
    authenticated: true,
    timestamp: new Date().toISOString(),
    expiresAt: new Date(Date.now() + 3_600_000).toISOString(),
    account: "test@example.com",
    grantedScopes: ["User.Read"],
    ...overrides,
  });
}

describe("MCP Server CLI", () => {
  describe("help", () => {
    it("prints usage information", async () => {
      await importIndex(["--help"]);

      expect(consoleLog).toHaveBeenCalledWith("Microsoft Graph MCP Server");
      expect(consoleLog).toHaveBeenCalledWith("Usage:");
      expect(consoleLog).toHaveBeenCalledWith(expect.stringContaining("authenticate"));
      expect(consoleLog).toHaveBeenCalledWith(expect.stringContaining("TEAMS_MCP_READ_ONLY"));
    });

    it("handles the help command and flag variants", async () => {
      await importIndex(["help"]);
      expect(consoleLog).toHaveBeenCalledWith("Usage:");

      consoleLog.mockClear();
      await importIndex(["-h"]);
      expect(consoleLog).toHaveBeenCalledWith("Usage:");
    });
  });

  describe("unknown command", () => {
    it("exits with an error", async () => {
      const exit = vi.spyOn(process, "exit").mockImplementation((() => undefined) as never);

      await importIndex(["bogus"]);

      expect(consoleError).toHaveBeenCalledWith("Unknown command: bogus");
      expect(exit).toHaveBeenCalledWith(1);
      exit.mockRestore();
    });
  });

  describe("check", () => {
    it("reports authentication details when credentials exist", async () => {
      await fs.writeFile(
        AUTH_INFO_PATH,
        authInfoFixture({ grantedScopes: ["Chat.ReadWrite", "User.Read"] })
      );

      await importIndex(["check"]);

      expect(consoleLog).toHaveBeenCalledWith("✅ Authentication found");
      expect(consoleLog).toHaveBeenCalledWith("👤 Account: test@example.com");
      expect(consoleLog).toHaveBeenCalledWith("🔒 Scope mode: full access");
    });

    it("reports read-only scope mode", async () => {
      await fs.writeFile(
        AUTH_INFO_PATH,
        authInfoFixture({ grantedScopes: ["User.Read", "Chat.Read"] })
      );

      await importIndex(["check"]);

      expect(consoleLog).toHaveBeenCalledWith("🔒 Scope mode: read-only");
    });

    it("reports not authenticated when no credentials exist", async () => {
      await importIndex(["check"]);

      expect(consoleLog).toHaveBeenCalledWith("❌ No authentication found");
    });
  });

  describe("logout", () => {
    it("removes the stored credentials", async () => {
      await fs.writeFile(AUTH_INFO_PATH, authInfoFixture());
      await fs.writeFile(TOKEN_CACHE_PATH, "{}");

      await importIndex(["logout"]);

      await expect(fs.access(AUTH_INFO_PATH)).rejects.toThrow();
      await expect(fs.access(TOKEN_CACHE_PATH)).rejects.toThrow();
      expect(consoleLog).toHaveBeenCalledWith("✅ Successfully logged out");
    });

    it("succeeds even when no credentials exist", async () => {
      await importIndex(["logout"]);

      expect(consoleLog).toHaveBeenCalledWith("✅ Successfully logged out");
    });
  });

  describe("authenticate", () => {
    it("runs the device code flow and stores credentials", async () => {
      msalMocks.acquireTokenByDeviceCode.mockResolvedValue({
        account: { username: "test@example.com" },
        scopes: ["User.Read", "Chat.ReadWrite"],
        expiresOn: new Date(Date.now() + 3_600_000),
      });

      await importIndex(["authenticate"]);

      expect(msalMocks.acquireTokenByDeviceCode).toHaveBeenCalledTimes(1);
      expect(consoleLog).toHaveBeenCalledWith(expect.stringContaining("Authentication successful"));

      const stored = JSON.parse(await fs.readFile(AUTH_INFO_PATH, "utf8"));
      expect(stored).toMatchObject({
        authenticated: true,
        account: "test@example.com",
        grantedScopes: ["User.Read", "Chat.ReadWrite"],
      });
    });

    it("passes read-only scopes when --read-only is given", async () => {
      msalMocks.acquireTokenByDeviceCode.mockResolvedValue({
        account: { username: "test@example.com" },
        scopes: ["User.Read"],
        expiresOn: new Date(Date.now() + 3_600_000),
      });

      await importIndex(["authenticate", "--read-only"]);

      const { scopes } = msalMocks.acquireTokenByDeviceCode.mock.calls[0][0];
      expect(scopes).not.toContain("Chat.ReadWrite");
      expect(scopes).not.toContain("ChannelMessage.Send");
      expect(scopes).toContain("User.Read");
    });
    it("logs the device code instructions", async () => {
      msalMocks.acquireTokenByDeviceCode.mockImplementation(async ({ deviceCodeCallback }) => {
        deviceCodeCallback({
          verificationUri: "https://microsoft.com/devicelogin",
          userCode: "ABC123",
        });
        return null;
      });

      await importIndex(["authenticate"]);

      expect(consoleLog).toHaveBeenCalledWith("🔑 Enter code: ABC123");
      await expect(fs.access(AUTH_INFO_PATH)).rejects.toThrow();
    });

    it("runs the interactive flow when --interactive is given", async () => {
      msalMocks.acquireTokenInteractive.mockImplementation(async ({ openBrowser }) => {
        await openBrowser("https://login.microsoftonline.com/authorize");
        return {
          account: { username: "test@example.com" },
          scopes: ["User.Read"],
          expiresOn: new Date(Date.now() + 3_600_000),
        };
      });

      await importIndex(["authenticate", "--interactive"]);

      expect(msalMocks.acquireTokenByDeviceCode).not.toHaveBeenCalled();
      expect(openMock).toHaveBeenCalledWith("https://login.microsoftonline.com/authorize");
      expect(consoleLog).toHaveBeenCalledWith("🔁 Flow: interactive (auth code + PKCE)");

      const stored = JSON.parse(await fs.readFile(AUTH_INFO_PATH, "utf8"));
      expect(stored).toMatchObject({ authenticated: true, account: "test@example.com" });
      if (process.platform !== "win32") {
        expect((await fs.stat(AUTH_INFO_PATH)).mode & 0o777).toBe(0o600);
      }
    });

    it.each([
      ["AADSTS50020: user not found", "User account not in tenant"],
      ["AADSTS65001: consent required", "Admin consent required"],
      ["AADSTS530032: blocked by CA", "Blocked by Conditional Access"],
    ])("explains %s and exits", async (message, expected) => {
      const exit = vi.spyOn(process, "exit").mockImplementation((() => undefined) as never);
      msalMocks.acquireTokenByDeviceCode.mockRejectedValue(new Error(message));

      await importIndex(["authenticate"]);

      expect(consoleError).toHaveBeenCalledWith(expect.stringContaining(expected));
      expect(exit).toHaveBeenCalledWith(1);
      exit.mockRestore();
    });

    it("suggests --interactive when Conditional Access blocks the device code flow", async () => {
      const exit = vi.spyOn(process, "exit").mockImplementation((() => undefined) as never);
      msalMocks.acquireTokenByDeviceCode.mockRejectedValue(new Error("AADSTS530032: blocked"));

      await importIndex(["authenticate"]);

      expect(consoleError).toHaveBeenCalledWith(expect.stringContaining("--interactive"));
      exit.mockRestore();
    });

    it("points to the tenant policy when Conditional Access blocks the interactive flow", async () => {
      const exit = vi.spyOn(process, "exit").mockImplementation((() => undefined) as never);
      msalMocks.acquireTokenInteractive.mockRejectedValue(new Error("AADSTS530032: blocked"));

      await importIndex(["authenticate", "--interactive"]);

      expect(consoleError).toHaveBeenCalledWith(
        expect.stringContaining("Blocked by Conditional Access")
      );
      expect(consoleError).toHaveBeenCalledWith(expect.stringContaining("sign-in logs"));
      expect(consoleError).not.toHaveBeenCalledWith(expect.stringContaining("--interactive"));
      exit.mockRestore();
    });

    it("reports unknown errors from the interactive flow and exits", async () => {
      const exit = vi.spyOn(process, "exit").mockImplementation((() => undefined) as never);
      msalMocks.acquireTokenInteractive.mockRejectedValue("user cancelled");

      await importIndex(["authenticate", "--interactive"]);

      expect(consoleError).toHaveBeenCalledWith("\n❌ Authentication failed:", "user cancelled");
      expect(exit).toHaveBeenCalledWith(1);
      exit.mockRestore();
    });
  });
});

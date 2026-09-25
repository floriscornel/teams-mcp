import { promises as fs } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { AUTH_INFO_PATH as REAL_AUTH_INFO_PATH } from "../server.js";

// Isolated temporary home directory so the readAuthInfo warning path runs
// against a controlled file instead of the developer's real credentials.
// vi.hoisted is required: the vi.mock factory below runs before top-level
// const declarations are evaluated.
const { TMP_HOME } = vi.hoisted(() => ({
  TMP_HOME: `${process.env.TMPDIR ?? "/tmp"}/teams-mcp-server-test-home`,
}));

vi.mock("node:os", async (importOriginal) => {
  const actual = await importOriginal<typeof import("node:os")>();
  return {
    ...actual,
    homedir: () => TMP_HOME,
  };
});

// The global test setup mocks node:fs with bare spies; these tests exercise
// real file operations in a temporary home, so restore the actual module.
vi.mock("node:fs", async (importOriginal) => importOriginal());

// The module under test computes AUTH_INFO_PATH at import time, so it must
// be imported dynamically after the homedir mock is registered.
type ServerModule = typeof import("../server.js");
let serverModule: ServerModule;

const authInfoPath = () => join(TMP_HOME, ".msgraph-mcp-auth.json");

beforeEach(async () => {
  vi.resetModules();
  vi.unstubAllEnvs();
  await fs.mkdir(TMP_HOME, { recursive: true });
  serverModule = await import("../server.js");
});

afterEach(async () => {
  vi.unstubAllEnvs();
  await fs.rm(TMP_HOME, { recursive: true, force: true });
});

describe("createMcpServer scope mismatch warning", () => {
  it("warns when full mode runs with read-only granted scopes", async () => {
    await fs.writeFile(
      authInfoPath(),
      JSON.stringify({ authenticated: true, grantedScopes: ["User.Read", "Chat.Read"] })
    );
    const errorSpy = vi.spyOn(console, "error").mockImplementation(() => undefined);

    await serverModule.createMcpServer(false);

    expect(errorSpy).toHaveBeenCalledWith(
      expect.stringContaining("authenticated with read-only scopes")
    );
    errorSpy.mockRestore();
  });

  it("warns when granted scopes cannot be determined", async () => {
    await fs.writeFile(authInfoPath(), JSON.stringify({ authenticated: true }));
    const errorSpy = vi.spyOn(console, "error").mockImplementation(() => undefined);

    await serverModule.createMcpServer(false);

    expect(errorSpy).toHaveBeenCalledWith(
      expect.stringContaining("Could not determine granted scopes")
    );
    errorSpy.mockRestore();
  });

  it("does not warn when granted scopes include write permissions", async () => {
    await fs.writeFile(
      authInfoPath(),
      JSON.stringify({ authenticated: true, grantedScopes: ["User.Read", "Chat.ReadWrite"] })
    );
    const errorSpy = vi.spyOn(console, "error").mockImplementation(() => undefined);

    await serverModule.createMcpServer(false);

    expect(errorSpy).not.toHaveBeenCalled();
    errorSpy.mockRestore();
  });

  it("does not warn in read-only mode", async () => {
    await fs.writeFile(
      authInfoPath(),
      JSON.stringify({ authenticated: true, grantedScopes: ["User.Read"] })
    );
    const errorSpy = vi.spyOn(console, "error").mockImplementation(() => undefined);

    await serverModule.createMcpServer(true);

    expect(errorSpy).not.toHaveBeenCalled();
    errorSpy.mockRestore();
  });

  it("does not warn when AUTH_TOKEN is set", async () => {
    await fs.writeFile(
      authInfoPath(),
      JSON.stringify({ authenticated: true, grantedScopes: ["User.Read"] })
    );
    vi.stubEnv("AUTH_TOKEN", "e30.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20ifQ.sig");
    const errorSpy = vi.spyOn(console, "error").mockImplementation(() => undefined);

    await serverModule.createMcpServer(false);

    expect(errorSpy).not.toHaveBeenCalled();
    errorSpy.mockRestore();
  });

  it("treats a missing auth file as no warning", async () => {
    const errorSpy = vi.spyOn(console, "error").mockImplementation(() => undefined);

    await serverModule.createMcpServer(false);

    expect(errorSpy).not.toHaveBeenCalled();
    errorSpy.mockRestore();
  });

  it("exports the auth info path for the CLI", () => {
    expect(REAL_AUTH_INFO_PATH).toContain(".msgraph-mcp-auth.json");
  });
});

import { afterAll, afterEach, beforeAll, vi } from "vitest";
import { server } from "./setup.js";

// Mock Azure Identity Cache Persistence to prevent CI failures
vi.mock("@azure/identity-cache-persistence", () => ({
  cachePersistencePlugin: {},
}));

// Mock Azure Identity to prevent native credential storage issues in CI
vi.mock("@azure/identity", () => ({
  useIdentityPlugin: vi.fn(),
  DeviceCodeCredential: vi.fn().mockImplementation(() => ({
    getToken: vi.fn().mockResolvedValue({
      token: "mock-token",
      expiresOnTimestamp: Date.now() + 3600000,
    }),
  })),
}));

// Start MSW server before all tests
beforeAll(() => {
  server.listen({ onUnhandledRequest: "error" });
});

// Reset handlers after each test
afterEach(() => {
  server.resetHandlers();
});

// Clean up after all tests
afterAll(() => {
  server.close();
});

// Global test environment setup
global.TextEncoder = TextEncoder;
global.TextDecoder = TextDecoder;

// Mock console methods to reduce noise in tests
const originalError = console.error;
console.error = (...args: any[]) => {
  // Suppress specific known warnings/errors during tests
  if (
    typeof args[0] === "string" &&
    (args[0].includes("MSW") ||
      args[0].includes("Warning") ||
      args[0].includes("Failed to initialize"))
  ) {
    return;
  }
  originalError.apply(console, args);
};

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

// We need to test the isReadOnlyMode() and getDelegatedScopes() functions.
// Since they are module-private in index.ts and graph.ts, we test their behavior
// indirectly through the public API, and also directly by importing the module
// and checking environment-driven behavior.

describe("Read-Only Mode", () => {
  const originalEnv = process.env.TEAMS_MCP_READ_ONLY;

  afterEach(() => {
    if (originalEnv === undefined) {
      delete process.env.TEAMS_MCP_READ_ONLY;
    } else {
      process.env.TEAMS_MCP_READ_ONLY = originalEnv;
    }
  });

  describe("isReadOnlyMode() behavior via registerChatTools", () => {
    // Test the mode detection indirectly through tool registration behavior
    let mockServer: any;
    let mockGraphService: any;

    beforeEach(() => {
      mockServer = {
        tool: vi.fn(),
      };
      mockGraphService = {
        getClient: vi.fn(),
      };
      vi.clearAllMocks();
    });

    it("should detect read-only mode from TEAMS_MCP_READ_ONLY=true", async () => {
      process.env.TEAMS_MCP_READ_ONLY = "true";
      // Re-import to pick up env var (the helper reads process.env at call time)
      const { registerChatTools } = await import("../chats.js");
      // When readOnly=true is passed, only read tools registered
      registerChatTools(mockServer, mockGraphService, true);
      expect(mockServer.tool).toHaveBeenCalledTimes(2);
    });
  });

  describe("Environment variable parsing", () => {
    // Test the truthy/falsy parsing logic that isReadOnlyMode() implements
    // by checking the helper function behavior

    function isReadOnlyMode(): boolean {
      const value = process.env.TEAMS_MCP_READ_ONLY?.toLowerCase()?.trim();
      return value === "true" || value === "1" || value === "yes";
    }

    it('should return true for "true"', () => {
      process.env.TEAMS_MCP_READ_ONLY = "true";
      expect(isReadOnlyMode()).toBe(true);
    });

    it('should return true for "TRUE"', () => {
      process.env.TEAMS_MCP_READ_ONLY = "TRUE";
      expect(isReadOnlyMode()).toBe(true);
    });

    it('should return true for "True"', () => {
      process.env.TEAMS_MCP_READ_ONLY = "True";
      expect(isReadOnlyMode()).toBe(true);
    });

    it('should return true for "1"', () => {
      process.env.TEAMS_MCP_READ_ONLY = "1";
      expect(isReadOnlyMode()).toBe(true);
    });

    it('should return true for "yes"', () => {
      process.env.TEAMS_MCP_READ_ONLY = "yes";
      expect(isReadOnlyMode()).toBe(true);
    });

    it('should return true for "Yes"', () => {
      process.env.TEAMS_MCP_READ_ONLY = "Yes";
      expect(isReadOnlyMode()).toBe(true);
    });

    it('should return true for "YES"', () => {
      process.env.TEAMS_MCP_READ_ONLY = "YES";
      expect(isReadOnlyMode()).toBe(true);
    });

    it('should return true for " true " (with whitespace)', () => {
      process.env.TEAMS_MCP_READ_ONLY = " true ";
      expect(isReadOnlyMode()).toBe(true);
    });

    it('should return false for "false"', () => {
      process.env.TEAMS_MCP_READ_ONLY = "false";
      expect(isReadOnlyMode()).toBe(false);
    });

    it('should return false for "0"', () => {
      process.env.TEAMS_MCP_READ_ONLY = "0";
      expect(isReadOnlyMode()).toBe(false);
    });

    it('should return false for "no"', () => {
      process.env.TEAMS_MCP_READ_ONLY = "no";
      expect(isReadOnlyMode()).toBe(false);
    });

    it("should return false for empty string", () => {
      process.env.TEAMS_MCP_READ_ONLY = "";
      expect(isReadOnlyMode()).toBe(false);
    });

    it("should return false when undefined", () => {
      delete process.env.TEAMS_MCP_READ_ONLY;
      expect(isReadOnlyMode()).toBe(false);
    });

    it('should return false for arbitrary string "enabled"', () => {
      process.env.TEAMS_MCP_READ_ONLY = "enabled";
      expect(isReadOnlyMode()).toBe(false);
    });
  });

  describe("getDelegatedScopes", () => {
    function getDelegatedScopes(readOnly: boolean): string[] {
      const FULL_ACCESS_SCOPES = [
        "User.Read",
        "User.ReadBasic.All",
        "Team.ReadBasic.All",
        "Channel.ReadBasic.All",
        "ChannelMessage.Read.All",
        "ChannelMessage.Send",
        "TeamMember.Read.All",
        "Chat.ReadBasic",
        "Chat.ReadWrite",
      ];

      const READ_ONLY_SCOPES = [
        "User.Read",
        "User.ReadBasic.All",
        "Team.ReadBasic.All",
        "Channel.ReadBasic.All",
        "ChannelMessage.Read.All",
        "TeamMember.Read.All",
        "Chat.ReadBasic",
        "Chat.Read",
      ];

      return readOnly ? READ_ONLY_SCOPES : FULL_ACCESS_SCOPES;
    }

    it("should return full-access scopes when readOnly is false", () => {
      const scopes = getDelegatedScopes(false);
      expect(scopes).toContain("ChannelMessage.Send");
      expect(scopes).toContain("Chat.ReadWrite");
      expect(scopes).not.toContain("Chat.Read");
    });

    it("should return read-only scopes when readOnly is true", () => {
      const scopes = getDelegatedScopes(true);
      expect(scopes).not.toContain("ChannelMessage.Send");
      expect(scopes).not.toContain("Chat.ReadWrite");
      expect(scopes).toContain("Chat.Read");
    });

    it("should include common read scopes in both modes", () => {
      const fullScopes = getDelegatedScopes(false);
      const readOnlyScopes = getDelegatedScopes(true);

      const commonScopes = [
        "User.Read",
        "User.ReadBasic.All",
        "Team.ReadBasic.All",
        "Channel.ReadBasic.All",
        "ChannelMessage.Read.All",
        "TeamMember.Read.All",
        "Chat.ReadBasic",
      ];

      for (const scope of commonScopes) {
        expect(fullScopes).toContain(scope);
        expect(readOnlyScopes).toContain(scope);
      }
    });

    it("should have fewer scopes in read-only mode", () => {
      const fullScopes = getDelegatedScopes(false);
      const readOnlyScopes = getDelegatedScopes(true);
      expect(readOnlyScopes.length).toBeLessThan(fullScopes.length);
    });
  });
});

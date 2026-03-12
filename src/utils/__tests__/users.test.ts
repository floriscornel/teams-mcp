import { beforeEach, describe, expect, it, vi } from "vitest";
import type { GraphService } from "../../services/graph.js";
import type {
  ChannelMention,
  ChannelMentionMapping,
  UserMention,
  UserMentionMapping,
} from "../../types/mentions.js";
import {
  isUserMentionMapping,
  isChannelMentionMapping,
  convertMentionInputsToMappings,
} from "../../types/mentions.js";
import {
  getUserByEmail,
  getUserById,
  parseMentions,
  processMentionsInHtml,
  searchUsers,
} from "../users.js";

const mockGraphService = {
  getClient: vi.fn(),
} as unknown as GraphService;

const mockClient = {
  api: vi.fn(),
};

describe("User Utilities", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    (mockGraphService.getClient as any).mockResolvedValue(mockClient);
  });

  describe("searchUsers", () => {
    it("should search users by display name", async () => {
      const mockUsers = [
        { id: "1", displayName: "John Doe", userPrincipalName: "john.doe@company.com" },
        { id: "2", displayName: "John Smith", userPrincipalName: "john.smith@company.com" },
      ];

      mockClient.api.mockReturnValue({
        get: vi.fn().mockResolvedValue({ value: mockUsers }),
      });

      const result = await searchUsers(mockGraphService, "John", 10);

      expect(result).toEqual([
        { id: "1", displayName: "John Doe", userPrincipalName: "john.doe@company.com" },
        { id: "2", displayName: "John Smith", userPrincipalName: "john.smith@company.com" },
      ]);

      expect(mockClient.api).toHaveBeenCalledWith(
        "/users?$filter=startswith(displayName,'John') or startswith(userPrincipalName,'John')&$top=10&$select=id,displayName,userPrincipalName"
      );
    });

    it("should return empty array when no users found", async () => {
      mockClient.api.mockReturnValue({
        get: vi.fn().mockResolvedValue({ value: [] }),
      });

      const result = await searchUsers(mockGraphService, "NonExistent", 10);
      expect(result).toEqual([]);
    });

    it("should handle errors gracefully", async () => {
      mockClient.api.mockReturnValue({
        get: vi.fn().mockRejectedValue(new Error("Graph API error")),
      });

      const consoleSpy = vi.spyOn(console, "error").mockImplementation(() => {
        // Mock implementation - do nothing
      });
      const result = await searchUsers(mockGraphService, "John", 10);

      expect(result).toEqual([]);
      expect(consoleSpy).toHaveBeenCalledWith("Error searching users:", expect.any(Error));

      consoleSpy.mockRestore();
    });
  });

  describe("getUserByEmail", () => {
    it("should get user by email", async () => {
      const mockUser = {
        id: "1",
        displayName: "John Doe",
        userPrincipalName: "john.doe@company.com",
      };

      mockClient.api.mockReturnValue({
        get: vi.fn().mockResolvedValue(mockUser),
      });

      const result = await getUserByEmail(mockGraphService, "john.doe@company.com");

      expect(result).toEqual({
        id: "1",
        displayName: "John Doe",
        userPrincipalName: "john.doe@company.com",
      });

      expect(mockClient.api).toHaveBeenCalledWith("/users/john.doe@company.com");
    });

    it("should return null when user not found", async () => {
      mockClient.api.mockReturnValue({
        get: vi.fn().mockRejectedValue(new Error("User not found")),
      });

      const result = await getUserByEmail(mockGraphService, "nonexistent@company.com");
      expect(result).toBeNull();
    });
  });

  describe("getUserById", () => {
    it("should get user by ID", async () => {
      const mockUser = {
        id: "1",
        displayName: "John Doe",
        userPrincipalName: "john.doe@company.com",
      };

      mockClient.api.mockReturnValue({
        select: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      });

      const result = await getUserById(mockGraphService, "1");

      expect(result).toEqual({
        id: "1",
        displayName: "John Doe",
        userPrincipalName: "john.doe@company.com",
      });

      expect(mockClient.api).toHaveBeenCalledWith("/users/1");
    });

    it("should return null when user not found", async () => {
      mockClient.api.mockReturnValue({
        select: vi.fn().mockReturnValue({
          get: vi.fn().mockRejectedValue(new Error("User not found")),
        }),
      });

      const result = await getUserById(mockGraphService, "nonexistent");
      expect(result).toBeNull();
    });
  });

  describe("parseMentions", () => {
    it("should parse simple @mentions", async () => {
      const mockUsers = [
        { id: "1", displayName: "John Doe", userPrincipalName: "john.doe@company.com" },
      ];

      mockClient.api.mockReturnValue({
        get: vi.fn().mockResolvedValue({ value: mockUsers }),
      });

      const result = await parseMentions("Hello @john.doe how are you?", mockGraphService);

      expect(result).toEqual([
        {
          mention: "john.doe",
          users: [{ id: "1", displayName: "John Doe", userPrincipalName: "john.doe@company.com" }],
        },
      ]);
    });

    it("should parse email @mentions", async () => {
      const mockUser = {
        id: "1",
        displayName: "John Doe",
        userPrincipalName: "john.doe@company.com",
      };

      mockClient.api.mockReturnValue({
        get: vi.fn().mockResolvedValue(mockUser),
      });

      const result = await parseMentions("Hello @john.doe@company.com", mockGraphService);

      expect(result).toEqual([
        {
          mention: "john.doe@company.com",
          users: [{ id: "1", displayName: "John Doe", userPrincipalName: "john.doe@company.com" }],
        },
      ]);
    });

    it("should parse quoted @mentions", async () => {
      const mockUsers = [
        { id: "1", displayName: "John Doe", userPrincipalName: "john.doe@company.com" },
      ];

      mockClient.api.mockReturnValue({
        get: vi.fn().mockResolvedValue({ value: mockUsers }),
      });

      const result = await parseMentions('Hello @"John Doe", how are you?', mockGraphService);

      expect(result).toEqual([
        {
          mention: "John Doe",
          users: [{ id: "1", displayName: "John Doe", userPrincipalName: "john.doe@company.com" }],
        },
      ]);
    });

    it("should handle multiple @mentions", async () => {
      mockClient.api
        .mockReturnValueOnce({
          get: vi.fn().mockResolvedValue({
            value: [
              { id: "1", displayName: "John Doe", userPrincipalName: "john.doe@company.com" },
            ],
          }),
        })
        .mockReturnValueOnce({
          get: vi.fn().mockResolvedValue({
            value: [
              { id: "2", displayName: "Jane Smith", userPrincipalName: "jane.smith@company.com" },
            ],
          }),
        });

      const result = await parseMentions("Hello @john.doe and @jane", mockGraphService);

      expect(result).toHaveLength(2);
      expect(result[0].mention).toBe("john.doe");
      expect(result[1].mention).toBe("jane");
    });

    it("should return empty array when no mentions found", async () => {
      const result = await parseMentions("Hello world, no mentions here!", mockGraphService);
      expect(result).toEqual([]);
    });
  });

  describe("processMentionsInHtml", () => {
    it("should process @mentions in HTML content", () => {
      const html = "<p>Hello @john.doe, how are you?</p>";
      const mentionMappings = [{ mention: "john.doe", userId: "1", displayName: "John Doe" }];

      const result = processMentionsInHtml(html, mentionMappings);

      expect(result.content).toBe('<p>Hello <at id="0">John Doe</at>, how are you?</p>');
      expect(result.mentions).toEqual([
        {
          id: 0,
          mentionText: "John Doe",
          mentioned: { user: { id: "1" } },
        },
      ]);
    });

    it("should process quoted @mentions in HTML content", () => {
      const html = '<p>Hello @"John Doe", how are you?</p>';
      const mentionMappings = [{ mention: "John Doe", userId: "1", displayName: "John Doe" }];

      const result = processMentionsInHtml(html, mentionMappings);

      expect(result.content).toBe('<p>Hello <at id="0">John Doe</at>, how are you?</p>');
      expect(result.mentions).toEqual([
        {
          id: 0,
          mentionText: "John Doe",
          mentioned: { user: { id: "1" } },
        },
      ]);
    });

    it("should handle multiple @mentions", () => {
      const html = "<p>Hello @john.doe and @jane.smith!</p>";
      const mentionMappings = [
        { mention: "john.doe", userId: "1", displayName: "John Doe" },
        { mention: "jane.smith", userId: "2", displayName: "Jane Smith" },
      ];

      const result = processMentionsInHtml(html, mentionMappings);

      expect(result.content).toBe(
        '<p>Hello <at id="0">John Doe</at> and <at id="1">Jane Smith</at>!</p>'
      );
      expect(result.mentions).toHaveLength(2);
    });

    it("should return unchanged content when no mappings provided", () => {
      const html = "<p>Hello @john.doe, how are you?</p>";
      const result = processMentionsInHtml(html, []);

      expect(result.content).toBe(html);
      expect(result.mentions).toEqual([]);
    });

    describe("channel mentions", () => {
      it("should process single channel mention", () => {
        const html = "<p>Attention @General channel</p>";
        const mappings: ChannelMentionMapping[] = [
          {
            mention: "General",
            channelId: "19:abc123@thread.tacv2",
            displayName: "General",
          },
        ];

        const result = processMentionsInHtml(html, mappings);

        expect(result.content).toBe('<p>Attention <at id="0">General</at> channel</p>');
        expect(result.mentions).toHaveLength(1);

        const mention = result.mentions[0] as ChannelMention;
        expect(mention.id).toBe(0);
        expect(mention.mentionText).toBe("General");
        expect(mention.mentioned.conversation).toBeDefined();
        expect(mention.mentioned.conversation.id).toBe("19:abc123@thread.tacv2");
        expect(mention.mentioned.conversation.displayName).toBe("General");
        expect(mention.mentioned.conversation.conversationIdentityType).toBe("channel");
      });

      it("should process multiple channel mentions", () => {
        const html = "<p>Check @General and @Engineering channels</p>";
        const mappings: ChannelMentionMapping[] = [
          {
            mention: "General",
            channelId: "19:general@thread.tacv2",
            displayName: "General",
          },
          {
            mention: "Engineering",
            channelId: "19:engineering@thread.tacv2",
            displayName: "Engineering",
          },
        ];

        const result = processMentionsInHtml(html, mappings);

        expect(result.content).toBe(
          '<p>Check <at id="0">General</at> and <at id="1">Engineering</at> channels</p>'
        );
        expect(result.mentions).toHaveLength(2);
      });

      it("should handle quoted channel mentions", () => {
        const html = '<p>Check @"Product Team" channel</p>';
        const mappings: ChannelMentionMapping[] = [
          {
            mention: "Product Team",
            channelId: "19:product@thread.tacv2",
            displayName: "Product Team",
          },
        ];

        const result = processMentionsInHtml(html, mappings);

        expect(result.content).toBe('<p>Check <at id="0">Product Team</at> channel</p>');
      });
    });

    describe("mixed user and channel mentions", () => {
      it("should process both user and channel mentions together", () => {
        const html = "<p>@General - @John Doe will present the update</p>";
        const mappings = [
          {
            mention: "General",
            channelId: "19:abc@thread.tacv2",
            displayName: "General",
          } as ChannelMentionMapping,
          {
            mention: "John Doe",
            userId: "user-123",
            displayName: "John Doe",
          } as UserMentionMapping,
        ];

        const result = processMentionsInHtml(html, mappings);

        expect(result.content).toBe(
          '<p><at id="0">General</at> - <at id="1">John Doe</at> will present the update</p>'
        );
        expect(result.mentions).toHaveLength(2);

        // First mention should be a channel mention
        const channelMention = result.mentions[0] as ChannelMention;
        expect(channelMention.mentioned.conversation).toBeDefined();
        expect(channelMention.mentioned.conversation.conversationIdentityType).toBe("channel");

        // Second mention should be a user mention
        const userMention = result.mentions[1] as UserMention;
        expect(userMention.mentioned.user).toBeDefined();
        expect(userMention.mentioned.user.id).toBe("user-123");
      });

      it("should assign sequential IDs to mixed mentions", () => {
        const html = "<p>@alice @General @bob @Engineering</p>";
        const mappings = [
          { mention: "alice", userId: "u1", displayName: "Alice" } as UserMentionMapping,
          { mention: "General", channelId: "c1", displayName: "General" } as ChannelMentionMapping,
          { mention: "bob", userId: "u2", displayName: "Bob" } as UserMentionMapping,
          {
            mention: "Engineering",
            channelId: "c2",
            displayName: "Engineering",
          } as ChannelMentionMapping,
        ];

        const result = processMentionsInHtml(html, mappings);

        expect(result.mentions[0].id).toBe(0);
        expect(result.mentions[1].id).toBe(1);
        expect(result.mentions[2].id).toBe(2);
        expect(result.mentions[3].id).toBe(3);
      });
    });

    describe("edge cases", () => {
      it("should handle special characters in mention text", () => {
        const html = "<p>Check @Dev.Team channel</p>";
        const mappings: ChannelMentionMapping[] = [
          {
            mention: "Dev.Team",
            channelId: "19:devteam@thread.tacv2",
            displayName: "Dev.Team",
          },
        ];

        const result = processMentionsInHtml(html, mappings);

        expect(result.content).toBe('<p>Check <at id="0">Dev.Team</at> channel</p>');
      });

      it("should handle same mention appearing multiple times", () => {
        const html = "<p>@General is great. I love @General!</p>";
        const mappings: ChannelMentionMapping[] = [
          {
            mention: "General",
            channelId: "19:general@thread.tacv2",
            displayName: "General",
          },
        ];

        const result = processMentionsInHtml(html, mappings);

        expect(result.content).toBe(
          '<p><at id="0">General</at> is great. I love <at id="0">General</at>!</p>'
        );
        // Only one mention object in the array even though it appears twice
        expect(result.mentions).toHaveLength(1);
      });
    });
  });

  describe("Type Guards", () => {
    describe("isUserMentionMapping", () => {
      it("should return true for user mention mappings", () => {
        const mapping: UserMentionMapping = {
          mention: "John",
          userId: "user-123",
          displayName: "John Doe",
        };
        expect(isUserMentionMapping(mapping)).toBe(true);
      });

      it("should return false for channel mention mappings", () => {
        const mapping: ChannelMentionMapping = {
          mention: "General",
          channelId: "19:abc@thread.tacv2",
          displayName: "General",
        };
        expect(isUserMentionMapping(mapping)).toBe(false);
      });
    });

    describe("isChannelMentionMapping", () => {
      it("should return true for channel mention mappings", () => {
        const mapping: ChannelMentionMapping = {
          mention: "General",
          channelId: "19:abc@thread.tacv2",
          displayName: "General",
        };
        expect(isChannelMentionMapping(mapping)).toBe(true);
      });

      it("should return false for user mention mappings", () => {
        const mapping: UserMentionMapping = {
          mention: "John",
          userId: "user-123",
          displayName: "John Doe",
        };
        expect(isChannelMentionMapping(mapping)).toBe(false);
      });
    });
  });

  describe("convertMentionInputsToMappings", () => {
    it("should convert user mention inputs to mappings", () => {
      const inputs = [{ mention: "John", userId: "user-123" }];
      const result = convertMentionInputsToMappings(inputs);

      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({
        mention: "John",
        userId: "user-123",
        displayName: "John",
      });
    });

    it("should convert channel mention inputs to mappings", () => {
      const inputs = [{ mention: "General", channelId: "19:abc@thread.tacv2" }];
      const result = convertMentionInputsToMappings(inputs);

      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({
        mention: "General",
        channelId: "19:abc@thread.tacv2",
        displayName: "General",
      });
    });

    it("should convert mixed mention inputs to mappings", () => {
      const inputs = [
        { mention: "John", userId: "user-123" },
        { mention: "General", channelId: "19:abc@thread.tacv2" },
      ];
      const result = convertMentionInputsToMappings(inputs);

      expect(result).toHaveLength(2);
      expect(isUserMentionMapping(result[0])).toBe(true);
      expect(isChannelMentionMapping(result[1])).toBe(true);
    });

    it("should throw error for invalid mention input (neither userId nor channelId)", () => {
      const inputs = [{ mention: "Invalid" }];
      expect(() => convertMentionInputsToMappings(inputs)).toThrow(
        'Invalid mention: "Invalid" must have either userId or channelId'
      );
    });
  });
});

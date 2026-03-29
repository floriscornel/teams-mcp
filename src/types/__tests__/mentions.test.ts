import { describe, expect, it } from "vitest";
import { z } from "zod";

/**
 * Tests for the mentions Zod schema validation.
 * This recreates the schema from teams.ts to test validation rules.
 */

// Recreate the mentions schema from teams.ts for testing
const mentionsSchema = z
  .array(
    z.object({
      mention: z
        .string()
        .describe("The @mention text as it appears in the message (e.g., 'John Doe' or 'General')"),
      userId: z
        .string()
        .optional()
        .describe("Azure AD User ID for user mentions - mutually exclusive with channelId"),
      channelId: z
        .string()
        .optional()
        .describe(
          "Channel ID for channel mentions (e.g., '19:abc123@thread.tacv2') - mutually exclusive with userId"
        ),
    })
  )
  .refine(
    (mentions) =>
      mentions.every((m) => {
        const hasUserId = m.userId !== undefined && m.userId !== "";
        const hasChannelId = m.channelId !== undefined && m.channelId !== "";
        return (hasUserId && !hasChannelId) || (!hasUserId && hasChannelId);
      }),
    {
      message:
        "Invalid mention configuration: Each mention must specify exactly one of 'userId' (for user mentions) or 'channelId' (for channel mentions). " +
        "Found mention with both or neither. Check that each mention in the array has either userId OR channelId, not both and not neither.",
    }
  )
  .optional()
  .describe(
    "Array of mentions - each must specify either userId (for user mentions) or channelId (for channel mentions)"
  );

describe("Mentions Schema Validation", () => {
  describe("valid inputs", () => {
    it("should accept mention with userId only", () => {
      const input = [{ mention: "John Doe", userId: "user-123" }];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data).toEqual(input);
      }
    });

    it("should accept mention with channelId only", () => {
      const input = [{ mention: "General", channelId: "19:abc123@thread.tacv2" }];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data).toEqual(input);
      }
    });

    it("should accept mixed user and channel mentions", () => {
      const input = [
        { mention: "General", channelId: "19:abc123@thread.tacv2" },
        { mention: "John Doe", userId: "user-123" },
      ];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data).toHaveLength(2);
      }
    });

    it("should accept empty array", () => {
      const result = mentionsSchema.safeParse([]);

      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data).toEqual([]);
      }
    });

    it("should accept undefined (optional)", () => {
      const result = mentionsSchema.safeParse(undefined);

      expect(result.success).toBe(true);
    });

    it("should accept multiple user mentions", () => {
      const input = [
        { mention: "John", userId: "user-1" },
        { mention: "Jane", userId: "user-2" },
        { mention: "Bob", userId: "user-3" },
      ];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(true);
    });

    it("should accept multiple channel mentions", () => {
      const input = [
        { mention: "General", channelId: "19:general@thread.tacv2" },
        { mention: "Engineering", channelId: "19:engineering@thread.tacv2" },
      ];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(true);
    });
  });

  describe("invalid inputs", () => {
    it("should reject mention with both userId and channelId", () => {
      const input = [
        {
          mention: "Test",
          userId: "user-123",
          channelId: "19:channel@thread.tacv2",
        },
      ];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.errors[0].message).toContain("Invalid mention configuration");
        expect(result.error.errors[0].message).toContain("userId");
        expect(result.error.errors[0].message).toContain("channelId");
      }
    });

    it("should reject mention with neither userId nor channelId", () => {
      const input = [{ mention: "Test" }];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.errors[0].message).toContain("Invalid mention configuration");
      }
    });

    it("should reject if any mention in array is invalid (missing both)", () => {
      const input = [
        { mention: "Valid", userId: "user-123" },
        { mention: "Invalid" }, // missing both userId and channelId
      ];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(false);
    });

    it("should reject if any mention in array is invalid (has both)", () => {
      const input = [
        { mention: "Valid", userId: "user-123" },
        { mention: "Invalid", userId: "user-456", channelId: "19:channel@thread.tacv2" },
      ];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(false);
    });

    it("should reject mention with empty string userId", () => {
      const input = [{ mention: "Test", userId: "" }];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(false);
    });

    it("should reject mention with empty string channelId", () => {
      const input = [{ mention: "Test", channelId: "" }];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(false);
    });

    it("should reject non-array input", () => {
      const result = mentionsSchema.safeParse("not an array");

      expect(result.success).toBe(false);
    });

    it("should reject mention without mention text", () => {
      const input = [{ userId: "user-123" }];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(false);
    });
  });

  describe("edge cases", () => {
    it("should accept mention with special characters in mention text", () => {
      const input = [{ mention: "John.Doe@company.com", userId: "user-123" }];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(true);
    });

    it("should accept mention with spaces in mention text", () => {
      const input = [{ mention: "John Doe Jr.", userId: "user-123" }];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(true);
    });

    it("should accept channel mention with standard Teams channel ID format", () => {
      const input = [
        {
          mention: "General",
          channelId: "19:8a5f5e8b-7c5d-4f3e-9a2b-1c0d3e4f5a6b@thread.tacv2",
        },
      ];
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(true);
    });

    it("should handle large number of mentions", () => {
      const input = Array.from({ length: 50 }, (_, i) => ({
        mention: `User ${i}`,
        userId: `user-${i}`,
      }));
      const result = mentionsSchema.safeParse(input);

      expect(result.success).toBe(true);
    });
  });
});

/**
 * Types for handling @mentions in Teams messages
 * Supports both user mentions and channel mentions
 */

// ============================================================================
// Mention Mapping Types (input from API consumers)
// ============================================================================

/**
 * User mention mapping - maps @mention text to a user
 */
export interface UserMentionMapping {
  mention: string;
  userId: string;
  displayName: string;
}

/**
 * Channel mention mapping - maps @mention text to a channel
 */
export interface ChannelMentionMapping {
  mention: string;
  channelId: string;
  displayName: string;
}

/**
 * Combined mention input from tool parameters
 * Either userId OR channelId must be provided, not both
 */
export interface MentionInput {
  mention: string;
  userId?: string;
  channelId?: string;
}

/**
 * Union type for all mention mapping types
 */
export type MentionMapping = UserMentionMapping | ChannelMentionMapping;

// ============================================================================
// Graph API Mention Types (output to Microsoft Graph API)
// ============================================================================

/**
 * Graph API mention format for users
 * @see https://learn.microsoft.com/en-us/graph/api/resources/chatmessagemention
 */
export interface UserMention {
  id: number;
  mentionText: string;
  mentioned: {
    user: {
      id: string;
    };
  };
}

/**
 * Graph API mention format for channels
 * @see https://learn.microsoft.com/en-us/graph/api/resources/chatmessagemention
 */
export interface ChannelMention {
  id: number;
  mentionText: string;
  mentioned: {
    conversation: {
      id: string;
      displayName: string;
      conversationIdentityType: "channel";
    };
  };
}

/**
 * Union type for all Graph API mention types
 */
export type GraphMention = UserMention | ChannelMention;

// ============================================================================
// Type Guards
// ============================================================================

/**
 * Type guard to check if a mention mapping is for a user
 */
export function isUserMentionMapping(
  mapping: MentionMapping
): mapping is UserMentionMapping {
  return "userId" in mapping;
}

/**
 * Type guard to check if a mention mapping is for a channel
 */
export function isChannelMentionMapping(
  mapping: MentionMapping
): mapping is ChannelMentionMapping {
  return "channelId" in mapping;
}

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Convert MentionInput array to MentionMapping array
 * Validates that each input has either userId or channelId
 */
export function convertMentionInputsToMappings(
  inputs: MentionInput[]
): MentionMapping[] {
  return inputs.map((input) => {
    if (input.userId) {
      return {
        mention: input.mention,
        userId: input.userId,
        displayName: input.mention,
      } as UserMentionMapping;
    }
    if (input.channelId) {
      return {
        mention: input.mention,
        channelId: input.channelId,
        displayName: input.mention,
      } as ChannelMentionMapping;
    }
    throw new Error(
      `Invalid mention: "${input.mention}" must have either userId or channelId`
    );
  });
}

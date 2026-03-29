import type { GraphService } from "../services/graph.js";
import type { User } from "../types/graph.js";
import type {
  MentionMapping,
  GraphMention,
  UserMentionMapping,
  ChannelMentionMapping,
} from "../types/mentions.js";
import { isUserMentionMapping } from "../types/mentions.js";

export interface UserInfo {
  id: string;
  displayName: string;
  userPrincipalName?: string;
}

/**
 * Search for users by display name or email
 */
export async function searchUsers(
  graphService: GraphService,
  query: string,
  limit = 10
): Promise<UserInfo[]> {
  try {
    const client = await graphService.getClient();

    // Use filter query to search users by displayName or userPrincipalName
    const searchQuery = `$filter=startswith(displayName,'${query}') or startswith(userPrincipalName,'${query}')&$top=${limit}&$select=id,displayName,userPrincipalName`;

    const response = await client.api(`/users?${searchQuery}`).get();

    if (!response?.value?.length) {
      return [];
    }

    return response.value.map((user: User) => ({
      id: user.id || "",
      displayName: user.displayName || "Unknown User",
      userPrincipalName: user.userPrincipalName || undefined,
    }));
  } catch (error) {
    console.error("Error searching users:", error);
    return [];
  }
}

/**
 * Get user by exact email or UPN
 */
export async function getUserByEmail(
  graphService: GraphService,
  email: string
): Promise<UserInfo | null> {
  try {
    const client = await graphService.getClient();

    const response = await client.api(`/users/${email}`).get();

    return {
      id: response.id,
      displayName: response.displayName || "Unknown User",
      userPrincipalName: response.userPrincipalName,
    };
  } catch (_error) {
    // User not found or access denied
    return null;
  }
}

/**
 * Get user by ID
 */
export async function getUserById(
  graphService: GraphService,
  userId: string
): Promise<UserInfo | null> {
  try {
    const client = await graphService.getClient();

    const response = await client
      .api(`/users/${userId}`)
      .select("id,displayName,userPrincipalName")
      .get();

    return {
      id: response.id,
      displayName: response.displayName || "Unknown User",
      userPrincipalName: response.userPrincipalName,
    };
  } catch (_error) {
    // User not found or access denied
    return null;
  }
}

/**
 * Parse @mentions from text and return user lookup suggestions
 * @param text - Message text containing @mentions
 * @param graphService - Graph service instance
 * @returns Array of mention patterns found and suggested users
 */
export async function parseMentions(
  text: string,
  graphService: GraphService
): Promise<Array<{ mention: string; users: UserInfo[] }>> {
  // Match @mentions in the format @username, @email@domain.com, or @"User Name"
  const mentionRegex = /@(?:"([^"]+)"|([^\s@]+(?:@[^\s@]+\.[^\s@]+)?|[^\s@]+))/g;
  const mentions: Array<{ mention: string; users: UserInfo[] }> = [];

  let match: RegExpExecArray | null = mentionRegex.exec(text);
  while (match !== null) {
    const mentionText = match[1] || match[2]; // Quoted name or unquoted

    let users: UserInfo[] = [];

    // If it looks like an email, try exact lookup first
    if (mentionText.includes("@") && mentionText.includes(".")) {
      const user = await getUserByEmail(graphService, mentionText);
      if (user) {
        users = [user];
      }
    }

    // If no exact match found, search by name
    if (users.length === 0) {
      users = await searchUsers(graphService, mentionText, 5);
    }

    mentions.push({
      mention: mentionText,
      users,
    });

    match = mentionRegex.exec(text);
  }

  return mentions;
}

/**
 * Generate HTML content with @mentions converted to proper format.
 * Supports both user mentions and channel mentions.
 */
export function processMentionsInHtml(
  html: string,
  mentionMappings: MentionMapping[]
): {
  content: string;
  mentions: GraphMention[];
} {
  let processedContent = html;
  const mentions: GraphMention[] = [];

  mentionMappings.forEach((mapping, index) => {
    // Replace @mention with HTML mention format
    const mentionRegex = new RegExp(
      `@(?:"${escapeRegex(mapping.mention)}"|${escapeRegex(mapping.mention)})`,
      "g"
    );
    const mentionId = index;

    processedContent = processedContent.replace(
      mentionRegex,
      `<at id="${mentionId}">${mapping.displayName}</at>`
    );

    if (isUserMentionMapping(mapping)) {
      // User mention
      mentions.push({
        id: mentionId,
        mentionText: mapping.displayName,
        mentioned: {
          user: {
            id: mapping.userId,
          },
        },
      });
    } else {
      // Channel mention
      mentions.push({
        id: mentionId,
        mentionText: mapping.displayName,
        mentioned: {
          conversation: {
            id: mapping.channelId,
            displayName: mapping.displayName,
            conversationIdentityType: "channel",
          },
        },
      });
    }
  });

  return { content: processedContent, mentions };
}

// Re-export types for convenience
export type { MentionMapping, GraphMention, UserMentionMapping, ChannelMentionMapping };

function escapeRegex(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

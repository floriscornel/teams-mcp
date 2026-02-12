import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphService } from "../services/graph.js";
import type {
  Chat,
  ChatMessage,
  ChatSummary,
  ConversationMember,
  CreateChatPayload,
  GraphApiResponse,
  MessageSummary,
  User,
} from "../types/graph.js";
import { markdownToHtml } from "../utils/markdown.js";
import { processMentionsInHtml } from "../utils/users.js";

export function registerChatTools(server: McpServer, graphService: GraphService, readOnly = false) {
  // List user's chats
  server.tool(
    "list_chats",
    "List all recent chats (1:1 conversations and group chats) that the current user participates in. Returns chat topics, types, and participant information.",
    {},
    async () => {
      try {
        // Build query parameters
        const queryParams: string[] = ["$expand=members"];

        const queryString = queryParams.join("&");

        const client = await graphService.getClient();
        const response = (await client
          .api(`/me/chats?${queryString}`)
          .get()) as GraphApiResponse<Chat>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No chats found.",
              },
            ],
          };
        }

        const chatList: ChatSummary[] = response.value.map((chat: Chat) => ({
          id: chat.id,
          topic: chat.topic || "No topic",
          chatType: chat.chatType,
          members:
            chat.members?.map((member: ConversationMember) => member.displayName).join(", ") ||
            "No members",
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(chatList, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Get chat messages with pagination support
  server.tool(
    "get_chat_messages",
    "Retrieve recent messages from a specific chat conversation. Returns message content, sender information, and timestamps.",
    {
      chatId: z.string().describe("Chat ID (e.g. 19:meeting_Njhi..j@thread.v2"),
      limit: z
        .number()
        .min(1)
        .max(2000)
        .optional()
        .default(20)
        .describe("Number of messages to retrieve (default: 20, max: 2000)"),
      since: z.string().optional().describe("Get messages since this ISO datetime"),
      until: z.string().optional().describe("Get messages until this ISO datetime"),
      fromUser: z.string().optional().describe("Filter messages from specific user ID"),
      orderBy: z
        .enum(["createdDateTime", "lastModifiedDateTime"])
        .optional()
        .default("createdDateTime")
        .describe("Sort order"),
      descending: z
        .boolean()
        .optional()
        .default(true)
        .describe("Sort in descending order (newest first)"),
      fetchAll: z
        .boolean()
        .optional()
        .default(false)
        .describe(
          "Fetch all messages using pagination (up to limit). When true, follows @odata.nextLink to get more messages."
        ),
    },
    async ({ chatId, limit, since, until, fromUser, orderBy, descending, fetchAll }) => {
      try {
        const client = await graphService.getClient();

        // Apply defaults for parameters (in case Zod validation is bypassed)
        const effectiveLimit = limit ?? 20;
        const effectiveOrderBy = orderBy ?? "createdDateTime";
        const effectiveDescending = descending ?? true;
        const effectiveFetchAll = fetchAll ?? false;

        // Build query parameters - use smaller page size for pagination
        const pageSize = effectiveFetchAll ? 50 : Math.min(effectiveLimit, 50);
        const queryParams: string[] = [`$top=${pageSize}`];

        // Add ordering - Graph API only supports descending order for datetime fields in chat messages
        if (
          (effectiveOrderBy === "createdDateTime" || effectiveOrderBy === "lastModifiedDateTime") &&
          !effectiveDescending
        ) {
          return {
            content: [
              {
                type: "text",
                text: `❌ Error: QueryOptions to order by '${effectiveOrderBy === "createdDateTime" ? "CreatedDateTime" : "LastModifiedDateTime"}' in 'Ascending' direction is not supported.`,
              },
            ],
          };
        }

        const sortDirection = effectiveDescending ? "desc" : "asc";
        queryParams.push(`$orderby=${effectiveOrderBy} ${sortDirection}`);

        // Add filters (only user filter is supported reliably)
        const filters: string[] = [];
        if (fromUser) {
          filters.push(`from/user/id eq '${fromUser}'`);
        }

        if (filters.length > 0) {
          queryParams.push(`$filter=${filters.join(" and ")}`);
        }

        const queryString = queryParams.join("&");

        // Fetch messages with pagination support
        const allMessages: ChatMessage[] = [];
        let nextLink: string | undefined;
        let pageCount = 0;
        const maxPages = 100; // Safety limit to prevent infinite loops

        // First request
        let response = (await client
          .api(`/me/chats/${chatId}/messages?${queryString}`)
          .get()) as GraphApiResponse<ChatMessage>;

        if (response?.value) {
          allMessages.push(...response.value);
        }

        // Follow pagination if fetchAll is enabled
        if (effectiveFetchAll) {
          nextLink = response["@odata.nextLink"];

          while (nextLink && allMessages.length < effectiveLimit && pageCount < maxPages) {
            pageCount++;

            try {
              response = (await client.api(nextLink).get()) as GraphApiResponse<ChatMessage>;

              if (response?.value) {
                allMessages.push(...response.value);
              }

              nextLink = response["@odata.nextLink"];
            } catch (pageError) {
              console.error(`Error fetching page ${pageCount}:`, pageError);
              break;
            }
          }
        }

        if (allMessages.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: "No messages found in this chat with the specified filters.",
              },
            ],
          };
        }

        // Apply client-side date filtering since server-side filtering is not supported
        let filteredMessages = allMessages;

        if (since || until) {
          filteredMessages = allMessages.filter((message: ChatMessage) => {
            if (!message.createdDateTime) return true;

            const messageDate = new Date(message.createdDateTime);
            if (since) {
              const sinceDate = new Date(since);
              if (messageDate <= sinceDate) return false;
            }
            if (until) {
              const untilDate = new Date(until);
              if (messageDate >= untilDate) return false;
            }
            return true;
          });
        }

        // Apply limit after filtering
        const limitedMessages = filteredMessages.slice(0, effectiveLimit);

        const messageList: MessageSummary[] = limitedMessages.map((message: ChatMessage) => ({
          id: message.id,
          content: message.body?.content,
          from: message.from?.user?.displayName,
          createdDateTime: message.createdDateTime,
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  filters: { since, until, fromUser },
                  filteringMethod: since || until ? "client-side" : "server-side",
                  paginationEnabled: fetchAll,
                  pagesRetrieved: pageCount + 1,
                  totalRetrieved: allMessages.length,
                  totalReturned: messageList.length,
                  hasMore: !!response["@odata.nextLink"] || filteredMessages.length > limit,
                  messages: messageList,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Write tools - only registered in full-access mode
  if (!readOnly) {
    // Send chat message
    server.tool(
      "send_chat_message",
      "Send a message to a specific chat conversation. Supports text and markdown formatting, mentions, and importance levels.",
      {
        chatId: z.string().describe("Chat ID"),
        message: z.string().describe("Message content"),
        importance: z.enum(["normal", "high", "urgent"]).optional().describe("Message importance"),
        format: z.enum(["text", "markdown"]).optional().describe("Message format (text or markdown)"),
        mentions: z
          .array(
            z.object({
              mention: z
                .string()
                .describe("The @mention text (e.g., 'john.doe' or 'john.doe@company.com')"),
              userId: z.string().describe("Azure AD User ID of the mentioned user"),
            })
          )
          .optional()
          .describe("Array of @mentions to include in the message"),
      },
      async ({ chatId, message, importance = "normal", format = "text", mentions }) => {
        try {
          const client = await graphService.getClient();

          // Process message content based on format
          let content: string;
          let contentType: "text" | "html";

          if (format === "markdown") {
            content = await markdownToHtml(message);
            contentType = "html";
          } else {
            content = message;
            contentType = "text";
          }

          // Process @mentions if provided
          const mentionMappings: Array<{ mention: string; userId: string; displayName: string }> = [];
          if (mentions && mentions.length > 0) {
            // Convert provided mentions to mappings with display names
            for (const mention of mentions) {
              try {
                // Get user info to get display name
                const userResponse = await client
                  .api(`/users/${mention.userId}`)
                  .select("displayName")
                  .get();
                mentionMappings.push({
                  mention: mention.mention,
                  userId: mention.userId,
                  displayName: userResponse.displayName || mention.mention,
                });
              } catch (_error) {
                console.warn(
                  `Could not resolve user ${mention.userId}, using mention text as display name`
                );
                mentionMappings.push({
                  mention: mention.mention,
                  userId: mention.userId,
                  displayName: mention.mention,
                });
              }
            }
          }

          // Process mentions in HTML content
          let finalMentions: Array<{
            id: number;
            mentionText: string;
            mentioned: { user: { id: string } };
          }> = [];
          if (mentionMappings.length > 0) {
            const result = processMentionsInHtml(content, mentionMappings);
            content = result.content;
            finalMentions = result.mentions;

            // Ensure we're using HTML content type when mentions are present
            contentType = "html";
          }

          // Build message payload
          const messagePayload: any = {
            body: {
              content,
              contentType,
            },
            importance,
          };

          if (finalMentions.length > 0) {
            messagePayload.mentions = finalMentions;
          }

          const result = (await client
            .api(`/me/chats/${chatId}/messages`)
            .post(messagePayload)) as ChatMessage;

          // Build success message
          const successText = `✅ Message sent successfully. Message ID: ${result.id}${
            finalMentions.length > 0
              ? `\n📱 Mentions: ${finalMentions.map((m) => m.mentionText).join(", ")}`
              : ""
          }`;

          return {
            content: [
              {
                type: "text" as const,
                text: successText,
              },
            ],
          };
        } catch (error: any) {
          return {
            content: [
              {
                type: "text" as const,
                text: `❌ Failed to send message: ${error.message}`,
              },
            ],
            isError: true,
          };
        }
      }
    );

    // Create new chat (1:1 or group)
    server.tool(
      "create_chat",
      "Create a new chat conversation. Can be a 1:1 chat (with one other user) or a group chat (with multiple users). Group chats can optionally have a topic.",
      {
        userEmails: z.array(z.string()).describe("Array of user email addresses to add to chat"),
        topic: z.string().optional().describe("Chat topic (for group chats)"),
      },
      async ({ userEmails, topic }) => {
        try {
          const client = await graphService.getClient();

          // Get current user ID
          const me = (await client.api("/me").get()) as User;

          // Create members array
          const members: ConversationMember[] = [
            {
              "@odata.type": "#microsoft.graph.aadUserConversationMember",
              user: {
                id: me?.id,
              },
              roles: ["owner"],
            } as ConversationMember,
          ];

          // Add other users as members
          for (const email of userEmails) {
            const user = (await client.api(`/users/${email}`).get()) as User;
            members.push({
              "@odata.type": "#microsoft.graph.aadUserConversationMember",
              user: {
                id: user?.id,
              },
              roles: ["member"],
            } as ConversationMember);
          }

          const chatData: CreateChatPayload = {
            chatType: userEmails.length === 1 ? "oneOnOne" : "group",
            members,
          };

          if (topic && userEmails.length > 1) {
            chatData.topic = topic;
          }

          const newChat = (await client.api("/chats").post(chatData)) as Chat;

          return {
            content: [
              {
                type: "text",
                text: `✅ Chat created successfully. Chat ID: ${newChat?.id}`,
              },
            ],
          };
        } catch (error: unknown) {
          const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
          return {
            content: [
              {
                type: "text",
                text: `❌ Error: ${errorMessage}`,
              },
            ],
          };
        }
      }
    );
  } // end if (!readOnly)
}

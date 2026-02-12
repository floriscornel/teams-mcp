import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphService } from "../services/graph.js";
import type {
  Channel,
  ChannelSummary,
  ChatMessage,
  ConversationMember,
  GraphApiResponse,
  MemberSummary,
  MessageSummary,
  Team,
  TeamSummary,
} from "../types/graph.js";
import {
  type ImageAttachment,
  imageUrlToBase64,
  isValidImageType,
  uploadImageAsHostedContent,
} from "../utils/attachments.js";
import { markdownToHtml } from "../utils/markdown.js";
import { processMentionsInHtml, searchUsers, type UserInfo } from "../utils/users.js";

export function registerTeamsTools(server: McpServer, graphService: GraphService, readOnly = false) {
  // List user's teams
  server.tool(
    "list_teams",
    "List all Microsoft Teams that the current user is a member of. Returns team names, descriptions, and IDs.",
    {},
    async () => {
      try {
        const client = await graphService.getClient();
        const response = (await client.api("/me/joinedTeams").get()) as GraphApiResponse<Team>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No teams found.",
              },
            ],
          };
        }

        const teamList: TeamSummary[] = response.value.map((team: Team) => ({
          id: team.id,
          displayName: team.displayName,
          description: team.description,
          isArchived: team.isArchived,
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(teamList, null, 2),
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

  // List channels in a team
  server.tool(
    "list_channels",
    "List all channels in a specific Microsoft Team. Returns channel names, descriptions, types, and IDs for the specified team.",
    {
      teamId: z.string().describe("Team ID"),
    },
    async ({ teamId }) => {
      try {
        const client = await graphService.getClient();
        const response = (await client
          .api(`/teams/${teamId}/channels`)
          .get()) as GraphApiResponse<Channel>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No channels found in this team.",
              },
            ],
          };
        }

        const channelList: ChannelSummary[] = response.value.map((channel: Channel) => ({
          id: channel.id,
          displayName: channel.displayName,
          description: channel.description,
          membershipType: channel.membershipType,
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(channelList, null, 2),
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

  // Get channel messages
  server.tool(
    "get_channel_messages",
    "Retrieve recent messages from a specific channel in a Microsoft Team. Returns message content, sender information, and timestamps.",
    {
      teamId: z.string().describe("Team ID"),
      channelId: z.string().describe("Channel ID"),
      limit: z
        .number()
        .min(1)
        .max(50)
        .optional()
        .default(20)
        .describe("Number of messages to retrieve (default: 20)"),
    },
    async ({ teamId, channelId, limit }) => {
      try {
        const client = await graphService.getClient();

        // Build query parameters - Teams channel messages API has limited query support
        // Only $top is supported, no $orderby, $filter, etc.
        const queryParams: string[] = [`$top=${limit}`];
        const queryString = queryParams.join("&");

        const response = (await client
          .api(`/teams/${teamId}/channels/${channelId}/messages?${queryString}`)
          .get()) as GraphApiResponse<ChatMessage>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No messages found in this channel.",
              },
            ],
          };
        }

        const messageList: MessageSummary[] = response.value.map((message: ChatMessage) => ({
          id: message.id,
          content: message.body?.content,
          from: message.from?.user?.displayName,
          createdDateTime: message.createdDateTime,
          importance: message.importance,
        }));

        // Sort messages by creation date (newest first) since API doesn't support orderby
        messageList.sort((a, b) => {
          const dateA = new Date(a.createdDateTime || 0).getTime();
          const dateB = new Date(b.createdDateTime || 0).getTime();
          return dateB - dateA;
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  totalReturned: messageList.length,
                  hasMore: !!response["@odata.nextLink"],
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
    // Send message to channel
    server.tool(
      "send_channel_message",
      "Send a message to a specific channel in a Microsoft Team. Supports text and markdown formatting, mentions, and importance levels.",
      {
        teamId: z.string().describe("Team ID"),
        channelId: z.string().describe("Channel ID"),
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
        imageUrl: z.string().optional().describe("URL of an image to attach to the message"),
        imageData: z.string().optional().describe("Base64 encoded image data to attach"),
        imageContentType: z
          .string()
          .optional()
          .describe("MIME type of the image (e.g., 'image/jpeg', 'image/png')"),
        imageFileName: z.string().optional().describe("Name for the attached image file"),
      },
      async ({
        teamId,
        channelId,
        message,
        importance = "normal",
        format = "text",
        mentions,
        imageUrl,
        imageData,
        imageContentType,
        imageFileName,
      }) => {
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

          // Handle image attachment
          const attachments: ImageAttachment[] = [];
          if (imageUrl || imageData) {
            let imageInfo: { data: string; contentType: string } | null = null;

            if (imageUrl) {
              imageInfo = await imageUrlToBase64(imageUrl);
              if (!imageInfo) {
                return {
                  content: [
                    {
                      type: "text" as const,
                      text: `❌ Failed to download image from URL: ${imageUrl}`,
                    },
                  ],
                  isError: true,
                };
              }
            } else if (imageData && imageContentType) {
              if (!isValidImageType(imageContentType)) {
                return {
                  content: [
                    {
                      type: "text" as const,
                      text: `❌ Unsupported image type: ${imageContentType}`,
                    },
                  ],
                  isError: true,
                };
              }
              imageInfo = { data: imageData, contentType: imageContentType };
            }

            if (imageInfo) {
              const uploadResult = await uploadImageAsHostedContent(
                graphService,
                teamId,
                channelId,
                imageInfo.data,
                imageInfo.contentType,
                imageFileName
              );

              if (uploadResult) {
                attachments.push(uploadResult.attachment);
              } else {
                return {
                  content: [
                    {
                      type: "text" as const,
                      text: "❌ Failed to upload image attachment",
                    },
                  ],
                  isError: true,
                };
              }
            }
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

          if (attachments.length > 0) {
            messagePayload.attachments = attachments;
          }

          const result = (await client
            .api(`/teams/${teamId}/channels/${channelId}/messages`)
            .post(messagePayload)) as ChatMessage;

          // Build success message
          const successText = `✅ Message sent successfully. Message ID: ${result.id}${
            finalMentions.length > 0
              ? `\n📱 Mentions: ${finalMentions.map((m) => m.mentionText).join(", ")}`
              : ""
          }${attachments.length > 0 ? `\n🖼️ Image attached: ${attachments[0].name}` : ""}`;

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
  } // end if (!readOnly) for send_channel_message

  // Get replies to a message in a channel
  server.tool(
    "get_channel_message_replies",
    "Get all replies to a specific message in a channel. Returns reply content, sender information, and timestamps.",
    {
      teamId: z.string().describe("Team ID"),
      channelId: z.string().describe("Channel ID"),
      messageId: z.string().describe("Message ID to get replies for"),
      limit: z
        .number()
        .min(1)
        .max(50)
        .optional()
        .default(20)
        .describe("Number of replies to retrieve (default: 20)"),
    },
    async ({ teamId, channelId, messageId, limit }) => {
      try {
        const client = await graphService.getClient();

        // Only $top is supported for message replies
        const queryParams: string[] = [`$top=${limit}`];
        const queryString = queryParams.join("&");

        const response = (await client
          .api(
            `/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies?${queryString}`
          )
          .get()) as GraphApiResponse<ChatMessage>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No replies found for this message.",
              },
            ],
          };
        }

        const repliesList: MessageSummary[] = response.value.map((reply: ChatMessage) => ({
          id: reply.id,
          content: reply.body?.content,
          from: reply.from?.user?.displayName,
          createdDateTime: reply.createdDateTime,
          importance: reply.importance,
        }));

        // Sort replies by creation date (oldest first for replies)
        repliesList.sort((a, b) => {
          const dateA = new Date(a.createdDateTime || 0).getTime();
          const dateB = new Date(b.createdDateTime || 0).getTime();
          return dateA - dateB;
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  parentMessageId: messageId,
                  totalReplies: repliesList.length,
                  hasMore: !!response["@odata.nextLink"],
                  replies: repliesList,
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

  if (!readOnly) {
    // Reply to a message in a channel
    server.tool(
      "reply_to_channel_message",
      "Reply to a specific message in a channel. Supports text and markdown formatting, mentions, and importance levels.",
      {
        teamId: z.string().describe("Team ID"),
        channelId: z.string().describe("Channel ID"),
        messageId: z.string().describe("Message ID to reply to"),
        message: z.string().describe("Reply content"),
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
          .describe("Array of @mentions to include in the reply"),
        imageUrl: z.string().optional().describe("URL of an image to attach to the reply"),
        imageData: z.string().optional().describe("Base64 encoded image data to attach"),
        imageContentType: z
          .string()
          .optional()
          .describe("MIME type of the image (e.g., 'image/jpeg', 'image/png')"),
        imageFileName: z.string().optional().describe("Name for the attached image file"),
      },
      async ({
        teamId,
        channelId,
        messageId,
        message,
        importance = "normal",
        format = "text",
        mentions,
        imageUrl,
        imageData,
        imageContentType,
        imageFileName,
      }) => {
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

          // Handle image attachment
          const attachments: ImageAttachment[] = [];
          if (imageUrl || imageData) {
            let imageInfo: { data: string; contentType: string } | null = null;

            if (imageUrl) {
              imageInfo = await imageUrlToBase64(imageUrl);
              if (!imageInfo) {
                return {
                  content: [
                    {
                      type: "text" as const,
                      text: `❌ Failed to download image from URL: ${imageUrl}`,
                    },
                  ],
                  isError: true,
                };
              }
            } else if (imageData && imageContentType) {
              if (!isValidImageType(imageContentType)) {
                return {
                  content: [
                    {
                      type: "text" as const,
                      text: `❌ Unsupported image type: ${imageContentType}`,
                    },
                  ],
                  isError: true,
                };
              }
              imageInfo = { data: imageData, contentType: imageContentType };
            }

            if (imageInfo) {
              const uploadResult = await uploadImageAsHostedContent(
                graphService,
                teamId,
                channelId,
                imageInfo.data,
                imageInfo.contentType,
                imageFileName
              );

              if (uploadResult) {
                attachments.push(uploadResult.attachment);
              } else {
                return {
                  content: [
                    {
                      type: "text" as const,
                      text: "❌ Failed to upload image attachment",
                    },
                  ],
                  isError: true,
                };
              }
            }
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

          if (attachments.length > 0) {
            messagePayload.attachments = attachments;
          }

          const result = (await client
            .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`)
            .post(messagePayload)) as ChatMessage;

          // Build success message
          const successText = `✅ Reply sent successfully. Reply ID: ${result.id}${
            finalMentions.length > 0
              ? `\n📱 Mentions: ${finalMentions.map((m) => m.mentionText).join(", ")}`
              : ""
          }${attachments.length > 0 ? `\n🖼️ Image attached: ${attachments[0].name}` : ""}`;

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
                text: `❌ Failed to send reply: ${error.message}`,
              },
            ],
            isError: true,
          };
        }
      }
    );
  } // end if (!readOnly)

  // List team members
  server.tool(
    "list_team_members",
    "List all members of a specific Microsoft Team. Returns member names, email addresses, roles, and IDs.",
    {
      teamId: z.string().describe("Team ID"),
    },
    async ({ teamId }) => {
      try {
        const client = await graphService.getClient();
        const response = (await client
          .api(`/teams/${teamId}/members`)
          .get()) as GraphApiResponse<ConversationMember>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No members found in this team.",
              },
            ],
          };
        }

        const memberList: MemberSummary[] = response.value.map((member: ConversationMember) => ({
          id: member.id,
          displayName: member.displayName,
          roles: member.roles,
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(memberList, null, 2),
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

  // Search users for @mentions
  server.tool(
    "search_users_for_mentions",
    "Search for users to mention in messages. Returns users with their display names, email addresses, and mention IDs.",
    {
      query: z.string().describe("Search query (name or email)"),
      limit: z
        .number()
        .min(1)
        .max(50)
        .optional()
        .default(10)
        .describe("Maximum number of results to return"),
    },
    async ({ query, limit }) => {
      try {
        const users = await searchUsers(graphService, query, limit);

        if (users.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: `No users found matching "${query}".`,
              },
            ],
          };
        }

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  query,
                  totalResults: users.length,
                  users: users.map((user: UserInfo) => ({
                    id: user.id,
                    displayName: user.displayName,
                    userPrincipalName: user.userPrincipalName,
                    mentionText:
                      user.userPrincipalName?.split("@")[0] ||
                      user.displayName.toLowerCase().replace(/\s+/g, ""),
                  })),
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

  // Download hosted content (images) from a Teams message
  server.tool(
    "download_message_hosted_content",
    "Download hosted content (such as images) from a Teams channel message. Returns the content as base64 encoded data along with metadata. Use this to retrieve images or other inline content embedded in messages.",
    {
      teamId: z.string().describe("Team ID"),
      channelId: z.string().describe("Channel ID"),
      messageId: z.string().describe("Message ID containing the hosted content"),
      hostedContentId: z
        .string()
        .optional()
        .describe(
          "Specific hosted content ID to download. If not provided, downloads all hosted contents from the message."
        ),
      savePath: z
        .string()
        .optional()
        .describe(
          "Optional file path to save the content. Supports UNC paths (e.g., \\\\wsl.localhost\\Ubuntu\\tmp\\file.png)."
        ),
    },
    async ({ teamId, channelId, messageId, hostedContentId, savePath }) => {
      try {
        const client = await graphService.getClient();

        // First, get the message to find hosted content references
        const message = (await client
          .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}`)
          .get()) as ChatMessage;

        if (!message) {
          return {
            content: [
              {
                type: "text",
                text: "❌ Message not found.",
              },
            ],
            isError: true,
          };
        }

        // Extract hosted content IDs from the message body
        const bodyContent = message.body?.content || "";
        const hostedContentRegex = /hostedContents\/([a-zA-Z0-9_=-]+)\/\$value|itemid="([^"]+)"/gi;
        const matches: string[] = [];
        let match: RegExpExecArray | null;

        // biome-ignore lint/suspicious/noAssignInExpressions: needed for regex extraction
        while ((match = hostedContentRegex.exec(bodyContent)) !== null) {
          const contentId = match[1] || match[2];
          if (contentId && !matches.includes(contentId)) {
            matches.push(contentId);
          }
        }

        if (matches.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: "❌ No hosted content found in this message.",
              },
            ],
            isError: true,
          };
        }

        // If a specific hosted content ID is provided, filter to just that one
        const contentIds = hostedContentId ? [hostedContentId] : matches;

        // Download each hosted content
        const results: Array<{
          id: string;
          contentType: string;
          size: number;
          base64Data?: string;
          savedTo?: string;
          error?: string;
        }> = [];

        for (const contentId of contentIds) {
          try {
            // Get the hosted content binary data
            const response = await client
              .api(
                `/teams/${teamId}/channels/${channelId}/messages/${messageId}/hostedContents/${contentId}/$value`
              )
              .responseType("arraybuffer" as any)
              .get();

            // Convert ArrayBuffer to Buffer and then to base64
            const buffer = Buffer.from(response as ArrayBuffer);
            const base64Data = buffer.toString("base64");

            // Determine content type from the response or default to image/png
            const contentType = "image/png"; // Graph API doesn't return content-type header easily

            const result: {
              id: string;
              contentType: string;
              size: number;
              base64Data?: string;
              savedTo?: string;
            } = {
              id: contentId,
              contentType,
              size: buffer.length,
            };

            // If savePath is provided, save the file
            if (savePath) {
              const fs = await import("node:fs/promises");
              const path = await import("node:path");

              // Debug: log the savePath to stderr
              console.error(`[DEBUG] savePath received: "${savePath}"`);

              // Normalize path: JSON escaping can cause double backslashes (4 chars -> 2 chars)
              // \\\\wsl.localhost\\... -> \\wsl.localhost\...
              const normalizedPath = savePath.replace(/\\\\/g, "\\");
              console.error(`[DEBUG] normalizedPath after fix: "${normalizedPath}"`);

              // Check if path starts with \\ (UNC on Windows) or //
              const isUncPath =
                normalizedPath.startsWith("\\\\") || normalizedPath.startsWith("//");
              console.error(`[DEBUG] isUncPath: ${isUncPath}`);

              // If multiple files, append index to filename
              let finalPath = normalizedPath;
              if (contentIds.length > 1) {
                const ext = path.extname(normalizedPath);
                const base = ext ? normalizedPath.slice(0, -ext.length) : normalizedPath;
                const index = contentIds.indexOf(contentId);
                finalPath = `${base}_${index}${ext}`;
              }

              // For UNC paths, don't use path.resolve as it can mess up the path
              const targetPath = isUncPath ? finalPath : path.resolve(finalPath);
              console.error(`[DEBUG] targetPath: "${targetPath}"`);

              await fs.writeFile(targetPath, buffer);
              result.savedTo = targetPath;
            } else {
              // Return base64 data if not saving to file
              result.base64Data = base64Data;
            }

            results.push(result);
          } catch (downloadError) {
            const errorMsg =
              downloadError instanceof Error ? downloadError.message : "Unknown error";
            results.push({
              id: contentId,
              contentType: "unknown",
              size: 0,
              error: errorMsg,
            });
          }
        }

        // Build response
        const successCount = results.filter((r) => !r.error).length;
        const errorCount = results.filter((r) => r.error).length;

        let summary = `📥 Downloaded ${successCount} of ${contentIds.length} hosted content(s)`;
        if (errorCount > 0) {
          summary += ` (${errorCount} failed)`;
        }

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  summary,
                  messageId,
                  totalContentItems: contentIds.length,
                  successCount,
                  errorCount,
                  contents: results,
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
}

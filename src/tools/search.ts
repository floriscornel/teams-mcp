import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphService } from "../services/graph.js";
import type { SearchHit, SearchRequest, SearchResponse } from "../types/graph.js";

export function registerSearchTools(server: McpServer, graphService: GraphService) {
  // Search messages across Teams using Microsoft Search API
  server.tool(
    "search_messages",
    "Search for messages across all Microsoft Teams channels and chats using Microsoft Search API. Supports advanced KQL syntax for filtering by sender, mentions, attachments, and more.",
    {
      query: z
        .string()
        .describe(
          "Search query. Supports KQL syntax like 'from:user mentions:userId hasAttachment:true'"
        ),
      scope: z
        .enum(["all", "channels", "chats"])
        .optional()
        .default("all")
        .describe("Scope of search"),
      limit: z
        .number()
        .min(1)
        .max(100)
        .optional()
        .default(25)
        .describe("Number of results to return"),
      enableTopResults: z
        .boolean()
        .optional()
        .default(true)
        .describe("Enable relevance-based ranking"),
    },
    async ({ query, scope, limit, enableTopResults }) => {
      try {
        const client = await graphService.getClient();

        // Build the search request
        const searchRequest: SearchRequest = {
          entityTypes: ["chatMessage"],
          query: {
            queryString: query,
          },
          from: 0,
          size: limit,
          enableTopResults,
        };

        // Add scope-specific filters to the query if needed
        let enhancedQuery = query;
        if (scope === "channels") {
          enhancedQuery = `${query} AND (channelIdentity/channelId:*)`;
        } else if (scope === "chats") {
          enhancedQuery = `${query} AND (chatId:* AND NOT channelIdentity/channelId:*)`;
        }

        searchRequest.query.queryString = enhancedQuery;

        const response = (await client
          .api("/search/query")
          .post({ requests: [searchRequest] })) as SearchResponse;

        if (!response?.value?.length || !response.value[0]?.hitsContainers?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No messages found matching your search criteria.",
              },
            ],
          };
        }

        const hits = response.value[0].hitsContainers[0].hits;
        const searchResults = hits.map((hit: SearchHit) => ({
          id: hit.resource.id,
          summary: hit.summary,
          rank: hit.rank,
          content: hit.resource.body?.content || "No content",
          from: hit.resource.from?.user?.displayName || "Unknown",
          createdDateTime: hit.resource.createdDateTime,
          chatId: hit.resource.chatId,
          teamId: hit.resource.channelIdentity?.teamId,
          channelId: hit.resource.channelIdentity?.channelId,
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  query,
                  scope,
                  totalResults: response.value[0].hitsContainers[0].total,
                  results: searchResults,
                  moreResultsAvailable: response.value[0].hitsContainers[0].moreResultsAvailable,
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
              text: `❌ Error searching messages: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Get recent messages with advanced filtering
  server.tool(
    "get_recent_messages",
    "Get recent messages from across Teams with advanced filtering options. Can filter by time range, scope (channels vs chats), teams, channels, and users.",
    {
      hours: z
        .number()
        .min(1)
        .max(168)
        .optional()
        .default(24)
        .describe("Get messages from the last N hours (max 168 = 1 week)"),
      limit: z
        .number()
        .min(1)
        .max(100)
        .optional()
        .default(50)
        .describe("Maximum number of messages to return"),
      mentionsUser: z.string().optional().describe("Filter messages that mention this user ID"),
      fromUser: z.string().optional().describe("Filter messages from this user ID"),
      hasAttachments: z.boolean().optional().describe("Filter messages with attachments"),
      importance: z
        .enum(["low", "normal", "high", "urgent"])
        .optional()
        .describe("Filter by message importance"),
      includeChannels: z.boolean().optional().default(true).describe("Include channel messages"),
      includeChats: z.boolean().optional().default(true).describe("Include chat messages"),
      teamIds: z.array(z.string()).optional().describe("Specific team IDs to search in"),
      keywords: z.string().optional().describe("Keywords to search for in message content"),
    },
    async ({
      hours,
      limit,
      mentionsUser,
      fromUser,
      hasAttachments,
      importance,
      includeChannels,
      includeChats,
      teamIds,
      keywords,
    }) => {
      try {
        const client = await graphService.getClient();

        let attemptedAdvancedSearch = false;

        // Try using the Search API first for rich filtering
        if (keywords || mentionsUser || hasAttachments !== undefined || importance) {
          attemptedAdvancedSearch = true;
          // Calculate the date threshold
          const since = new Date(Date.now() - hours * 60 * 60 * 1000).toISOString();

          // Build KQL query for Microsoft Search API
          const queryParts: string[] = [];

          // Add time filter - use a more permissive date format
          queryParts.push(`sent>=${since.split("T")[0]}`); // Use just the date part

          // Add user filters
          if (mentionsUser) {
            queryParts.push(`mentions:${mentionsUser}`);
          }
          if (fromUser) {
            queryParts.push(`from:${fromUser}`);
          }

          // Add content filters
          if (hasAttachments !== undefined) {
            queryParts.push(`hasAttachment:${hasAttachments}`);
          }
          if (importance) {
            queryParts.push(`importance:${importance}`);
          }

          // Add keyword search
          if (keywords) {
            queryParts.push(`"${keywords}"`);
          }

          // If no specific filters, search for all recent messages
          if (queryParts.length === 1) {
            // Only has the time filter
            queryParts.push("*"); // Match all messages
          }

          const searchQuery = queryParts.join(" AND ");

          const searchRequest: SearchRequest = {
            entityTypes: ["chatMessage"],
            query: {
              queryString: searchQuery,
            },
            from: 0,
            size: Math.min(limit, 100),
            enableTopResults: false, // For recent messages, prefer chronological order
          };

          try {
            const response = (await client
              .api("/search/query")
              .post({ requests: [searchRequest] })) as SearchResponse;

            if (response?.value?.length && response.value[0]?.hitsContainers?.length) {
              const hits = response.value[0].hitsContainers[0].hits;
              const recentMessages = hits
                .filter((hit) => {
                  // Apply scope filters
                  const isChannelMessage = hit.resource.channelIdentity?.channelId;
                  const isChatMessage = hit.resource.chatId && !isChannelMessage;

                  if (!includeChannels && isChannelMessage) return false;
                  if (!includeChats && isChatMessage) return false;

                  // Apply team filter if specified
                  if (teamIds?.length && isChannelMessage) {
                    return teamIds.includes(hit.resource.channelIdentity?.teamId || "");
                  }

                  return true;
                })
                .map((hit: SearchHit) => ({
                  id: hit.resource.id,
                  content: hit.resource.body?.content || "No content",
                  from: hit.resource.from?.user?.displayName || "Unknown",
                  fromUserId: hit.resource.from?.user?.id,
                  createdDateTime: hit.resource.createdDateTime,
                  chatId: hit.resource.chatId,
                  teamId: hit.resource.channelIdentity?.teamId,
                  channelId: hit.resource.channelIdentity?.channelId,
                  type: hit.resource.channelIdentity?.channelId ? "channel" : "chat",
                }))
                .slice(0, limit); // Apply final limit after filtering

              // Check if Search API returned poor quality results (No content/Unknown)
              const poorQualityResults = recentMessages.filter(
                (msg) => msg.content === "No content" || msg.from === "Unknown"
              ).length;

              const qualityThreshold = 0.5; // If more than 50% of results are poor quality, fall back
              if (
                recentMessages.length > 0 &&
                poorQualityResults / recentMessages.length > qualityThreshold
              ) {
                console.log(
                  "Search API returned poor quality results, falling back to direct queries"
                );
                // Fall through to direct chat queries
              } else {
                return {
                  content: [
                    {
                      type: "text",
                      text: JSON.stringify(
                        {
                          method: "search_api",
                          timeRange: `Last ${hours} hours`,
                          filters: {
                            mentionsUser,
                            fromUser,
                            hasAttachments,
                            importance,
                            keywords,
                          },
                          totalFound: recentMessages.length,
                          messages: recentMessages,
                        },
                        null,
                        2
                      ),
                    },
                  ],
                };
              }
            }
          } catch (searchError) {
            console.error("Search API failed, falling back to direct queries:", searchError);
          }
        }

        // Fallback: Get recent messages from user's chats directly
        // This method is more reliable but doesn't support advanced filtering
        const chatsResponse = await client.api("/me/chats?$expand=members").get();
        const chats = chatsResponse?.value || [];

        const allMessages: Array<{
          id: string;
          content: string;
          from: string;
          fromUserId?: string;
          createdDateTime: string;
          chatId: string;
          type: string;
        }> = [];
        const since = new Date(Date.now() - hours * 60 * 60 * 1000);

        // Get recent messages from each chat
        for (const chat of chats.slice(0, 10)) {
          // Limit to first 10 chats to avoid rate limits
          try {
            let queryString = `$top=${Math.min(limit, 50)}&$orderby=createdDateTime desc`;

            // Apply user filter if specified
            if (fromUser) {
              queryString += `&$filter=from/user/id eq '${fromUser}'`;
            }

            const messagesResponse = await client
              .api(`/me/chats/${chat.id}/messages?${queryString}`)
              .get();

            const messages = messagesResponse?.value || [];

            for (const message of messages) {
              // Filter by time
              if (message.createdDateTime) {
                const messageDate = new Date(message.createdDateTime);
                if (messageDate < since) continue;
              }

              // Apply scope filter for chats
              if (!includeChats) {
                continue; // Skip chat messages if includeChats is false
              }

              // Apply keyword filter (simple text search)
              if (
                keywords &&
                message.body?.content &&
                !message.body.content.toLowerCase().includes(keywords.toLowerCase())
              ) {
                continue;
              }

              allMessages.push({
                id: message.id || "",
                content: message.body?.content || "No content",
                from: message.from?.user?.displayName || "Unknown",
                fromUserId: message.from?.user?.id,
                createdDateTime: message.createdDateTime || "",
                chatId: message.chatId || "",
                type: "chat",
              });

              if (allMessages.length >= limit) break;
            }

            if (allMessages.length >= limit) break;
          } catch (chatError) {
            console.error(`Error getting messages from chat ${chat.id}:`, chatError);
          }
        }

        // Sort by creation date (newest first)
        allMessages.sort(
          (a, b) => new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime()
        );

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  method: attemptedAdvancedSearch
                    ? "direct_chat_queries_fallback"
                    : "direct_chat_queries",
                  timeRange: `Last ${hours} hours`,
                  filters: {
                    mentionsUser,
                    fromUser,
                    hasAttachments,
                    importance,
                    keywords,
                  },
                  note: attemptedAdvancedSearch
                    ? "Search API returned poor quality results, using direct chat queries as fallback"
                    : "Using direct chat queries for better content reliability",
                  totalFound: allMessages.slice(0, limit).length,
                  messages: allMessages.slice(0, limit),
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
              text: `❌ Error getting recent messages: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Search for messages mentioning the current user
  server.tool(
    "get_my_mentions",
    "Find all recent messages where the current user was mentioned (@mentioned) across Teams channels and chats.",
    {
      hours: z
        .number()
        .min(1)
        .max(168)
        .optional()
        .default(24)
        .describe("Get mentions from the last N hours"),
      limit: z
        .number()
        .min(1)
        .max(50)
        .optional()
        .default(20)
        .describe("Maximum number of mentions to return"),
      scope: z
        .enum(["all", "channels", "chats"])
        .optional()
        .default("all")
        .describe("Scope of search"),
    },
    async ({ hours, limit, scope }) => {
      try {
        const client = await graphService.getClient();

        // Get current user ID first
        const me = await client.api("/me").get();
        const userId = me?.id;

        if (!userId) {
          return {
            content: [
              {
                type: "text",
                text: "❌ Error: Could not determine current user ID",
              },
            ],
          };
        }

        const since = new Date(Date.now() - hours * 60 * 60 * 1000).toISOString();

        // Build query to find mentions of current user
        const queryParts = [
          `sent>=${since.split("T")[0]}`, // Use just the date part to avoid time parsing issues
          `mentions:${userId}`,
        ];

        const searchQuery = queryParts.join(" AND ");

        const searchRequest: SearchRequest = {
          entityTypes: ["chatMessage"],
          query: {
            queryString: searchQuery,
          },
          from: 0,
          size: Math.min(limit, 50),
          enableTopResults: false,
        };

        const response = (await client
          .api("/search/query")
          .post({ requests: [searchRequest] })) as SearchResponse;

        if (
          !response?.value?.length ||
          !response.value[0]?.hitsContainers?.length ||
          !response.value[0].hitsContainers[0]?.hits
        ) {
          return {
            content: [
              {
                type: "text",
                text: "No recent mentions found.",
              },
            ],
          };
        }

        const hits = response.value[0].hitsContainers[0].hits || [];
        if (hits.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: "No recent mentions found.",
              },
            ],
          };
        }

        const mentions = hits
          .filter((hit) => {
            // Apply scope filters
            const isChannelMessage = hit.resource.channelIdentity?.channelId;
            const isChatMessage = hit.resource.chatId && !isChannelMessage;

            if (scope === "channels" && !isChannelMessage) return false;
            if (scope === "chats" && !isChatMessage) return false;

            return true;
          })
          .map((hit: SearchHit) => ({
            id: hit.resource.id,
            content: hit.resource.body?.content || "No content",
            summary: hit.summary,
            from: hit.resource.from?.user?.displayName || "Unknown",
            fromUserId: hit.resource.from?.user?.id,
            createdDateTime: hit.resource.createdDateTime,
            chatId: hit.resource.chatId,
            teamId: hit.resource.channelIdentity?.teamId,
            channelId: hit.resource.channelIdentity?.channelId,
            type: hit.resource.channelIdentity?.channelId ? "channel" : "chat",
          }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  timeRange: `Last ${hours} hours`,
                  mentionedUser: me?.displayName || "Current User",
                  scope,
                  totalMentions: mentions.length,
                  mentions,
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
              text: `❌ Error getting mentions: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );
}

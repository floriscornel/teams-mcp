import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphService } from "../services/graph.js";
import type { SearchHit, SearchRequest, SearchResponse } from "../types/graph.js";
import { formatMessageContent } from "../utils/html-to-markdown.js";

/**
 * Maps raw SearchHit objects from the Microsoft Search API into a
 * consistent, flat shape for tool responses.
 *
 * @param hits - Array of search hits from the Microsoft Search API
 * @param contentFormat - Format for message content: "markdown" or "raw"
 */
export function formatSearchHits(
  hits: SearchHit[],
  contentFormat: "raw" | "markdown" = "markdown"
) {
  return hits.map((hit) => ({
    id: hit.resource.id,
    summary: hit.summary,
    rank: hit.rank,
    content: formatMessageContent(hit.resource.body?.content, contentFormat),
    from: hit.resource.from?.user?.displayName,
    fromUserId: hit.resource.from?.user?.id,
    createdDateTime: hit.resource.createdDateTime,
    importance: hit.resource.importance,
    webLink: hit.resource.webLink,
    chatId: hit.resource.chatId,
    teamId: hit.resource.channelIdentity?.teamId,
    channelId: hit.resource.channelIdentity?.channelId,
  }));
}

export function registerSearchTools(server: McpServer, graphService: GraphService) {
  server.tool(
    "search_messages",
    [
      "Search for messages across all Microsoft Teams channels and chats using the Microsoft Search API.",
      "The query string supports KQL (Keyword Query Language) syntax for advanced filtering:",
      "  from:<name>              — messages sent by a person (e.g. from:bob)",
      "  to:<name>                — messages sent to a person",
      "  mentions:<userId>        — messages that mention a specific user ID (without dashes)",
      "  IsMentioned:true         — messages that mention the current user",
      "  hasAttachment:true|false — filter by attachment presence",
      "  IsRead:true|false        — filter by read status",
      "  sent>=YYYY-MM-DD         — messages sent on or after a date",
      "  sent<=YYYY-MM-DD         — messages sent on or before a date",
      "Examples:",
      '  "quarterly report" from:alice sent>=2025-01-01',
      "  hasAttachment:true from:bob",
      "  project update sent>=2025-02-01",
      "Use get_chat_messages or get_channel_messages for browsing a specific conversation.",
    ].join("\n"),
    {
      query: z.string().describe("Search query string. Supports KQL syntax (see tool description)"),
      from: z
        .number()
        .min(0)
        .optional()
        .default(0)
        .describe("Offset for pagination (0-based). Use with size to paginate through results"),
      size: z
        .number()
        .min(1)
        .max(100)
        .optional()
        .default(25)
        .describe("Number of results to return (max 100)"),
      enableTopResults: z
        .boolean()
        .optional()
        .default(true)
        .describe("When true, results are ranked by relevance. When false, results are unranked"),
      contentFormat: z
        .enum(["raw", "markdown"])
        .optional()
        .default("markdown")
        .describe(
          'Format for message content. "markdown" (default) converts Teams HTML to clean Markdown optimized for LLMs. "raw" returns original HTML from Graph API.'
        ),
    },
    async ({ query, from, size, enableTopResults, contentFormat }) => {
      try {
        const client = await graphService.getClient();

        const searchRequest: SearchRequest = {
          entityTypes: ["chatMessage"],
          query: { queryString: query },
          from,
          size,
          enableTopResults,
        };

        const response = (await client
          .api("/search/query")
          .post({ requests: [searchRequest] })) as SearchResponse;

        const container = response?.value?.[0]?.hitsContainers?.[0];
        if (!container?.hits?.length) {
          return {
            content: [{ type: "text", text: "No messages found matching your search criteria." }],
          };
        }

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  query,
                  from,
                  size,
                  total: container.total,
                  moreResultsAvailable: container.moreResultsAvailable,
                  results: formatSearchHits(container.hits, contentFormat ?? "markdown"),
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
          content: [{ type: "text", text: `❌ Error searching messages: ${errorMessage}` }],
        };
      }
    }
  );

  server.tool(
    "get_my_mentions",
    "Find recent messages where the current user was @mentioned across all Teams channels and chats.",
    {
      hours: z
        .number()
        .min(1)
        .max(168)
        .optional()
        .default(24)
        .describe("Look back this many hours (max 168 = 1 week)"),
      size: z
        .number()
        .min(1)
        .max(100)
        .optional()
        .default(25)
        .describe("Maximum number of mentions to return"),
      contentFormat: z
        .enum(["raw", "markdown"])
        .optional()
        .default("markdown")
        .describe(
          'Format for message content. "markdown" (default) converts Teams HTML to clean Markdown optimized for LLMs. "raw" returns original HTML from Graph API.'
        ),
    },
    async ({ hours, size, contentFormat }) => {
      try {
        const client = await graphService.getClient();

        // Resolve current user
        const me = await client.api("/me").get();
        const userId = me?.id;
        if (!userId) {
          return {
            content: [{ type: "text", text: "❌ Error: Could not determine current user ID" }],
          };
        }

        const sinceDate = new Date(Date.now() - hours * 60 * 60 * 1000).toISOString().split("T")[0];

        const searchRequest: SearchRequest = {
          entityTypes: ["chatMessage"],
          query: { queryString: `IsMentioned:true sent>=${sinceDate}` },
          from: 0,
          size,
          enableTopResults: false,
        };

        const response = (await client
          .api("/search/query")
          .post({ requests: [searchRequest] })) as SearchResponse;

        const container = response?.value?.[0]?.hitsContainers?.[0];
        if (!container?.hits?.length) {
          return { content: [{ type: "text", text: "No recent mentions found." }] };
        }

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  timeRange: `Last ${hours} hours`,
                  mentionedUser: me?.displayName || "Current User",
                  total: container.total,
                  moreResultsAvailable: container.moreResultsAvailable,
                  mentions: formatSearchHits(container.hits, contentFormat ?? "markdown"),
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
          content: [{ type: "text", text: `❌ Error getting mentions: ${errorMessage}` }],
        };
      }
    }
  );
}

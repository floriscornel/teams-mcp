import TurndownService from "turndown";
import { gfm } from "turndown-plugin-gfm";

// Create a singleton TurndownService instance configured for Teams HTML
const turndownService = new TurndownService({
  headingStyle: "atx",
  hr: "---",
  bulletListMarker: "-",
  codeBlockStyle: "fenced",
  emDelimiter: "*",
  strongDelimiter: "**",
});

// Enable GitHub Flavored Markdown (tables, strikethrough, task lists)
turndownService.use(gfm);

// Custom rule: Teams @mentions — <at id="N">Name</at> → @Name
turndownService.addRule("teamsMention", {
  filter: (node: HTMLElement) => node.nodeName === "AT",
  replacement: (content: string) => `@${content}`,
});

// Custom rule: Teams attachments — <attachment id="N"></attachment> → remove
turndownService.addRule("teamsAttachment", {
  filter: (node: HTMLElement) => node.nodeName === "ATTACHMENT",
  replacement: () => "",
});

// Custom rule: Teams system events — <systemEventMessage/> → remove
turndownService.addRule("teamsSystemEvent", {
  filter: (node: HTMLElement) =>
    node.nodeName === "SYSTEMEVENTMESSAGE" ||
    (node.textContent || "").trim() === "systemEventMessage/",
  replacement: () => "",
});

/**
 * Converts Teams HTML content to clean Markdown.
 * Handles Teams-specific elements like @mentions, attachments, and system events.
 *
 * @param html - Raw HTML content from Microsoft Graph API
 * @returns Clean Markdown string
 */
export function htmlToMarkdown(html: string): string {
  // Handle empty/null content
  if (!html || html.trim() === "") {
    return "";
  }

  return turndownService.turndown(html).trim();
}

/**
 * Formats message content based on the requested format.
 *
 * @param content - Raw message content from Graph API (HTML)
 * @param format - "markdown" to convert HTML to Markdown, "raw" to return as-is
 * @returns Formatted content string, or the original value if null/undefined
 */
export function formatMessageContent(
  content: string | null | undefined,
  format: "raw" | "markdown"
): string | null | undefined {
  if (content == null) {
    return content;
  }

  if (format === "raw") {
    return content;
  }

  return htmlToMarkdown(content);
}

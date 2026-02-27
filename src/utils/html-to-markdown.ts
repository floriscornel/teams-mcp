import type { ChatMessageMention } from "@microsoft/microsoft-graph-types";
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

// Custom rule: Teams attachments — pre-processed <span data-attachment="N"> → [attachment:N]
// Original <attachment> tags are converted to spans in htmlToMarkdown() because
// JSDOM does not recognize custom elements like <attachment>.
turndownService.addRule("teamsAttachment", {
  filter: (node: HTMLElement) => node.nodeName === "SPAN" && node.hasAttribute("data-attachment"),
  replacement: (content: string) => content,
});

// Custom rule: Teams system events — <systemEventMessage/> → remove
turndownService.addRule("teamsSystemEvent", {
  filter: (node: HTMLElement) =>
    node.nodeName === "SYSTEMEVENTMESSAGE" ||
    (node.textContent || "").trim() === "systemEventMessage/",
  replacement: () => "",
});

/**
 * Merges consecutive `<at>` tags that belong to the same user.
 *
 * Teams splits multi-word display names into separate `<at>` tags
 * (one per word), e.g. `<at id="0">John</at>&nbsp;<at id="1">Doe</at>`.
 * Without merging this produces `@John @Doe` instead of `@John Doe`.
 *
 * The `mentions` array from the Graph API maps each `<at id>` to a user ID
 * and provides the full display name, allowing us to reconstruct the
 * original mention.
 */
function mergeConsecutiveMentions(html: string, mentions: ChatMessageMention[]): string {
  if (mentions.length === 0) return html;

  // Build lookup: atId → userId (for grouping consecutive tags by person)
  const atIdToUserId = new Map<string, string>();

  for (const m of mentions) {
    if (m.id == null || !m.mentioned?.user?.id) continue;
    atIdToUserId.set(String(m.id), m.mentioned.user.id);
  }

  // Collect every <at> tag with its position and resolved userId
  interface AtTag {
    start: number;
    end: number;
    id: string;
    text: string;
    userId: string | undefined;
  }
  const atRegex = /<at\s+id="(\d+)">([^<]*)<\/at>/g;
  const tags: AtTag[] = [];
  for (const exec of html.matchAll(atRegex)) {
    tags.push({
      start: exec.index,
      end: exec.index + exec[0].length,
      id: exec[1],
      text: exec[2],
      userId: atIdToUserId.get(exec[1]),
    });
  }

  if (tags.length < 2) return html;

  // Group consecutive tags separated only by &nbsp; with the same userId
  const groups: AtTag[][] = [[tags[0]]];
  for (let i = 1; i < tags.length; i++) {
    const prev = tags[i - 1];
    const curr = tags[i];
    const between = html.substring(prev.end, curr.start);
    const lastGroup = groups[groups.length - 1];

    if (between === "&nbsp;" && prev.userId && curr.userId && prev.userId === curr.userId) {
      lastGroup.push(curr);
    } else {
      groups.push([curr]);
    }
  }

  // Rebuild HTML, replacing multi-tag groups with a single merged tag
  let result = "";
  let lastEnd = 0;

  for (const group of groups) {
    if (group.length <= 1) continue;

    const groupStart = group[0].start;
    const groupEnd = group[group.length - 1].end;
    // Always join the tag texts — the API returns mentionText per-word,
    // so the concatenated <at> contents are the most reliable full name.
    const fullName = group.map((t) => t.text).join(" ");

    result += html.substring(lastEnd, groupStart);
    result += `<at id="${group[0].id}">${fullName}</at>`;
    lastEnd = groupEnd;
  }

  result += html.substring(lastEnd);
  return result;
}

/**
 * Converts Teams HTML content to clean Markdown.
 * Handles Teams-specific elements like @mentions, attachments, and system events.
 *
 * @param html - Raw HTML content from Microsoft Graph API
 * @param mentions - Optional mentions array from ChatMessage for multi-word name merging
 * @returns Clean Markdown string
 */
export function htmlToMarkdown(html: string, mentions?: ChatMessageMention[] | null): string {
  // Handle empty/null content
  if (!html || html.trim() === "") {
    return "";
  }

  // Pre-process: merge consecutive <at> tags for the same user
  let preprocessed = mentions ? mergeConsecutiveMentions(html, mentions) : html;

  // Pre-process: ensure adjacent <at> tags have a space between them.
  // Teams sometimes omits any separator between mentions of different people,
  // e.g. `</at><at id="3">` which Turndown would render as `@Name1@Name2`.
  preprocessed = preprocessed.replace(/<\/at>(<at\s)/g, "</at> $1");

  // Pre-process: replace <attachment id="X"></attachment> with inline markers
  // JSDOM does not recognize custom <attachment> elements, so we convert them
  // to plain text markers before Turndown processes the HTML.
  preprocessed = preprocessed.replace(
    /<attachment(?:\s+id="([^"]*)")?><\/attachment>/gi,
    (_match, id) =>
      id
        ? `<span data-attachment="${id}">{attachment:${id}}</span>`
        : "<span data-attachment>{attachment}</span>"
  );

  return turndownService
    .turndown(preprocessed)
    .replace(/\u00A0/g, " ") // Normalize non-breaking spaces for LLM consumption
    .trim();
}

/**
 * Formats message content based on the requested format.
 *
 * @param content - Raw message content from Graph API (HTML)
 * @param format - "markdown" to convert HTML to Markdown, "raw" to return as-is
 * @param mentions - Optional mentions array from ChatMessage for multi-word name merging
 * @returns Formatted content string, or the original value if null/undefined
 */
export function formatMessageContent(
  content: string | null | undefined,
  format: "raw" | "markdown",
  mentions?: ChatMessageMention[] | null
): string | null | undefined {
  if (content == null) {
    return content;
  }

  if (format === "raw") {
    return content;
  }

  return htmlToMarkdown(content, mentions);
}

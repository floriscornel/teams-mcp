import { describe, expect, it } from "vitest";
import { formatMessageContent, htmlToMarkdown } from "../html-to-markdown.js";

describe("htmlToMarkdown", () => {
  it("should return empty string for empty input", () => {
    expect(htmlToMarkdown("")).toBe("");
  });

  it("should return empty string for whitespace-only input", () => {
    expect(htmlToMarkdown("   ")).toBe("");
  });

  it("should convert bold text", () => {
    expect(htmlToMarkdown("<strong>Bold</strong>")).toBe("**Bold**");
  });

  it("should convert italic text", () => {
    expect(htmlToMarkdown("<em>italic</em>")).toBe("*italic*");
  });

  it("should convert links", () => {
    expect(htmlToMarkdown('<a href="https://example.com">Link</a>')).toBe(
      "[Link](https://example.com)"
    );
  });

  it("should convert unordered lists", () => {
    const html = "<ul><li>Item 1</li><li>Item 2</li></ul>";
    const result = htmlToMarkdown(html);
    expect(result).toMatch(/-\s+Item 1/);
    expect(result).toMatch(/-\s+Item 2/);
  });

  it("should convert ordered lists", () => {
    const html = "<ol><li>First</li><li>Second</li></ol>";
    const result = htmlToMarkdown(html);
    expect(result).toMatch(/1\.\s+First/);
    expect(result).toMatch(/2\.\s+Second/);
  });

  it("should convert headings", () => {
    expect(htmlToMarkdown("<h1>Title</h1>")).toBe("# Title");
    expect(htmlToMarkdown("<h2>Subtitle</h2>")).toBe("## Subtitle");
    expect(htmlToMarkdown("<h3>Section</h3>")).toBe("### Section");
  });

  it("should convert inline code", () => {
    expect(htmlToMarkdown("<code>console.log()</code>")).toBe("`console.log()`");
  });

  it("should convert code blocks", () => {
    const html = "<pre><code>const x = 1;\nconst y = 2;</code></pre>";
    const result = htmlToMarkdown(html);
    expect(result).toContain("```");
    expect(result).toContain("const x = 1;");
    expect(result).toContain("const y = 2;");
  });

  it("should convert horizontal rules", () => {
    expect(htmlToMarkdown("<hr>")).toBe("---");
  });

  it("should convert blockquotes", () => {
    expect(htmlToMarkdown("<blockquote>Quoted text</blockquote>")).toBe("> Quoted text");
  });

  it("should convert strikethrough (GFM)", () => {
    const result = htmlToMarkdown("<del>deleted</del>");
    expect(result).toMatch(/~+deleted~+/);
  });

  it("should convert tables (GFM)", () => {
    const html = `
      <table>
        <thead><tr><th>Name</th><th>Value</th></tr></thead>
        <tbody><tr><td>A</td><td>1</td></tr></tbody>
      </table>
    `;
    const result = htmlToMarkdown(html);
    expect(result).toContain("| Name | Value |");
    expect(result).toContain("| --- | --- |");
    expect(result).toContain("| A | 1 |");
  });

  it("should decode HTML entities", () => {
    const result = htmlToMarkdown("<p>A &amp; B &lt; C &gt; D</p>");
    expect(result).toContain("A & B < C > D");
  });

  it("should handle plain text wrapped in paragraphs", () => {
    expect(htmlToMarkdown("<p>Hello world</p>")).toBe("Hello world");
  });

  it("should handle mixed formatting", () => {
    const html = "<p><strong>Bold</strong> and <em>italic</em> and <code>code</code></p>";
    const result = htmlToMarkdown(html);
    expect(result).toBe("**Bold** and *italic* and `code`");
  });

  // Teams-specific elements

  it("should convert Teams @mentions to @Name", () => {
    expect(htmlToMarkdown('<at id="0">John Doe</at>')).toBe("@John Doe");
  });

  it("should handle multiple mentions", () => {
    const html = '<p><at id="0">Alice</at> and <at id="1">Bob</at> are here</p>';
    const result = htmlToMarkdown(html);
    expect(result).toBe("@Alice and @Bob are here");
  });

  it("should merge consecutive at-tags for the same user using mentions array", () => {
    const html = '<p><at id="0">Brunno</at>&nbsp;<at id="1">Joyran</at>&nbsp;hello</p>';
    // Graph API returns mentionText per-word (one entry per <at> tag)
    const mentions = [
      { id: 0, mentionText: "Brunno", mentioned: { user: { id: "user-A" } } },
      { id: 1, mentionText: "Joyran", mentioned: { user: { id: "user-A" } } },
    ];
    const result = htmlToMarkdown(html, mentions);
    expect(result).toBe("@Brunno Joyran hello");
  });

  it("should merge two different people with multi-word names", () => {
    const html =
      '<p><at id="0">Brunno</at>&nbsp;<at id="1">Joyran</at>&nbsp;<at id="2">Pereira</at>&nbsp;<at id="3">Soares</at></p>';
    const mentions = [
      { id: 0, mentionText: "Brunno", mentioned: { user: { id: "user-A" } } },
      { id: 1, mentionText: "Joyran", mentioned: { user: { id: "user-A" } } },
      { id: 2, mentionText: "Pereira", mentioned: { user: { id: "user-B" } } },
      { id: 3, mentionText: "Soares", mentioned: { user: { id: "user-B" } } },
    ];
    const result = htmlToMarkdown(html, mentions);
    expect(result).toBe("@Brunno Joyran @Pereira Soares");
  });

  it("should not merge at-tags for different users", () => {
    const html = '<p><at id="0">Alice</at>&nbsp;<at id="1">Bob</at></p>';
    const mentions = [
      { id: 0, mentionText: "Alice", mentioned: { user: { id: "user-A" } } },
      { id: 1, mentionText: "Bob", mentioned: { user: { id: "user-B" } } },
    ];
    const result = htmlToMarkdown(html, mentions);
    expect(result).toBe("@Alice @Bob");
  });

  it("should work without mentions array (backwards compatible)", () => {
    const html = '<p><at id="0">Brunno</at>&nbsp;<at id="1">Joyran</at>&nbsp;hello</p>';
    const result = htmlToMarkdown(html);
    expect(result).toBe("@Brunno @Joyran hello");
  });

  it("should separate adjacent at-tags with no separator between them", () => {
    // Teams sometimes omits &nbsp; between different people's mentions
    const html =
      '<p><at id="0">Lindeberg</at>&nbsp;<at id="1">Pessoa</at>&nbsp;<at id="2">Leite</at><at id="3">Bruno</at>&nbsp;<at id="4">Werneck</at>&nbsp;<at id="5">Pinto</at>&nbsp;<at id="6">Hoelz</at></p>';
    const mentions = [
      { id: 0, mentionText: "Lindeberg", mentioned: { user: { id: "user-A" } } },
      { id: 1, mentionText: "Pessoa", mentioned: { user: { id: "user-A" } } },
      { id: 2, mentionText: "Leite", mentioned: { user: { id: "user-A" } } },
      { id: 3, mentionText: "Bruno", mentioned: { user: { id: "user-B" } } },
      { id: 4, mentionText: "Werneck", mentioned: { user: { id: "user-B" } } },
      { id: 5, mentionText: "Pinto", mentioned: { user: { id: "user-B" } } },
      { id: 6, mentionText: "Hoelz", mentioned: { user: { id: "user-B" } } },
    ];
    const result = htmlToMarkdown(html, mentions);
    expect(result).toBe("@Lindeberg Pessoa Leite @Bruno Werneck Pinto Hoelz");
  });

  it("should convert Teams attachment elements to inline markers", () => {
    const html = '<p>See the file</p><attachment id="abc123"></attachment>';
    const result = htmlToMarkdown(html);
    expect(result).toContain("See the file");
    expect(result).toContain("{attachment:abc123}");
  });

  it("should handle attachment elements without id", () => {
    const html = "<p>See the file</p><attachment></attachment>";
    const result = htmlToMarkdown(html);
    expect(result).toContain("See the file");
    expect(result).toContain("{attachment}");
  });

  it("should remove systemEventMessage elements", () => {
    const html = "<systemEventMessage/>";
    const result = htmlToMarkdown(html);
    expect(result).toBe("");
  });

  it("should not discard messages that mention systemEventMessage in text", () => {
    const html = "<p>The &lt;systemEventMessage&gt; tag is used for system events</p>";
    const result = htmlToMarkdown(html);
    expect(result).toContain("systemEventMessage");
    expect(result.length).toBeGreaterThan(0);
  });

  it("should handle a realistic Teams message with mentions and formatting", () => {
    const html = `<p><at id="0">Alice</at>&nbsp;<at id="1">Bob</at>&nbsp;on the audit service</p>
<ul><li><strong>Latency</strong>: each call triggers validation</li></ul>
<table><thead><tr><th>Change</th><th>Verdict</th></tr></thead>
<tbody><tr><td>Remove auth</td><td>Approved</td></tr></tbody></table>`;
    const result = htmlToMarkdown(html);
    expect(result).toContain("@Alice");
    expect(result).toContain("@Bob");
    expect(result).toContain("**Latency**");
    expect(result).toMatch(/-\s+\*\*Latency\*\*: each call triggers validation/);
    expect(result).toContain("| Change | Verdict |");
    expect(result).toContain("| Remove auth | Approved |");
    // Should NOT contain raw HTML tags
    expect(result).not.toContain("<at");
    expect(result).not.toContain("<strong>");
    expect(result).not.toContain("<table>");
    expect(result).not.toContain("&nbsp;");
  });

  it("should handle images", () => {
    const html = '<img src="https://example.com/img.png" alt="screenshot">';
    const result = htmlToMarkdown(html);
    expect(result).toBe("![screenshot](https://example.com/img.png)");
  });
});

describe("formatMessageContent", () => {
  it("should return null for null input", () => {
    expect(formatMessageContent(null, "markdown")).toBeNull();
  });

  it("should return undefined for undefined input", () => {
    expect(formatMessageContent(undefined, "markdown")).toBeUndefined();
  });

  it("should return null for null input with raw format", () => {
    expect(formatMessageContent(null, "raw")).toBeNull();
  });

  it("should return raw HTML when format is raw", () => {
    const html = "<p><strong>Bold</strong></p>";
    expect(formatMessageContent(html, "raw")).toBe(html);
  });

  it("should convert to markdown when format is markdown", () => {
    const html = "<p><strong>Bold</strong></p>";
    expect(formatMessageContent(html, "markdown")).toBe("**Bold**");
  });

  it("should pass through empty string in raw mode", () => {
    expect(formatMessageContent("", "raw")).toBe("");
  });

  it("should return empty string for empty string in markdown mode", () => {
    expect(formatMessageContent("", "markdown")).toBe("");
  });
});

import { describe, expect, it } from "vitest";
import type { ChatMessage } from "../../types/graph.js";
import {
  AUDIO_CARD_CONTENT_TYPE,
  extractHostedContentRefs,
  resolveHostedContentType,
} from "../hosted-content.js";

const audioCard = (...urls: unknown[]) =>
  JSON.stringify({ duration: "PT1M32S", media: urls.map((url) => ({ url })) });

describe("extractHostedContentRefs", () => {
  it("returns an empty list when the message has no hosted content", () => {
    const message = { body: { content: "Plain text" } } as ChatMessage;
    expect(extractHostedContentRefs(message)).toEqual([]);
  });

  it("extracts inline hosted content from the body", () => {
    const message = {
      body: {
        content:
          '<img src="https://graph.microsoft.com/v1.0/chats/c/messages/m/hostedContents/img1/$value"><img itemid="img2">',
      },
    } as ChatMessage;
    expect(extractHostedContentRefs(message)).toEqual([
      { id: "img1", source: "body" },
      { id: "img2", source: "body" },
    ]);
  });

  it("extracts voice message content from audio card attachments", () => {
    const message = {
      body: { content: '<div><attachment id="att1"></attachment></div>' },
      attachments: [
        {
          id: "att1",
          contentType: AUDIO_CARD_CONTENT_TYPE,
          content: audioCard(
            "https://graph.microsoft.com/v1.0/chats/c/messages/m/hostedContents/aWQ9LHR5cGU9Mw==/$value"
          ),
        },
      ],
    } as ChatMessage;
    expect(extractHostedContentRefs(message)).toEqual([
      { id: "aWQ9LHR5cGU9Mw==", source: "audioCard" },
    ]);
  });

  it("decodes percent-encoded audio card URLs", () => {
    const message = {
      attachments: [
        {
          contentType: AUDIO_CARD_CONTENT_TYPE,
          content: audioCard(
            "https://graph.microsoft.com/v1.0/chats/c/messages/m/hostedContents/aWQ9Mw%3D%3D/$value"
          ),
        },
      ],
    } as ChatMessage;
    expect(extractHostedContentRefs(message)).toEqual([{ id: "aWQ9Mw==", source: "audioCard" }]);
  });

  it("preserves '+' and encoded slashes in audio card IDs", () => {
    const message = {
      attachments: [
        {
          contentType: AUDIO_CARD_CONTENT_TYPE,
          content: audioCard(
            "https://graph.microsoft.com/v1.0/chats/c/messages/m/hostedContents/a+b%2Fc%3D/$value",
            "https://graph.microsoft.com/v1.0/chats/c/messages/m/hostedContents/bad%E0%A4%A/$value"
          ),
        },
      ],
    } as ChatMessage;
    expect(extractHostedContentRefs(message)).toEqual([
      { id: "a+b/c=", source: "audioCard" },
      { id: "bad%E0%A4%A", source: "audioCard" },
    ]);
  });

  it("deduplicates IDs and keeps the first source", () => {
    const url = "https://graph.microsoft.com/v1.0/chats/c/messages/m/hostedContents/dup/$value";
    const message = {
      body: { content: `<img src="${url}">` },
      attachments: [{ contentType: AUDIO_CARD_CONTENT_TYPE, content: audioCard(url) }],
    } as ChatMessage;
    expect(extractHostedContentRefs(message)).toEqual([{ id: "dup", source: "body" }]);
  });

  it("ignores malformed or unrelated attachments", () => {
    const message = {
      body: { content: "" },
      attachments: [
        { contentType: AUDIO_CARD_CONTENT_TYPE, content: "not json" },
        { contentType: AUDIO_CARD_CONTENT_TYPE, content: null },
        { contentType: AUDIO_CARD_CONTENT_TYPE, content: JSON.stringify({ media: "x" }) },
        {
          contentType: AUDIO_CARD_CONTENT_TYPE,
          content: audioCard(42, "https://example.com/a.mp4"),
        },
        {
          contentType: "reference",
          content: audioCard(
            "https://graph.microsoft.com/v1.0/chats/c/messages/m/hostedContents/other/$value"
          ),
        },
      ],
    } as ChatMessage;
    expect(extractHostedContentRefs(message)).toEqual([]);
  });
});

describe("resolveHostedContentType", () => {
  it("reports MP4 from an audio card as audio", () => {
    expect(resolveHostedContentType("video/mp4", "audioCard")).toBe("audio/mp4");
    expect(resolveHostedContentType("application/octet-stream", "audioCard")).toBe("audio/mp4");
  });

  it("keeps the detected type otherwise", () => {
    expect(resolveHostedContentType("video/mp4", "body")).toBe("video/mp4");
    expect(resolveHostedContentType("image/png", "audioCard")).toBe("image/png");
    expect(resolveHostedContentType("image/png", undefined)).toBe("image/png");
  });
});

# Bug Report: Teams MCP `subject` Field Missing from Channel Message Tools

## Summary
The `teams-mcp` connector's `get_channel_messages` and `get_channel_message_replies` tools silently drop the `subject` field from returned message payloads. This field is populated and accessible via `Microsoft 365:read_resource` for the same message, proving the data exists upstream in the Microsoft Graph API response but is being filtered out by the MCP tool's response transformation.

## Impact
Subjects on Teams channel messages are a primary user-facing identifier (they render as bold headers above the message body in the Teams UI). Dropping them causes downstream consumers — including LLM agents summarizing or classifying messages — to lose critical context that is often *not* duplicated in the body. In a test case, a message used the subject line to declare a specific date range that the body did not repeat; without the subject, the agent misclassified the message's content. This is a general class of failure: any workflow that depends on subject-line metadata will silently produce wrong results.

## Reproduction

**Environment:** `teams-mcp` connector, Microsoft 365 tenant with a standard Team + channel.

**Steps:**
1. Post a message to a Teams channel with a non-empty subject line (Teams' "New conversation" → "Add a subject" field).
2. Call `teams-mcp:get_channel_messages` with the team ID and channel ID.
3. Inspect the returned message object.
4. Call `Microsoft 365:read_resource` with URI `teams:///teams/{teamId}/channels/{channelId}/messages/{messageId}` for the same message.
5. Compare the two responses.

**Expected:** Both responses include `subject`.

**Actual:** `get_channel_messages` response omits `subject` entirely. `read_resource` returns the populated `subject` value.

## Concrete example

A single channel message, retrieved via both tools.

`get_channel_messages` returned (abbreviated):
```json
{
  "id": "<messageId>",
  "content": "<ul><li>...body content...",
  "from": "<Display Name>",
  "createdDateTime": "2026-04-17T23:21:06.587Z",
  "importance": "normal",
  "reactions": []
}
```

`read_resource` on the same message returned (abbreviated):
```json
{
  "id": "<messageId>",
  "subject": "<subject line text>",
  "summary": null,
  "body": { "contentType": "html", "content": "<ul>..." },
  "lastModifiedDateTime": "...",
  "lastEditedDateTime": "...",
  "webUrl": "...",
  "mentions": [],
  "attachments": [],
  "replyToId": null,
  ...
}
```

Fields present in `read_resource` but absent from `get_channel_messages`: `subject`, `summary`, `body.contentType`, `lastModifiedDateTime`, `lastEditedDateTime`, `webUrl`, `mentions`, `attachments`, `replyToId`.

## Also affected: `get_channel_message_replies`

Same issue on replies, though testing suggested Teams may not allow subjects on reply messages (they inherit the parent thread). Worth verifying whether the field is simply always null for replies vs. being dropped in transformation.

## Suggested fix
Pass through `subject` (and ideally `summary`, `webUrl`, `lastEditedDateTime`, `mentions`, `attachments`, `replyToId`) in the response schema for `get_channel_messages` and `get_channel_message_replies`. These are all present in the Microsoft Graph `chatMessage` resource and useful for agent reasoning. The minimal fix is just `subject`; the broader fix is to stop selectively filtering fields and let the caller decide what to ignore.

## Workaround
For every root message that might be relevant, make a follow-up `Microsoft 365:read_resource` call to fetch the full message payload including the subject. This is wasteful — one extra API call per message — and requires the agent to know the bug exists.

## Scope note
Worth checking whether related tools (`get_chat_messages`, `search_messages`) have the same field-filtering issue on their response paths. If so, the root cause is probably a shared response mapper and one fix could cover all of them.

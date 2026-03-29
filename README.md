# Teams MCP

[![npm version](https://img.shields.io/npm/v/@floriscornel/teams-mcp.svg)](https://www.npmjs.com/package/@floriscornel/teams-mcp)
[![npm downloads](https://img.shields.io/npm/dm/@floriscornel/teams-mcp.svg)](https://www.npmjs.com/package/@floriscornel/teams-mcp)
[![codecov](https://codecov.io/gh/floriscornel/teams-mcp/graph/badge.svg)](https://app.codecov.io/gh/floriscornel/teams-mcp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![GitHub stars](https://img.shields.io/github/stars/floriscornel/teams-mcp.svg)](https://github.com/floriscornel/teams-mcp/stargazers)

A Model Context Protocol (MCP) server that provides seamless integration with Microsoft Graph APIs, enabling AI assistants to interact with Microsoft Teams, users, chats, files, and organizational data.

<a href="https://glama.ai/mcp/servers/@floriscornel/teams-mcp">
  <img width="380" height="200" src="https://glama.ai/mcp/servers/@floriscornel/teams-mcp/badge" alt="Teams MCP server" />
</a>

## 📦 Installation

To use this MCP server in Cursor/Claude/VS Code, add the following configuration:

```json
{
  "mcpServers": {
    "teams-mcp": {
      "command": "npx",
      "args": ["-y", "@floriscornel/teams-mcp@latest"]
    }
  }
}
```

## 🚀 Features

### 🔐 Authentication
- OAuth 2.0 device code authentication flow with Microsoft Graph
- Secure token management, cache persistence, and refresh token renewal
- Authentication status checking and logout support
- Read-only mode with reduced scopes
- Direct `AUTH_TOKEN` support for pre-issued Microsoft Graph access tokens

### 👥 User Management
- Get current user information
- Search users by name or email
- Retrieve detailed user profiles
- Access organizational directory data

### 🏢 Microsoft Teams Integration
- **Teams Management**
  - List user's joined teams
  - Access team details and metadata

- **Channel Operations**
  - List channels within teams
  - Retrieve channel messages and replies
  - Send messages to team channels
  - Reply to existing channel threads
  - Edit and soft delete channel messages and replies
  - Support for message importance levels (`normal`, `high`, `urgent`)
  - Support for inline image attachments via URL or base64 data

- **Team Members**
  - List team members and their roles
  - Access member information
  - Search users for `@mentions`

### 💬 Chat & Messaging
- **1:1 and Group Chats**
  - List user's chats
  - Create new 1:1 or group conversations
  - Retrieve chat message history with filtering, ordering, and pagination
  - Fetch all available messages via `@odata.nextLink` pagination
  - Send messages to existing chats
  - Edit previously sent chat messages
  - Soft delete chat messages

### ✏️ Message Management
- **Edit & Delete**
  - Update (edit) sent messages in chats and channels
  - Soft delete messages in chats and channels (marks as deleted without permanent removal)
  - Only message senders can update/delete their own messages
  - Support for Markdown formatting, mentions, and importance levels on edits

### 📎 Media & Attachments
- **Hosted Content**
  - Download hosted content (images, files) from chat and channel messages
  - Access inline images and attachments shared in conversations
  - Optionally save hosted content directly to disk

- **File Upload**
  - Upload and send any file type (PDF, DOCX, XLSX, ZIP, images, etc.) to channels and chats
  - Large file support (>4 MB) via resumable upload sessions
  - Channel uploads go to SharePoint and chat uploads go to OneDrive
  - Optional message text, custom filename, formatting, and importance levels

### 🔍 Advanced Search & Discovery
- **Message Search**
  - Search across all Teams channels and chats using Microsoft Search API
  - Support for KQL (Keyword Query Language) syntax
  - Filter by sender, mentions, attachments, read state, and date ranges
  - Get recent messages with advanced filtering options
  - Find messages mentioning the current user

## Rich Message Formatting Support

The following tools support rich message formatting in Teams channels and chats:
- `send_channel_message`
- `send_chat_message`
- `reply_to_channel_message`
- `update_channel_message`
- `update_chat_message`
- `send_file_to_channel`
- `send_file_to_chat`

### Format Options

You can specify the `format` parameter to control the message formatting:
- `text` (default): Plain text
- `markdown`: Markdown formatting (bold, italic, lists, links, code, etc.) converted to sanitized HTML

When `format` is set to `markdown`, the message content is converted to HTML using a secure markdown parser and sanitized to remove potentially dangerous content before being sent to Teams.

If `format` is not specified, the message will be sent as plain text.

### Example Usage

```json
{
  "teamId": "...",
  "channelId": "...",
  "message": "**Bold text** and _italic text_\n\n- List item 1\n- List item 2\n\n[Link](https://example.com)",
  "format": "markdown",
  "importance": "high"
}
```

```json
{
  "chatId": "...",
  "message": "Simple plain text message",
  "format": "text"
}
```

### Security Features

- **HTML Sanitization**: All markdown content is converted to HTML and sanitized to remove potentially dangerous elements (scripts, event handlers, etc.)
- **Allowed Tags**: Only safe HTML tags are permitted (p, strong, em, a, ul, ol, li, h1-h6, code, pre, etc.)
- **Safe Attributes**: Only safe attributes are allowed
- **XSS Prevention**: Content is automatically sanitized to prevent cross-site scripting attacks

### Supported Markdown Features

- **Text formatting**: Bold (`**text**`), italic (`_text_`), strikethrough (`~~text~~`)
- **Links**: `[text](url)`
- **Lists**: Bulleted (`- item`) and numbered (`1. item`)
- **Code**: Inline `` `code` `` and fenced code blocks
- **Headings**: `# H1` through `###### H6`
- **Blockquotes**: `> quoted text`
- **Tables**: GitHub-flavored markdown tables

## LLM-Friendly Content Format

Messages retrieved from the Microsoft Graph API are returned as raw HTML containing Teams-specific tags. To make this content more consumable by AI assistants, the following tools support automatic HTML-to-Markdown conversion:

- `get_chat_messages`
- `get_channel_messages`
- `get_channel_message_replies`
- `search_messages`
- `get_my_mentions`

### Content Format Options

Use the `contentFormat` parameter to control how message content is returned:
- `markdown` (default): Converts Teams HTML to clean Markdown, optimized for LLM consumption
- `raw`: Returns the original HTML from the Microsoft Graph API

### What Gets Converted

| HTML Element                           | Markdown Output                                           |
| -------------------------------------- | --------------------------------------------------------- |
| `<at id="0">Name</at>` (Teams mention) | `@Name` (multi-word names merged using mentions metadata) |
| `<strong>text</strong>`                | `**text**`                                                |
| `<em>text</em>`                        | `*text*`                                                  |
| `<code>text</code>`                    | `` `text` ``                                              |
| `<a href="url">text</a>`               | `[text](url)`                                             |
| `<ul><li>item</li></ul>`               | `- item`                                                  |
| `<table>...</table>`                   | GFM Markdown table                                        |
| `<attachment id="...">`                | `{attachment:id}`                                         |
| `<systemEventMessage/>`                | *(removed)*                                               |
| `<hr>`                                 | `---`                                                     |
| `&nbsp;`, `&amp;`, etc.                | Decoded to plain characters                               |

### Attachment Metadata

Messages that contain file attachments or inline images include an `attachments` array in the response with metadata for each attachment (id, name, contentType, contentUrl, thumbnailUrl). The inline `{attachment:id}` markers in the markdown content correlate with entries in this array, allowing consumers to identify and download attachments via `download_message_hosted_content` or `download_chat_hosted_content`.

### Example Usage

```json
{
  "chatId": "19:meeting_...",
  "limit": 10,
  "contentFormat": "markdown"
}
```

To get the original HTML:

```json
{
  "chatId": "19:meeting_...",
  "limit": 10,
  "contentFormat": "raw"
}
```

## 📦 Installation

```bash
# Install dependencies
npm install

# Build the project
npm run build

# Set up authentication
npm run auth
```

## 🔧 Configuration

### Prerequisites
- Node.js 18+
- Microsoft 365 account with appropriate permissions
- Microsoft Graph delegated permissions for the scopes below

### Required Microsoft Graph Permissions

**Full mode (default):**
- `User.Read` - Read user profile
- `User.ReadBasic.All` - Read basic user info
- `Team.ReadBasic.All` - Read team information
- `Channel.ReadBasic.All` - Read channel information
- `ChannelMessage.Read.All` - Read channel messages
- `ChannelMessage.Send` - Send channel messages and replies
- `ChannelMessage.ReadWrite` - Edit and delete channel messages
- `Chat.Read` - Read chat messages (included via read-only scopes)
- `Chat.ReadWrite` - Create and manage chats, send/edit/delete chat messages (supersedes `Chat.Read`)
- `TeamMember.Read.All` - Read team members
- `Files.ReadWrite.All` - Required for file uploads to channels and chats

**Read-only mode** (`TEAMS_MCP_READ_ONLY=true`) — only these scopes are requested:
- `User.Read`
- `User.ReadBasic.All`
- `Team.ReadBasic.All`
- `Channel.ReadBasic.All`
- `ChannelMessage.Read.All`
- `TeamMember.Read.All`
- `Chat.Read`

### Authentication Modes

**Full access:**

```bash
npx @floriscornel/teams-mcp@latest authenticate
```

**Read-only access:**

```bash
npx @floriscornel/teams-mcp@latest authenticate --read-only
```

**Direct token injection with an existing Microsoft Graph JWT:**

```json
{
  "mcpServers": {
    "teams-mcp": {
      "command": "npx",
      "args": ["-y", "@floriscornel/teams-mcp@latest"],
      "env": {
        "AUTH_TOKEN": "<jwt-for-https://graph.microsoft.com>"
      }
    }
  }
}
```

### Token Storage

- Auth metadata is stored locally at `~/.msgraph-mcp-auth.json`
- Token cache is stored locally at `~/.teams-mcp-token-cache.json`

## 🛠️ Usage

### Starting the Server
```bash
# Development mode with hot reload
npm run dev

# Production mode
npm run build && node dist/index.js

# Start in read-only mode (disables all write tools)
TEAMS_MCP_READ_ONLY=true node dist/index.js
```

### CLI Commands

```bash
npx @floriscornel/teams-mcp@latest authenticate              # Authenticate with full scopes
npx @floriscornel/teams-mcp@latest authenticate --read-only  # Authenticate with read-only scopes
npx @floriscornel/teams-mcp@latest check                     # Check authentication status
npx @floriscornel/teams-mcp@latest logout                    # Clear authentication
npx @floriscornel/teams-mcp@latest auth                      # Alias for authenticate
npx @floriscornel/teams-mcp@latest                           # Start MCP server (default)
```

### Environment Variables

- `TEAMS_MCP_READ_ONLY=true` - Start the MCP server in read-only mode
- `AUTH_TOKEN=<jwt>` - Use a pre-existing Microsoft Graph access token instead of MSAL login

### Read-Only Mode

The server supports a read-only mode that disables all write operations (sending messages, creating chats, uploading files, editing/deleting messages) and requests only read-permission scopes from Microsoft Graph.

**Enable read-only mode** using either:
- Environment variable: `TEAMS_MCP_READ_ONLY=true`
- CLI flag: `--read-only`

**Authenticate with reduced scopes:**
```bash
npx @floriscornel/teams-mcp@latest authenticate --read-only
```

**MCP server configuration (read-only):**
```json
{
  "mcpServers": {
    "teams-mcp": {
      "command": "npx",
      "args": ["-y", "@floriscornel/teams-mcp@latest"],
      "env": {
        "TEAMS_MCP_READ_ONLY": "true"
      }
    }
  }
}
```

**Switching modes:** When switching from read-only to full mode, the server detects the scope mismatch and warns you to re-authenticate:
```bash
npx @floriscornel/teams-mcp@latest authenticate
```

**Read-only tools (16):**
`auth_status`, `get_current_user`, `search_users`, `get_user`, `list_teams`, `list_channels`, `get_channel_messages`, `get_channel_message_replies`, `list_team_members`, `search_users_for_mentions`, `download_message_hosted_content`, `list_chats`, `get_chat_messages`, `download_chat_hosted_content`, `search_messages`, `get_my_mentions`

**Write tools disabled in read-only mode (10):**
`send_channel_message`, `reply_to_channel_message`, `update_channel_message`, `delete_channel_message`, `send_file_to_channel`, `send_chat_message`, `create_chat`, `update_chat_message`, `delete_chat_message`, `send_file_to_chat`

### Available MCP Tools

#### Authentication
- `auth_status` - Check current authentication status

#### User Operations
- `get_current_user` - Get authenticated user information
- `search_users` - Search for users by name or email
- `get_user` - Get detailed user information by ID or email

#### Teams Operations
- `list_teams` - List user's joined teams
- `list_channels` - List channels in a specific team
- `get_channel_messages` - Retrieve messages from a team channel with attachment summaries and content format selection
- `get_channel_message_replies` - Get replies to a specific channel message
- `send_channel_message` - Send a message to a team channel with optional mentions, importance, and image attachments
- `reply_to_channel_message` - Reply to an existing channel message
- `update_channel_message` - Edit a previously sent channel message or reply
- `delete_channel_message` - Soft delete a channel message or reply
- `list_team_members` - List members of a specific team
- `search_users_for_mentions` - Search for team members to @mention in messages
- `send_file_to_channel` - Upload a local file and send it as a message to a channel

#### Chat Operations
- `list_chats` - List user's chats (1:1 and group)
- `get_chat_messages` - Retrieve messages from a specific chat with pagination, filters, ordering, and `fetchAll`
- `send_chat_message` - Send a message to a chat
- `create_chat` - Create a new 1:1 or group chat
- `update_chat_message` - Edit a previously sent chat message
- `delete_chat_message` - Soft delete a chat message
- `send_file_to_chat` - Upload a local file and send it as a message to a chat

#### Media Operations
- `download_message_hosted_content` - Download hosted content (images, files) from channel messages
- `download_chat_hosted_content` - Download hosted content (images, files) from chat messages

#### Search Operations
- `search_messages` - Search across all Teams messages using KQL syntax
- `get_my_mentions` - Find recent messages mentioning the current user

## 📋 Examples

### Authentication

First, authenticate with Microsoft Graph:

```bash
# Full access (default)
npx @floriscornel/teams-mcp@latest authenticate

# Read-only (reduced permission scopes)
npx @floriscornel/teams-mcp@latest authenticate --read-only
```

Check your authentication status:

```bash
npx @floriscornel/teams-mcp@latest check
```

Logout if needed:

```bash
npx @floriscornel/teams-mcp@latest logout
```

### Chat Pagination Example

```json
{
  "chatId": "19:meeting_...",
  "limit": 100,
  "fetchAll": true,
  "orderBy": "createdDateTime",
  "descending": true,
  "contentFormat": "markdown"
}
```

### Channel Message with Mentions and Image

```json
{
  "teamId": "team-id",
  "channelId": "channel-id",
  "message": "Please review **today's update**",
  "format": "markdown",
  "importance": "high",
  "mentions": [
    {
      "mention": "alex.chen",
      "userId": "00000000-0000-0000-0000-000000000000"
    }
  ],
  "imageUrl": "https://example.com/status.png"
}
```

### File Upload Example

```json
{
  "chatId": "19:meeting_...",
  "filePath": "/absolute/path/to/report.pdf",
  "message": "Please review the attached report",
  "format": "markdown"
}
```

### Integrating with Cursor/Claude

This MCP server is designed to work with AI assistants like Claude/Cursor/VS Code through the Model Context Protocol.

```json
{
  "mcpServers": {
    "teams-mcp": {
      "command": "npx",
      "args": ["-y", "@floriscornel/teams-mcp@latest"]
    }
  }
}
```

## 🔒 Security

- All authentication is handled through Microsoft's OAuth 2.0 flow or a caller-provided Microsoft Graph token
- **Refresh token support**: Access tokens are automatically renewed using cached refresh tokens, so you don't need to re-authenticate every hour
- Token cache is stored locally at `~/.teams-mcp-token-cache.json`
- Auth metadata is stored locally at `~/.msgraph-mcp-auth.json`
- Markdown content is sanitized before sending HTML to Teams
- `AUTH_TOKEN` is validated to ensure it targets `https://graph.microsoft.com`
- No sensitive data is logged or exposed
- Follows Microsoft Graph API security best practices

## 📝 License

MIT License - see LICENSE file for details

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Run build, linting, and tests
5. Submit a pull request

## 📞 Support

For issues and questions:
- Check the existing GitHub issues
- Review Microsoft Graph API documentation
- Ensure proper authentication and permissions are configured

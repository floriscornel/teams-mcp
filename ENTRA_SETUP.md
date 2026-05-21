# Entra ID setup for teams-mcp (interactive auth-code flow)

This fork adds an `authenticate --interactive` flow that uses **OAuth auth-code + PKCE
with a loopback redirect**, sidestepping the upstream device-code flow that gets blocked
by tenants which enable the Microsoft-managed Conditional Access policy
*"Block device code flow"* (often enforced by tenant Conditional Access policy).

The interactive flow requires an Entra app registration **in your tenant**, because the
upstream default (Microsoft Graph CLI app, `14d82eec-204b-4c2f-b7e8-296a70dab67e`) does
not have `http://localhost` registered as a public-client redirect URI in your tenant.

## Which path applies to you?

The app registration is **per tenant, not per user** — one registration serves the whole
organisation. Tokens are user-delegated and stored per user (`~/.teams-mcp-token-cache.json`),
so every signed-in user only sees their own Teams data.

- **Path A — first person in the tenant / admin:** no existing registration yet. Follow
  sections 1 → 5 below to create it, grant admin consent, and authenticate.
- **Path B — your org already has a teams-mcp app registration:** ask whoever set it up
  (or your Entra admin) for the **client ID** and **tenant ID**. Skip sections 1 – 3 and
  jump straight to [section 4 (Authenticate)](#4-authenticate).

The client ID is **not a secret** — it's safe to share internally (Teams/Slack/wiki),
but don't publish it to public GitHub.

## 1. Register the app

Microsoft Entra admin center → **Identity → Applications → App registrations → New registration**

| Field                 | Value                                                          |
|-----------------------|----------------------------------------------------------------|
| Name                  | `teams-mcp` (or org-specific, e.g. `yourcompany-teams-mcp`)    |
| Supported accounts    | **Single tenant only** (`Accounts in this organizational directory only`) |
| Redirect URI platform | **Public client/native (mobile & desktop)**                    |
| Redirect URI value    | `http://localhost` *(no port — Entra allows any loopback port at runtime per RFC 8252)* |

> If the new registration UI defaulted to **Web** platform, fix it after creation:
> *Authentication* blade → delete the Web redirect URI → **Add a platform** → *Mobile and
> desktop applications* → `http://localhost`.

After creation, on the **Authentication** blade → **Settings** tab → confirm
**"Allow public client flows"** is **Enabled**. MSAL will refuse the flow otherwise.

## 2. Configure API permissions

**API permissions** blade → **Add a permission** → **Microsoft Graph** → **Delegated permissions**.

Required for full functionality:

| Permission                              | Why                                                |
|-----------------------------------------|----------------------------------------------------|
| `User.Read`                             | Read signed-in user profile (auth status check)    |
| `User.ReadBasic.All`                    | Resolve users referenced in chats / messages       |
| `Team.ReadBasic.All`                    | List teams, resolve team ids                       |
| `Channel.ReadBasic.All`                 | List channels                                      |
| `ChannelMessage.Read.All`               | Read channel messages (admin-consent required)     |
| `ChannelMessage.Send`                   | Post replies                                       |
| `ChannelMessage.ReadWrite`              | Edit / react / unreact on channel messages         |
| `Chat.ReadWrite`                        | Read + send 1:1 / group chat messages              |
| `TeamMember.Read.All`                   | Resolve team member identities                     |
| `Files.ReadWrite.All`                   | Upload attachments to channels / chats             |
| `TeamsAppInstallation.ReadWriteSelfForChat` *(optional)* | Allow the app to install itself in a chat (some attachment flows) |

A few of these (`ChannelMessage.Read.All`, `TeamsAppInstallation.*`) are flagged
**"Admin consent required"**. After adding, click **Grant admin consent for
\<your tenant\>** at the top of the permissions list. All rows should show
**Status: Granted for \<your tenant\>** (green check).

For a read-only deployment (start MCP with `--read-only`), the minimum set is
`User.Read`, `User.ReadBasic.All`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`,
`ChannelMessage.Read.All`, `TeamMember.Read.All`, `Chat.Read`.

## 3. Note the IDs

**Overview** blade → copy:

- **Application (client) ID** → set as `TEAMS_MCP_CLIENT_ID`
- **Directory (tenant) ID** → use in `TEAMS_MCP_AUTHORITY` =
  `https://login.microsoftonline.com/<tenant-id>`

Single-tenant authority is **required** when the app registration is single-tenant;
the upstream default `/common` will reject sign-in with `AADSTS50194`.

## 4. Authenticate

You need the **client ID** and **tenant ID** from section 3 (or from your admin if you
are on Path B). From the cloned/built fork:

```bash
TEAMS_MCP_CLIENT_ID=<your-client-id> \
TEAMS_MCP_AUTHORITY=https://login.microsoftonline.com/<your-tenant-id> \
  node dist/index.js authenticate --interactive
```

A browser tab opens against `login.microsoftonline.com`, you sign in (passing
Conditional Access naturally because the flow is interactive + carries device
identity), and the redirect lands on `http://localhost:<ephemeral-port>` where
MSAL's loopback listener exchanges the code for a token.

The refresh token is cached at `~/.teams-mcp-token-cache.json`. The MCP server
re-uses this cache silently on every start, refreshing access tokens as they
expire — no further sign-in required until the refresh token's lifetime ends
(typically 90 days).

## 5. Wire into Claude Code / MCP client

Edit `~/.claude.json` (or your MCP client's config):

```json
{
  "mcpServers": {
    "teams-mcp": {
      "type": "stdio",
      "command": "node",
      "args": ["/absolute/path/to/teams-mcp/dist/index.js"],
      "env": {
        "TEAMS_MCP_CLIENT_ID": "<your-client-id>",
        "TEAMS_MCP_AUTHORITY": "https://login.microsoftonline.com/<your-tenant-id>"
      }
    }
  }
}
```

`AUTH_TOKEN=<jwt>` is still supported and takes priority over the cache.

## Troubleshooting

- **`AADSTS65001 / consent required`** — admin hasn't consented to the
  permissions, or you added a new permission after the last consent. Re-run
  *Grant admin consent* in the Entra portal.
- **`AADSTS50194 / not configured as multi-tenant`** — you set
  `TEAMS_MCP_AUTHORITY` to `/common` but the app is single-tenant. Use the
  tenant-specific authority `https://login.microsoftonline.com/<tenant-id>`.
- **`AADSTS530032 / blocked by Conditional Access`** — the interactive flow
  is still being blocked. Check the CA report-only / blocking policies on the
  user; consider adding the user to an exclusion group or enrolling the device
  in Intune.
- **Browser shows mojibake on the success page** — clear the token cache
  (`teams-mcp logout`) and rebuild; the success template was fixed to declare
  UTF-8 charset.

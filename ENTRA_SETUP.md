# Entra ID setup for interactive sign-in

`authenticate --interactive` signs in with the **OAuth 2.0 authorization code flow with
PKCE and a loopback redirect**, instead of the default device code flow. Use it when
your organization blocks device code sign-in through Conditional Access (for example
with a *"Block device code flow"* policy). Device code sign-in then fails with
`AADSTS530032`.

Because the sign-in happens in your browser, Conditional Access evaluates it like any
other browser sign-in (MFA, device compliance, sign-in risk).

The interactive flow needs an app registration in your own tenant: the default
Microsoft Graph CLI app does not let you register `http://localhost` as a redirect URI.

> **Already have an app registration?** An app registration serves the whole tenant,
> while tokens are stored per user. If someone in your organization already created
> one for teams-mcp, ask them for the client ID and tenant ID and skip to
> [step 4](#4-sign-in).

## 1. Register the app

In the Microsoft Entra admin center, go to **App registrations → New registration**:

| Field                   | Value                                                            |
|-------------------------|------------------------------------------------------------------|
| Name                    | Any name, e.g. `teams-mcp`                                       |
| Supported account types | *Accounts in this organizational directory only* (single tenant) |
| Redirect URI platform   | **Public client/native (mobile & desktop)**                      |
| Redirect URI            | `http://localhost`                                               |

Leave the port off the redirect URI. Entra accepts any port on a loopback redirect, and
MSAL picks a free one at sign-in time.

If the redirect URI ended up under the **Web** platform, remove it and add it again on
the **Authentication** page with **Add a platform → Mobile and desktop applications**.

On the same **Authentication** page, set **Allow public client flows** to **Yes**.

## 2. Add API permissions

Under **API permissions → Add a permission → Microsoft Graph → Delegated permissions**,
add the permissions teams-mcp requests:

| Permission                 | Used for                                           |
|----------------------------|----------------------------------------------------|
| `User.Read`                | Signed-in user profile                             |
| `User.ReadBasic.All`       | Looking up users                                   |
| `Team.ReadBasic.All`       | Listing teams                                      |
| `Channel.ReadBasic.All`    | Listing channels                                   |
| `ChannelMessage.Read.All`  | Reading channel messages                           |
| `TeamMember.Read.All`      | Listing team members                               |
| `Chat.Read`                | Reading chats                                      |
| `ChannelMessage.Send`      | Sending channel messages and replies               |
| `ChannelMessage.ReadWrite` | Editing, deleting and reacting to channel messages |
| `Chat.ReadWrite`           | Creating chats and sending chat messages           |
| `Files.ReadWrite.All`      | Uploading files to channels and chats              |

In read-only mode (`--read-only` or `TEAMS_MCP_READ_ONLY=true`) only the first seven are
requested. If your app has a different set, request exactly that set with
`TEAMS_MCP_SCOPES`.

Several of these permissions require admin consent. An admin can grant it with
**Grant admin consent for \<tenant\>** on the same page. Every permission should then
show as granted.

## 3. Copy the IDs

On the app's **Overview** page, copy:

- **Application (client) ID**: use it as `TEAMS_MCP_CLIENT_ID`
- **Directory (tenant) ID**: use it as `TEAMS_MCP_TENANT_ID`

For a single-tenant app the tenant ID is required. Without it the `common` authority is
used, and sign-in fails with `AADSTS50194`.

## 4. Sign in

```bash
TEAMS_MCP_CLIENT_ID=<client-id> \
TEAMS_MCP_TENANT_ID=<tenant-id> \
npx @floriscornel/teams-mcp@latest authenticate --interactive
```

Your browser opens the Microsoft sign-in page. After you sign in, the browser is
redirected to a temporary listener on `http://localhost`, and the CLI stores the tokens
in `~/.teams-mcp-token-cache.json`. The MCP server reuses and refreshes these tokens on
every start, so you only sign in again when the refresh token expires or is revoked.

## 5. Configure your MCP client

Pass the same environment variables to the MCP server, so it refreshes tokens with the
same app. For example:

```json
{
  "mcpServers": {
    "teams-mcp": {
      "command": "npx",
      "args": ["-y", "@floriscornel/teams-mcp@latest"],
      "env": {
        "TEAMS_MCP_CLIENT_ID": "<client-id>",
        "TEAMS_MCP_TENANT_ID": "<tenant-id>"
      }
    }
  }
}
```

## Troubleshooting

- **`AADSTS65001` (consent required):** admin consent has not been granted, or a
  permission was added after consent was given. Grant admin consent again.
- **`AADSTS50194` (not configured as multi-tenant):** `TEAMS_MCP_TENANT_ID` is not set
  while the app is single-tenant. Set it to the directory (tenant) ID.
- **`AADSTS50011` (redirect URI mismatch):** `http://localhost` is missing or registered
  under the wrong platform. See [step 1](#1-register-the-app).
- **`AADSTS7000218` (client assertion or secret required):** **Allow public client
  flows** is off. See [step 1](#1-register-the-app).
- **`AADSTS530032` with `--interactive`:** Conditional Access also blocks the browser
  sign-in, for example because the device is not compliant. Check the sign-in logs in
  Entra or ask your administrator which policy applies.

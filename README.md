# mcp-outlook

A production-grade [Model Context Protocol](https://modelcontextprotocol.io) server that connects AI agents to Microsoft Outlook — email, calendar, attachments, SharePoint, and inbox rules — via the Microsoft Graph API.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Tests](https://img.shields.io/badge/tests-262%20passing-brightgreen)]()
[![Coverage](https://img.shields.io/badge/coverage-61%25-yellow)]()
[![Node](https://img.shields.io/badge/node-%3E%3D18-blue)]()

## What it does

Give any MCP-compatible AI agent (Claude, GPT, Gemini, etc.) the ability to read, send, and manage Outlook email; check and create calendar events; download and parse attachments; access SharePoint files; and manage inbox rules — all through natural language.

**46 tools** across 6 categories:

| Category | Count | Highlights |
|----------|-------|------------|
| **Email** | 15 | List, search, send, reply, forward, draft, move, flag, categorize, batch operations |
| **Calendar** | 17 | Events, recurring meetings, availability, online meetings, timezone handling |
| **Attachments** | 4 | List, download with auto-parsing (PDF/Word/Excel/PowerPoint), upload, scan |
| **Folders** | 4 | List, create, rename, stats |
| **SharePoint** | 3 | Access files via sharing links or direct IDs, resolve links |
| **Rules** | 3 | List, create, delete server-side inbox rules |

## Quick start

### 1. Register an Azure app (one-time, 5 minutes)

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) → **New registration**
2. Name it anything (e.g., `MCP Outlook`), pick your account type:
   - **Work/school**: "Accounts in this organizational directory only"
   - **Personal** (outlook.com): "Accounts in any organizational directory and personal Microsoft accounts"
3. Redirect URI: select **Web**, enter `http://localhost/callback`
4. After registration, go to **Authentication** → enable **Allow public client flows** → Save
5. Go to **API permissions** → Add **Microsoft Graph** → **Delegated permissions**:
   ```
   Mail.Read  Mail.ReadWrite  Mail.Send
   Calendars.Read  Calendars.ReadWrite
   User.Read  MailboxSettings.Read
   Files.Read.All  Sites.Read.All
   offline_access
   ```
6. Copy your **Application (client) ID** and **Directory (tenant) ID** from the Overview page

> No client secret needed — this uses OAuth 2.0 with PKCE (see [How authentication works](#how-authentication-works)).

### 2. Install and configure

```bash
git clone https://github.com/w21180239/mcp-outlook.git
cd mcp-outlook
npm install
```

Then add the server to your AI tool's MCP configuration:

<details>
<summary><b>Claude Code</b></summary>

```bash
claude mcp add outlook -- node /absolute/path/to/mcp-outlook/server/index.js \
  --env AZURE_CLIENT_ID=your-client-id \
  --env AZURE_TENANT_ID=your-tenant-id
```

Or add to `~/.claude.json`:
```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/absolute/path/to/mcp-outlook/server/index.js"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```
</details>

<details>
<summary><b>Claude Desktop</b></summary>

Add to your Claude Desktop config (`~/Library/Application Support/Claude/claude_desktop_config.json` on macOS):
```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/absolute/path/to/mcp-outlook/server/index.js"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```
</details>

<details>
<summary><b>Cursor</b></summary>

Add to `.cursor/mcp.json` in your project or `~/.cursor/mcp.json` globally:
```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/absolute/path/to/mcp-outlook/server/index.js"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```
</details>

<details>
<summary><b>Windsurf / Other MCP clients</b></summary>

The server speaks standard MCP over stdio. Point your client at:
```
command: node
args: ["/absolute/path/to/mcp-outlook/server/index.js"]
env: AZURE_CLIENT_ID=..., AZURE_TENANT_ID=...
```
</details>

### 3. Authenticate

The first time any tool is called, a browser window opens for Microsoft login. After that, tokens are cached and refreshed automatically — no re-login needed between sessions.

---

## How authentication works

This server uses **OAuth 2.0 Authorization Code flow with PKCE** (Proof Key for Code Exchange), the recommended pattern for public clients (desktop/CLI apps) that cannot securely store a client secret.

```
Agent calls a tool
       │
       ▼
┌─ ensureAuthenticated() ─────────────────────────────┐
│  Token cached & valid?  ──yes──▶  Use cached token   │
│        │ no                                          │
│  Refresh token works?   ──yes──▶  Silent refresh     │
│        │ no                                          │
│  Start PKCE flow:                                    │
│    1. Generate random code_verifier + code_challenge  │
│    2. Open browser → Microsoft login page             │
│    3. User authenticates, Azure redirects to          │
│       localhost:{random_port}/callback with auth code  │
│    4. Exchange code + code_verifier for tokens         │
│    5. Store tokens encrypted on disk                  │
└──────────────────────────────────────────────────────┘
```

**Key properties:**
- **No client secret** — the PKCE challenge/verifier pair proves the caller is the same entity that started the flow
- **Tokens encrypted at rest** — AES-256 encryption using OS keychain (via `keytar`) or a random persistent key
- **Scoped to `/me/`** — all Graph API calls are scoped to the authenticated user's own mailbox, calendar, and files
- **Automatic refresh** — expired tokens are silently refreshed using the stored refresh token; browser re-login only happens when the refresh token itself expires

---

## Example prompts

Once connected, just talk to your agent naturally:

**Email**
- "Show me unread emails from this week"
- "Find all emails from Alice about the budget"
- "Reply to that email thanking her for the update"
- "Draft an email to the team with meeting notes"

**Calendar**
- "What meetings do I have tomorrow?"
- "Schedule a 30-min call with Bob next Tuesday at 2pm"
- "Am I free on Friday afternoon?"

**Attachments**
- "Download and summarize the PDF from the latest Finance email"
- "What's in the Excel file attached to that report?"

**SharePoint**
- "Get the contents of this SharePoint link: [paste]"

**Rules**
- "Show my inbox rules"
- "Create a rule to move emails from noreply@example.com to the Archive folder"

---

## Configuration

| Environment variable | Required | Description |
|---------------------|----------|-------------|
| `AZURE_CLIENT_ID` | Yes | Azure AD application client ID |
| `AZURE_TENANT_ID` | Yes | Azure AD directory (tenant) ID |
| `MCP_OUTLOOK_WORK_DIR` | No | Directory for saving large files (defaults to system temp) |
| `DEBUG` | No | Set to any value to enable debug logging on stderr |

---

## For agent builders

### Claude Code skill

If you use Claude Code and want Outlook tools always available, create a skill file:

```markdown
# ~/.claude/skills/outlook/SKILL.md
---
name: outlook
description: Use when the user wants to read/send emails, check calendar, download attachments, or manage inbox rules via Outlook
---

The user has an Outlook MCP server configured. Use the `outlook_*` tools to interact with their Outlook account. Available tool categories:

- Email: outlook_list_emails, outlook_search_emails, outlook_send_email, outlook_reply_to_email, outlook_create_draft, etc.
- Calendar: outlook_list_events, outlook_create_event, outlook_check_availability, etc.
- Attachments: outlook_list_attachments, outlook_download_attachment
- Folders: outlook_list_folders, outlook_create_folder
- SharePoint: outlook_get_sharepoint_file, outlook_list_sharepoint_files
- Rules: outlook_list_rules, outlook_create_rule, outlook_delete_rule

Always call outlook_list_emails or outlook_search_emails before trying to operate on specific messages.
```

### AGENTS.md / system prompt guidance

If building an agent that uses this server, add to your system prompt:

```
You have access to Outlook via MCP tools prefixed with `outlook_`.
- Always search/list before operating on specific items (you need message IDs).
- Email send/reply is a high-stakes action — confirm with the user before sending.
- Attachment downloads may return parsed text content (PDF, Word, Excel) directly.
- Use outlook_search_emails for targeted queries; outlook_list_emails for browsing.
```

---

## Development

### Project structure

```
mcp-outlook/
├── server/
│   ├── index.js              # MCP server entry (114 lines)
│   ├── auth/                  # OAuth 2.0 PKCE authentication
│   │   ├── auth.js            # Core auth manager (338 lines)
│   │   ├── browserLauncher.js # Platform-specific browser open
│   │   ├── templates.js       # OAuth callback HTML pages
│   │   ├── tokenManager.js    # Encrypted token persistence
│   │   └── config.js          # OAuth configuration
│   ├── graph/                 # Microsoft Graph API client
│   ├── schemas/               # MCP tool schema definitions
│   ├── tools/
│   │   ├── dispatcher.js      # Tool name → handler registry
│   │   ├── common/            # Shared utilities (file parsing, logging)
│   │   ├── email/             # Email tools
│   │   ├── calendar/          # Calendar tools
│   │   ├── attachments/       # Attachment tools
│   │   ├── folders/           # Folder tools
│   │   ├── sharepoint/        # SharePoint tools
│   │   └── rules/             # Inbox rules tools
│   ├── utils/                 # Error handling, validation, caching
│   └── tests/                 # 262 tests (vitest)
├── package.json
└── vitest.config.js
```

### Running tests

```bash
npm test                      # Run all 262 tests
npm run test:watch            # Watch mode
npx vitest run --coverage     # With coverage report (~61%)
```

### Architecture decisions

- **Pure ESM** — no CommonJS, no transpilation
- **No client secret** — OAuth PKCE only, suitable for local/CLI deployment
- **Dispatcher pattern** — tool routing via registry map, not a switch statement
- **Conditional logging** — debug output gated behind `DEBUG` env var; `warn` level always on
- **Characterization tests** — tests written against existing behavior before refactoring, then maintained

---

## Security

- Tokens encrypted at rest (OS keychain or AES-256 with random key)
- No sensitive data in MCP tool responses (OAuth errors sanitized, stack traces removed)
- Email recipient validation before sending
- All Graph API calls scoped to authenticated user (`/me/` prefix)
- Debug logging gated behind `DEBUG` env var to prevent accidental data exposure

Report security issues via [GitHub Issues](https://github.com/w21180239/mcp-outlook/issues) (private disclosure preferred for critical issues).

---

## Acknowledgements

This project was originally forked from [XenoXilus/outlook-mcp](https://github.com/XenoXilus/outlook-mcp). It has since been substantially rewritten with a modular architecture, comprehensive test suite, security hardening, and new features.

## License

[MIT](LICENSE)

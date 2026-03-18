# mcp-outlook-lite

The lightest way to connect AI agents to Microsoft Outlook. No client secret. No complex OAuth. Just a Client ID and you're done.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Tests](https://img.shields.io/badge/tests-769%20passing-brightgreen)]()
[![Coverage](https://img.shields.io/badge/coverage-81%25-brightgreen)]()
[![TypeScript](https://img.shields.io/badge/TypeScript-strict-blue)]()
[![Node](https://img.shields.io/badge/node-%3E%3D18-blue)]()

> **Tired of getting stuck on Outlook MCP auth?** Most Outlook MCP servers require client secrets, complex permission grants, and multi-step OAuth configurations that break silently. This one uses **PKCE** — the browser handles login, no secrets stored anywhere. If you can create an Azure app registration, you can use this.

---

## Why this one?

| | mcp-outlook-lite | Other Outlook MCPs |
|---|---|---|
| **Auth setup** | Client ID only, zero secrets | Client ID + Client Secret + certificates |
| **Auth flow** | PKCE (browser popup) + device code (headless) | Complex OAuth requiring manual token management |
| **First-time experience** | Register app > paste ID > done | Register app > create secret > configure redirect > manage tokens > debug errors |
| **Token management** | Auto-refresh, encrypted at rest, zero maintenance | Often manual refresh or re-auth required |
| **Token efficiency** | Focused tool schemas, minimal response payloads | Verbose responses eating your context window |
| **Headless support** | Auto-detects SSH/containers, prints device code | Browser-only or manual token injection |

**The auth problem is real.** If you've tried other Outlook MCPs and got stuck after creating the Azure app — authorization failures, redirect URI mismatches, token exchange errors — that's because they use flows designed for server apps. PKCE is designed for exactly this use case: local tools that can't store secrets.

---

## 3-step setup

### Step 1: Register an Azure app (5 min, one-time)

1. [Azure Portal > App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) > **New registration**
2. Name: anything (e.g. `Outlook MCP`). Account type:
   - Work/school: *"Accounts in this organizational directory only"*
   - Personal: *"Accounts in any org directory and personal Microsoft accounts"*
3. Redirect URI: **Web** > `http://localhost/callback`
4. **Authentication** > enable **Allow public client flows** > Save
5. **API permissions** > Add **Microsoft Graph** delegated permissions:
   ```
   Mail.Read  Mail.ReadWrite  Mail.Send  Calendars.Read  Calendars.ReadWrite
   User.Read  MailboxSettings.Read  Files.Read.All  Sites.Read.All  offline_access
   ```
6. Copy **Application (client) ID** from the Overview page
7. Determine your **Tenant ID**:
   - **Personal account** (outlook.com / hotmail.com / live.com): use `consumers`
   - **Work/school account**: use the **Directory (tenant) ID** from the Overview page
   - **Both**: use `common`

> That's it. No client secret. No certificates. No admin consent (for personal accounts).
>
> ⚠️ **Personal account users**: You **must** set `AZURE_TENANT_ID=consumers`. Using the Directory (tenant) ID from Azure Portal will authenticate successfully but Graph API calls will return 401 because your mailbox lives in the consumer identity system, not in that Azure AD tenant.

### Step 2: Install

```bash
npx mcp-outlook-lite
```

Or add to your MCP client config:

<details>
<summary><b>Claude Code</b></summary>

```bash
# Personal account (outlook.com / hotmail.com / live.com)
claude mcp add outlook \
  -e AZURE_CLIENT_ID=your-client-id \
  -e AZURE_TENANT_ID=consumers \
  -- npx mcp-outlook-lite

# Work/school account
claude mcp add outlook \
  -e AZURE_CLIENT_ID=your-client-id \
  -e AZURE_TENANT_ID=your-directory-tenant-id \
  -- npx mcp-outlook-lite
```
</details>

<details>
<summary><b>Claude Desktop</b></summary>

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": ["mcp-outlook-lite"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_TENANT_ID": "consumers"
      }
    }
  }
}
```

> Replace `consumers` with your Directory (tenant) ID for work/school accounts.
</details>

<details>
<summary><b>Cursor / Windsurf / Other MCP clients</b></summary>

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": ["mcp-outlook-lite"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_TENANT_ID": "consumers"
      }
    }
  }
}
```

> Replace `consumers` with your Directory (tenant) ID for work/school accounts.
</details>

### Step 3: Use it

The first tool call triggers auth automatically:
- **Desktop**: browser opens for Microsoft login
- **SSH / container**: device code printed to stderr — follow the link

After that, tokens refresh silently. No re-login between sessions.

---

## 46 tools, 6 categories

| Category | Count | Highlights |
|----------|-------|------------|
| **Email** | 15 | List, search, send, reply, forward, draft, move, flag, categorize, batch |
| **Calendar** | 17 | Events, recurring meetings, availability, online meetings, timezone handling |
| **Attachments** | 4 | List, download with auto-parsing (PDF/Word/Excel/PPT), upload, scan |
| **Folders** | 4 | List, create, rename, stats |
| **SharePoint** | 3 | Access files via sharing links or direct IDs |
| **Rules** | 3 | List, create, delete server-side inbox rules |

### Example prompts

```
"Show me unread emails from this week"
"Find all emails from Alice about the budget"
"Reply to that email thanking her for the update"
"What meetings do I have tomorrow?"
"Schedule a 30-min call with Bob next Tuesday at 2pm"
"Download and summarize the PDF from the latest Finance email"
```

---

## How PKCE auth works

```
Agent calls a tool
       |
       v
  Token cached?  ---yes--->  Use it
       | no
  Refresh works? ---yes--->  Silent refresh (no browser)
       | no
  PKCE flow:
    1. Generate code_verifier + code_challenge
    2. Browser opens -> Microsoft login
    3. Redirect to localhost with auth code
    4. Exchange code + verifier for tokens
    5. Encrypt and store tokens locally
```

**No client secret anywhere in this flow.** The PKCE challenge/verifier pair cryptographically proves the caller's identity. Tokens are encrypted at rest using the OS keychain or AES-256 with a random key.

---

## Configuration

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_CLIENT_ID` | Yes | Application (client) ID from Azure |
| `AZURE_TENANT_ID` | Yes | `consumers` for personal accounts, Directory (tenant) ID for work/school, or `common` for both |
| `MCP_OUTLOOK_DEVICE_CODE` | No | Set to `1` to force device code flow |
| `MCP_OUTLOOK_WORK_DIR` | No | Directory for large file downloads |
| `DEBUG` | No | Enable debug logging on stderr |

---

## Development

**TypeScript** with `noImplicitAny`. 769 tests, 81% coverage.

```bash
npm test              # Run tests
npm run typecheck     # Type check
npm run build         # Compile to dist/
npm run dev           # Dev mode with tsx
```

<details>
<summary>Project structure</summary>

```
server/
  index.ts              # MCP server entry
  types.ts              # Shared interfaces
  auth/                 # PKCE + device code auth
  graph/                # Microsoft Graph client with rate limiting
  tools/                # 46 tool handlers
  schemas/              # MCP tool schemas
  utils/                # Validation, caching, error handling
  tests/                # 769 tests
```
</details>

---

## Security

- Tokens encrypted at rest (OS keychain or AES-256)
- All Graph API calls scoped to `/me/` (your mailbox only)
- No sensitive data in tool responses
- Recipient validation before sending emails

Report vulnerabilities via [GitHub private reporting](https://github.com/w21180239/mcp-outlook-lite/security). See [SECURITY.md](SECURITY.md).

---

## License

[MIT](LICENSE)

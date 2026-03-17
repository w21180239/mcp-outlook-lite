# mcp-outlook — Agent Instructions

This is an MCP server that provides 46 Outlook tools via Microsoft Graph API.

## Architecture

- Node.js ESM, no TypeScript
- Tests: vitest (`npm test`)
- Entry: `server/index.js` → dispatcher → tool handlers
- Auth: OAuth 2.0 PKCE in `server/auth/`
- All logging goes to stderr (stdout reserved for MCP protocol)

## Key conventions

- Tool handlers return MCP responses (never throw) — use `createToolError()` / `convertErrorToToolError()`
- Debug output: use `debug()` from `server/utils/logger.js` (gated by `DEBUG` env)
- Warnings: use `warn()` from `server/utils/logger.js` (always outputs)
- Never use `console.log` in production code (corrupts MCP stdio transport)
- All Graph API calls scoped to `/me/` (authenticated user only)

## Testing

```bash
npm test              # 262 tests
npm run test:watch    # watch mode
```

Test files live in `server/tests/`, mirroring the source structure. Use `server/tests/helpers/mockAuthManager.js` for mocking auth in tool tests.

## Adding a new tool

1. Create handler in `server/tools/<category>/`
2. Add schema in `server/schemas/<category>Schemas.js`
3. Export from `server/tools/index.js`
4. Register in `server/tools/dispatcher.js`
5. Write tests in `server/tests/tools/<category>/`

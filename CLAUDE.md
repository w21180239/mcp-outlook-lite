# mcp-outlook-lite — Agent Instructions

The lightest Outlook MCP server — zero secrets, PKCE auth, 46 tools via Microsoft Graph API.

## Architecture

- TypeScript, compiled to ES2022 ESM (`npm run build` → `dist/`)
- Tests: vitest (`npm test`) — 769 tests, ~81% coverage
- Entry: `server/index.ts` → dispatcher → tool handlers
- Auth: OAuth 2.0 PKCE + device code flow in `server/auth/`
- All logging goes to stderr (stdout reserved for MCP protocol)

## Key conventions

- Tool handlers return MCP responses (never throw) — use `createToolError()` / `convertErrorToToolError()`
- Debug output: use `debug()` from `server/utils/logger.js` (gated by `DEBUG` env)
- Warnings: use `warn()` from `server/utils/logger.js` (always outputs)
- Never use `console.log` in production code (corrupts MCP stdio transport)
- All Graph API calls scoped to `/me/` (authenticated user only)
- Shared types in `server/types.ts` — use `MCPResponse`, `ToolHandler`, etc.

## Testing

```bash
npm test              # 769 tests
npm run test:watch    # watch mode
npm run typecheck     # type checking
npm run build         # compile to dist/
```

Test files live in `server/tests/`, mirroring the source structure. Use `server/tests/helpers/mockAuthManager.js` for mocking auth in tool tests.

## Adding a new tool

1. Create handler in `server/tools/<category>/newTool.ts`
2. Add schema in `server/schemas/<category>Schemas.ts`
3. Export from `server/tools/index.ts`
4. Register in `server/tools/dispatcher.ts`
5. Write tests in `server/tests/tools/<category>/`

# mcp-outlook Project Improvements Design

## Goal

Improve mcp-outlook across 6 areas: npm publishing, CI/CD, security disclosure, TypeScript migration, test coverage, and headless authentication.

## Scope

### 1. npm Package Publishing (High Value)

**Decision:** Add `bin` field to package.json, add `prepublishOnly` script, publish as `mcp-outlook` on npm.

- Add `#!/usr/bin/env node` shebang to `server/index.js`
- Add `bin`, `files`, `engines`, `type`, `keywords`, `repository`, `bugs`, `homepage` fields to package.json
- Add `.npmignore` to exclude tests, coverage, .tokens, docs
- Users can then use `npx mcp-outlook` directly

### 2. GitHub Actions CI/CD (High Value)

**Decision:** Add `ci.yml` workflow running tests on push/PR to main.

- Node 18 + 20 matrix
- Install deps, run tests with coverage
- Upload coverage report as artifact
- Keep existing `release.yml` for DXT packaging

### 3. SECURITY.md (High Value)

**Decision:** Add standard SECURITY.md with GitHub private vulnerability reporting instructions.

### 4. TypeScript Migration (Medium Value)

**Decision:** Incremental migration using `allowJs: true` in tsconfig. Convert all `.js` files to `.ts`.

- Add `typescript`, `tsx`, `@types/node` as dev dependencies
- Add `tsconfig.json` with strict mode, ESM output
- Convert all source files in `server/` to TypeScript
- Keep tests in JavaScript initially, convert test helpers only
- Use `tsx` for runtime (no build step needed for development)
- Add `build` script using `tsc` for npm distribution
- Update package.json `main` to point to compiled output

**Key typing decisions:**
- Define interfaces for all Graph API response types
- Type all tool handler signatures as `(authManager: OutlookAuthManager, args: Record<string, unknown>) => Promise<MCPResponse>`
- Use strict null checks

### 5. Test Coverage Improvement (Medium Value)

**Decision:** Target 80% coverage. Focus on:

- Auth flow edge cases (token refresh failure, PKCE validation, concurrent auth)
- Graph client error paths (rate limiting, 5xx retries, timeout)
- Input validation edge cases
- Tool error handling paths

### 6. Device Code Flow for Headless Auth (Medium Value)

**Decision:** Add device code flow as fallback when no browser is available.

- Detect headless environment (no DISPLAY, SSH session, CI)
- Use `urn:ietf:params:oauth:grant-type:device_code` grant type
- Print device code + verification URL to stderr
- Poll token endpoint until user completes auth
- Same token storage/encryption as PKCE flow

## Architecture Impact

- Entry point changes from `server/index.js` to compiled `dist/server/index.js` for npm
- Dev workflow uses `tsx` for zero-build iteration
- All existing functionality preserved; TypeScript is additive
- Device code flow is an alternative auth path, not a replacement

## Testing Strategy

- Existing 262 tests continue passing throughout migration
- New tests added for device code flow, auth edge cases
- Coverage gate at 80% in CI

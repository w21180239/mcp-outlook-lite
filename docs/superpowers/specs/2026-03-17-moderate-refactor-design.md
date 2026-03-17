# Outlook-MCP Moderate Refactor

**Date:** 2026-03-17
**Author:** Wei Liu
**Status:** Approved

## Context

Freshly forked from XenoXilus/outlook-mcp (inactive since Jan 2025). The codebase works but has accumulated tech debt: a god-file dispatcher, duplicated utilities, oversized modules, ~25% test coverage, and inconsistent error handling. Goal is to bring it to maintainable quality with TDD.

## Approach: Hybrid characterization-test refactor

For each module: write tests against current behavior first, then refactor and verify tests pass. Module by module, not all at once.

## Pre-requisite: Install coverage tooling

Add `@vitest/coverage-v8` to devDependencies so we can measure the current baseline and track the 60% target. Remove unused `@azure/identity` from dependencies in the same commit.

## Phases

### Phase 1: Cleanup quick wins

**Changes:**
- Extract duplicated file utilities from **both** `tools/sharepoint/getSharePointFile.js` and `tools/attachments/downloadAttachment.js` into `tools/common/fileTypeUtils.js`:
  - `isTextContent()` — duplicated near-identically in both files
  - `isExcelFile()` — duplicated near-identically in both files
  - `formatFileSize()` — local copy in `downloadAttachment.js`; `getSharePointFile.js` uses `graphHelpers.general.formatFileSize()` instead (a third copy lives in `graph/graphHelpers.js`)
  - `parseExcelContent()`, `isOfficeDocument()`, `parseOfficeDocument()`, `decodeContent()` — duplicated between both files
- Update both consumers + attachment tools to import from the shared module
- Audit `graphHelpers.general.formatFileSize()` — if equivalent, replace callers with the shared version

**Tests:**
- Unit tests for each extracted utility (extension detection, MIME type matching, size formatting, content decoding)

**Risk:** Low. Mechanical extraction, but scope is larger than initially estimated due to 3+ copies of some functions.

### Phase 2: Extract tool dispatch from index.js

**Changes:**
- Create `tools/dispatcher.js` with a tool registry map: `{toolName → handlerFunction}`
- The existing barrel file `tools/index.js` already exports all tool functions — use it as the source, mapping MCP tool names (e.g. `outlook_list_emails`) to handler functions (e.g. `listEmailsTool`)
- Remove the redundant destructuring on lines 57-109 of `index.js` and the 40-case switch
- `index.js` retains only: MCP server init, transport setup, top-level request routing

**Tests:**
- Integration: every tool name in the registry resolves to a callable function
- Integration: tool schema names from `allToolSchemas` match registry keys exactly (schema-registry alignment)
- Integration: unknown tool name returns proper error
- Integration: tool dispatch calls handler with correct (authManager, args) signature

**Note:** This phase is a test-guarded refactor rather than strict TDD — the "before" and "after" tests are identical regression guards.

**Risk:** Medium. The switch statement is the spine of the server. Schema-registry alignment test is critical.

### Phase 3: Split auth.js (572 lines)

**Changes:**
- Extract HTML page templates (~200 lines of inline HTML) → `auth/templates.js`
- Extract browser launch logic → `auth/browserLauncher.js`
- Fix duplicate `this.isAuthenticated = false` assignment (lines 19-20)
- `auth.js` retains: PKCE flow, token exchange, silent refresh, ensureAuthenticated

**Tests:**
- Unit: `authenticate()` with valid cached token → returns success without browser
- Unit: `authenticate()` with expired token + valid refresh → silent refresh, no browser
- Unit: `authenticate()` with expired token + failed refresh → falls back to interactive
- Unit: `refreshAccessToken()` success path stores new tokens
- Unit: `refreshAccessToken()` failure clears tokens
- Unit: `exchangeCodeForToken()` success, HTTP error, and malformed response paths
- Unit: `getAuthorizationCode()` sets expected URL parameters (client_id, scope, code_challenge, state)
- Unit: templates render expected HTML strings

**Risk:** Medium. Auth is critical path. Tests must lock down behavior before splitting.

### Phase 4: Split getSharePointFile.js (1024 lines) and downloadAttachment.js (637 lines)

**Changes:**
- Extract document parsing (PDF, Word, PowerPoint) → `tools/common/documentParser.js`
- Extract Excel parsing → `tools/common/excelParser.js`
- Update **both** `getSharePointFile.js` and `downloadAttachment.js` to consume the shared parsers (Phase 1 handles utility dedup; this phase handles the parsing logic that uses those utilities)
- Original files retain: SharePoint URL resolution / attachment download, orchestration

**Known fragility:** `officeparser` uses non-standard `(data, err)` callback signature (data first). Characterization tests should pin this behavior.

**Tests:**
- Unit: `documentParser` handles each file type (docx, pptx, pdf)
- Unit: `documentParser` returns error for unsupported types
- Unit: `excelParser` extracts sheet data, handles empty sheets
- Unit: both tool files wire to shared parsers correctly (mock parsers, verify call args)

**Risk:** Medium. Document parsing has edge cases. Need test fixtures.

### Phase 5: Consistency pass

**Moved before tool test coverage phase to avoid test churn.**

**Changes:**
- Standardize error handling: resolve the return-vs-throw inconsistency across tools. Some tools return `createToolError(...)` while others throw `convertErrorToToolError(...)`. Pick one pattern (return for tools, throw for auth/internal) and apply consistently.
- Replace raw `console.error` debug logging with conditional logger that checks `process.env.DEBUG`
- Remove sensitive data from log output (email content, attachment data)

**Tests:**
- Unit: logger respects DEBUG env var
- Unit: error responses match expected MCP format across tool categories
- Verify existing 111 tests still pass after format changes

**Risk:** Low-medium. Error handling changes touch many files but are mechanical. Doing this before Phase 6 means the tool tests we write next will test the final, standardized behavior.

### Phase 6: Add tool test coverage

**Changes:** No production code changes. Test-only phase.

**Tests to add:**
- Email tools: listEmails, getEmail, searchEmails, sendEmail, replyToEmail, createDraft, markAsRead, deleteEmail
- Calendar tools: listEvents, createEvent, updateEvent, deleteEvent, respondToInvite
- Attachment tools: listAttachments, downloadAttachment
- Folder tools: listFolders, createFolder, moveEmail, renameFolder
- Rules tools: listRules, createRule, deleteRule

**Approach:** Mock `authManager.ensureAuthenticated()` to return a Graph client stub. Mock Graph API responses with realistic payloads.

**Target:** 60%+ overall coverage (measured via `@vitest/coverage-v8`).

**Risk:** Low. No production code changes. Tests assert against the standardized behavior from Phase 5.

## Out of scope

- New features (no new tools or capabilities)
- Schema changes (existing schemas stay as-is)
- `graph/` directory refactoring — `graphClient.js` (541 lines), `graphHelpers.js` (599 lines), `httpConfig.js` (414 lines) work correctly and are excluded from line-count targets for this pass
- `utils/` directory refactoring — `InputValidator.js` (611 lines), `ErrorHandler.js` (541 lines) are large but focused single-purpose modules; excluded from line-count targets
- `tools/common/sharedUtils.js` (623 lines) — styling/signature cache, leave as-is
- `tools/calendar/calendarUtils.js` (699 lines) — works, splitting deferred to a future pass
- Performance optimization (not a bottleneck)
- Migration to TypeScript (too large for this pass)

## Success criteria

1. All 111 existing tests continue to pass throughout
2. New tests bring measured coverage to 60%+ (via `vitest --coverage`)
3. Files targeted by this refactor are under 400 lines: `index.js`, `auth.js`, `getSharePointFile.js`, `downloadAttachment.js`
4. Zero duplicated utility functions between sharepoint and attachment tools
5. `index.js` under 100 lines
6. Live MCP server works: email list, calendar list, authenticate — verified manually after each phase

## Execution order

Phases are sequential. Each phase ends with: all tests pass, commit, manual smoke test. No phase starts until the previous one is clean.

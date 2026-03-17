# Outlook-MCP Moderate Refactor

**Date:** 2026-03-17
**Author:** Wei Liu
**Status:** Approved

## Context

Freshly forked from XenoXilus/outlook-mcp (inactive since Jan 2025). The codebase works but has accumulated tech debt: a god-file dispatcher, duplicated utilities, oversized modules, ~25% test coverage, and inconsistent error handling. Goal is to bring it to maintainable quality with TDD.

## Approach: Hybrid characterization-test refactor

For each module: write tests against current behavior first, then refactor and verify tests pass. Module by module, not all at once.

## Phases

### Phase 1: Cleanup quick wins

**Changes:**
- Remove unused `@azure/identity` dependency from package.json
- Extract duplicated file utilities (`isTextContent`, `isExcelFile`, `formatFileSize`) from `tools/sharepoint/getSharePointFile.js` and `tools/attachments/downloadAttachment.js` into `tools/common/fileTypeUtils.js`
- Update both consumers to import from the shared module

**Tests:**
- Unit tests for `fileTypeUtils.js` (extension detection, MIME type matching, size formatting)

**Risk:** Low. Mechanical extraction, easy to verify.

### Phase 2: Extract tool dispatch from index.js

**Changes:**
- Create `tools/dispatcher.js` with a tool registry map: `{toolName → handlerFunction}`
- Replace 40-case switch statement in `index.js` with registry lookup
- `index.js` retains only MCP server init, transport setup, and top-level request routing

**Tests:**
- Integration test: every tool name in the registry resolves to a callable function
- Integration test: unknown tool name returns proper error
- Integration test: tool dispatch calls handler with correct args

**Risk:** Medium. The switch statement is the spine of the server. Tests must verify every tool name before and after.

### Phase 3: Split auth.js (572 lines)

**Changes:**
- Extract HTML page templates (success, error, CSRF pages) → `auth/templates.js`
- Extract browser launch logic → `auth/browserLauncher.js`
- `auth.js` retains: PKCE flow, token exchange, silent refresh, ensureAuthenticated

**Tests:**
- Unit: `authenticate()` with valid cached token → returns success without browser
- Unit: `authenticate()` with expired token + valid refresh → silent refresh, no browser
- Unit: `authenticate()` with expired token + failed refresh → falls back to interactive
- Unit: `refreshAccessToken()` success path
- Unit: `refreshAccessToken()` failure clears tokens
- Unit: templates render expected HTML strings

**Risk:** Medium. Auth is critical path. Tests must lock down behavior before splitting.

### Phase 4: Split getSharePointFile.js (1024 lines)

**Changes:**
- Extract document parsing (PDF, Word, PowerPoint) → `tools/common/documentParser.js`
- Extract Excel parsing → `tools/common/excelParser.js`
- Original file retains: SharePoint URL resolution, file download, orchestration

**Tests:**
- Unit: `documentParser` handles each file type (docx, pptx, pdf)
- Unit: `documentParser` returns error for unsupported types
- Unit: `excelParser` extracts sheet data, handles empty sheets
- Unit: main tool wires parsers correctly (mock parsers, verify call args)

**Risk:** Medium. Document parsing has edge cases. Need test fixtures.

### Phase 5: Add tool test coverage

**Changes:** No production code changes. Test-only phase.

**Tests to add:**
- Email tools: listEmails, getEmail, searchEmails, sendEmail, replyToEmail, createDraft, markAsRead, deleteEmail
- Calendar tools: listEvents, createEvent, updateEvent, deleteEvent, respondToInvite
- Attachment tools: listAttachments, downloadAttachment
- Folder tools: listFolders, createFolder, moveEmail, renameFolder
- Rules tools: listRules, createRule, deleteRule

**Approach:** Mock `authManager.ensureAuthenticated()` to return a Graph client stub. Mock Graph API responses with realistic payloads.

**Target:** 60%+ overall coverage (up from ~25%).

**Risk:** Low. No production code changes.

### Phase 6: Consistency pass

**Changes:**
- Standardize error responses: all tools return MCP-format errors via `convertErrorToToolError()`
- Replace raw `console.error` debug logging with conditional logger that checks `process.env.DEBUG`
- Remove sensitive data from log output (email content, attachment data)

**Tests:**
- Unit: logger respects DEBUG env var
- Unit: error responses match expected MCP format across tool categories

**Risk:** Low. Mostly formatting and logging changes.

## Out of scope

- New features (no new tools or capabilities)
- Schema changes (existing schemas stay as-is)
- Graph client refactoring (rate limiting and retry logic works, leave it)
- Performance optimization (not a bottleneck)
- Migration to TypeScript (too large for this pass)

## Success criteria

1. All 111 existing tests continue to pass throughout
2. New tests bring coverage to 60%+
3. No file over 400 lines in the final state (except schemas, which are declarative)
4. Zero duplicated utility functions
5. `index.js` under 100 lines
6. Live MCP server works: email list, calendar list, authenticate — all verified manually after each phase

## Execution order

Phases are sequential. Each phase ends with: tests pass, commit, manual smoke test. No phase starts until the previous one is clean.

# Agent Guidelines for mcp-outlook

If you are an AI agent with access to this MCP server's tools (prefixed `outlook_`), follow these guidelines.

## Before operating on items

Always list or search first — you need the `messageId`, `eventId`, or `attachmentId` before you can read, modify, or delete anything.

## High-stakes actions

These actions have real-world consequences. Confirm with the user before executing:
- `outlook_send_email` — sends an actual email
- `outlook_reply_to_email` / `outlook_reply_all` / `outlook_forward_email` — sends a reply
- `outlook_delete_email` / `outlook_delete_event` / `outlook_delete_rule` — permanent deletion
- `outlook_create_rule` — affects future mail routing

## Attachment handling

`outlook_download_attachment` with `includeContent: true` automatically parses:
- PDF → extracted text
- Word (.docx) → extracted text
- PowerPoint (.pptx) → slide text
- Excel (.xlsx) → structured sheet data

Set `decodeContent: false` to get raw base64 instead.

## Search tips

- `outlook_search_emails` supports free-text search across all folders
- Use `from`, `subject`, `startDate`, `endDate` for targeted queries
- Set `includeBody: true` to get full email content (max 5 results)
- `outlook_list_emails` with `folder` parameter for browsing specific folders

## Error handling

If a tool returns `isError: true`, the error message describes what went wrong. Common causes:
- Authentication expired (will auto-refresh on next call)
- Missing required parameters
- Resource not found (wrong ID)

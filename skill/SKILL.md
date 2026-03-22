---
name: docx-mcp
description: "Use when editing existing Word (.docx) documents with track changes, comments, footnotes, or structural validation. Triggers: reviewing contracts, marking up reports, adding revision comments, validating document structure, removing watermarks. Requires the docx-mcp MCP server to be running."
---

# Editing Word Documents with docx-mcp

## Overview

The `docx-mcp` MCP server provides 18 tools for reading and editing .docx files with proper OOXML markup. Edits appear as real revisions in Microsoft Word — red strikethrough for deletions, green underline for insertions, comments in the sidebar.

## When to Use

- Editing existing .docx files with tracked changes
- Adding comments or footnotes to documents
- Reviewing/auditing document structure
- Removing watermarks
- Any task where changes must be visible as Word revisions

**Do NOT use for:** Creating new .docx from scratch (use docx-js instead), PDFs, spreadsheets.

## Workflow

```
1. open_document("/path/to/file.docx")
2. get_headings() or get_document_info()     → understand structure
3. search_text("clause text")                → find target paragraphs
4. delete_text(para_id, "old text")          → tracked deletion
5. insert_text(para_id, "new text")          → tracked insertion
6. add_comment(para_id, "Reason for change") → explain the edit
7. audit_document()                          → verify integrity
8. save_document("/path/to/output.docx")     → save (or omit path to overwrite)
```

Always `audit_document()` before saving to catch structural issues.

## Tool Quick Reference

| Tool | Purpose | Key args |
|------|---------|----------|
| `open_document` | Open .docx for editing | `path` |
| `close_document` | Close and clean up | — |
| `get_document_info` | Stats overview | — |
| `get_headings` | Heading tree with paraIds | — |
| `search_text` | Find text in body/footnotes/comments | `query`, `regex` |
| `get_paragraph` | Full text of one paragraph | `para_id` |
| `insert_text` | Tracked insertion (green underline) | `para_id`, `text`, `position` |
| `delete_text` | Tracked deletion (red strikethrough) | `para_id`, `text` |
| `add_comment` | Comment anchored to paragraph | `para_id`, `text` |
| `reply_to_comment` | Threaded reply | `parent_id`, `text` |
| `get_comments` | List all comments | — |
| `add_footnote` | Footnote with superscript ref | `para_id`, `text` |
| `get_footnotes` | List all footnotes | — |
| `validate_footnotes` | Cross-ref footnote IDs | — |
| `validate_paraids` | Check paraId uniqueness | — |
| `remove_watermark` | Remove DRAFT watermarks | — |
| `audit_document` | Full structural audit | — |
| `save_document` | Save to .docx | `output_path` (optional) |

## Tips

- **paraId** is an 8-char hex string (e.g., `"1A2B3C4D"`). Get them from `get_headings()` or `search_text()`.
- **position** in `insert_text`: `"start"`, `"end"`, or a substring to insert after.
- **author** defaults to `"Claude"` for all tracked changes and comments.
- **Replace text**: `delete_text` then `insert_text` on the same paragraph — Word shows both marks.
- **Save to new file** to preserve the original: `save_document("/path/to/revised.docx")`.

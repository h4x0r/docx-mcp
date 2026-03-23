# Design: `create_document` and `create_from_markdown` Tools

**Date**: 2026-03-23
**Status**: Approved
**Tool count**: 43 → 45

---

## Problem

docx-mcp can only edit existing .docx files. Users need to create documents from scratch — either blank or from markdown content. This is the #1 gap identified in competitive analysis (5+ competitors offer document creation).

## Design Decisions

| Decision | Choice | Rationale |
|---|---|---|
| Markdown features | Full GitHub-Flavored Markdown | Covers headings, bold/italic/strikethrough, links, images, lists, code, tables, footnotes, task lists, blockquotes |
| Input modes | File path OR raw text | File path for existing .md files, raw text for agent-generated content |
| Styles | Built-in + custom fallback | `Heading 1`–`9`, `Normal`, `List Bullet`, `List Number` are built-in. `CodeBlock`, `BlockQuote`, `FootnoteRef` are custom styles for constructs without built-in equivalents |
| Images | Local embed, remote as hyperlinks | Local paths get embedded. Remote URLs become clickable hyperlinks with alt text. No network calls in MCP server |
| Templates | Optional .dotx support | `create_document` and `create_from_markdown` accept optional `template_path` to inherit styles/headers/footers/page setup |
| Smart typography | Auto-convert quotes and dashes | Straight quotes → curly quotes, `--` → en dash, `---` → em dash, `...` → ellipsis. Not applied inside code spans/blocks |

---

## Architecture

### Approach: New mixin + standalone converter

Two new modules:

1. **`docx_mcp/document/creation.py`** — `CreationMixin` added to `DocxDocument`. Handles blank .docx skeleton and template-based creation.
2. **`docx_mcp/markdown.py`** — Standalone `MarkdownConverter` class. Parses markdown, walks AST, populates a `DocxDocument` via internal mixin methods.

```
Agent calls create_from_markdown(output_path, markdown="# Hello\n...")
    → server.py creates blank doc via CreationMixin.create()
    → MarkdownConverter.convert(doc, markdown_text)
        → mistune parses markdown → AST
        → walker calls doc's internal methods per node type
    → doc is auto-opened as the active document
    → agent can immediately edit/save/audit
```

### New dependency

`mistune>=3.0` — pure Python markdown parser, supports GFM plugins (tables, footnotes, strikethrough, task lists).

---

## Tool 1: `create_document`

### Signature

```python
@mcp.tool()
def create_document(
    output_path: str,
    template_path: str | None = None,
) -> str:
```

### Blank mode (no template)

Builds a minimal valid .docx ZIP containing:

| Part | Purpose |
|---|---|
| `[Content_Types].xml` | Content type mappings |
| `_rels/.rels` | Root relationships |
| `word/document.xml` | Empty body with one blank paragraph (paraId generated) |
| `word/_rels/document.xml.rels` | Relationships to styles, settings, footnotes, endnotes, header |
| `word/styles.xml` | Built-in styles + custom styles (CodeBlock, BlockQuote, FootnoteRef) |
| `word/settings.xml` | Basic settings with track-changes author |
| `word/footnotes.xml` | Separator/continuation boilerplate (id=-1, id=0) |
| `word/endnotes.xml` | Separator/continuation boilerplate (id=-1, id=0) |
| `word/header1.xml` | Empty header |
| `docProps/core.xml` | creator="docx-mcp", created=now |

### Template mode

1. Copy the .dotx file to `output_path` (as .docx)
2. Open it normally via the existing `open_document` path

### Post-creation behavior

- Document is automatically opened (set as the `_doc` singleton)
- If a document is already open, it is closed first
- Returns document info (same as `get_document_info`)

---

## Tool 2: `create_from_markdown`

### Signature

```python
@mcp.tool()
def create_from_markdown(
    output_path: str,
    md_path: str | None = None,
    markdown: str | None = None,
    template_path: str | None = None,
) -> str:
```

- `md_path` and `markdown` are mutually exclusive; exactly one required
- `template_path` is optional — if provided, document starts from template

### Markdown → OOXML Mapping

| Markdown construct | OOXML output | Implementation |
|---|---|---|
| `# Heading` (1–6) | `w:p` with `Heading N` pStyle | Direct XML |
| Paragraph text | `w:p` with `Normal` style | Direct XML |
| `**bold**` | `w:rPr > w:b` | Direct XML (run properties) |
| `*italic*` | `w:rPr > w:i` | Direct XML |
| `~~strikethrough~~` | `w:rPr > w:strike` | Direct XML |
| `[link](url)` | `w:hyperlink` with relationship | Direct XML + rels |
| `![img](local_path)` | Embedded image | ImagesMixin internals |
| `![img](remote_url)` | Hyperlink with alt text | Direct XML + rels |
| `- bullet` | `w:p` with `List Bullet` + `w:numPr` | ListsMixin internals |
| `1. numbered` | `w:p` with `List Number` + `w:numPr` | ListsMixin internals |
| Nested lists | `w:ilvl` for indent level | Direct XML |
| `` `inline code` `` | `w:rPr` with Courier New font | Direct XML |
| ```` ```code block``` ```` | `w:p` with `CodeBlock` custom style | Direct XML |
| `> blockquote` | `w:p` with `BlockQuote` custom style | Direct XML |
| `---` horizontal rule | `w:p` with bottom border | Direct XML |
| `\| table \|` | `w:tbl` with rows/cells | TablesMixin internals |
| `[^1]` footnote ref | `w:footnoteReference` | FootnotesMixin internals |
| `[^1]: text` footnote def | `w:footnote` in footnotes.xml | FootnotesMixin internals |
| `- [x]` task list | Bullet with checkbox character (☐/☑) | Direct XML |

### Smart Typography

Applied to all text content except code spans and code blocks:

| Input | Output | Unicode |
|---|---|---|
| `"text"` | \u201Ctext\u201D | Left/right double quotes |
| `'text'` | \u2018text\u2019 | Left/right single quotes |
| `it's` | it\u2019s | Smart apostrophe |
| `--` | \u2013 | En dash |
| `---` | \u2014 | Em dash |
| `...` | \u2026 | Ellipsis |

### Processing flow

1. If `template_path` provided: create from template, else create blank
2. Parse markdown via mistune with GFM plugins (table, footnote, strikethrough, task_list)
3. Remove the initial blank paragraph from the skeleton
4. Walk AST nodes, generating OOXML elements via mixin internals
5. All paragraphs get paraIds via `_new_para_id()`
6. Call `audit_document()` internally to validate the result
7. Auto-open the document as the active `_doc` singleton
8. Return document info

---

## Custom Styles

Defined in the blank skeleton's `word/styles.xml`:

### CodeBlock

- Font: Courier New, 9pt
- Background shading: light gray (#F2F2F2)
- Spacing: 0 before/after (code blocks are contiguous)
- No smart typography applied

### BlockQuote

- Left indent: 0.5 inch
- Left border: 3pt solid gray (#AAAAAA)
- Italic
- Color: dark gray (#555555)

### FootnoteRef

- Superscript
- Font size: inherit minus 2pt

---

## Files Changed/Created

| File | Action | Description |
|---|---|---|
| `docx_mcp/document/creation.py` | **New** | `CreationMixin` with `create()` method |
| `docx_mcp/document/__init__.py` | Edit | Add `CreationMixin` to `DocxDocument` bases |
| `docx_mcp/markdown.py` | **New** | `MarkdownConverter` class |
| `docx_mcp/server.py` | Edit | Add `create_document` and `create_from_markdown` tools |
| `docx_mcp/skill/SKILL.md` | Edit | Add new tools to reference tables |
| `pyproject.toml` | Edit | Add `mistune>=3.0`, bump version |
| `README.md` | Edit | Update tool count (43→45), add tools to tables |
| `tests/test_creation.py` | **New** | Blank skeleton, template, auto-open, styles, paraIds |
| `tests/test_markdown.py` | **New** | One test per markdown construct + edge cases |
| `tests/test_e2e.py` | Edit | Roundtrip tests for both new tools |
| `tests/conftest.py` | Edit | Add template fixture if needed |

---

## Testing Strategy

### `tests/test_creation.py`

- Blank document has valid XML structure (all required parts present)
- Blank document passes `audit_document()` with no issues
- All expected styles exist (Normal, Heading 1–9, List Bullet, List Number, CodeBlock, BlockQuote)
- ParaIds are valid (unique, < 0x80000000, 8 hex digits)
- Template mode copies .dotx and opens as .docx
- Template mode preserves template styles and headers
- Auto-opens document after creation (singleton is set)
- Closes previous document if one is open
- Returns document info

### `tests/test_markdown.py`

One test per construct:
- Headings (H1–H6)
- Paragraph text
- Bold, italic, strikethrough, bold+italic combo
- Links (external hyperlinks)
- Images (local path embedded)
- Images (remote URL as hyperlink)
- Bullet lists (flat and nested)
- Numbered lists (flat and nested)
- Inline code
- Code blocks (with and without language tag)
- Blockquotes (including nested)
- Horizontal rules
- Tables (with header row)
- Footnotes (reference + definition)
- Task lists (checked and unchecked)
- Smart typography (quotes, dashes, ellipsis)
- Smart typography NOT applied in code
- Mixed constructs (heading + paragraph + list + table in one document)
- Empty input
- Mutually exclusive input validation (both md_path and markdown provided → error)

### `tests/test_e2e.py` additions

- `create_document` → save → reopen → verify structure
- `create_from_markdown` → save → reopen → verify all content types survive roundtrip
- `create_from_markdown` with template → verify template styles preserved
- `create_from_markdown` → edit with track changes → save → verify revisions on created content

### Coverage target: 100%

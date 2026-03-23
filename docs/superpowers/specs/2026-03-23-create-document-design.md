# Design: `create_document` and `create_from_markdown` Tools

**Date**: 2026-03-23
**Status**: Approved (post spec-review fixes)
**Tool count**: 43 → 45

---

## Problem

docx-mcp can only edit existing .docx files. Users need to create documents from scratch — either blank or from markdown content. This is the #1 gap identified in competitive analysis (5+ competitors offer document creation).

## Design Decisions

| Decision | Choice | Rationale |
|---|---|---|
| Markdown features | Full GitHub-Flavored Markdown | Covers headings, bold/italic/strikethrough, links, images, lists, code, tables, footnotes, task lists, blockquotes |
| Input modes | File path OR raw text | File path for existing .md files, raw text for agent-generated content |
| Styles | Built-in + custom fallback | `Heading 1`–`6`, `Normal`, `List Bullet`, `List Number` are built-in. `CodeBlock`, `BlockQuote` are custom styles for constructs without built-in equivalents |
| Images | Local embed, remote as hyperlinks | Local paths get embedded. Remote URLs become clickable hyperlinks with alt text. No network calls in MCP server |
| Templates | Optional .dotx support | `create_document` and `create_from_markdown` accept optional `template_path` to inherit styles/headers/footers/page setup |
| Smart typography | Auto-convert quotes and dashes | Straight quotes → curly quotes, `--` → en dash, `---` → em dash, `...` → ellipsis. Not applied inside code spans/blocks |

---

## Architecture

### Approach: New mixin + standalone converter

Two new modules:

1. **`docx_mcp/document/creation.py`** — `CreationMixin` added to `DocxDocument`. Handles blank .docx skeleton and template-based creation.
2. **`docx_mcp/markdown.py`** — Standalone `MarkdownConverter` class. Parses markdown, walks AST, builds OOXML directly (not via public mixin tool methods).

### Creation flow (how it integrates with BaseMixin)

`CreationMixin.create(output_path, template_path=None)` works as follows:

1. **Blank mode**: Build the skeleton ZIP from XML string templates, write it to `output_path`
2. **Template mode**: Copy the .dotx to `output_path` as .docx
3. Call `self.open()` on the now-existing file — this uses the normal `BaseMixin.__init__` + `open()` path since the file exists on disk
4. The document is now fully loaded in memory, ready for editing

In `server.py`, the tool does:
```python
global _doc
if _doc is not None:
    _doc.close()
_doc = DocxDocument(output_path)
_doc.create(output_path, template_path)  # writes skeleton, then calls self.open()
```

Wait — this requires the file to exist before `DocxDocument(path)` validates it. Instead:

```python
global _doc
if _doc is not None:
    _doc.close()
# CreationMixin.create is a classmethod that writes the file, then returns an opened instance
_doc = DocxDocument.create(output_path, template_path)
```

**`create()` is a `@classmethod`** that:
1. Writes the skeleton ZIP to `output_path` (or copies template)
2. Instantiates `DocxDocument(output_path)`
3. Calls `.open()` on it
4. Returns the instance

This avoids modifying `BaseMixin.__init__` or `open()`.

### Markdown converter interaction with mixins

The `MarkdownConverter` does **NOT** call public mixin methods like `add_list()`, `insert_image()`, or `add_footnote()`. Those methods are designed for editing existing documents (they locate paragraphs by paraId, generate tracked changes, etc.).

Instead, the converter **builds OOXML directly** using lxml, the same way the mixins do internally. It has access to:
- `doc._trees` — the cached XML trees
- `doc._new_para_id()` / `doc._new_text_id()` — ID generators
- `doc.workdir` — for image embedding
- Namespace constants from `base.py`

This is the same pattern used within the mixin files themselves — they all manipulate `self._trees` directly.

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

| Part | Purpose | Content-Type Override |
|---|---|---|
| `[Content_Types].xml` | Content type mappings (with all Override entries below) | N/A |
| `_rels/.rels` | Root relationships | N/A |
| `word/document.xml` | Empty body with one blank paragraph (paraId + textId generated) | `application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml` |
| `word/_rels/document.xml.rels` | Relationships to styles, settings, footnotes, endnotes, header, numbering | N/A |
| `word/styles.xml` | Built-in styles + custom styles (CodeBlock, BlockQuote) | `application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml` |
| `word/settings.xml` | Basic settings with track-changes author | `application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml` |
| `word/numbering.xml` | Multi-level list definitions for bullets and numbered lists (9 levels) | `application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml` |
| `word/footnotes.xml` | Separator/continuation boilerplate (id=-1, id=0) | `application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml` |
| `word/endnotes.xml` | Separator/continuation boilerplate (id=-1, id=0) | `application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml` |
| `word/header1.xml` | Empty header | `application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml` |
| `docProps/core.xml` | creator="docx-mcp", created=now (uses `dc:`, `dcterms:`, `cp:` namespace prefixes from base.py constants) | `application/vnd.openxmlformats-package.core-properties+xml` |

All paragraphs get both `w14:paraId` and `w14:textId` attributes.

### numbering.xml — Multi-level list definitions

The blank skeleton includes `word/numbering.xml` with two `w:abstractNum` entries:

1. **Bullet list** (abstractNumId="0"): 9 levels (`w:ilvl` 0–8), each with appropriate bullet character and increasing indent
2. **Numbered list** (abstractNumId="1"): 9 levels, alternating decimal/lowerLetter/lowerRoman patterns with increasing indent

Plus corresponding `w:num` entries (numId="1" for bullets, numId="2" for numbered). The relationship is registered in `document.xml.rels` and the content type in `[Content_Types].xml`.

### Template mode

1. Copy the .dotx file to `output_path` (as .docx)
2. Open it normally via the existing `open()` path
3. If template lacks custom styles (CodeBlock, BlockQuote), add them
4. If template lacks `numbering.xml`, bootstrap it with the multi-level definitions

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
- If `md_path` is provided, image paths in the markdown are resolved relative to the markdown file's directory

### Markdown → OOXML Mapping

| Markdown construct | OOXML output | Implementation |
|---|---|---|
| `# Heading` (1–6) | `w:p` with `Heading N` pStyle | Direct XML |
| Paragraph text | `w:p` with `Normal` style | Direct XML |
| `**bold**` | `w:rPr > w:b` | Direct XML (run properties) |
| `*italic*` | `w:rPr > w:i` | Direct XML |
| `~~strikethrough~~` | `w:rPr > w:strike` | Direct XML |
| `[link](url)` | `w:hyperlink` with relationship | Direct XML + rels |
| `![img](local_path)` | Embedded image (error if file not found → skip with `[Image not found: path]` placeholder text) | Direct XML + rels + media copy |
| `![img](remote_url)` | Hyperlink with alt text | Direct XML + rels |
| `- bullet` | `w:p` with `w:numPr` (numId=1, ilvl=0) | Direct XML |
| `1. numbered` | `w:p` with `w:numPr` (numId=2, ilvl=0) | Direct XML |
| Nested lists | `w:ilvl` matching nesting depth (0–8) using multi-level abstractNum | Direct XML |
| `` `inline code` `` | `w:rPr` with Courier New font | Direct XML |
| ```` ```code block``` ```` | `w:p` with `CodeBlock` custom style, one `w:p` per line | Direct XML |
| `> blockquote` | `w:p` with `BlockQuote` custom style | Direct XML |
| Nested `>> blockquote` | `BlockQuote` style with increasing left indent per nesting level | Direct XML |
| `---` horizontal rule | `w:p` with bottom border | Direct XML |
| `\| table \|` | `w:tbl` with rows/cells, first row bold (header) | Direct XML |
| `[^1]` footnote ref | `w:footnoteReference` with `w:rStyle="FootnoteReference"` (Word built-in) | Direct XML |
| `[^1]: text` footnote def | `w:footnote` in footnotes.xml with paraId/textId | Direct XML |
| `- [x]` task list | Bullet with checkbox character (☐ U+2610 / ☑ U+2611) | Direct XML |

Note: Mistune only parses heading levels 1–6. Heading styles 7–9 exist in `styles.xml` for manual use via editing tools but are never produced by the markdown converter.

### Smart Typography

Applied to all text content **except** code spans and code blocks:

| Input | Output | Unicode |
|---|---|---|
| `"text"` | \u201Ctext\u201D | Left/right double quotes |
| `'text'` | \u2018text\u2019 | Left/right single quotes |
| `it's` | it\u2019s | Smart apostrophe |
| `--` | \u2013 | En dash |
| `---` | \u2014 | Em dash |
| `...` | \u2026 | Ellipsis |

**Apostrophe vs. quote heuristic:** A single `'` is treated as:
- **Left single quote** if preceded by whitespace, start-of-string, or opening punctuation (`([{`)
- **Apostrophe (right single quote)** in all other cases (after a letter, digit, or closing punctuation)

This handles `'twas` (apostrophe), `she said 'hello'` (quotes), and `it's` (apostrophe) correctly. Edge cases that are ambiguous default to apostrophe since contractions are more common than single-quoted strings.

### Processing flow

1. Validate inputs: exactly one of `md_path` or `markdown` must be provided
2. If `md_path`: read file, set base directory for image resolution
3. If `template_path` provided: `create_document(output_path, template_path)`, else `create_document(output_path)`
4. Parse markdown via mistune with GFM plugins (table, footnote, strikethrough, task_list)
5. Remove the initial blank paragraph from the skeleton
6. Walk AST nodes, generating OOXML elements directly in `doc._trees["word/document.xml"]`
7. All paragraphs get both `w14:paraId` and `w14:textId` via `_new_para_id()` / `_new_text_id()`
8. Auto-open the document as the active `_doc` singleton
9. Return document info

Note: `audit_document()` is **not** called internally during creation. The audit tool checks for artifact markers (DRAFT, TODO, FIXME) which may appear legitimately in markdown content. Users should call `audit_document()` explicitly if they want validation after creation.

### Template mode specifics

When `template_path` is provided with `create_from_markdown`:
- Template body content is **cleared** (all `w:p` and `w:tbl` elements removed from `w:body`, preserving only `w:sectPr`)
- Markdown content is then inserted into the empty body
- Template styles are preserved; custom styles (CodeBlock, BlockQuote) are added only if missing
- If template lacks `footnotes.xml`, `endnotes.xml`, or `numbering.xml`, they are bootstrapped

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

Note: Footnote references use Word's built-in `FootnoteReference` character style (superscript). No custom style needed.

---

## Files Changed/Created

| File | Action | Description |
|---|---|---|
| `docx_mcp/document/creation.py` | **New** | `CreationMixin` with `create()` classmethod |
| `docx_mcp/document/__init__.py` | Edit | Add `CreationMixin` to `DocxDocument` bases |
| `docx_mcp/markdown.py` | **New** | `MarkdownConverter` class |
| `docx_mcp/server.py` | Edit | Add `create_document` and `create_from_markdown` tools |
| `docx_mcp/skill/SKILL.md` | Edit | Add new tools to reference tables |
| `pyproject.toml` | Edit | Add `mistune>=3.0`, bump version to 0.3.0 |
| `README.md` | Edit | Update tool count (43→45), add tools to tables |
| `tests/test_creation.py` | **New** | Blank skeleton, template, auto-open, styles, paraIds |
| `tests/test_markdown.py` | **New** | One test per markdown construct + edge cases |
| `tests/test_e2e.py` | Edit | Roundtrip tests for both new tools |
| `tests/conftest.py` | Edit | Add template fixture if needed |

---

## Testing Strategy

### `tests/test_creation.py`

- Blank document has valid XML structure (all required parts present)
- All `[Content_Types].xml` Override entries are present
- `word/numbering.xml` exists with multi-level bullet and numbered definitions
- All expected styles exist (Normal, Heading 1–6, List Bullet, List Number, CodeBlock, BlockQuote)
- ParaIds and textIds are valid (unique, < 0x80000000, 8 hex digits)
- Template mode copies .dotx and opens as .docx
- Template mode preserves template styles and headers
- Template mode adds missing custom styles
- Template mode adds missing numbering.xml
- Auto-opens document after creation (singleton is set)
- Closes previous document if one is open
- Returns document info
- `create()` is a classmethod returning DocxDocument instance

### `tests/test_markdown.py`

One test per construct:
- Headings (H1–H6)
- Paragraph text
- Bold, italic, strikethrough, bold+italic combo
- Links (external hyperlinks)
- Images (local path embedded)
- Images (remote URL as hyperlink)
- Images (missing local file → placeholder text)
- Bullet lists (flat and nested 3 levels deep)
- Numbered lists (flat and nested 3 levels deep)
- Inline code
- Code blocks (with and without language tag)
- Blockquotes (flat and nested)
- Horizontal rules
- Tables (with header row, bold first row)
- Footnotes (reference + definition, correct cross-refs)
- Task lists (checked and unchecked)
- Smart typography (quotes, dashes, ellipsis)
- Smart typography NOT applied in code
- Apostrophe vs single quote heuristic
- Mixed constructs (heading + paragraph + list + table in one document)
- Empty input produces document with no body paragraphs
- Mutually exclusive input validation (both md_path and markdown provided → error)
- Neither md_path nor markdown provided → error
- Image paths resolve relative to markdown file directory

### `tests/test_e2e.py` additions

- `create_document` → save → reopen → verify structure
- `create_from_markdown` → save → reopen → verify all content types survive roundtrip
- `create_from_markdown` with template → verify template styles preserved, body content replaced
- `create_from_markdown` → edit with track changes → save → verify revisions on created content

### Coverage target: 100%

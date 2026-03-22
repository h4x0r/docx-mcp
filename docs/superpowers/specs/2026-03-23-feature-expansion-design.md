# docx-mcp Feature Expansion Design

**Date:** 2026-03-23
**Approach:** Modular mixin-based package (Approach B)
**Method:** TDD — tests first, 100% coverage maintained throughout
**Current state:** 18 tools, ~800-line document.py, 100% coverage

## Overview

Expand docx-mcp from 18 to ~45 tools covering tables, formatting, images, document properties, headers/footers, styles, lists, endnotes, sections, merge, cross-references, and protection. Restructure document.py into a mixin-based package to keep files manageable.

## Architecture

### Package Structure After Refactor

```
docx_mcp/
  __init__.py              # unchanged
  __main__.py              # unchanged
  server.py                # thin tool wrappers (grows from 299 to ~600 lines)
  document/
    __init__.py            # DocxDocument(all mixins), re-exports
    base.py                # BaseMixin: open/close/save, _parts cache, namespaces, paraId generation
    reading.py             # ReadingMixin: get_headings, search_text, get_paragraph, get_document_info
    tracks.py              # TracksMixin: insert_text, delete_text, accept_changes, reject_changes
    comments.py            # CommentsMixin: get_comments, add_comment, reply_to_comment
    footnotes.py           # FootnotesMixin: get_footnotes, add_footnote, validate_footnotes
    validation.py          # ValidationMixin: validate_paraids, remove_watermark, audit_document
    tables.py              # TablesMixin: get_tables, add_table, modify_cell, add_table_row, delete_table_row
    formatting.py          # FormattingMixin: set_formatting
    styles.py              # StylesMixin: get_styles
    headers_footers.py     # HeadersFootersMixin: get_headers_footers, edit_header_footer
    properties.py          # PropertiesMixin: get_properties, set_properties
    images.py              # ImagesMixin: get_images, insert_image
    endnotes.py            # EndnotesMixin: get_endnotes, add_endnote, validate_endnotes
    lists.py               # ListsMixin: add_list
    sections.py            # SectionsMixin: add_section_break, set_section_properties, add_page_break
    merge.py               # MergeMixin: merge_documents
    references.py          # ReferencesMixin: add_cross_reference
    protection.py          # ProtectionMixin: set_document_protection
```

### Mixin Composition

```python
# document/__init__.py
from .base import BaseMixin
from .reading import ReadingMixin
from .tracks import TracksMixin
# ... all mixins

class DocxDocument(
    BaseMixin,
    ReadingMixin,
    TracksMixin,
    CommentsMixin,
    FootnotesMixin,
    ValidationMixin,
    TablesMixin,
    FormattingMixin,
    StylesMixin,
    HeadersFootersMixin,
    PropertiesMixin,
    ImagesMixin,
    EndnotesMixin,
    ListsMixin,
    SectionsMixin,
    MergeMixin,
    ReferencesMixin,
    ProtectionMixin,
):
    """Word document editor with OOXML-level control."""
    pass
```

Each mixin accesses shared state via `self._parts`, `self._tmp_dir`, `self._new_para_id()`, etc., defined in `BaseMixin`.

### Import Compatibility

`server.py` currently does `from docx_mcp.document import DocxDocument`. After refactor, `document/__init__.py` exports `DocxDocument` — zero changes needed in server.py or tests.

## Phases

### Phase 0: Refactor (no behavior changes)

Extract existing document.py into package. No new features. Verify 100% coverage unchanged.

**Steps:**
1. Create `docx_mcp/document/` package
2. Move lifecycle/cache/namespace code to `base.py`
3. Move existing methods to corresponding mixin files
4. Compose `DocxDocument` in `__init__.py`
5. Run existing tests — must pass identically
6. Verify 100% coverage

### Phase 1: Reading (6 new tools)

Read-only tools. No XML mutation. Zero risk to existing functionality.

| Tool | Method | Returns |
|------|--------|---------|
| `get_tables` | `tables()` | List of dicts: index, row_count, col_count, header_row text, all cells as nested lists |
| `get_styles` | `styles()` | List of dicts: id, name, type (paragraph/character/table), base_style |
| `get_headers_footers` | `headers_footers()` | List of dicts: section_index, type (default/first/even), location (header/footer), text |
| `get_properties` | `properties()` | Dict: title, creator, subject, description, created, modified, revision, lastModifiedBy |
| `get_images` | `images()` | List of dicts: rId, filename, content_type, width_emu, height_emu |
| `get_endnotes` | `endnotes()` | List of dicts: id, text (excludes separator endnotes id=0, -1) |

**Test fixture:** Extend conftest.py test .docx to include a 2x3 table, custom styles, header/footer, core.xml properties, an embedded image, and endnotes.

**Test file:** `tests/test_reading.py`

### Phase 2: Track Changes Complete (3 new tools)

| Tool | Method | Args | Behavior |
|------|--------|------|----------|
| `accept_changes` | `accept_changes()` | `scope` (all/by_id/by_author), `ids` (optional list), `author` (optional) | Removes `w:ins` wrapper keeping content, removes `w:del` elements entirely, removes `w:rPrChange` keeping new formatting, handles `w:moveTo` (keep content) and `w:moveFrom` (remove) |
| `reject_changes` | `reject_changes()` | Same as accept | Removes `w:ins` elements entirely, removes `w:del` wrapper keeping content, reverts `w:rPrChange` to original formatting, handles `w:moveFrom` (keep content) and `w:moveTo` (remove) |
| `set_formatting` | `set_formatting()` | `para_id`, `text` (substring to format), `bold`, `italic`, `underline`, `font_name`, `font_size`, `color` | Splits run if needed, applies new `w:rPr` properties, stores original formatting as `w:rPrChange` child inside the new `w:rPr` (rPrChange is a child of rPr, not a wrapper around it) |

**Test file:** `tests/test_tracks.py`

### Phase 3: Tables Write (4 new tools)

| Tool | Method | Args | Behavior |
|------|--------|------|----------|
| `add_table` | `add_table()` | `para_id` (insert after), `rows`, `cols`, `header` (optional list), `data` (optional nested list) | Creates `w:tbl` with `w:tblGrid`, `w:tr`, `w:tc` elements. Wrapped in `w:ins` for tracking. |
| `modify_cell` | `modify_cell()` | `table_index`, `row`, `col`, `text` | Tracked delete of old text + tracked insert of new text in the cell's paragraph |
| `add_table_row` | `add_table_row()` | `table_index`, `position` (index or "end"), `cells` (list of text) | Insert `w:tr` wrapped in `w:ins` |
| `delete_table_row` | `delete_table_row()` | `table_index`, `row` | Mark row as deleted by wrapping each cell's paragraph runs in `w:del` and marking paragraph marks with `w:rPr > w:del` (w:del cannot wrap w:tr directly — must be applied per-cell content) |

**Test file:** `tests/test_tables.py`

### Phase 4: Content Creation (5 new tools)

| Tool | Method | Args | Behavior |
|------|--------|------|----------|
| `add_list` | `add_list()` | `para_id`, `items` (list of strings), `list_type` ("bullet"/"number") | Creates `w:abstractNum` + `w:num` in numbering.xml, inserts paragraphs with `w:numPr` after target para. Wrapped in `w:ins`. |
| `insert_image` | `insert_image()` | `para_id`, `image_path`, `width` (optional, EMU), `height` (optional, EMU), `alt_text` | Copies image to word/media/, adds relationship, adds content type, inserts `w:drawing` with `wp:inline`. Wrapped in `w:ins`. |
| `edit_header_footer` | `edit_header_footer()` | `section_index`, `location` ("header"/"footer"), `type` ("default"/"first"/"even"), `text` | Tracked delete of old content + tracked insert of new text |
| `add_endnote` | `add_endnote()` | `para_id`, `text` | Mirrors add_footnote but targets endnotes.xml. Creates endnote element + superscript reference in body. |
| `validate_endnotes` | `validate_endnotes()` | — | Cross-ref endnote IDs between body references and endnotes.xml |

**Test file:** `tests/test_content.py`

### Phase 5: Document Structure (4 new tools)

| Tool | Method | Args | Behavior |
|------|--------|------|----------|
| `add_section_break` | `add_section_break()` | `para_id`, `break_type` ("nextPage"/"continuous"/"evenPage"/"oddPage") | Insert `w:sectPr` inside the target paragraph's `w:pPr` (OOXML requires sectPr as child of pPr, not a sibling of w:p) |
| `set_section_properties` | `set_section_properties()` | `section_index`, `width`, `height`, `orientation`, `margins` (dict) | Modify `w:pgSz` and `w:pgMar` in section's `w:sectPr` |
| `add_page_break` | `add_page_break()` | `para_id` | Insert `w:br w:type="page"` in a new run after target paragraph |
| `add_cross_reference` | `add_cross_reference()` | `para_id`, `target_para_id`, `text` (display text) | Add bookmark at target if not present, insert `w:hyperlink w:anchor` at source with display text |

**Test file:** `tests/test_structure.py`

### Phase 6: Protection, Properties Write & Merge (3 new tools)

| Tool | Method | Args | Behavior |
|------|--------|------|----------|
| `set_document_protection` | `set_document_protection()` | `mode` ("trackedChanges"/"readOnly"/"comments"/"forms"), `password` (optional) | Add/update `w:documentProtection` in settings.xml with enforcement and optional SHA-512 hash |
| `set_properties` | `set_properties()` | `title`, `creator`, `subject`, `description` (all optional) | Update `dc:title`, `dc:creator`, `dc:subject`, `dc:description` in core.xml |
| `merge_documents` | `merge_documents()` | `source_path` | See Merge Sub-Spec below. Placed last because it depends on all other features being stable (remaps paraIds, rIds, footnotes, endnotes, comments, numbering, styles, images). |

**Test file:** `tests/test_protection.py`

## Merge Sub-Spec (`merge_documents`)

The most complex operation. Steps in order:

1. **Open source .docx** into a second temporary directory, parse all XML parts
2. **Remap paraIds** — collect all paraIds from target doc, generate new unique IDs for every paraId in source doc, apply via search-replace across all source XML parts
3. **Remap relationship IDs** — source rIds (rId1, rId2...) may collide with target. Offset all source rIds by max(target rIds) + 1, update all references in source XML
4. **Copy media files** — copy source `word/media/*` to target, renaming if filename collisions exist, update source rels accordingly
5. **Merge footnotes** — append source footnote elements to target footnotes.xml, remap footnote IDs to avoid collision, update body references
6. **Merge endnotes** — same pattern as footnotes
7. **Merge comments** — append source comments, remap comment IDs, update comment range markers in body
8. **Merge numbering** — append source abstractNum/num entries with remapped IDs
9. **Merge styles** — append source styles that don't exist in target (by styleId). Skip duplicates.
10. **Merge content types** — add any new extensions/overrides from source `[Content_Types].xml`
11. **Append body content** — insert all source `w:body` children (excluding final `w:sectPr`) after target's last paragraph but before target's final `w:sectPr`
12. **Merge relationships** — append remapped source relationships to target `document.xml.rels`

**Error handling:** If source .docx can't be opened, return error. If paraId collision can't be resolved (exhausted ID space), return error. Always clean up source temp directory.

**Test approach:** Create two minimal test .docx files with overlapping paraIds, rIds, footnotes, and styles. Verify merge produces valid document with no collisions. Verify audit_document passes on merged result.

## Additional Namespaces

Add to `base.py` namespace constants:

- `WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` — required for `wp:inline`, `wp:extent` in image insertion
- `DC = "http://purl.org/dc/elements/1.1/"` — required for `dc:title`, `dc:creator` in core properties
- `DCTERMS = "http://purl.org/dc/terms/"` — required for `dcterms:created`, `dcterms:modified`
- `CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"` — core properties namespace

## Infrastructure Notes

- **Numbering part bootstrap:** If numbering.xml doesn't exist in the source .docx (not all documents have it), `add_list` must create it from scratch: create the XML file, add the content type override in `[Content_Types].xml`, and add the relationship in `document.xml.rels`
- **Section enumeration:** `get_headers_footers` and `edit_header_footer` need section infrastructure (finding `w:sectPr` in body pPr elements + the final body sectPr). This is shared infrastructure extracted into a `_get_sections()` helper in `base.py`
- **Password hashing:** `set_document_protection` uses SHA-512 with salt, iterated 100,000 times, per ECMA-376 4th edition. The hash, salt, spin count, and algorithm name are stored as attributes on `w:documentProtection`

## Audit Extension

`audit_document()` grows with each phase:

- Phase 1: Report table count, image count, endnote count in summary
- Phase 3: Validate table structure (consistent col count per row, non-empty tblGrid)
- Phase 4: Validate image relationships (rId exists in rels, file exists in media/)
- Phase 4: Validate endnote cross-references
- Phase 5: Validate section break consistency
- Phase 6: Report protection status

## Test Strategy

- **TDD:** Write failing test first, implement minimum code to pass, refactor
- **100% coverage:** Enforced by `fail_under = 100` in pyproject.toml — every phase must maintain this
- **Fixture growth:** conftest.py `test_docx` fixture grows per phase to include necessary XML parts
- **Separate test files per phase:** Keeps test files focused and manageable
- **Each test class:** One tool, multiple scenarios (happy path, edge cases, error cases)

## Version Bump Plan

- Phase 0 (refactor): no version bump (internal only)
- Phase 1-2: v0.2.0 (reading + track changes complete)
- Phase 3-4: v0.3.0 (tables + content creation)
- Phase 5-6: v0.4.0 (structure + protection)

## Dependencies

No new dependencies. All features use lxml (already required) for XML manipulation and stdlib zipfile/shutil for file operations. Image insertion uses only stdlib (no Pillow — dimensions passed by caller).

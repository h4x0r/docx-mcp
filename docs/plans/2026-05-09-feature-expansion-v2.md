# Feature Expansion v2 Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Expand docx-mcp from 45 to ~100 tools across 8 phased releases, closing che-word-mcp gaps and opening new territory in legal automation, litigation support, and document forensics.

**Architecture:** Each feature domain gets its own mixin file in `docx_mcp/document/`. New mixins are composed into `DocxDocument` via `__init__.py`. All server tools registered in `docx_mcp/server.py`. All tests in `tests/` following the existing pytest + 100% coverage contract.

**Tech Stack:** Python 3.10+, lxml, FastMCP, pytest, pytest-cov, mistune, optional: presidio, spacy, tesseract (pytesseract), latex2mathml

**TDD contract:** Every task: RED commit (failing tests) then GREEN commit (passing implementation). Run `pytest --tb=short -q` after each step. Coverage must stay at 100% (`pytest --cov=docx_mcp --cov-fail-under=100`).

---

## Phase 0: scrub_pii Option B — regex fallback (v0.4.1)

**Goal:** `scrub_pii` works out-of-the-box with no optional deps using regex patterns. Presidio upgrades accuracy when installed.

**New file:** `docx_mcp/document/pii_regex.py`
**Modify:** `docx_mcp/document/pii.py`
**Test:** `tests/test_pii.py` (extend existing)

### Task 0.1: Regex PII engine

**Step 1: Write failing tests**
```python
# tests/test_pii.py — add to existing file
class TestRegexFallback:
    def test_scrub_email_no_presidio(self, tmp_path, monkeypatch):
        """scrub_pii works with regex even when presidio is unavailable."""
        import sys
        monkeypatch.setitem(sys.modules, "presidio_analyzer", None)
        # build doc with email in a run
        doc = _make_doc_with_text("contact us at alice@example.com for details")
        result = doc.scrub_pii(entities=["EMAIL_ADDRESS"])
        assert result["redacted_count"] >= 1

    def test_regex_patterns_covered(self):
        from docx_mcp.document.pii_regex import find_pii_spans
        text = "Email: foo@bar.com, Phone: +1-555-867-5309, SSN: 123-45-6789"
        spans = find_pii_spans(text)
        types = {s["type"] for s in spans}
        assert "EMAIL_ADDRESS" in types
        assert "PHONE_NUMBER" in types
        assert "US_SSN" in types

    def test_regex_credit_card(self):
        from docx_mcp.document.pii_regex import find_pii_spans
        spans = find_pii_spans("Card: 4111 1111 1111 1111")
        assert any(s["type"] == "CREDIT_CARD" for s in spans)

    def test_presidio_takes_precedence_when_available(self, tmp_path):
        """When presidio IS installed, it is still used (higher recall)."""
        doc = _make_doc_with_text("John Smith lives at 42 Main St")
        result = doc.scrub_pii()  # should not raise
        assert isinstance(result["redacted_count"], int)
```

Run: `pytest tests/test_pii.py::TestRegexFallback -v`
Expected: FAIL — `pii_regex` module not found.

**Step 2: Create `docx_mcp/document/pii_regex.py`**
```python
"""Regex-based PII detection — zero-dependency fallback for scrub_pii."""
from __future__ import annotations
import re

_PATTERNS: list[tuple[str, str]] = [
    ("EMAIL_ADDRESS",  r"\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b"),
    ("PHONE_NUMBER",   r"\b(\+?1[\s.\-]?)?(\(?\d{3}\)?[\s.\-]?)?\d{3}[\s.\-]?\d{4}\b"),
    ("US_SSN",         r"\b\d{3}-\d{2}-\d{4}\b"),
    ("CREDIT_CARD",    r"\b(?:\d[ \-]?){13,16}\b"),
    ("IP_ADDRESS",     r"\b(?:\d{1,3}\.){3}\d{1,3}\b"),
    ("US_ITIN",        r"\b9\d{2}-\d{2}-\d{4}\b"),
]

def find_pii_spans(text: str, entities: list[str] | None = None) -> list[dict]:
    """Return list of {type, start, end, text} dicts for PII found in text."""
    results: list[dict] = []
    for entity_type, pattern in _PATTERNS:
        if entities and entity_type not in entities:
            continue
        for m in re.finditer(pattern, text):
            results.append({
                "type": entity_type,
                "start": m.start(),
                "end": m.end(),
                "text": m.group(),
            })
    # sort by start, remove overlaps (keep longer match)
    results.sort(key=lambda x: (x["start"], -(x["end"] - x["start"])))
    deduped: list[dict] = []
    last_end = -1
    for span in results:
        if span["start"] >= last_end:
            deduped.append(span)
            last_end = span["end"]
    return deduped
```

**Step 3: Modify `pii.py` — use regex when presidio unavailable**

In `_get_analyzer()`, instead of raising `ImportError`, return `None` when unavailable.
In `PiiMixin.scrub_pii()`, fall back to `find_pii_spans` when analyzer is `None`.

Key change to `_get_analyzer`:
```python
def _get_analyzer():
    global _analyzer
    if _analyzer is None:
        try:
            from presidio_analyzer import AnalyzerEngine
            _analyzer = AnalyzerEngine()
        except ImportError:
            _analyzer = "regex"  # sentinel: use regex fallback
    return _analyzer
```

In `scrub_pii` method, branch on `_analyzer == "regex"`:
```python
analyzer = _get_analyzer()
if analyzer == "regex":
    from .pii_regex import find_pii_spans
    # use find_pii_spans() for span detection
    ...
else:
    # existing presidio path
    ...
```

**Step 4: Run and confirm GREEN**
`pytest tests/test_pii.py -v`

**Step 5: RED commit, GREEN commit**
```bash
# RED was committed before implementation
git add docx_mcp/document/pii_regex.py docx_mcp/document/pii.py tests/test_pii.py
git commit -m "GREEN: scrub_pii regex fallback — works without presidio/spaCy install"
```

**Step 6: Bump version to 0.4.1, push, tag**

---

## Phase 1: Tier C Foundation (v0.5.0)

**Goal:** Stable error taxonomy, raw XML part access, XPath query escape hatch. These primitives make every subsequent phase easier.

**New files:**
- `docx_mcp/document/errors.py` — error code enum + DocxMcpError
- `docx_mcp/document/rawparts.py` — RawPartsMixin
- `docx_mcp/document/query.py` — XPathMixin

### Task 1.1: Error taxonomy

**New file: `docx_mcp/document/errors.py`**
```python
"""Stable error codes for docx-mcp tools."""
from __future__ import annotations
from enum import Enum

class ErrCode(str, Enum):
    STYLE_NOT_FOUND       = "STYLE_NOT_FOUND"
    PARA_NOT_FOUND        = "PARA_NOT_FOUND"
    BOOKMARK_NOT_FOUND    = "BOOKMARK_NOT_FOUND"
    BOOKMARK_DANGLING     = "BOOKMARK_DANGLING"
    PART_NOT_FOUND        = "PART_NOT_FOUND"
    INVALID_RELATIONSHIP  = "INVALID_REL"
    NUMBERING_ORPHAN      = "NUMBERING_ORPHAN"
    OOXML_INVALID         = "OOXML_INVALID"
    PII_DEPS_MISSING      = "PII_DEPS_MISSING"
    NO_OPEN_DOCUMENT      = "NO_OPEN_DOCUMENT"
    XPATH_ERROR           = "XPATH_ERROR"

class DocxMcpError(Exception):
    def __init__(self, code: ErrCode, message: str, hint: str = ""):
        self.code = code
        self.hint = hint
        super().__init__(message)

    def to_dict(self) -> dict:
        return {"error": self.code.value, "message": str(self), "hint": self.hint}
```

**Test: `tests/test_errors.py`**
```python
from docx_mcp.document.errors import DocxMcpError, ErrCode

def test_error_to_dict():
    e = DocxMcpError(ErrCode.STYLE_NOT_FOUND, "Style 'Foo' not found",
                     hint="Use get_styles() to list available styles")
    d = e.to_dict()
    assert d["error"] == "STYLE_NOT_FOUND"
    assert "hint" in d

def test_error_is_exception():
    with pytest.raises(DocxMcpError) as exc_info:
        raise DocxMcpError(ErrCode.NO_OPEN_DOCUMENT, "No doc open")
    assert exc_info.value.code == ErrCode.NO_OPEN_DOCUMENT
```

### Task 1.2: Raw part read/write

**New file: `docx_mcp/document/rawparts.py`**
```python
"""Raw XML part read/write — power-user escape hatch."""
from __future__ import annotations
from lxml import etree
from .errors import DocxMcpError, ErrCode

class RawPartsMixin:
    def read_part(self, part_path: str) -> dict:
        """Return raw XML of any part (e.g. 'word/document.xml')."""
        tree = self._tree(part_path)
        if tree is None:
            raise DocxMcpError(ErrCode.PART_NOT_FOUND, f"Part not found: {part_path}",
                               hint="Use list_parts() to see available parts.")
        return {"part": part_path,
                "xml": etree.tostring(tree, pretty_print=True).decode()}

    def write_part(self, part_path: str, xml: str) -> dict:
        """Replace a part's XML. Validates well-formedness before writing."""
        try:
            tree = etree.fromstring(xml.encode())
        except etree.XMLSyntaxError as e:
            raise DocxMcpError(ErrCode.OOXML_INVALID, f"Invalid XML: {e}") from e
        self._zin._parts[part_path] = etree.tostring(tree)  # write to in-memory zip
        self._dirty.add(part_path)
        return {"part": part_path, "bytes_written": len(etree.tostring(tree))}

    def list_parts(self) -> list[str]:
        """List all parts (file paths) in the DOCX zip."""
        return sorted(self._zip.namelist())
```

**Test: `tests/test_rawparts.py`**
```python
def test_read_part_document_xml(tmp_path):
    doc = _make_minimal_doc(tmp_path)
    result = doc.read_part("word/document.xml")
    assert "<w:document" in result["xml"]

def test_read_part_not_found_raises(tmp_path):
    doc = _make_minimal_doc(tmp_path)
    with pytest.raises(DocxMcpError) as exc_info:
        doc.read_part("word/nonexistent.xml")
    assert exc_info.value.code == ErrCode.PART_NOT_FOUND

def test_write_part_invalid_xml_raises(tmp_path):
    doc = _make_minimal_doc(tmp_path)
    with pytest.raises(DocxMcpError) as exc_info:
        doc.write_part("word/document.xml", "<unclosed>")
    assert exc_info.value.code == ErrCode.OOXML_INVALID

def test_list_parts_includes_document(tmp_path):
    doc = _make_minimal_doc(tmp_path)
    parts = doc.list_parts()
    assert "word/document.xml" in parts
```

### Task 1.3: XPath query

**New file: `docx_mcp/document/query.py`**
```python
"""XPath query escape hatch."""
from __future__ import annotations
from lxml import etree
from .base import W, W14
from .errors import DocxMcpError, ErrCode

_NS = {
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
}

class XPathMixin:
    def xpath_query(self, xpath: str, part: str = "word/document.xml") -> dict:
        """Run XPath against any part. Returns matching elements as XML snippets.

        Namespaces pre-bound: w, w14, r, wp, a.
        Example: xpath_query("//w:p[w:pPr/w:pStyle/@w:val='Heading1']")
        """
        tree = self._require(part)
        try:
            matches = tree.xpath(xpath, namespaces=_NS)
        except etree.XPathError as e:
            raise DocxMcpError(ErrCode.XPATH_ERROR, f"XPath error: {e}",
                               hint="Check namespace prefixes: w, w14, r, wp, a") from e
        results = []
        for m in matches[:50]:  # cap at 50 to avoid flooding context
            if isinstance(m, etree._Element):
                results.append(etree.tostring(m, pretty_print=True).decode())
            else:
                results.append(str(m))
        return {"xpath": xpath, "part": part, "count": len(matches), "results": results}
```

**Test: `tests/test_query.py`**
```python
def test_xpath_finds_paragraphs(tmp_path):
    doc = _make_doc_with_heading(tmp_path, "Test Heading", level=1)
    result = doc.xpath_query("//w:p")
    assert result["count"] >= 1

def test_xpath_finds_by_style(tmp_path):
    doc = _make_doc_with_heading(tmp_path, "My Heading", level=1)
    result = doc.xpath_query("//w:p[w:pPr/w:pStyle/@w:val='Heading1']")
    assert result["count"] >= 1

def test_xpath_invalid_expression_raises(tmp_path):
    doc = _make_minimal_doc(tmp_path)
    with pytest.raises(DocxMcpError) as exc_info:
        doc.xpath_query("//[invalid")
    assert exc_info.value.code == ErrCode.XPATH_ERROR

def test_xpath_on_styles_part(tmp_path):
    doc = _make_minimal_doc(tmp_path)
    result = doc.xpath_query("//w:style", part="word/styles.xml")
    assert result["count"] >= 1
```

**Server tools to add (server.py):**
```python
@mcp.tool()
def read_part(part_path: str) -> str:
    """Read raw XML of any DOCX part (e.g. 'word/document.xml')."""
    return _js(_require_doc().read_part(part_path))

@mcp.tool()
def write_part(part_path: str, xml: str) -> str:
    """Replace a DOCX part with new XML. Validates well-formedness first."""
    return _js(_require_doc().write_part(part_path, xml))

@mcp.tool()
def list_parts() -> str:
    """List all parts (files) in the open DOCX zip."""
    return _js(_require_doc().list_parts())

@mcp.tool()
def xpath_query(xpath: str, part: str = "word/document.xml") -> str:
    """Run XPath against any DOCX part. Namespaces: w, w14, r, wp, a."""
    return _js(_require_doc().xpath_query(xpath, part))
```

**Register in `__init__.py`:**
Add `RawPartsMixin`, `XPathMixin` to `DocxDocument` bases.

---

## Phase 2: Hyperlinks, Bookmarks, Fields (v0.5.1)

**Goal:** Full CRUD for hyperlinks and bookmarks; field insertion (PAGE, SEQ, REF, IF, STYLEREF, update_fields).

**New files:**
- `docx_mcp/document/hyperlinks.py`
- `docx_mcp/document/bookmarks.py`
- `docx_mcp/document/fields.py`

### Task 2.1: Bookmark CRUD

**New file: `docx_mcp/document/bookmarks.py`**

Key OOXML: `<w:bookmarkStart w:id="1" w:name="MyBookmark"/>` ... `<w:bookmarkEnd w:id="1"/>`

```python
class BookmarksMixin:
    def list_bookmarks(self) -> list[dict]:
        """Return all bookmarks: {id, name, para_id}."""
        ...

    def add_bookmark(self, para_id: str, name: str) -> dict:
        """Wrap paragraph in bookmark start/end markers."""
        ...

    def remove_bookmark(self, name: str) -> dict:
        """Remove bookmarkStart+End pair by name."""
        ...

    def get_bookmarked_text(self, name: str) -> dict:
        """Return text content between bookmark markers."""
        ...
```

**Tests:**
```python
class TestBookmarks:
    def test_list_bookmarks_empty(self, tmp_path): ...
    def test_add_bookmark(self, tmp_path): ...
    def test_add_duplicate_name_raises(self, tmp_path): ...
    def test_remove_bookmark(self, tmp_path): ...
    def test_get_bookmarked_text(self, tmp_path): ...
    def test_bookmark_id_uniqueness(self, tmp_path): ...
```

### Task 2.2: Hyperlink CRUD

**New file: `docx_mcp/document/hyperlinks.py`**

Key OOXML: `<w:hyperlink r:id="rId5">` for external (relationship-based), `<w:hyperlink w:anchor="BookmarkName">` for internal.

```python
class HyperlinksMixin:
    def list_hyperlinks(self) -> list[dict]:
        """Return all hyperlinks: {id, url_or_anchor, text, para_id}."""
        ...

    def add_hyperlink(self, para_id: str, text: str, url: str) -> dict:
        """Insert external hyperlink after paragraph's last run.
        Creates r:id relationship in document.xml.rels."""
        ...

    def add_internal_link(self, para_id: str, text: str, bookmark: str) -> dict:
        """Insert w:hyperlink w:anchor pointing to bookmark."""
        ...

    def remove_hyperlink(self, hyperlink_id: str) -> dict:
        """Unwrap hyperlink (keep text runs, remove w:hyperlink wrapper)."""
        ...

    def update_hyperlink(self, hyperlink_id: str, url: str) -> dict:
        """Update the target URL in the relationship."""
        ...
```

**Tests:**
```python
class TestHyperlinks:
    def test_add_external_hyperlink(self, tmp_path): ...
    def test_hyperlink_creates_relationship(self, tmp_path): ...
    def test_add_internal_link_to_bookmark(self, tmp_path): ...
    def test_remove_hyperlink_preserves_text(self, tmp_path): ...
    def test_update_hyperlink_url(self, tmp_path): ...
    def test_list_hyperlinks(self, tmp_path): ...
```

### Task 2.3: Field insertion

**New file: `docx_mcp/document/fields.py`**

Key OOXML for complex fields:
```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>1</w:t></w:r>  <!-- cached result -->
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

```python
class FieldsMixin:
    def add_field(self, para_id: str, field_code: str,
                  cached_value: str = "") -> dict:
        """Insert a Word field at end of paragraph.

        Common field_code values:
          "PAGE"
          "NUMPAGES"
          "DATE \\@ \"MMMM d, yyyy\""
          "SEQ Figure \\* ARABIC"
          "REF MyBookmark \\h"
          "STYLEREF \"Heading 1\" \\n"
          'IF { MERGEFIELD Status } = "Active" "Yes" "No"'
        """
        ...

    def update_fields(self) -> dict:
        """Mark all fields w:dirty='true' so Word recalculates on open.
        Deterministic fields (SEQ) are recalculated immediately."""
        ...

    def list_fields(self) -> list[dict]:
        """Return all fields: {code, cached_value, para_id, location}."""
        ...
```

**Tests:**
```python
class TestFields:
    def test_add_page_field(self, tmp_path): ...
    def test_add_seq_field(self, tmp_path): ...
    def test_add_ref_field(self, tmp_path): ...
    def test_field_structure_begin_separate_end(self, tmp_path): ...
    def test_update_fields_sets_dirty(self, tmp_path): ...
    def test_list_fields(self, tmp_path): ...
    def test_add_if_field(self, tmp_path): ...
```

**Server tools:** `add_field`, `update_fields`, `list_fields`, `list_bookmarks`, `add_bookmark`, `remove_bookmark`, `add_hyperlink`, `add_internal_link`, `remove_hyperlink`, `update_hyperlink`, `list_hyperlinks`

---

## Phase 3: Content Controls + ToC (v0.5.2)

**Goal:** SDT content controls (checkbox, dropdown, date, text) and Table of Contents generation.

**New files:**
- `docx_mcp/document/contentcontrols.py`
- `docx_mcp/document/toc.py`

### Task 3.1: Content controls

**New file: `docx_mcp/document/contentcontrols.py`**

Key OOXML:
```xml
<!-- Checkbox (w14:checkbox) -->
<w:sdt>
  <w:sdtPr>
    <w:tag w:val="myTag"/>
    <w:alias w:val="My Label"/>
    <w14:checkbox><w14:checked w14:val="0"/></w14:checkbox>
  </w:sdtPr>
  <w:sdtContent><w:p>...</w:p></w:sdtContent>
</w:sdt>

<!-- Dropdown -->
<w:sdt>
  <w:sdtPr>
    <w:dropDownList>
      <w:listItem w:displayText="Option A" w:value="A"/>
    </w:dropDownList>
  </w:sdtPr>
  <w:sdtContent>...</w:sdtContent>
</w:sdt>
```

```python
class ContentControlsMixin:
    def add_content_control(self, para_id: str, tag: str,
                             control_type: str,  # "text"|"checkbox"|"dropdown"|"date"
                             label: str = "",
                             options: list[str] | None = None,
                             default: str = "") -> dict: ...

    def get_content_controls(self) -> list[dict]: ...

    def set_content_control_value(self, tag: str, value: str) -> dict: ...

    def lock_content_control(self, tag: str,
                              lock: str = "sdtLocked") -> dict: ...
```

**Tests:**
```python
class TestContentControls:
    def test_add_checkbox_control(self, tmp_path): ...
    def test_add_dropdown_control(self, tmp_path): ...
    def test_add_date_picker_control(self, tmp_path): ...
    def test_add_text_control(self, tmp_path): ...
    def test_set_control_value(self, tmp_path): ...
    def test_list_content_controls(self, tmp_path): ...
    def test_lock_content_control(self, tmp_path): ...
    def test_duplicate_tag_raises(self, tmp_path): ...
```

### Task 3.2: Table of Contents

**New file: `docx_mcp/document/toc.py`**

A ToC is an SDT wrapping a `TOC \o "1-3" \h \z` field followed by cached TOC entries. Each entry is a paragraph styled `TOC 1`/`TOC 2`/`TOC 3` with a tab and `PAGEREF` field.

```python
class TocMixin:
    def generate_toc(self, max_level: int = 3,
                      title: str = "Table of Contents") -> dict:
        """Insert a ToC field at start of document body.
        Generates cached entries from current headings.
        Requires Word/LibreOffice to recalculate page numbers."""
        ...

    def generate_list_of_figures(self) -> dict:
        """Insert TOC field for SEQ Figure captions."""
        ...

    def generate_list_of_tables(self) -> dict:
        """Insert TOC field for SEQ Table captions."""
        ...

    def update_toc(self) -> dict:
        """Regenerate TOC entries from current headings (page nums marked dirty)."""
        ...
```

**Tests:**
```python
class TestToc:
    def test_generate_toc_creates_field(self, tmp_path): ...
    def test_toc_entries_match_headings(self, tmp_path): ...
    def test_toc_max_level_filtering(self, tmp_path): ...
    def test_update_toc_reflects_new_headings(self, tmp_path): ...
    def test_generate_lof_requires_seq_captions(self, tmp_path): ...
```

**Server tools:** `add_content_control`, `get_content_controls`, `set_content_control_value`, `lock_content_control`, `generate_toc`, `update_toc`, `generate_list_of_figures`, `generate_list_of_tables`

---

## Phase 4: Floating Images + Multilevel Lists + Table Improvements (v0.5.3)

### Task 4.1: Floating image

**Modify: `docx_mcp/document/images.py`**

Add `insert_floating_image(path, width_cm, height_cm, h_pos, v_pos, wrap)`.

Key OOXML: `<wp:anchor>` vs current `<wp:inline>`. Anchor needs `wp:positionH/V`, `wp:wrapSquare` or `wp:wrapTopAndBottom`, `distT/B/L/R` attributes.

```python
def insert_floating_image(self, para_id: str, image_path: str,
                           width_cm: float, height_cm: float,
                           h_pos: float = 0.0, v_pos: float = 0.0,
                           wrap: str = "square") -> dict: ...
```

**Tests:**
```python
class TestFloatingImage:
    def test_insert_floating_creates_anchor(self, tmp_path, sample_png): ...
    def test_wrap_square_attribute(self, tmp_path, sample_png): ...
    def test_wrap_topbottom(self, tmp_path, sample_png): ...
    def test_position_set_correctly(self, tmp_path, sample_png): ...
```

### Task 4.2: Multilevel list

**Modify: `docx_mcp/document/lists.py`**

Add `create_multilevel_list(name, levels)` and `apply_numbering_to_headings(abstract_num_id)`.

Each level dict: `{num_fmt, lvl_text, indent, hanging, style}`.

```python
def create_multilevel_list(self, name: str,
                            levels: list[dict]) -> dict:
    """Create abstractNum + num entries in numbering.xml."""
    ...

def restart_numbering(self, para_id: str, level: int, start: int = 1) -> dict:
    """Add lvlOverride with startOverride for this paragraph."""
    ...

def suppress_numbering(self, para_id: str) -> dict:
    """Set w:numPr/w:numId val='0' to suppress list numbering."""
    ...
```

**Tests:**
```python
class TestMultilevelList:
    def test_create_multilevel_list_3_levels(self, tmp_path): ...
    def test_numbering_xml_abstract_num_created(self, tmp_path): ...
    def test_restart_numbering(self, tmp_path): ...
    def test_suppress_numbering(self, tmp_path): ...
    def test_heading_numbering_binding(self, tmp_path): ...
```

### Task 4.3: Table cell merge/split + header row

**Modify: `docx_mcp/document/tables.py`**

```python
def merge_cells(self, table_index: int,
                start_row: int, start_col: int,
                end_row: int, end_col: int) -> dict:
    """Horizontal: w:gridSpan. Vertical: w:vMerge."""
    ...

def set_header_row(self, table_index: int) -> dict:
    """Set w:tblHeader on first row — repeats across page breaks."""
    ...

def set_column_widths(self, table_index: int,
                      widths_cm: list[float]) -> dict:
    """Set w:tcW for each column in w:tblGrid."""
    ...

def csv_to_table(self, para_id: str, csv_text: str,
                 header_row: bool = True) -> dict:
    """Insert table from CSV string."""
    ...

def table_to_csv(self, table_index: int) -> dict:
    """Export table to CSV string."""
    ...
```

**Tests:**
```python
class TestTableImprovements:
    def test_merge_cells_horizontal(self, tmp_path): ...
    def test_merge_cells_vertical(self, tmp_path): ...
    def test_set_header_row(self, tmp_path): ...
    def test_set_column_widths(self, tmp_path): ...
    def test_csv_to_table_roundtrip(self, tmp_path): ...
```

**Server tools:** `insert_floating_image`, `create_multilevel_list`, `restart_numbering`, `suppress_numbering`, `merge_cells`, `set_header_row`, `set_column_widths`, `csv_to_table`, `table_to_csv`

---

## Phase 5: Legal Automation (v0.6.0)

**Goal:** fill_template (data-driven contract generation), bates_number, redact_text (true XML deletion + black box), generate_privilege_log.

**New files:**
- `docx_mcp/document/template.py`
- `docx_mcp/document/litigation.py`

### Task 5.1: fill_template

**New file: `docx_mcp/document/template.py`**

How it works:
1. Walk all `<w:sdt>` tags in the document
2. For each SDT, read `w:tag/@w:val` to get the key name
3. Look up key in provided `data: dict[str, str]`
4. Replace `w:sdtContent` with a plain run containing the value
5. Support repeating sections: `w:sdt[@w14:repeatingSectionItem]` — clone N times for list values

```python
class TemplateMixin:
    def fill_template(self, data: dict[str, str | list[str]],
                       remove_empty: bool = False) -> dict:
        """Fill all SDT content controls from data dict.

        Args:
            data: {"CLIENT_NAME": "Acme Corp", "DATE": "2026-06-01"}
                  For repeating sections, pass list: {"PARTIES": ["Alice", "Bob"]}
            remove_empty: if True, remove SDTs with no matching key

        Returns: {filled: N, unfilled: [...tags]}
        """
        ...

    def list_template_fields(self) -> list[dict]:
        """List all SDT tags/aliases in the document — the template schema."""
        ...

    def validate_template_data(self, data: dict) -> dict:
        """Check data dict covers all required template fields."""
        ...
```

**Tests:**
```python
class TestFillTemplate:
    def test_fill_text_control(self, tmp_path): ...
    def test_fill_multiple_fields(self, tmp_path): ...
    def test_unfilled_fields_reported(self, tmp_path): ...
    def test_remove_empty_controls(self, tmp_path): ...
    def test_list_template_fields(self, tmp_path): ...
    def test_validate_missing_fields(self, tmp_path): ...
    def test_repeating_section_list(self, tmp_path): ...
```

### Task 5.2: Bates numbering

**New file: `docx_mcp/document/litigation.py`** (first tool in it)

Bates numbering works by adding a footer field/run to every section with sequential numbering. Since we don't paginate, we embed the stamp text as a field in the footer with a SEQ counter.

Actually, the canonical approach in OOXML:
- For each section in the doc, add/update its footer
- Add a run with `prefix + zero-padded number` computed from SEQ counter
- Alternatively: insert the Bates stamp as a `DOCPROPERTY` or `SET` field

Simpler reliable approach: add a footer run with the Bates prefix + `SEQ BatesNum \* MERGEFORMAT \# "000000"` field. Word will calculate page-sequential values. We also emit a separate manifest.

```python
class LitigationMixin:
    def bates_number(self, prefix: str, start: int = 1,
                      digits: int = 6, position: str = "footer-right") -> dict:
        """Stamp Bates numbers on every section footer.

        Returns: {prefix, start, digits, sections_stamped}
        """
        ...

    def redact_text(self, pattern: str | None = None,
                    para_ids: list[str] | None = None,
                    exact_text: str | None = None,
                    reason: str = "") -> dict:
        """True redaction: delete text from XML, replace run with black rectangle.

        The underlying text is removed (not just highlighted) — a forensically
        clean redaction. Generates a redaction log entry for each redacted span.

        Args:
            pattern: regex to match text runs
            para_ids: limit search to these paragraphs
            exact_text: exact string to redact
            reason: exemption code for log (e.g. "Attorney-Client Privilege")

        Returns: {redacted_count, log: [{para_id, original_length, reason}]}
        """
        ...

    def generate_redaction_log(self, output_path: str = "") -> dict:
        """Write a DOCX table listing all redactions made this session.

        Columns: #, Location (para_id), Characters Removed, Reason, Reviewer, Date
        """
        ...

    def generate_privilege_log(self, output_path: str = "") -> dict:
        """Generate a privilege log DOCX from document metadata.

        Extracts: author, last-modified-by, creation date, tracked-change authors.
        Produces a table: Bates Range | Author | Recipients | Date | Subject | Basis
        """
        ...
```

**Tests:**
```python
class TestBatesNumbering:
    def test_bates_footer_contains_prefix(self, tmp_path): ...
    def test_bates_start_override(self, tmp_path): ...
    def test_bates_digits_padding(self, tmp_path): ...

class TestRedactText:
    def test_redact_by_exact_text(self, tmp_path): ...
    def test_redacted_text_removed_from_xml(self, tmp_path): ...
    def test_redaction_replaced_with_black_rect(self, tmp_path): ...
    def test_redaction_log_generated(self, tmp_path): ...
    def test_redact_by_regex(self, tmp_path): ...

class TestPrivilegeLog:
    def test_privilege_log_is_valid_docx(self, tmp_path): ...
    def test_privilege_log_columns_present(self, tmp_path): ...
    def test_redaction_log_docx_output(self, tmp_path): ...
```

**Server tools:** `fill_template`, `list_template_fields`, `validate_template_data`, `bates_number`, `redact_text`, `generate_redaction_log`, `generate_privilege_log`

---

## Phase 6: Equations + Charts (v0.7.0)

**Goal:** LaTeX→OMML equation insertion; native DOCX chart from data series.

**New files:**
- `docx_mcp/document/equations.py`
- `docx_mcp/document/charts.py`

**Optional dependency:** `latex2mathml` (pure Python, ~100KB) for LaTeX→MathML; custom XSLT for MathML→OMML (Microsoft provides MML2OMML.XSL, MIT-compatible redistribution).

### Task 6.1: Equations

```python
class EquationsMixin:
    def add_equation(self, para_id: str, latex: str) -> dict:
        """Insert an equation as Office Math (OMML).

        Pipeline: LaTeX → MathML (via latex2mathml) → OMML (via XSL transform)
        Fallback: if latex2mathml not installed, raise DocxMcpError with install hint.
        """
        ...

    def get_equations(self) -> list[dict]:
        """Return all equations as {para_id, omml_xml, latex_approx}."""
        ...
```

**Tests:**
```python
class TestEquations:
    def test_add_simple_equation(self, tmp_path): ...
    def test_omml_in_document_xml(self, tmp_path): ...
    def test_get_equations_roundtrip(self, tmp_path): ...
    def test_missing_dep_graceful_error(self, tmp_path, monkeypatch): ...
```

### Task 6.2: Native charts

DOCX charts are an embedded `charts/chartN.xml` part (DrawingML chart) referenced from `document.xml` via a relationship. The chart part contains data series inline (no Excel dependency in our case).

```python
class ChartsMixin:
    def insert_bar_chart(self, para_id: str, title: str,
                          series: list[dict],  # [{"name": "Q1", "values": [10,20,30]}]
                          categories: list[str],
                          width_cm: float = 14.0, height_cm: float = 9.0) -> dict:
        """Insert a native bar chart (no Excel required)."""
        ...

    def insert_line_chart(self, para_id: str, title: str,
                           series: list[dict], categories: list[str],
                           width_cm: float = 14.0, height_cm: float = 9.0) -> dict: ...

    def insert_pie_chart(self, para_id: str, title: str,
                          series: list[dict], categories: list[str]) -> dict: ...

    def update_chart_data(self, chart_id: str, series: list[dict]) -> dict:
        """Replace data series in an existing chart."""
        ...
```

**Tests:**
```python
class TestCharts:
    def test_insert_bar_chart_creates_part(self, tmp_path): ...
    def test_chart_relationship_in_document(self, tmp_path): ...
    def test_series_data_in_chart_xml(self, tmp_path): ...
    def test_insert_line_chart(self, tmp_path): ...
    def test_insert_pie_chart(self, tmp_path): ...
    def test_update_chart_data(self, tmp_path): ...
```

**Server tools:** `add_equation`, `get_equations`, `insert_bar_chart`, `insert_line_chart`, `insert_pie_chart`, `update_chart_data`

---

## Phase 7: Advanced Review Tools (v0.8.0)

**Goal:** merge_review_rounds (N reviewer copies → one consolidated), clause-aligned contract diff.

### Task 7.1: merge_review_rounds

```python
# docx_mcp/document/reviewmerge.py
class ReviewMergeMixin:
    def merge_review_rounds(self, reviewer_paths: list[str],
                             base_path: str | None = None) -> dict:
        """Merge tracked changes from N reviewer copies into the open document.

        Algorithm:
        1. For each reviewer doc, extract w:ins/w:del with author+date
        2. Deduplicate identical changes (same text + position)
        3. Merge non-conflicting changes into open doc
        4. Flag conflicts (same range, different content) for manual resolution

        Returns: {merged: N, conflicts: [...], output_path}
        """
        ...
```

### Task 7.2: Clause-aligned contract diff

```python
# docx_mcp/document/clausediff.py
class ClauseDiffMixin:
    def compare_contracts(self, other_path: str,
                           output_path: str = "",
                           align_by: str = "heading") -> dict:
        """Clause-aware diff: align by heading, then diff within each clause.

        Unlike compare_documents (which does LCS line-by-line), this:
        1. Extracts logical clauses (heading → sub-paragraphs) from both docs
        2. Matches clauses by heading text (fuzzy) across both docs
        3. Produces tracked-change output aligned at clause level
        4. Flags reordered, added, deleted, or renamed clauses

        Returns: {output_path, clauses_compared, clauses_changed, reordered}
        """
        ...
```

---

## Phase 8: Remaining Tier C (v0.9.0)

### Task 8.1: Session log + replay

```python
# docx_mcp/document/sessionlog.py
class SessionLogMixin:
    def get_session_log(self) -> list[dict]:
        """Return all operations performed this session as replayable JSON."""
        ...

    def export_session_script(self, output_path: str) -> dict:
        """Write session as a Python script using the MCP tool API."""
        ...
```

### Task 8.2: ECMA-376 schema validation

```python
# docx_mcp/document/schemavalidation.py
class SchemaValidationMixin:
    def validate_schema(self, strict: bool = False) -> dict:
        """Validate document.xml against ECMA-376 XSD.
        strict=True uses strict conformance; False uses transitional."""
        ...
```

### Task 8.3: Accessibility audit

```python
# docx_mcp/document/accessibility.py
class AccessibilityMixin:
    def accessibility_audit(self) -> dict:
        """WCAG 2.1 / Section 508 checks:
        - Images missing alt-text (wp:docPr/@descr empty)
        - Tables missing header row (w:tblHeader)
        - Heading sequence gaps (H1→H3)
        - Language not set (w:lang)
        - Color contrast (theme-based estimate)
        Returns: {pass: [...], fail: [...], warn: [...]}
        """
        ...

    def fix_alt_text(self, image_id: str, alt_text: str) -> dict: ...
```

---

## Phase 9: Image PII + Pentest Tools (v0.9.x — optional, high effort)

### Task 9.1: Image PII via OCR

**Optional dep:** `pytesseract` + `Pillow`

```python
# docx_mcp/document/imgpii.py
def scrub_image_pii(self, confidence_threshold: float = 0.7) -> dict:
    """OCR all embedded images, detect PII text regions, pixelate/blackbox."""
    ...
```

### Task 9.2: Pentest tools

**New file: `docx_mcp/document/pentest.py`**

```python
class PentestMixin:
    def add_finding(self, title: str, severity: str,
                     cvss_vector: str | None = None,
                     affected: str = "",
                     description: str = "",
                     evidence: str = "",
                     remediation: str = "") -> dict:
        """Insert a structured finding block with styled headings."""
        ...

    def compute_cvss(self, vector: str) -> dict:
        """Compute CVSS 3.1 or 4.0 base score from vector string."""
        ...

    def build_findings_summary_table(self) -> dict:
        """Auto-generate executive table from all finding blocks."""
        ...

    def embed_command_evidence(self, para_id: str, command_output: str,
                                caption: str = "") -> dict:
        """Insert terminal output in Courier New / fixed-width style block
        with smart-quote conversion DISABLED."""
        ...

    def hash_evidence_chain(self) -> dict:
        """SHA-256 all embedded images; embed hashes in their captions."""
        ...
```

---

## Execution Order

Run phases sequentially. Each phase:
1. Start gitsign credential cache: `gitsign credential-cache start`
2. Write failing tests → RED commit
3. Implement → GREEN commit
4. `pytest --cov=docx_mcp --cov-fail-under=100 -q`
5. Bump patch version → tag → push → verify CI green

| Phase | Version | New tools | Est. tests |
|-------|---------|-----------|------------|
| 0     | 0.4.1   | 0 (fixes) | +8         |
| 1     | 0.5.0   | 4         | +15        |
| 2     | 0.5.1   | 11        | +25        |
| 3     | 0.5.2   | 8         | +20        |
| 4     | 0.5.3   | 9         | +20        |
| 5     | 0.6.0   | 7         | +25        |
| 6     | 0.7.0   | 6         | +20        |
| 7     | 0.8.0   | 2         | +15        |
| 8     | 0.9.0   | 5         | +20        |
| 9     | 0.9.x   | 8         | +25        |
| **Σ** | **1.0** | **~60**   | **+193**   |

Total after all phases: ~105 tools, ~615 tests.

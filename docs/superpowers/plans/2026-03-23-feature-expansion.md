# docx-mcp Feature Expansion Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Expand docx-mcp from 18 to 45 tools with tables, formatting, images, properties, headers/footers, styles, lists, endnotes, sections, merge, cross-references, and protection — all with TDD and 100% coverage.

**Architecture:** Refactor monolithic `document.py` into a mixin-based package (`docx_mcp/document/`). Each feature domain gets its own mixin file. `DocxDocument` composes all mixins. Import path unchanged: `from docx_mcp.document import DocxDocument`.

**Tech Stack:** Python 3.10+, lxml, FastMCP, pytest, pytest-cov

**Spec:** `docs/superpowers/specs/2026-03-23-feature-expansion-design.md`

---

## Phase 0: Refactor document.py into package

### Task 0.1: Create package structure and base mixin

**Files:**
- Create: `docx_mcp/document/__init__.py`
- Create: `docx_mcp/document/base.py`
- Delete: `docx_mcp/document.py` (after migration)

- [ ] **Step 1: Create `docx_mcp/document/base.py`**

Move all imports, namespace constants, module-level helpers (`_now_iso`, `_preserve`), and the `DocxDocument` class skeleton with these methods into `base.py`:
- `__init__`, `open`, `close`, `save`
- `_tree`, `_require`, `_mark`, `_text`, `_find_para`, `_new_para_id`, `_next_markup_id`, `_next_comment_id`, `_make_run`
- All namespace constants: `W`, `W14`, `W15`, `R`, `V`, `A`, `CT`, `RELS`, `XML_SPACE`, `NSMAP`, `REL_TYPES`, `CT_TYPES`

Rename the class to `BaseMixin`.

```python
# docx_mcp/document/base.py
"""Base mixin: lifecycle, XML cache, namespace constants, shared helpers."""

from __future__ import annotations

import contextlib
import os
import random
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

# ── OOXML namespace constants ───────────────────────────────────────────────
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
W15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
V = "{urn:schemas-microsoft-com:vml}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
RELS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

REL_TYPES = {
    "comments": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "commentsExtended": "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    "footnotes": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
}

CT_TYPES = {
    "comments": "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    "commentsExtended": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
    ),
}


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _preserve(t_el: etree._Element, text: str) -> None:
    """Set text on a <w:t> or <w:delText> element with xml:space=preserve."""
    t_el.text = text
    t_el.set(XML_SPACE, "preserve")


class BaseMixin:
    """Lifecycle, XML cache, and shared helpers."""

    def __init__(self, path: str):
        self.source_path = Path(path).resolve()
        self.workdir: Path | None = None
        self._trees: dict[str, etree._Element] = {}
        self._modified: set[str] = set()

    # Copy ALL existing lifecycle methods (open, close, save) and
    # ALL private helpers (_tree, _require, _mark, _text, _find_para,
    # _new_para_id, _next_markup_id, _next_comment_id, _make_run,
    # _real_footnotes, _create_comments_part) from document.py verbatim.
    # ... (exact code from current document.py lines 77-919)
```

- [ ] **Step 2: Create reading mixin**

```python
# docx_mcp/document/reading.py
"""Reading mixin: headings, search, paragraph access, document info."""

from __future__ import annotations

import contextlib
import re

from lxml import etree

from .base import W, W14, R, A, RELS


class ReadingMixin:
    """Read-only document inspection methods."""

    def get_info(self) -> dict:
        # Copy from document.py lines 138-161

    def get_headings(self) -> list[dict]:
        # Copy from document.py lines 165-167

    def _find_headings(self, root: etree._Element) -> list[dict]:
        # Copy from document.py lines 169-190

    def search_text(self, query: str, *, regex: bool = False) -> list[dict]:
        # Copy from document.py lines 194-229

    def get_paragraph(self, para_id: str) -> dict:
        # Copy from document.py lines 233-249
```

- [ ] **Step 3: Create tracks mixin**

```python
# docx_mcp/document/tracks.py
"""Track changes mixin: insert, delete."""

from __future__ import annotations

from lxml import etree

from .base import W, W14, _now_iso, _preserve


class TracksMixin:
    """Insert/delete text with tracked changes markup."""

    def insert_text(self, para_id, text, *, position="end", author="Claude") -> dict:
        # Copy from document.py lines 402-454

    def delete_text(self, para_id, text, *, author="Claude") -> dict:
        # Copy from document.py lines 456-525
```

- [ ] **Step 4: Create comments mixin**

```python
# docx_mcp/document/comments.py
"""Comments mixin: get, add, reply."""

from __future__ import annotations

from lxml import etree

from .base import W, W14, W15, _now_iso, _preserve


class CommentsMixin:
    """Comment operations."""

    def get_comments(self) -> list[dict]:
        # Copy from document.py lines 529-541

    def add_comment(self, para_id, text, *, author="Claude") -> dict:
        # Copy from document.py lines 543-612

    def reply_to_comment(self, parent_id, text, *, author="Claude") -> dict:
        # Copy from document.py lines 614-669
```

- [ ] **Step 5: Create footnotes mixin**

```python
# docx_mcp/document/footnotes.py
"""Footnotes mixin: get, add, validate."""

from __future__ import annotations

from lxml import etree

from .base import W, W14, _preserve


class FootnotesMixin:
    """Footnote operations."""

    def get_footnotes(self) -> list[dict]:
        # Copy from document.py lines 253-265

    def add_footnote(self, para_id: str, text: str) -> dict:
        # Copy from document.py lines 267-319

    def validate_footnotes(self) -> dict:
        # Copy from document.py lines 321-346
```

- [ ] **Step 6: Create validation mixin**

```python
# docx_mcp/document/validation.py
"""Validation mixin: paraids, watermark, audit."""

from __future__ import annotations

from lxml import etree

from .base import W, W14, R, V, A, RELS


class ValidationMixin:
    """Structural validation and audit."""

    def validate_paraids(self) -> dict:
        # Copy from document.py lines 350-375

    def remove_watermark(self) -> dict:
        # Copy from document.py lines 379-398

    def audit(self) -> dict:
        # Copy from document.py lines 673-753
```

- [ ] **Step 7: Create `__init__.py` composing all mixins**

```python
# docx_mcp/document/__init__.py
"""DocxDocument: mixin composition and public API."""

from .base import (
    A,
    CT,
    CT_TYPES,
    NSMAP,
    R,
    RELS,
    REL_TYPES,
    V,
    W,
    W14,
    W15,
    XML_SPACE,
    BaseMixin,
    _now_iso,
    _preserve,
)
from .comments import CommentsMixin
from .footnotes import FootnotesMixin
from .reading import ReadingMixin
from .tracks import TracksMixin
from .validation import ValidationMixin


class DocxDocument(
    BaseMixin,
    ReadingMixin,
    TracksMixin,
    CommentsMixin,
    FootnotesMixin,
    ValidationMixin,
):
    """Word document editor with OOXML-level control."""

    pass


__all__ = [
    "DocxDocument",
    "W", "W14", "W15", "R", "V", "A", "CT", "RELS", "XML_SPACE",
    "NSMAP", "REL_TYPES", "CT_TYPES",
    "_now_iso", "_preserve",
]
```

- [ ] **Step 8: Delete old `docx_mcp/document.py`**

Remove the monolithic file now that the package replaces it.

- [ ] **Step 9: Run all tests — must pass identically**

Run: `python -m pytest tests/test_e2e.py -v`
Expected: 82 passed

- [ ] **Step 10: Verify 100% coverage**

Run: `python -m pytest tests/ --cov=docx_mcp --cov-report=term-missing --cov-fail-under=100`
Expected: 100% coverage, 82 passed

- [ ] **Step 11: Commit**

```bash
git add docx_mcp/document/ && git rm docx_mcp/document.py
git add tests/
git commit -m "refactor: extract document.py into mixin-based package

No behavior changes. All 82 tests pass. 100% coverage maintained.
Prepares for feature expansion with per-domain mixin files."
```

---

## Phase 1: Reading tools (6 new tools)

### Task 1.1: Extend test fixture with tables, styles, endnotes, properties, image, headers/footers

**Files:**
- Modify: `tests/conftest.py`

- [ ] **Step 1: Add new XML templates to conftest.py**

Add these new XML template strings and update existing ones:

```python
# Add to conftest.py - new XML templates

_STYLES_XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
  </w:style>
  <w:style w:type="character" w:styleId="FootnoteReference">
    <w:name w:val="footnote reference"/>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
  </w:style>
</w:styles>
"""

_ENDNOTES_XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:endnote w:type="separator" w:id="-1">
    <w:p w14:paraId="00000E11" w14:textId="77777777">
      <w:r><w:separator/></w:r>
    </w:p>
  </w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0">
    <w:p w14:paraId="00000E12" w14:textId="77777777">
      <w:r><w:continuationSeparator/></w:r>
    </w:p>
  </w:endnote>
  <w:endnote w:id="1">
    <w:p w14:paraId="00000E13" w14:textId="77777777">
      <w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>
      <w:r><w:t>Endnote reference material.</w:t></w:r>
    </w:p>
  </w:endnote>
</w:endnotes>
"""

_CORE_XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Test Document</dc:title>
  <dc:creator>Test Author</dc:creator>
  <dc:subject>Test Subject</dc:subject>
  <dc:description>Test Description</dc:description>
  <cp:lastModifiedBy>Test Editor</cp:lastModifiedBy>
  <cp:revision>3</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2025-06-15T12:00:00Z</dcterms:modified>
</cp:coreProperties>
"""

_SETTINGS_XML = """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:settings>
"""
```

Add a 2x3 table to `_DOCUMENT_XML` (before `</w:body>`):

```xml
    <w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>
      <w:tr w14:paraId="00000T01" w14:textId="77777777">
        <w:tc><w:p w14:paraId="00000T02" w14:textId="77777777"><w:r><w:t>Header A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p w14:paraId="00000T03" w14:textId="77777777"><w:r><w:t>Header B</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr w14:paraId="00000T04" w14:textId="77777777">
        <w:tc><w:p w14:paraId="00000T05" w14:textId="77777777"><w:r><w:t>Row 1 A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p w14:paraId="00000T06" w14:textId="77777777"><w:r><w:t>Row 1 B</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr w14:paraId="00000T07" w14:textId="77777777">
        <w:tc><w:p w14:paraId="00000T08" w14:textId="77777777"><w:r><w:t>Row 2 A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p w14:paraId="00000T09" w14:textId="77777777"><w:r><w:t>Row 2 B</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
```

Add a 1x1 PNG image (smallest valid PNG — 67 bytes) to the fixture and an image reference in document.xml. Add an endnote reference to paragraph 00000005.

Update `_build_fixture` to include the new files: `word/styles.xml`, `word/endnotes.xml`, `docProps/core.xml`, `word/settings.xml`, and `word/media/image1.png`.

Update `_CONTENT_TYPES` with overrides for styles, endnotes, core properties, settings.

Update `_DOC_RELS` with relationships for styles, endnotes, settings, and image.

Update paragraph count assertions in `TestOpen.test_open_returns_info` and `TestInfo.test_get_info` (table cells add paragraphs).

- [ ] **Step 2: Run tests to verify fixture changes don't break existing tests**

Run: `python -m pytest tests/test_e2e.py -v`
Expected: All existing tests pass (may need paragraph count adjustments for the table paragraphs)

- [ ] **Step 3: Commit fixture expansion**

```bash
git add tests/conftest.py tests/test_e2e.py
git commit -m "test: expand fixture with table, styles, endnotes, properties, image"
```

### Task 1.2: get_tables tool

**Files:**
- Create: `docx_mcp/document/tables.py`
- Modify: `docx_mcp/document/__init__.py` (add TablesMixin)
- Modify: `docx_mcp/server.py` (add get_tables tool)
- Create: `tests/test_reading.py`

- [ ] **Step 1: Write failing test**

```python
# tests/test_reading.py
"""Tests for Phase 1 read-only tools."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server


def _j(result: str) -> dict | list:
    return json.loads(result)


class TestGetTables:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_returns_tables(self):
        tables = _j(server.get_tables())
        assert len(tables) == 1
        t = tables[0]
        assert t["index"] == 0
        assert t["row_count"] == 3
        assert t["col_count"] == 2
        assert t["cells"][0] == ["Header A", "Header B"]
        assert t["cells"][1] == ["Row 1 A", "Row 1 B"]

    def test_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.get_tables()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_reading.py::TestGetTables -v`
Expected: FAIL — `server.get_tables` doesn't exist

- [ ] **Step 3: Implement tables mixin**

```python
# docx_mcp/document/tables.py
"""Tables mixin: read and write table operations."""

from __future__ import annotations

from lxml import etree

from .base import W, W14


class TablesMixin:
    """Table operations."""

    def get_tables(self) -> list[dict]:
        """Get all tables with their content."""
        doc = self._require("word/document.xml")
        tables = []
        for idx, tbl in enumerate(doc.iter(f"{W}tbl")):
            rows = []
            for tr in tbl.findall(f"{W}tr"):
                cells = []
                for tc in tr.findall(f"{W}tc"):
                    cells.append(self._text(tc))
                rows.append(cells)
            col_count = len(rows[0]) if rows else 0
            tables.append({
                "index": idx,
                "row_count": len(rows),
                "col_count": col_count,
                "cells": rows,
            })
        return tables
```

Add `TablesMixin` to `__init__.py` composition and add `get_tables` tool to `server.py`:

```python
# server.py addition
@mcp.tool()
def get_tables() -> str:
    """Get all tables with row/column counts and cell text content."""
    return _js(_require_doc().get_tables())
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_reading.py::TestGetTables -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add docx_mcp/document/tables.py docx_mcp/document/__init__.py docx_mcp/server.py tests/test_reading.py
git commit -m "feat: add get_tables tool for reading table content"
```

### Task 1.3: get_styles tool

**Files:**
- Create: `docx_mcp/document/styles.py`
- Modify: `docx_mcp/document/__init__.py`
- Modify: `docx_mcp/server.py`
- Modify: `tests/test_reading.py`

- [ ] **Step 1: Write failing test**

```python
class TestGetStyles:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_returns_styles(self):
        styles = _j(server.get_styles())
        assert len(styles) >= 3  # Heading1, Heading2, FootnoteReference, TableGrid
        ids = {s["id"] for s in styles}
        assert "Heading1" in ids
        assert "TableGrid" in ids

    def test_style_fields(self):
        styles = _j(server.get_styles())
        h1 = next(s for s in styles if s["id"] == "Heading1")
        assert h1["name"] == "heading 1"
        assert h1["type"] == "paragraph"
        assert h1["base_style"] == "Normal"

    def test_no_styles_xml(self, tmp_path: Path):
        """Document without styles.xml returns empty list."""
        # Build minimal docx without styles.xml
        import zipfile
        path = tmp_path / "nostyles.docx"
        # (minimal fixture without word/styles.xml)
        # ...
        server.open_document(str(path))
        styles = _j(server.get_styles())
        assert styles == []
```

- [ ] **Step 2: Run to verify fail, Step 3: Implement**

```python
# docx_mcp/document/styles.py
"""Styles mixin: enumerate document styles."""

from __future__ import annotations

from .base import W


class StylesMixin:
    """Style inspection."""

    def get_styles(self) -> list[dict]:
        tree = self._tree("word/styles.xml")
        if tree is None:
            return []
        result = []
        for s in tree.findall(f"{W}style"):
            name_el = s.find(f"{W}name")
            based_el = s.find(f"{W}basedOn")
            result.append({
                "id": s.get(f"{W}styleId", ""),
                "name": name_el.get(f"{W}val", "") if name_el is not None else "",
                "type": s.get(f"{W}type", ""),
                "base_style": based_el.get(f"{W}val", "") if based_el is not None else "",
            })
        return result
```

- [ ] **Step 4: Run to verify pass, Step 5: Commit**

### Task 1.4: get_headers_footers tool

**Files:**
- Create: `docx_mcp/document/headers_footers.py`
- Modify: `docx_mcp/document/__init__.py`, `docx_mcp/server.py`
- Modify: `tests/test_reading.py`

- [ ] **Step 1: Write failing test**

```python
class TestGetHeadersFooters:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_returns_headers(self):
        hf = _j(server.get_headers_footers())
        assert len(hf) >= 1
        h = hf[0]
        assert h["part"] == "word/header1.xml"
        assert h["location"] == "header"
        # Text may be empty after watermark (just the DRAFT shape)
```

- [ ] **Step 2-3: Implement**

```python
# docx_mcp/document/headers_footers.py
"""Headers/footers mixin."""

from __future__ import annotations

from .base import W


class HeadersFootersMixin:
    """Header and footer operations."""

    def get_headers_footers(self) -> list[dict]:
        results = []
        for rel_path, tree in self._trees.items():
            if not rel_path.startswith("word/header") and not rel_path.startswith("word/footer"):
                continue
            location = "header" if "header" in rel_path else "footer"
            text = self._text(tree)
            results.append({
                "part": rel_path,
                "location": location,
                "text": text,
            })
        return sorted(results, key=lambda x: x["part"])
```

- [ ] **Step 4-5: Verify pass, commit**

### Task 1.5: get_properties tool

**Files:**
- Create: `docx_mcp/document/properties.py`
- Modify: `docx_mcp/document/base.py` (add `DC`, `DCTERMS`, `CP` namespace constants; parse `docProps/core.xml` in `open()`)
- Modify: `docx_mcp/document/__init__.py`, `docx_mcp/server.py`
- Modify: `tests/test_reading.py`

- [ ] **Step 1: Write failing test**

```python
class TestGetProperties:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_returns_properties(self):
        props = _j(server.get_properties())
        assert props["title"] == "Test Document"
        assert props["creator"] == "Test Author"
        assert props["subject"] == "Test Subject"
        assert props["description"] == "Test Description"
        assert props["last_modified_by"] == "Test Editor"
        assert props["revision"] == "3"
        assert "2025-01-01" in props["created"]
        assert "2025-06-15" in props["modified"]
```

- [ ] **Step 2-3: Implement**

Add namespace constants to `base.py`:

```python
DC = "{http://purl.org/dc/elements/1.1/}"
DCTERMS = "{http://purl.org/dc/terms/}"
CP = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
```

Add `docProps/core.xml` to the parsing list in `open()`.

```python
# docx_mcp/document/properties.py
"""Properties mixin: read/write core document properties."""

from __future__ import annotations

from .base import DC, DCTERMS, CP


class PropertiesMixin:
    """Document property operations."""

    def get_properties(self) -> dict:
        tree = self._tree("docProps/core.xml")
        if tree is None:
            return {}

        def _val(tag: str) -> str:
            el = tree.find(tag)
            return el.text if el is not None and el.text else ""

        return {
            "title": _val(f"{DC}title"),
            "creator": _val(f"{DC}creator"),
            "subject": _val(f"{DC}subject"),
            "description": _val(f"{DC}description"),
            "last_modified_by": _val(f"{CP}lastModifiedBy"),
            "revision": _val(f"{CP}revision"),
            "created": _val(f"{DCTERMS}created"),
            "modified": _val(f"{DCTERMS}modified"),
        }
```

- [ ] **Step 4-5: Verify pass, commit**

### Task 1.6: get_images tool

**Files:**
- Create: `docx_mcp/document/images.py`
- Modify: `docx_mcp/document/base.py` (add `WP` namespace constant)
- Modify: `docx_mcp/document/__init__.py`, `docx_mcp/server.py`
- Modify: `tests/test_reading.py`

- [ ] **Step 1: Write failing test**

```python
class TestGetImages:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_returns_images(self):
        images = _j(server.get_images())
        assert len(images) >= 1
        img = images[0]
        assert "rId" in img
        assert img["filename"] == "image1.png"
        assert img["content_type"] == "image/png"
```

- [ ] **Step 2-3: Implement**

Add to `base.py`:
```python
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
```

```python
# docx_mcp/document/images.py
"""Images mixin: list and insert images."""

from __future__ import annotations

from .base import W, R, A, WP, RELS


class ImagesMixin:
    """Image operations."""

    def get_images(self) -> list[dict]:
        doc = self._tree("word/document.xml")
        rels = self._tree("word/_rels/document.xml.rels")
        if doc is None:
            return []
        images = []
        for blip in doc.iter(f"{A}blip"):
            embed = blip.get(f"{R}embed")
            if not embed:
                continue
            info = {"rId": embed, "filename": "", "content_type": ""}
            if rels is not None:
                from .base import RELS as RELS_NS
                rel = rels.find(f'{RELS_NS}Relationship[@Id="{embed}"]')
                if rel is not None:
                    info["filename"] = rel.get("Target", "").split("/")[-1]
            # Get dimensions from wp:extent
            drawing = blip.getparent()
            while drawing is not None and drawing.tag != f"{W}drawing":
                drawing = drawing.getparent()
            if drawing is not None:
                extent = drawing.find(f".//{WP}extent")
                if extent is not None:
                    info["width_emu"] = int(extent.get("cx", "0"))
                    info["height_emu"] = int(extent.get("cy", "0"))
            # Content type from [Content_Types].xml
            ct = self._tree("[Content_Types].xml")
            if ct is not None:
                ext = info["filename"].rsplit(".", 1)[-1] if "." in info["filename"] else ""
                from .base import CT
                for default in ct.findall(f"{CT}Default"):
                    if default.get("Extension") == ext:
                        info["content_type"] = default.get("ContentType", "")
            images.append(info)
        return images
```

- [ ] **Step 4-5: Verify pass, commit**

### Task 1.7: get_endnotes tool

**Files:**
- Create: `docx_mcp/document/endnotes.py`
- Modify: `docx_mcp/document/__init__.py`, `docx_mcp/server.py`
- Modify: `tests/test_reading.py`

- [ ] **Step 1: Write failing test**

```python
class TestGetEndnotes:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_returns_endnotes(self):
        endnotes = _j(server.get_endnotes())
        assert len(endnotes) == 1
        assert endnotes[0]["id"] == 1
        assert "Endnote reference" in endnotes[0]["text"]

    def test_no_endnotes_xml(self):
        """Without endnotes.xml, returns empty list."""
        # Use a fixture without endnotes.xml
```

- [ ] **Step 2-3: Implement**

```python
# docx_mcp/document/endnotes.py
"""Endnotes mixin: read, add, validate endnotes."""

from __future__ import annotations

from .base import W


class EndnotesMixin:
    """Endnote operations."""

    def get_endnotes(self) -> list[dict]:
        tree = self._tree("word/endnotes.xml")
        if tree is None:
            return []
        return [
            {"id": int(en.get(f"{W}id", "0")), "text": self._text(en)}
            for en in tree.findall(f"{W}endnote")
            if en.get(f"{W}id") not in ("0", "-1")
        ]
```

- [ ] **Step 4-5: Verify pass, commit**

### Task 1.8: Coverage check and Phase 1 commit

- [ ] **Step 1: Run full test suite with coverage**

Run: `python -m pytest tests/ --cov=docx_mcp --cov-report=term-missing --cov-fail-under=100`
Expected: 100% coverage

- [ ] **Step 2: Fix any coverage gaps** — add tests for edge cases (empty tables, no images, missing parts)

- [ ] **Step 3: Commit Phase 1 complete**

```bash
git add -A
git commit -m "feat: add 6 read-only tools (tables, styles, headers/footers, properties, images, endnotes)"
```

---

## Phase 2: Track changes complete (3 new tools)

### Task 2.1: accept_changes tool

**Files:**
- Modify: `docx_mcp/document/tracks.py`
- Modify: `docx_mcp/server.py`
- Create: `tests/test_tracks.py`

- [ ] **Step 1: Write failing tests**

```python
# tests/test_tracks.py
"""Tests for accept/reject changes and formatting."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server


def _j(result: str) -> dict | list:
    return json.loads(result)


class TestAcceptChanges:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_accept_all(self):
        # Make some tracked changes first
        server.insert_text("00000004", "NEW ", position="start")
        server.delete_text("00000004", "30 days")
        result = _j(server.accept_changes(scope="all"))
        assert result["accepted"] >= 2
        # Verify: inserted text kept, deleted text gone
        para = _j(server.get_paragraph("00000004"))
        assert "NEW" in para["text"]
        assert "30 days" not in para["text"]

    def test_accept_by_author(self):
        server.insert_text("00000004", "added", author="Alice")
        server.insert_text("00000004", "other", author="Bob")
        result = _j(server.accept_changes(scope="by_author", author="Alice"))
        assert result["accepted"] >= 1

    def test_accept_no_document(self):
        server.close_document()
        with pytest.raises(RuntimeError, match="No document"):
            server.accept_changes(scope="all")
```

- [ ] **Step 2: Run to verify fail**

- [ ] **Step 3: Implement accept_changes**

```python
# Add to tracks.py
def accept_changes(self, *, scope: str = "all", ids: list[int] | None = None,
                   author: str | None = None) -> dict:
    """Accept tracked changes. Removes w:ins wrappers (keeps content),
    removes w:del elements (removes content), handles w:moveTo/w:moveFrom."""
    doc = self._require("word/document.xml")
    accepted = 0

    def _should_process(el: etree._Element) -> bool:
        if scope == "all":
            return True
        if scope == "by_id" and ids:
            eid = el.get(f"{W}id")
            return eid is not None and int(eid) in ids
        if scope == "by_author" and author:
            return el.get(f"{W}author") == author
        return False

    # Accept insertions: unwrap w:ins (keep children)
    for ins in list(doc.iter(f"{W}ins")):
        if not _should_process(ins):
            continue
        parent = ins.getparent()
        idx = list(parent).index(ins)
        for child in list(ins):
            parent.insert(idx, child)
            idx += 1
        parent.remove(ins)
        accepted += 1

    # Accept deletions: remove w:del entirely
    for del_el in list(doc.iter(f"{W}del")):
        if not _should_process(del_el):
            continue
        del_el.getparent().remove(del_el)
        accepted += 1

    # Accept moves: keep w:moveTo content, remove w:moveFrom
    for move_to in list(doc.iter(f"{W}moveTo")):
        if not _should_process(move_to):
            continue
        parent = move_to.getparent()
        idx = list(parent).index(move_to)
        for child in list(move_to):
            parent.insert(idx, child)
            idx += 1
        parent.remove(move_to)
        accepted += 1

    for move_from in list(doc.iter(f"{W}moveFrom")):
        if not _should_process(move_from):
            continue
        move_from.getparent().remove(move_from)
        accepted += 1

    if accepted:
        self._mark("word/document.xml")
    return {"accepted": accepted, "scope": scope}
```

- [ ] **Step 4: Run to verify pass, Step 5: Commit**

### Task 2.2: reject_changes tool

Same pattern as accept but inverted: remove w:ins content, unwrap w:del content.

- [ ] **Step 1: Write failing test**
- [ ] **Step 2: Implement (inverse of accept)**
- [ ] **Step 3: Verify pass, commit**

### Task 2.3: set_formatting tool

**Files:**
- Create: `docx_mcp/document/formatting.py`
- Modify: `docx_mcp/document/__init__.py`, `docx_mcp/server.py`
- Modify: `tests/test_tracks.py`

- [ ] **Step 1: Write failing test**

```python
class TestSetFormatting:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_bold(self):
        result = _j(server.set_formatting(
            para_id="00000004", text="contract", bold=True
        ))
        assert result["formatted"] is True

    def test_multiple_properties(self):
        result = _j(server.set_formatting(
            para_id="00000004", text="30 days",
            bold=True, italic=True, color="FF0000"
        ))
        assert result["formatted"] is True
```

- [ ] **Step 2-3: Implement**

The `set_formatting` method must:
1. Find the run containing the target text
2. Split the run if text is a substring (reuse `_make_run` pattern)
3. Clone original `w:rPr` as `w:rPrChange` child
4. Apply new properties to `w:rPr`
5. Wrap in tracked change attributes

- [ ] **Step 4-5: Verify pass, commit**

### Task 2.4: Phase 2 coverage check

- [ ] **Step 1: Run full coverage**
- [ ] **Step 2: Fix gaps**
- [ ] **Step 3: Commit**

---

## Phase 3: Tables write (4 new tools)

### Task 3.1: add_table tool

- [ ] **Step 1: Write failing test** — create table after paraId, verify XML structure
- [ ] **Step 2: Implement** — build `w:tbl` with `w:tblPr`, `w:tblGrid`, `w:tr`/`w:tc`, wrap in `w:ins`
- [ ] **Step 3: Verify pass, commit**

### Task 3.2: modify_cell tool

- [ ] **Step 1: Write failing test** — modify cell text with tracked changes
- [ ] **Step 2: Implement** — tracked delete old + tracked insert new in cell paragraph
- [ ] **Step 3: Verify pass, commit**

### Task 3.3: add_table_row tool

- [ ] **Step 1: Write failing test** — add row at end and at index
- [ ] **Step 2: Implement** — new `w:tr` wrapped in `w:ins`
- [ ] **Step 3: Verify pass, commit**

### Task 3.4: delete_table_row tool

- [ ] **Step 1: Write failing test** — delete row, verify cell content wrapped in `w:del`
- [ ] **Step 2: Implement** — mark each cell's runs with `w:del`, paragraph mark with `w:rPr > w:del`
- [ ] **Step 3: Verify pass, commit**

### Task 3.5: Phase 3 coverage check and commit

---

## Phase 4: Content creation (5 new tools)

### Task 4.1: add_list tool

**Files:**
- Create: `docx_mcp/document/lists.py`
- Modify: `docx_mcp/document/base.py` (add numbering part bootstrap helper)

- [ ] **Step 1: Write failing test** — bullet list and numbered list
- [ ] **Step 2: Implement** — create `w:abstractNum` + `w:num` in numbering.xml (bootstrap if missing), insert paragraphs with `w:numPr`
- [ ] **Step 3: Verify pass, commit**

### Task 4.2: insert_image tool

- [ ] **Step 1: Write failing test** — insert image at paraId, verify file in media/ and relationship
- [ ] **Step 2: Implement** — copy to word/media/, add rels, add content type, insert `w:drawing` > `wp:inline` > `a:graphic` > `pic:pic` > `a:blip`
- [ ] **Step 3: Verify pass, commit**

### Task 4.3: edit_header_footer tool

- [ ] **Step 1: Write failing test** — edit header text with tracked changes
- [ ] **Step 2: Implement** — tracked delete old + tracked insert new
- [ ] **Step 3: Verify pass, commit**

### Task 4.4: add_endnote and validate_endnotes tools

- [ ] **Step 1: Write failing test** — add endnote, validate cross-references
- [ ] **Step 2: Implement** — mirror footnote pattern but target endnotes.xml
- [ ] **Step 3: Verify pass, commit**

### Task 4.5: Phase 4 coverage check and commit

---

## Phase 5: Document structure (4 new tools)

### Task 5.1: add_page_break tool

**Files:**
- Create: `docx_mcp/document/sections.py`

- [ ] **Step 1: Write failing test** — insert page break after paraId
- [ ] **Step 2: Implement** — insert `w:p` with `w:r` > `w:br w:type="page"` after target
- [ ] **Step 3: Verify pass, commit**

### Task 5.2: add_section_break tool

- [ ] **Step 1: Write failing test** — insert section break, verify `w:sectPr` in `w:pPr`
- [ ] **Step 2: Implement** — add `w:sectPr` with `w:type` inside target paragraph's `w:pPr`
- [ ] **Step 3: Verify pass, commit**

### Task 5.3: set_section_properties tool

- [ ] **Step 1: Write failing test** — modify page size/orientation
- [ ] **Step 2: Implement** — find section's `w:sectPr`, update `w:pgSz` and `w:pgMar`
- [ ] **Step 3: Verify pass, commit**

### Task 5.4: add_cross_reference tool

**Files:**
- Create: `docx_mcp/document/references.py`

- [ ] **Step 1: Write failing test** — add cross-reference to heading
- [ ] **Step 2: Implement** — add bookmark at target if missing, insert `w:hyperlink` with `w:anchor`
- [ ] **Step 3: Verify pass, commit**

### Task 5.5: Phase 5 coverage check and commit

---

## Phase 6: Protection, properties write, and merge (3 new tools)

### Task 6.1: set_document_protection tool

**Files:**
- Create: `docx_mcp/document/protection.py`
- Create: `tests/test_protection.py`

- [ ] **Step 1: Write failing test** — set trackedChanges protection, verify settings.xml
- [ ] **Step 2: Implement** — add/update `w:documentProtection` in settings.xml with SHA-512 hash
- [ ] **Step 3: Verify pass, commit**

### Task 6.2: set_properties tool

- [ ] **Step 1: Write failing test** — set title and creator, verify core.xml
- [ ] **Step 2: Implement** — update `dc:title`, `dc:creator` etc in core.xml (create elements if missing)
- [ ] **Step 3: Verify pass, commit**

### Task 6.3: merge_documents tool

**Files:**
- Create: `docx_mcp/document/merge.py`

This is the most complex tool. Break into sub-steps:

- [ ] **Step 1: Write failing test** — merge two docx, verify combined content, no paraId collisions
- [ ] **Step 2: Implement paraId remapping** — collect all target IDs, remap source IDs
- [ ] **Step 3: Implement rId remapping** — offset source rIds
- [ ] **Step 4: Implement media merge** — copy source images to target
- [ ] **Step 5: Implement footnote/endnote merge** — remap IDs, append elements
- [ ] **Step 6: Implement body content merge** — append source body children before target's final sectPr
- [ ] **Step 7: Implement relationship merge** — append source rels with remapped IDs
- [ ] **Step 8: Verify pass, commit**

### Task 6.4: Phase 6 coverage check and commit

---

## Phase 7: Final validation

### Task 7.1: Extend audit_document

- [ ] **Step 1: Add audit checks** for tables (consistent col count), images (rId + file), endnotes (cross-ref), sections, protection status
- [ ] **Step 2: Write tests for each new audit check**
- [ ] **Step 3: Verify pass, commit**

### Task 7.2: Update README with new tools

- [ ] **Step 1: Update Available Tools section** with all 45 tools
- [ ] **Step 2: Update skill/SKILL.md** with new tool reference
- [ ] **Step 3: Commit**

### Task 7.3: Final coverage and version bump

- [ ] **Step 1: Run full coverage** — must be 100%

Run: `python -m pytest tests/ --cov=docx_mcp --cov-report=term-missing --cov-fail-under=100`

- [ ] **Step 2: Bump version to 0.2.0 in pyproject.toml**
- [ ] **Step 3: Run ruff** — `ruff check docx_mcp/ tests/`
- [ ] **Step 4: Final commit and tag**

```bash
git add -A
git commit -m "feat: expand to 45 tools with tables, formatting, images, properties, and more

Phase 0: Refactored document.py into mixin-based package
Phase 1: 6 read-only tools (tables, styles, headers/footers, properties, images, endnotes)
Phase 2: accept/reject changes, set_formatting
Phase 3: Table write operations (add, modify, add/delete rows)
Phase 4: Lists, image insert, header/footer edit, endnotes
Phase 5: Sections, page breaks, cross-references
Phase 6: Protection, property write, document merge
100% test coverage maintained throughout."

git tag -a v0.2.0 -m "v0.2.0: 45 tools"
git push && git push --tags
```

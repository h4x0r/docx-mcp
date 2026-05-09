"""Tests for TocMixin — generate_toc, update_toc, generate_list_of_figures, generate_list_of_tables."""
from __future__ import annotations

import uuid
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument, W, W14


# ── Helpers ──────────────────────────────────────────────────────────────────


def _make_doc(tmp_path: Path) -> DocxDocument:
    """Create a fresh document."""
    out = str(tmp_path / "test.docx")
    return DocxDocument.create(out)


def _add_heading(doc: DocxDocument, text: str, level: int) -> str:
    """Inject a Heading N paragraph into the document body (before sectPr)."""
    tree = doc._tree("word/document.xml")
    body = tree.find(f"{W}body")
    para_id = uuid.uuid4().hex[:8].upper()
    p = etree.Element(f"{W}p")
    p.set(f"{W14}paraId", para_id)
    pPr = etree.SubElement(p, f"{W}pPr")
    pStyle = etree.SubElement(pPr, f"{W}pStyle")
    pStyle.set(f"{W}val", f"Heading {level}")
    r = etree.SubElement(p, f"{W}r")
    t = etree.SubElement(r, f"{W}t")
    t.text = text
    # Insert before last child (sectPr)
    body.insert(len(body) - 1, p)
    doc._mark("word/document.xml")
    return para_id


def _get_instr_texts(doc: DocxDocument) -> list[str]:
    """Return all instrText strings in the document."""
    tree = doc._tree("word/document.xml")
    return [el.text or "" for el in tree.iter(f"{W}instrText")]


def _get_toc_entry_styles(doc: DocxDocument) -> list[str]:
    """Return pStyle values of paragraphs that have a TOC N entry style (TOC1, TOC2, TOC3)."""
    import re
    tree = doc._tree("word/document.xml")
    styles = []
    for p in tree.iter(f"{W}p"):
        pPr = p.find(f"{W}pPr")
        if pPr is None:
            continue
        pStyle = pPr.find(f"{W}pStyle")
        if pStyle is None:
            continue
        val = pStyle.get(f"{W}val", "")
        if re.match(r"^TOC\d+$", val):
            styles.append(val)
    return styles


# ── Tests ─────────────────────────────────────────────────────────────────────


class TestToc:
    def test_generate_toc_creates_field(self, tmp_path: Path):
        """generate_toc inserts a paragraph with TOC field instrText."""
        doc = _make_doc(tmp_path)
        _add_heading(doc, "Introduction", 1)
        result = doc.generate_toc()

        assert isinstance(result, dict)
        instr_texts = _get_instr_texts(doc)
        assert any(" TOC " in t for t in instr_texts), (
            f"No instrText containing ' TOC ' found; got: {instr_texts}"
        )

    def test_toc_entries_match_headings(self, tmp_path: Path):
        """generate_toc inserts one entry per heading found."""
        doc = _make_doc(tmp_path)
        _add_heading(doc, "Chapter One", 1)
        _add_heading(doc, "Section 1.1", 2)
        result = doc.generate_toc()

        assert result["entry_count"] == 2
        toc_styles = _get_toc_entry_styles(doc)
        assert len(toc_styles) == 2

    def test_toc_max_level_filtering(self, tmp_path: Path):
        """generate_toc with max_level=1 only includes Heading 1 entries."""
        doc = _make_doc(tmp_path)
        _add_heading(doc, "Top Level", 1)
        _add_heading(doc, "Sub Level", 2)
        result = doc.generate_toc(max_level=1)

        assert result["entry_count"] == 1
        toc_styles = _get_toc_entry_styles(doc)
        assert len(toc_styles) == 1
        assert toc_styles[0] == "TOC1"

    def test_update_toc_reflects_new_headings(self, tmp_path: Path):
        """update_toc regenerates entries after new headings are added."""
        doc = _make_doc(tmp_path)
        _add_heading(doc, "First Heading", 1)
        first_result = doc.generate_toc()
        assert first_result["entry_count"] == 1

        # Add another heading after ToC was generated
        _add_heading(doc, "Second Heading", 1)
        update_result = doc.update_toc()

        assert update_result["updated"] is True
        assert update_result["entry_count"] == 2
        toc_styles = _get_toc_entry_styles(doc)
        assert len(toc_styles) == 2

    def test_generate_lof_requires_seq_captions(self, tmp_path: Path):
        """generate_list_of_figures inserts a TOC \\c "Figure" field."""
        doc = _make_doc(tmp_path)
        result = doc.generate_list_of_figures()

        assert isinstance(result, dict)
        instr_texts = _get_instr_texts(doc)
        assert any("Figure" in t for t in instr_texts), (
            f"No instrText containing 'Figure' found; got: {instr_texts}"
        )
        # Blank doc has no SEQ Figure captions → entry_count == 0
        assert result["entry_count"] == 0

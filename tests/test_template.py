"""Tests for TemplateMixin — fill_template, list_template_fields, validate_template_data."""
from __future__ import annotations

import uuid
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument, W, W14


# ── Helpers ──────────────────────────────────────────────────────────────────


def _make_doc(tmp_path: Path) -> DocxDocument:
    out = str(tmp_path / "test.docx")
    return DocxDocument.create(out)


def _inject_para(doc: DocxDocument, text: str = "placeholder") -> str:
    """Inject a plain paragraph into the body and return its para_id."""
    tree = doc._tree("word/document.xml")
    body = tree.find(f"{W}body")
    para_id = uuid.uuid4().hex[:8].upper()
    p = etree.Element(f"{W}p")
    p.set(f"{W14}paraId", para_id)
    r = etree.SubElement(p, f"{W}r")
    t = etree.SubElement(r, f"{W}t")
    t.text = text
    # Insert before last child (sectPr or last element)
    body.insert(len(body) - 1, p)
    doc._mark("word/document.xml")
    return para_id


def _make_template_doc(tmp_path: Path) -> DocxDocument:
    """Create a doc with two text SDT controls tagged CLIENT_NAME and DATE."""
    doc = _make_doc(tmp_path)
    p1_id = _inject_para(doc, "Client: placeholder")
    p2_id = _inject_para(doc, "Date: placeholder")
    doc.add_content_control(p1_id, "CLIENT_NAME", "text", label="Client Name")
    doc.add_content_control(p2_id, "DATE", "text", label="Date")
    return doc


# ── Tests ─────────────────────────────────────────────────────────────────────


class TestFillTemplate:
    def test_fill_text_control(self, tmp_path):
        """fill_template fills a single SDT text control."""
        doc = _make_template_doc(tmp_path)
        result = doc.fill_template({"CLIENT_NAME": "Acme Corp", "DATE": "2026-06-01"})
        assert result["filled"] == 2
        assert result["unfilled"] == []
        # Verify text is in the XML
        tree = doc._tree("word/document.xml")
        texts = [t.text for t in tree.iter(f"{W}t") if t.text]
        assert "Acme Corp" in texts

    def test_fill_multiple_fields(self, tmp_path):
        """fill_template fills all matching SDTs."""
        doc = _make_template_doc(tmp_path)
        result = doc.fill_template({"CLIENT_NAME": "Beta LLC", "DATE": "2026-01-01"})
        assert result["filled"] == 2
        tree = doc._tree("word/document.xml")
        texts = [t.text for t in tree.iter(f"{W}t") if t.text]
        assert "Beta LLC" in texts
        assert "2026-01-01" in texts

    def test_unfilled_fields_reported(self, tmp_path):
        """fill_template reports tags not present in data."""
        doc = _make_template_doc(tmp_path)
        result = doc.fill_template({"CLIENT_NAME": "X Corp"})
        assert result["filled"] == 1
        assert "DATE" in result["unfilled"]

    def test_remove_empty_controls(self, tmp_path):
        """remove_empty=True removes SDTs with no matching data key."""
        doc = _make_template_doc(tmp_path)
        doc.fill_template({"CLIENT_NAME": "Gamma"}, remove_empty=True)
        # DATE has no match and remove_empty=True -> should be removed
        tree = doc._tree("word/document.xml")
        sdts = list(tree.iter(f"{W}sdt"))
        tags = []
        for sdt in sdts:
            sdtPr = sdt.find(f"{W}sdtPr")
            if sdtPr is not None:
                tag_el = sdtPr.find(f"{W}tag")
                if tag_el is not None:
                    tags.append(tag_el.get(f"{W}val", ""))
        assert "DATE" not in tags

    def test_list_template_fields(self, tmp_path):
        """list_template_fields returns all SDT tags."""
        doc = _make_template_doc(tmp_path)
        fields = doc.list_template_fields()
        tags = {f["tag"] for f in fields}
        assert "CLIENT_NAME" in tags
        assert "DATE" in tags

    def test_validate_missing_fields(self, tmp_path):
        """validate_template_data reports missing and extra keys."""
        doc = _make_template_doc(tmp_path)
        result = doc.validate_template_data({"CLIENT_NAME": "X", "EXTRA_KEY": "Y"})
        assert result["valid"] is False
        assert "DATE" in result["missing"]
        assert "EXTRA_KEY" in result["extra"]

    def test_repeating_section_list(self, tmp_path):
        """fill_template with list value creates multiple SDT clones."""
        doc = _make_doc(tmp_path)
        p_id = _inject_para(doc, "Item placeholder")
        doc.add_content_control(p_id, "ITEM", "text")
        result = doc.fill_template({"ITEM": ["Alpha", "Beta", "Gamma"]})
        assert result["filled"] >= 1
        # Should have text from all 3 items in the XML
        tree = doc._tree("word/document.xml")
        texts = [t.text for t in tree.iter(f"{W}t") if t.text]
        assert "Alpha" in texts
        assert "Beta" in texts
        assert "Gamma" in texts

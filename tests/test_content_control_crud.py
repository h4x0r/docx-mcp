"""RED tests for content control CRUD — delete, get, update by control_id (w:id)."""
from __future__ import annotations

from pathlib import Path

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument, W, W14
from docx_mcp.document.errors import DocxMcpError


# ── helpers ──────────────────────────────────────────────────────────────────


def _make_doc_with_sdt(tmp_path: Path, control_id: str = "42") -> DocxDocument:
    """Create a document, wrap first para in a text SDT, inject w:id."""
    out = str(tmp_path / "test.docx")
    doc = DocxDocument.create(out)
    tree = doc._tree("word/document.xml")

    # Find first paragraph with a paraId
    para_id = None
    for p in tree.iter(f"{W}p"):
        pid = p.get(f"{W14}paraId")
        if pid is not None:
            para_id = pid
            break
    assert para_id is not None

    # Wrap it in a text SDT
    doc.add_content_control(para_id, "mytag", "text", label="My Label")

    # Inject w:id into sdtPr
    tree = doc._tree("word/document.xml")
    sdt = tree.find(f".//{W}sdt")
    assert sdt is not None
    sdtPr = sdt.find(f"{W}sdtPr")
    assert sdtPr is not None
    id_el = etree.SubElement(sdtPr, f"{W}id")
    id_el.set(f"{W}val", control_id)

    return doc


# ── delete_content_control ────────────────────────────────────────────────────


class TestDeleteContentControl:
    def test_delete_content_control_unwraps_content(self, tmp_path: Path):
        """SDT wrapper removed; child content (w:p) remains in parent."""
        doc = _make_doc_with_sdt(tmp_path, "99")
        result = doc.delete_content_control("99")

        assert result == {"control_id": "99", "deleted": True}

        tree = doc._tree("word/document.xml")
        # No more w:sdt elements
        sdts = list(tree.iter(f"{W}sdt"))
        assert len(sdts) == 0, f"Expected 0 SDTs after unwrap, got {len(sdts)}"

        # At least one w:p remains in the body
        body = tree.find(f".//{W}body")
        assert body is not None
        paras = list(body.iter(f"{W}p"))
        assert len(paras) > 0, "Expected content paragraphs to remain"

    def test_delete_content_control_not_found_raises(self, tmp_path: Path):
        """Bad control_id raises ValueError."""
        doc = _make_doc_with_sdt(tmp_path, "10")
        with pytest.raises(ValueError, match="not found"):
            doc.delete_content_control("NONEXISTENT")


# ── get_content_control ────────────────────────────────────────────────────────


class TestGetContentControl:
    def test_get_content_control_returns_single(self, tmp_path: Path):
        """Returns dict with expected keys for a known control_id."""
        doc = _make_doc_with_sdt(tmp_path, "7")
        result = doc.get_content_control("7")

        assert isinstance(result, dict)
        assert result["control_id"] == "7"
        assert result["type"] == "text"
        assert result["title"] == "My Label"
        assert result["tag"] == "mytag"
        assert "value" in result

    def test_get_content_control_not_found_raises(self, tmp_path: Path):
        """Bad control_id raises ValueError."""
        doc = _make_doc_with_sdt(tmp_path, "5")
        with pytest.raises(ValueError, match="not found"):
            doc.get_content_control("MISSING")


# ── update_content_control ─────────────────────────────────────────────────────


class TestUpdateContentControl:
    def test_update_content_control_title(self, tmp_path: Path):
        """Updating title sets w:alias/@w:val in sdtPr."""
        doc = _make_doc_with_sdt(tmp_path, "3")
        result = doc.update_content_control("3", title="New Title")

        assert result["control_id"] == "3"
        assert result["title"] == "New Title"

        tree = doc._tree("word/document.xml")
        sdt = tree.find(f".//{W}sdt")
        assert sdt is not None
        sdtPr = sdt.find(f"{W}sdtPr")
        alias_el = sdtPr.find(f"{W}alias")
        assert alias_el is not None
        assert alias_el.get(f"{W}val") == "New Title"

    def test_update_content_control_tag(self, tmp_path: Path):
        """Updating tag sets w:tag/@w:val in sdtPr."""
        doc = _make_doc_with_sdt(tmp_path, "4")
        result = doc.update_content_control("4", tag="newtag")

        assert result["control_id"] == "4"
        assert result["tag"] == "newtag"

        tree = doc._tree("word/document.xml")
        sdt = tree.find(f".//{W}sdt")
        sdtPr = sdt.find(f"{W}sdtPr")
        tag_el = sdtPr.find(f"{W}tag")
        assert tag_el is not None
        assert tag_el.get(f"{W}val") == "newtag"

    def test_update_content_control_not_found_raises(self, tmp_path: Path):
        """Bad control_id raises ValueError."""
        doc = _make_doc_with_sdt(tmp_path, "1")
        with pytest.raises(ValueError, match="not found"):
            doc.update_content_control("BOGUS", title="x")

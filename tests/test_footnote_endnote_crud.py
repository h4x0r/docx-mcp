"""Tests for footnote and endnote CRUD operations (update, delete)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server


def _j(result: str) -> dict | list:
    return json.loads(result)


# ═══════════════════════════════════════════════════════════════════════════
#  TestFootnoteCRUD
# ═══════════════════════════════════════════════════════════════════════════


class TestFootnoteCRUD:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_update_footnote(self):
        """Update existing footnote #1 text and verify the result."""
        result = _j(server.update_footnote(1, "Updated footnote text."))
        assert result["footnote_id"] == 1
        assert result["text"] == "Updated footnote text."
        # Verify via get_footnotes
        footnotes = _j(server.get_footnotes())
        fn1 = next(f for f in footnotes if f["id"] == 1)
        assert "Updated footnote text." in fn1["text"]

    def test_update_footnote_not_found(self):
        """Updating a non-existent footnote raises ValueError."""
        with pytest.raises(ValueError, match="not found"):
            server.update_footnote(999, "Should fail")

    def test_update_footnote_builtin_rejected(self):
        """Updating built-in footnote (id < 1) raises ValueError."""
        with pytest.raises(ValueError):
            server.update_footnote(0, "Should fail")

    def test_delete_footnote_removes_from_xml(self):
        """delete_footnote removes the definition from footnotes.xml."""
        result = _j(server.delete_footnote(1))
        assert result["deleted"] == 1
        footnotes = _j(server.get_footnotes())
        ids = [f["id"] for f in footnotes]
        assert 1 not in ids

    def test_delete_footnote_removes_reference(self):
        """delete_footnote also removes the footnoteReference run in document.xml."""
        server.delete_footnote(1)
        # validate_footnotes should still report valid (no dangling refs)
        validation = _j(server.validate_footnotes())
        assert validation["valid"] is True
        assert 1 not in validation.get("missing_definitions", [])

    def test_delete_footnote_not_found(self):
        """Deleting a non-existent footnote raises ValueError."""
        with pytest.raises(ValueError, match="not found"):
            server.delete_footnote(999)

    def test_update_footnote_then_read_back(self):
        """Round-trip: add a new footnote, update it, confirm text changed."""
        add_result = _j(server.add_footnote("00000004", "Initial text"))
        fid = add_result["footnote_id"]
        _j(server.update_footnote(fid, "Revised text"))
        footnotes = _j(server.get_footnotes())
        fn = next(f for f in footnotes if f["id"] == fid)
        assert "Revised text" in fn["text"]


# ═══════════════════════════════════════════════════════════════════════════
#  TestEndnoteCRUD
# ═══════════════════════════════════════════════════════════════════════════


class TestEndnoteCRUD:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    def test_update_endnote(self):
        """Update existing endnote #1 text and verify the result."""
        result = _j(server.update_endnote(1, "Updated endnote text."))
        assert result["endnote_id"] == 1
        assert result["text"] == "Updated endnote text."
        # Verify via get_endnotes
        endnotes = _j(server.get_endnotes())
        en1 = next(e for e in endnotes if e["id"] == 1)
        assert "Updated endnote text." in en1["text"]

    def test_update_endnote_not_found(self):
        """Updating a non-existent endnote raises ValueError."""
        with pytest.raises(ValueError, match="not found"):
            server.update_endnote(999, "Should fail")

    def test_update_endnote_builtin_rejected(self):
        """Updating built-in endnote (id < 1) raises ValueError."""
        with pytest.raises(ValueError):
            server.update_endnote(0, "Should fail")

    def test_delete_endnote_removes_from_xml(self):
        """delete_endnote removes the definition from endnotes.xml."""
        result = _j(server.delete_endnote(1))
        assert result["deleted"] == 1
        endnotes = _j(server.get_endnotes())
        ids = [e["id"] for e in endnotes]
        assert 1 not in ids

    def test_delete_endnote_removes_reference(self):
        """delete_endnote also removes the endnoteReference run in document.xml."""
        server.delete_endnote(1)
        # validate_endnotes should report valid (no dangling refs)
        validation = _j(server.validate_endnotes())
        assert validation["valid"] is True
        assert 1 not in validation.get("orphaned_refs", [])

    def test_delete_endnote_not_found(self):
        """Deleting a non-existent endnote raises ValueError."""
        with pytest.raises(ValueError, match="not found"):
            server.delete_endnote(999)

    def test_update_endnote_then_read_back(self):
        """Round-trip: add a new endnote, update it, confirm text changed."""
        add_result = _j(server.add_endnote("00000004", "Initial endnote"))
        eid = add_result["endnote_id"]
        _j(server.update_endnote(eid, "Revised endnote"))
        endnotes = _j(server.get_endnotes())
        en = next(e for e in endnotes if e["id"] == eid)
        assert "Revised endnote" in en["text"]

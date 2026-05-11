"""RED tests for set_different_first_page and set_odd_even_headers."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument, W


def _make_doc(tmp_path: Path) -> DocxDocument:
    return DocxDocument.create(str(tmp_path / "test.docx"))


class TestSetDifferentFirstPage:
    def test_set_different_first_page_enabled(self, tmp_path):
        """Enabling should add w:titlePg to the target sectPr."""
        doc = _make_doc(tmp_path)
        result = doc.set_different_first_page(0, True)
        assert result == {"section_index": 0, "different_first_page": True}

        xml_doc = doc._require("word/document.xml")
        body = xml_doc.find(f"{W}body")
        sectprs = doc._collect_sectprs(body)
        sect_pr, _ = sectprs[0]
        assert sect_pr.find(f"{W}titlePg") is not None

    def test_set_different_first_page_disabled(self, tmp_path):
        """After enabling, disabling should remove w:titlePg."""
        doc = _make_doc(tmp_path)
        doc.set_different_first_page(0, True)

        result = doc.set_different_first_page(0, False)
        assert result == {"section_index": 0, "different_first_page": False}

        xml_doc = doc._require("word/document.xml")
        body = xml_doc.find(f"{W}body")
        sectprs = doc._collect_sectprs(body)
        sect_pr, _ = sectprs[0]
        assert sect_pr.find(f"{W}titlePg") is None

    def test_set_different_first_page_invalid_index(self, tmp_path):
        """Out-of-range index raises ValueError."""
        doc = _make_doc(tmp_path)
        with pytest.raises(ValueError):
            doc.set_different_first_page(99, True)

    def test_set_different_first_page_idempotent_enable(self, tmp_path):
        """Enabling twice should be safe (only one w:titlePg present)."""
        doc = _make_doc(tmp_path)
        doc.set_different_first_page(0, True)
        doc.set_different_first_page(0, True)

        xml_doc = doc._require("word/document.xml")
        body = xml_doc.find(f"{W}body")
        sectprs = doc._collect_sectprs(body)
        sect_pr, _ = sectprs[0]
        count = len(sect_pr.findall(f"{W}titlePg"))
        assert count == 1

    def test_set_different_first_page_marks_dirty(self, tmp_path):
        """Calling the method should mark word/document.xml dirty."""
        doc = _make_doc(tmp_path)
        doc._modified.clear()
        doc.set_different_first_page(0, True)
        assert "word/document.xml" in doc._modified


class TestSetOddEvenHeaders:
    def test_set_odd_even_headers_enabled(self, tmp_path):
        """Enabling should add w:evenAndOddHeaders to settings.xml root."""
        doc = _make_doc(tmp_path)
        result = doc.set_odd_even_headers(True)
        assert result == {"odd_even_headers": True}

        settings = doc._tree("word/settings.xml")
        assert settings is not None
        assert settings.find(f"{W}evenAndOddHeaders") is not None

    def test_set_odd_even_headers_disabled(self, tmp_path):
        """After enabling, disabling should remove w:evenAndOddHeaders."""
        doc = _make_doc(tmp_path)
        doc.set_odd_even_headers(True)

        result = doc.set_odd_even_headers(False)
        assert result == {"odd_even_headers": False}

        settings = doc._tree("word/settings.xml")
        assert settings is not None
        assert settings.find(f"{W}evenAndOddHeaders") is None

    def test_set_odd_even_headers_idempotent(self, tmp_path):
        """Enabling twice should be safe (only one element present)."""
        doc = _make_doc(tmp_path)
        doc.set_odd_even_headers(True)
        doc.set_odd_even_headers(True)

        settings = doc._tree("word/settings.xml")
        count = len(settings.findall(f"{W}evenAndOddHeaders"))
        assert count == 1

    def test_set_odd_even_headers_disable_when_not_set(self, tmp_path):
        """Disabling when not set should be a no-op (no error)."""
        doc = _make_doc(tmp_path)
        result = doc.set_odd_even_headers(False)
        assert result == {"odd_even_headers": False}

    def test_set_odd_even_headers_marks_dirty(self, tmp_path):
        """Calling should mark word/settings.xml dirty."""
        doc = _make_doc(tmp_path)
        doc._modified.clear()
        doc.set_odd_even_headers(True)
        assert "word/settings.xml" in doc._modified

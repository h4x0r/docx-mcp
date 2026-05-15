"""Tests for page setup convenience methods: set_page_size, set_page_margins, set_page_orientation."""  # noqa: E501

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server
from docx_mcp.document import W


def _j(result: str) -> dict | list:
    return json.loads(result)


# DXA conversion helper (mirrors implementation)
def _mm_to_dxa(mm: float) -> int:
    return round(mm * 1440 / 25.4)


class TestPageSetup:
    @pytest.fixture(autouse=True)
    def _open(self, test_docx: Path):
        server.open_document(str(test_docx))

    # ── set_page_size ─────────────────────────────────────────────────────────

    def test_set_page_size_a4(self):
        result = _j(server.set_page_size(210.0, 297.0))
        expected_w = _mm_to_dxa(210.0)  # 11906
        expected_h = _mm_to_dxa(297.0)  # 16838
        assert result["width_mm"] == 210.0
        assert result["height_mm"] == 297.0
        assert result["width_dxa"] == expected_w
        assert result["height_dxa"] == expected_h
        # Verify XML was written
        doc = server._doc._trees["word/document.xml"]
        body = doc.find(f"{W}body")
        sect_pr = body.find(f"{W}sectPr")
        pg_sz = sect_pr.find(f"{W}pgSz")
        assert pg_sz is not None
        assert int(pg_sz.get(f"{W}w")) == expected_w
        assert int(pg_sz.get(f"{W}h")) == expected_h

    def test_set_page_size_letter(self):
        result = _j(server.set_page_size(215.9, 279.4))
        expected_w = _mm_to_dxa(215.9)
        expected_h = _mm_to_dxa(279.4)
        assert result["width_mm"] == 215.9
        assert result["height_mm"] == 279.4
        assert result["width_dxa"] == expected_w
        assert result["height_dxa"] == expected_h

    # ── set_page_margins ──────────────────────────────────────────────────────

    def test_set_page_margins_all(self):
        result = _j(server.set_page_margins(
            top_mm=25.4, bottom_mm=25.4, left_mm=31.75, right_mm=31.75
        ))
        margins = result["margins_mm"]
        assert margins["top"] == 25.4
        assert margins["bottom"] == 25.4
        assert margins["left"] == 31.75
        assert margins["right"] == 31.75
        # Verify XML
        doc = server._doc._trees["word/document.xml"]
        body = doc.find(f"{W}body")
        sect_pr = body.find(f"{W}sectPr")
        pg_mar = sect_pr.find(f"{W}pgMar")
        assert pg_mar is not None
        assert int(pg_mar.get(f"{W}top")) == _mm_to_dxa(25.4)
        assert int(pg_mar.get(f"{W}bottom")) == _mm_to_dxa(25.4)
        assert int(pg_mar.get(f"{W}left")) == _mm_to_dxa(31.75)
        assert int(pg_mar.get(f"{W}right")) == _mm_to_dxa(31.75)

    def test_set_page_margins_partial(self):
        result = _j(server.set_page_margins(top_mm=20.0))
        margins = result["margins_mm"]
        assert "top" in margins
        assert margins["top"] == 20.0
        # Only top should be present in response
        assert "bottom" not in margins
        assert "left" not in margins
        assert "right" not in margins

    # ── set_page_orientation ──────────────────────────────────────────────────

    def test_set_page_orientation_landscape(self):
        # First set a known size (portrait A4)
        server.set_page_size(210.0, 297.0)
        result = _j(server.set_page_orientation("landscape"))
        assert result["orientation"] == "landscape"
        # landscape: width > height
        assert result["width_dxa"] > result["height_dxa"]
        # Verify XML
        doc = server._doc._trees["word/document.xml"]
        body = doc.find(f"{W}body")
        sect_pr = body.find(f"{W}sectPr")
        pg_sz = sect_pr.find(f"{W}pgSz")
        assert pg_sz.get(f"{W}orient") == "landscape"
        assert int(pg_sz.get(f"{W}w")) > int(pg_sz.get(f"{W}h"))

    def test_set_page_orientation_portrait(self):
        # Start in portrait, go landscape, then back to portrait
        server.set_page_size(210.0, 297.0)
        server.set_page_orientation("landscape")
        result = _j(server.set_page_orientation("portrait"))
        assert result["orientation"] == "portrait"
        # portrait: width < height
        assert result["width_dxa"] < result["height_dxa"]
        # Verify XML
        doc = server._doc._trees["word/document.xml"]
        body = doc.find(f"{W}body")
        sect_pr = body.find(f"{W}sectPr")
        pg_sz = sect_pr.find(f"{W}pgSz")
        # orient attribute should be "portrait" or absent
        orient = pg_sz.get(f"{W}orient")
        assert orient in (None, "portrait")
        assert int(pg_sz.get(f"{W}w")) < int(pg_sz.get(f"{W}h"))

    def test_set_page_orientation_invalid(self):
        with pytest.raises(ValueError, match="portrait.*landscape|landscape.*portrait|orientation"):
            server._doc.set_page_orientation("sideways")

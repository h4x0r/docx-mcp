"""Tests for paragraph formatting: indentation, line spacing, get_paragraph_format."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"

CM_TO_TWIPS = 567  # 1 cm = 567 twips (round(1440/2.54))


def _open(path: Path) -> None:
    server._doc = None
    server.open_document(str(path))


def _root():
    return server._doc._tree("word/document.xml")


def _para(para_id: str):
    return server._doc._find_para(_root(), para_id)


class TestSetParagraphIndentation:
    def test_set_paragraph_indentation_left_right(self, test_docx: Path) -> None:
        _open(test_docx)
        result = json.loads(server.set_paragraph_indentation("00000002", left_cm=2.0, right_cm=1.0))
        assert result["para_id"] == "00000002"
        assert result["left_cm"] == 2.0
        assert result["right_cm"] == 1.0

        para = _para("00000002")
        ppr = para.find(f"{W}pPr")
        assert ppr is not None
        ind = ppr.find(f"{W}ind")
        assert ind is not None
        assert ind.get(f"{W}left") == str(round(2.0 * CM_TO_TWIPS))
        assert ind.get(f"{W}right") == str(round(1.0 * CM_TO_TWIPS))

    def test_set_paragraph_indentation_first_line(self, test_docx: Path) -> None:
        _open(test_docx)
        result = json.loads(server.set_paragraph_indentation("00000002", first_line_cm=1.27))
        assert result["first_line_cm"] == 1.27

        para = _para("00000002")
        ppr = para.find(f"{W}pPr")
        ind = ppr.find(f"{W}ind")
        assert ind is not None
        assert ind.get(f"{W}firstLine") == str(round(1.27 * CM_TO_TWIPS))
        # hanging should NOT be set
        assert ind.get(f"{W}hanging") is None

    def test_set_paragraph_indentation_both_raises(self, test_docx: Path) -> None:
        _open(test_docx)
        with pytest.raises(ValueError, match="mutually exclusive"):
            server._doc.set_paragraph_indentation("00000002", first_line_cm=1.0, hanging_cm=0.5)


class TestSetLineSpacing:
    def test_set_line_spacing_auto(self, test_docx: Path) -> None:
        _open(test_docx)
        # 1.5 lines = 360 in 240ths-of-a-line units
        result = json.loads(server.set_line_spacing("00000002", line_rule="auto", line_value=360))
        assert result["para_id"] == "00000002"
        assert result["line_rule"] == "auto"
        assert result["line_value"] == 360

        para = _para("00000002")
        ppr = para.find(f"{W}pPr")
        assert ppr is not None
        spacing = ppr.find(f"{W}spacing")
        assert spacing is not None
        assert spacing.get(f"{W}lineRule") == "auto"
        assert spacing.get(f"{W}line") == "360"

    def test_set_line_spacing_space_before_after(self, test_docx: Path) -> None:
        _open(test_docx)
        result = json.loads(
            server.set_line_spacing("00000003", space_before_pt=12.0, space_after_pt=6.0)
        )
        assert result["space_before_pt"] == 12.0
        assert result["space_after_pt"] == 6.0

        para = _para("00000003")
        ppr = para.find(f"{W}pPr")
        spacing = ppr.find(f"{W}spacing")
        assert spacing is not None
        # 12 pt * 20 = 240 twips; 6 pt * 20 = 120 twips
        assert spacing.get(f"{W}before") == "240"
        assert spacing.get(f"{W}after") == "120"


class TestGetParagraphFormat:
    def test_get_paragraph_format_full(self, test_docx: Path) -> None:
        _open(test_docx)
        # Set up some formatting first
        server._doc.set_paragraph_indentation("00000002", left_cm=1.0, right_cm=0.5)
        server._doc.set_line_spacing(
            "00000002", line_rule="auto", line_value=240, space_before_pt=6.0, space_after_pt=3.0
        )
        server._doc.set_paragraph_border("00000002", ["top"])
        server._doc.set_paragraph_shading("00000002", "FF0000")

        result = json.loads(server.get_paragraph_format("00000002"))

        assert "indentation" in result
        assert result["indentation"]["left_twips"] == round(1.0 * CM_TO_TWIPS)
        assert result["indentation"]["right_twips"] == round(0.5 * CM_TO_TWIPS)

        assert "spacing" in result
        assert result["spacing"]["line_value"] == 240
        assert result["spacing"]["line_rule"] == "auto"
        assert result["spacing"]["before_twips"] == 120  # 6pt * 20
        assert result["spacing"]["after_twips"] == 60  # 3pt * 20

        assert result["border"] is True
        assert result["shading"] is True

    def test_get_paragraph_format_bare(self, test_docx: Path) -> None:
        _open(test_docx)
        # Paragraph with no explicit formatting
        result = json.loads(server.get_paragraph_format("00000001"))

        assert isinstance(result["style"], str)
        assert isinstance(result["alignment"], str)

        ind = result["indentation"]
        assert ind["left_twips"] == 0
        assert ind["right_twips"] == 0
        assert ind["first_line_twips"] == 0
        assert ind["hanging_twips"] == 0

        sp = result["spacing"]
        assert sp["before_twips"] == 0
        assert sp["after_twips"] == 0
        assert sp["line_value"] == 0
        assert sp["line_rule"] == ""

        assert result["border"] is False
        assert result["shading"] is False
        assert result["numPr"] is None

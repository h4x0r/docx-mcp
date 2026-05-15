"""Tests for run-level formatting API: get_runs and set_run_* methods."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"


class TestRunFormatting:
    def _open(self, path: Path) -> None:
        server._doc = None
        server.open_document(str(path))

    # ── get_runs ───────────────────────────────────────────────────────────

    def test_get_runs_returns_correct_count(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.get_runs("00000006"))
        assert len(result) == 3

    def test_get_runs_returns_text(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.get_runs("00000006"))
        texts = [r["text"] for r in result]
        assert texts[0] == "First "
        assert texts[1] == "bold"
        assert texts[2] == " last"

    def test_get_runs_run_idx_sequential(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.get_runs("00000006"))
        assert [r["run_idx"] for r in result] == [0, 1, 2]

    def test_get_runs_bold_detected(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.get_runs("00000006"))
        assert result[0]["bold"] is False
        assert result[1]["bold"] is True
        assert result[2]["bold"] is False

    def test_get_runs_font_none_when_absent(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.get_runs("00000001"))
        assert result[0]["font"] is None

    def test_get_runs_size_none_when_absent(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.get_runs("00000001"))
        assert result[0]["size_pt"] is None

    def test_get_runs_color_none_when_absent(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.get_runs("00000001"))
        assert result[0]["color"] is None

    def test_get_runs_raises_on_bad_para_id(self, test_docx: Path) -> None:
        self._open(test_docx)
        with pytest.raises(ValueError, match="not found"):
            server._doc.get_runs("DEADBEEF")

    # ── set_run_font ───────────────────────────────────────────────────────

    def test_set_run_font_sets_rfonts_attributes(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_run_font("00000006", 0, "Arial"))
        assert result["para_id"] == "00000006"
        assert result["run_idx"] == 0
        assert result["font"] == "Arial"

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        assert rpr is not None
        rfonts = rpr.find(f"{W}rFonts")
        assert rfonts is not None
        assert rfonts.get(f"{W}ascii") == "Arial"
        assert rfonts.get(f"{W}hAnsi") == "Arial"
        assert rfonts.get(f"{W}cs") == "Arial"

    def test_set_run_font_raises_index_error_on_bad_run_idx(self, test_docx: Path) -> None:
        self._open(test_docx)
        with pytest.raises(IndexError):
            server._doc.set_run_font("00000001", 99, "Arial")

    def test_set_run_font_raises_value_error_on_bad_para_id(self, test_docx: Path) -> None:
        self._open(test_docx)
        with pytest.raises(ValueError, match="not found"):
            server._doc.set_run_font("DEADBEEF", 0, "Arial")

    # ── set_run_color ──────────────────────────────────────────────────────

    def test_set_run_color_sets_color_val(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_run_color("00000006", 1, "FF0000"))
        assert result["para_id"] == "00000006"
        assert result["run_idx"] == 1
        assert result["color"] == "FF0000"

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[1].find(f"{W}rPr")
        assert rpr is not None
        color_el = rpr.find(f"{W}color")
        assert color_el is not None
        assert color_el.get(f"{W}val") == "FF0000"

    # ── set_run_size ───────────────────────────────────────────────────────

    def test_set_run_size_sets_sz_in_half_points(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_run_size("00000006", 0, 12.0))
        assert result["para_id"] == "00000006"
        assert result["run_idx"] == 0
        assert result["size_pt"] == 12.0

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        assert rpr is not None
        sz = rpr.find(f"{W}sz")
        assert sz is not None
        assert sz.get(f"{W}val") == "24"

    def test_set_run_size_also_sets_sz_cs(self, test_docx: Path) -> None:
        self._open(test_docx)
        server.set_run_size("00000006", 0, 14.0)

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        assert rpr is not None
        sz_cs = rpr.find(f"{W}szCs")
        assert sz_cs is not None
        assert sz_cs.get(f"{W}val") == "28"

    # ── set_character_spacing ──────────────────────────────────────────────

    def test_set_character_spacing_sets_twips(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_character_spacing("00000006", 0, 1.0))
        assert result["para_id"] == "00000006"
        assert result["run_idx"] == 0
        assert result["spacing_pt"] == 1.0

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        assert rpr is not None
        spacing_el = rpr.find(f"{W}spacing")
        assert spacing_el is not None
        assert spacing_el.get(f"{W}val") == "20"

    # ── set_character_position ─────────────────────────────────────────────

    def test_set_character_position_sets_half_points(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_character_position("00000006", 0, 3.0))
        assert result["para_id"] == "00000006"
        assert result["run_idx"] == 0
        assert result["position_pt"] == 3.0

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        assert rpr is not None
        pos_el = rpr.find(f"{W}position")
        assert pos_el is not None
        assert pos_el.get(f"{W}val") == "6"

    def test_set_character_position_negative_value(self, test_docx: Path) -> None:
        self._open(test_docx)
        server.set_character_position("00000006", 0, -2.0)

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        pos_el = rpr.find(f"{W}position")
        assert pos_el.get(f"{W}val") == "-4"

    # ── get_runs with font/color/size present ──────────────────────────────

    def test_get_runs_reflects_font_after_set(self, test_docx: Path) -> None:
        self._open(test_docx)
        server.set_run_font("00000006", 0, "Times New Roman")
        result = json.loads(server.get_runs("00000006"))
        assert result[0]["font"] == "Times New Roman"

    def test_get_runs_reflects_color_after_set(self, test_docx: Path) -> None:
        self._open(test_docx)
        server.set_run_color("00000006", 0, "0000FF")
        result = json.loads(server.get_runs("00000006"))
        assert result[0]["color"] == "0000FF"

    def test_get_runs_reflects_size_after_set(self, test_docx: Path) -> None:
        self._open(test_docx)
        server.set_run_size("00000006", 0, 10.0)
        result = json.loads(server.get_runs("00000006"))
        assert result[0]["size_pt"] == 10.0

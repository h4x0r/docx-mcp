"""Tests for run text effects: highlight, strikethrough, super/subscript, underline."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


class TestRunEffects:
    def _open(self, path: Path) -> None:
        server._doc = None
        server.open_document(str(path))

    def _rpr(self, test_docx, para_id, run_idx):
        self._open(test_docx)
        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, para_id)
        runs = para.findall(f"{W}r")
        return runs[run_idx].find(f"{W}rPr")

    # ── set_run_highlight ──────────────────────────────────────────────────

    def test_set_run_highlight_sets_yellow(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_run_highlight("00000006", 0, "yellow"))
        assert result["para_id"] == "00000006"
        assert result["run_idx"] == 0
        assert result["color"] == "yellow"

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        assert rpr is not None
        hl = rpr.find(f"{W}highlight")
        assert hl is not None
        assert hl.get(f"{W}val") == "yellow"

    def test_set_run_highlight_raises_index_error_on_bad_run_idx(self, test_docx: Path) -> None:
        self._open(test_docx)
        with pytest.raises(IndexError):
            server._doc.set_run_highlight("00000001", 99, "yellow")

    # ── set_run_strikethrough ──────────────────────────────────────────────

    def test_set_run_strikethrough_single_sets_strike(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_run_strikethrough("00000006", 0, False))
        assert result["para_id"] == "00000006"
        assert result["run_idx"] == 0
        assert result["double"] is False

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        assert rpr is not None
        assert rpr.find(f"{W}strike") is not None

    def test_set_run_strikethrough_double_sets_dstrike_removes_strike(
        self, test_docx: Path
    ) -> None:  # noqa: E501
        self._open(test_docx)
        server.set_run_strikethrough("00000006", 0, False)
        result = json.loads(server.set_run_strikethrough("00000006", 0, True))
        assert result["double"] is True

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        assert rpr is not None
        assert rpr.find(f"{W}dstrike") is not None
        assert rpr.find(f"{W}strike") is None

    # ── set_run_superscript ────────────────────────────────────────────────

    def test_set_run_superscript_sets_vert_align(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_run_superscript("00000006", 1))
        assert result["para_id"] == "00000006"
        assert result["run_idx"] == 1
        assert result["valign"] == "superscript"

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[1].find(f"{W}rPr")
        assert rpr is not None
        va = rpr.find(f"{W}vertAlign")
        assert va is not None
        assert va.get(f"{W}val") == "superscript"

    # ── set_run_subscript ──────────────────────────────────────────────────

    def test_set_run_subscript_sets_vert_align(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_run_subscript("00000006", 2))
        assert result["para_id"] == "00000006"
        assert result["run_idx"] == 2
        assert result["valign"] == "subscript"

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[2].find(f"{W}rPr")
        assert rpr is not None
        va = rpr.find(f"{W}vertAlign")
        assert va is not None
        assert va.get(f"{W}val") == "subscript"

    # ── set_run_underline ──────────────────────────────────────────────────

    def test_set_run_underline_single_sets_u_val(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_run_underline("00000006", 0, "single"))
        assert result["para_id"] == "00000006"
        assert result["run_idx"] == 0
        assert result["style"] == "single"

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        assert rpr is not None
        u = rpr.find(f"{W}u")
        assert u is not None
        assert u.get(f"{W}val") == "single"

    def test_set_run_underline_double_sets_correct_val(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_run_underline("00000006", 0, "double"))
        assert result["style"] == "double"

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        u = rpr.find(f"{W}u")
        assert u is not None
        assert u.get(f"{W}val") == "double"

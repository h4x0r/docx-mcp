"""RED tests for run advanced tools: clear_run_formatting, set_run_language, set_text_case."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


class TestRunAdvanced:

    def _open(self, path: Path) -> None:
        server._doc = None
        server.open_document(str(path))

    def _get_run(self, test_docx, para_id, run_idx):
        self._open(test_docx)
        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, para_id)
        runs = para.findall(f"{W}r")
        return runs[run_idx]

    # ── clear_run_formatting ───────────────────────────────────────────────

    def test_clear_run_formatting_removes_rPr_children(self, test_docx: Path) -> None:
        self._open(test_docx)
        # First apply bold+italic to create rPr with children
        server._doc.set_run_font("00000006", 0, "Arial")
        server._doc.set_run_color("00000006", 0, "FF0000")

        result = json.loads(server.clear_run_formatting("00000006", 0))
        assert result["para_id"] == "00000006"
        assert result["run_idx"] == 0
        assert result["cleared"] is True

        # rPr should be absent or have no children
        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        if rpr is not None:
            assert len(list(rpr)) == 0

    def test_clear_run_formatting_out_of_range_raises(self, test_docx: Path) -> None:
        self._open(test_docx)
        with pytest.raises((ValueError, IndexError)):
            server._doc.clear_run_formatting("00000001", 99)

    # ── set_run_language ──────────────────────────────────────────────────

    def test_set_run_language_sets_lang_val(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_run_language("00000006", 0, "en-US"))
        assert result["para_id"] == "00000006"
        assert result["run_idx"] == 0
        assert result["language"] == "en-US"

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        assert rpr is not None
        lang = rpr.find(f"{W}lang")
        assert lang is not None
        assert lang.get(f"{W}val") == "en-US"

    def test_set_run_language_creates_rPr_if_absent(self, test_docx: Path) -> None:
        self._open(test_docx)
        # Use a paragraph where the run has no rPr
        # First clear any rPr on run 0 of 00000001
        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000001")
        runs = para.findall(f"{W}r")
        if runs:
            rpr = runs[0].find(f"{W}rPr")
            if rpr is not None:
                runs[0].remove(rpr)

        result = json.loads(server.set_run_language("00000001", 0, "fr-FR"))
        assert result["language"] == "fr-FR"

        root2 = doc._tree("word/document.xml")
        para2 = doc._find_para(root2, "00000001")
        runs2 = para2.findall(f"{W}r")
        rpr2 = runs2[0].find(f"{W}rPr")
        assert rpr2 is not None
        lang2 = rpr2.find(f"{W}lang")
        assert lang2 is not None
        assert lang2.get(f"{W}val") == "fr-FR"

    # ── set_text_case ─────────────────────────────────────────────────────

    def test_set_text_case_upper(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_text_case("00000006", 0, "upper"))
        assert result["para_id"] == "00000006"
        assert result["run_idx"] == 0
        assert result["case"] == "upper"

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        assert rpr is not None
        assert rpr.find(f"{W}caps") is not None
        assert rpr.find(f"{W}smallCaps") is None

    def test_set_text_case_small(self, test_docx: Path) -> None:
        self._open(test_docx)
        result = json.loads(server.set_text_case("00000006", 0, "small"))
        assert result["case"] == "small"

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        assert rpr is not None
        assert rpr.find(f"{W}smallCaps") is not None
        assert rpr.find(f"{W}caps") is None

    def test_set_text_case_none_removes_both(self, test_docx: Path) -> None:
        self._open(test_docx)
        # Set upper first, then clear
        server.set_text_case("00000006", 0, "upper")
        result = json.loads(server.set_text_case("00000006", 0, "none"))
        assert result["case"] == "none"

        doc = server._doc
        root = doc._tree("word/document.xml")
        para = doc._find_para(root, "00000006")
        runs = para.findall(f"{W}r")
        rpr = runs[0].find(f"{W}rPr")
        if rpr is not None:
            assert rpr.find(f"{W}caps") is None
            assert rpr.find(f"{W}smallCaps") is None

    def test_set_text_case_invalid_raises(self, test_docx: Path) -> None:
        self._open(test_docx)
        with pytest.raises(ValueError):
            server._doc.set_text_case("00000006", 0, "UPPER")

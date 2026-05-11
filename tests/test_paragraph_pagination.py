"""Tests for paragraph pagination controls: keep_with_next, keep_lines_together,
page_break_before, widow_control."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _open(path: Path) -> None:
    server._doc = None
    server.open_document(str(path))


def _root():
    return server._doc._tree("word/document.xml")


def _para(para_id: str):
    return server._doc._find_para(_root(), para_id)


def _ppr(para_id: str):
    p = _para(para_id)
    assert p is not None
    return p.find(f"{W}pPr")


class TestSetKeepWithNext:

    def test_set_keep_with_next_enabled(self, test_docx: Path) -> None:
        _open(test_docx)
        result = json.loads(server.set_keep_with_next("00000002", True))
        assert result["para_id"] == "00000002"
        assert result["enabled"] is True

        ppr = _ppr("00000002")
        assert ppr is not None
        assert ppr.find(f"{W}keepNext") is not None

    def test_set_keep_with_next_disabled(self, test_docx: Path) -> None:
        _open(test_docx)
        # Enable first, then disable
        server.set_keep_with_next("00000002", True)
        result = json.loads(server.set_keep_with_next("00000002", False))
        assert result["enabled"] is False

        ppr = _ppr("00000002")
        assert ppr is None or ppr.find(f"{W}keepNext") is None


class TestSetKeepLinesTogether:

    def test_set_keep_lines_together_enabled(self, test_docx: Path) -> None:
        _open(test_docx)
        result = json.loads(server.set_keep_lines_together("00000002", True))
        assert result["para_id"] == "00000002"
        assert result["enabled"] is True

        ppr = _ppr("00000002")
        assert ppr is not None
        assert ppr.find(f"{W}keepLines") is not None


class TestSetPageBreakBefore:

    def test_set_page_break_before_enabled(self, test_docx: Path) -> None:
        _open(test_docx)
        result = json.loads(server.set_page_break_before("00000002", True))
        assert result["para_id"] == "00000002"
        assert result["enabled"] is True

        ppr = _ppr("00000002")
        assert ppr is not None
        assert ppr.find(f"{W}pageBreakBefore") is not None


class TestSetWidowControl:

    def test_set_widow_control_enabled(self, test_docx: Path) -> None:
        _open(test_docx)
        result = json.loads(server.set_widow_control("00000002", True))
        assert result["para_id"] == "00000002"
        assert result["enabled"] is True

        ppr = _ppr("00000002")
        assert ppr is not None
        assert ppr.find(f"{W}widowControl") is not None


class TestPaginationNotFound:

    def test_pagination_not_found_raises(self, test_docx: Path) -> None:
        _open(test_docx)
        with pytest.raises(ValueError, match="not found"):
            server._doc.set_keep_with_next("DEADBEEF", True)

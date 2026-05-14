"""Security tests for search_text ReDoS protection."""
from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument


def _open_doc(tmp_path: Path) -> DocxDocument:
    from tests.conftest import _build_fixture
    path = tmp_path / "test.docx"
    _build_fixture(path)
    doc = DocxDocument(str(path))
    doc.open()
    return doc


def test_invalid_regex_raises_value_error(tmp_path: Path):
    """search_text with invalid regex raises ValueError."""
    doc = _open_doc(tmp_path)
    with pytest.raises(ValueError, match="regex"):
        doc.search_text("[unclosed", regex=True)
    doc.close()


def test_valid_regex_works(tmp_path: Path):
    """search_text with a valid regex returns results."""
    doc = _open_doc(tmp_path)
    results = doc.search_text(r"\w+", regex=True)
    assert isinstance(results, list)
    doc.close()


def test_regex_no_match_returns_empty(tmp_path: Path):
    """search_text with regex that doesn't match returns empty list."""
    doc = _open_doc(tmp_path)
    results = doc.search_text(r"XYZZY_NOMATCH_12345", regex=True)
    assert results == []
    doc.close()

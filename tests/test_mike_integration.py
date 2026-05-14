"""Tests for mike integration: get_body_text, replace_text return, hyperlink, NBSP."""
from __future__ import annotations
import json
import pytest
from docx_mcp import server
from docx_mcp.document.tracks import _flatten_para
from lxml import etree


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"


def test_flatten_para_includes_hyperlink_text(mike_corpus_docx):
    """_flatten_para must include text inside w:hyperlink runs."""
    server.open_document(str(mike_corpus_docx))
    doc = server._doc._require("word/document.xml")
    # Find paragraph 00000101 which contains a w:hyperlink with "thirty calendar days"
    para = None
    for p in doc.iter(f"{W}p"):
        if p.get(f"{W14}paraId") == "00000101":
            para = p
            break
    assert para is not None, "paragraph 00000101 not found"
    slots = _flatten_para(para)
    text = "".join(s.char for s in slots)
    assert "thirty calendar days" in text, (
        f"hyperlink text not found in _flatten_para output; got: {text!r}"
    )


def test_get_body_text_returns_body_string(mike_corpus_docx):
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.get_body_text())
    assert "body" in result
    assert isinstance(result["body"], str)
    assert len(result["body"]) > 0


def test_get_body_text_includes_hyperlink_text(mike_corpus_docx):
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.get_body_text())
    # Hyperlink text must appear in body (not invisible)
    assert "thirty calendar days" in result["body"]


def test_get_body_text_accepted_view_excludes_deleted(mike_corpus_docx):
    """Accepted view: w:del text must NOT appear, w:ins text MUST appear."""
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.get_body_text())
    assert "one hundred" not in result["body"]   # inside w:del → excluded
    assert "two hundred" in result["body"]         # inside w:ins → included


def test_get_body_text_returns_footnotes(mike_corpus_docx):
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.get_body_text())
    assert "footnotes" in result
    # mike_corpus uses the same _FOOTNOTES_XML as test_docx
    # Check the footnotes XML to see what text it contains
    assert isinstance(result["footnotes"], str)


def test_get_body_text_real_contract(tmp_path):
    """Smoke test against real externally-sourced document."""
    import os
    fixture = os.path.join(os.path.dirname(__file__), "fixtures", "real_contract.docx")
    if not os.path.exists(fixture):
        pytest.skip("real fixture not present")
    server.open_document(fixture)
    result = json.loads(server.get_body_text())
    assert len(result["body"]) > 100

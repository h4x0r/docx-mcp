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

"""Tests for markdown-to-DOCX conversion."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import W14, DocxDocument, W
from docx_mcp.markdown import MarkdownConverter


@pytest.fixture()
def blank_doc(tmp_path: Path) -> DocxDocument:
    out = tmp_path / "test.docx"
    doc = DocxDocument.create(str(out))
    return doc


class TestHeadings:
    def test_h1(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "# Title")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        # Should have replaced the blank paragraph with the heading
        assert len(paras) == 1
        ppr = paras[0].find(f"{W}pPr")
        style = ppr.find(f"{W}pStyle")
        assert style.get(f"{W}val") == "Heading1"
        assert blank_doc._text(paras[0]) == "Title"

    def test_h1_through_h6(self, blank_doc: DocxDocument):
        md = "\n\n".join(f"{'#' * i} Heading {i}" for i in range(1, 7))
        MarkdownConverter.convert(blank_doc, md)
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 6
        for i, para in enumerate(paras, 1):
            style = para.find(f"{W}pPr/{W}pStyle")
            assert style.get(f"{W}val") == f"Heading{i}"

    def test_all_paragraphs_have_para_ids(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "# H1\n\nParagraph\n\n## H2")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        for para in body.findall(f"{W}p"):
            pid = para.get(f"{W14}paraId")
            assert pid is not None
            assert len(pid) == 8
            assert int(pid, 16) < 0x80000000


class TestParagraphs:
    def test_simple_paragraph(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "Hello world")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 1
        assert blank_doc._text(paras[0]) == "Hello world"

    def test_multiple_paragraphs(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "First\n\nSecond\n\nThird")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 3

    def test_empty_input(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        paras = body.findall(f"{W}p")
        assert len(paras) == 0


class TestInlineFormatting:
    def test_bold(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "**bold**")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        runs = body.findall(f".//{W}r")
        # Find the run with bold text
        bold_runs = [r for r in runs if r.find(f"{W}rPr/{W}b") is not None]
        assert len(bold_runs) >= 1
        assert blank_doc._text(bold_runs[0]) == "bold"

    def test_italic(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "*italic*")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        runs = body.findall(f".//{W}r")
        italic_runs = [r for r in runs if r.find(f"{W}rPr/{W}i") is not None]
        assert len(italic_runs) >= 1
        assert blank_doc._text(italic_runs[0]) == "italic"

    def test_strikethrough(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "~~struck~~")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        runs = body.findall(f".//{W}r")
        strike_runs = [r for r in runs if r.find(f"{W}rPr/{W}strike") is not None]
        assert len(strike_runs) >= 1

    def test_bold_italic_combo(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "***both***")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        runs = body.findall(f".//{W}r")
        both = [
            r for r in runs
            if r.find(f"{W}rPr/{W}b") is not None
            and r.find(f"{W}rPr/{W}i") is not None
        ]
        assert len(both) >= 1

    def test_inline_code(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, "Use `print()` here")
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        runs = body.findall(f".//{W}r")
        code_runs = [r for r in runs if r.find(f"{W}rPr/{W}rFonts") is not None]
        assert len(code_runs) >= 1
        # Should have Courier New font
        font = code_runs[0].find(f"{W}rPr/{W}rFonts")
        assert font.get(f"{W}ascii") == "Courier New"

    def test_smart_typography_applied(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, '"quoted"')
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        text = blank_doc._text(body)
        assert "\u201c" in text  # left double quote
        assert "\u201d" in text  # right double quote

    def test_smart_typography_not_in_code(self, blank_doc: DocxDocument):
        MarkdownConverter.convert(blank_doc, '`"not smart"`')
        body = blank_doc._trees["word/document.xml"].find(f"{W}body")
        runs = body.findall(f".//{W}r")
        code_runs = [r for r in runs if r.find(f"{W}rPr/{W}rFonts") is not None]
        assert any('"not smart"' in (blank_doc._text(r) or "") for r in code_runs)

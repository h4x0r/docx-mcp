"""Tests for XPath query escape hatch — xpath_query tool."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_mcp import server
from docx_mcp.document.errors import DocxMcpError, ErrCode


def _j(result: str) -> dict | list:
    return json.loads(result)


def _open(path: Path) -> None:
    server.open_document(str(path))


def _require_doc():
    assert server._doc is not None, "No document open"
    return server._doc


# ── helpers ─────────────────────────────────────────────────────────────────


def _make_doc_with_paragraphs(tmp_path: Path, n: int) -> Path:
    """Build a minimal DOCX with n plain body paragraphs."""
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    W14 = "http://schemas.microsoft.com/office/word/2010/wordml"

    def para(text: str, para_id: str) -> str:
        return (
            f'<w:p xmlns:w="{W}" xmlns:w14="{W14}" '
            f'w14:paraId="{para_id}" w14:textId="1">'
            f"<w:r><w:t>{text}</w:t></w:r></w:p>"
        )

    body_paras = "".join(para(f"Para {i}", f"{i:08X}") for i in range(1, n + 1))

    document_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}" xmlns:w14="{W14}">
  <w:body>
    {body_paras}
    <w:sectPr/>
  </w:body>
</w:document>"""

    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    top_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"""

    import zipfile

    path = tmp_path / "test.docx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", top_rels)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/document.xml", document_xml)
    return path


def _make_doc_with_heading(tmp_path: Path) -> Path:
    """Build a minimal DOCX with a Heading1 paragraph and a body paragraph."""
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    W14 = "http://schemas.microsoft.com/office/word/2010/wordml"

    document_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}" xmlns:w14="{W14}">
  <w:body>
    <w:p xmlns:w="{W}" xmlns:w14="{W14}" w14:paraId="00000001" w14:textId="1">
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Introduction</w:t></w:r>
    </w:p>
    <w:p xmlns:w="{W}" xmlns:w14="{W14}" w14:paraId="00000002" w14:textId="1">
      <w:r><w:t>Body text here</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>"""

    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

    top_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
</Relationships>"""

    styles_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="{W}">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>"""

    import zipfile

    path = tmp_path / "heading.docx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", top_rels)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/styles.xml", styles_xml)
    return path


# ── TestXpathQuery ───────────────────────────────────────────────────────────


class TestXpathQuery:
    def test_finds_all_paragraphs(self, tmp_path: Path):
        path = _make_doc_with_paragraphs(tmp_path, 3)
        _open(path)
        result = _require_doc().xpath_query("//w:p")
        assert result["count"] >= 1
        assert result["returned"] >= 1
        assert isinstance(result["results"][0], str)
        assert "<w:p" in result["results"][0]

    def test_finds_by_style(self, tmp_path: Path):
        path = _make_doc_with_heading(tmp_path)
        _open(path)
        result = _require_doc().xpath_query("//w:p[w:pPr/w:pStyle/@w:val='Heading1']")
        assert result["count"] >= 1
        assert result["returned"] >= 1
        assert "Heading1" in result["results"][0]

    def test_text_content_xpath(self, tmp_path: Path):
        path = _make_doc_with_paragraphs(tmp_path, 2)
        _open(path)
        result = _require_doc().xpath_query("//w:t/text()")
        assert result["count"] >= 1
        # text nodes should be plain strings, not XML
        for r in result["results"]:
            assert not r.strip().startswith("<"), "Expected plain string, got XML element"

    def test_invalid_xpath_raises_xpath_error(self, tmp_path: Path):
        path = _make_doc_with_paragraphs(tmp_path, 1)
        _open(path)
        with pytest.raises(DocxMcpError) as exc_info:
            _require_doc().xpath_query("//[invalid")
        assert exc_info.value.code == ErrCode.XPATH_ERROR

    def test_part_not_found_raises(self, tmp_path: Path):
        path = _make_doc_with_paragraphs(tmp_path, 1)
        _open(path)
        with pytest.raises(DocxMcpError) as exc_info:
            _require_doc().xpath_query("//w:p", part="word/nonexistent.xml")
        assert exc_info.value.code == ErrCode.PART_NOT_FOUND

    def test_results_capped_at_50(self, tmp_path: Path):
        path = _make_doc_with_paragraphs(tmp_path, 60)
        _open(path)
        result = _require_doc().xpath_query("//w:p")
        assert result["count"] >= 60
        assert result["returned"] == 50
        assert len(result["results"]) == 50

    def test_returns_correct_keys(self, tmp_path: Path):
        path = _make_doc_with_paragraphs(tmp_path, 1)
        _open(path)
        result = _require_doc().xpath_query("//w:p")
        assert set(result.keys()) == {"xpath", "part", "count", "returned", "results"}

    def test_styles_part_queryable(self, tmp_path: Path):
        path = _make_doc_with_heading(tmp_path)
        _open(path)
        result = _require_doc().xpath_query("//w:style", part="word/styles.xml")
        assert result["count"] >= 1
        assert result["part"] == "word/styles.xml"

    def test_count_vs_returned_distinction(self, tmp_path: Path):
        path = _make_doc_with_paragraphs(tmp_path, 60)
        _open(path)
        result = _require_doc().xpath_query("//w:p")
        assert result["count"] > result["returned"]
        assert result["returned"] == 50

    def test_xpath_and_part_in_result(self, tmp_path: Path):
        path = _make_doc_with_paragraphs(tmp_path, 1)
        _open(path)
        result = _require_doc().xpath_query("//w:p", part="word/document.xml")
        assert result["xpath"] == "//w:p"
        assert result["part"] == "word/document.xml"


# ── TestXpathQueryServer ─────────────────────────────────────────────────────


class TestXpathQueryServer:
    def test_server_tool_returns_json(self, tmp_path: Path):
        path = _make_doc_with_paragraphs(tmp_path, 2)
        server.open_document(str(path))
        raw = server.xpath_query("//w:p")
        result = json.loads(raw)
        assert isinstance(result, dict)
        assert result["count"] >= 1
        assert "results" in result

    def test_server_tool_invalid_xpath(self, tmp_path: Path):
        path = _make_doc_with_paragraphs(tmp_path, 1)
        server.open_document(str(path))
        with pytest.raises(DocxMcpError) as exc_info:
            server.xpath_query("//[bad")
        assert exc_info.value.code == ErrCode.XPATH_ERROR

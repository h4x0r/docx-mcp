"""Tests for ThemeMixin and CaptionMixin."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp import server
from docx_mcp.document import DocxDocument

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

_THEME_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    </a:clrScheme>
  </a:themeElements>
</a:theme>"""

_MINIMAL_DOCX_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="00000001" w14:textId="77777777">
      <w:r><w:t>First paragraph.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000002" w14:textId="77777777">
      <w:r><w:t>Second paragraph.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="00000003" w14:textId="77777777">
      <w:pPr><w:pStyle w:val="Caption"/></w:pPr>
      <w:r><w:t>Figure 1: Existing caption</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""


def _build_docx(path: Path, include_theme: bool = True) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""
    top_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""
    doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
</Relationships>"""
    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Caption">
    <w:name w:val="caption"/>
  </w:style>
</w:styles>"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types.strip())
        zf.writestr("_rels/.rels", top_rels.strip())
        zf.writestr("word/document.xml", _MINIMAL_DOCX_XML.strip())
        zf.writestr("word/_rels/document.xml.rels", doc_rels.strip())
        zf.writestr("word/styles.xml", styles.strip())
        if include_theme:
            zf.writestr("word/theme/theme1.xml", _THEME_XML)


def _open_doc(path: Path) -> DocxDocument:
    server._doc = None
    server.open_document(str(path))
    return server._doc


class TestThemeMixin:
    def test_get_theme_colors_returns_dict_with_slots(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc._trees["word/theme/theme1.xml"] = etree.fromstring(_THEME_XML)
        colors = doc.get_theme_colors()
        assert isinstance(colors, dict)
        assert colors["dk1"] == "000000"
        assert colors["lt1"] == "FFFFFF"
        assert colors["accent1"] == "4472C4"

    def test_get_theme_colors_returns_empty_when_no_theme(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p, include_theme=False)
        doc = _open_doc(p)
        doc._trees.pop("word/theme/theme1.xml", None)
        colors = doc.get_theme_colors()
        assert colors == {}

    def test_set_theme_color_updates_existing_slot(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc._trees["word/theme/theme1.xml"] = etree.fromstring(_THEME_XML)
        result = doc.set_theme_color("accent1", "FF0000")
        assert result == {"slot": "accent1", "hex_color": "FF0000"}
        colors = doc.get_theme_colors()
        assert colors["accent1"] == "FF0000"

    def test_set_theme_color_marks_dirty(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc._trees["word/theme/theme1.xml"] = etree.fromstring(_THEME_XML)
        doc.set_theme_color("dk1", "123456")
        assert "word/theme/theme1.xml" in doc._modified

    def test_set_theme_color_raises_on_unknown_slot(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc._trees["word/theme/theme1.xml"] = etree.fromstring(_THEME_XML)
        with pytest.raises(ValueError, match="unknown slot"):
            doc.set_theme_color("bogus", "AABBCC")

    def test_set_theme_color_uses_srgbclr(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc._trees["word/theme/theme1.xml"] = etree.fromstring(_THEME_XML)
        doc.set_theme_color("lt1", "CCCCCC")
        theme = doc._trees["word/theme/theme1.xml"]
        ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        clr_scheme = theme.find(f"{{{ns}}}themeElements/{{{ns}}}clrScheme")
        slot_el = clr_scheme.find(f"{{{ns}}}lt1")
        srgb = slot_el.find(f"{{{ns}}}srgbClr")
        assert srgb is not None
        assert srgb.get("val") == "CCCCCC"


class TestCaptionMixin:
    def test_insert_caption_inserts_after_target(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.insert_caption("00000002", "A new figure")
        assert isinstance(result, dict)
        root = doc._tree("word/document.xml")
        body = root.find(f"{W}body")
        paras = [el for el in body if el.tag == f"{W}p"]
        ids = [el.get(f"{W14}paraId") for el in paras]
        idx_target = ids.index("00000002")
        assert ids[idx_target + 1] == result["para_id"]

    def test_insert_caption_uses_caption_style(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.insert_caption("00000001", "Test fig")
        root = doc._tree("word/document.xml")
        new_para = doc._find_para(root, result["para_id"])
        ppr = new_para.find(f"{W}pPr")
        assert ppr is not None
        pstyle = ppr.find(f"{W}pStyle")
        assert pstyle is not None
        assert pstyle.get(f"{W}val") == "Caption"

    def test_insert_caption_auto_increments_seq_num(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        r1 = doc.insert_caption("00000001", "First new caption")
        r2 = doc.insert_caption("00000002", "Second new caption")
        assert r1["seq_num"] == 2
        assert r2["seq_num"] == 3

    def test_insert_caption_raises_on_bad_para_id(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        with pytest.raises(ValueError, match="not found"):
            doc.insert_caption("DEADBEEF", "caption text")

    def test_insert_caption_default_label_is_figure(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.insert_caption("00000001", "something")
        assert result["label"] == "Figure"

    def test_insert_caption_custom_label(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.insert_caption("00000001", "data table", label="Table")
        assert result["label"] == "Table"
        root = doc._tree("word/document.xml")
        new_para = doc._find_para(root, result["para_id"])
        texts = [t.text for t in new_para.iter(f"{W}t") if t.text]
        assert any("Table" in t for t in texts)

"""Tests for generate_tof and generate_tot methods on TocMixin."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp import server
from docx_mcp.document import DocxDocument

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"

_DOC_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="A0000001" w14:textId="77777777">
      <w:r><w:t>Intro paragraph.</w:t></w:r>
    </w:p>
    <w:p w14:paraId="A0000002" w14:textId="77777777">
      <w:pPr><w:pStyle w:val="Caption"/></w:pPr>
      <w:r><w:t>Figure 1: First figure caption</w:t></w:r>
    </w:p>
    <w:p w14:paraId="A0000003" w14:textId="77777777">
      <w:pPr><w:pStyle w:val="Caption"/></w:pPr>
      <w:r><w:t>Figure 2: Second figure caption</w:t></w:r>
    </w:p>
    <w:p w14:paraId="A0000004" w14:textId="77777777">
      <w:pPr><w:pStyle w:val="Caption"/></w:pPr>
      <w:r><w:t>Table 1: First table caption</w:t></w:r>
    </w:p>
    <w:p w14:paraId="A0000005" w14:textId="77777777">
      <w:pPr><w:pStyle w:val="caption"/></w:pPr>
      <w:r><w:t>Table 2: Second table caption (lowercase style)</w:t></w:r>
    </w:p>
    <w:p w14:paraId="A0000006" w14:textId="77777777">
      <w:r><w:t>End paragraph.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""


def _build_docx(path: Path) -> None:
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
        zf.writestr("word/document.xml", _DOC_XML.strip())
        zf.writestr("word/_rels/document.xml.rels", doc_rels.strip())
        zf.writestr("word/styles.xml", styles.strip())


def _open_doc(path: Path) -> DocxDocument:
    server._doc = None
    server.open_document(str(path))
    return server._doc


class TestGenerateTof:

    def test_inserts_field_with_figure_instruction(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc.generate_tof("A0000001")
        root = doc._tree("word/document.xml")
        instr_texts = [it.text for it in root.iter(f"{W}instrText") if it.text]
        assert any('\\c "Figure"' in t for t in instr_texts)

    def test_collects_figure_captions_as_entries(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.generate_tof("A0000001")
        assert result["entry_count"] == 2

    def test_inserts_after_target_paragraph(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc.generate_tof("A0000001")
        root = doc._tree("word/document.xml")
        body = root.find(f"{W}body")
        children = list(body)
        target_idx = next(
            i for i, el in enumerate(children)
            if el.get(f"{W14}paraId") == "A0000001"
        )
        inserted_texts = []
        for el in children[target_idx + 1:]:
            inserted_texts.extend(it.text for it in el.iter(f"{W}instrText") if it.text)
        assert any('\\c "Figure"' in t for t in inserted_texts)
        pre_target_texts = []
        for el in children[:target_idx]:
            pre_target_texts.extend(it.text for it in el.iter(f"{W}instrText") if it.text)
        assert not any('\\c "Figure"' in t for t in pre_target_texts)

    def test_raises_value_error_on_bad_para_id(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        with pytest.raises(ValueError):
            doc.generate_tof("DEADBEEF")

    def test_returns_correct_dict_structure(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.generate_tof("A0000001", title="List of Figures")
        assert result["para_id"] == "A0000001"
        assert result["title"] == "List of Figures"
        assert isinstance(result["entry_count"], int)

    def test_does_not_insert_at_start_of_doc(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc.generate_tof("A0000006")
        root = doc._tree("word/document.xml")
        body = root.find(f"{W}body")
        first_child = list(body)[0]
        instr_texts = [it.text for it in first_child.iter(f"{W}instrText") if it.text]
        assert not any('\\c "Figure"' in t for t in instr_texts)

    def test_entries_styled_toc1(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc.generate_tof("A0000001")
        root = doc._tree("word/document.xml")
        body = root.find(f"{W}body")
        toc1_paras = []
        for child in body:
            pPr = child.find(f"{W}pPr")
            if pPr is not None:
                pStyle = pPr.find(f"{W}pStyle")
                if pStyle is not None and pStyle.get(f"{W}val") == "TOC1":
                    toc1_paras.append(child)
        assert len(toc1_paras) == 2


class TestGenerateTot:

    def test_inserts_field_with_table_instruction(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc.generate_tot("A0000001")
        root = doc._tree("word/document.xml")
        instr_texts = [it.text for it in root.iter(f"{W}instrText") if it.text]
        assert any('\\c "Table"' in t for t in instr_texts)

    def test_collects_table_captions_as_entries(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.generate_tot("A0000001")
        assert result["entry_count"] == 2

    def test_returns_correct_dict_structure(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.generate_tot("A0000001", title="List of Tables")
        assert result["para_id"] == "A0000001"
        assert result["title"] == "List of Tables"
        assert isinstance(result["entry_count"], int)

    def test_does_not_collect_figure_captions(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.generate_tot("A0000001")
        assert result["entry_count"] == 2

    def test_raises_value_error_on_bad_para_id(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        with pytest.raises(ValueError):
            doc.generate_tot("BADID999")

    def test_inserts_after_target_paragraph(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        doc.generate_tot("A0000006")
        root = doc._tree("word/document.xml")
        body = root.find(f"{W}body")
        children = list(body)
        target_idx = next(
            i for i, el in enumerate(children)
            if el.get(f"{W14}paraId") == "A0000006"
        )
        inserted_texts = []
        for el in children[target_idx + 1:]:
            inserted_texts.extend(it.text for it in el.iter(f"{W}instrText") if it.text)
        assert any('\\c "Table"' in t for t in inserted_texts)
        pre_target_texts = []
        for el in children[:target_idx]:
            pre_target_texts.extend(it.text for it in el.iter(f"{W}instrText") if it.text)
        assert not any('\\c "Table"' in t for t in pre_target_texts)

    def test_case_insensitive_caption_style_match(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docx"
        _build_docx(p)
        doc = _open_doc(p)
        result = doc.generate_tot("A0000001")
        assert result["entry_count"] == 2

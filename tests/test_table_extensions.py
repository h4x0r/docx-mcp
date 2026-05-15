"""Tests for Task #26: delete_header/footer + table extension methods."""

from __future__ import annotations

import zipfile

import pytest

from docx_mcp.document import W14, DocxDocument, W

# ── helpers ──────────────────────────────────────────────────────────────────


def _make_doc_with_table(tmp_path, rows=2, cols=2):
    doc = DocxDocument.create(str(tmp_path / "test.docx"))
    tree = doc._tree("word/document.xml")
    paras = tree.findall(f".//{W}p")
    para_id = paras[0].get(f"{W14}paraId")
    doc.add_table(para_id, rows, cols)
    return doc, 0


def _doc_with_header_footer(tmp_path):
    """Build a DOCX that has a header and footer referenced in sectPr."""
    path = tmp_path / "hf.docx"

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

    top_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>  # noqa: E501
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>  # noqa: E501
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>  # noqa: E501
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>  # noqa: E501
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>  # noqa: E501
</Relationships>"""

    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="11111111" w14:textId="77777777">
      <w:r><w:t>Hello</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId3"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
      <w:footerReference w:type="default" r:id="rId4"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
    </w:sectPr>
  </w:body>
</w:document>"""

    header_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Header text</w:t></w:r></w:p>
</w:hdr>"""

    footer_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Footer text</w:t></w:r></w:p>
</w:ftr>"""

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"""

    settings_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"""

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types.strip())
        zf.writestr("_rels/.rels", top_rels.strip())
        zf.writestr("word/document.xml", document_xml.strip())
        zf.writestr("word/_rels/document.xml.rels", doc_rels.strip())
        zf.writestr("word/header1.xml", header_xml.strip())
        zf.writestr("word/footer1.xml", footer_xml.strip())
        zf.writestr("word/styles.xml", styles_xml.strip())
        zf.writestr("word/settings.xml", settings_xml.strip())

    doc = DocxDocument(str(path))
    doc.open()
    return doc


# ── delete_header / delete_footer ────────────────────────────────────────────


class TestDeleteHeader:
    def test_delete_header_returns_deleted_location(self, tmp_path):
        doc = _doc_with_header_footer(tmp_path)
        result = doc.delete_header(location="default")
        assert result == {"deleted": "default"}

    def test_delete_header_removes_headerReference_from_sectPr(self, tmp_path):
        doc = _doc_with_header_footer(tmp_path)
        doc.delete_header(location="default")
        tree = doc._tree("word/document.xml")
        sect = tree.find(f".//{W}sectPr")
        refs = sect.findall(f"{W}headerReference")
        assert refs == []

    def test_delete_header_not_found_returns_none(self, tmp_path):
        doc = _doc_with_header_footer(tmp_path)
        result = doc.delete_header(location="first")
        assert result == {"deleted": None}

    def test_delete_footer_returns_deleted_location(self, tmp_path):
        doc = _doc_with_header_footer(tmp_path)
        result = doc.delete_footer(location="default")
        assert result == {"deleted": "default"}

    def test_delete_footer_removes_footerReference_from_sectPr(self, tmp_path):
        doc = _doc_with_header_footer(tmp_path)
        doc.delete_footer(location="default")
        tree = doc._tree("word/document.xml")
        sect = tree.find(f".//{W}sectPr")
        refs = sect.findall(f"{W}footerReference")
        assert refs == []

    def test_delete_footer_not_found_returns_none(self, tmp_path):
        doc = _doc_with_header_footer(tmp_path)
        result = doc.delete_footer(location="even")
        assert result == {"deleted": None}


# ── delete_table ─────────────────────────────────────────────────────────────


class TestDeleteTable:
    def test_delete_table_removes_tbl_element(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        tree = doc._tree("word/document.xml")
        assert len(list(tree.iter(f"{W}tbl"))) == 1
        doc.delete_table(0)
        assert len(list(tree.iter(f"{W}tbl"))) == 0

    def test_delete_table_returns_deleted_idx(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        result = doc.delete_table(0)
        assert result == {"deleted": 0}

    def test_delete_table_out_of_range_raises(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        with pytest.raises(IndexError):
            doc.delete_table(5)


# ── add_column_to_table ───────────────────────────────────────────────────────


class TestAddColumnToTable:
    def test_add_column_increases_col_count(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path, rows=2, cols=2)
        result = doc.add_column_to_table(idx, header_text="NewCol")
        assert result["columns"] == 3

    def test_add_column_returns_correct_keys(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path, rows=2, cols=2)
        result = doc.add_column_to_table(idx, header_text="H")
        assert result["table_idx"] == idx
        assert "columns" in result

    def test_add_column_first_row_has_header_text(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path, rows=2, cols=2)
        doc.add_column_to_table(idx, header_text="MyHeader")
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        rows = tbl.findall(f"{W}tr")
        first_row_cells = rows[0].findall(f"{W}tc")
        last_cell = first_row_cells[-1]
        texts = [t.text or "" for t in last_cell.iter(f"{W}t")]
        assert "MyHeader" in texts

    def test_add_column_all_rows_get_new_cell(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path, rows=3, cols=2)
        doc.add_column_to_table(idx)
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        for tr in tbl.findall(f"{W}tr"):
            assert len(tr.findall(f"{W}tc")) == 3


# ── delete_column_from_table ─────────────────────────────────────────────────


class TestDeleteColumnFromTable:
    def test_delete_column_decreases_col_count(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path, rows=2, cols=3)
        result = doc.delete_column_from_table(idx, col_idx=1)
        assert result == {"table_idx": idx, "col_idx": 1}
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        for tr in tbl.findall(f"{W}tr"):
            assert len(tr.findall(f"{W}tc")) == 2

    def test_delete_column_out_of_range_raises(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path, rows=2, cols=2)
        with pytest.raises(IndexError):
            doc.delete_column_from_table(idx, col_idx=5)


# ── set_cell_width ────────────────────────────────────────────────────────────


class TestSetCellWidth:
    def test_set_cell_width_returns_dxa(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        result = doc.set_cell_width(idx, 0, 0, width_mm=25.4)
        assert result["width_dxa"] == 1440
        assert result["table_idx"] == idx
        assert result["row_idx"] == 0
        assert result["col_idx"] == 0

    def test_set_cell_width_sets_tcW_element(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        doc.set_cell_width(idx, 0, 0, width_mm=25.4)
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        tc = tbl.findall(f"{W}tr")[0].findall(f"{W}tc")[0]
        tc_pr = tc.find(f"{W}tcPr")
        assert tc_pr is not None
        tc_w = tc_pr.find(f"{W}tcW")
        assert tc_w is not None
        assert tc_w.get(f"{W}w") == "1440"
        assert tc_w.get(f"{W}type") == "dxa"


# ── set_cell_vertical_alignment ───────────────────────────────────────────────


class TestSetCellVerticalAlignment:
    def test_set_valign_center(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        result = doc.set_cell_vertical_alignment(idx, 0, 0, "center")
        assert result["alignment"] == "center"
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        tc = tbl.findall(f"{W}tr")[0].findall(f"{W}tc")[0]
        tc_pr = tc.find(f"{W}tcPr")
        v_align = tc_pr.find(f"{W}vAlign")
        assert v_align is not None
        assert v_align.get(f"{W}val") == "center"

    def test_set_valign_returns_correct_keys(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        result = doc.set_cell_vertical_alignment(idx, 0, 1, "bottom")
        assert result == {"table_idx": idx, "row_idx": 0, "col_idx": 1, "alignment": "bottom"}


# ── set_row_height ────────────────────────────────────────────────────────────


class TestSetRowHeight:
    def test_set_row_height_dxa_value(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        result = doc.set_row_height(idx, 0, height_mm=25.4, rule="exact")
        assert result["height_dxa"] == 1440
        assert result["table_idx"] == idx
        assert result["row_idx"] == 0

    def test_set_row_height_sets_trHeight_element(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        doc.set_row_height(idx, 0, height_mm=25.4, rule="atLeast")
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        tr = tbl.findall(f"{W}tr")[0]
        tr_pr = tr.find(f"{W}trPr")
        assert tr_pr is not None
        tr_height = tr_pr.find(f"{W}trHeight")
        assert tr_height is not None
        assert tr_height.get(f"{W}val") == "1440"
        assert tr_height.get(f"{W}hRule") == "atLeast"


# ── set_table_alignment ───────────────────────────────────────────────────────


class TestSetTableAlignment:
    def test_set_table_alignment_center(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        result = doc.set_table_alignment(idx, "center")
        assert result == {"table_idx": idx, "alignment": "center"}
        tree = doc._tree("word/document.xml")
        tbl = list(tree.iter(f"{W}tbl"))[0]
        tbl_pr = tbl.find(f"{W}tblPr")
        jc = tbl_pr.find(f"{W}jc")
        assert jc is not None
        assert jc.get(f"{W}val") == "center"

    def test_set_table_alignment_right(self, tmp_path):
        doc, idx = _make_doc_with_table(tmp_path)
        result = doc.set_table_alignment(idx, "right")
        assert result["alignment"] == "right"

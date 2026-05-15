"""Tests for P9.2: set_image_border — image border via a:ln in pic:spPr."""

from __future__ import annotations

import pytest

from docx_mcp.document import DocxDocument

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"


def _open(test_docx):
    doc = DocxDocument(str(test_docx))
    doc.open()
    return doc


def _get_rId(doc):
    images = doc.get_images()
    assert images, "fixture must have at least one image"
    return images[0]["rId"]


def _sp_pr(doc, rId):
    """Return the pic:spPr element for the given rId."""
    from docx_mcp.document.base import A as BA
    from docx_mcp.document.base import R
    tree = doc._tree("word/document.xml")
    for blip in tree.iter(f"{BA}blip"):
        if blip.get(f"{R}embed") == rId:
            pic_el = blip.getparent()
            while pic_el is not None and pic_el.tag != f"{PIC}pic":
                pic_el = pic_el.getparent()
            if pic_el is not None:
                return pic_el.find(f"{PIC}spPr")
    return None


class TestSetImageBorder:
    def test_set_border_creates_ln_element(self, test_docx):
        doc = _open(test_docx)
        rId = _get_rId(doc)
        doc.set_image_border(rId, 1.0)
        sp = _sp_pr(doc, rId)
        assert sp is not None
        ln = sp.find(f"{A}ln")
        assert ln is not None
        assert ln.get("w") == str(round(1.0 * 12700))

    def test_set_border_sets_color(self, test_docx):
        doc = _open(test_docx)
        rId = _get_rId(doc)
        doc.set_image_border(rId, 2.0, color="FF0000")
        sp = _sp_pr(doc, rId)
        ln = sp.find(f"{A}ln")
        assert ln is not None
        solid = ln.find(f"{A}solidFill")
        assert solid is not None
        clr = solid.find(f"{A}srgbClr")
        assert clr is not None
        assert clr.get("val") == "FF0000"

    def test_set_border_default_color_is_black(self, test_docx):
        doc = _open(test_docx)
        rId = _get_rId(doc)
        doc.set_image_border(rId, 1.0)
        sp = _sp_pr(doc, rId)
        ln = sp.find(f"{A}ln")
        solid = ln.find(f"{A}solidFill")
        clr = solid.find(f"{A}srgbClr")
        assert clr.get("val") == "000000"

    def test_set_border_zero_removes_ln(self, test_docx):
        doc = _open(test_docx)
        rId = _get_rId(doc)
        doc.set_image_border(rId, 1.0)
        doc.set_image_border(rId, 0)
        sp = _sp_pr(doc, rId)
        assert sp is not None
        assert sp.find(f"{A}ln") is None

    def test_returns_correct_dict(self, test_docx):
        doc = _open(test_docx)
        rId = _get_rId(doc)
        result = doc.set_image_border(rId, 1.5, "AABBCC")
        assert result == {"rId": rId, "border_pt": 1.5, "color": "AABBCC"}

    def test_zero_border_returns_empty_color(self, test_docx):
        doc = _open(test_docx)
        rId = _get_rId(doc)
        result = doc.set_image_border(rId, 0)
        assert result == {"rId": rId, "border_pt": 0, "color": ""}

    def test_rId_not_found_raises(self, test_docx):
        doc = _open(test_docx)
        with pytest.raises(ValueError, match="not found"):
            doc.set_image_border("rId999", 1.0)

    def test_marks_document_dirty(self, test_docx):
        doc = _open(test_docx)
        rId = _get_rId(doc)
        doc._modified.discard("word/document.xml")
        doc.set_image_border(rId, 1.0)
        assert "word/document.xml" in doc._modified

"""Tests for insert_floating_image — wp:anchor floating image support."""

from __future__ import annotations

import pytest

from docx_mcp.document import W14, WP, DocxDocument, W

_TINY_PNG = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"
    b"\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02"
    b"\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx"
    b"\x9cc\xf8\x0f\x00\x00\x01\x01\x00\x05\x18\xd8N"
    b"\x00\x00\x00\x00IEND\xaeB`\x82"
)


@pytest.fixture
def sample_png(tmp_path):
    p = tmp_path / "sample.png"
    p.write_bytes(_TINY_PNG)
    return p


class TestFloatingImage:
    def test_insert_floating_creates_anchor(self, tmp_path, sample_png):
        """insert_floating_image creates a wp:anchor element (not wp:inline)."""
        doc = DocxDocument.create(str(tmp_path / "test.docx"))
        tree = doc._tree("word/document.xml")
        paras = tree.findall(f".//{W}p")
        para_id = paras[0].get(f"{W14}paraId")
        result = doc.insert_floating_image(para_id, str(sample_png), 5.0, 3.0)
        # Verify result dict
        assert "rId" in result
        assert result["wrap"] == "square"
        # Verify anchor in XML
        doc2 = doc._tree("word/document.xml")
        anchors = doc2.findall(f".//{WP}anchor")
        assert len(anchors) == 1

    def test_wrap_square_attribute(self, tmp_path, sample_png):
        """wrap='square' produces wp:wrapSquare element."""
        doc = DocxDocument.create(str(tmp_path / "test.docx"))
        tree = doc._tree("word/document.xml")
        paras = tree.findall(f".//{W}p")
        para_id = paras[0].get(f"{W14}paraId")
        doc.insert_floating_image(para_id, str(sample_png), 5.0, 3.0, wrap="square")
        doc2 = doc._tree("word/document.xml")
        wrap_els = doc2.findall(f".//{WP}wrapSquare")
        assert len(wrap_els) == 1

    def test_wrap_topbottom(self, tmp_path, sample_png):
        """wrap='topbottom' produces wp:wrapTopAndBottom element."""
        doc = DocxDocument.create(str(tmp_path / "test.docx"))
        tree = doc._tree("word/document.xml")
        paras = tree.findall(f".//{W}p")
        para_id = paras[0].get(f"{W14}paraId")
        doc.insert_floating_image(para_id, str(sample_png), 4.0, 2.0, wrap="topbottom")
        doc2 = doc._tree("word/document.xml")
        wrap_els = doc2.findall(f".//{WP}wrapTopAndBottom")
        assert len(wrap_els) == 1
        # Also confirm no wrapSquare
        assert len(doc2.findall(f".//{WP}wrapSquare")) == 0

    def test_position_set_correctly(self, tmp_path, sample_png):
        """h_pos and v_pos set correct EMU values in wp:positionH/V."""
        doc = DocxDocument.create(str(tmp_path / "test.docx"))
        tree = doc._tree("word/document.xml")
        paras = tree.findall(f".//{W}p")
        para_id = paras[0].get(f"{W14}paraId")
        h_cm = 2.5
        v_cm = 3.0
        doc.insert_floating_image(para_id, str(sample_png), 5.0, 3.0, h_pos=h_cm, v_pos=v_cm)
        doc2 = doc._tree("word/document.xml")
        expected_h = int(h_cm * 914400 / 2.54)
        expected_v = int(v_cm * 914400 / 2.54)
        pos_h = doc2.find(f".//{WP}positionH")
        pos_v = doc2.find(f".//{WP}positionV")
        assert pos_h is not None
        assert pos_v is not None
        offset_h = pos_h.find(f"{WP}posOffset")
        offset_v = pos_v.find(f"{WP}posOffset")
        assert offset_h is not None
        assert offset_v is not None
        assert int(offset_h.text) == expected_h
        assert int(offset_v.text) == expected_v

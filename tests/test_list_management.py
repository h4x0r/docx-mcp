"""Tests for list management tools: get_lists, promote_list_item, demote_list_item."""
from __future__ import annotations

import uuid

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument, W, W14


def _make_doc(tmp_path):
    doc = DocxDocument.create(str(tmp_path / "test.docx"))
    return doc


def _add_para(doc: DocxDocument, text: str) -> str:
    """Inject a plain paragraph into the document body and return its paraId."""
    tree = doc._tree("word/document.xml")
    body = tree.find(f"{W}body")
    para_id = uuid.uuid4().hex[:8].upper()
    p = etree.Element(f"{W}p")
    p.set(f"{W14}paraId", para_id)
    r = etree.SubElement(p, f"{W}r")
    t = etree.SubElement(r, f"{W}t")
    t.text = text
    # Insert before the last sectPr or at end
    sect = body.find(f"{W}sectPr")
    if sect is not None:
        sect.addprevious(p)
    else:
        body.append(p)
    return para_id


def _add_list_para(doc: DocxDocument, text: str, ilvl: int = 0) -> str:
    """Inject a paragraph that is already a list item at the given ilvl."""
    # First create a numbering entry via add_list so numId=1 exists
    # Then manually set ilvl
    para_id = _add_para(doc, text)
    tree = doc._tree("word/document.xml")
    body = tree.find(f"{W}body")
    # Find the paragraph
    para = None
    for p in body.iter(f"{W}p"):
        if p.get(f"{W14}paraId") == para_id:
            para = p
            break
    assert para is not None

    # Add pPr/numPr/ilvl + numId
    ppr = etree.SubElement(para, f"{W}pPr")
    para.remove(ppr)
    para.insert(0, ppr)
    num_pr = etree.SubElement(ppr, f"{W}numPr")
    ilvl_el = etree.SubElement(num_pr, f"{W}ilvl")
    ilvl_el.set(f"{W}val", str(ilvl))
    num_id_el = etree.SubElement(num_pr, f"{W}numId")
    num_id_el.set(f"{W}val", "1")

    return para_id


class TestGetLists:
    def test_get_lists_returns_abstract_nums(self, tmp_path):
        """get_lists returns list of dicts for each abstractNum in numbering.xml."""
        doc = _make_doc(tmp_path)
        # Bootstrap a numbering.xml with two abstractNums via add_list
        pid1 = _add_para(doc, "item1")
        pid2 = _add_para(doc, "item2")
        doc.add_list([pid1], style="bullet")
        doc.add_list([pid2], style="numbered")

        result = doc.get_lists()

        assert isinstance(result, list)
        assert len(result) == 2
        abstract_ids = [r["abstract_num_id"] for r in result]
        assert 0 in abstract_ids
        assert 1 in abstract_ids
        # Check required keys
        for item in result:
            assert "abstract_num_id" in item
            assert "num_format" in item
            assert "levels" in item
            assert item["levels"] >= 1
        # First is bullet, second is decimal
        by_id = {r["abstract_num_id"]: r for r in result}
        assert by_id[0]["num_format"] == "bullet"
        assert by_id[1]["num_format"] == "decimal"

    def test_get_lists_empty_when_no_numbering(self, tmp_path):
        """get_lists returns [] when numbering.xml is absent."""
        doc = _make_doc(tmp_path)
        # No add_list calls, numbering.xml won't exist
        result = doc.get_lists()
        assert result == []


class TestPromoteListItem:
    def test_promote_list_item_decreases_ilvl(self, tmp_path):
        """promote_list_item decreases ilvl by 1."""
        doc = _make_doc(tmp_path)
        para_id = _add_list_para(doc, "nested item", ilvl=2)

        result = doc.promote_list_item(para_id)

        assert result == {"para_id": para_id, "ilvl": 1}
        # Verify XML was updated
        tree = doc._tree("word/document.xml")
        body = tree.find(f"{W}body")
        para = None
        for p in body.iter(f"{W}p"):
            if p.get(f"{W14}paraId") == para_id:
                para = p
                break
        ilvl_el = para.find(f".//{W}ilvl")
        assert ilvl_el.get(f"{W}val") == "1"

    def test_promote_list_item_clamps_at_zero(self, tmp_path):
        """promote_list_item at ilvl=0 stays at 0."""
        doc = _make_doc(tmp_path)
        para_id = _add_list_para(doc, "top-level item", ilvl=0)

        result = doc.promote_list_item(para_id)

        assert result == {"para_id": para_id, "ilvl": 0}

    def test_promote_list_item_not_list_raises(self, tmp_path):
        """promote_list_item on non-list paragraph raises ValueError."""
        doc = _make_doc(tmp_path)
        para_id = _add_para(doc, "plain paragraph")

        with pytest.raises(ValueError, match="not a list item"):
            doc.promote_list_item(para_id)


class TestDemoteListItem:
    def test_demote_list_item_increases_ilvl(self, tmp_path):
        """demote_list_item increases ilvl by 1."""
        doc = _make_doc(tmp_path)
        para_id = _add_list_para(doc, "item at level 1", ilvl=1)

        result = doc.demote_list_item(para_id)

        assert result == {"para_id": para_id, "ilvl": 2}
        # Verify XML
        tree = doc._tree("word/document.xml")
        body = tree.find(f"{W}body")
        para = None
        for p in body.iter(f"{W}p"):
            if p.get(f"{W14}paraId") == para_id:
                para = p
                break
        ilvl_el = para.find(f".//{W}ilvl")
        assert ilvl_el.get(f"{W}val") == "2"

    def test_demote_list_item_clamps_at_eight(self, tmp_path):
        """demote_list_item at ilvl=8 stays at 8."""
        doc = _make_doc(tmp_path)
        para_id = _add_list_para(doc, "deepest item", ilvl=8)

        result = doc.demote_list_item(para_id)

        assert result == {"para_id": para_id, "ilvl": 8}

    def test_demote_list_item_not_list_raises(self, tmp_path):
        """demote_list_item on non-list paragraph raises ValueError."""
        doc = _make_doc(tmp_path)
        para_id = _add_para(doc, "plain paragraph")

        with pytest.raises(ValueError, match="not a list item"):
            doc.demote_list_item(para_id)

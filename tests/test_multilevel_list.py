"""Tests for multilevel list methods: create_multilevel_list, restart_numbering, suppress_numbering."""  # noqa: E501

from __future__ import annotations

import uuid

from lxml import etree

from docx_mcp.document import W14, DocxDocument, W


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
    # Insert before last child (sectPr)
    body.insert(len(body) - 1, p)
    doc._mark("word/document.xml")
    return para_id


class TestMultilevelList:
    def test_create_multilevel_list_3_levels(self, tmp_path):
        """create_multilevel_list creates 3 w:lvl elements in abstractNum."""
        levels = [
            {"num_fmt": "decimal", "lvl_text": "%1.", "indent": 720, "hanging": 360},
            {"num_fmt": "decimal", "lvl_text": "%1.%2.", "indent": 1440, "hanging": 360},
            {"num_fmt": "decimal", "lvl_text": "%1.%2.%3.", "indent": 2160, "hanging": 360},
        ]
        doc = _make_doc(tmp_path)
        result = doc.create_multilevel_list("MyList", levels)
        assert result["level_count"] == 3
        assert result["name"] == "MyList"
        # Verify in numbering.xml
        num_tree = doc._tree("word/numbering.xml")
        abs_id = str(result["abstract_num_id"])
        abstract = None
        for a in num_tree.findall(f"{W}abstractNum"):
            if a.get(f"{W}abstractNumId") == abs_id:
                abstract = a
                break
        assert abstract is not None
        lvls = abstract.findall(f"{W}lvl")
        assert len(lvls) == 3

    def test_numbering_xml_abstract_num_created(self, tmp_path):
        """abstractNum and num entries are both created in numbering.xml."""
        doc = _make_doc(tmp_path)
        result = doc.create_multilevel_list(
            "TestList",
            [
                {"num_fmt": "bullet", "lvl_text": "•", "indent": 720, "hanging": 360},
            ],
        )
        num_tree = doc._tree("word/numbering.xml")
        # Both abstractNum and num should exist
        abs_nums = num_tree.findall(f"{W}abstractNum")
        nums = num_tree.findall(f"{W}num")
        assert len(abs_nums) >= 1
        assert len(nums) >= 1
        # num references the abstractNum
        found = None
        for n in nums:
            if n.get(f"{W}numId") == str(result["num_id"]):
                found = n
                break
        assert found is not None
        ref = found.find(f"{W}abstractNumId")
        assert ref.get(f"{W}val") == str(result["abstract_num_id"])

    def test_restart_numbering(self, tmp_path):
        """restart_numbering adds lvlOverride with startOverride."""
        doc = _make_doc(tmp_path)
        # Create a list and apply it
        levels = [{"num_fmt": "decimal", "lvl_text": "%1.", "indent": 720, "hanging": 360}]
        doc.create_multilevel_list("RestartTest", levels)
        # Add a paragraph and apply the list to it
        para_id = _add_para(doc, "Item 1")
        doc.add_list([para_id], style="numbered")
        # Now get the actual num_id from the paragraph's numPr
        tree = doc._tree("word/document.xml")
        para = doc._find_para(tree, para_id)
        num_id_el = para.find(f".//{W}numId")
        assert num_id_el is not None
        # restart at 5
        result = doc.restart_numbering(para_id, level=0, start=5)
        assert result["start"] == 5
        # Verify lvlOverride exists on the specific w:num the paragraph references
        num_tree = doc._tree("word/numbering.xml")
        num_id_val = num_id_el.get(f"{W}val")
        target_num = None
        for n in num_tree.findall(f"{W}num"):
            if n.get(f"{W}numId") == num_id_val:
                target_num = n
                break
        assert target_num is not None, f"w:num with numId={num_id_val} not found"
        override = target_num.find(f"{W}lvlOverride")
        assert override is not None
        start_override = override.find(f"{W}startOverride")
        assert start_override is not None
        assert start_override.get(f"{W}val") == "5"

    def test_suppress_numbering(self, tmp_path):
        """suppress_numbering sets numId to 0."""
        doc = _make_doc(tmp_path)
        para_id = _add_para(doc, "Item to suppress")
        doc.add_list([para_id], style="bullet")
        result = doc.suppress_numbering(para_id)
        assert result["suppressed"] is True
        tree = doc._tree("word/document.xml")
        para = doc._find_para(tree, para_id)
        num_id_el = para.find(f".//{W}numId")
        assert num_id_el is not None
        assert num_id_el.get(f"{W}val") == "0"

    def test_heading_numbering_binding(self, tmp_path):
        """create_multilevel_list with style binds abstractNum to heading style."""
        doc = _make_doc(tmp_path)
        levels = [
            {
                "num_fmt": "decimal",
                "lvl_text": "%1.",
                "indent": 0,
                "hanging": 360,
                "style": "Heading 1",
            },  # noqa: E501
        ]
        result = doc.create_multilevel_list("HeadingList", levels)
        num_tree = doc._tree("word/numbering.xml")
        # Find abstractNum with our abstract_num_id
        abs_id = str(result["abstract_num_id"])
        abstract = None
        for a in num_tree.findall(f"{W}abstractNum"):
            if a.get(f"{W}abstractNumId") == abs_id:
                abstract = a
                break
        assert abstract is not None
        lvl = abstract.find(f"{W}lvl")
        pPr = lvl.find(f"{W}pPr")
        assert pPr is not None
        pStyle = pPr.find(f"{W}pStyle")  # inside pPr, not lvl
        assert pStyle is not None
        assert pStyle.get(f"{W}val") == "Heading 1"

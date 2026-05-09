"""Tests for ContentControlsMixin — SDT content controls."""
from __future__ import annotations

from pathlib import Path

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument, W, W14
from docx_mcp.document.errors import DocxMcpError, ErrCode

_W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"


def _make_doc(tmp_path: Path) -> tuple[DocxDocument, str]:
    """Create a fresh document and return (doc, first_para_id)."""
    out = str(tmp_path / "test.docx")
    doc = DocxDocument.create(out)
    tree = doc._tree("word/document.xml")
    para_id = None
    for p in tree.iter(f"{W}p"):
        pid = p.get(f"{W14}paraId")
        if pid is not None:
            para_id = pid
            break
    assert para_id is not None, "No paragraph with w14:paraId found"
    return doc, para_id


class TestContentControls:
    def test_add_checkbox_control(self, tmp_path: Path):
        """Wraps paragraph in checkbox SDT; returns correct dict."""
        doc, para_id = _make_doc(tmp_path)
        result = doc.add_content_control(para_id, "chk1", "checkbox", label="Accept Terms")
        assert result["tag"] == "chk1"
        assert result["type"] == "checkbox"
        assert result["label"] == "Accept Terms"

        tree = doc._tree("word/document.xml")
        sdt = tree.find(f".//{W}sdt")
        assert sdt is not None, "w:sdt element not found"
        sdtPr = sdt.find(f"{W}sdtPr")
        assert sdtPr is not None

        # tag
        tag_el = sdtPr.find(f"{W}tag")
        assert tag_el is not None
        assert tag_el.get(f"{W}val") == "chk1"

        # alias / label
        alias_el = sdtPr.find(f"{W}alias")
        assert alias_el is not None
        assert alias_el.get(f"{W}val") == "Accept Terms"

        # w14:checkbox
        checkbox_el = sdtPr.find(f"{{{_W14_NS}}}checkbox")
        assert checkbox_el is not None

        # sdtContent contains the paragraph with default unchecked char
        sdtContent = sdt.find(f"{W}sdtContent")
        assert sdtContent is not None
        text = "".join(t.text for t in sdtContent.iter(f"{W}t") if t.text)
        assert text == "☐"

    def test_add_dropdown_control(self, tmp_path: Path):
        """Dropdown SDT has listItem entries for each option."""
        doc, para_id = _make_doc(tmp_path)
        result = doc.add_content_control(
            para_id, "dd1", "dropdown",
            label="Pick one",
            options=["Alpha", "Beta", "Gamma"],
            default="Alpha",
        )
        assert result["tag"] == "dd1"
        assert result["type"] == "dropdown"

        tree = doc._tree("word/document.xml")
        sdt = tree.find(f".//{W}sdt")
        sdtPr = sdt.find(f"{W}sdtPr")

        ddl = sdtPr.find(f"{W}dropDownList")
        assert ddl is not None

        items = ddl.findall(f"{W}listItem")
        assert len(items) == 3
        display_texts = [i.get(f"{W}displayText") for i in items]
        values = [i.get(f"{W}value") for i in items]
        assert display_texts == ["Alpha", "Beta", "Gamma"]
        assert values == ["Alpha", "Beta", "Gamma"]

        # sdtContent default text
        sdtContent = sdt.find(f"{W}sdtContent")
        text = "".join(t.text for t in sdtContent.iter(f"{W}t") if t.text)
        assert text == "Alpha"

    def test_add_date_picker_control(self, tmp_path: Path):
        """Date SDT has w:date element in sdtPr."""
        doc, para_id = _make_doc(tmp_path)
        result = doc.add_content_control(
            para_id, "dt1", "date",
            label="Pick Date",
            default="January 1, 2024",
        )
        assert result["tag"] == "dt1"
        assert result["type"] == "date"

        tree = doc._tree("word/document.xml")
        sdt = tree.find(f".//{W}sdt")
        sdtPr = sdt.find(f"{W}sdtPr")

        date_el = sdtPr.find(f"{W}date")
        assert date_el is not None

        date_fmt = date_el.find(f"{W}dateFormat")
        assert date_fmt is not None
        assert date_fmt.get(f"{W}val") == "MMMM d, yyyy"

        sdtContent = sdt.find(f"{W}sdtContent")
        text = "".join(t.text for t in sdtContent.iter(f"{W}t") if t.text)
        assert text == "January 1, 2024"

    def test_add_text_control(self, tmp_path: Path):
        """Text SDT has w:text element in sdtPr."""
        doc, para_id = _make_doc(tmp_path)
        result = doc.add_content_control(
            para_id, "txt1", "text",
            label="Enter text",
            default="placeholder",
        )
        assert result["tag"] == "txt1"
        assert result["type"] == "text"

        tree = doc._tree("word/document.xml")
        sdt = tree.find(f".//{W}sdt")
        sdtPr = sdt.find(f"{W}sdtPr")

        text_el = sdtPr.find(f"{W}text")
        assert text_el is not None

        sdtContent = sdt.find(f"{W}sdtContent")
        content_text = "".join(t.text for t in sdtContent.iter(f"{W}t") if t.text)
        assert content_text == "placeholder"

    def test_set_control_value(self, tmp_path: Path):
        """set_content_control_value updates w:t text (and checkbox state)."""
        doc, para_id = _make_doc(tmp_path)
        doc.add_content_control(para_id, "txt2", "text", default="old text")
        result = doc.set_content_control_value("txt2", "new text")
        assert result["tag"] == "txt2"
        assert result["value"] == "new text"

        tree = doc._tree("word/document.xml")
        sdt = tree.find(f".//{W}sdt")
        sdtContent = sdt.find(f"{W}sdtContent")
        text = "".join(t.text for t in sdtContent.iter(f"{W}t") if t.text)
        assert text == "new text"

        # checkbox checked/unchecked
        doc2, para_id2 = _make_doc(tmp_path / "cb")
        (tmp_path / "cb").mkdir(exist_ok=True)
        doc2.add_content_control(para_id2, "chk2", "checkbox")
        doc2.set_content_control_value("chk2", "true")
        tree2 = doc2._tree("word/document.xml")
        sdt2 = tree2.find(f".//{W}sdt")
        sdtPr2 = sdt2.find(f"{W}sdtPr")
        checked_el = sdtPr2.find(f".//{{{_W14_NS}}}checked")
        assert checked_el is not None
        assert checked_el.get(f"{{{_W14_NS}}}val") == "1"
        sdtContent2 = sdt2.find(f"{W}sdtContent")
        text2 = "".join(t.text for t in sdtContent2.iter(f"{W}t") if t.text)
        assert text2 == "☑"

    def test_list_content_controls(self, tmp_path: Path):
        """get_content_controls returns all controls with tag/type/label/value."""
        out = str(tmp_path / "multi.docx")
        doc = DocxDocument.create(out)
        tree = doc._tree("word/document.xml")

        # Collect two distinct para_ids
        para_ids = []
        for p in tree.iter(f"{W}p"):
            pid = p.get(f"{W14}paraId")
            if pid:
                para_ids.append(pid)
            if len(para_ids) >= 2:
                break

        # If only one paragraph, add more text to get a second
        # Use the same para multiple times is allowed since after wrapping
        # the first control, the original para is inside sdtContent — grab a fresh one.
        doc.add_content_control(para_ids[0], "tagA", "text", label="LabelA", default="valA")

        # Get a fresh para_id after first wrap (first is now inside sdt)
        tree = doc._tree("word/document.xml")
        second_pid = None
        for p in tree.iter(f"{W}p"):
            pid = p.get(f"{W14}paraId")
            if pid and pid not in (para_ids[0],):
                second_pid = pid
                break

        if second_pid is None:
            pytest.skip("Document has only one paragraph")

        doc.add_content_control(second_pid, "tagB", "checkbox", label="LabelB")

        controls = doc.get_content_controls()
        assert len(controls) >= 2

        tags = {c["tag"] for c in controls}
        assert "tagA" in tags
        assert "tagB" in tags

        by_tag = {c["tag"]: c for c in controls}
        assert by_tag["tagA"]["type"] == "text"
        assert by_tag["tagA"]["label"] == "LabelA"
        assert by_tag["tagA"]["value"] == "valA"
        assert by_tag["tagB"]["type"] == "checkbox"
        assert by_tag["tagB"]["label"] == "LabelB"

    def test_lock_content_control(self, tmp_path: Path):
        """lock_content_control adds w:lock to sdtPr."""
        doc, para_id = _make_doc(tmp_path)
        doc.add_content_control(para_id, "lockme", "text")
        result = doc.lock_content_control("lockme", "sdtLocked")
        assert result["tag"] == "lockme"
        assert result["lock"] == "sdtLocked"

        tree = doc._tree("word/document.xml")
        sdt = tree.find(f".//{W}sdt")
        sdtPr = sdt.find(f"{W}sdtPr")
        lock_el = sdtPr.find(f"{W}lock")
        assert lock_el is not None
        assert lock_el.get(f"{W}val") == "sdtLocked"

        # contentLocked variant
        doc2, para_id2 = _make_doc(tmp_path / "lock2")
        (tmp_path / "lock2").mkdir(exist_ok=True)
        doc2.add_content_control(para_id2, "lockme2", "text")
        doc2.lock_content_control("lockme2", "contentLocked")
        tree2 = doc2._tree("word/document.xml")
        sdt2 = tree2.find(f".//{W}sdt")
        sdtPr2 = sdt2.find(f"{W}sdtPr")
        lock2 = sdtPr2.find(f"{W}lock")
        assert lock2 is not None
        assert lock2.get(f"{W}val") == "contentLocked"

    def test_duplicate_tag_raises(self, tmp_path: Path):
        """Adding a control with an existing tag raises OOXML_INVALID."""
        doc, para_id = _make_doc(tmp_path)
        doc.add_content_control(para_id, "dup", "text")

        # Get another paragraph (the original is now inside sdt)
        tree = doc._tree("word/document.xml")
        second_pid = None
        for p in tree.iter(f"{W}p"):
            pid = p.get(f"{W14}paraId")
            if pid and pid != para_id:
                second_pid = pid
                break

        if second_pid is None:
            # No second para available — use same para_id with any para
            # The duplicate tag check happens before para lookup in implementation;
            # just use a bogus para_id if needed:
            second_pid = "DEADBEEF"

        with pytest.raises(DocxMcpError) as exc_info:
            doc.add_content_control(second_pid, "dup", "text")
        assert exc_info.value.code == ErrCode.OOXML_INVALID

    def test_para_not_found_raises(self, tmp_path: Path):
        """add_content_control raises PARA_NOT_FOUND for unknown para_id."""
        doc, _ = _make_doc(tmp_path)
        with pytest.raises(DocxMcpError) as exc_info:
            doc.add_content_control("DEADBEEF", "t1", "text")
        assert exc_info.value.code == ErrCode.PARA_NOT_FOUND

    def test_set_value_tag_not_found_raises(self, tmp_path: Path):
        """set_content_control_value raises BOOKMARK_NOT_FOUND for missing tag."""
        doc, _ = _make_doc(tmp_path)
        with pytest.raises(DocxMcpError) as exc_info:
            doc.set_content_control_value("nonexistent", "value")
        assert exc_info.value.code == ErrCode.BOOKMARK_NOT_FOUND

    def test_lock_tag_not_found_raises(self, tmp_path: Path):
        """lock_content_control raises BOOKMARK_NOT_FOUND for missing tag."""
        doc, _ = _make_doc(tmp_path)
        with pytest.raises(DocxMcpError) as exc_info:
            doc.lock_content_control("nonexistent")
        assert exc_info.value.code == ErrCode.BOOKMARK_NOT_FOUND

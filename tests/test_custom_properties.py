"""Tests for custom document properties (docProps/custom.xml)."""

from __future__ import annotations

import pytest
from lxml import etree

from docx_mcp import server

CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
CUSTOM = f"{{{CUSTOM_NS}}}"
VT = f"{{{VT_NS}}}"


def _inject_custom(doc, props: dict) -> None:
    root = etree.Element(f"{CUSTOM}Properties", nsmap={
        None: CUSTOM_NS,
        "vt": VT_NS,
    })
    for pid_offset, (name, value) in enumerate(props.items()):
        prop = etree.SubElement(root, f"{CUSTOM}property")
        prop.set("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        prop.set("pid", str(2 + pid_offset))
        prop.set("name", name)
        vt_el = etree.SubElement(prop, f"{VT}lpwstr")
        vt_el.text = value
    doc._trees["docProps/custom.xml"] = root


@pytest.fixture()
def doc(test_docx):
    server.open_document(str(test_docx))
    return server._doc


def test_get_custom_properties_missing_file(doc):
    doc._trees.pop("docProps/custom.xml", None)
    result = doc.get_custom_properties()
    assert result == {}


def test_get_custom_properties_returns_name_value_dict(doc):
    _inject_custom(doc, {"Author": "Alice", "Department": "Legal"})
    result = doc.get_custom_properties()
    assert result == {"Author": "Alice", "Department": "Legal"}


def test_set_custom_property_adds_new(doc):
    doc._trees.pop("docProps/custom.xml", None)
    result = doc.set_custom_property("Status", "Draft")
    assert result["name"] == "Status"
    assert result["value"] == "Draft"
    assert result["vt_type"] == "lpwstr"
    props = doc.get_custom_properties()
    assert props["Status"] == "Draft"


def test_set_custom_property_updates_existing(doc):
    _inject_custom(doc, {"Status": "Draft"})
    result = doc.set_custom_property("Status", "Final")
    assert result["value"] == "Final"
    props = doc.get_custom_properties()
    assert props["Status"] == "Final"
    assert len(props) == 1


def test_set_custom_property_vt_type_respected(doc):
    doc._trees.pop("docProps/custom.xml", None)
    result = doc.set_custom_property("Count", "42", vt_type="i4")
    assert result["vt_type"] == "i4"
    root = doc._trees["docProps/custom.xml"]
    prop = root.find(f"{CUSTOM}property[@name='Count']")
    assert prop is not None
    child = prop[0]
    assert child.tag == f"{VT}i4"
    assert child.text == "42"


def test_set_custom_property_bootstraps_missing_file(doc):
    doc._trees.pop("docProps/custom.xml", None)
    doc.set_custom_property("Project", "Alpha")
    assert "docProps/custom.xml" in doc._trees
    root = doc._trees["docProps/custom.xml"]
    assert root.tag == f"{CUSTOM}Properties"


def test_delete_custom_property_removes(doc):
    _inject_custom(doc, {"Foo": "bar", "Baz": "qux"})
    result = doc.delete_custom_property("Foo")
    assert result == {"deleted": "Foo"}
    props = doc.get_custom_properties()
    assert "Foo" not in props
    assert "Baz" in props


def test_delete_custom_property_raises_on_missing(doc):
    _inject_custom(doc, {"Foo": "bar"})
    with pytest.raises(ValueError):
        doc.delete_custom_property("NonExistent")

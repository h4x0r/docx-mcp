"""RED tests for split_document — split DOCX at heading boundaries."""

from __future__ import annotations

import copy
import json
import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"


def _para(text: str, style: str | None = None) -> etree._Element:
    p = etree.Element(f"{W}p")
    if style:
        ppr = etree.SubElement(p, f"{W}pPr")
        ps = etree.SubElement(ppr, f"{W}pStyle")
        ps.set(f"{W}val", style)
    r = etree.SubElement(p, f"{W}r")
    t = etree.SubElement(r, f"{W}t")
    t.text = text
    return p


def _make_doc_with_sections(tmp_path: Path):
    """Create a DocxDocument with 3 heading sections."""
    from docx_mcp.document import DocxDocument

    out = str(tmp_path / "source.docx")
    doc = DocxDocument.create(out)

    body = doc._require("word/document.xml").find(f"{W}body")
    for el in list(body):
        body.remove(el)

    body.append(_para("Introduction", "Heading1"))
    body.append(_para("Intro content"))
    body.append(_para("Section Two", "Heading1"))
    body.append(_para("Section two content"))
    body.append(_para("Section Three", "Heading1"))
    body.append(_para("Final content"))
    etree.SubElement(body, f"{W}sectPr")

    doc._mark("word/document.xml")
    doc.save()
    return doc


def _make_doc_with_preamble(tmp_path: Path):
    """Create a DocxDocument with content before first heading."""
    from docx_mcp.document import DocxDocument

    out = str(tmp_path / "preamble.docx")
    doc = DocxDocument.create(out)

    body = doc._require("word/document.xml").find(f"{W}body")
    for el in list(body):
        body.remove(el)

    body.append(_para("Preamble content"))
    body.append(_para("More preamble"))
    body.append(_para("Chapter One", "Heading1"))
    body.append(_para("Chapter content"))
    body.append(_para("Chapter Two", "Heading1"))
    body.append(_para("Chapter two content"))
    etree.SubElement(body, f"{W}sectPr")

    doc._mark("word/document.xml")
    doc.save()
    return doc


def _make_doc_with_heading2(tmp_path: Path):
    """Create a DocxDocument with Heading2 sections."""
    from docx_mcp.document import DocxDocument

    out = str(tmp_path / "h2source.docx")
    doc = DocxDocument.create(out)

    body = doc._require("word/document.xml").find(f"{W}body")
    for el in list(body):
        body.remove(el)

    body.append(_para("Top Level", "Heading1"))
    body.append(_para("Sub A", "Heading2"))
    body.append(_para("Content A"))
    body.append(_para("Sub B", "Heading2"))
    body.append(_para("Content B"))
    etree.SubElement(body, f"{W}sectPr")

    doc._mark("word/document.xml")
    doc.save()
    return doc


# ── Tests ─────────────────────────────────────────────────────────────────


def test_split_produces_correct_file_count(tmp_path):
    doc = _make_doc_with_sections(tmp_path)
    out_dir = str(tmp_path / "out")
    result = doc.split_document(output_dir=out_dir)
    assert result["parts"] == 3


def test_split_output_files_exist_on_disk(tmp_path):
    doc = _make_doc_with_sections(tmp_path)
    out_dir = str(tmp_path / "out")
    result = doc.split_document(output_dir=out_dir)
    for fpath in result["files"]:
        assert Path(fpath).exists(), f"Missing: {fpath}"


def test_split_return_dict_keys(tmp_path):
    doc = _make_doc_with_sections(tmp_path)
    out_dir = str(tmp_path / "out")
    result = doc.split_document(output_dir=out_dir)
    assert "files" in result
    assert "parts" in result
    assert "output_dir" in result


def test_split_parts_equals_heading_count(tmp_path):
    doc = _make_doc_with_sections(tmp_path)
    out_dir = str(tmp_path / "out")
    result = doc.split_document(output_dir=out_dir)
    assert result["parts"] == len(result["files"])


def test_split_output_files_are_valid_zip(tmp_path):
    doc = _make_doc_with_sections(tmp_path)
    out_dir = str(tmp_path / "out")
    result = doc.split_document(output_dir=out_dir)
    for fpath in result["files"]:
        assert zipfile.is_zipfile(fpath), f"Not a valid ZIP: {fpath}"


def test_split_at_heading2(tmp_path):
    doc = _make_doc_with_heading2(tmp_path)
    out_dir = str(tmp_path / "out2")
    result = doc.split_document(output_dir=out_dir, at_heading_level=2)
    # H1 paragraph before first H2 becomes non-empty preamble: 1 preamble + 2 H2 sections
    assert result["parts"] == 3


def test_split_default_output_dir(tmp_path):
    doc = _make_doc_with_sections(tmp_path)
    result = doc.split_document()
    assert result["output_dir"]
    assert Path(result["output_dir"]).exists()


def test_split_preamble_included_when_nonempty(tmp_path):
    doc = _make_doc_with_preamble(tmp_path)
    out_dir = str(tmp_path / "outp")
    result = doc.split_document(output_dir=out_dir)
    # preamble + 2 headings = 3 parts
    assert result["parts"] == 3


def test_split_file_names_are_docx(tmp_path):
    doc = _make_doc_with_sections(tmp_path)
    out_dir = str(tmp_path / "outn")
    result = doc.split_document(output_dir=out_dir)
    for fpath in result["files"]:
        assert fpath.endswith(".docx"), f"Not a .docx: {fpath}"

"""Splitting mixin: split a DOCX into multiple files at heading boundaries."""

from __future__ import annotations

import copy
import os
import re
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from .base import W


def _slugify(text: str, max_len: int = 40) -> str:
    text = re.sub(r"[^\w\s-]", "", text).strip()
    text = re.sub(r"\s+", "_", text)
    return text[:max_len] or "part"


def _is_heading(paragraph: etree._Element, level: int) -> bool:
    ppr = paragraph.find(f"{W}pPr")
    if ppr is None:
        return False
    ps = ppr.find(f"{W}pStyle")
    if ps is None:
        return False
    val = (ps.get(f"{W}val") or "").lower()
    return val in (f"heading{level}", f"heading {level}")


def _para_text(paragraph: etree._Element) -> str:
    return "".join(t.text for t in paragraph.iter(f"{W}t") if t.text)


class SplittingMixin:
    """Split document at heading boundaries."""

    def split_document(self, output_dir: str = "", at_heading_level: int = 1) -> dict:
        """Split the open document into multiple DOCX files, one per heading section.

        Args:
            output_dir: Directory for output files. Defaults to <workdir>/split_output.
            at_heading_level: Heading level to split on (default 1).

        Returns:
            {"output_dir": str, "files": list[str], "parts": int}
        """
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")

        out_path = Path(output_dir) if output_dir else Path(self.workdir) / "split_output"
        out_path.mkdir(parents=True, exist_ok=True)

        # Collect sectPr (last child if it's sectPr)
        sect_pr = None
        children = list(body)
        if children and children[-1].tag == f"{W}sectPr":
            sect_pr = children[-1]
            children = children[:-1]

        # Walk children and collect sections
        sections: list[dict] = []
        current: list[etree._Element] = []
        current_title: str | None = None

        for child in children:
            if child.tag == f"{W}p" and _is_heading(child, at_heading_level):
                if current:
                    sections.append({"title": current_title, "elements": current})
                current = [child]
                current_title = _para_text(child)
            else:
                current.append(child)

        if current:
            sections.append({"title": current_title, "elements": current})

        # Filter empty preamble (no title = preamble section)
        final_sections = []
        for sec in sections:
            if sec["title"] is None:
                has_content = any(
                    _para_text(el).strip()
                    for el in sec["elements"]
                    if el.tag == f"{W}p"
                )
                if not has_content:
                    continue
            final_sections.append(sec)

        # Save the parent to a temp zip so we can copy it
        tmp_docx = Path(tempfile.mktemp(suffix=".docx"))
        self.save(str(tmp_docx), backup=False)

        files: list[str] = []
        used_names: dict[str, int] = {}

        for idx, sec in enumerate(final_sections):
            if sec["title"] is not None:
                base = _slugify(sec["title"])
            else:
                base = "preamble"

            # Deduplicate names
            if base in used_names:
                used_names[base] += 1
                name = f"{base}_{used_names[base]}"
            else:
                used_names[base] = 0
                name = base

            out_file = str(out_path / f"{name}.docx")

            # Build new document.xml body content
            body_children_xml = b""
            for el in sec["elements"]:
                body_children_xml += etree.tostring(copy.deepcopy(el))
            if sect_pr is not None:
                body_children_xml += etree.tostring(copy.deepcopy(sect_pr))

            # Read the parent zip, replace word/document.xml body
            with zipfile.ZipFile(str(tmp_docx), "r") as zin:
                with zipfile.ZipFile(out_file, "w", zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        if item.filename == "word/document.xml":
                            orig_xml = zin.read(item.filename)
                            new_xml = _replace_body(orig_xml, body_children_xml)
                            zout.writestr(item, new_xml)
                        else:
                            zout.writestr(item, zin.read(item.filename))

            files.append(out_file)

        # Clean up temp file
        tmp_docx.unlink(missing_ok=True)

        return {
            "output_dir": str(out_path),
            "files": files,
            "parts": len(files),
        }


def _replace_body(doc_xml_bytes: bytes, body_children_xml: bytes) -> bytes:
    """Parse document.xml, replace body children, return serialized bytes."""
    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(doc_xml_bytes, parser)
    body = root.find(f"{W}body")
    for child in list(body):
        body.remove(child)
    # Parse and append new children
    wrapper = etree.fromstring(b"<root>" + body_children_xml + b"</root>")
    for child in wrapper:
        body.append(child)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

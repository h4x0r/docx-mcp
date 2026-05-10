"""Watermark mixin: insert and remove VML watermarks in the default header."""

from __future__ import annotations

import contextlib

from lxml import etree

from .base import V, W, RELS

_V_NS = "urn:schemas-microsoft-com:vml"
_HEADER_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
)
_HEADER_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
)

_STYLE_DIAGONAL = (
    "position:absolute;"
    "margin-left:0;"
    "margin-top:0;"
    "width:527.85pt;"
    "height:131.95pt;"
    "z-index:-251654144;"
    "mso-position-horizontal:center;"
    "mso-position-horizontal-relative:margin;"
    "mso-position-vertical:center;"
    "mso-position-vertical-relative:margin"
)

_STYLE_HORIZONTAL = (
    "position:absolute;"
    "margin-left:0;"
    "margin-top:0;"
    "width:527.85pt;"
    "height:131.95pt;"
    "z-index:-251654144;"
    "mso-position-horizontal:center;"
    "mso-position-horizontal-relative:margin;"
    "mso-position-vertical:center;"
    "mso-position-vertical-relative:margin;"
    "mso-rotation:0"
)


class WatermarkMixin:
    """Insert and remove VML watermarks in the document's default header."""

    def _default_header_path(self) -> str | None:
        rels = self._tree("word/_rels/document.xml.rels")
        if rels is None:
            return None
        for rel in rels.findall(f"{RELS}Relationship"):
            if rel.get("Type") == _HEADER_REL_TYPE:
                target = rel.get("Target", "")
                return f"word/{target}"
        return None

    def _ensure_default_header(self) -> str:
        path = self._default_header_path()
        if path and path in self._trees:
            return path

        if path is None:
            path = "word/header1.xml"

        hdr_ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = etree.Element(
            f"{W}hdr",
            nsmap={
                "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "v": _V_NS,
            },
        )
        etree.SubElement(root, f"{W}p")
        self._trees[path] = root

        import os
        if self.workdir:
            fp = self.workdir / path
            fp.parent.mkdir(parents=True, exist_ok=True)
            etree.ElementTree(root).write(
                str(fp), xml_declaration=True, encoding="UTF-8"
            )

        ct = self._tree("[Content_Types].xml")
        CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
        if ct is not None:
            existing = {e.get("PartName") for e in ct.findall(f"{CT}Override")}
            part_name = f"/{path}"
            if part_name not in existing:
                ov = etree.SubElement(ct, f"{CT}Override")
                ov.set("PartName", part_name)
                ov.set("ContentType", _HEADER_CT)
                self._mark("[Content_Types].xml")

        rels_tree = self._tree("word/_rels/document.xml.rels")
        if rels_tree is not None:
            existing_targets = {
                r.get("Target") for r in rels_tree.findall(f"{RELS}Relationship")
            }
            fname = os.path.basename(path)
            if fname not in existing_targets:
                max_rid = 0
                for r in rels_tree.findall(f"{RELS}Relationship"):
                    rid = r.get("Id", "")
                    if rid.startswith("rId"):
                        with contextlib.suppress(ValueError):
                            max_rid = max(max_rid, int(rid[3:]))
                rel = etree.SubElement(rels_tree, f"{RELS}Relationship")
                rel.set("Id", f"rId{max_rid + 1}")
                rel.set("Type", _HEADER_REL_TYPE)
                rel.set("Target", fname)
                self._mark("word/_rels/document.xml.rels")

        self._mark(path)
        return path

    def insert_watermark(self, text: str, diagonal: bool = True) -> dict:
        """Insert a VML watermark into the default header.

        Args:
            text: Watermark text (e.g. "DRAFT").
            diagonal: If True, diagonal orientation; if False, horizontal.

        Returns:
            {"header": "default", "text": text, "diagonal": diagonal}
        """
        path = self._ensure_default_header()
        hdr = self._trees[path]

        para = hdr.find(f"{W}p")
        if para is None:
            para = etree.SubElement(hdr, f"{W}p")

        style_str = _STYLE_DIAGONAL if diagonal else _STYLE_HORIZONTAL

        run = etree.SubElement(para, f"{W}r")

        rpr = etree.SubElement(run, f"{W}rPr")
        rpr_style = etree.SubElement(rpr, f"{W}rStyle")
        rpr_style.set(f"{W}val", "WatermarkText")

        pict = etree.SubElement(run, f"{W}pict")

        shape = etree.SubElement(
            pict,
            f"{{{_V_NS}}}shape",
            nsmap={"v": _V_NS},
        )
        shape.set("id", "watermark1")
        shape.set("type", "#_x0000_t136")
        shape.set("style", style_str)
        shape.set("fillcolor", "#d8d8d8")
        shape.set("stroked", "f")

        tp = etree.SubElement(shape, f"{{{_V_NS}}}textpath")
        tp.set("style", "font-family:'Calibri';font-size:1pt")
        tp.set("string", text)
        tp.set("trim", "t")
        tp.set("fitshape", "t")

        etree.SubElement(shape, f"{{{_V_NS}}}imagedata")

        self._mark(path)
        return {"header": "default", "text": text, "diagonal": diagonal}

    def remove_watermark(self) -> dict:
        """Remove VML watermark runs from all document headers.

        Returns:
            {"removed": int} count of watermark <w:r> elements removed.
        """
        removed = 0
        for rel_path, tree in self._trees.items():
            if not rel_path.startswith("word/header"):
                continue
            for run in list(tree.iter(f"{W}r")):
                pict = run.find(f"{W}pict")
                if pict is None:
                    continue
                has_textpath = any(
                    True for _ in pict.iter(f"{{{_V_NS}}}textpath")
                )
                if not has_textpath:
                    continue
                parent = run.getparent()
                if parent is not None:
                    parent.remove(run)
                    removed += 1
                    self._mark(rel_path)
        return {"removed": removed}

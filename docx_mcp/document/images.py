"""Images mixin: list and insert embedded images."""

from __future__ import annotations

import shutil
from pathlib import Path

from lxml import etree

from .base import CT, RELS, W14, WP, A, R, W

# Additional namespace constants for floating images
_PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"  # bare URI for nsmap


class ImagesMixin:
    """Image operations."""

    def get_images(self) -> list[dict]:
        """Get all embedded images with metadata."""
        doc = self._tree("word/document.xml")
        rels = self._tree("word/_rels/document.xml.rels")
        if doc is None:
            return []
        images = []
        for blip in doc.iter(f"{A}blip"):
            embed = blip.get(f"{R}embed")
            if not embed:
                continue
            info: dict = {"rId": embed, "filename": "", "content_type": ""}
            if rels is not None:
                rel = rels.find(f'{RELS}Relationship[@Id="{embed}"]')
                if rel is not None:
                    info["filename"] = rel.get("Target", "").split("/")[-1]
            # Get dimensions from wp:extent
            drawing = blip.getparent()
            while drawing is not None and drawing.tag != f"{W}drawing":
                drawing = drawing.getparent()
            if drawing is not None:
                extent = drawing.find(f".//{WP}extent")
                if extent is not None:
                    info["width_emu"] = int(extent.get("cx", "0"))
                    info["height_emu"] = int(extent.get("cy", "0"))
            # Content type from [Content_Types].xml
            ct_tree = self._tree("[Content_Types].xml")
            if ct_tree is not None:
                ext = info["filename"].rsplit(".", 1)[-1] if "." in info["filename"] else ""
                for default in ct_tree.findall(f"{CT}Default"):
                    if default.get("Extension") == ext:
                        info["content_type"] = default.get("ContentType", "")
            images.append(info)
        return images

    def insert_image(
        self,
        para_id: str,
        image_path: str,
        *,
        width_emu: int = 2000000,
        height_emu: int = 2000000,
    ) -> dict:
        """Insert an image into the document after a paragraph.

        Args:
            para_id: paraId of the paragraph to insert after.
            image_path: Path to the image file on disk.
            width_emu: Image width in EMUs (914400 = 1 inch).
            height_emu: Image height in EMUs.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        src = Path(image_path)
        ext = src.suffix.lstrip(".")

        # Copy image to word/media/
        media_dir = self.workdir / "word" / "media"
        media_dir.mkdir(parents=True, exist_ok=True)
        existing = list(media_dir.glob("image*.*"))
        idx = len(existing) + 1
        filename = f"image{idx}.{ext}"
        shutil.copy2(str(src), str(media_dir / filename))

        # Add relationship
        rels = self._tree("word/_rels/document.xml.rels")
        existing_rids = [r.get("Id") for r in rels.findall(f"{RELS}Relationship")]
        rid_num = (
            max(
                (int(r.replace("rId", "")) for r in existing_rids if r.startswith("rId")),
                default=0,
            )
            + 1
        )
        rid = f"rId{rid_num}"
        rel = etree.SubElement(rels, f"{RELS}Relationship")
        rel.set("Id", rid)
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        rel.set("Target", f"media/{filename}")
        self._mark("word/_rels/document.xml.rels")

        # Ensure content type exists
        ct_tree = self._tree("[Content_Types].xml")
        ct_map = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif"}
        content_type = ct_map.get(ext, f"image/{ext}")
        has_ext = any(d.get("Extension") == ext for d in ct_tree.findall(f"{CT}Default"))
        if not has_ext:
            default = etree.SubElement(ct_tree, f"{CT}Default")
            default.set("Extension", ext)
            default.set("ContentType", content_type)
            self._mark("[Content_Types].xml")

        # Build drawing XML
        ns_wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
        ns_pic = "http://schemas.openxmlformats.org/drawingml/2006/picture"
        ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        new_para = etree.Element(f"{W}p")
        new_para.set(f"{W14}paraId", self._new_para_id())
        new_para.set(f"{W14}textId", "77777777")
        run = etree.SubElement(new_para, f"{W}r")
        drawing = etree.SubElement(run, f"{W}drawing")
        inline = etree.SubElement(drawing, f"{{{ns_wp}}}inline")
        extent = etree.SubElement(inline, f"{{{ns_wp}}}extent")
        extent.set("cx", str(width_emu))
        extent.set("cy", str(height_emu))
        graphic = etree.SubElement(inline, f"{{{ns_a}}}graphic")
        gdata = etree.SubElement(
            graphic,
            f"{{{ns_a}}}graphicData",
            uri=ns_pic,
        )
        pic = etree.SubElement(gdata, f"{{{ns_pic}}}pic")
        blip_fill = etree.SubElement(pic, f"{{{ns_pic}}}blipFill")
        blip = etree.SubElement(blip_fill, f"{{{ns_a}}}blip")
        blip.set(f"{{{ns_r}}}embed", rid)

        para.addnext(new_para)
        self._mark("word/document.xml")

        return {
            "filename": filename,
            "rId": rid,
            "width_emu": width_emu,
            "height_emu": height_emu,
        }

    # ── Floating image helpers ───────────────────────────────────────────────

    def _next_drawing_id(self, doc_tree: etree._Element) -> int:
        """Return max existing wp:docPr id + 1 for unique id allocation."""
        max_id = 0
        for el in doc_tree.iter(f"{WP}docPr"):
            try:
                max_id = max(max_id, int(el.get("id", "0")))
            except ValueError:
                pass
        return max_id + 1

    def insert_floating_image(
        self,
        para_id: str,
        image_path: str,
        width_cm: float,
        height_cm: float,
        h_pos: float = 0.0,
        v_pos: float = 0.0,
        wrap: str = "square",
    ) -> dict:
        """Insert a floating (anchored) image positioned at (h_pos, v_pos) cm from page origin.

        Returns {"filename": str, "rId": str, "width_emu": int, "height_emu": int, "wrap": str}.
        Raises ValueError if para_id not found.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        # EMU conversion: 1 cm = 914400/2.54 EMU
        width_emu = int(width_cm * 914400 / 2.54)
        height_emu = int(height_cm * 914400 / 2.54)
        h_pos_emu = int(h_pos * 914400 / 2.54)
        v_pos_emu = int(v_pos * 914400 / 2.54)

        src = Path(image_path)
        ext = src.suffix.lstrip(".")

        # Copy image to word/media/
        media_dir = self.workdir / "word" / "media"
        media_dir.mkdir(parents=True, exist_ok=True)
        existing = list(media_dir.glob("image*.*"))
        idx = len(existing) + 1
        filename = f"image{idx}.{ext}"
        shutil.copy2(str(src), str(media_dir / filename))

        # Add relationship
        rels = self._tree("word/_rels/document.xml.rels")
        existing_rids = [r.get("Id") for r in rels.findall(f"{RELS}Relationship")]
        rid_num = (
            max(
                (int(r.replace("rId", "")) for r in existing_rids if r.startswith("rId")),
                default=0,
            )
            + 1
        )
        rid = f"rId{rid_num}"
        rel = etree.SubElement(rels, f"{RELS}Relationship")
        rel.set("Id", rid)
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        rel.set("Target", f"media/{filename}")
        self._mark("word/_rels/document.xml.rels")

        # Ensure content type exists
        ct_tree = self._tree("[Content_Types].xml")
        ct_map = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif"}
        content_type = ct_map.get(ext, f"image/{ext}")
        has_ext = any(d.get("Extension") == ext for d in ct_tree.findall(f"{CT}Default"))
        if not has_ext:
            default_ct = etree.SubElement(ct_tree, f"{CT}Default")
            default_ct.set("Extension", ext)
            default_ct.set("ContentType", content_type)
            self._mark("[Content_Types].xml")

        # Allocate unique drawing id
        drawing_id = self._next_drawing_id(doc)

        # Bare namespace URI strings (for nsmap/attribute values — not Clark notation)
        ns_wp_bare = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        ns_a_bare = "http://schemas.openxmlformats.org/drawingml/2006/main"
        ns_pic_bare = "http://schemas.openxmlformats.org/drawingml/2006/picture"

        # Build wp:anchor element
        anchor = etree.Element(f"{WP}anchor")
        anchor.set("distT", "0")
        anchor.set("distB", "0")
        anchor.set("distL", "114300")
        anchor.set("distR", "114300")
        anchor.set("simplePos", "0")
        anchor.set("locked", "0")
        anchor.set("layoutInCell", "1")
        anchor.set("allowOverlap", "1")
        anchor.set("behindDoc", "0")
        anchor.set("relativeHeight", "251658240")

        simple_pos = etree.SubElement(anchor, f"{WP}simplePos")
        simple_pos.set("x", "0")
        simple_pos.set("y", "0")

        pos_h = etree.SubElement(anchor, f"{WP}positionH")
        pos_h.set("relativeFrom", "page")
        pos_offset_h = etree.SubElement(pos_h, f"{WP}posOffset")
        pos_offset_h.text = str(h_pos_emu)

        pos_v = etree.SubElement(anchor, f"{WP}positionV")
        pos_v.set("relativeFrom", "page")
        pos_offset_v = etree.SubElement(pos_v, f"{WP}posOffset")
        pos_offset_v.text = str(v_pos_emu)

        extent = etree.SubElement(anchor, f"{WP}extent")
        extent.set("cx", str(width_emu))
        extent.set("cy", str(height_emu))

        effect_extent = etree.SubElement(anchor, f"{WP}effectExtent")
        effect_extent.set("l", "0")
        effect_extent.set("t", "0")
        effect_extent.set("r", "0")
        effect_extent.set("b", "0")

        # Wrap element
        if wrap == "square":
            wrap_el = etree.SubElement(anchor, f"{WP}wrapSquare")
            wrap_el.set("wrapText", "bothSides")
        elif wrap == "topbottom":
            etree.SubElement(anchor, f"{WP}wrapTopAndBottom")
        else:  # "none"
            etree.SubElement(anchor, f"{WP}wrapNone")

        doc_pr = etree.SubElement(anchor, f"{WP}docPr")
        doc_pr.set("id", str(drawing_id))
        doc_pr.set("name", f"Image {drawing_id}")

        cnv_frame = etree.SubElement(anchor, f"{WP}cNvGraphicFramePr")
        frame_locks = etree.SubElement(cnv_frame, f"{A}graphicFrameLocks")
        frame_locks.set("noChangeAspect", "1")

        graphic = etree.SubElement(anchor, f"{A}graphic")
        gdata = etree.SubElement(graphic, f"{A}graphicData")
        gdata.set("uri", ns_pic_bare)

        pic = etree.SubElement(gdata, f"{_PIC}pic")
        blip_fill = etree.SubElement(pic, f"{_PIC}blipFill")
        blip = etree.SubElement(blip_fill, f"{A}blip")
        blip.set(f"{R}embed", rid)

        sp_pr = etree.SubElement(pic, f"{_PIC}spPr")
        xfrm = etree.SubElement(sp_pr, f"{A}xfrm")
        off = etree.SubElement(xfrm, f"{A}off")
        off.set("x", "0")
        off.set("y", "0")
        ext_el = etree.SubElement(xfrm, f"{A}ext")
        ext_el.set("cx", str(width_emu))
        ext_el.set("cy", str(height_emu))
        prst_geom = etree.SubElement(sp_pr, f"{A}prstGeom")
        prst_geom.set("prst", "rect")
        etree.SubElement(prst_geom, f"{A}avLst")

        # Wrap anchor in w:drawing > w:r > w:p
        new_para = etree.Element(f"{W}p")
        new_para.set(f"{W14}paraId", self._new_para_id())
        new_para.set(f"{W14}textId", "77777777")
        run = etree.SubElement(new_para, f"{W}r")
        drawing_el = etree.SubElement(run, f"{W}drawing")
        drawing_el.append(anchor)

        para.addnext(new_para)
        self._mark("word/document.xml")

        return {
            "filename": filename,
            "rId": rid,
            "width_emu": width_emu,
            "height_emu": height_emu,
            "wrap": wrap,
        }

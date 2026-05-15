"""Text boxes mixin: insert inline text boxes via wps:wsp."""

from __future__ import annotations

from lxml import etree

from .base import A, W, W14, WP
from .errors import DocxMcpError, ErrCode

# WordprocessingShape namespace constants
WPS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"
_WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
_A_URI = "http://schemas.openxmlformats.org/drawingml/2006/main"
_GRAPHIC_DATA_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

_CM_TO_EMU = 360000  # 1 cm = 914400/2.54 ≈ 360000 EMU


class TextBoxesMixin:
    """Insert inline text boxes into a document."""

    def insert_text_box(
        self,
        para_id: str,
        text: str,
        width_cm: float = 5.0,
        height_cm: float = 2.0,
    ) -> dict:
        """Insert a new paragraph after para_id containing an inline text box.

        Args:
            para_id: paraId of the reference paragraph (insert AFTER this).
            text: Text content for the text box.
            width_cm: Width in centimetres (default 5.0).
            height_cm: Height in centimetres (default 2.0).

        Returns:
            {"para_id": str, "text": str, "width_cm": float, "height_cm": float}
            where para_id is the NEW paragraph's paraId.

        Raises:
            DocxMcpError(PARA_NOT_FOUND): if para_id is not found.
        """
        doc = self._require("word/document.xml")

        ref_para = self._find_para(doc, para_id)
        if ref_para is None:
            raise DocxMcpError(
                ErrCode.PARA_NOT_FOUND,
                f"Paragraph '{para_id}' not found",
                hint="Check the paraId value with list_paragraphs.",
            )

        cx = round(width_cm * _CM_TO_EMU)
        cy = round(height_cm * _CM_TO_EMU)

        # Allocate a unique drawing id
        drawing_id = self._next_drawing_id(doc)  # available via ImagesMixin in DocxDocument

        new_pid = self._new_para_id()

        # ── Build <w:p> ──────────────────────────────────────────────────────
        new_para = etree.Element(f"{W}p")
        new_para.set(f"{W14}paraId", new_pid)
        new_para.set(f"{W14}textId", "77777777")

        # <w:r>
        run = etree.SubElement(new_para, f"{W}r")

        # <w:drawing>
        drawing = etree.SubElement(run, f"{W}drawing")

        # <wp:inline>
        inline = etree.SubElement(drawing, f"{WP}inline")
        inline.set("distT", "0")
        inline.set("distB", "0")
        inline.set("distL", "0")
        inline.set("distR", "0")

        # <wp:extent cx cy>
        extent = etree.SubElement(inline, f"{WP}extent")
        extent.set("cx", str(cx))
        extent.set("cy", str(cy))

        # <wp:effectExtent>
        eff = etree.SubElement(inline, f"{WP}effectExtent")
        eff.set("l", "0")
        eff.set("t", "0")
        eff.set("r", "0")
        eff.set("b", "0")

        # <wp:docPr id name>
        doc_pr = etree.SubElement(inline, f"{WP}docPr")
        doc_pr.set("id", str(drawing_id))
        doc_pr.set("name", f"TextBox {drawing_id}")

        # <wp:cNvGraphicFramePr/>
        etree.SubElement(inline, f"{WP}cNvGraphicFramePr")

        # <a:graphic>
        graphic = etree.SubElement(inline, f"{A}graphic")

        # <a:graphicData uri="...wordprocessingShape">
        graphic_data = etree.SubElement(graphic, f"{A}graphicData")
        graphic_data.set("uri", _GRAPHIC_DATA_URI)

        # <wps:wsp>
        wsp = etree.SubElement(graphic_data, f"{WPS}wsp")

        # <wps:cNvSpPr txBx="1">
        cnv_sp_pr = etree.SubElement(wsp, f"{WPS}cNvSpPr")
        cnv_sp_pr.set("txBx", "1")
        sp_locks = etree.SubElement(cnv_sp_pr, f"{A}spLocks")
        sp_locks.set("noChangeArrowheads", "1")

        # <wps:spPr>
        sp_pr = etree.SubElement(wsp, f"{WPS}spPr")
        xfrm = etree.SubElement(sp_pr, f"{A}xfrm")
        off = etree.SubElement(xfrm, f"{A}off")
        off.set("x", "0")
        off.set("y", "0")
        ext = etree.SubElement(xfrm, f"{A}ext")
        ext.set("cx", str(cx))
        ext.set("cy", str(cy))
        prst_geom = etree.SubElement(sp_pr, f"{A}prstGeom")
        prst_geom.set("prst", "rect")
        etree.SubElement(prst_geom, f"{A}avLst")

        # <wps:txbx>
        txbx = etree.SubElement(wsp, f"{WPS}txbx")
        txbx_content = etree.SubElement(txbx, f"{W}txbxContent")
        inner_p = etree.SubElement(txbx_content, f"{W}p")
        inner_r = etree.SubElement(inner_p, f"{W}r")
        inner_t = etree.SubElement(inner_r, f"{W}t")
        inner_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        inner_t.text = text

        # <wps:bodyPr/>
        etree.SubElement(wsp, f"{WPS}bodyPr")

        # Insert after reference paragraph
        ref_para.addnext(new_para)

        self._mark("word/document.xml")

        return {
            "para_id": new_pid,
            "text": text,
            "width_cm": width_cm,
            "height_cm": height_cm,
        }

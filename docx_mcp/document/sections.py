"""Sections mixin: page breaks, section breaks, section properties."""

from __future__ import annotations

from lxml import etree

from .base import W14, W


class SectionsMixin:
    """Section and page break operations."""

    def add_page_break(self, para_id: str) -> dict:
        """Insert a page break after a paragraph.

        Creates a new paragraph containing a page break element.

        Args:
            para_id: paraId of the paragraph to insert after.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        new_para = etree.Element(f"{W}p")
        new_pid = self._new_para_id()
        new_para.set(f"{W14}paraId", new_pid)
        new_para.set(f"{W14}textId", "77777777")
        run = etree.SubElement(new_para, f"{W}r")
        br = etree.SubElement(run, f"{W}br")
        br.set(f"{W}type", "page")

        para.addnext(new_para)
        self._mark("word/document.xml")

        return {"para_id": new_pid}

    def add_section_break(
        self,
        para_id: str,
        break_type: str = "nextPage",
    ) -> dict:
        """Add a section break at a paragraph.

        Inserts w:sectPr inside the paragraph's w:pPr to mark it as the last
        paragraph of its section.

        Args:
            para_id: paraId of the paragraph to place the section break on.
            break_type: "nextPage", "continuous", "evenPage", or "oddPage".
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ppr = para.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.SubElement(para, f"{W}pPr")
            # Move pPr to be first child
            para.remove(ppr)
            para.insert(0, ppr)

        sect_pr = etree.SubElement(ppr, f"{W}sectPr")
        type_el = etree.SubElement(sect_pr, f"{W}type")
        type_el.set(f"{W}val", break_type)

        self._mark("word/document.xml")

        return {"para_id": para_id, "break_type": break_type}

    def set_section_properties(
        self,
        *,
        para_id: str | None = None,
        width: int | None = None,
        height: int | None = None,
        orientation: str | None = None,
        margin_top: int | None = None,
        margin_bottom: int | None = None,
        margin_left: int | None = None,
        margin_right: int | None = None,
    ) -> dict:
        """Modify section properties (page size, orientation, margins).

        Args:
            para_id: paraId of paragraph with section break. None = body section.
            width: Page width in DXA (twips). 12240 = 8.5".
            height: Page height in DXA. 15840 = 11".
            orientation: "portrait" or "landscape".
            margin_top: Top margin in DXA.
            margin_bottom: Bottom margin in DXA.
            margin_left: Left margin in DXA.
            margin_right: Right margin in DXA.
        """
        doc = self._require("word/document.xml")

        if para_id is not None:
            # Paragraph-level section
            para = self._find_para(doc, para_id)
            if para is None:
                raise ValueError(f"Paragraph '{para_id}' not found")
            ppr = para.find(f"{W}pPr")
            sect_pr = ppr.find(f"{W}sectPr") if ppr is not None else None
            if sect_pr is None:
                raise ValueError(
                    f"No section break on paragraph '{para_id}'. Use add_section_break first."
                )
        else:
            # Body-level section
            body = doc.find(f"{W}body")
            sect_pr = body.find(f"{W}sectPr")
            if sect_pr is None:
                sect_pr = etree.SubElement(body, f"{W}sectPr")

        # Page size
        if width is not None or height is not None or orientation is not None:
            pg_sz = sect_pr.find(f"{W}pgSz")
            if pg_sz is None:
                pg_sz = etree.SubElement(sect_pr, f"{W}pgSz")
            if width is not None:
                pg_sz.set(f"{W}w", str(width))
            if height is not None:
                pg_sz.set(f"{W}h", str(height))
            if orientation is not None:
                pg_sz.set(f"{W}orient", orientation)

        # Margins
        margin_vals = {
            "top": margin_top,
            "bottom": margin_bottom,
            "left": margin_left,
            "right": margin_right,
        }
        if any(v is not None for v in margin_vals.values()):
            pg_mar = sect_pr.find(f"{W}pgMar")
            if pg_mar is None:
                pg_mar = etree.SubElement(sect_pr, f"{W}pgMar")
            for attr, val in margin_vals.items():
                if val is not None:
                    pg_mar.set(f"{W}{attr}", str(val))

        self._mark("word/document.xml")

        # Build response from current state
        result: dict = {}
        pg_sz = sect_pr.find(f"{W}pgSz")
        if pg_sz is not None:
            w_val = pg_sz.get(f"{W}w")
            h_val = pg_sz.get(f"{W}h")
            if w_val:
                result["width"] = int(w_val)
            if h_val:
                result["height"] = int(h_val)
            orient = pg_sz.get(f"{W}orient")
            if orient:
                result["orientation"] = orient

        pg_mar = sect_pr.find(f"{W}pgMar")
        if pg_mar is not None:
            for attr in ("top", "bottom", "left", "right"):
                val = pg_mar.get(f"{W}{attr}")
                if val:
                    result[f"margin_{attr}"] = int(val)

        return result

    def set_document_language(self, language_code: str) -> dict:
        """Set the default document language via the default paragraph style's rPr.

        Writes ``<w:lang w:val="language_code"/>`` into the ``<w:rPr>`` of
        the ``w:style[@w:type='paragraph'][@w:default='1']`` element in
        ``word/styles.xml``.

        Args:
            language_code: BCP-47 language tag, e.g. "en-US", "fr-FR", "de-DE".

        Returns:
            {"language": language_code}
        """
        styles = self._require("word/styles.xml")

        # Find the default paragraph style
        default_style = None
        for s in styles.findall(f"{W}style"):
            if s.get(f"{W}type") == "paragraph" and s.get(f"{W}default") == "1":
                default_style = s
                break

        if default_style is None:
            raise RuntimeError("No default paragraph style found in word/styles.xml")

        # Get or create w:rPr
        rpr = default_style.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.SubElement(default_style, f"{W}rPr")

        # Get or create w:lang
        lang = rpr.find(f"{W}lang")
        if lang is None:
            lang = etree.SubElement(rpr, f"{W}lang")

        lang.set(f"{W}val", language_code)
        self._mark("word/styles.xml")
        return {"language": language_code}

    # ── Convenience wrappers ─────────────────────────────────────────────────

    @staticmethod
    def _mm_to_dxa(mm: float) -> int:
        """Convert millimetres to DXA (twips). 1 inch = 1440 twips = 25.4 mm."""
        return round(mm * 1440 / 25.4)

    def set_page_size(
        self,
        width_mm: float,
        height_mm: float,
        *,
        para_id: str | None = None,
    ) -> dict:
        """Set page size from millimetre values.

        Args:
            width_mm: Page width in mm (e.g. 210 for A4).
            height_mm: Page height in mm (e.g. 297 for A4).
            para_id: paraId of paragraph with section break. None = body section.
        """
        width_dxa = self._mm_to_dxa(width_mm)
        height_dxa = self._mm_to_dxa(height_mm)
        self.set_section_properties(para_id=para_id, width=width_dxa, height=height_dxa)
        return {
            "width_mm": width_mm,
            "height_mm": height_mm,
            "width_dxa": width_dxa,
            "height_dxa": height_dxa,
        }

    def set_page_margins(
        self,
        *,
        top_mm: float | None = None,
        bottom_mm: float | None = None,
        left_mm: float | None = None,
        right_mm: float | None = None,
        para_id: str | None = None,
    ) -> dict:
        """Set page margins from millimetre values.

        Args:
            top_mm: Top margin in mm. None = unchanged.
            bottom_mm: Bottom margin in mm. None = unchanged.
            left_mm: Left margin in mm. None = unchanged.
            right_mm: Right margin in mm. None = unchanged.
            para_id: paraId of paragraph with section break. None = body section.
        """
        self.set_section_properties(
            para_id=para_id,
            margin_top=self._mm_to_dxa(top_mm) if top_mm is not None else None,
            margin_bottom=self._mm_to_dxa(bottom_mm) if bottom_mm is not None else None,
            margin_left=self._mm_to_dxa(left_mm) if left_mm is not None else None,
            margin_right=self._mm_to_dxa(right_mm) if right_mm is not None else None,
        )
        margins: dict = {}
        if top_mm is not None:
            margins["top"] = top_mm
        if bottom_mm is not None:
            margins["bottom"] = bottom_mm
        if left_mm is not None:
            margins["left"] = left_mm
        if right_mm is not None:
            margins["right"] = right_mm
        return {"margins_mm": margins}

    def set_page_orientation(
        self,
        orientation: str,
        *,
        para_id: str | None = None,
    ) -> dict:
        """Set page orientation, swapping width/height if needed.

        Args:
            orientation: "portrait" or "landscape".
            para_id: paraId of paragraph with section break. None = body section.
        """
        if orientation not in ("portrait", "landscape"):
            raise ValueError(
                f"orientation must be 'portrait' or 'landscape', got {orientation!r}"
            )

        doc = self._require("word/document.xml")

        if para_id is not None:
            para = self._find_para(doc, para_id)
            if para is None:
                raise ValueError(f"Paragraph '{para_id}' not found")
            ppr = para.find(f"{W}pPr")
            sect_pr = ppr.find(f"{W}sectPr") if ppr is not None else None
            if sect_pr is None:
                raise ValueError(
                    f"No section break on paragraph '{para_id}'. Use add_section_break first."
                )
        else:
            body = doc.find(f"{W}body")
            sect_pr = body.find(f"{W}sectPr")
            if sect_pr is None:
                sect_pr = etree.SubElement(body, f"{W}sectPr")

        # Get or create pgSz
        pg_sz = sect_pr.find(f"{W}pgSz")
        if pg_sz is None:
            pg_sz = etree.SubElement(sect_pr, f"{W}pgSz")

        w_val = int(pg_sz.get(f"{W}w", "0") or "0")
        h_val = int(pg_sz.get(f"{W}h", "0") or "0")

        if orientation == "landscape":
            if w_val < h_val or (w_val == 0 and h_val == 0):
                # Swap
                new_w, new_h = h_val, w_val
            else:
                new_w, new_h = w_val, h_val
            pg_sz.set(f"{W}w", str(new_w))
            pg_sz.set(f"{W}h", str(new_h))
            pg_sz.set(f"{W}orient", "landscape")
        else:  # portrait
            if w_val > h_val:
                # Swap
                new_w, new_h = h_val, w_val
            else:
                new_w, new_h = w_val, h_val
            pg_sz.set(f"{W}w", str(new_w))
            pg_sz.set(f"{W}h", str(new_h))
            # Set to portrait explicitly (or remove orient attr)
            pg_sz.set(f"{W}orient", "portrait")

        self._mark("word/document.xml")

        return {
            "orientation": orientation,
            "width_dxa": int(pg_sz.get(f"{W}w", "0")),
            "height_dxa": int(pg_sz.get(f"{W}h", "0")),
        }

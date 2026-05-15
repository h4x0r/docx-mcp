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

    def set_different_first_page(self, section_index: int, enabled: bool) -> dict:
        """Enable or disable a different first-page header/footer for a section.

        Args:
            section_index: Zero-based section index.
            enabled: True to add w:titlePg (enable), False to remove it.

        Returns:
            {"section_index": section_index, "different_first_page": enabled}

        Raises:
            ValueError: If section_index is out of range.
        """
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        sectprs = self._collect_sectprs(body)

        # Auto-create body-level sectPr if document has none yet
        if len(sectprs) == 0:
            sect_pr_el = etree.SubElement(body, f"{W}sectPr")
            sectprs = [(sect_pr_el, True)]

        if section_index < 0 or section_index >= len(sectprs):
            raise ValueError(
                f"section_index {section_index} out of range (0..{len(sectprs) - 1})"
            )
        sect_pr, _ = sectprs[section_index]

        existing = sect_pr.find(f"{W}titlePg")
        if enabled:
            if existing is None:
                etree.SubElement(sect_pr, f"{W}titlePg")
        else:
            if existing is not None:
                sect_pr.remove(existing)

        self._mark("word/document.xml")
        return {"section_index": section_index, "different_first_page": enabled}

    def set_odd_even_headers(self, enabled: bool) -> dict:
        """Enable or disable different odd/even page headers (document-level setting).

        Args:
            enabled: True to add w:evenAndOddHeaders to settings.xml, False to remove it.

        Returns:
            {"odd_even_headers": enabled}
        """
        import contextlib

        from .base import CT, RELS

        settings = self._tree("word/settings.xml")
        if settings is None and not enabled:
            # Nothing to remove — return early without creating settings.xml
            return {"odd_even_headers": False}
        if settings is None:
            settings = etree.Element(
                f"{W}settings",
                nsmap={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
            )
            self._trees["word/settings.xml"] = settings

            fp = self.workdir / "word" / "settings.xml"
            fp.parent.mkdir(parents=True, exist_ok=True)
            etree.ElementTree(settings).write(
                str(fp), xml_declaration=True, encoding="UTF-8"
            )

            ct = self._tree("[Content_Types].xml")
            if ct is not None:
                existing_parts = {e.get("PartName") for e in ct.findall(f"{CT}Override")}
                if "/word/settings.xml" not in existing_parts:
                    ov = etree.SubElement(ct, f"{CT}Override")
                    ov.set("PartName", "/word/settings.xml")
                    ov.set(
                        "ContentType",
                        "application/vnd.openxmlformats-officedocument"
                        ".wordprocessingml.settings+xml",
                    )
                    self._mark("[Content_Types].xml")

            rels = self._tree("word/_rels/document.xml.rels")
            if rels is not None:
                existing_targets = {
                    r.get("Target") for r in rels.findall(f"{RELS}Relationship")
                }
                if "settings.xml" not in existing_targets:
                    max_rid = 0
                    for r in rels.findall(f"{RELS}Relationship"):
                        rid = r.get("Id", "")
                        if rid.startswith("rId"):
                            with contextlib.suppress(ValueError):
                                max_rid = max(max_rid, int(rid[3:]))
                    rel = etree.SubElement(rels, f"{RELS}Relationship")
                    rel.set("Id", f"rId{max_rid + 1}")
                    rel.set(
                        "Type",
                        "http://schemas.openxmlformats.org/officeDocument"
                        "/2006/relationships/settings",
                    )
                    rel.set("Target", "settings.xml")
                    self._mark("word/_rels/document.xml.rels")

        tag = f"{W}evenAndOddHeaders"
        existing = settings.find(tag)
        if enabled:
            if existing is None:
                etree.SubElement(settings, tag)
                self._mark("word/settings.xml")
        else:
            if existing is not None:
                settings.remove(existing)
                self._mark("word/settings.xml")
        return {"odd_even_headers": enabled}

    # ── Section enumeration helpers ──────────────────────────────────────────

    def _collect_sectprs(self, body):
        """Return list of (sectPr, is_final) tuples in document order."""
        result = []
        for child in body:
            if child.tag == f"{W}p":
                ppr = child.find(f"{W}pPr")
                if ppr is not None:
                    spr = ppr.find(f"{W}sectPr")
                    if spr is not None:
                        result.append((spr, False))
        final = body.find(f"{W}sectPr")
        if final is not None:
            result.append((final, True))
        return result

    @staticmethod
    def _parse_sectpr(sectpr, index: int, is_final: bool) -> dict:
        """Extract structured info from a sectPr element."""
        W_NS = W  # alias for use inside staticmethod (module-level constant)

        # break_type
        if is_final:
            break_type = ""
        else:
            type_el = sectpr.find(f"{W_NS}type")
            break_type = type_el.get(f"{W_NS}val", "continuous") if type_el is not None else "continuous"  # noqa: E501

        # page size
        pg_sz = sectpr.find(f"{W_NS}pgSz")
        if pg_sz is not None:
            page_width = int(pg_sz.get(f"{W_NS}w", "0") or "0")
            page_height = int(pg_sz.get(f"{W_NS}h", "0") or "0")
            orient_val = pg_sz.get(f"{W_NS}orient", "")
            orientation = "landscape" if orient_val == "landscape" else "portrait"
        else:
            page_width = 0
            page_height = 0
            orientation = "portrait"

        # columns
        cols_el = sectpr.find(f"{W_NS}cols")
        if cols_el is not None:
            num_attr = cols_el.get(f"{W_NS}num")
            if num_attr:
                columns = int(num_attr)
            else:
                col_children = [c for c in cols_el if c.tag == f"{W_NS}col"]
                columns = len(col_children) if col_children else 1
        else:
            columns = 1

        # margins
        pg_mar = sectpr.find(f"{W_NS}pgMar")
        if pg_mar is not None:
            margin_top = int(pg_mar.get(f"{W_NS}top", "0") or "0")
            margin_bottom = int(pg_mar.get(f"{W_NS}bottom", "0") or "0")
        else:
            margin_top = 0
            margin_bottom = 0

        return {
            "index": index,
            "break_type": break_type,
            "page_width": page_width,
            "page_height": page_height,
            "orientation": orientation,
            "columns": columns,
            "margin_top": margin_top,
            "margin_bottom": margin_bottom,
        }

    def get_sections(self) -> list[dict]:
        """Return structured info about each section in document order.

        Each section dict contains index, break_type, page_width, page_height,
        orientation, columns, margin_top, margin_bottom.
        """
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        sectprs = self._collect_sectprs(body)
        return [self._parse_sectpr(spr, i, is_final) for i, (spr, is_final) in enumerate(sectprs)]

    def set_section_columns(
        self,
        section_index: int,
        num_columns: int,
        equal_width: bool = True,
    ) -> dict:
        """Set multi-column layout on a section.

        Args:
            section_index: Zero-based section index.
            num_columns: Number of columns.
            equal_width: If True, set equalWidth="1" and remove w:col children.
        """
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        sectprs = self._collect_sectprs(body)
        if section_index < 0 or section_index >= len(sectprs):
            raise ValueError(
                f"section_index {section_index} out of range (0..{len(sectprs) - 1})"
            )
        sect_pr, _ = sectprs[section_index]

        # Remove existing w:cols
        old_cols = sect_pr.find(f"{W}cols")
        if old_cols is not None:
            sect_pr.remove(old_cols)

        cols_el = etree.SubElement(sect_pr, f"{W}cols")
        cols_el.set(f"{W}num", str(num_columns))
        if equal_width:
            cols_el.set(f"{W}equalWidth", "1")
            # Remove any w:col children (already fresh element, none exist)

        self._mark("word/document.xml")
        return {
            "section_index": section_index,
            "num_columns": num_columns,
            "equal_width": equal_width,
        }

    def delete_section_break(self, para_id: str) -> dict:
        """Remove the section break from a paragraph.

        Args:
            para_id: w14:paraId of the paragraph containing the section break.

        Returns:
            {"para_id": para_id, "deleted": True}

        Raises:
            ValueError: If no section break found in the paragraph.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ppr = para.find(f"{W}pPr")
        sect_pr = ppr.find(f"{W}sectPr") if ppr is not None else None
        if sect_pr is None:
            raise ValueError("No section break found in paragraph")

        ppr.remove(sect_pr)
        # Clean up empty pPr
        if ppr is not None and len(ppr) == 0:
            para.remove(ppr)

        self._mark("word/document.xml")
        return {"para_id": para_id, "deleted": True}

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

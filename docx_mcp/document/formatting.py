"""Formatting mixin: apply character formatting with tracked changes."""

from __future__ import annotations

from lxml import etree

from .base import W, W14, _now_iso, _preserve


class FormattingMixin:
    """Apply formatting to text with tracked change markup."""

    def set_formatting(
        self,
        para_id: str,
        text: str,
        *,
        bold: bool = False,
        italic: bool = False,
        underline: str | None = None,
        color: str | None = None,
        author: str = "Claude",
    ) -> dict:
        """Apply character formatting to a text substring with tracked changes.

        Finds the run containing `text`, splits if needed, clones the original
        rPr as rPrChange, and applies the new formatting properties.

        Args:
            para_id: Target paragraph paraId.
            text: Exact text to format.
            bold: Apply bold.
            italic: Apply italic.
            underline: Underline style (e.g., "single", "double").
            color: Font color as hex (e.g., "FF0000").
            author: Author name for the revision.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        now = _now_iso()
        cid = self._next_markup_id(doc)

        for run_el in list(para.findall(f"{W}r")):
            t_el = run_el.find(f"{W}t")
            if t_el is None or t_el.text is None:
                continue
            full = t_el.text
            if text not in full:
                continue

            idx = full.index(text)
            old_rpr = run_el.find(f"{W}rPr")
            rpr_bytes = etree.tostring(old_rpr) if old_rpr is not None else None
            parent = run_el.getparent()
            pos = list(parent).index(run_el)
            parent.remove(run_el)

            insert_at = pos

            # Text before match — plain run
            if idx > 0:
                before = self._make_run(full[:idx], rpr_bytes)
                parent.insert(insert_at, before)
                insert_at += 1

            # Formatted run with rPrChange for tracked change
            fmt_run = etree.Element(f"{W}r")
            new_rpr = etree.SubElement(fmt_run, f"{W}rPr")

            # Apply requested formatting
            if bold:
                etree.SubElement(new_rpr, f"{W}b")
            if italic:
                etree.SubElement(new_rpr, f"{W}i")
            if underline:
                u = etree.SubElement(new_rpr, f"{W}u")
                u.set(f"{W}val", underline)
            if color:
                c = etree.SubElement(new_rpr, f"{W}color")
                c.set(f"{W}val", color)

            # Record original formatting as rPrChange
            rpr_change = etree.SubElement(new_rpr, f"{W}rPrChange")
            rpr_change.set(f"{W}id", str(cid))
            rpr_change.set(f"{W}author", author)
            rpr_change.set(f"{W}date", now)
            if rpr_bytes:
                rpr_change.append(etree.fromstring(rpr_bytes))
            else:
                etree.SubElement(rpr_change, f"{W}rPr")

            fmt_t = etree.SubElement(fmt_run, f"{W}t")
            _preserve(fmt_t, text)
            parent.insert(insert_at, fmt_run)
            insert_at += 1

            # Text after match — plain run
            end = idx + len(text)
            if end < len(full):
                after = self._make_run(full[end:], rpr_bytes)
                parent.insert(insert_at, after)

            self._mark("word/document.xml")
            return {"formatted": True}

        raise ValueError(f"Text '{text}' not found in a single run of paragraph '{para_id}'.")

    def insert_paragraph(
        self,
        after_para_id: str,
        text: str,
        style: str | None = None,
    ) -> dict:
        doc = self._require("word/document.xml")
        target = self._find_para(doc, after_para_id)
        if target is None:
            raise ValueError(f"Paragraph '{after_para_id}' not found")

        new_pid = self._new_para_id()
        new_p = etree.Element(f"{W}p")
        new_p.set(f"{W14}paraId", new_pid)

        if style is not None:
            ppr = etree.SubElement(new_p, f"{W}pPr")
            ps = etree.SubElement(ppr, f"{W}pStyle")
            ps.set(f"{W}val", style)

        run = etree.SubElement(new_p, f"{W}r")
        t_el = etree.SubElement(run, f"{W}t")
        _preserve(t_el, text)

        target.addnext(new_p)
        self._mark("word/document.xml")
        return {"para_id": new_pid, "text": text}

    def update_paragraph(
        self,
        para_id: str,
        text: str | None = None,
        style: str | None = None,
    ) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        if text is not None:
            for run in list(para.findall(f"{W}r")):
                para.remove(run)
            run = etree.SubElement(para, f"{W}r")
            t_el = etree.SubElement(run, f"{W}t")
            _preserve(t_el, text)

        if style is not None:
            ppr = para.find(f"{W}pPr")
            if ppr is None:
                ppr = etree.Element(f"{W}pPr")
                para.insert(0, ppr)
            pstyle = ppr.find(f"{W}pStyle")
            if pstyle is None:
                pstyle = etree.Element(f"{W}pStyle")
                ppr.insert(0, pstyle)
            pstyle.set(f"{W}val", style)

        if text is not None or style is not None:
            self._mark("word/document.xml")
        return {"para_id": para_id}

    def delete_paragraph(self, para_id: str) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")
        para.getparent().remove(para)
        self._mark("word/document.xml")
        return {"deleted": para_id}

    def set_paragraph_border(
        self,
        para_id: str,
        sides: list[str],
        color: str = "000000",
        size: int = 4,
    ) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ppr = para.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            para.insert(0, ppr)

        pbdr = ppr.find(f"{W}pBdr")
        if pbdr is None:
            pbdr = etree.Element(f"{W}pBdr")
            # Insert before w:shd to maintain CT_PPr schema order
            shd = ppr.find(f"{W}shd")
            if shd is not None:
                shd.addprevious(pbdr)
            else:
                ppr.append(pbdr)

        for side in sides:
            el = pbdr.find(f"{W}{side}")
            if el is None:
                el = etree.SubElement(pbdr, f"{W}{side}")
            el.set(f"{W}val", "single")
            el.set(f"{W}sz", str(size))
            el.set(f"{W}space", "0")
            el.set(f"{W}color", color)

        self._mark("word/document.xml")
        return {"para_id": para_id, "sides": sides}

    def set_paragraph_shading(
        self,
        para_id: str,
        fill_color: str,
        pattern: str = "clear",
    ) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        ppr = para.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            para.insert(0, ppr)

        shd = ppr.find(f"{W}shd")
        if shd is None:
            shd = etree.SubElement(ppr, f"{W}shd")
        shd.set(f"{W}val", pattern)
        shd.set(f"{W}color", "auto")
        shd.set(f"{W}fill", fill_color)

        self._mark("word/document.xml")
        return {"para_id": para_id, "fill_color": fill_color}

    def _get_run(self, para, run_idx: int):
        runs = para.findall(f"{W}r")
        if run_idx < 0 or run_idx >= len(runs):
            raise IndexError(f"Run index {run_idx} out of range (have {len(runs)})")
        return runs[run_idx]

    def _upsert_rpr(self, run):
        rpr = run.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.Element(f"{W}rPr")
            run.insert(0, rpr)
        return rpr

    def get_runs(self, para_id: str) -> list:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        result = []
        for idx, run in enumerate(para.findall(f"{W}r")):
            text = "".join(t.text or "" for t in run.findall(f"{W}t"))
            rpr = run.find(f"{W}rPr")

            bold = False
            italic = False
            font = None
            size_pt = None
            color = None

            if rpr is not None:
                b_el = rpr.find(f"{W}b")
                bold = b_el is not None and b_el.get(f"{W}val") not in ("0", "false", "off")

                i_el = rpr.find(f"{W}i")
                italic = i_el is not None and i_el.get(f"{W}val") not in ("0", "false", "off")

                rfonts = rpr.find(f"{W}rFonts")
                if rfonts is not None:
                    font = rfonts.get(f"{W}ascii")

                sz_el = rpr.find(f"{W}sz")
                if sz_el is not None:
                    val = sz_el.get(f"{W}val")
                    if val is not None:
                        size_pt = int(val) / 2

                color_el = rpr.find(f"{W}color")
                if color_el is not None:
                    color = color_el.get(f"{W}val")

            result.append({
                "run_idx": idx,
                "text": text,
                "bold": bold,
                "italic": italic,
                "font": font,
                "size_pt": size_pt,
                "color": color,
            })

        return result

    def set_run_font(self, para_id: str, run_idx: int, font_name: str) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        rfonts = rpr.find(f"{W}rFonts")
        if rfonts is None:
            rfonts = etree.SubElement(rpr, f"{W}rFonts")
        rfonts.set(f"{W}ascii", font_name)
        rfonts.set(f"{W}hAnsi", font_name)
        rfonts.set(f"{W}cs", font_name)

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "font": font_name}

    def set_run_color(self, para_id: str, run_idx: int, color: str) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        color_el = rpr.find(f"{W}color")
        if color_el is None:
            color_el = etree.SubElement(rpr, f"{W}color")
        color_el.set(f"{W}val", color)

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "color": color}

    def set_run_size(self, para_id: str, run_idx: int, size_pt: float) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        half_pts = str(round(size_pt * 2))

        sz_el = rpr.find(f"{W}sz")
        if sz_el is None:
            sz_el = etree.SubElement(rpr, f"{W}sz")
        sz_el.set(f"{W}val", half_pts)

        sz_cs_el = rpr.find(f"{W}szCs")
        if sz_cs_el is None:
            sz_cs_el = etree.SubElement(rpr, f"{W}szCs")
        sz_cs_el.set(f"{W}val", half_pts)

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "size_pt": size_pt}

    def set_character_spacing(self, para_id: str, run_idx: int, spacing_pt: float) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        spacing_el = rpr.find(f"{W}spacing")
        if spacing_el is None:
            spacing_el = etree.SubElement(rpr, f"{W}spacing")
        spacing_el.set(f"{W}val", str(round(spacing_pt * 20)))

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "spacing_pt": spacing_pt}

    def set_character_position(self, para_id: str, run_idx: int, position_pt: float) -> dict:
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        run = self._get_run(para, run_idx)
        rpr = self._upsert_rpr(run)

        pos_el = rpr.find(f"{W}position")
        if pos_el is None:
            pos_el = etree.SubElement(rpr, f"{W}position")
        pos_el.set(f"{W}val", str(round(position_pt * 2)))

        self._mark("word/document.xml")
        return {"para_id": para_id, "run_idx": run_idx, "position_pt": position_pt}

    def find_replace_formatted(
        self,
        find: str,
        replace: str,
        *,
        bold: bool | None = None,
        italic: bool | None = None,
        color: str | None = None,
        size_pt: float | None = None,
    ) -> dict:
        if not find:
            raise ValueError("find must be non-empty")

        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        if body is None:
            body = doc

        count = 0

        for para in body.iter(f"{W}p"):
            changed = True
            while changed:
                changed = False
                for run_el in list(para):
                    if run_el.tag != f"{W}r":
                        continue
                    t_el = run_el.find(f"{W}t")
                    if t_el is None or t_el.text is None:
                        continue
                    full = t_el.text
                    if find not in full:
                        continue

                    idx = full.index(find)
                    old_rpr = run_el.find(f"{W}rPr")
                    rpr_bytes = etree.tostring(old_rpr) if old_rpr is not None else None
                    parent = run_el.getparent()
                    pos = list(parent).index(run_el)
                    parent.remove(run_el)

                    insert_at = pos

                    before_text = full[:idx]
                    if before_text:
                        before_run = self._make_run(before_text, rpr_bytes)
                        parent.insert(insert_at, before_run)
                        insert_at += 1

                    fmt_run = etree.Element(f"{W}r")
                    has_fmt = bold is not None or italic is not None or color is not None or size_pt is not None
                    new_rpr = etree.SubElement(fmt_run, f"{W}rPr") if has_fmt else None

                    if bold is True:
                        etree.SubElement(new_rpr, f"{W}b")
                    elif bold is False:
                        b_el = etree.SubElement(new_rpr, f"{W}b")
                        b_el.set(f"{W}val", "0")

                    if italic is True:
                        etree.SubElement(new_rpr, f"{W}i")
                    elif italic is False:
                        i_el = etree.SubElement(new_rpr, f"{W}i")
                        i_el.set(f"{W}val", "0")

                    if color is not None:
                        c_el = etree.SubElement(new_rpr, f"{W}color")
                        c_el.set(f"{W}val", color)

                    if size_pt is not None:
                        half = str(round(size_pt * 2))
                        sz_el = etree.SubElement(new_rpr, f"{W}sz")
                        sz_el.set(f"{W}val", half)
                        sz_cs_el = etree.SubElement(new_rpr, f"{W}szCs")
                        sz_cs_el.set(f"{W}val", half)

                    repl_t = etree.SubElement(fmt_run, f"{W}t")
                    _preserve(repl_t, replace)
                    parent.insert(insert_at, fmt_run)
                    insert_at += 1

                    after_text = full[idx + len(find):]
                    if after_text:
                        after_run = self._make_run(after_text, rpr_bytes)
                        parent.insert(insert_at, after_run)

                    count += 1
                    changed = True
                    break

        if count > 0:
            self._mark("word/document.xml")

        return {"find": find, "replace": replace, "count": count}

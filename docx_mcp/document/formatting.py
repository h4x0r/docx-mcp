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
            pbdr = etree.SubElement(ppr, f"{W}pBdr")

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

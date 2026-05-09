"""FieldsMixin — Word complex field insertion and management."""
from __future__ import annotations

from lxml import etree

from .base import W, W14, XML_SPACE, _preserve
from .errors import DocxMcpError, ErrCode


class FieldsMixin:
    """Insert and manage Word complex fields (fldChar / instrText)."""

    def add_field(
        self, para_id: str, field_code: str, cached_value: str = ""
    ) -> dict:
        """Insert a Word complex field at the end of a paragraph.

        Inserts the five-run structure:
          begin fldChar → instrText → separate fldChar → cached w:t → end fldChar

        Args:
            para_id: w14:paraId of the target paragraph.
            field_code: The field instruction text (e.g. "PAGE", "SEQ Figure").
            cached_value: Display text cached in the document. Defaults to "0".

        Returns:
            {"code": str, "cached_value": str, "para_id": str}

        Raises:
            DocxMcpError(PARA_NOT_FOUND) if para_id not found.
        """
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        para = self._find_para(body, para_id)
        if para is None:
            raise DocxMcpError(
                ErrCode.PARA_NOT_FOUND,
                f"Paragraph with paraId '{para_id}' not found.",
            )

        stored_value = cached_value if cached_value else "0"

        # Run 1: begin fldChar (dirty=true so Word recalculates on open)
        r_begin = etree.SubElement(para, f"{W}r")
        fc_begin = etree.SubElement(r_begin, f"{W}fldChar")
        fc_begin.set(f"{W}fldCharType", "begin")
        fc_begin.set(f"{W}dirty", "true")

        # Run 2: instrText
        r_instr = etree.SubElement(para, f"{W}r")
        instr = etree.SubElement(r_instr, f"{W}instrText")
        instr.set(XML_SPACE, "preserve")
        instr.text = f" {field_code} "

        # Run 3: separate fldChar
        r_sep = etree.SubElement(para, f"{W}r")
        fc_sep = etree.SubElement(r_sep, f"{W}fldChar")
        fc_sep.set(f"{W}fldCharType", "separate")

        # Run 4: cached value text
        r_cached = etree.SubElement(para, f"{W}r")
        t_cached = etree.SubElement(r_cached, f"{W}t")
        _preserve(t_cached, stored_value)

        # Run 5: end fldChar
        r_end = etree.SubElement(para, f"{W}r")
        fc_end = etree.SubElement(r_end, f"{W}fldChar")
        fc_end.set(f"{W}fldCharType", "end")

        self._mark("word/document.xml")
        return {"code": field_code, "cached_value": stored_value, "para_id": para_id}

    def update_fields(self) -> dict:
        """Set w:dirty='true' on all begin fldChar elements in the document.

        Returns:
            {"updated_count": int}
        """
        doc = self._require("word/document.xml")
        count = 0
        for fc in doc.iter(f"{W}fldChar"):
            if fc.get(f"{W}fldCharType") == "begin":
                fc.set(f"{W}dirty", "true")
                count += 1
        if count:
            self._mark("word/document.xml")
        return {"updated_count": count}

    def list_fields(self) -> list[dict]:
        """Walk document XML and return all complex fields.

        Each entry contains:
          - code: stripped instrText content
          - cached_value: text from runs between separate and end fldChar
          - para_id: w14:paraId of the enclosing <w:p>, or None

        Returns:
            List of {"code": str, "cached_value": str, "para_id": str | None}
        """
        doc = self._tree("word/document.xml")
        if doc is None:
            return []

        results: list[dict] = []

        # Find all begin fldChar runs
        for fld_char in doc.iter(f"{W}fldChar"):
            if fld_char.get(f"{W}fldCharType") != "begin":
                continue

            # The begin fldChar is inside a w:r; we need the w:p parent to find siblings
            begin_run = fld_char.getparent()
            if begin_run is None:
                continue
            para = begin_run.getparent()
            if para is None:
                continue

            # Get para_id from the enclosing w:p (walk up if needed)
            ancestor = para
            while ancestor is not None and ancestor.tag != f"{W}p":
                ancestor = ancestor.getparent()
            para_id = ancestor.get(f"{W14}paraId") if ancestor is not None else None

            # Walk siblings of begin_run to collect instrText and cached value
            children = list(para)
            try:
                start_idx = children.index(begin_run)
            except ValueError:
                continue

            code_parts: list[str] = []
            cached_parts: list[str] = []
            state = "instr"  # "instr" -> after begin, "cached" -> after separate

            for sibling in children[start_idx + 1:]:
                # Check for fldChar in this run
                fc = sibling.find(f"{W}fldChar")
                if fc is not None:
                    ftype = fc.get(f"{W}fldCharType")
                    if ftype == "separate":
                        state = "cached"
                        continue
                    elif ftype == "end":
                        break  # done with this field

                if state == "instr":
                    for it in sibling.iter(f"{W}instrText"):
                        if it.text:
                            code_parts.append(it.text)
                elif state == "cached":
                    for t in sibling.iter(f"{W}t"):
                        if t.text:
                            cached_parts.append(t.text)

            results.append({
                "code": "".join(code_parts).strip(),
                "cached_value": "".join(cached_parts),
                "para_id": para_id,
            })

        return results

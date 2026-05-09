"""FieldsMixin — Word complex field insertion and management."""
from __future__ import annotations

from lxml import etree

from .base import W, W14, _preserve
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
        instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
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
        doc = self._require("word/document.xml")
        results: list[dict] = []

        in_field = False
        collecting_cached = False
        current: dict | None = None
        current_para_id: str | None = None

        for elem in doc.iter():
            tag = elem.tag

            # Track enclosing paragraph para_id
            if tag == f"{W}p":
                current_para_id = elem.get(f"{W14}paraId")
                continue

            if tag == f"{W}fldChar":
                fc_type = elem.get(f"{W}fldCharType")
                if fc_type == "begin":
                    in_field = True
                    collecting_cached = False
                    current = {
                        "code": "",
                        "cached_value": "",
                        "para_id": current_para_id,
                    }
                elif fc_type == "separate" and in_field:
                    collecting_cached = True
                elif fc_type == "end" and in_field:
                    if current is not None:
                        current["code"] = current["code"].strip()
                        results.append(current)
                    current = None
                    in_field = False
                    collecting_cached = False

            elif tag == f"{W}instrText" and in_field and not collecting_cached:
                if current is not None and elem.text:
                    current["code"] += elem.text

            elif tag == f"{W}t" and collecting_cached:
                if current is not None and elem.text:
                    current["cached_value"] += elem.text

        return results

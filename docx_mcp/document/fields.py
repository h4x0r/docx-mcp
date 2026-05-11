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
          - field_id: positional identifier (e.g. "field_0")
          - code: stripped instrText content
          - type: first word of code (field type keyword)
          - instruction: full instrText content (same as code)
          - result: text from runs between separate and end fldChar
          - cached_value: alias for result (backward compat)
          - para_id: w14:paraId of the enclosing <w:p>, or None

        Returns:
            List of field dicts.
        """
        doc = self._tree("word/document.xml")
        if doc is None:
            return []

        results: list[dict] = []
        field_index = 0

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

            code = "".join(code_parts).strip()
            field_type = code.split()[0] if code else ""
            cached = "".join(cached_parts)
            field_id = f"field_{field_index}"
            field_index += 1

            results.append({
                "field_id": field_id,
                "code": code,
                "type": field_type,
                "instruction": code,
                "result": cached,
                "cached_value": cached,
                "para_id": para_id,
            })

        return results

    def get_field(self, field_id: str) -> dict:
        """Return details of a single field by field_id.

        Args:
            field_id: Positional identifier returned by list_fields (e.g. "field_0").

        Returns:
            {"field_id": str, "type": str, "instruction": str, "result": str,
             "code": str, "cached_value": str, "para_id": str | None}

        Raises:
            ValueError if field_id not found.
        """
        fields = self.list_fields()
        for f in fields:
            if f["field_id"] == field_id:
                return f
        raise ValueError(f"Field '{field_id}' not found.")

    def delete_field(self, field_id: str) -> dict:
        """Remove a complete complex field (begin through end runs).

        Args:
            field_id: Positional identifier returned by list_fields (e.g. "field_0").

        Returns:
            {"field_id": str, "deleted": True}

        Raises:
            ValueError if field_id not found.
        """
        doc = self._require("word/document.xml")

        field_index = 0
        for fld_char in doc.iter(f"{W}fldChar"):
            if fld_char.get(f"{W}fldCharType") != "begin":
                continue

            current_id = f"field_{field_index}"
            field_index += 1

            if current_id != field_id:
                continue

            # Found — collect all runs from begin through end
            begin_run = fld_char.getparent()
            if begin_run is None:
                raise ValueError(f"Field '{field_id}' has no parent run.")
            para = begin_run.getparent()
            if para is None:
                raise ValueError(f"Field '{field_id}' run has no parent paragraph.")

            children = list(para)
            try:
                start_idx = children.index(begin_run)
            except ValueError:
                raise ValueError(f"Field '{field_id}' begin run not in paragraph.")

            runs_to_remove = [begin_run]
            for sibling in children[start_idx + 1:]:
                runs_to_remove.append(sibling)
                fc = sibling.find(f"{W}fldChar")
                if fc is not None and fc.get(f"{W}fldCharType") == "end":
                    break

            for run in runs_to_remove:
                para.remove(run)

            self._mark("word/document.xml")
            return {"field_id": field_id, "deleted": True}

        raise ValueError(f"Field '{field_id}' not found.")

    def _build_field_runs(self, instr: str) -> list:
        """Return [begin_run, instr_run, separate_run, end_run] elements."""
        runs = []
        for fchar_type, text in [
            ("begin", None),
            ("instrText", instr),
            ("separate", None),
            ("end", None),
        ]:
            r = etree.Element(f"{W}r")
            if fchar_type == "instrText":
                it = etree.SubElement(r, f"{W}instrText")
                it.set(XML_SPACE, "preserve")
                it.text = text
            else:
                fc = etree.SubElement(r, f"{W}fldChar")
                fc.set(f"{W}fldCharType", fchar_type)
                if fchar_type == "begin":
                    fc.set(f"{W}dirty", "true")
            runs.append(r)
        return runs

    def insert_date_field(
        self,
        para_id: str,
        date_format: str = r'\@ "MMMM d, yyyy"',
    ) -> dict:
        """Insert a DATE complex field at the end of a paragraph.

        Args:
            para_id: w14:paraId of the target paragraph.
            date_format: Date picture switch (default: \\@ "MMMM d, yyyy").

        Returns:
            {"para_id": str, "field_type": "DATE", "format": str}

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

        instr = f" DATE {date_format} "
        for run in self._build_field_runs(instr):
            para.append(run)

        self._mark("word/document.xml")
        return {"para_id": para_id, "field_type": "DATE", "format": date_format}

    def insert_page_number_field(self, para_id: str) -> dict:
        """Insert a PAGE complex field at the end of a paragraph.

        Args:
            para_id: w14:paraId of the target paragraph.

        Returns:
            {"para_id": str, "field_type": "PAGE"}

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

        instr = " PAGE "
        for run in self._build_field_runs(instr):
            para.append(run)

        self._mark("word/document.xml")
        return {"para_id": para_id, "field_type": "PAGE"}

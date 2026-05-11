"""Content Controls (SDT) mixin — checkbox, dropdown, date picker, plain text."""
from __future__ import annotations

from lxml import etree

from .base import W, W14, NSMAP, _preserve
from .errors import DocxMcpError, ErrCode

_TRUTHY = {"true", "1", "yes", "on"}


def _make_sdt(
    tag: str,
    control_type: str,
    label: str = "",
    options: list[str] | None = None,
    default: str = "",
) -> etree._Element:
    """Build and return a w:sdt element (not yet attached to a parent)."""
    sdt = etree.Element(f"{W}sdt")

    # ── sdtPr ──────────────────────────────────────────────────────────────
    sdtPr = etree.SubElement(sdt, f"{W}sdtPr")

    # w:tag
    tag_el = etree.SubElement(sdtPr, f"{W}tag")
    tag_el.set(f"{W}val", tag)

    # w:alias (label)
    if label:
        alias_el = etree.SubElement(sdtPr, f"{W}alias")
        alias_el.set(f"{W}val", label)

    if control_type == "checkbox":
        checkbox_el = etree.SubElement(
            sdtPr,
            f"{W14}checkbox",
            nsmap={"w14": NSMAP["w14"]},
        )
        checked_el = etree.SubElement(checkbox_el, f"{W14}checked")
        checked_el.set(f"{W14}val", "0")

    elif control_type == "dropdown":
        ddl = etree.SubElement(sdtPr, f"{W}dropDownList")
        for opt in (options or []):
            item = etree.SubElement(ddl, f"{W}listItem")
            item.set(f"{W}displayText", opt)
            item.set(f"{W}value", opt)

    elif control_type == "date":
        date_el = etree.SubElement(sdtPr, f"{W}date")
        if default:
            full_date = default if "T" in default else f"{default}T00:00:00Z"
            date_el.set(f"{W}fullDate", full_date)
        date_fmt = etree.SubElement(date_el, f"{W}dateFormat")
        date_fmt.set(f"{W}val", "MMMM d, yyyy")

    elif control_type == "text":
        etree.SubElement(sdtPr, f"{W}text")

    # ── sdtContent ─────────────────────────────────────────────────────────
    # (paragraph will be moved in here by the caller)
    etree.SubElement(sdt, f"{W}sdtContent")

    return sdt


def _find_sdt_by_id(doc: etree._Element, control_id: str) -> etree._Element | None:
    """Return the w:sdt whose sdtPr/w:id/@w:val equals control_id, or None."""
    for sdt in doc.iter(f"{W}sdt"):
        sdtPr = sdt.find(f"{W}sdtPr")
        if sdtPr is None:
            continue
        id_el = sdtPr.find(f"{W}id")
        if id_el is not None and id_el.get(f"{W}val") == control_id:
            return sdt
    return None


def _find_sdt_by_tag(doc: etree._Element, tag: str) -> etree._Element | None:
    """Return the w:sdt whose sdtPr/w:tag/@w:val equals tag, or None."""
    for sdt in doc.iter(f"{W}sdt"):
        sdtPr = sdt.find(f"{W}sdtPr")
        if sdtPr is None:
            continue
        tag_el = sdtPr.find(f"{W}tag")
        if tag_el is not None and tag_el.get(f"{W}val") == tag:
            return sdt
    return None


def _sdt_type(sdtPr: etree._Element) -> str:
    if sdtPr.find(f"{W14}checkbox") is not None:
        return "checkbox"
    if sdtPr.find(f"{W}dropDownList") is not None:
        return "dropdown"
    if sdtPr.find(f"{W}date") is not None:
        return "date"
    if sdtPr.find(f"{W}text") is not None:
        return "text"
    return "unknown"


def _sdt_value(sdt: etree._Element, ctrl_type: str) -> str:
    """Extract the current display value from an sdt."""
    sdtContent = sdt.find(f"{W}sdtContent")
    if sdtContent is None:
        return ""
    text = "".join(t.text for t in sdtContent.iter(f"{W}t") if t.text)
    if ctrl_type == "checkbox":
        sdtPr = sdt.find(f"{W}sdtPr")
        if sdtPr is not None:
            checked_el = sdtPr.find(f".//{W14}checked")
            if checked_el is not None:
                return "true" if checked_el.get(f"{W14}val") == "1" else "false"
    return text


class ContentControlsMixin:
    """CRUD operations for Word SDT content controls."""

    # ──────────────────────────────────────────────────────────────────────
    # Public API
    # ──────────────────────────────────────────────────────────────────────

    def add_content_control(
        self,
        para_id: str,
        tag: str,
        control_type: str,
        label: str = "",
        options: list[str] | None = None,
        default: str = "",
    ) -> dict:
        """Wrap the paragraph in an SDT content control.

        Returns {"tag": str, "type": str, "label": str}.
        Raises DocxMcpError(PARA_NOT_FOUND) if para_id not found.
        Raises DocxMcpError(OOXML_INVALID, "Duplicate tag") if tag already exists.
        """
        doc = self._require("word/document.xml")  # type: ignore[attr-defined]

        # Duplicate tag check
        if _find_sdt_by_tag(doc, tag) is not None:
            raise DocxMcpError(
                ErrCode.OOXML_INVALID,
                f"Duplicate tag: '{tag}' already exists in this document.",
            )

        # Find the paragraph
        para = self._find_para(doc, para_id)  # type: ignore[attr-defined]
        if para is None:
            raise DocxMcpError(
                ErrCode.PARA_NOT_FOUND,
                f"Paragraph '{para_id}' not found.",
            )

        # Build the SDT
        sdt = _make_sdt(tag, control_type, label=label, options=options, default=default)

        # Move para into sdtContent
        sdtContent = sdt.find(f"{W}sdtContent")
        parent = para.getparent()
        idx = list(parent).index(para)
        parent.remove(para)
        sdtContent.append(para)
        parent.insert(idx, sdt)

        # Set default text / checked state in sdtContent
        if control_type == "checkbox":
            # Ensure there's a w:t with ☐
            self._sdt_set_text(sdtContent, "☐")
        elif control_type in ("text", "dropdown", "date"):
            if default:
                self._sdt_set_text(sdtContent, default)
            elif control_type == "dropdown" and options:
                self._sdt_set_text(sdtContent, options[0])

        self._mark("word/document.xml")  # type: ignore[attr-defined]

        return {"tag": tag, "type": control_type, "label": label}

    def get_content_controls(self) -> list[dict]:
        """Return all SDT controls: [{"tag", "type", "label", "value"}]."""
        doc = self._require("word/document.xml")  # type: ignore[attr-defined]
        result: list[dict] = []
        for sdt in doc.iter(f"{W}sdt"):
            sdtPr = sdt.find(f"{W}sdtPr")
            if sdtPr is None:
                continue
            tag_el = sdtPr.find(f"{W}tag")
            if tag_el is None:
                continue
            tag = tag_el.get(f"{W}val", "")
            alias_el = sdtPr.find(f"{W}alias")
            label = alias_el.get(f"{W}val", "") if alias_el is not None else ""
            ctrl_type = _sdt_type(sdtPr)
            value = _sdt_value(sdt, ctrl_type)
            result.append({"tag": tag, "type": ctrl_type, "label": label, "value": value})
        return result

    def set_content_control_value(self, tag: str, value: str) -> dict:
        """Update the display text / checked state of an existing control.

        For checkbox: "true"/"1" → checked (☑), else → unchecked (☐).
        For others: sets the w:t text in sdtContent.
        Returns {"tag": str, "value": str}.
        Raises DocxMcpError(BOOKMARK_NOT_FOUND) if tag missing.
        """
        doc = self._require("word/document.xml")  # type: ignore[attr-defined]
        sdt = _find_sdt_by_tag(doc, tag)
        if sdt is None:
            raise DocxMcpError(
                ErrCode.BOOKMARK_NOT_FOUND,
                f"Tag not found: {tag}",
            )

        sdtPr = sdt.find(f"{W}sdtPr")
        ctrl_type = _sdt_type(sdtPr) if sdtPr is not None else "unknown"
        sdtContent = sdt.find(f"{W}sdtContent")
        if sdtContent is None:
            sdtContent = etree.SubElement(sdt, f"{W}sdtContent")

        if ctrl_type == "checkbox":
            checked = value.lower() in _TRUTHY
            # Update w14:checked/@w14:val
            if sdtPr is not None:
                checked_el = sdtPr.find(f".//{W14}checked")
                if checked_el is not None:
                    checked_el.set(f"{W14}val", "1" if checked else "0")
            display = "☑" if checked else "☐"
            self._sdt_set_text(sdtContent, display)
            self._mark("word/document.xml")  # type: ignore[attr-defined]
            return {"tag": tag, "value": "true" if checked else "false"}
        else:
            self._sdt_set_text(sdtContent, value)

        self._mark("word/document.xml")  # type: ignore[attr-defined]
        return {"tag": tag, "value": value}

    def lock_content_control(self, tag: str, lock: str = "sdtLocked") -> dict:
        """Add w:lock to sdtPr.

        lock values: sdtLocked | contentLocked | sdtContentLocked.
        Returns {"tag": str, "lock": str}.
        Raises DocxMcpError(BOOKMARK_NOT_FOUND) if tag missing.
        """
        doc = self._require("word/document.xml")  # type: ignore[attr-defined]
        sdt = _find_sdt_by_tag(doc, tag)
        if sdt is None:
            raise DocxMcpError(
                ErrCode.BOOKMARK_NOT_FOUND,
                f"Tag not found: {tag}",
            )

        sdtPr = sdt.find(f"{W}sdtPr")
        if sdtPr is None:
            sdtPr = etree.Element(f"{W}sdtPr")
            sdt.insert(0, sdtPr)

        # Add or replace w:lock
        lock_el = sdtPr.find(f"{W}lock")
        if lock_el is None:
            lock_el = etree.SubElement(sdtPr, f"{W}lock")
        lock_el.set(f"{W}val", lock)

        self._mark("word/document.xml")  # type: ignore[attr-defined]
        return {"tag": tag, "lock": lock}

    def delete_content_control(self, control_id: str) -> dict:
        """Remove the SDT wrapper but keep its content in place (unwrap).

        Returns {"control_id": str, "deleted": True}.
        Raises ValueError if control_id not found.
        """
        doc = self._require("word/document.xml")  # type: ignore[attr-defined]
        sdt = _find_sdt_by_id(doc, control_id)
        if sdt is None:
            raise ValueError(f"Content control '{control_id}' not found.")

        sdtContent = sdt.find(f"{W}sdtContent")
        parent = sdt.getparent()
        idx = list(parent).index(sdt)

        if sdtContent is not None:
            children = list(sdtContent)
            for i, child in enumerate(children):
                parent.insert(idx + i, child)

        parent.remove(sdt)
        self._mark("word/document.xml")  # type: ignore[attr-defined]
        return {"control_id": control_id, "deleted": True}

    def get_content_control(self, control_id: str) -> dict:
        """Return details of a single content control by its w:id.

        Returns {"control_id", "type", "title", "tag", "value"}.
        Raises ValueError if control_id not found.
        """
        doc = self._require("word/document.xml")  # type: ignore[attr-defined]
        sdt = _find_sdt_by_id(doc, control_id)
        if sdt is None:
            raise ValueError(f"Content control '{control_id}' not found.")

        sdtPr = sdt.find(f"{W}sdtPr")
        tag_el = sdtPr.find(f"{W}tag") if sdtPr is not None else None
        tag = tag_el.get(f"{W}val", "") if tag_el is not None else ""
        alias_el = sdtPr.find(f"{W}alias") if sdtPr is not None else None
        title = alias_el.get(f"{W}val", "") if alias_el is not None else ""
        ctrl_type = _sdt_type(sdtPr) if sdtPr is not None else "unknown"
        value = _sdt_value(sdt, ctrl_type)

        return {
            "control_id": control_id,
            "type": ctrl_type,
            "title": title,
            "tag": tag,
            "value": value,
        }

    def update_content_control(
        self,
        control_id: str,
        *,
        title: str | None = None,
        tag: str | None = None,
        placeholder_text: str | None = None,
    ) -> dict:
        """Modify properties of an existing content control.

        Returns {"control_id", "title", "tag", "placeholder_text"}.
        Raises ValueError if control_id not found.
        """
        doc = self._require("word/document.xml")  # type: ignore[attr-defined]
        sdt = _find_sdt_by_id(doc, control_id)
        if sdt is None:
            raise ValueError(f"Content control '{control_id}' not found.")

        sdtPr = sdt.find(f"{W}sdtPr")
        if sdtPr is None:
            sdtPr = etree.Element(f"{W}sdtPr")
            sdt.insert(0, sdtPr)

        if title is not None:
            alias_el = sdtPr.find(f"{W}alias")
            if alias_el is None:
                alias_el = etree.SubElement(sdtPr, f"{W}alias")
            alias_el.set(f"{W}val", title)

        if tag is not None:
            tag_el = sdtPr.find(f"{W}tag")
            if tag_el is None:
                tag_el = etree.SubElement(sdtPr, f"{W}tag")
            tag_el.set(f"{W}val", tag)

        if placeholder_text is not None:
            ph_el = sdtPr.find(f"{W}placeholder")
            if ph_el is None:
                ph_el = etree.SubElement(sdtPr, f"{W}placeholder")
            r_el = ph_el.find(f"{W}r")
            if r_el is None:
                r_el = etree.SubElement(ph_el, f"{W}r")
            t_el = r_el.find(f"{W}t")
            if t_el is None:
                t_el = etree.SubElement(r_el, f"{W}t")
            _preserve(t_el, placeholder_text)

        self._mark("word/document.xml")  # type: ignore[attr-defined]
        return {
            "control_id": control_id,
            "title": title,
            "tag": tag,
            "placeholder_text": placeholder_text,
        }

    # ──────────────────────────────────────────────────────────────────────
    # Private helpers
    # ──────────────────────────────────────────────────────────────────────

    @staticmethod
    def _sdt_set_text(sdtContent: etree._Element, text: str) -> None:
        """Find or create the first w:p/w:r/w:t in sdtContent and set its text."""
        # Look for existing w:t
        for t in sdtContent.iter(f"{W}t"):
            _preserve(t, text)
            return

        # No w:t found — ensure w:p > w:r > w:t exist
        para = sdtContent.find(f"{W}p")
        if para is None:
            para = etree.SubElement(sdtContent, f"{W}p")
        run = para.find(f"{W}r")
        if run is None:
            run = etree.SubElement(para, f"{W}r")
        t = run.find(f"{W}t")
        if t is None:
            t = etree.SubElement(run, f"{W}t")
        _preserve(t, text)

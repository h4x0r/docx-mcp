"""Equation support: LaTeX → MathML → OMML (Office Math Markup Language)."""
from __future__ import annotations

from lxml import etree

from .base import W, W14
from .errors import DocxMcpError, ErrCode

M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


# ── MathML → OMML converter ─────────────────────────────────────────────────


def _extract_text(node: etree._Element) -> str:
    """Recursively extract all text from a MathML element."""
    parts = []
    if node.text:
        parts.append(node.text)
    for child in node:
        if not callable(child.tag):
            parts.append(_extract_text(child))
    return "".join(parts)


def _convert_node(node: etree._Element) -> etree._Element:
    """Convert a single MathML element to its OMML equivalent."""
    local = etree.QName(node).localname

    if local in ("mi", "mn", "mo"):
        r = etree.Element(f"{M}r")
        t = etree.SubElement(r, f"{M}t")
        t.text = node.text or ""
        return r

    elif local == "mtext":
        r = etree.Element(f"{M}r")
        rPr = etree.SubElement(r, f"{M}rPr")
        etree.SubElement(rPr, f"{M}nor")
        t = etree.SubElement(r, f"{M}t")
        t.text = node.text or ""
        return r

    elif local == "mrow":
        # mrow has no direct OMML equivalent — flatten to run with concatenated text
        group = etree.Element(f"{M}r")
        t = etree.SubElement(group, f"{M}t")
        t.text = _extract_text(node)
        return group

    elif local == "mfrac":
        children = [c for c in node if not callable(c.tag)]
        f = etree.Element(f"{M}f")
        num = etree.SubElement(f, f"{M}num")
        den = etree.SubElement(f, f"{M}den")
        if len(children) >= 1:
            num.append(_convert_node(children[0]))
        if len(children) >= 2:
            den.append(_convert_node(children[1]))
        return f

    elif local == "msup":
        children = [c for c in node if not callable(c.tag)]
        ssup = etree.Element(f"{M}sSup")
        e_el = etree.SubElement(ssup, f"{M}e")
        sup_el = etree.SubElement(ssup, f"{M}sup")
        if len(children) >= 1:
            e_el.append(_convert_node(children[0]))
        if len(children) >= 2:
            sup_el.append(_convert_node(children[1]))
        return ssup

    elif local == "msub":
        children = [c for c in node if not callable(c.tag)]
        ssub = etree.Element(f"{M}sSub")
        e_el = etree.SubElement(ssub, f"{M}e")
        sub_el = etree.SubElement(ssub, f"{M}sub")
        if len(children) >= 1:
            e_el.append(_convert_node(children[0]))
        if len(children) >= 2:
            sub_el.append(_convert_node(children[1]))
        return ssub

    elif local == "msqrt":
        rad = etree.Element(f"{M}rad")
        radPr = etree.SubElement(rad, f"{M}radPr")
        degHide = etree.SubElement(radPr, f"{M}degHide")
        degHide.set(f"{M}val", "1")
        etree.SubElement(rad, f"{M}deg")
        e_el = etree.SubElement(rad, f"{M}e")
        _convert_children(node, e_el)
        return rad

    else:
        # Fallback: extract all text into a run
        r = etree.Element(f"{M}r")
        t = etree.SubElement(r, f"{M}t")
        t.text = _extract_text(node)
        return r


def _convert_children(src: etree._Element, dst: etree._Element) -> None:
    """Recursively convert MathML children and append OMML to dst."""
    for child in src:
        if callable(child.tag):  # skip comments/PIs
            continue
        el = _convert_node(child)
        dst.append(el)


def _mathml_to_omml(node: etree._Element) -> etree._Element:
    """Convert MathML root (math element) to m:oMath."""
    omath = etree.Element(f"{M}oMath")
    _convert_children(node, omath)
    return omath


# ── Mixin ────────────────────────────────────────────────────────────────────


class EquationsMixin:
    """Add/retrieve OMML (Office Math) equations in a docx."""

    def add_equation(self, para_id: str, latex: str) -> dict:
        """Insert a LaTeX equation as OMML (Office Math) after a paragraph.

        Pipeline: LaTeX → MathML (via latex2mathml) → OMML (custom converter).

        Raises:
            DocxMcpError(PII_DEPS_MISSING): if latex2mathml is not installed.
            DocxMcpError(OOXML_INVALID): if conversion fails.

        Returns:
            {"para_id": str, "latex": str, "omml_tag": str}
        """
        try:
            from latex2mathml.converter import convert
        except ImportError:
            raise DocxMcpError(
                ErrCode.PII_DEPS_MISSING,
                "latex2mathml required: pip install latex2mathml",
            )
        try:
            mathml_str = convert(latex)
        except Exception as e:
            raise DocxMcpError(ErrCode.OOXML_INVALID, f"LaTeX conversion failed: {e}")

        mathml_tree = etree.fromstring(mathml_str.encode())
        omml = _mathml_to_omml(mathml_tree)

        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise DocxMcpError(ErrCode.PARA_NOT_FOUND, f"Paragraph {para_id!r} not found")

        omath_para = etree.Element(f"{M}oMathPara")
        omath_para.append(omml)
        para.addnext(omath_para)
        self._mark("word/document.xml")

        return {"para_id": para_id, "latex": latex, "omml_tag": f"{M}oMathPara"}

    def get_equations(self) -> list[dict]:
        """Return all m:oMath elements found in the document.

        Returns:
            [{"omml_xml": str, "para_id": str | None}]
        """
        doc = self._tree("word/document.xml")
        if doc is None:
            return []
        results = []
        for omath in doc.iter(f"{M}oMath"):
            # Find the containing m:oMathPara first
            omath_para = omath.getparent()
            while omath_para is not None and omath_para.tag != f"{M}oMathPara":
                omath_para = omath_para.getparent()

            para_id = None
            if omath_para is not None:
                prev = omath_para.getprevious()
                if prev is not None and prev.tag == f"{W}p":
                    para_id = prev.get(f"{W14}paraId")
            results.append({
                "omml_xml": etree.tostring(omath, encoding="unicode"),
                "para_id": para_id,
            })
        return results

"""XPath query escape hatch — run raw XPath against any DOCX part."""
from __future__ import annotations

from lxml import etree

from .errors import DocxMcpError, ErrCode

# Pre-bound namespace map for all queries
_NS: dict[str, str] = {
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

_MAX_RESULTS = 50


class XPathMixin:

    def xpath_query(
        self,
        xpath: str,
        part: str = "word/document.xml",
    ) -> dict:
        """Run XPath against any DOCX part. Returns up to 50 matching elements.

        Pre-bound namespace prefixes: w, w14, r, wp, a, mc.
        Examples:
          "//w:p"  — all paragraphs
          "//w:p[w:pPr/w:pStyle/@w:val='Heading1']"  — heading 1 paragraphs
          "//w:t/text()"  — all text content

        Args:
            xpath: XPath 1.0 expression
            part:  Part path to query (default: word/document.xml)

        Returns:
            {
              "xpath": str,
              "part": str,
              "count": int,        # total matches (may exceed 50)
              "returned": int,     # number of results in "results" list
              "results": list[str] # XML snippets or string values, capped at 50
            }

        Raises:
            DocxMcpError(ErrCode.PART_NOT_FOUND) if part not in document
            DocxMcpError(ErrCode.XPATH_ERROR) if xpath is invalid
        """
        tree = self._tree(part)
        if tree is None:
            raise DocxMcpError(
                ErrCode.PART_NOT_FOUND,
                f"Part not found: {part}",
                hint="Use list_parts() to see available parts.",
            )

        try:
            matches = tree.xpath(xpath, namespaces=_NS)
        except etree.XPathError as exc:
            raise DocxMcpError(
                ErrCode.XPATH_ERROR,
                f"XPath error: {exc}",
                hint="Check namespace prefixes: w, w14, r, wp, a, mc",
            ) from exc

        total = len(matches) if isinstance(matches, list) else 1
        if not isinstance(matches, list):
            matches = [matches]

        results: list[str] = []
        for m in matches[:_MAX_RESULTS]:
            if isinstance(m, etree._Element):
                results.append(
                    etree.tostring(m, pretty_print=True).decode()
                )
            else:
                results.append(str(m))

        return {
            "xpath": xpath,
            "part": part,
            "count": total,
            "returned": len(results),
            "results": results,
        }

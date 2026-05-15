"""Charts mixin: native DrawingML chart insertion (no Excel required)."""

from __future__ import annotations

from lxml import etree

from .base import CT, RELS, W14, WP, A, R, W
from .errors import DocxMcpError, ErrCode

# ── Chart-specific namespace constants ──────────────────────────────────────
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
_C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"  # bare URI
_A_NS = A[1:-1]  # bare URI
_WP_NS = WP[1:-1]  # bare URI
_R_NS = R[1:-1]  # bare URI

_CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
_CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"


def _col_letter(n: int) -> str:
    """Convert 1-based column index to Excel-style letter (1=A, 26=Z, 27=AA)."""
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def _build_chart_xml(
    title: str,
    series: list[dict],
    categories: list[str],
    chart_type: str,
) -> etree._Element:
    """Build a minimal chartSpace XML element."""
    cs = etree.Element(
        f"{C}chartSpace",
        nsmap={
            "c": _C_NS,
            "a": _A_NS,
            "r": _R_NS,
        },
    )
    chart = etree.SubElement(cs, f"{C}chart")

    # Title
    title_el = etree.SubElement(chart, f"{C}title")
    tx = etree.SubElement(title_el, f"{C}tx")
    rich = etree.SubElement(tx, f"{C}rich")
    etree.SubElement(rich, f"{A}bodyPr")
    etree.SubElement(rich, f"{A}lstStyle")
    p = etree.SubElement(rich, f"{A}p")
    r = etree.SubElement(p, f"{A}r")
    t = etree.SubElement(r, f"{A}t")
    t.text = title
    etree.SubElement(title_el, f"{C}overlay", {"val": "0"})

    etree.SubElement(chart, f"{C}autoTitleDeleted", {"val": "0"})

    plot_area = etree.SubElement(chart, f"{C}plotArea")

    # Chart-type element
    if chart_type == "bar":
        chart_el = etree.SubElement(plot_area, f"{C}barChart")
        etree.SubElement(chart_el, f"{C}barDir", {"val": "col"})
        etree.SubElement(chart_el, f"{C}grouping", {"val": "clustered"})
    elif chart_type == "line":
        chart_el = etree.SubElement(plot_area, f"{C}lineChart")
        etree.SubElement(chart_el, f"{C}grouping", {"val": "standard"})
    elif chart_type == "pie":
        chart_el = etree.SubElement(plot_area, f"{C}pieChart")
        etree.SubElement(chart_el, f"{C}firstSliceAng", {"val": "0"})
    else:
        raise ValueError(f"Unknown chart_type: {chart_type!r}")

    # Add series
    for idx, ser_data in enumerate(series):
        ser = etree.SubElement(chart_el, f"{C}ser")
        etree.SubElement(ser, f"{C}idx", {"val": str(idx)})
        etree.SubElement(ser, f"{C}order", {"val": str(idx)})

        # Series name
        tx2 = etree.SubElement(ser, f"{C}tx")
        str_ref = etree.SubElement(tx2, f"{C}strRef")
        etree.SubElement(str_ref, f"{C}f").text = f"Sheet1!${_col_letter(idx + 2)}$1"
        str_cache = etree.SubElement(str_ref, f"{C}strCache")
        etree.SubElement(str_cache, f"{C}ptCount", {"val": "1"})
        pt = etree.SubElement(str_cache, f"{C}pt", {"idx": "0"})
        etree.SubElement(pt, f"{C}v").text = ser_data.get("name", f"Series {idx + 1}")

        # Categories
        if categories:
            cat = etree.SubElement(ser, f"{C}cat")
            cat_ref = etree.SubElement(cat, f"{C}strRef")
            etree.SubElement(cat_ref, f"{C}f").text = f"Sheet1!$A$2:$A${len(categories) + 1}"
            cat_cache = etree.SubElement(cat_ref, f"{C}strCache")
            etree.SubElement(cat_cache, f"{C}ptCount", {"val": str(len(categories))})
            for ci, cat_name in enumerate(categories):
                cpt = etree.SubElement(cat_cache, f"{C}pt", {"idx": str(ci)})
                etree.SubElement(cpt, f"{C}v").text = str(cat_name)

        # Values
        val_el = etree.SubElement(ser, f"{C}val")
        num_ref = etree.SubElement(val_el, f"{C}numRef")
        values = ser_data.get("values", [])
        etree.SubElement(
            num_ref, f"{C}f"
        ).text = f"Sheet1!${_col_letter(idx + 2)}$2:${_col_letter(idx + 2)}${len(values) + 1}"
        num_cache = etree.SubElement(num_ref, f"{C}numCache")
        etree.SubElement(num_cache, f"{C}formatCode").text = "General"
        etree.SubElement(num_cache, f"{C}ptCount", {"val": str(len(values))})
        for vi, v in enumerate(values):
            vpt = etree.SubElement(num_cache, f"{C}pt", {"idx": str(vi)})
            etree.SubElement(vpt, f"{C}v").text = str(v)

    # Legend
    legend = etree.SubElement(chart, f"{C}legend")
    etree.SubElement(legend, f"{C}legendPos", {"val": "b"})

    return cs


def _rebuild_series(chart_el: etree._Element, series: list[dict]) -> None:
    """Remove existing c:ser elements and rebuild from series list."""
    for ser in list(chart_el.findall(f"{C}ser")):
        chart_el.remove(ser)

    for idx, ser_data in enumerate(series):
        ser = etree.SubElement(chart_el, f"{C}ser")
        etree.SubElement(ser, f"{C}idx", {"val": str(idx)})
        etree.SubElement(ser, f"{C}order", {"val": str(idx)})

        # Series name
        tx2 = etree.SubElement(ser, f"{C}tx")
        str_ref = etree.SubElement(tx2, f"{C}strRef")
        etree.SubElement(str_ref, f"{C}f").text = f"Sheet1!${_col_letter(idx + 2)}$1"
        str_cache = etree.SubElement(str_ref, f"{C}strCache")
        etree.SubElement(str_cache, f"{C}ptCount", {"val": "1"})
        pt = etree.SubElement(str_cache, f"{C}pt", {"idx": "0"})
        etree.SubElement(pt, f"{C}v").text = ser_data.get("name", f"Series {idx + 1}")

        # Values
        val_el = etree.SubElement(ser, f"{C}val")
        num_ref = etree.SubElement(val_el, f"{C}numRef")
        values = ser_data.get("values", [])
        etree.SubElement(
            num_ref, f"{C}f"
        ).text = f"Sheet1!${_col_letter(idx + 2)}$2:${_col_letter(idx + 2)}${len(values) + 1}"
        num_cache = etree.SubElement(num_ref, f"{C}numCache")
        etree.SubElement(num_cache, f"{C}formatCode").text = "General"
        etree.SubElement(num_cache, f"{C}ptCount", {"val": str(len(values))})
        for vi, v in enumerate(values):
            vpt = etree.SubElement(num_cache, f"{C}pt", {"idx": str(vi)})
            etree.SubElement(vpt, f"{C}v").text = str(v)


class ChartsMixin:
    """Native DrawingML chart insertion — no Excel required."""

    # ── Public API ───────────────────────────────────────────────────────────

    def insert_bar_chart(
        self,
        para_id: str,
        title: str,
        series: list[dict],
        categories: list[str],
        width_cm: float = 14.0,
        height_cm: float = 9.0,
    ) -> dict:
        """Insert a native bar chart (no Excel required).

        Returns: {"chart_id": str, "rId": str, "part_path": str}
        """
        return self._insert_chart(para_id, title, series, categories, width_cm, height_cm, "bar")

    def insert_line_chart(
        self,
        para_id: str,
        title: str,
        series: list[dict],
        categories: list[str],
        width_cm: float = 14.0,
        height_cm: float = 9.0,
    ) -> dict:
        """Insert a native line chart.

        Returns: {"chart_id": str, "rId": str, "part_path": str}
        """
        return self._insert_chart(para_id, title, series, categories, width_cm, height_cm, "line")

    def insert_pie_chart(
        self,
        para_id: str,
        title: str,
        series: list[dict],
        categories: list[str],
    ) -> dict:
        """Insert a native pie chart (single series, fixed 14×9 cm).

        Returns: {"chart_id": str, "rId": str, "part_path": str}
        """
        return self._insert_chart(para_id, title, series, categories, 14.0, 9.0, "pie")

    def update_chart_data(self, chart_id: str, series: list[dict]) -> dict:
        """Replace data series in an existing chart by chart_id.

        Returns: {"chart_id": str, "updated": True}
        Raises DocxMcpError(PART_NOT_FOUND) if chart not found.
        """
        chart_path = f"word/charts/{chart_id}.xml"

        # Try from cache first, then from disk
        chart_tree = self._tree(chart_path)
        if chart_tree is None:
            fp = self.workdir / chart_path
            if not fp.exists():
                raise DocxMcpError(
                    ErrCode.PART_NOT_FOUND,
                    f"Chart {chart_id!r} not found",
                )
            from lxml import etree as _etree

            parser = _etree.XMLParser(remove_blank_text=False)
            chart_tree = _etree.parse(str(fp), parser).getroot()
            self._trees[chart_path] = chart_tree

        # Find the chart-type element and rebuild series
        chart_el = None
        for chart_el in chart_tree.iter(f"{C}barChart", f"{C}lineChart", f"{C}pieChart"):
            _rebuild_series(chart_el, series)
            break
        if chart_el is None:
            raise DocxMcpError(
                ErrCode.OOXML_INVALID,
                f"No recognised chart type in {chart_id!r}",
            )

        # Write back to disk immediately (tests read disk directly)
        fp = self.workdir / chart_path
        fp.parent.mkdir(parents=True, exist_ok=True)
        etree.ElementTree(chart_tree).write(
            str(fp),
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        self._mark(chart_path)
        return {"chart_id": chart_id, "updated": True}

    # ── Shared insertion helper ──────────────────────────────────────────────

    def _insert_chart(
        self,
        para_id: str,
        title: str,
        series: list[dict],
        categories: list[str],
        width_cm: float,
        height_cm: float,
        chart_type: str,
    ) -> dict:
        # Step 1: Allocate chart number
        chart_dir = self.workdir / "word" / "charts"
        chart_dir.mkdir(parents=True, exist_ok=True)
        existing = list(chart_dir.glob("chart*.xml"))
        chart_num = len(existing) + 1
        chart_path = f"word/charts/chart{chart_num}.xml"
        chart_id = f"chart{chart_num}"

        # Step 2: Build chart XML
        chart_xml_el = _build_chart_xml(title, series, categories, chart_type)

        # Step 3: Write chart part to disk and cache
        chart_xml_bytes = etree.tostring(
            chart_xml_el,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        (self.workdir / chart_path).write_bytes(chart_xml_bytes)
        self._trees[chart_path] = chart_xml_el
        self._mark(chart_path)

        # Step 4: Add relationship to document.xml.rels
        rels = self._require("word/_rels/document.xml.rels")
        existing_rids = [r.get("Id") for r in rels.findall(f"{RELS}Relationship")]
        rid_num = (
            max(
                (int(r.replace("rId", "")) for r in existing_rids if r and r.startswith("rId")),
                default=0,
            )
            + 1
        )
        rid = f"rId{rid_num}"
        rel = etree.SubElement(rels, f"{RELS}Relationship")
        rel.set("Id", rid)
        rel.set("Type", _CHART_REL_TYPE)
        rel.set("Target", f"charts/chart{chart_num}.xml")
        self._mark("word/_rels/document.xml.rels")

        # Step 5: Update [Content_Types].xml
        ct_tree = self._require("[Content_Types].xml")
        chart_part_name = f"/word/charts/chart{chart_num}.xml"
        existing_overrides = ct_tree.findall(f"{CT}Override[@PartName='{chart_part_name}']")
        if not existing_overrides:
            override = etree.SubElement(ct_tree, f"{CT}Override")
            override.set("PartName", chart_part_name)
            override.set("ContentType", _CHART_CT)
            self._mark("[Content_Types].xml")

        # Step 6: Insert drawing paragraph into document.xml
        width_emu = int(width_cm * 914400 / 2.54)
        height_emu = int(height_cm * 914400 / 2.54)

        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise DocxMcpError(
                ErrCode.PARA_NOT_FOUND,
                f"Paragraph {para_id!r} not found",
            )

        # docPr id must be unique across ALL drawing objects (images + charts)
        existing_doc_pr_ids = [
            int(dp.get("id", "0")) for dp in doc.iter(f"{WP}docPr") if dp.get("id", "").isdigit()
        ]
        doc_pr_id = max(existing_doc_pr_ids, default=0) + 1

        new_para = etree.Element(f"{W}p")
        new_para.set(f"{W14}paraId", self._new_para_id())
        run = etree.SubElement(new_para, f"{W}r")
        drawing = etree.SubElement(run, f"{W}drawing")
        inline = etree.SubElement(
            drawing,
            f"{WP}inline",
            distT="0",
            distB="0",
            distL="0",
            distR="0",
        )
        extent = etree.SubElement(inline, f"{WP}extent")
        extent.set("cx", str(width_emu))
        extent.set("cy", str(height_emu))

        doc_pr = etree.SubElement(inline, f"{WP}docPr")
        doc_pr.set("id", str(doc_pr_id))
        doc_pr.set("name", f"Chart {chart_num}")

        graphic = etree.SubElement(inline, f"{A}graphic")
        graphic_data = etree.SubElement(graphic, f"{A}graphicData")
        graphic_data.set("uri", _C_NS)

        chart_ref = etree.SubElement(graphic_data, f"{{{_C_NS}}}chart")
        chart_ref.set(f"{R}id", rid)

        para.addnext(new_para)
        self._mark("word/document.xml")

        return {
            "chart_id": chart_id,
            "rId": rid,
            "part_path": chart_path,
        }

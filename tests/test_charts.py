"""Tests for native chart insertion (Task #17)."""
from __future__ import annotations

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument, W, W14, WP
from docx_mcp.document.base import RELS

C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"


def _make_doc(tmp_path):
    return DocxDocument.create(str(tmp_path / "test.docx"))


def _get_para_id(doc):
    tree = doc._tree("word/document.xml")
    return tree.findall(f".//{W}p")[0].get(f"{W14}paraId")


_SERIES = [{"name": "Q1", "values": [10, 20, 30]}]
_CATS = ["Jan", "Feb", "Mar"]


class TestCharts:
    def test_insert_bar_chart_creates_part(self, tmp_path):
        """insert_bar_chart creates word/charts/chart1.xml."""
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        result = doc.insert_bar_chart(para_id, "My Bar Chart", _SERIES, _CATS)
        assert "chart_id" in result
        assert (doc.workdir / "word" / "charts" / "chart1.xml").exists()

    def test_chart_relationship_in_document(self, tmp_path):
        """A chart relationship is added to document.xml.rels."""
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        result = doc.insert_bar_chart(para_id, "Chart", _SERIES, _CATS)
        rels = doc._tree("word/_rels/document.xml.rels")
        chart_rels = [
            r for r in rels.iter(f"{RELS}Relationship")
            if "chart" in r.get("Type", "").lower()
        ]
        assert len(chart_rels) >= 1

    def test_series_data_in_chart_xml(self, tmp_path):
        """Chart XML contains the series values."""
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        doc.insert_bar_chart(para_id, "Chart", _SERIES, _CATS)
        chart_xml = (doc.workdir / "word" / "charts" / "chart1.xml").read_bytes()
        tree = etree.fromstring(chart_xml)
        vals = [el.text for el in tree.iter(f"{C}v")]
        assert "10" in vals
        assert "20" in vals
        assert "30" in vals

    def test_insert_line_chart(self, tmp_path):
        """insert_line_chart creates a lineChart element in the chart XML."""
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        result = doc.insert_line_chart(para_id, "Line", _SERIES, _CATS)
        assert "chart_id" in result
        chart_xml = (doc.workdir / "word" / "charts" / "chart1.xml").read_bytes()
        tree = etree.fromstring(chart_xml)
        assert len(list(tree.iter(f"{C}lineChart"))) >= 1

    def test_insert_pie_chart(self, tmp_path):
        """insert_pie_chart creates a pieChart element in the chart XML."""
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        pie_series = [{"name": "Data", "values": [30, 40, 30]}]
        result = doc.insert_pie_chart(para_id, "Pie", pie_series, _CATS)
        assert "chart_id" in result
        chart_xml = (doc.workdir / "word" / "charts" / "chart1.xml").read_bytes()
        tree = etree.fromstring(chart_xml)
        assert len(list(tree.iter(f"{C}pieChart"))) >= 1

    def test_update_chart_data(self, tmp_path):
        """update_chart_data replaces series values in the chart."""
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        result = doc.insert_bar_chart(para_id, "Chart", _SERIES, _CATS)
        chart_id = result["chart_id"]
        new_series = [{"name": "Q2", "values": [99, 88, 77]}]
        update = doc.update_chart_data(chart_id, new_series)
        assert update["updated"] is True
        chart_xml = (doc.workdir / "word" / "charts" / f"{chart_id}.xml").read_bytes()
        tree = etree.fromstring(chart_xml)
        vals = [el.text for el in tree.iter(f"{C}v")]
        assert "99" in vals

    def test_many_series_col_letters_valid(self, tmp_path):
        """Column letters must stay valid past 24 series (no chr overflow)."""
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        many = [{"name": f"S{i}", "values": [i]} for i in range(26)]
        result = doc.insert_bar_chart(para_id, "Big", many, ["X"])
        chart_xml = (doc.workdir / "word" / "charts" / "chart1.xml").read_text()
        assert "[" not in chart_xml  # chr(91) sentinel for overflow
        assert "\\" not in chart_xml

    def test_update_chart_data_unknown_type_raises(self, tmp_path):
        """update_chart_data raises OOXML_INVALID if chart type not recognised."""
        from lxml import etree as _et
        from docx_mcp.document.errors import DocxMcpError, ErrCode
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        result = doc.insert_bar_chart(para_id, "Chart", _SERIES, _CATS)
        chart_id = result["chart_id"]
        # Replace barChart with areaChart to simulate unknown type
        chart_path = doc.workdir / "word" / "charts" / f"{chart_id}.xml"
        content = chart_path.read_text()
        chart_path.write_text(content.replace("barChart", "areaChart"))
        doc._trees.pop(f"word/charts/{chart_id}.xml", None)
        with pytest.raises(DocxMcpError) as exc:
            doc.update_chart_data(chart_id, _SERIES)
        assert exc.value.code == ErrCode.OOXML_INVALID

    def test_doc_pr_id_unique_across_charts(self, tmp_path):
        """Each chart drawing gets a unique docPr/@id (no collision)."""
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        doc.insert_bar_chart(para_id, "C1", _SERIES, _CATS)
        doc.insert_line_chart(para_id, "C2", _SERIES, _CATS)
        tree = doc._tree("word/document.xml")
        ids = [dp.get("id") for dp in tree.iter(f"{WP}docPr")]
        assert len(ids) == len(set(ids)), f"duplicate docPr ids: {ids}"

    def test_chart_attributes_are_unqualified(self, tmp_path):
        """Chart element attributes must be unqualified (no c: namespace prefix)."""
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        doc.insert_bar_chart(para_id, "Chart", _SERIES, _CATS)
        chart_xml = (doc.workdir / "word" / "charts" / "chart1.xml").read_text(encoding="utf-8")
        # Attributes like val and idx on chart elements must NOT have the c: prefix
        assert 'c:val=' not in chart_xml
        assert 'c:idx=' not in chart_xml
        # But the actual values should be present as plain attributes
        assert 'val=' in chart_xml
        assert 'idx=' in chart_xml

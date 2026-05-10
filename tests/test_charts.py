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

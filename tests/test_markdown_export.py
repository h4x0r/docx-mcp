"""Tests for export_markdown."""

from __future__ import annotations

import zipfile
from pathlib import Path

from docx_mcp.document import DocxDocument


def _make_doc(tmp_path: Path) -> DocxDocument:
    """Build a minimal valid DOCX in tmp_path and return an open DocxDocument."""
    path = tmp_path / "src.docx"
    W_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    W14_ns = "http://schemas.microsoft.com/office/word/2010/wordml"
    rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"

    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_ns}" xmlns:w14="{W14_ns}">'
        "<w:body>"
        f'<w:p w14:paraId="00000001"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
        f"<w:r><w:t>Title One</w:t></w:r></w:p>"
        f'<w:p w14:paraId="00000002"><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>'
        f"<w:r><w:t>Section Two</w:t></w:r></w:p>"
        f'<w:p w14:paraId="00000003"><w:pPr><w:pStyle w:val="Heading3"/></w:pPr>'
        f"<w:r><w:t>Subsection Three</w:t></w:r></w:p>"
        f'<w:p w14:paraId="00000004">'
        f"<w:r><w:rPr><w:b/></w:rPr><w:t>BoldWord</w:t></w:r>"
        f"</w:p>"
        f'<w:p w14:paraId="00000005">'
        f"<w:r><w:rPr><w:i/></w:rPr><w:t>ItalicWord</w:t></w:r>"
        f"</w:p>"
        f'<w:p w14:paraId="00000006">'
        f"<w:r><w:t>Plain text here</w:t></w:r>"
        f"</w:p>"
        f'<w:p w14:paraId="00000007">'
        f'<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
        f"<w:r><w:t>List item</w:t></w:r>"
        f"</w:p>"
        f'<w:p w14:paraId="00000008"/>'
        f"<w:tbl>"
        f"<w:tr><w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc>"
        f"<w:tc><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:tc></w:tr>"
        f"<w:tr><w:tc><w:p><w:r><w:t>R1C1</w:t></w:r></w:p></w:tc>"
        f"<w:tc><w:p><w:r><w:t>R1C2</w:t></w:r></w:p></w:tc></w:tr>"
        f"</w:tbl>"
        "<w:sectPr/>"
        "</w:body>"
        "</w:document>"
    )
    rels_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{rels_ns}"/>'
    )
    top_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{rels_ns}">'
        f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'  # noqa: E501
        "</Relationships>"
    )
    ct_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{ct_ns}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'  # noqa: E501
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'  # noqa: E501
        "</Types>"
    )
    core_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">'
        "</cp:coreProperties>"
    )

    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", ct_xml)
        zf.writestr("_rels/.rels", top_rels)
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/_rels/document.xml.rels", rels_xml)
        zf.writestr("docProps/core.xml", core_xml)

    doc = DocxDocument(str(path))
    doc.open()
    return doc


class TestMarkdownExport:
    def test_heading1_produces_h1(self, tmp_path: Path) -> None:
        doc = _make_doc(tmp_path)
        result = doc.export_markdown(str(tmp_path / "out.md"))
        content = Path(result["output_path"]).read_text()
        assert "# Title One" in content

    def test_heading2_produces_h2(self, tmp_path: Path) -> None:
        doc = _make_doc(tmp_path)
        result = doc.export_markdown(str(tmp_path / "out.md"))
        content = Path(result["output_path"]).read_text()
        assert "## Section Two" in content

    def test_heading3_produces_h3(self, tmp_path: Path) -> None:
        doc = _make_doc(tmp_path)
        result = doc.export_markdown(str(tmp_path / "out.md"))
        content = Path(result["output_path"]).read_text()
        assert "### Subsection Three" in content

    def test_bold_run_produces_double_asterisk(self, tmp_path: Path) -> None:
        doc = _make_doc(tmp_path)
        result = doc.export_markdown(str(tmp_path / "out.md"))
        content = Path(result["output_path"]).read_text()
        assert "**BoldWord**" in content

    def test_italic_run_produces_single_asterisk(self, tmp_path: Path) -> None:
        doc = _make_doc(tmp_path)
        result = doc.export_markdown(str(tmp_path / "out.md"))
        content = Path(result["output_path"]).read_text()
        assert "*ItalicWord*" in content

    def test_plain_paragraph(self, tmp_path: Path) -> None:
        doc = _make_doc(tmp_path)
        result = doc.export_markdown(str(tmp_path / "out.md"))
        content = Path(result["output_path"]).read_text()
        assert "Plain text here" in content

    def test_list_paragraph_produces_bullet(self, tmp_path: Path) -> None:
        doc = _make_doc(tmp_path)
        result = doc.export_markdown(str(tmp_path / "out.md"))
        content = Path(result["output_path"]).read_text()
        assert "- List item" in content

    def test_empty_paragraph_produces_blank_line(self, tmp_path: Path) -> None:
        doc = _make_doc(tmp_path)
        result = doc.export_markdown(str(tmp_path / "out.md"))
        content = Path(result["output_path"]).read_text()
        assert "\n\n" in content

    def test_table_produces_gfm(self, tmp_path: Path) -> None:
        doc = _make_doc(tmp_path)
        result = doc.export_markdown(str(tmp_path / "out.md"))
        content = Path(result["output_path"]).read_text()
        assert "| H1 | H2 |" in content
        assert "|" in content and "---" in content
        assert "| R1C1 | R1C2 |" in content

    def test_return_dict_keys(self, tmp_path: Path) -> None:
        doc = _make_doc(tmp_path)
        result = doc.export_markdown(str(tmp_path / "out.md"))
        assert set(result.keys()) == {"output_path", "paragraphs", "tables"}

    def test_tables_count(self, tmp_path: Path) -> None:
        doc = _make_doc(tmp_path)
        result = doc.export_markdown(str(tmp_path / "out.md"))
        assert result["tables"] == 1

    def test_default_output_path(self, tmp_path: Path) -> None:
        doc = _make_doc(tmp_path)
        result = doc.export_markdown()
        assert result["output_path"].endswith("export.md")
        assert Path(result["output_path"]).exists()

    def test_heading_space_variant(self, tmp_path: Path) -> None:
        """heading 1 (with space) variant also maps to #."""
        W_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        W14_ns = "http://schemas.microsoft.com/office/word/2010/wordml"
        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"

        doc_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document xmlns:w="{W_ns}" xmlns:w14="{W14_ns}">'
            "<w:body>"
            f'<w:p w14:paraId="00000001"><w:pPr><w:pStyle w:val="heading 1"/></w:pPr>'
            f"<w:r><w:t>Space Variant</w:t></w:r></w:p>"
            "<w:sectPr/>"
            "</w:body>"
            "</w:document>"
        )
        path2 = tmp_path / "space.docx"
        rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{rels_ns}"/>'
        )
        top_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{rels_ns}">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'  # noqa: E501
            "</Relationships>"
        )
        ct_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Types xmlns="{ct_ns}">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'  # noqa: E501
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'  # noqa: E501
            "</Types>"
        )
        core_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">'
            "</cp:coreProperties>"
        )
        with zipfile.ZipFile(path2, "w") as zf:
            zf.writestr("[Content_Types].xml", ct_xml)
            zf.writestr("_rels/.rels", top_rels)
            zf.writestr("word/document.xml", doc_xml)
            zf.writestr("word/_rels/document.xml.rels", rels_xml)
            zf.writestr("docProps/core.xml", core_xml)

        doc2 = DocxDocument(str(path2))
        doc2.open()
        result = doc2.export_markdown(str(tmp_path / "out2.md"))
        content = Path(result["output_path"]).read_text()
        assert "# Space Variant" in content

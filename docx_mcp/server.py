"""docx-mcp: MCP server for Word document editing.

Provides tools for reading and editing .docx files with full support for
track changes (w:ins/w:del), comments, footnotes, and structural validation.
"""

from __future__ import annotations

import json

from mcp.server.fastmcp import FastMCP

from docx_mcp.document import DocxDocument

mcp = FastMCP(
    "docx-mcp",
    instructions=(
        "This server edits Word (.docx) documents. Open a document first with "
        "open_document, then use other tools to read, edit, and save. Changes are "
        "made with proper Word track-changes markup so they appear as revisions "
        "in Microsoft Word / LibreOffice."
    ),
)

_doc: DocxDocument | None = None


def _js(obj: object) -> str:
    """Serialize to compact JSON for MCP responses."""
    return json.dumps(obj, indent=2, ensure_ascii=False)


def _require_doc() -> DocxDocument:
    if _doc is None:
        raise RuntimeError("No document is open. Call open_document first.")
    return _doc


# ── Document lifecycle ──────────────────────────────────────────────────────


@mcp.tool()
def open_document(path: str) -> str:
    """Open a .docx file for reading and editing.

    Unpacks the DOCX archive, parses all XML parts, and caches them in memory.
    Only one document can be open at a time; opening a new one closes the previous.

    Args:
        path: Absolute path to the .docx file.
    """
    global _doc
    if _doc is not None:
        _doc.close()
    _doc = DocxDocument(path)
    info = _doc.open()
    return _js(info)


@mcp.tool()
def close_document() -> str:
    """Close the current document and clean up temporary files."""
    global _doc
    if _doc is not None:
        _doc.close()
        _doc = None
    return "Document closed."


@mcp.tool()
def create_document(
    output_path: str,
    template_path: str | None = None,
) -> str:
    """Create a new blank .docx document (or from a .dotx template).

    The document is automatically opened for editing after creation.
    Use save_document to save changes, or start editing immediately
    with insert_text, add_table, etc.

    Args:
        output_path: Path for the new .docx file.
        template_path: Optional path to a .dotx template file.
    """
    global _doc
    if _doc is not None:
        _doc.close()
    _doc = DocxDocument.create(output_path, template_path=template_path)
    return _js(_doc.get_info())


@mcp.tool()
def create_from_markdown(
    output_path: str,
    md_path: str | None = None,
    markdown: str | None = None,
    template_path: str | None = None,
) -> str:
    """Create a new .docx document from markdown content.

    Supports full GitHub-Flavored Markdown: headings, bold/italic/strikethrough,
    links, images, bullet/numbered/nested lists, code blocks, blockquotes,
    tables, footnotes, and task lists. Smart typography (curly quotes, em/en
    dashes, ellipses) is applied automatically.

    Provide exactly one of md_path or markdown. The document is automatically
    opened for editing after creation.

    Args:
        output_path: Path for the new .docx file.
        md_path: Path to a .md file. Mutually exclusive with markdown.
        markdown: Raw markdown text. Mutually exclusive with md_path.
        template_path: Optional path to a .dotx template file.
    """
    if md_path and markdown:
        return "Error: md_path and markdown are mutually exclusive — provide one, not both."
    if not md_path and not markdown:
        return "Error: Either md_path or markdown must be provided."

    global _doc
    if _doc is not None:
        _doc.close()

    _doc = DocxDocument.create(output_path, template_path=template_path)

    base_dir = None
    if md_path:
        from pathlib import Path

        p = Path(md_path)
        if not p.exists():
            return f"Error: Markdown file not found: {md_path}"
        markdown = p.read_text(encoding="utf-8")
        base_dir = p.parent

    from docx_mcp.markdown import MarkdownConverter

    MarkdownConverter.convert(_doc, markdown, base_dir=base_dir)
    _doc.save(backup=False)
    return _js(_doc.get_info())


@mcp.tool()
def get_document_info() -> str:
    """Get overview stats: paragraph count, headings, footnotes, comments, images."""
    return _js(_require_doc().get_info())


# ── Reading ─────────────────────────────────────────────────────────────────


@mcp.tool()
def get_headings() -> str:
    """Get the document heading structure with levels, text, and paraIds.

    Returns a list of headings in document order, each with:
    - level (1-9)
    - text (heading content)
    - style (e.g., "Heading1")
    - paraId (unique paragraph identifier for targeting edits)
    """
    return _js(_require_doc().get_headings())


@mcp.tool()
def search_text(query: str, regex: bool = False) -> str:
    """Search for text across the document body, footnotes, and comments.

    Args:
        query: Text to search for (case-insensitive), or a regex pattern.
        regex: If true, treat query as a Python regular expression.

    Returns matching paragraphs with their paraId, source part, and context.
    """
    return _js(_require_doc().search_text(query, regex=regex))


@mcp.tool()
def get_paragraph(para_id: str) -> str:
    """Get the full text and style of a specific paragraph by its paraId.

    Args:
        para_id: The 8-character hex paraId (e.g., "1A2B3C4D").
    """
    return _js(_require_doc().get_paragraph(para_id))


# ── Tables ─────────────────────────────────────────────────────────────────


@mcp.tool()
def get_tables() -> str:
    """Get all tables with row/column counts and cell text content."""
    return _js(_require_doc().get_tables())


@mcp.tool()
def add_table(
    para_id: str,
    rows: int,
    cols: int,
    author: str = "Claude",
) -> str:
    """Insert a new table after a paragraph with tracked insertion.

    Args:
        para_id: paraId of the paragraph to insert after.
        rows: Number of rows.
        cols: Number of columns.
        author: Author name for the revision.
    """
    return _js(_require_doc().add_table(para_id, rows, cols, author=author))


@mcp.tool()
def modify_cell(
    table_idx: int,
    row: int,
    col: int,
    text: str,
    author: str = "Claude",
) -> str:
    """Modify a table cell with tracked changes (delete old, insert new).

    Args:
        table_idx: Table index (0-based).
        row: Row index (0-based).
        col: Column index (0-based).
        text: New cell text.
        author: Author name for the revision.
    """
    return _js(_require_doc().modify_cell(table_idx, row, col, text, author=author))


@mcp.tool()
def add_table_row(
    table_idx: int,
    row_idx: int = -1,
    cells: list[str] | None = None,
    author: str = "Claude",
) -> str:
    """Add a row to a table with tracked insertion.

    Args:
        table_idx: Table index (0-based).
        row_idx: Insert at this row index. -1 = append at end.
        cells: Cell text content. Empty = empty cells.
        author: Author name for the revision.
    """
    idx = row_idx if row_idx >= 0 else None
    return _js(_require_doc().add_table_row(table_idx, row_idx=idx, cells=cells, author=author))


@mcp.tool()
def delete_table_row(
    table_idx: int,
    row_idx: int,
    author: str = "Claude",
) -> str:
    """Delete a table row with tracked changes.

    Args:
        table_idx: Table index (0-based).
        row_idx: Row index to delete (0-based).
        author: Author name for the revision.
    """
    return _js(_require_doc().delete_table_row(table_idx, row_idx, author=author))


@mcp.tool()
def merge_cells(
    table_index: int,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> str:
    """Merge a rectangular range of cells. Horizontal: gridSpan. Vertical: vMerge."""
    doc = _require_doc()
    return _js(doc.merge_cells(table_index, start_row, start_col, end_row, end_col))


@mcp.tool()
def set_header_row(table_index: int) -> str:
    """Mark the first row as a repeating header row."""
    doc = _require_doc()
    return _js(doc.set_header_row(table_index))


@mcp.tool()
def set_column_widths(table_index: int, widths_cm: list[float]) -> str:
    """Set column widths in cm. len(widths_cm) must match column count."""
    doc = _require_doc()
    return _js(doc.set_column_widths(table_index, widths_cm))


@mcp.tool()
def csv_to_table(para_id: str, csv_text: str, header_row: bool = True) -> str:
    """Insert a table from CSV text."""
    doc = _require_doc()
    return _js(doc.csv_to_table(para_id, csv_text, header_row))


@mcp.tool()
def table_to_csv(table_index: int) -> str:
    """Export a table as CSV string."""
    doc = _require_doc()
    return _js(doc.table_to_csv(table_index))


@mcp.tool()
def delete_table(table_idx: int) -> str:
    """Delete a table by index (0-based). Raises IndexError if out of range."""
    return _js(_require_doc().delete_table(table_idx))


@mcp.tool()
def add_column_to_table(table_idx: int, header_text: str = "") -> str:
    """Add a new column to every row of a table. First row gets header_text."""
    return _js(_require_doc().add_column_to_table(table_idx, header_text=header_text))


@mcp.tool()
def delete_column_from_table(table_idx: int, col_idx: int) -> str:
    """Delete a column (0-based) from every row of a table."""
    return _js(_require_doc().delete_column_from_table(table_idx, col_idx))


@mcp.tool()
def set_cell_width(table_idx: int, row_idx: int, col_idx: int, width_mm: float) -> str:
    """Set the width of a table cell in millimetres (stored as DXA)."""
    return _js(_require_doc().set_cell_width(table_idx, row_idx, col_idx, width_mm))


@mcp.tool()
def set_cell_vertical_alignment(
    table_idx: int, row_idx: int, col_idx: int, alignment: str
) -> str:
    """Set vertical alignment of a table cell: top, center, or bottom."""
    return _js(
        _require_doc().set_cell_vertical_alignment(table_idx, row_idx, col_idx, alignment)
    )


@mcp.tool()
def set_row_height(
    table_idx: int, row_idx: int, height_mm: float, rule: str = "exact"
) -> str:
    """Set row height in millimetres. rule: exact, atLeast, or auto."""
    return _js(_require_doc().set_row_height(table_idx, row_idx, height_mm, rule=rule))


@mcp.tool()
def set_table_alignment(table_idx: int, alignment: str) -> str:
    """Set table alignment: left, center, or right."""
    return _js(_require_doc().set_table_alignment(table_idx, alignment))


@mcp.tool()
def set_table_borders(
    table_idx: int,
    border_style: str = "single",
    color: str = "000000",
    size: int = 4,
) -> str:
    """Set borders on all six sides of a table (top, bottom, left, right, insideH, insideV)."""
    return _js(_require_doc().set_table_borders(table_idx, border_style=border_style, color=color, size=size))


@mcp.tool()
def set_cell_shading(
    table_idx: int,
    row_idx: int,
    col_idx: int,
    fill_color: str,
    pattern: str = "clear",
) -> str:
    """Set background shading fill color on a table cell."""
    return _js(_require_doc().set_cell_shading(table_idx, row_idx, col_idx, fill_color, pattern=pattern))


@mcp.tool()
def set_table_style(table_idx: int, style_name: str) -> str:
    """Apply a named table style (e.g. TableGrid, LightShading-Accent1) to a table."""
    return _js(_require_doc().set_table_style(table_idx, style_name))


# ── Lists ──────────────────────────────────────────────────────────────────


@mcp.tool()
def add_list(
    para_ids: list[str],
    style: str = "bullet",
) -> str:
    """Apply list formatting to paragraphs (bullet or numbered).

    Creates numbering definitions and sets w:numPr on each target paragraph.

    Args:
        para_ids: List of paraIds to format as list items.
        style: "bullet" or "numbered".
    """
    return _js(_require_doc().add_list(para_ids, style=style))


# ── Styles ─────────────────────────────────────────────────────────────────


@mcp.tool()
def get_styles() -> str:
    """Get all defined styles with ID, name, type, and base style."""
    return _js(_require_doc().get_styles())


@mcp.tool()
def create_style(
    name: str,
    style_type: str,
    based_on: str | None = None,
    next_style: str | None = None,
) -> str:
    """Create a new style in the document.

    Args:
        name: Style name (used as styleId after removing spaces).
        style_type: "paragraph", "character", "table", or "numbering".
        based_on: Optional styleId this style inherits from.
        next_style: Optional styleId applied to the next paragraph.
    """
    return _js(_require_doc().create_style(name, style_type, based_on=based_on, next_style=next_style))


@mcp.tool()
def update_style(
    name: str,
    based_on: str | None = None,
    next_style: str | None = None,
) -> str:
    """Update an existing style's basedOn and/or next properties.

    Args:
        name: Style name or styleId (case-insensitive).
        based_on: New basedOn styleId (replaces existing).
        next_style: New next styleId (replaces existing).
    """
    return _js(_require_doc().update_style(name, based_on=based_on, next_style=next_style))


@mcp.tool()
def delete_style(name: str) -> str:
    """Delete a style from the document.

    Args:
        name: Style name or styleId (case-insensitive).
    """
    return _js(_require_doc().delete_style(name))


# ── Headers / Footers ──────────────────────────────────────────────────────


@mcp.tool()
def get_headers_footers() -> str:
    """Get all headers and footers with their text content."""
    return _js(_require_doc().get_headers_footers())


@mcp.tool()
def edit_header_footer(
    location: str,
    old_text: str,
    new_text: str,
    author: str = "Claude",
) -> str:
    """Edit text in a header or footer with tracked changes.

    Args:
        location: "header" or "footer" (matches first found).
        old_text: Text to find and replace.
        new_text: Replacement text.
        author: Author name for the revision.
    """
    return _js(_require_doc().edit_header_footer(location, old_text, new_text, author=author))


@mcp.tool()
def delete_header(location: str = "default") -> str:
    """Delete a header by location: default, first, or even."""
    return _js(_require_doc().delete_header(location=location))


@mcp.tool()
def delete_footer(location: str = "default") -> str:
    """Delete a footer by location: default, first, or even."""
    return _js(_require_doc().delete_footer(location=location))


# ── Properties ─────────────────────────────────────────────────────────────


@mcp.tool()
def get_properties() -> str:
    """Get core document properties (title, creator, subject, dates, revision)."""
    return _js(_require_doc().get_properties())


@mcp.tool()
def set_properties(
    title: str = "",
    creator: str = "",
    subject: str = "",
    description: str = "",
) -> str:
    """Set core document properties.

    Args:
        title: Document title. Empty = unchanged.
        creator: Document author/creator. Empty = unchanged.
        subject: Document subject. Empty = unchanged.
        description: Document description. Empty = unchanged.
    """
    return _js(
        _require_doc().set_properties(
            title=title or None,
            creator=creator or None,
            subject=subject or None,
            description=description or None,
        )
    )


# ── Images ─────────────────────────────────────────────────────────────────


@mcp.tool()
def get_images() -> str:
    """Get all embedded images with rId, filename, content type, and dimensions."""
    return _js(_require_doc().get_images())


@mcp.tool()
def insert_image(
    para_id: str,
    image_path: str,
    width_emu: int = 2000000,
    height_emu: int = 2000000,
) -> str:
    """Insert an image into the document after a paragraph.

    Args:
        para_id: paraId of the paragraph to insert after.
        image_path: Absolute path to the image file.
        width_emu: Image width in EMUs (914400 = 1 inch).
        height_emu: Image height in EMUs.
    """
    return _js(
        _require_doc().insert_image(para_id, image_path, width_emu=width_emu, height_emu=height_emu)
    )


@mcp.tool()
def insert_floating_image(
    para_id: str,
    image_path: str,
    width_cm: float,
    height_cm: float,
    h_pos: float = 0.0,
    v_pos: float = 0.0,
    wrap: str = "square",
) -> str:
    """Insert a floating (anchored) image. wrap: square|topbottom|none."""
    doc = _require_doc()
    return _js(doc.insert_floating_image(para_id, image_path, width_cm, height_cm, h_pos, v_pos, wrap))


# ── Endnotes ───────────────────────────────────────────────────────────────


@mcp.tool()
def get_endnotes() -> str:
    """Get all endnotes with their ID and text content."""
    return _js(_require_doc().get_endnotes())


@mcp.tool()
def add_endnote(para_id: str, text: str) -> str:
    """Add an endnote to a paragraph.

    Creates the endnote definition in endnotes.xml and adds a superscript
    reference in the target paragraph.

    Args:
        para_id: paraId of the paragraph to attach the endnote to.
        text: Endnote text content.
    """
    return _js(_require_doc().add_endnote(para_id, text))


@mcp.tool()
def validate_endnotes() -> str:
    """Cross-reference endnote IDs between document.xml and endnotes.xml.

    Checks that every endnote reference has a matching definition,
    and flags orphaned definitions with no reference.
    """
    return _js(_require_doc().validate_endnotes())


# ── Footnotes ───────────────────────────────────────────────────────────────


@mcp.tool()
def get_footnotes() -> str:
    """List all footnotes with their ID and text content."""
    return _js(_require_doc().get_footnotes())


@mcp.tool()
def add_footnote(para_id: str, text: str) -> str:
    """Add a footnote to a paragraph.

    Creates the footnote definition in footnotes.xml and adds a superscript
    reference number at the end of the target paragraph.

    Args:
        para_id: paraId of the paragraph to attach the footnote to.
        text: Footnote text content.
    """
    return _js(_require_doc().add_footnote(para_id, text))


@mcp.tool()
def validate_footnotes() -> str:
    """Cross-reference footnote IDs between document.xml and footnotes.xml.

    Checks that every footnote reference in the document body has a matching
    definition, and flags orphaned definitions with no reference.
    """
    return _js(_require_doc().validate_footnotes())


@mcp.tool()
def update_footnote(footnote_id: int, text: str) -> str:
    """Update the text of an existing footnote.

    Args:
        footnote_id: The numeric ID of the footnote to update (must be >= 1).
        text: The new text content for the footnote.
    """
    return _js(_require_doc().update_footnote(footnote_id, text))


@mcp.tool()
def delete_footnote(footnote_id: int) -> str:
    """Delete a footnote and its in-body reference.

    Removes the footnote definition from footnotes.xml and removes the
    footnoteReference run from the document body.

    Args:
        footnote_id: The numeric ID of the footnote to delete.
    """
    return _js(_require_doc().delete_footnote(footnote_id))


@mcp.tool()
def update_endnote(endnote_id: int, text: str) -> str:
    """Update the text of an existing endnote.

    Args:
        endnote_id: The numeric ID of the endnote to update (must be >= 1).
        text: The new text content for the endnote.
    """
    return _js(_require_doc().update_endnote(endnote_id, text))


@mcp.tool()
def delete_endnote(endnote_id: int) -> str:
    """Delete an endnote and its in-body reference.

    Removes the endnote definition from endnotes.xml and removes the
    endnoteReference run from the document body.

    Args:
        endnote_id: The numeric ID of the endnote to delete.
    """
    return _js(_require_doc().delete_endnote(endnote_id))


# ── Sections / Page breaks ─────────────────────────────────────────────


@mcp.tool()
def add_page_break(para_id: str) -> str:
    """Insert a page break after a paragraph.

    Creates a new paragraph containing a page break element.

    Args:
        para_id: paraId of the paragraph to insert after.
    """
    return _js(_require_doc().add_page_break(para_id))


@mcp.tool()
def add_section_break(
    para_id: str,
    break_type: str = "nextPage",
) -> str:
    """Add a section break at a paragraph.

    Inserts a section properties element in the paragraph, marking it as
    the last paragraph of its section.

    Args:
        para_id: paraId of the paragraph to place the section break on.
        break_type: "nextPage", "continuous", "evenPage", or "oddPage".
    """
    return _js(_require_doc().add_section_break(para_id, break_type=break_type))


@mcp.tool()
def set_section_properties(
    para_id: str = "",
    width: int = 0,
    height: int = 0,
    orientation: str = "",
    margin_top: int = 0,
    margin_bottom: int = 0,
    margin_left: int = 0,
    margin_right: int = 0,
) -> str:
    """Modify section properties (page size, orientation, margins).

    Args:
        para_id: paraId of paragraph with section break. Empty = body section.
        width: Page width in DXA (twips). 12240 = 8.5 inches.
        height: Page height in DXA. 15840 = 11 inches.
        orientation: "portrait" or "landscape". Empty = unchanged.
        margin_top: Top margin in DXA. 0 = unchanged.
        margin_bottom: Bottom margin in DXA. 0 = unchanged.
        margin_left: Left margin in DXA. 0 = unchanged.
        margin_right: Right margin in DXA. 0 = unchanged.
    """
    return _js(
        _require_doc().set_section_properties(
            para_id=para_id or None,
            width=width or None,
            height=height or None,
            orientation=orientation or None,
            margin_top=margin_top or None,
            margin_bottom=margin_bottom or None,
            margin_left=margin_left or None,
            margin_right=margin_right or None,
        )
    )



@mcp.tool()
def set_page_size(width_mm: float, height_mm: float, para_id: str | None = None) -> str:
    """Set page size from millimetre values.

    Args:
        width_mm: Page width in mm (e.g. 210 for A4, 215.9 for Letter).
        height_mm: Page height in mm (e.g. 297 for A4, 279.4 for Letter).
        para_id: paraId of paragraph with section break. None = body section.
    """
    return _js(_require_doc().set_page_size(width_mm, height_mm, para_id=para_id))


@mcp.tool()
def set_page_margins(
    top_mm: float | None = None,
    bottom_mm: float | None = None,
    left_mm: float | None = None,
    right_mm: float | None = None,
    para_id: str | None = None,
) -> str:
    """Set page margins from millimetre values.

    Args:
        top_mm: Top margin in mm. None = unchanged.
        bottom_mm: Bottom margin in mm. None = unchanged.
        left_mm: Left margin in mm. None = unchanged.
        right_mm: Right margin in mm. None = unchanged.
        para_id: paraId of paragraph with section break. None = body section.
    """
    return _js(
        _require_doc().set_page_margins(
            top_mm=top_mm,
            bottom_mm=bottom_mm,
            left_mm=left_mm,
            right_mm=right_mm,
            para_id=para_id,
        )
    )


@mcp.tool()
def set_page_orientation(orientation: str, para_id: str | None = None) -> str:
    """Set page orientation, swapping width/height dimensions if needed.

    Args:
        orientation: "portrait" or "landscape".
        para_id: paraId of paragraph with section break. None = body section.
    """
    return _js(_require_doc().set_page_orientation(orientation, para_id=para_id))


# ── Cross-references ──────────────────────────────────────────────────


@mcp.tool()
def add_cross_reference(
    source_para_id: str,
    target_para_id: str,
    text: str,
) -> str:
    """Add a cross-reference link from one paragraph to another.

    Creates a bookmark at the target (if none exists) and inserts a hyperlink
    in the source paragraph with the given display text.

    Args:
        source_para_id: paraId of the paragraph where the link appears.
        target_para_id: paraId of the paragraph being referenced.
        text: Display text for the cross-reference link.
    """
    return _js(_require_doc().add_cross_reference(source_para_id, target_para_id, text))


# ── Protection ─────────────────────────────────────────────────────────────


@mcp.tool()
def set_document_protection(
    edit: str,
    password: str = "",
) -> str:
    """Set document protection in settings.xml.

    Args:
        edit: Protection type — "trackedChanges", "comments", "readOnly",
              "forms", or "none" (removes protection).
        password: Optional password (hashed with SHA-512). Empty = no password.
    """
    return _js(_require_doc().set_document_protection(edit, password=password or None))


# ── Merge ──────────────────────────────────────────────────────────────────


@mcp.tool()
def merge_documents(source_path: str) -> str:
    """Merge another DOCX document's content into the current document.

    Appends body paragraphs and tables from the source. ParaIds are
    automatically remapped to avoid collisions.

    Args:
        source_path: Absolute path to the DOCX file to merge in.
    """
    return _js(_require_doc().merge_documents(source_path))


# ── Validation ──────────────────────────────────────────────────────────────


@mcp.tool()
def validate_paraids() -> str:
    """Check paraId uniqueness across all document parts.

    ParaIds must be unique across document.xml, footnotes.xml, headers, footers,
    and comments. They must also be valid 8-digit hex values < 0x80000000.
    """
    return _js(_require_doc().validate_paraids())


@mcp.tool()
def insert_watermark(text: str, diagonal: bool = True) -> str:
    """Insert a VML watermark into the document's default header.

    Places a <v:shape> with a <v:textpath> inside the default header, which is
    the standard Word watermark pattern.

    Args:
        text: Watermark text (e.g. "DRAFT", "CONFIDENTIAL").
        diagonal: If True (default), diagonal orientation; if False, horizontal.
    """
    return _js(_require_doc().insert_watermark(text, diagonal=diagonal))


@mcp.tool()
def remove_watermark() -> str:
    """Remove VML watermarks (e.g., DRAFT) from all document headers.

    Detects and removes <v:shape> elements with <v:textpath> inside header XML
    files — the standard pattern for Word watermarks.
    """
    return _js(_require_doc().remove_watermark())


@mcp.tool()
def audit_document() -> str:
    """Run a comprehensive structural audit of the document.

    Checks:
    - Footnote cross-references (references vs definitions)
    - ParaId uniqueness and range validity
    - Heading level continuity (no skips like H2 -> H4)
    - Bookmark pairing (start/end matching)
    - Relationship targets (all referenced files exist)
    - Image references (all embedded images exist)
    - Residual artifacts (DRAFT, TODO, FIXME markers)

    Returns a detailed report with an overall valid/invalid status.
    """
    return _js(_require_doc().audit())


# ── Track changes ───────────────────────────────────────────────────────────


@mcp.tool()
def insert_text(
    para_id: str,
    text: str,
    position: str = "end",
    author: str = "Claude",
    context_before: str = "",
    context_after: str = "",
    ignore_case: bool = False,
) -> str:
    """Insert text with Word track-changes markup (appears as a green underlined insertion in Word).

    Args:
        para_id: paraId of the target paragraph.
        text: Text to insert.
        position: Where to insert — "start", "end", or a substring to insert after.
        author: Author name for the revision (shown in Word's review pane).
        context_before: Text immediately before the insertion point (for precise anchoring).
        context_after: Text immediately after the insertion point (for precise anchoring).
        ignore_case: If True, match context_before/context_after case-insensitively.
    """
    return _js(_require_doc().insert_text(
        para_id, text, position=position, author=author,
        context_before=context_before, context_after=context_after,
        ignore_case=ignore_case,
    ))


@mcp.tool()
def delete_text(
    para_id: str,
    text: str,
    author: str = "Claude",
    context_before: str = "",
    context_after: str = "",
    ignore_case: bool = False,
) -> str:
    """Mark text as deleted with Word track-changes markup (appears as red strikethrough in Word).

    Finds the text within the paragraph (across run boundaries if needed) and wraps it in
    deletion markup. Provide context_before/context_after to disambiguate when the same text
    appears multiple times, or when the text contains smart quotes / special whitespace.

    Args:
        para_id: paraId of the target paragraph.
        text: Text to mark as deleted (ASCII quotes/dashes/spaces match their Unicode equivalents).
        author: Author name for the revision.
        context_before: Text immediately before the target (for precise anchoring).
        context_after: Text immediately after the target (for precise anchoring).
        ignore_case: If True, match text and context case-insensitively (output preserves original casing).
    """
    return _js(_require_doc().delete_text(
        para_id, text, author=author,
        context_before=context_before, context_after=context_after,
        ignore_case=ignore_case,
    ))


@mcp.tool()
def replace_text(
    para_id: str,
    find: str,
    replace: str,
    author: str = "Claude",
    context_before: str = "",
    context_after: str = "",
    ignore_case: bool = False,
) -> str:
    """Replace text with tracked changes markup (deletion + insertion).

    Only the actually-changed portion is marked; common leading/trailing text is
    left as plain runs (collapseDiff behaviour).

    Args:
        para_id: paraId of the target paragraph.
        find: Text to find and replace (may span run boundaries).
        replace: Replacement text.
        author: Author name for the revision.
        context_before: Text immediately before the target (for precise anchoring).
        context_after: Text immediately after the target (for precise anchoring).
    """
    return _js(_require_doc().replace_text(
        para_id, find=find, replace=replace, author=author,
        context_before=context_before, context_after=context_after,
        ignore_case=ignore_case,
    ))


@mcp.tool()
def get_tracked_changes() -> str:
    """Return all pending tracked changes (insertions and deletions) as a JSON list.

    Each entry contains: type, change_id, author, date, para_id, text.
    Changes are returned in document order.
    """
    return _js(_require_doc().get_tracked_changes())


@mcp.tool()
def accept_changes(author: str = "") -> str:
    """Accept tracked changes — keep insertions, remove deletions.

    Args:
        author: If set, only accept changes by this author. Empty = all changes.
    """
    return _js(_require_doc().accept_changes(author=author or None))


@mcp.tool()
def reject_changes(author: str = "") -> str:
    """Reject tracked changes — remove insertions, restore deleted text.

    Args:
        author: If set, only reject changes by this author. Empty = all changes.
    """
    return _js(_require_doc().reject_changes(author=author or None))


@mcp.tool()
def accept_change(change_id: int) -> str:
    """Accept a single tracked change by its change_id.

    For insertions: keeps the inserted text (unwraps w:ins).
    For deletions: discards the deleted text (removes w:del).

    Args:
        change_id: The integer id attribute of the w:ins or w:del element.
    """
    return _js(_require_doc().accept_change(change_id))


@mcp.tool()
def reject_change(change_id: int) -> str:
    """Reject a single tracked change by its change_id.

    For insertions: discards the inserted text (removes w:ins).
    For deletions: keeps the deleted text (unwraps w:del, restoring text).

    Args:
        change_id: The integer id attribute of the w:ins or w:del element.
    """
    return _js(_require_doc().reject_change(change_id))


@mcp.tool()
def accept_all_changes() -> str:
    """Accept all tracked changes in document order.

    Returns a JSON object with the count of accepted changes: {"accepted": int}.
    """
    return _js(_require_doc().accept_all_changes())


@mcp.tool()
def reject_all_changes() -> str:
    """Reject all tracked changes in document order.

    Returns a JSON object with the count of rejected changes: {"rejected": int}.
    """
    return _js(_require_doc().reject_all_changes())


@mcp.tool()
def set_formatting(
    para_id: str,
    text: str,
    bold: bool = False,
    italic: bool = False,
    underline: str = "",
    color: str = "",
    author: str = "Claude",
) -> str:
    """Apply character formatting to text with tracked-change markup.

    Finds the text within the paragraph, splits the run if needed, and applies
    formatting with rPrChange so it appears as a format revision in Word.

    Args:
        para_id: paraId of the target paragraph.
        text: Exact text to format.
        bold: Apply bold formatting.
        italic: Apply italic formatting.
        underline: Underline style (e.g., "single", "double"). Empty = no underline.
        color: Font color as hex (e.g., "FF0000"). Empty = no color change.
        author: Author name for the revision.
    """
    return _js(
        _require_doc().set_formatting(
            para_id,
            text,
            bold=bold,
            italic=italic,
            underline=underline or None,
            color=color or None,
            author=author,
        )
    )


# ── Comments ────────────────────────────────────────────────────────────────


@mcp.tool()
def get_comments() -> str:
    """List all comments with their ID, author, date, and text."""
    return _js(_require_doc().get_comments())


@mcp.tool()
def add_comment(
    para_id: str,
    text: str,
    author: str = "Claude",
) -> str:
    """Add a comment anchored to a paragraph.

    Creates the comment in comments.xml and adds range markers
    (commentRangeStart/End) around the paragraph content.

    Args:
        para_id: paraId of the paragraph to comment on.
        text: Comment text.
        author: Author name (shown in Word's comment sidebar).
    """
    return _js(_require_doc().add_comment(para_id, text, author=author))


@mcp.tool()
def reply_to_comment(
    parent_id: int,
    text: str,
    author: str = "Claude",
) -> str:
    """Reply to an existing comment (creates a threaded reply).

    Args:
        parent_id: ID of the comment to reply to.
        text: Reply text.
        author: Author name.
    """
    return _js(_require_doc().reply_to_comment(parent_id, text, author=author))


@mcp.tool()
def update_comment(comment_id: int, text: str) -> str:
    """Replace the text of an existing comment.

    Args:
        comment_id: ID of the comment to update.
        text: New comment text.
    """
    return _js(_require_doc().update_comment(comment_id, text))


@mcp.tool()
def delete_comment(comment_id: int) -> str:
    """Delete a comment and remove its range markers from the document.

    Args:
        comment_id: ID of the comment to delete.
    """
    return _js(_require_doc().delete_comment(comment_id))


@mcp.tool()
def resolve_comment(comment_id: int) -> str:
    """Mark a comment as resolved (sets w15:done='1' in commentsExtended.xml).

    Args:
        comment_id: ID of the comment to resolve.
    """
    return _js(_require_doc().resolve_comment(comment_id))


@mcp.tool()
def list_comment_threads() -> str:
    """List all comment threads (root comments with their replies).

    Returns a list of thread dicts: {root: {id, author, date, text}, replies: [...]}.
    """
    return _js(_require_doc().list_comment_threads())


# ── Save ────────────────────────────────────────────────────────────────────


@mcp.tool()
def save_document(output_path: str = "") -> str:
    """Save all changes back to a .docx file.

    Serializes modified XML parts and repacks into a DOCX archive.

    Args:
        output_path: Path for the output file. If empty, overwrites the original.
    """
    doc = _require_doc()
    path = output_path if output_path else None
    return _js(doc.save(path))


@mcp.tool()
def scrub_pii(
    output_path: str = "",
    entities: list[str] | None = None,
    confidence_threshold: float = 0.35,
    dry_run: bool = False,
    also_sanitize_metadata: bool = True,
    redact_authors_as: str = "REDACTED",
) -> str:
    """Detect and permanently redact PII from the open document using Presidio + spaCy NER.

    NER model (en_core_web_trf, ~430MB) is downloaded automatically on first use.

    Detects: PERSON, EMAIL_ADDRESS, PHONE_NUMBER, CREDIT_CARD, SSN, IP_ADDRESS,
             IBAN_CODE, US_BANK_NUMBER, US_PASSPORT, and more via Presidio.

    Redacted text is replaced with a solid black DrawingML rectangle — true XML
    redaction where the original text is deleted entirely from the OOXML, not
    merely hidden by formatting.

    Args:
        output_path: Destination path. Required when dry_run=False.
        entities: Presidio entity types to redact. None = all detected types.
        confidence_threshold: Presidio score floor (default 0.35).
        dry_run: If True, detect only — return entity list, write no file.
        also_sanitize_metadata: Apply level-3 metadata sanitization (default True).
        redact_authors_as: Replacement author string for metadata pass.
    """
    return _js(_require_doc().scrub_pii(
        output_path,
        entities=entities,
        confidence_threshold=confidence_threshold,
        dry_run=dry_run,
        also_sanitize_metadata=also_sanitize_metadata,
        redact_authors_as=redact_authors_as,
    ))


@mcp.tool()
def sanitize_metadata(
    output_path: str,
    level: int = 1,
    redact_authors_as: str = "",
) -> str:
    """Write a sanitized copy of the open document to output_path.

    Level 1: Remove rsid session-fingerprint attributes from document.xml.
    Level 2: + Replace tracked-change author names (w:author on w:ins/w:del).
    Level 3: + Clear creator/lastModifiedBy/revision in docProps/core.xml
             + Clear Company in docProps/app.xml
             + Remove attachedTemplate reference from word/settings.xml

    Args:
        output_path: Destination path for the sanitized DOCX. Must be non-empty.
        level: Sanitization depth (1, 2, or 3). Default 1.
        redact_authors_as: Replacement author string for level 2+. Default "Anonymous".
    """
    return _js(_require_doc().sanitize_metadata(
        output_path,
        level=level,
        redact_authors_as=redact_authors_as,
    ))


@mcp.tool()
def compare_documents(
    base_path: str,
    revised_path: str,
    output_path: str = "",
) -> str:
    """Diff two DOCX files and produce a tracked-change document.

    Paragraph-level LCS diff:
      - Unchanged paragraphs copied verbatim.
      - Deleted paragraphs (in base, absent in revised) wrapped in w:del.
      - Inserted paragraphs (in revised, absent in base) wrapped in w:ins.
      - Modified paragraphs (1:1 replacement) get word-level del+ins inline.

    The output is a valid DOCX readable in Word/LibreOffice showing the changes
    as tracked revisions.

    Args:
        base_path: Path to the original DOCX.
        revised_path: Path to the revised DOCX.
        output_path: Destination path. Auto-generated if empty.
    """
    return _js(DocxDocument.compare_documents(base_path, revised_path, output_path))


@mcp.tool()
def list_parts() -> str:
    """List all XML parts (files) in the open DOCX zip."""
    return _js(_require_doc().list_parts())


@mcp.tool()
def read_part(part_path: str) -> str:
    """Read raw XML of any DOCX part (e.g. 'word/document.xml').
    Use list_parts() to discover available parts.
    """
    return _js(_require_doc().read_part(part_path))


@mcp.tool()
def write_part(part_path: str, xml: str) -> str:
    """Replace a DOCX part with new XML. Validates well-formedness first.
    WARNING: Direct XML manipulation can corrupt the document if used incorrectly.
    """
    return _js(_require_doc().write_part(part_path, xml))


@mcp.tool()
def xpath_query(xpath: str, part: str = "word/document.xml") -> str:
    """Run XPath against any DOCX part. Pre-bound namespaces: w, w14, r, wp, a, mc.

    Examples:
      xpath="//w:p" — all paragraphs
      xpath="//w:t/text()" — all text content
      xpath="//w:p[w:pPr/w:pStyle/@w:val='Heading1']" — Heading 1 paragraphs
    """
    return _js(_require_doc().xpath_query(xpath, part))


# ── Bookmarks ───────────────────────────────────────────────────────────────


@mcp.tool()
def list_bookmarks() -> str:
    """List all bookmarks in the document."""
    return _js(_require_doc().list_bookmarks())


@mcp.tool()
def add_bookmark(para_id: str, name: str) -> str:
    """Add a named bookmark wrapping the specified paragraph."""
    return _js(_require_doc().add_bookmark(para_id, name))


@mcp.tool()
def remove_bookmark(name: str) -> str:
    """Remove a bookmark by name (keeps paragraph content)."""
    return _js(_require_doc().remove_bookmark(name))


@mcp.tool()
def get_bookmarked_text(name: str) -> str:
    """Get the text content within a named bookmark."""
    return _js(_require_doc().get_bookmarked_text(name))


# ── Hyperlinks ──────────────────────────────────────────────────────────────


@mcp.tool()
def list_hyperlinks() -> str:
    """List all hyperlinks in the document.

    Returns a list of dicts with keys: id, url_or_anchor, text, para_id, type.
    type is "external" (r:id based) or "internal" (w:anchor based).
    """
    return _js(_require_doc().list_hyperlinks())


@mcp.tool()
def add_hyperlink(para_id: str, text: str, url: str) -> str:
    """Append an external hyperlink at the end of a paragraph.

    Creates a relationship entry in word/_rels/document.xml.rels and a
    w:hyperlink element with a Hyperlink-styled run in word/document.xml.

    Args:
        para_id: w14:paraId of the target paragraph.
        text: Display text for the hyperlink.
        url: The URL the hyperlink points to.
    """
    return _js(_require_doc().add_hyperlink(para_id, text, url))


@mcp.tool()
def add_internal_link(para_id: str, text: str, bookmark: str) -> str:
    """Append an internal anchor hyperlink (w:anchor) at the end of a paragraph.

    Internal links reference bookmarks by name and do NOT add a relationship.

    Args:
        para_id: w14:paraId of the target paragraph.
        text: Display text for the hyperlink.
        bookmark: Name of the bookmark to link to.
    """
    return _js(_require_doc().add_internal_link(para_id, text, bookmark))


@mcp.tool()
def remove_hyperlink(para_id: str, url_or_anchor: str) -> str:
    """Remove a hyperlink wrapper, preserving the text runs inside.

    The paragraph text remains; only the w:hyperlink element is unwrapped.

    Args:
        para_id: w14:paraId of the paragraph containing the hyperlink.
        url_or_anchor: URL (for external) or bookmark name (for internal) to match.
    """
    return _js(_require_doc().remove_hyperlink(para_id, url_or_anchor))


@mcp.tool()
def update_hyperlink(r_id: str, new_url: str) -> str:
    """Update the target URL of an existing external hyperlink relationship.

    Args:
        r_id: The relationship ID (e.g. "rId7") to update.
        new_url: The new target URL.
    """
    return _js(_require_doc().update_hyperlink(r_id, new_url))


# ── Fields ──────────────────────────────────────────────────────────────────


@mcp.tool()
def add_field(para_id: str, field_code: str, cached_value: str = "") -> str:
    """Insert a Word field at end of paragraph.

    Common field codes: PAGE, NUMPAGES, DATE, SEQ Figure, REF MyBookmark,
    STYLEREF Heading.

    Args:
        para_id: w14:paraId of the target paragraph.
        field_code: The field instruction text (e.g. "PAGE").
        cached_value: Optional display text cached in the document.
    """
    return _js(_require_doc().add_field(para_id, field_code, cached_value))


@mcp.tool()
def update_fields() -> str:
    """Mark all fields as dirty so Word recalculates on open."""
    return _js(_require_doc().update_fields())


@mcp.tool()
def list_fields() -> str:
    """List all fields in the document with their codes and cached values."""
    return _js(_require_doc().list_fields())


# ── Table of Contents ────────────────────────────────────────────────────────


@mcp.tool()
def generate_toc(max_level: int = 3, title: str = "Table of Contents") -> str:
    """Generate a Table of Contents from document headings."""
    doc = _require_doc()
    return _js(doc.generate_toc(max_level, title))


@mcp.tool()
def update_toc() -> str:
    """Regenerate ToC entries from current headings."""
    doc = _require_doc()
    return _js(doc.update_toc())


@mcp.tool()
def generate_list_of_figures() -> str:
    """Insert a List of Figures field (requires SEQ Figure captions)."""
    doc = _require_doc()
    return _js(doc.generate_list_of_figures())


@mcp.tool()
def generate_list_of_tables() -> str:
    """Insert a List of Tables field (requires SEQ Table captions)."""
    doc = _require_doc()
    return _js(doc.generate_list_of_tables())


# ── Content Controls ────────────────────────────────────────────────────────


@mcp.tool()
def add_content_control(
    para_id: str,
    tag: str,
    control_type: str,
    label: str = "",
    options: list[str] | None = None,
    default: str = "",
) -> str:
    """Wrap a paragraph in an SDT content control.

    Args:
        para_id: paraId of the paragraph to wrap.
        tag: Unique tag name for the control.
        control_type: One of text|checkbox|dropdown|date.
        label: Optional display label (w:alias).
        options: List of option strings for dropdown controls.
        default: Default display text (or date string for date controls).
    """
    doc = _require_doc()
    return _js(doc.add_content_control(para_id, tag, control_type, label, options, default))


@mcp.tool()
def get_content_controls() -> str:
    """List all SDT content controls in the document."""
    doc = _require_doc()
    return _js(doc.get_content_controls())


@mcp.tool()
def set_content_control_value(tag: str, value: str) -> str:
    """Update the value/text of a content control by its tag.

    Args:
        tag: Tag name of the control to update.
        value: New value. For checkbox: 'true'/'1' = checked, 'false'/'0' = unchecked.
    """
    doc = _require_doc()
    return _js(doc.set_content_control_value(tag, value))


@mcp.tool()
def lock_content_control(tag: str, lock: str = "sdtLocked") -> str:
    """Lock a content control to prevent editing.

    Args:
        tag: Tag name of the control to lock.
        lock: Lock type — sdtLocked|contentLocked|sdtContentLocked.
    """
    doc = _require_doc()
    return _js(doc.lock_content_control(tag, lock))


# ── Template Filling ─────────────────────────────────────────────────────────


@mcp.tool()
def fill_template(data: dict[str, str | list[str]], remove_empty: bool = False) -> str:
    """Fill SDT content controls from data dict. Keys match w:tag values.

    Args:
        data: Mapping of tag names to values. Use list[str] for repeating sections.
        remove_empty: If True, remove SDTs with no matching key in data.
    """
    doc = _require_doc()
    return _js(doc.fill_template(data, remove_empty))


@mcp.tool()
def list_template_fields() -> str:
    """List all SDT template fields (tag, label, type) in the document."""
    doc = _require_doc()
    return _js(doc.list_template_fields())


@mcp.tool()
def validate_template_data(data: dict) -> str:
    """Validate data dict covers all template fields. Returns missing and extra keys."""
    doc = _require_doc()
    return _js(doc.validate_template_data(data))


# ── Multilevel Lists ─────────────────────────────────────────────────────────


@mcp.tool()
def create_multilevel_list(name: str, levels: list[dict]) -> str:
    """Create a multilevel list in numbering.xml. Each level dict: {num_fmt, lvl_text, indent, hanging, style?}."""
    doc = _require_doc()
    return _js(doc.create_multilevel_list(name, levels))


@mcp.tool()
def restart_numbering(para_id: str, level: int = 0, start: int = 1) -> str:
    """Restart list numbering at a paragraph. Adds lvlOverride with startOverride."""
    doc = _require_doc()
    return _js(doc.restart_numbering(para_id, level, start))


@mcp.tool()
def suppress_numbering(para_id: str) -> str:
    """Remove list numbering from a paragraph by setting numId to 0."""
    doc = _require_doc()
    return _js(doc.suppress_numbering(para_id))


# ── Litigation Tools ────────────────────────────────────────────────────────


@mcp.tool()
def bates_number(prefix: str, start: int = 1, digits: int = 6, position: str = "footer-right") -> str:
    """Add Bates numbering stamp to document footer.

    Args:
        prefix: Bates prefix string (e.g. "ACME-").
        start: Starting Bates number.
        digits: Zero-padding width for the number.
        position: Hint for stamp position (currently "footer-right").
    """
    doc = _require_doc()
    return _js(doc.bates_number(prefix, start, digits, position))


@mcp.tool()
def redact_text(
    pattern: str | None = None,
    para_ids: list[str] | None = None,
    exact_text: str | None = None,
    reason: str = "",
) -> str:
    """True redaction: remove text and replace with black rectangle. Use exact_text or pattern.

    Args:
        pattern: Regex pattern to match run text.
        para_ids: Optional list of paragraph paraId values to limit scope.
        exact_text: Exact string to match against run text.
        reason: Reason for redaction (stored in log).
    """
    doc = _require_doc()
    return _js(doc.redact_text(pattern, para_ids, exact_text, reason))


@mcp.tool()
def generate_redaction_log(output_path: str = "") -> str:
    """Write a DOCX table of all redactions made this session.

    Args:
        output_path: Destination path. If empty, writes to a temp file.
    """
    doc = _require_doc()
    return _js(doc.generate_redaction_log(output_path))


@mcp.tool()
def generate_privilege_log(output_path: str = "") -> str:
    """Generate a privilege log DOCX from document metadata.

    Args:
        output_path: Destination path. If empty, writes to a temp file.
    """
    doc = _require_doc()
    return _js(doc.generate_privilege_log(output_path))


# ── Equations ───────────────────────────────────────────────────────────────


@mcp.tool()
def add_equation(para_id: str, latex: str) -> str:
    """Insert a LaTeX equation as OMML. Requires: pip install latex2mathml.

    Args:
        para_id: paraId of the paragraph after which the equation is inserted.
        latex: LaTeX source string (e.g. r"\\frac{1}{2}").
    """
    doc = _require_doc()
    return _js(doc.add_equation(para_id, latex))


@mcp.tool()
def get_equations() -> str:
    """Return all equations in the document as OMML XML strings."""
    doc = _require_doc()
    return _js(doc.get_equations())


# ── Charts ───────────────────────────────────────────────────────────────────


@mcp.tool()
def insert_bar_chart(
    para_id: str,
    title: str,
    series: list[dict],
    categories: list[str],
    width_cm: float = 14.0,
    height_cm: float = 9.0,
) -> str:
    """Insert a native bar chart (no Excel required).

    series: [{"name": str, "values": [float, ...]}]
    """
    doc = _require_doc()
    return _js(doc.insert_bar_chart(para_id, title, series, categories, width_cm, height_cm))


@mcp.tool()
def insert_line_chart(
    para_id: str,
    title: str,
    series: list[dict],
    categories: list[str],
    width_cm: float = 14.0,
    height_cm: float = 9.0,
) -> str:
    """Insert a native line chart.

    series: [{"name": str, "values": [float, ...]}]
    """
    doc = _require_doc()
    return _js(doc.insert_line_chart(para_id, title, series, categories, width_cm, height_cm))


@mcp.tool()
def insert_pie_chart(
    para_id: str,
    title: str,
    series: list[dict],
    categories: list[str],
) -> str:
    """Insert a native pie chart (single series, fixed 14x9 cm).

    series: [{"name": str, "values": [float, ...]}]
    """
    doc = _require_doc()
    return _js(doc.insert_pie_chart(para_id, title, series, categories))


@mcp.tool()
def update_chart_data(chart_id: str, series: list[dict]) -> str:
    """Replace data series in an existing chart by chart_id.

    series: [{"name": str, "values": [float, ...]}]
    """
    doc = _require_doc()
    return _js(doc.update_chart_data(chart_id, series))


@mcp.tool()
def merge_review_rounds(reviewer_paths: list[str], base_path: str | None = None) -> str:
    """Merge tracked changes from N reviewer copies into the open document."""
    return _js(_require_doc().merge_review_rounds(reviewer_paths, base_path))


@mcp.tool()
def compare_contracts(other_path: str, output_path: str = "", align_by: str = "heading") -> str:
    """Clause-aware diff between the open contract and another .docx file."""
    return _js(_require_doc().compare_contracts(other_path, output_path, align_by))


# ── Session log ──────────────────────────────────────────────────────────────


@mcp.tool()
def get_session_log() -> str:
    """Return all operations performed this session as replayable JSON."""
    return _js(_require_doc().get_session_log())


@mcp.tool()
def export_session_script(output_path: str) -> str:
    """Write session operations as a Python replay script.

    Returns: {"output_path": str, "operations": int}
    """
    return _js(_require_doc().export_session_script(output_path))


# ── Paragraph CRUD + border/shading ────────────────────────────────────────


@mcp.tool()
def insert_paragraph(
    after_para_id: str,
    text: str,
    style: str = "",
) -> str:
    """Insert a new paragraph after the paragraph with the given paraId.

    Args:
        after_para_id: paraId of the paragraph to insert after.
        text: Text content for the new paragraph.
        style: Optional paragraph style name (e.g., "Heading1", "Normal").
    """
    return _js(_require_doc().insert_paragraph(after_para_id, text, style=style or None))


@mcp.tool()
def update_paragraph(
    para_id: str,
    text: str = "",
    style: str = "",
) -> str:
    """Update the text and/or style of an existing paragraph.

    Args:
        para_id: paraId of the paragraph to update.
        text: New text content. Empty string leaves text unchanged.
        style: New paragraph style name. Empty string leaves style unchanged.
    """
    return _js(
        _require_doc().update_paragraph(
            para_id,
            text=text or None,
            style=style or None,
        )
    )


@mcp.tool()
def delete_paragraph(para_id: str) -> str:
    """Delete the paragraph with the given paraId.

    Args:
        para_id: paraId of the paragraph to remove.
    """
    return _js(_require_doc().delete_paragraph(para_id))


@mcp.tool()
def set_paragraph_border(
    para_id: str,
    sides: list[str],
    color: str = "000000",
    size: int = 4,
) -> str:
    """Set borders on one or more sides of a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        sides: List of sides: "top", "bottom", "left", "right", "between".
        color: Border color as 6-digit hex (default "000000").
        size: Border width in eighths of a point (default 4 = 0.5pt).
    """
    return _js(_require_doc().set_paragraph_border(para_id, sides, color=color, size=size))


@mcp.tool()
def set_paragraph_shading(
    para_id: str,
    fill_color: str,
    pattern: str = "clear",
) -> str:
    """Set background shading on a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        fill_color: Fill color as 6-digit hex (e.g., "FFFF00").
        pattern: Shading pattern (default "clear").
    """
    return _js(_require_doc().set_paragraph_shading(para_id, fill_color, pattern=pattern))


# ── Run-level formatting ───────────────────────────────────────────────────


@mcp.tool()
def get_runs(para_id: str) -> str:
    """Get all runs in a paragraph with their formatting properties.

    Args:
        para_id: paraId of the target paragraph.
    """
    return _js(_require_doc().get_runs(para_id))


@mcp.tool()
def set_run_font(para_id: str, run_idx: int, font_name: str) -> str:
    """Set the font of a specific run (zero-based index) in a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        font_name: Font name (e.g., "Arial", "Times New Roman").
    """
    return _js(_require_doc().set_run_font(para_id, run_idx, font_name))


@mcp.tool()
def set_run_color(para_id: str, run_idx: int, color: str) -> str:
    """Set the font color of a specific run in a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        color: Hex color without # (e.g., "FF0000").
    """
    return _js(_require_doc().set_run_color(para_id, run_idx, color))


@mcp.tool()
def set_run_size(para_id: str, run_idx: int, size_pt: float) -> str:
    """Set the font size of a specific run in a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        size_pt: Font size in points (e.g., 12.0).
    """
    return _js(_require_doc().set_run_size(para_id, run_idx, size_pt))


@mcp.tool()
def set_character_spacing(para_id: str, run_idx: int, spacing_pt: float) -> str:
    """Set character spacing (tracking) for a specific run in a paragraph.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        spacing_pt: Spacing in points (positive = expanded, negative = condensed).
    """
    return _js(_require_doc().set_character_spacing(para_id, run_idx, spacing_pt))


@mcp.tool()
def set_character_position(para_id: str, run_idx: int, position_pt: float) -> str:
    """Set vertical character position (raised/lowered) for a specific run.

    Args:
        para_id: paraId of the target paragraph.
        run_idx: Zero-based index of the run.
        position_pt: Offset in points (positive = raised, negative = lowered).
    """
    return _js(_require_doc().set_character_position(para_id, run_idx, position_pt))


@mcp.tool()
def export_markdown(output_path: str = "") -> str:
    """Export the open document as Markdown.

    Converts the document body (headings, bold/italic runs, lists, tables,
    plain paragraphs) to GitHub-Flavoured Markdown and writes it to a file.

    Args:
        output_path: Destination path for the .md file.
            Defaults to <workdir>/export.md when empty.

    Returns: {"output_path": str, "paragraphs": int, "tables": int}
    """
    return _js(_require_doc().export_markdown(output_path))




@mcp.tool()
def get_theme_colors() -> str:
    """Return the named color slots from word/theme/theme1.xml.

    Returns a dict mapping slot names (dk1, lt1, accent1, ...) to 6-digit hex strings.
    Returns an empty dict if the document has no theme file.
    """
    return _js(_require_doc().get_theme_colors())


@mcp.tool()
def set_theme_color(slot: str, hex_color: str) -> str:
    """Update a named color slot in the document theme.

    Args:
        slot: One of dk1, lt1, dk2, lt2, accent1-accent6, hlink, folHlink.
        hex_color: 6-character hex string without # (e.g. "FF0000").
    """
    return _js(_require_doc().set_theme_color(slot, hex_color))


@mcp.tool()
def insert_caption(after_para_id: str, text: str, label: str = "Figure") -> str:
    """Insert a caption paragraph after the specified paragraph.

    Args:
        after_para_id: paraId of the paragraph to insert after.
        text: Caption description text (without label/number prefix).
        label: Caption label, e.g. "Figure" or "Table" (default "Figure").
    """
    return _js(_require_doc().insert_caption(after_para_id, text, label=label))
# ── Entry point ─────────────────────────────────────────────────────────────


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()

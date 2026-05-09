# Competitive Analysis: Word/DOCX MCP Servers

**Date**: 2026-03-23
**Our Server**: docx-mcp / docx-mcp-server (PyPI) — 43 tools, Python, cross-platform

---

## Landscape Overview

~18 Word/DOCX MCP servers exist. The market is fragmented: most are thin CRUD wrappers around python-docx. Only 3-4 (including ours) do deep OOXML-level work.

**Key finding**: The #1 competitor by stars (Office-Word-MCP-Server, 1,800 stars) was **archived on March 3, 2026**, creating a vacuum.

---

## Tier 1 Competitors

### 1. che-word-mcp (PsychQuant/kiki830621) — BIGGEST THREAT
| Attribute | Value |
|---|---|
| Language | Swift (native binary) |
| Stars | 3 GitHub |
| Tools | **146** (v1.17.0, 2026-03-11) |
| URL | https://github.com/kiki830621/che-word-mcp |
| Platform | macOS only (Swift requirement) |
| Velocity | v1.0 to v1.17 in ~2 months |

**Tool categories (146 tools):**
- Document Management (6): create, open, save, close, list_open, get_info
- Content Operations (6): get_text, get_paragraphs, insert/update/delete_paragraph, replace_text
- Formatting (3): format_text, set_paragraph_format, apply_style
- Tables (6): insert_table, get_tables, update_cell, delete_table, merge_cells, set_table_style
- Images (6): insert_image, insert_floating_image, update/delete/list_images, set_image_style
- Headers & Footers (5): add/update_header, add/update_footer, insert_page_number
- Style Management (4): list/create/update/delete_styles
- Page Setup (5): page_size, margins, orientation, page_break, section_break
- Lists (3): bullet_list, numbered_list, set_list_level
- Footnotes & Endnotes (4): insert/delete_footnote, insert/delete_endnote
- Comments & Revisions (10): full comment CRUD + resolve_comment, enable/disable_track_changes, accept/reject_revision
- Hyperlinks & Bookmarks (6): insert/update/delete_hyperlink, internal_link, insert/delete_bookmark
- Advanced (9): **insert_toc**, text_field, **checkbox**, **dropdown**, **equation**, paragraph_border, paragraph_shading, character_spacing, text_effect
- Field Codes (7): **IF field**, **calculation field**, date_field, page_field, **merge_field**, **sequence_field**, **content_control (SDT)**
- Export (2): export_text, export_markdown
- Plus ~64 tools added in v1.3-v1.17: **compare_documents**, search_text, autosave, finalize_document, session_state, EMF-to-PNG, dual-mode access

### 2. Office-Word-MCP-Server (GongRzhe) — ARCHIVED
| Attribute | Value |
|---|---|
| Language | Python |
| Stars | 1,800 GitHub |
| Downloads | 172,500 PyPI |
| Tools | ~27 |
| Status | **ARCHIVED March 3, 2026** |
| URL | https://github.com/GongRzhe/Office-Word-MCP-Server |

**Notable tools:** convert_to_pdf, add_digital_signature, verify_document_integrity, add_password_protection, restricted_editing (with editable sections), create_custom_style, format_table, delete_paragraph, copy_document, merge_documents

### 3. hongkongkiwi/docx-mcp (Rust)
| Attribute | Value |
|---|---|
| Language | Rust |
| Stars | 20 GitHub |
| Tools | ~30+ |
| URL | https://github.com/hongkongkiwi/docx-mcp |

**Notable unique features:**
- Built-in PDF generation (no LibreOffice needed)
- Convert to images (PNG/JPG per page)
- Export to Markdown, HTML
- 7 professional templates (business letter, resume, report, invoice, contract, memo, newsletter)
- Document structure analysis, formatting consistency analysis
- Word count with reading time
- Sandbox mode, readonly mode, tool whitelisting/blacklisting, max-size limits, no-network mode

### 4. MCP-Doc (MeterLong)
| Attribute | Value |
|---|---|
| Language | Python (FastMCP) |
| Stars | 170 GitHub |
| Tools | ~15-20 |
| URL | https://github.com/MeterLong/MCP-Doc |

**Notable:** Merge/split table cells, find-and-replace, style preservation on edit

---

## Tier 2 Competitors

| Server | Lang | Stars | Notable Unique Features |
|---|---|---|---|
| **MCP-OPENAPI-DOCX** (Fu-Jie) | Python/FastAPI | 3 | RESTful+Swagger, version control, template mgmt, export PDF/HTML/MD, encryption, Celery tasks, K8s |
| **Rookie0x80/docx-mcp** | Python | — | Batch cell ops, cell merge+unmerge, table header search, range queries, data/style separation |
| **word-mcp** (PyPI) | Python | — | COM automation (actual Word control), Windows+macOS, multiple instances |
| **officemcp** (PyPI) | Python | — | Full Office suite COM (Word/Excel/PPT/Access/OneNote/Visio/Project/WPS) |
| **mario-andreschak** | TypeScript | — | COM Interop via winax, line spacing, table auto-format, header/footer types |
| **@docx-mcp/docx-mcp** (npm) | Node.js | — | JSON schema-driven, code blocks (180+ langs), blockquotes, text boxes |
| **Microsoft Work IQ** | Cloud | — | OneDrive/SharePoint, enterprise governance, M365 ecosystem (requires Copilot license) |
| **mcp-md-pdf** | Python | — | Markdown-to-Word/PDF, .dotx template support, batch processing |
| **aiexplorations/docx-mcp** | Python | — | 30+ tools, apply lists to existing paragraphs |
| **dvejsada/mcp-ms-office** | Python | — | {{placeholder}} templates, multi-format (docx/pptx/eml/xlsx), cloud upload (S3/GCS/Azure) |
| **consistent-docx-mcp** | Node.js | — | Format-preserving Markdown-to-DOCX |

---

## GAP ANALYSIS: What Competitors Have That We Don't

### HIGH PRIORITY — Table Stakes Gaps

| Feature | Competitors Who Have It | Est. Difficulty |
|---|---|---|
| **Find and replace** | che-word, GongRzhe, MeterLong, hongkongkiwi, Rookie0x80 | Easy |
| **Delete paragraph** | che-word, GongRzhe | Easy |
| **Delete table** | che-word | Easy |
| **Merge/split table cells** | che-word, MeterLong, Rookie0x80, mario | Medium |
| **Export to Markdown** | che-word, hongkongkiwi | Medium |
| **Convert to PDF** | GongRzhe, hongkongkiwi, Fu-Jie, mcp-md-pdf | Medium-Hard |
| **Table of contents (TOC)** | che-word, hongkongkiwi | Medium |
| **Bookmarks** (insert/delete) | che-word | Medium |
| **Hyperlinks** (insert/update/delete) | che-word | Medium |
| **Custom style CRUD** | che-word, GongRzhe | Medium |
| **Page size/orientation** | che-word | Easy |

### MEDIUM PRIORITY — Differentiation

| Feature | Who Has It | Est. Difficulty |
|---|---|---|
| **Form fields** (text, checkbox, dropdown) | che-word | Medium |
| **Math equations** | che-word | Hard |
| **Floating images with text wrap** | che-word | Medium |
| **Image borders/effects** | che-word | Medium |
| **Content controls (SDT)** | che-word | Medium |
| **Field codes** (IF, calculation, date, merge, sequence) | che-word | Medium-Hard |
| **Compare documents** (diff) | che-word, Fu-Jie | Hard |
| **Paragraph borders/shading** | che-word | Easy |
| **Character spacing** | che-word | Easy |
| **Word count / reading time** | hongkongkiwi | Easy |
| **Resolve comment** | che-word | Easy |
| **Delete footnote/endnote** | che-word | Easy |
| **Document templates** (built-in) | hongkongkiwi, dvejsada | Medium |
| **Search text across document** | che-word | Easy |

### LOW PRIORITY — Niche / Architecture

| Feature | Who Has It | Est. Difficulty |
|---|---|---|
| Digital signatures | GongRzhe | Hard |
| Document encryption | Fu-Jie, GongRzhe | Hard |
| Version control/history | Fu-Jie | Hard |
| COM automation (live Word) | word-mcp, officemcp, mario | N/A (different paradigm) |
| Cloud upload (S3/GCS/Azure) | dvejsada | Medium |
| Sandbox/readonly modes | hongkongkiwi | Medium |
| Code blocks w/ syntax highlighting | lihongjie | Medium |
| Text boxes | lihongjie | Medium |
| Blockquotes | lihongjie | Easy |
| Convert to images | hongkongkiwi | Hard |
| OneDrive/SharePoint integration | Microsoft | N/A (cloud-only) |

---

## Our Competitive Moat (Unique to Us)

These features are unique to docx-mcp or shared with at most one competitor:

1. **Track changes at OOXML level** (insert_text, delete_text with proper revision markup) — only us + che-word
2. **Paragraph ID validation** (validate_paraids) — unique to us
3. **Document auditing** (audit_document) — unique to us
4. **Watermark removal** (remove_watermark) — unique to us
5. **Cross-references** (add_cross_reference) — unique to us
6. **OOXML structural validation** — unique to us
7. **Comment threading** (reply_to_comment) — us + che-word
8. **Endnote support** (get/add/validate) — us + che-word
9. **Document protection with proper OOXML** — proper w:documentProtection vs basic password

---

## Recommended Roadmap (Priority Order)

### Phase 1 — Close Table-Stakes Gaps
1. Find and replace (5+ competitors have this)
2. Delete paragraph
3. Delete table / delete table row
4. Merge table cells
5. Page size and orientation control

### Phase 2 — Feature Parity with che-word-mcp
6. Bookmarks (insert/delete)
7. Hyperlinks (insert/update/delete, internal links)
8. Table of contents (insert_toc)
9. Custom style CRUD (create/update/delete)
10. Export to Markdown

### Phase 3 — Differentiation
11. Form fields (checkbox, dropdown, text field)
12. Field codes (IF, calculation, date, page number, merge field, sequence)
13. Content controls (SDT)
14. Compare documents (diff)
15. Convert to PDF

### Phase 4 — Polish
16. Floating images with text wrap
17. Word count / document statistics
18. Paragraph borders and shading
19. Character spacing and text effects
20. Resolve comment (mark as done)

---

## Sources

- [che-word-mcp (GitHub)](https://github.com/kiki830621/che-word-mcp)
- [Office-Word-MCP-Server (GitHub, ARCHIVED)](https://github.com/GongRzhe/Office-Word-MCP-Server)
- [hongkongkiwi/docx-mcp (GitHub)](https://github.com/hongkongkiwi/docx-mcp)
- [MCP-Doc (GitHub)](https://github.com/MeterLong/MCP-Doc)
- [MCP-OPENAPI-DOCX (GitHub)](https://github.com/Fu-Jie/MCP-OPENAPI-DOCX)
- [Rookie0x80/docx-mcp (GitHub)](https://github.com/Rookie0x80/docx-mcp)
- [word-mcp (PyPI)](https://pypi.org/project/word-mcp/)
- [officemcp (PyPI)](https://pypi.org/project/officemcp/)
- [mario-andreschak/mcp-msoffice-interop-word (GitHub)](https://github.com/mario-andreschak/mcp-msoffice-interop-word)
- [@docx-mcp/docx-mcp (npm)](https://www.npmjs.com/package/@docx-mcp/docx-mcp)
- [Microsoft Work IQ Word (Microsoft Learn)](https://learn.microsoft.com/en-us/microsoft-agent-365/mcp-server-reference/word)
- [mcp-md-pdf (PyPI)](https://pypi.org/project/mcp-md-pdf/)
- [aiexplorations/docx-mcp (GitHub)](https://github.com/aiexplorations/docx-mcp)
- [dvejsada/mcp-ms-office-documents (GitHub)](https://github.com/dvejsada/mcp-ms-office-documents)
- [Glama MCP Directory](https://glama.ai/mcp/servers)
- [Smithery MCP Directory](https://smithery.ai)
- [PulseMCP Directory](https://www.pulsemcp.com)

# tests/fixtures/

Real Word documents sourced externally for correctness validation.
These are the doer-checker gate — synthetic fixtures alone are insufficient.

- `real_contract.docx` — 50-page multi-section document with TOC, heading styles, headers, footers, and page numbers (1924 paragraphs). Source: https://sample-files.com/downloads/documents/docx/sample-files.com-large-document.docx
- `real_tracked_changes.docx` — 2-page document with tracked insertions (`w:ins`) and deletions (`w:del`), plus reviewer comments. Source: https://sample-files.com/downloads/documents/docx/sample-files.com-tracked-changes.docx
- `real_hyperlinks_footnotes.docx` — TEI Publisher test document with hyperlinks in the document body AND in footnotes (footnotes.xml contains 2 hyperlinks, document.xml contains 1). Exercises the footnotes.xml.rels vs document.xml.rels ID collision edge case. Source: https://github.com/eeditiones/tei-publisher-lib/files/7334303/test.docx (attached to issue #5)

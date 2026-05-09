# Design Spec Review: Feature Expansion (2026-03-23)

**Reviewer:** Claude Code (Senior Code Reviewer)
**Spec:** `2026-03-23-feature-expansion-design.md`
**Codebase state:** 18 tools, 919-line `document.py`, 100% coverage, Python 3.10+/FastMCP/lxml

---

## Verdict: Solid foundation with 5 critical fixes needed before implementation

The spec is well-structured with clear phase boundaries, consistent naming, and a sound TDD approach. The mixin decomposition is the right call for managing growth from 18 to 43 tools. However, several OOXML technical details are incorrect and will produce invalid documents if implemented as-written.

---

## CRITICAL (Must Fix Before Implementation)

### C1. `add_section_break` placement is wrong

The spec says: "Insert `w:sectPr` with `w:type` **after** target paragraph."

In OOXML, section breaks go **inside** the paragraph's `w:pPr`, not after it. The body-level `w:sectPr` (direct child of `w:body`) defines only the final section. All intermediate section breaks are `w:sectPr` children of the last paragraph's `w:pPr` in that section.

**Fix:** Insert `w:sectPr` as child of target paragraph's `w:pPr` (create `w:pPr` if absent). Document that the target paragraph becomes the last paragraph of the new section.

### C2. `merge_documents` needs a sub-specification

One-line behavior for the most complex operation in the entire expansion is insufficient. This tool must handle:
- paraId collision remapping across all parts (body, footnotes, endnotes, comments, headers, footers)
- rId remapping with namespace isolation
- Style merging (same name, different definition)
- `[Content_Types].xml` merging
- `numbering.xml` abstractNum ID collision
- Header/footer relationship remapping
- Section boundary between source and target documents

Recommendation: Write a dedicated sub-spec for `merge_documents`. Consider moving it to Phase 7 (after all other features stabilize), since it depends on every feature area.

### C3. `set_formatting` track changes mechanism is described incorrectly

The spec says "wraps in `w:rPrChange`." The correct OOXML pattern:
1. Clone existing `w:rPr` as the child of a new `w:rPrChange` element
2. Apply new formatting properties to the actual `w:rPr`
3. Insert `w:rPrChange` (with `w:id`, `w:author`, `w:date` attributes) **inside** the modified `w:rPr`

`w:rPrChange` is a child of `w:rPr`, not a wrapper around it.

### C4. `accept_changes`/`reject_changes` are missing move tracking

The spec only handles `w:ins`, `w:del`, and `w:rPrChange`. Real Word documents also contain `w:moveTo`/`w:moveFrom` (paragraph/run moves). At minimum, these must be handled. Table-level changes (`w:cellIns`, `w:cellDel`, `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`) can be documented as out of scope for now.

### C5. `delete_table_row` approach is invalid OOXML

The spec says "Wrap entire `w:tr` content in `w:del`." In OOXML, `w:del` cannot be a parent of `w:tr`. Word tracks deleted rows by wrapping each paragraph's runs within each cell in `w:del` elements (and optionally marking the row with `w:trPr/w:del`). The entire `w:tr` structure must remain intact.

---

## IMPORTANT (Should Fix)

### I1. Phase 0 refactor: path references are wrong

The spec references `src/docx_mcp/document/`. The actual layout is flat: `docx_mcp/` at repo root (no `src/` directory). Hatch config confirms: `packages = ["docx_mcp"]`. All refactor paths must reference `docx_mcp/document/`.

### I2. Mixin MRO needs explicit documentation

With 18 mixins, the class definition order determines Python's Method Resolution Order via C3 linearization. The spec places `BaseMixin` first. This works only if mixins do not call `super().__init__()`. Document the constraint: only `BaseMixin` defines `__init__`; all other mixins must NOT define `__init__`.

### I3. `get_images` requires missing namespace

`width_emu`/`height_emu` come from `wp:extent` (wordprocessingDrawing namespace). The codebase defines `A` (drawingml/main) but not `WP` (`http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing`). This namespace must be added in Phase 0 or Phase 1.

### I4. `add_table` missing required `w:tblPr`

A valid Word table needs `w:tblPr` with at minimum `w:tblW` (width). Without it, Word will auto-repair the file on open, which corrupts the track-changes intent.

### I5. `add_list` missing numbering part bootstrapping

If `numbering.xml` does not exist (common in simple documents), it must be created AND registered in `[Content_Types].xml` AND linked via `word/_rels/document.xml.rels`. This three-step bootstrap is not documented.

### I6. `set_document_protection` password hash algorithm unspecified

Word uses a specific algorithm: SHA-512 with base64-encoded random salt, iterative hash with configurable `spinCount` (typically 100000), per ECMA-376 Part 4 section 2.15.1.28. Implementing this incorrectly means Word rejects the password.

### I7. `edit_header_footer` has a hidden dependency on section enumeration

Section enumeration (mapping `section_index` to the correct `w:sectPr`) requires the same logic needed by Phase 5's section tools. Either:
- Extract section enumeration as shared infrastructure in Phase 0/1, or
- Move `edit_header_footer` to Phase 5 (after section tools exist)

### I8. Audit extension creates implicit mixin coupling

`audit_document()` in `ValidationMixin` will call `self.tables()` (Phase 3), `self.images()` etc. This works via Python's MRO but means `ValidationMixin` silently depends on other mixins being composed. Document this as a design constraint, or use `hasattr(self, 'tables')` guards.

---

## SUGGESTIONS

### S1. Standardize error response format for new tools
Existing tools raise `ValueError`/`RuntimeError`/`FileNotFoundError`. Document that new tools follow the same pattern for consistency.

### S2. Define bookmark naming convention for `add_cross_reference`
Bookmark names must be unique, no spaces, max 40 chars. Suggest: `_docxmcp_{paraId}`.

### S3. Clarify test fixture growth strategy
"conftest.py fixture grows per phase" risks unintended test interactions. Consider per-phase fixture builders that extend the base.

### S4. Add `author` parameter defaults for Phase 2 tools
`accept_changes`/`reject_changes` take an `author` filter but the default behavior is not specified.

### S5. Consider a `get_raw_xml` debug tool
Trivial to implement, invaluable for debugging OOXML issues during development.

### S6. Tool count precision
Spec says "~45". Actual count: 18 existing + 25 new = 43 tools. State precisely.

---

## Phase Dependency Issues

| Issue | Recommendation |
|-------|---------------|
| `merge_documents` (Phase 5) depends on ALL feature areas | Move to Phase 7 or make it the absolute last tool implemented |
| `edit_header_footer` (Phase 4) needs section enumeration | Move after Phase 5, or extract section enumeration to Phase 1 |
| `set_formatting` (Phase 2) needs run-splitting at arbitrary boundaries | Document as new capability beyond existing `delete_text` run splitting |
| Audit growth creates cross-mixin calls | Use `hasattr` guards or document mixin ordering requirements |

---

## What Was Done Well

- Mixin decomposition is the right architecture for 43 tools
- TDD with `fail_under = 100` is excellent discipline
- Phase 0 pure-refactor-then-verify approach is safe
- Import compatibility preservation (`document/__init__.py` re-exports) is smart
- No new dependencies (lxml + stdlib only) keeps the project lean
- Separate test files per phase is good organization
- Version bump plan tied to phase pairs makes sense
- Tool naming conventions are consistent throughout

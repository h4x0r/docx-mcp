"""Tests for mike integration: get_body_text, replace_text return, hyperlink, NBSP."""
from __future__ import annotations
import json
import pytest
from docx_mcp import server
from docx_mcp.document.tracks import _flatten_para
from lxml import etree


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"


def test_flatten_para_includes_hyperlink_text(mike_corpus_docx):
    """_flatten_para must include text inside w:hyperlink runs."""
    server.open_document(str(mike_corpus_docx))
    doc = server._doc._require("word/document.xml")
    # Find paragraph 00000101 which contains a w:hyperlink with "thirty calendar days"
    para = None
    for p in doc.iter(f"{W}p"):
        if p.get(f"{W14}paraId") == "00000101":
            para = p
            break
    assert para is not None, "paragraph 00000101 not found"
    slots = _flatten_para(para)
    text = "".join(s.char for s in slots)
    assert "thirty calendar days" in text, (
        f"hyperlink text not found in _flatten_para output; got: {text!r}"
    )


def test_get_body_text_returns_body_string(mike_corpus_docx):
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.get_body_text())
    assert "body" in result
    assert isinstance(result["body"], str)
    assert len(result["body"]) > 0


def test_get_body_text_includes_hyperlink_text(mike_corpus_docx):
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.get_body_text())
    # Hyperlink text must appear in body (not invisible)
    assert "thirty calendar days" in result["body"]


def test_get_body_text_accepted_view_excludes_deleted(mike_corpus_docx):
    """Accepted view: w:del text must NOT appear, w:ins text MUST appear."""
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.get_body_text())
    assert "one hundred" not in result["body"]   # inside w:del → excluded
    assert "two hundred" in result["body"]         # inside w:ins → included


def test_get_body_text_returns_footnotes(mike_corpus_docx):
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.get_body_text())
    assert "footnotes" in result
    # Footnote id=1 in _FOOTNOTES_XML contains "See appendix A for supporting evidence."
    # Reserved ids -1 (separator) and 0 (continuation) must be excluded.
    assert "See appendix A" in result["footnotes"]


def test_get_body_text_real_contract(tmp_path):
    """Smoke test against real externally-sourced document."""
    import os
    fixture = os.path.join(os.path.dirname(__file__), "fixtures", "real_contract.docx")
    if not os.path.exists(fixture):
        pytest.skip("real fixture not present")
    server.open_document(fixture)
    result = json.loads(server.get_body_text())
    assert len(result["body"]) > 100


def test_replace_text_returns_del_id_and_ins_id(mike_corpus_docx):
    """replace_text must return both del_id and ins_id for Accept/Reject card UI."""
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.replace_text(
        para_id="00000108",
        find="initial deposit",
        replace="upfront payment",
    ))
    assert "del_id" in result, "must include del_id"
    assert "ins_id" in result, "must include ins_id"
    assert result["del_id"] is not None
    assert result["ins_id"] is not None
    assert result["del_id"] != result["ins_id"]
    assert result["type"] == "replacement"


def test_replace_text_pure_deletion_has_no_ins_id(mike_corpus_docx):
    """Pure deletion (empty replace) must have ins_id=None."""
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.replace_text(
        para_id="00000109",
        find="ongoing ",
        replace="",
    ))
    assert result["del_id"] is not None
    assert result["ins_id"] is None


# ── Narrow no-break space normalization ──────────────────────────────────────

def test_replace_finds_text_with_narrow_nbsp(mike_corpus_docx):
    """replace_text must match '$5 000' when find uses regular space '$5 000'.
    Paragraph 00000102 contains U+202F narrow no-break space between '5' and '000'.
    """
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.replace_text(
        para_id="00000102",
        find="$5 000",          # normal space in find
        replace="$6 000",
    ))
    assert result["del_id"] is not None


# ── Smart quote normalization ─────────────────────────────────────────────────

def test_replace_finds_smart_quotes_with_ascii(mike_corpus_docx):
    """replace_text must match curly "Effective Date" when find uses ASCII double quotes."""
    server.open_document(str(mike_corpus_docx))
    result = json.loads(server.replace_text(
        para_id="00000103",
        find='"Effective Date"',   # ASCII double quotes
        replace='"Execution Date"',
    ))
    assert result["del_id"] is not None


# ── Context disambiguation ─────────────────────────────────────────────────────

def test_search_text_returns_both_ambiguous_matches(mike_corpus_docx):
    """'twelve months' appears in two paragraphs — search must return both."""
    server.open_document(str(mike_corpus_docx))
    # search_text returns a list of dicts with 'paraId' (camelCase) key
    results = json.loads(server.search_text(query="twelve months"))
    para_ids = [m["paraId"] for m in results]
    assert "00000104" in para_ids
    assert "00000105" in para_ids


def test_replace_with_context_disambiguates(mike_corpus_docx, tmp_path):
    """replace_text with context_before disambiguates between two identical phrases."""
    server.open_document(str(mike_corpus_docx))
    # Target only Section 1
    result = json.loads(server.replace_text(
        para_id="00000104",
        find="twelve months",
        replace="eighteen months",
        context_before="term shall be ",
    ))
    assert result["del_id"] is not None
    # Section 2 paragraph must be unchanged — save and re-read to verify
    out = tmp_path / "disambig_out.docx"
    server.save_document(str(out))
    server.open_document(str(out))
    body = json.loads(server.get_body_text())["body"]
    assert "twelve months" in body   # section 2 still has it


# ── Accepted-view with pre-existing tracked changes ───────────────────────────

def test_accepted_view_shows_ins_hides_del(mike_corpus_docx):
    server.open_document(str(mike_corpus_docx))
    body = json.loads(server.get_body_text())["body"]
    assert "two hundred" in body
    assert "one hundred" not in body


def test_replace_targets_accepted_view_text(mike_corpus_docx):
    """replace_text on text inside a pre-existing w:ins raises ValueError.
    Paragraph 00000107 has pre-existing w:del('one hundred') and w:ins('two hundred').
    The implementation blocks deletion of text inside an existing w:ins (must accept/reject first).
    """
    server.open_document(str(mike_corpus_docx))
    with pytest.raises(ValueError, match="existing w:ins"):
        server.replace_text(
            para_id="00000107",
            find="two hundred",
            replace="three hundred",
        )


# ── Batch edits ───────────────────────────────────────────────────────────────

def test_batch_three_replacements_in_one_session(mike_corpus_docx, tmp_path):
    """Three replace_text calls in one open/save session all persist."""
    server.open_document(str(mike_corpus_docx))
    server.replace_text(para_id="00000108", find="initial deposit", replace="upfront payment")
    server.replace_text(para_id="00000109", find="maintenance fee", replace="service charge")
    server.replace_text(para_id="0000010A", find="termination penalty", replace="exit fee")
    out = tmp_path / "batch_out.docx"
    server.save_document(str(out))
    # Re-open and verify all three changes are tracked
    server.open_document(str(out))
    # get_tracked_changes returns a list of dicts with keys: type, change_id, author, date, para_id, text
    changes = json.loads(server.get_tracked_changes())
    texts = [c["text"] for c in changes]
    # Either the deleted or inserted text should appear in tracked changes
    assert any("initial deposit" in t or "upfront payment" in t for t in texts)
    assert any("maintenance fee" in t or "service charge" in t for t in texts)
    assert any("termination penalty" in t or "exit fee" in t for t in texts)


# ── Accept/Reject roundtrip ───────────────────────────────────────────────────

def test_accept_change_from_replace_text(mike_corpus_docx, tmp_path):
    server.open_document(str(mike_corpus_docx))
    r = json.loads(server.replace_text(para_id="00000108", find="initial", replace="upfront"))
    del_id = r["del_id"]
    ins_id = r["ins_id"]
    # Accept del (accept the deletion — remove the old text markup)
    server.accept_change(del_id)
    # Accept ins (accept the insertion — make the new text permanent)
    server.accept_change(ins_id)
    out = tmp_path / "accepted.docx"
    server.save_document(str(out))
    server.open_document(str(out))
    body = json.loads(server.get_body_text())["body"]
    assert "upfront" in body
    assert "initial" not in body


# ── Real fixture smoke tests ──────────────────────────────────────────────────

@pytest.mark.parametrize("fixture_name", [
    "real_contract.docx",
    "real_tracked_changes.docx",
    "real_hyperlinks_footnotes.docx",
])
def test_get_body_text_real_fixtures(fixture_name, tmp_path):
    import os
    fixture = os.path.join(os.path.dirname(__file__), "fixtures", fixture_name)
    if not os.path.exists(fixture):
        pytest.skip(f"fixture {fixture_name} not present")
    server.open_document(fixture)
    result = json.loads(server.get_body_text())
    assert len(result["body"]) > 50, "real doc must have substantial body text"

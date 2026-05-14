"""Smoke tests for the atheris fuzzer harness — runs without atheris installed."""
from __future__ import annotations

import io
import zipfile


def _minimal_docx_bytes() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("_rels/.rels", "<Relationships/>")
        zf.writestr("word/document.xml", "<document/>")
    return buf.getvalue()


def test_fuzz_harness_handles_valid_input():
    from tests.fuzz.fuzz_open import TestOneInput
    TestOneInput(_minimal_docx_bytes())


def test_fuzz_harness_handles_garbage():
    from tests.fuzz.fuzz_open import TestOneInput
    TestOneInput(b"not a zip file at all")


def test_fuzz_harness_handles_empty():
    from tests.fuzz.fuzz_open import TestOneInput
    TestOneInput(b"")


def test_fuzz_harness_handles_truncated_zip():
    from tests.fuzz.fuzz_open import TestOneInput
    TestOneInput(b"PK\x03\x04\x14\x00")


def test_fuzz_harness_handles_zipslip_payload():
    """ZipSlip payload must not raise an unhandled exception."""
    import io
    from tests.fuzz.fuzz_open import TestOneInput
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("../../evil.txt", "pwned")
    TestOneInput(buf.getvalue())


def test_fuzz_harness_handles_xxe_payload():
    """XXE payload must not raise an unhandled exception."""
    import io
    from tests.fuzz.fuzz_open import TestOneInput
    evil_xml = (
        '<?xml version="1.0"?>'
        '<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>'
        "<document>&xxe;</document>"
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("word/document.xml", evil_xml)
    TestOneInput(buf.getvalue())

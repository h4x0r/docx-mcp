"""Atheris byte-level fuzzer for DocxDocument.open().

Usage:
    pip install atheris
    python tests/fuzz/fuzz_open.py -runs=10000
    python tests/fuzz/fuzz_open.py tests/fuzz/corpus/  # use corpus directory

Invariant: any input must either succeed or raise one of the expected
exceptions. An unhandled exception = crash = fuzzer finds a bug.
"""

from __future__ import annotations

import contextlib
import os
import sys
import tempfile
from pathlib import Path

try:
    import atheris

    ATHERIS_AVAILABLE = True
except ImportError:
    ATHERIS_AVAILABLE = False

from docx_mcp.document import DocxDocument
from docx_mcp.document.errors import DocxMcpError

_EXPECTED = (
    DocxMcpError,
    FileNotFoundError,
    ValueError,
    OSError,
    PermissionError,
)


def TestOneInput(data: bytes) -> None:
    """Feed raw bytes as a .docx file to DocxDocument.open()."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        f.write(data)
        tmp_path = f.name
    doc = None
    try:
        doc = DocxDocument(tmp_path)
        doc.open()
    except _EXPECTED:
        pass
    except Exception:
        raise
    finally:
        if doc is not None:
            with contextlib.suppress(Exception):
                doc.close()
        with contextlib.suppress(OSError):
            os.unlink(tmp_path)


if __name__ == "__main__":
    if not ATHERIS_AVAILABLE:
        print("atheris not installed. Run: pip install atheris", file=sys.stderr)
        sys.exit(1)

    corpus_dir = Path(__file__).parent / "corpus"
    atheris.Setup(sys.argv, TestOneInput)
    atheris.Fuzz()

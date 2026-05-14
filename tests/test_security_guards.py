"""Security guard tests — input validation."""
from __future__ import annotations

import pytest

from docx_mcp.document.errors import ErrCode


def test_malformed_input_code_exists():
    assert ErrCode.MALFORMED_INPUT == "MALFORMED_INPUT"


def test_unsafe_path_code_exists():
    assert ErrCode.UNSAFE_PATH == "UNSAFE_PATH"

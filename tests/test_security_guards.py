"""Security guard tests — input validation."""
from __future__ import annotations

import pytest

from docx_mcp.document.errors import ErrCode


def test_malformed_input_code_exists():
    assert ErrCode.MALFORMED_INPUT == "MALFORMED_INPUT"


def test_unsafe_path_code_exists():
    assert ErrCode.UNSAFE_PATH == "UNSAFE_PATH"


import re
import sys
from pathlib import Path

from hypothesis import given, settings
from hypothesis import strategies as st

from docx_mcp.document.guards import InputGuard


# ── para_id ──────────────────────────────────────────────────────────────────

def test_para_id_valid():
    assert InputGuard.para_id("1A2B3C4D") == "1A2B3C4D"


def test_para_id_valid_lowercase():
    assert InputGuard.para_id("abcdef01") == "abcdef01"


def test_para_id_empty_raises():
    with pytest.raises(ValueError, match="para_id"):
        InputGuard.para_id("")


def test_para_id_too_long_raises():
    with pytest.raises(ValueError, match="para_id"):
        InputGuard.para_id("AABBCCDD11")  # 10 chars — over limit


def test_para_id_non_hex_raises():
    with pytest.raises(ValueError, match="para_id"):
        InputGuard.para_id("GGGGGGGG")


def test_para_id_null_byte_raises():
    with pytest.raises(ValueError, match="para_id"):
        InputGuard.para_id("1A2B\x003C")


@given(st.text(min_size=0, max_size=200))
@settings(max_examples=500)
def test_para_id_arbitrary_input_never_crashes(s):
    try:
        InputGuard.para_id(s)
    except ValueError:
        pass  # expected for invalid inputs


# ── output_path ───────────────────────────────────────────────────────────────

def test_output_path_valid(tmp_path):
    p = str(tmp_path / "out.docx")
    result = InputGuard.output_path(p)
    assert isinstance(result, Path)
    assert result.suffix == ".docx"


def test_output_path_traversal_raises(tmp_path):
    with pytest.raises(ValueError, match="traversal"):
        InputGuard.output_path("../../etc/passwd.docx")


def test_output_path_wrong_suffix_raises(tmp_path):
    with pytest.raises(ValueError, match="suffix"):
        InputGuard.output_path(str(tmp_path / "out.pdf"))


def test_output_path_absolute_traversal_raises():
    with pytest.raises(ValueError):
        InputGuard.output_path("/etc/passwd")


@given(st.text(min_size=1, max_size=300))
@settings(max_examples=300)
def test_output_path_arbitrary_never_crashes(s):
    try:
        InputGuard.output_path(s)
    except (ValueError, OSError):
        pass


# ── color_hex ────────────────────────────────────────────────────────────────

def test_color_hex_valid():
    assert InputGuard.color_hex("FF0000") == "FF0000"


def test_color_hex_lowercase_valid():
    assert InputGuard.color_hex("a1b2c3") == "a1b2c3"


def test_color_hex_wrong_length_raises():
    with pytest.raises(ValueError, match="color"):
        InputGuard.color_hex("FFF")


def test_color_hex_non_hex_raises():
    with pytest.raises(ValueError, match="color"):
        InputGuard.color_hex("GGGGGG")


# ── bounded_int ───────────────────────────────────────────────────────────────

def test_bounded_int_valid():
    assert InputGuard.bounded_int(5, 1, 10, "width") == 5


def test_bounded_int_at_bounds():
    assert InputGuard.bounded_int(1, 1, 10, "width") == 1
    assert InputGuard.bounded_int(10, 1, 10, "width") == 10


def test_bounded_int_below_raises():
    with pytest.raises(ValueError, match="width"):
        InputGuard.bounded_int(0, 1, 10, "width")


def test_bounded_int_above_raises():
    with pytest.raises(ValueError, match="width"):
        InputGuard.bounded_int(11, 1, 10, "width")


@given(st.integers(min_value=-(2**62), max_value=2**62))
def test_bounded_int_arbitrary_never_crashes(n):
    try:
        InputGuard.bounded_int(n, 1, 100, "test")
    except ValueError:
        pass


# ── regex_pattern ─────────────────────────────────────────────────────────────

def test_regex_pattern_valid():
    pat = InputGuard.regex_pattern(r"\d+")
    assert isinstance(pat, re.Pattern)


def test_regex_pattern_invalid_raises():
    with pytest.raises(ValueError, match="regex"):
        InputGuard.regex_pattern("[unclosed")

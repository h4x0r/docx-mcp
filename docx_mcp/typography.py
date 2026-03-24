"""Smart typography: convert straight quotes and dashes to typographic equivalents."""

from __future__ import annotations


def smartify(text: str) -> str:
    """Convert straight quotes, dashes, and ellipses to smart typography.

    Rules:
    - "text" → \u201ctext\u201d (curly double quotes)
    - 'text' → \u2018text\u2019 (curly single quotes)
    - Apostrophe after letter → \u2019
    - --- → \u2014 (em dash)
    - -- → \u2013 (en dash)
    - ... → \u2026 (ellipsis)
    """
    # Order matters: longest patterns first
    # Em dash before en dash
    text = text.replace("---", "\u2014")
    text = text.replace("--", "\u2013")
    # Ellipsis
    text = text.replace("...", "\u2026")
    # Double quotes
    text = _convert_double_quotes(text)
    # Single quotes / apostrophes
    text = _convert_single_quotes(text)
    return text


def _convert_double_quotes(text: str) -> str:
    """Convert straight double quotes to curly."""
    result = []
    open_quote = True
    for char in text:
        if char == '"':
            result.append("\u201c" if open_quote else "\u201d")
            open_quote = not open_quote
        else:
            result.append(char)
    return "".join(result)


def _convert_single_quotes(text: str) -> str:
    """Convert straight single quotes to curly or apostrophe.

    Heuristic:
    - After a letter or digit → apostrophe (right single quote \u2019)
    - After whitespace, start-of-string, or opening punct → left single quote \u2018
    """
    result = []
    for i, char in enumerate(text):
        if char == "'":
            if i == 0 or (i > 0 and text[i - 1] in " \t\n\r([{"):
                result.append("\u2018")  # left single quote
            else:
                result.append("\u2019")  # apostrophe / right single quote
        else:
            result.append(char)
    return "".join(result)

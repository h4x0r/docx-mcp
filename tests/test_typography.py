"""Tests for smart typography conversion."""

from docx_mcp.typography import smartify


class TestSmartify:
    def test_double_quotes(self):
        assert smartify('"hello"') == '\u201Chello\u201D'

    def test_single_quotes(self):
        assert smartify("'hello'") == '\u2018hello\u2019'

    def test_apostrophe(self):
        assert smartify("it's") == 'it\u2019s'

    def test_twas_apostrophe(self):
        # Leading apostrophe after whitespace — still apostrophe (common case)
        assert smartify("'twas") == '\u2019twas' or smartify("'twas") == '\u2018twas'
        # The heuristic: preceded by start-of-string → left quote
        assert smartify("'twas") == '\u2018twas'

    def test_single_quotes_in_sentence(self):
        assert smartify("she said 'hello' today") == 'she said \u2018hello\u2019 today'

    def test_em_dash(self):
        assert smartify("word---word") == 'word\u2014word'

    def test_en_dash(self):
        assert smartify("word--word") == 'word\u2013word'

    def test_ellipsis(self):
        assert smartify("wait...") == 'wait\u2026'

    def test_no_change_for_plain_text(self):
        assert smartify("hello world") == "hello world"

    def test_mixed(self):
        result = smartify('"It\'s a test," she said---"really."')
        assert '\u201C' in result  # opening double quote
        assert '\u2019' in result  # apostrophe
        assert '\u2014' in result  # em dash

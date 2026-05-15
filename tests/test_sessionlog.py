"""RED tests — Phase 8.1: session log + replay script export."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument


class TestSessionLog:
    def test_get_session_log_empty(self, tmp_path: Path) -> None:
        doc = DocxDocument.create(str(tmp_path / "test.docx"))
        assert doc.get_session_log() == []

    def test_record_op_adds_entry(self, tmp_path: Path) -> None:
        doc = DocxDocument.create(str(tmp_path / "test.docx"))
        doc._record_op("open_document", {"path": "/tmp/a.docx"}, {"ok": True})
        log = doc.get_session_log()
        assert len(log) == 1
        entry = log[0]
        assert entry["tool"] == "open_document"
        assert entry["args"] == {"path": "/tmp/a.docx"}
        assert entry["result"] == {"ok": True}
        assert "timestamp" in entry
        # timestamp must be ISO-format string
        from datetime import datetime
        datetime.fromisoformat(entry["timestamp"])

    def test_record_op_multiple(self, tmp_path: Path) -> None:
        doc = DocxDocument.create(str(tmp_path / "test.docx"))
        doc._record_op("tool_a", {"x": 1}, {"done": True})
        doc._record_op("tool_b", {"y": 2}, {"done": False})
        log = doc.get_session_log()
        assert len(log) == 2
        assert log[0]["tool"] == "tool_a"
        assert log[1]["tool"] == "tool_b"

    def test_export_script_creates_file(self, tmp_path: Path) -> None:
        doc = DocxDocument.create(str(tmp_path / "test.docx"))
        doc._record_op("add_paragraph", {"text": "Hello"}, {"para_id": "ABC"})
        out = tmp_path / "replay.py"
        result = doc.export_session_script(str(out))
        assert out.exists()
        assert result["output_path"] == str(out)
        assert result["operations"] == 1

    def test_export_script_content(self, tmp_path: Path) -> None:
        doc = DocxDocument.create(str(tmp_path / "test.docx"))
        doc._record_op("add_paragraph", {"text": "Hello", "style": "Normal"}, {"para_id": "XYZ"})
        out = tmp_path / "replay.py"
        doc.export_session_script(str(out))
        content = out.read_text(encoding="utf-8")
        assert "from docx_mcp import server" in content
        assert "server.add_paragraph(" in content
        assert "text='Hello'" in content or 'text="Hello"' in content
        assert "style='Normal'" in content or 'style="Normal"' in content

    def test_export_script_empty(self, tmp_path: Path) -> None:
        doc = DocxDocument.create(str(tmp_path / "test.docx"))
        out = tmp_path / "replay_empty.py"
        result = doc.export_session_script(str(out))
        assert out.exists()
        assert result["operations"] == 0
        content = out.read_text(encoding="utf-8")
        assert "from docx_mcp import server" in content

    def test_get_session_log_returns_copy(self, tmp_path: Path) -> None:
        doc = DocxDocument.create(str(tmp_path / "test.docx"))
        doc._record_op("tool_a", {}, {})
        log1 = doc.get_session_log()
        log1.append({"extra": "entry"})
        log2 = doc.get_session_log()
        assert len(log2) == 1

    def test_no_document_raises(self) -> None:
        from docx_mcp import server
        server._doc = None
        with pytest.raises(RuntimeError, match="No document"):
            server.get_session_log()

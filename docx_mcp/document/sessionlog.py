"""Session log mixin: record and replay document operations."""

from __future__ import annotations

from datetime import datetime, timezone


class SessionLogMixin:
    """Record operations performed this session and export as a replay script."""

    def _record_op(self, tool: str, args: dict, result: dict) -> None:
        if not hasattr(self, "_session_log"):
            self._session_log: list[dict] = []
        self._session_log.append(
            {
                "tool": tool,
                "args": args,
                "result": result,
                "timestamp": datetime.now(timezone.utc).isoformat(),
            }
        )

    def get_session_log(self) -> list[dict]:
        """Return all operations performed this session as replayable JSON.

        Each entry: {"tool": str, "args": dict, "result": dict, "timestamp": str (ISO)}
        """
        return list(getattr(self, "_session_log", []))

    def export_session_script(self, output_path: str) -> dict:
        """Write session as a Python script using the MCP tool API.

        Returns: {"output_path": str, "operations": int}
        """
        log = getattr(self, "_session_log", [])
        lines = [
            "from docx_mcp import server",
            "",
        ]
        for entry in log:
            args_repr = ", ".join(f"{k}={v!r}" for k, v in entry["args"].items())
            lines.append(f"server.{entry['tool']}({args_repr})")
        script = "\n".join(lines) + "\n"
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(script)
        return {"output_path": output_path, "operations": len(log)}

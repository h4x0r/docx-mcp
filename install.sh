#!/usr/bin/env bash
# Install docx-mcp: MCP server + Claude Code skill
# Usage: curl -sSL https://raw.githubusercontent.com/SecurityRonin/docx-mcp/main/install.sh | bash
set -euo pipefail

echo "Installing docx-mcp..."

# 1. Add MCP server to Claude Code
if command -v claude &>/dev/null; then
  claude mcp add docx-mcp -- uvx docx-mcp-server
  echo "  ✓ MCP server added to Claude Code"
else
  echo "  ⚠ Claude Code CLI not found — add manually to your MCP config:"
  echo '    {"mcpServers":{"docx-mcp":{"command":"uvx","args":["docx-mcp-server"]}}}'
fi

# 2. Install skill
SKILL_DIR="${HOME}/.claude/skills/docx-mcp"
mkdir -p "$SKILL_DIR"

SKILL_URL="https://raw.githubusercontent.com/SecurityRonin/docx-mcp/main/skill/SKILL.md"
if command -v curl &>/dev/null; then
  curl -sSL "$SKILL_URL" -o "$SKILL_DIR/SKILL.md"
elif command -v wget &>/dev/null; then
  wget -q "$SKILL_URL" -O "$SKILL_DIR/SKILL.md"
else
  echo "  ⚠ Neither curl nor wget found — download skill manually:"
  echo "    $SKILL_URL → $SKILL_DIR/SKILL.md"
  exit 1
fi
echo "  ✓ Skill installed to $SKILL_DIR"

echo ""
echo "Done! Start a new Claude Code session to use docx-mcp."
echo "Try: \"Open contract.docx and change 'Net 30' to 'Net 60' with track changes\""

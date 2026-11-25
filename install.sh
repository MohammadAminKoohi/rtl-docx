#!/usr/bin/env bash
# Install the rtl-docx CLI and register this checkout as a user skill.
#
#   ./install.sh                 core (build/track/comments/validate/check/lint)
#   ./install.sh --with-redline  + compare/audit (43 MB Docxodus engine)
#   ./install.sh --with-all      + redline, merge, MCP server
set -euo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILLS="${AGENTS_SKILLS_DIR:-$HOME/.agents/skills}"
SKILL_SRC="$REPO/skills/rtl-docx"

command -v uv >/dev/null || {
  echo "error: uv is required — https://docs.astral.sh/uv/getting-started/installation/" >&2
  exit 69
}
[[ -f "$SKILL_SRC/SKILL.md" ]] || {
  echo "error: missing $SKILL_SRC/SKILL.md" >&2
  exit 66
}

WITH=()
case "${1:-}" in
  "") ;;
  --with-redline)
    WITH=(--with "python-redlines[docxodus]>=3.0.0")
    ;;
  --with-all|--all)
    WITH=(
      --with "python-redlines[docxodus]>=3.0.0"
      --with "docxcompose>=1.4.0"
      --with "mcp>=2.0.0"
    )
    ;;
  *)
    echo "usage: $0 [--with-redline|--with-all]" >&2
    exit 64
    ;;
esac

uv tool install --editable "$REPO" "${WITH[@]}" --force

mkdir -p "$SKILLS"
ln -sfn "$SKILL_SRC" "$SKILLS/rtl-docx"

echo "installed: $(command -v rtl-docx || echo "$HOME/.local/bin/rtl-docx")"
echo "skill:     $SKILLS/rtl-docx"
echo
rtl-docx doctor || true

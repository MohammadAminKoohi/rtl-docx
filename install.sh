#!/usr/bin/env bash
# Install the rtl-docx CLI and register this checkout as a user skill.
set -euo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILLS="${AGENTS_SKILLS_DIR:-$HOME/.agents/skills}"

command -v uv >/dev/null || {
  echo "error: uv is required — https://docs.astral.sh/uv/getting-started/installation/" >&2
  exit 69
}

uv tool install --editable "$REPO" --force

mkdir -p "$SKILLS"
ln -sfn "$REPO" "$SKILLS/rtl-docx"

echo "installed: $(command -v rtl-docx || echo "$HOME/.local/bin/rtl-docx")"
echo "skill:     $SKILLS/rtl-docx"
echo
rtl-docx doctor || true

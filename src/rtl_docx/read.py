"""Read .docx (and other office files) as token-efficient Markdown."""

from __future__ import annotations

from pathlib import Path


class ReadError(Exception):
    """MarkItDown is unavailable or conversion failed."""


def read_markdown(path: str | Path) -> str:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"not found: {path}")
    try:
        from markitdown import MarkItDown
    except ImportError as exc:  # pragma: no cover
        raise ReadError("MarkItDown is not installed; pip install 'markitdown[docx]'") from exc

    return MarkItDown(enable_plugins=False).convert(str(path)).text_content

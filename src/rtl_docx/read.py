"""Read .docx (and other office files) as token-efficient Markdown."""

from __future__ import annotations

from pathlib import Path


class ReadError(Exception):
    """MarkItDown is unavailable or conversion failed."""


MAX_INPUT = 200 * 1024 * 1024
MAX_RATIO = 200


def _check_archive(path: Path) -> None:
    from zipfile import BadZipFile, ZipFile

    if path.stat().st_size > MAX_INPUT:
        raise ReadError(f"file is larger than {MAX_INPUT} bytes")
    if path.suffix.lower() not in (".docx", ".docm", ".dotx"):
        return
    try:
        with ZipFile(path) as package:
            total = sum(info.file_size for info in package.infolist())
            if total > MAX_INPUT:
                raise ReadError(f"package expands to more than {MAX_INPUT} bytes")
            for info in package.infolist():
                if info.compress_size and info.file_size / info.compress_size > MAX_RATIO:
                    raise ReadError("suspicious compression ratio (possible zip bomb)")
    except BadZipFile as exc:
        raise ReadError(f"not a valid OOXML package: {path}") from exc


def read_markdown(path: str | Path) -> str:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"not found: {path}")
    _check_archive(path)
    try:
        from markitdown import MarkItDown
    except ImportError as exc:  # pragma: no cover
        raise ReadError("MarkItDown is not installed; pip install 'markitdown[docx]'") from exc

    return MarkItDown(enable_plugins=False).convert(str(path)).text_content

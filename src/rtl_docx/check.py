"""Round-trip check: compare a .docx against the Markdown it was built from."""

from __future__ import annotations

import difflib
import re
import shutil
import subprocess
from pathlib import Path

from .common import open_document, region_text
from .persian import normalize_for_search


class CheckError(Exception):
    """Round-trip comparison could not run."""


def _strip_html_wrappers(md: str) -> str:
    return re.sub(r"</?div[^>]*>", "", md)


def _expected_text(markdown: str | Path) -> str:
    markdown = Path(markdown)
    pandoc = shutil.which("pandoc")
    if not pandoc:
        raise CheckError("pandoc is required for round-trip checks")
    source = _strip_html_wrappers(markdown.read_text(encoding="utf-8"))
    proc = subprocess.run(
        [pandoc, "-f", "markdown", "-t", "plain", "--wrap=none"],
        input=source, capture_output=True, text=True, check=True,
    )
    return proc.stdout


def _normalize_lines(text: str) -> list[str]:
    lines = []
    for raw in text.splitlines():
        line = normalize_for_search(raw, ignore_zwnj=True)
        line = re.sub(r"\s+", " ", line).strip()
        line = re.sub(r"^(?:[-*•]|\d+[.)])\s+", "", line)
        line = re.sub(r"\[\d+\]", "", line).strip()
        if not line:
            continue
        if line.startswith("[") and line.endswith("]") and line.count("[") == 1 and line.count("]") == 1:
            line = line[1:-1].strip()
        if not line:
            continue
        if re.fullmatch(r"[-_=]{3,}", line):
            continue
        if len(line) > 3 and re.fullmatch(r"[\s|:+-]+", line):
            continue
        lines.append(line)
    return lines


def _actual_body_text(path: str | Path) -> str:
    doc = open_document(path)
    parts = [
        text for label, text in region_text(doc).items()
        if label.startswith("body/p")
    ]
    return "\n".join(parts)


def check_document(
    docx: str | Path,
    markdown: str | Path,
    max_report: int = 40,
    min_ratio: float = 0.95,
) -> dict:
    docx, markdown = Path(docx), Path(markdown)
    if not docx.exists():
        raise FileNotFoundError(f"not found: {docx}")
    if not markdown.exists():
        raise FileNotFoundError(f"not found: {markdown}")

    expected = _normalize_lines(_expected_text(markdown))
    actual = _normalize_lines(_actual_body_text(docx))

    diff = list(difflib.unified_diff(
        expected, actual,
        fromfile=f"markdown:{markdown.name}",
        tofile=f"docx:{docx.name}",
        lineterm="",
    ))
    changes = [line for line in diff if line.startswith(("+", "-"))
               and not line.startswith(("+++", "---"))]
    expected_words = " ".join(expected).split()
    actual_words = " ".join(actual).split()
    ratio = difflib.SequenceMatcher(None, expected_words, actual_words).ratio()

    return {
        "clean": not changes,
        "pass": ratio >= min_ratio,
        "min_ratio": min_ratio,
        "ratio": round(ratio, 4),
        "expected_lines": len(expected),
        "actual_lines": len(actual),
        "missing_lines": len([line for line in changes if line.startswith("-")]),
        "extra_lines": len([line for line in changes if line.startswith("+")]),
        "diff": "\n".join(diff[:max_report * 2]),
    }

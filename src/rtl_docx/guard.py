"""Structural reporting and cross-region verification for .docx files."""

from __future__ import annotations

import difflib
import re
from pathlib import Path

from .common import open_document, region_text


class CheckFailed(Exception):
    """A verification assertion did not hold."""


def inventory(path: str | Path) -> dict[str, int]:
    doc = open_document(path)
    tables = doc.tables
    cell_count = 0
    nested = 0
    queue = list(tables)
    while queue:
        table = queue.pop()
        rows = list(table.rows)
        cell_count += sum(len(row.cells) for row in rows)
        for row in rows:
            for cell in row.cells:
                nested += len(cell.tables)
                queue.extend(cell.tables)

    headers = footers = 0
    for section in doc.sections:
        for kind in ("header", "footer"):
            for prefix in ("", "first_page_", "even_page_"):
                part = getattr(section, f"{prefix}{kind}")
                if part.is_linked_to_previous:
                    continue
                has_text = any(p.text.strip() for p in part.paragraphs) or bool(part.tables)
                if has_text:
                    if kind == "header":
                        headers += 1
                    else:
                        footers += 1

    text = "\n".join(region_text(doc).values())
    return {
        "sections": len(doc.sections),
        "body_paragraphs": len(doc.paragraphs),
        "tables": len(tables),
        "nested_table_refs": nested,
        "table_cells": cell_count,
        "headers": headers,
        "footers": footers,
        "inline_images": len(doc.inline_shapes),
        "words": len(re.findall(r"\S+", text)),
    }


def count_matches(path: str | Path, needle: str, regex: bool = False) -> dict[str, int]:
    """Occurrences per region label, across body, tables, headers and footers."""
    needle_re = re.compile(needle) if regex else None
    out: dict[str, int] = {}
    for label, text in region_text(open_document(path)).items():
        n = len(needle_re.findall(text)) if regex else text.count(needle)
        if n:
            out[label] = n
    return out


def group_matches(hits: dict[str, int]) -> dict[str, int]:
    groups: dict[str, int] = {}
    for label, n in hits.items():
        key = label.split("/p")[0] if "/p" in label else label
        groups[key] = groups.get(key, 0) + n
    return groups


def assert_count(total: int, *, exact: int | None, minimum: int | None, allow_zero: bool) -> None:
    if exact is not None and total != exact:
        raise CheckFailed(f"expected exactly {exact} match(es), found {total}")
    if minimum is not None and total < minimum:
        raise CheckFailed(f"expected at least {minimum} match(es), found {total}")
    if total == 0 and not allow_zero:
        raise CheckFailed("no match found")


def dump_text(path: str | Path) -> str:
    lines: list[str] = []
    for label, text in region_text(open_document(path)).items():
        lines.append(f"### {label}")
        lines.append(text)
        lines.append("")
    return "\n".join(lines)


def diff_docs(a: str | Path, b: str | Path) -> tuple[str, int]:
    ta, tb = region_text(open_document(a)), region_text(open_document(b))
    labels = list(ta) + [x for x in tb if x not in ta]
    chunks: list[str] = []
    changed = 0
    for label in labels:
        before, after = ta.get(label, ""), tb.get(label, "")
        if before == after:
            continue
        changed += 1
        chunks.append(f"### {label}")
        chunks.extend(
            difflib.unified_diff(
                before.splitlines(), after.splitlines(),
                fromfile=f"old/{label}", tofile=f"new/{label}", lineterm="",
            )
        )
        chunks.append("")
    payload = "\n".join(chunks) if chunks else "(no textual differences)\n"
    return payload, changed

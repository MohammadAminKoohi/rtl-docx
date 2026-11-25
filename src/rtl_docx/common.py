"""Shared docx traversal helpers: body, tables, headers, footers."""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path
from zipfile import BadZipFile

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml.ns import qn
from docx.table import Table


def open_document(path: str | Path):
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"not found: {path}")
    try:
        return Document(str(path))
    except (PackageNotFoundError, BadZipFile) as exc:
        raise FileNotFoundError(f"not a valid .docx package: {path}") from exc


def _dedupe_cells(table: Table):
    seen = set()
    for row in table.rows:
        for cell in row.cells:
            key = id(cell._tc)
            if key in seen:
                continue
            seen.add(key)
            yield cell


def iter_table_paragraphs(tables, prefix: str) -> Iterator[tuple[str, object]]:
    for ti, table in enumerate(tables, 1):
        for ci, cell in enumerate(_dedupe_cells(table), 1):
            label = f"{prefix}table{ti}/cell{ci}"
            for pi, paragraph in enumerate(cell.paragraphs, 1):
                yield f"{label}/p{pi}", paragraph
            yield from iter_table_paragraphs(cell.tables, f"{label}/")


HEADER_FOOTER_KINDS = (
    "header", "footer",
    "first_page_header", "first_page_footer",
    "even_page_header", "even_page_footer",
)


def iter_regions(doc) -> Iterator[tuple[str, object]]:
    """Yield (region_label, paragraph) for every editable/readable paragraph.

    Covers body paragraphs, body tables (recursively), and headers/footers that
    are actually defined (linked ones inherit content from earlier parts).
    """
    for i, paragraph in enumerate(doc.paragraphs, 1):
        yield f"body/p{i}", paragraph

    yield from iter_table_paragraphs(doc.tables, "body/")

    for si, section in enumerate(doc.sections, 1):
        for kind in HEADER_FOOTER_KINDS:
            part = getattr(section, kind)
            if part.is_linked_to_previous:
                continue
            for pi, paragraph in enumerate(part.paragraphs, 1):
                yield f"section{si}/{kind}/p{pi}", paragraph
            yield from iter_table_paragraphs(part.tables, f"section{si}/{kind}/")


def paragraph_text(paragraph) -> str:
    """Text including tracked insertions, excluding tracked deletions.

    python-docx's ``paragraph.text`` only reads direct ``w:r`` children, so it
    drops runs wrapped in ``<w:ins>``. Verification must not lie about what a
    reader will see.
    """
    out: list[str] = []
    for node in paragraph._p.iter():
        if node.tag == qn("w:t"):
            out.append(node.text or "")
        elif node.tag == qn("w:tab"):
            out.append("\t")
        elif node.tag in (qn("w:br"), qn("w:cr")):
            out.append("\n")
    return "".join(out)


def paragraph_text_without_insertions(paragraph) -> str:
    """Visible text ignoring runs inside <w:ins> (for 'old text survived' checks)."""
    out: list[str] = []
    for node in paragraph._p.iter(qn("w:t")):
        if any(anc.tag == qn("w:ins") for anc in node.iterancestors()):
            continue
        out.append(node.text or "")
    return "".join(out)


def region_text(doc) -> dict[str, str]:
    """Region label -> visible text, preserving document order."""
    out: dict[str, str] = {}
    for label, paragraph in iter_regions(doc):
        text = paragraph_text(paragraph)
        if text.strip():
            out[label] = text
    return out

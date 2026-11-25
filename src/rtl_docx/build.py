"""Build a styled, RTL-first .docx from Markdown with pandoc."""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ASSETS = Path(__file__).resolve().parent / "assets"
DEFAULT_REFERENCE = ASSETS / "reference.docx"

PPR_AFTER_BIDI = (
    "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
    "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)
RPR_AFTER_RTL = ("w:cs", "w:em", "w:lang", "w:eastAsianLayout",
                 "w:specVanish", "w:oMath")


def _set_flag(parent, tag: str, value: str, successors: tuple[str, ...]) -> None:
    node = parent.find(qn(tag))
    if node is None:
        node = OxmlElement(tag)
        parent.insert_element_before(node, *successors)
    node.set(qn("w:val"), value)


def _fix_code_paragraph(paragraph) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    _set_flag(ppr, "w:bidi", "0", PPR_AFTER_BIDI)
    _set_flag(ppr, "w:jc", "left", (
        "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
        "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
    ))
    for run in paragraph.runs:
        _set_flag(run._r.get_or_add_rPr(), "w:rtl", "0", RPR_AFTER_RTL)


def force_code_ltr(path: str | Path) -> int:
    """Stamp LTR directly on code paragraphs; some renderers ignore style-only bidi."""
    doc = Document(str(path))
    fixed = 0

    def visit(paragraphs):
        nonlocal fixed
        for paragraph in paragraphs:
            style = paragraph.style
            if style is not None and style.style_id in ("SourceCode", "Source Code"):
                _fix_code_paragraph(paragraph)
                fixed += 1

    visit(doc.paragraphs)
    queue = list(doc.tables)
    while queue:
        table = queue.pop()
        for row in table.rows:
            for cell in row.cells:
                visit(cell.paragraphs)
                queue.extend(cell.tables)

    doc.save(str(path))
    return fixed


def build_docx(
    source: str | Path,
    output: str | Path,
    reference: str | Path | None = None,
    extra_pandoc_args: list[str] | None = None,
) -> Path:
    source, output = Path(source), Path(output)
    if not source.exists():
        raise FileNotFoundError(f"input not found: {source}")
    if source.suffix.lower() != ".md":
        raise ValueError("input must be Markdown (.md)")
    if output.suffix.lower() != ".docx":
        raise ValueError("output must end in .docx")
    if output.exists():
        raise FileExistsError(f"refusing to overwrite {output} (use a new path)")

    pandoc = shutil.which("pandoc")
    if not pandoc:
        raise RuntimeError("pandoc is not installed: https://pandoc.org/installing.html")

    ref = Path(reference) if reference else DEFAULT_REFERENCE
    if not ref.exists():
        raise FileNotFoundError(f"reference doc not found: {ref}")

    output.parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        pandoc, str(source),
        "--from=markdown", "--to=docx",
        f"--reference-doc={ref}",
        f"--resource-path={source.parent}",
        "-o", str(output),
        *(extra_pandoc_args or []),
    ]
    subprocess.run(cmd, check=True)
    force_code_ltr(output)
    return output


def export_reference(output: str | Path) -> Path:
    """Copy the bundled RTL reference template for editing in Word."""
    output = Path(output)
    if output.exists():
        raise FileExistsError(f"refusing to overwrite {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(DEFAULT_REFERENCE, output)
    return output

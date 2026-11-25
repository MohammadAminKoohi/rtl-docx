"""Build a styled, RTL-first .docx from Markdown with pandoc."""

from __future__ import annotations

import os
import re
import shutil
import subprocess
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ASSETS = Path(__file__).resolve().parent / "assets"
DEFAULT_REFERENCE = ASSETS / "reference.docx"
CODE_STYLES = ("SourceCode", "Source Code")

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


def _stamp_run(run_el, rtl: str) -> None:
    rpr = run_el.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        run_el.insert(0, rpr)
    _set_flag(rpr, "w:rtl", rtl, RPR_AFTER_RTL)


def _stamp_paragraph(paragraph, is_code: bool) -> None:
    if is_code:
        ppr = paragraph._p.get_or_add_pPr()
        _set_flag(ppr, "w:bidi", "0", PPR_AFTER_BIDI)
        _set_flag(ppr, "w:jc", "left", (
            "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
            "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
        ))
    for run_el in paragraph._p.iter(qn("w:r")):
        _stamp_run(run_el, "0" if is_code else "1")


def _is_code(paragraph) -> bool:
    style = paragraph.style
    return style is not None and style.style_id in CODE_STYLES


def _iter_paragraphs(doc):
    yield from doc.paragraphs
    queue = list(doc.tables)
    while queue:
        table = queue.pop()
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs
                queue.extend(cell.tables)

    for section in doc.sections:
        for kind in ("header", "footer", "first_page_header", "first_page_footer",
                     "even_page_header", "even_page_footer"):
            part = getattr(section, kind)
            if part.is_linked_to_previous:
                continue
            yield from part.paragraphs
            queue = list(part.tables)
            while queue:
                table = queue.pop()
                for row in table.rows:
                    for cell in row.cells:
                        yield from cell.paragraphs
                        queue.extend(cell.tables)


def apply_rtl_direction(path: str | Path) -> int:
    """Stamp w:rtl on runs (valid) rather than style rPr (invalid in OOXML)."""
    doc = Document(str(path))
    count = 0
    for paragraph in _iter_paragraphs(doc):
        _stamp_paragraph(paragraph, _is_code(paragraph))
        count += 1
    doc.save(str(path))
    return count


def drop_styles_with_effects(path: str | Path) -> bool:
    """Remove the optional stylesWithEffects part and its references."""
    path = Path(path)
    with ZipFile(path) as zin:
        names = zin.namelist()
        if "word/stylesWithEffects.xml" not in names:
            return False
        payloads = {name: zin.read(name) for name in names}

    payloads.pop("word/stylesWithEffects.xml")
    ct = payloads["[Content_Types].xml"].decode("utf-8")
    ct = re.sub(r'<Override[^>]*PartName="/word/stylesWithEffects\.xml"[^>]*/>', "", ct)
    payloads["[Content_Types].xml"] = ct.encode("utf-8")

    rels_name = "word/_rels/document.xml.rels"
    if rels_name in payloads:
        rels = payloads[rels_name].decode("utf-8")
        rels = re.sub(r'<Relationship[^>]*Target="stylesWithEffects\.xml"[^>]*/>', "", rels)
        payloads[rels_name] = rels.encode("utf-8")

    tmp = path.with_name(path.stem + ".tmp.docx")
    with ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
        for name, payload in payloads.items():
            zout.writestr(name, payload)
    os.replace(tmp, path)
    return True


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
    apply_rtl_direction(output)
    drop_styles_with_effects(output)
    return output


def export_reference(output: str | Path) -> Path:
    """Copy the bundled RTL reference template for editing in Word."""
    output = Path(output)
    if output.exists():
        raise FileExistsError(f"refusing to overwrite {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(DEFAULT_REFERENCE, output)
    return output

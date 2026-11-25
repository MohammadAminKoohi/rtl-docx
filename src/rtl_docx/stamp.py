"""Add headers, footers and page-number fields to every section."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}
PPR_AFTER_BIDI = (
    "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
    "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)
RPR_AFTER_RTL = ("w:cs", "w:em", "w:lang", "w:eastAsianLayout",
                 "w:specVanish", "w:oMath")


class StampError(Exception):
    """Header/footer operation failed."""


def _ensure_part_styles(doc) -> None:
    try:
        normal = doc.styles["Normal"]
    except KeyError:
        normal = None
    for name in ("Header", "Footer"):
        try:
            doc.styles[name]
        except KeyError:
            style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            style.element.set(qn("w:styleId"), name)
            if normal is not None:
                style.base_style = normal


def _rtl_paragraph(paragraph, align: str) -> None:
    paragraph.alignment = ALIGNMENTS.get(align, WD_ALIGN_PARAGRAPH.CENTER)
    ppr = paragraph._p.get_or_add_pPr()
    if ppr.find(qn("w:bidi")) is None:
        bidi = OxmlElement("w:bidi")
        ppr.insert_element_before(bidi, *PPR_AFTER_BIDI)


def _rtl_run(run) -> None:
    rpr = run._r.get_or_add_rPr()
    if rpr.find(qn("w:rtl")) is None:
        rtl = OxmlElement("w:rtl")
        rpr.insert_element_before(rtl, *RPR_AFTER_RTL)


def _clear(paragraph) -> None:
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)


def _remove_existing_page_field(paragraph) -> None:
    for field in paragraph._p.iter(qn("w:fldSimple")):
        if "PAGE" in (field.get(qn("w:instr")) or ""):
            parent = field.getparent()
            if parent is not None:
                parent.remove(field)


def _append_page_field(paragraph, label: str) -> None:
    if label:
        run = paragraph.add_run(label)
        _rtl_run(run)

    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), " PAGE ")
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rpr.append(OxmlElement("w:rtl"))
    run.append(rpr)
    text = OxmlElement("w:t")
    text.text = "1"
    run.append(text)
    field.append(run)
    paragraph._p.append(field)


def _set_part_text(part, text: str, align: str) -> None:
    paragraphs = part.paragraphs
    paragraph = paragraphs[0] if paragraphs else part.add_paragraph()
    _clear(paragraph)
    _rtl_paragraph(paragraph, align)
    run = paragraph.add_run(text)
    _rtl_run(run)


def stamp_document(
    path: str | Path,
    header: str | None = None,
    footer: str | None = None,
    page_numbers: bool = False,
    page_label: str = "",
    align: str = "center",
) -> dict:
    doc = Document(str(path))
    _ensure_part_styles(doc)
    sections = 0
    for index, section in enumerate(doc.sections):
        if header is not None and (index == 0 or not section.header.is_linked_to_previous):
            _set_part_text(section.header, header, align)
            sections += 1
        if footer is None and not page_numbers:
            continue
        if index > 0 and section.footer.is_linked_to_previous:
            continue
        paragraphs = section.footer.paragraphs
        paragraph = paragraphs[0] if paragraphs else section.footer.add_paragraph()
        if footer is not None:
            _clear(paragraph)
        _rtl_paragraph(paragraph, align)
        _remove_existing_page_field(paragraph)
        if footer is not None:
            run = paragraph.add_run(footer + (" " if page_numbers else ""))
            _rtl_run(run)
        if page_numbers:
            _append_page_field(paragraph, page_label)
            sections += 1

    doc.save(str(path))
    from .build import drop_styles_with_effects
    drop_styles_with_effects(path)

    from .validate import validate_docx
    result = validate_docx(path, schema=False)
    if not result["valid"]:
        raise StampError(
            "stamped package is structurally invalid:\n"
            + "\n".join(result["errors"][:5])
        )
    return {"sections": sections, "page_numbers": page_numbers}


def stamp_copy(source: str | Path, output: str | Path, **kwargs) -> Path:
    import shutil

    source, output = Path(source), Path(output)
    if not source.exists():
        raise FileNotFoundError(f"not found: {source}")
    if output.exists():
        raise FileExistsError(f"refusing to overwrite {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, output)
    stamp_document(output, **kwargs)
    return output

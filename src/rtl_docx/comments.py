"""Read and add Word comments with python-docx 1.2's native comment API."""

from __future__ import annotations

from pathlib import Path

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.shared import Pt

from .common import iter_regions, open_document


class CommentError(Exception):
    """Comment operation failed."""


def _as_datetime(value):
    return value.isoformat() if hasattr(value, "isoformat") else (str(value) if value else None)


def _anchors(doc) -> dict[int, str]:
    """Map comment id -> text between commentRangeStart/End."""
    anchors: dict[int, str] = {}
    current: int | None = None
    buffer: list[str] = []
    for node in doc.element.body.iter():
        if node.tag == qn("w:commentRangeStart"):
            raw = node.get(qn("w:id"))
            current = int(raw) if raw is not None else None
            buffer = []
        elif node.tag == qn("w:t") and current is not None:
            buffer.append(node.text or "")
        elif node.tag == qn("w:commentRangeEnd") and current is not None:
            anchors[current] = "".join(buffer)
            current = None
    return anchors


def list_comments(path: str | Path) -> list[dict]:
    doc = open_document(path)
    anchors = _anchors(doc)
    return [
        {
            "id": comment.comment_id,
            "author": comment.author,
            "initials": comment.initials,
            "date": _as_datetime(comment.timestamp),
            "text": comment.text,
            "anchor": anchors.get(comment.comment_id, ""),
        }
        for comment in doc.comments
    ]


def _ensure_comment_styles(doc) -> None:
    """python-docx references CommentReference/CommentText but does not create them."""
    try:
        default_font = doc.styles["Default Paragraph Font"]
    except KeyError:
        default_font = None
    try:
        normal = doc.styles["Normal"]
    except KeyError:
        normal = None

    try:
        doc.styles["CommentReference"]
    except KeyError:
        style = doc.styles.add_style("CommentReference", WD_STYLE_TYPE.CHARACTER)
        style.element.set(qn("w:styleId"), "CommentReference")
        style.font.superscript = True
        style.font.hidden = True
        if default_font is not None:
            style.base_style = default_font

    try:
        doc.styles["CommentText"]
    except KeyError:
        style = doc.styles.add_style("CommentText", WD_STYLE_TYPE.PARAGRAPH)
        style.element.set(qn("w:styleId"), "CommentText")
        if normal is not None:
            style.base_style = normal
        style.font.size = Pt(10)


def _find_runs(paragraph, anchor: str, skip: int) -> tuple[list, int] | None:
    runs = paragraph.runs
    full = "".join(run.text for run in runs)
    idx = -1
    for _ in range(skip + 1):
        idx = full.find(anchor, idx + 1)
        if idx < 0:
            return None
    end = idx + len(anchor)

    pos = 0
    start_i = end_i = None
    for i, run in enumerate(runs):
        run_end = pos + len(run.text)
        if start_i is None and pos <= idx < run_end:
            start_i = i
        if pos < end <= run_end:
            end_i = i
            break
        pos = run_end
    if start_i is None or end_i is None:
        return None
    return runs[start_i:end_i + 1], idx


def add_comment(
    source: str | Path,
    output: str | Path,
    anchor: str,
    text: str,
    author: str,
    initials: str | None = None,
    occurrence: int = 1,
) -> dict:
    source, output = Path(source), Path(output)
    if not source.exists():
        raise FileNotFoundError(f"not found: {source}")
    if output.exists():
        raise FileExistsError(f"refusing to overwrite {output}")
    if source.resolve() == output.resolve():
        raise CommentError("refusing in-place comment; choose a new output path")
    if not anchor.strip():
        raise CommentError("anchor must not be empty")
    if occurrence < 1:
        raise CommentError("occurrence must be >= 1")

    doc = open_document(source)
    _ensure_comment_styles(doc)
    seen = 0
    for label, paragraph in iter_regions(doc):
        full = "".join(run.text for run in paragraph.runs)
        in_paragraph = full.count(anchor)
        if not in_paragraph:
            continue
        if seen + in_paragraph < occurrence:
            seen += in_paragraph
            continue
        local = occurrence - seen - 1
        found = _find_runs(paragraph, anchor, local)
        if found is None:
            raise CommentError(
                f"anchor found in {label} but not inside editable runs "
                "(it may be in a hyperlink or field)"
            )
        runs, _idx = found
        try:
            comment = doc.add_comment(
                runs, text=text, author=author, initials=initials or ""
            )
        except Exception as exc:  # noqa: BLE001
            raise CommentError(f"could not add comment: {exc}") from exc

        output.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output))

        from .build import drop_styles_with_effects
        drop_styles_with_effects(output)

        from .validate import validate_docx
        result = validate_docx(output, schema=False)
        if not result["valid"]:
            raise CommentError(
                "commented document failed validation:\n" + "\n".join(result["errors"][:5])
            )
        return {
            "output": str(output),
            "comment_id": comment.comment_id,
            "anchor": anchor,
            "region": label,
        }

    raise CommentError(
        f"anchor not found inside a single paragraph ({occurrence} occurrence(s) requested)"
    )

"""Read, accept, reject, and audit Word tracked changes across every story part."""

from __future__ import annotations

import difflib
import re
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from lxml import etree

from .common import open_document, region_text
from .persian import normalize_for_search

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
REVISION_TAGS = {
    f"{W}ins": "insertion",
    f"{W}del": "deletion",
    f"{W}moveTo": "move_to",
    f"{W}moveFrom": "move_from",
}
REVISION_PART_RE = re.compile(
    r"^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$"
)


class RevisionError(Exception):
    """Revision operation failed."""


def _xml_part(name: str) -> bool:
    return bool(REVISION_PART_RE.match(name))


def _revision_parts(package: ZipFile) -> list[str]:
    return sorted(n for n in package.namelist() if _xml_part(n))


def _text_of(element) -> str:
    parts = []
    for node in element.iter():
        if node.tag in (f"{W}t", f"{W}delText", f"{W}instrText", f"{W}delInstrText"):
            parts.append(node.text or "")
    return "".join(parts)


def _paragraph_of(element):
    node = element.getparent()
    while node is not None and node.tag != f"{W}p":
        node = node.getparent()
    return node


def _as_datetime(value):
    return value if value else None


def list_revisions(path: str | Path) -> list[dict]:
    items: list[dict] = []
    with ZipFile(path) as package:
        for name in _revision_parts(package):
            root = etree.fromstring(package.read(name))
            for element in root.iter():
                kind = REVISION_TAGS.get(element.tag)
                if not kind:
                    continue
                paragraph = _paragraph_of(element)
                preview = _text_of(paragraph) if paragraph is not None else ""
                items.append({
                    "type": kind,
                    "text": _text_of(element) or "(paragraph mark)",
                    "author": element.get(f"{W}author", ""),
                    "date": _as_datetime(element.get(f"{W}date")),
                    "revision_id": element.get(f"{W}id"),
                    "part": name,
                    "paragraph": preview[:200],
                })
    return items


def revision_summary(path: str | Path) -> dict:
    items = list_revisions(path)
    by_type: dict[str, int] = {}
    by_author: dict[str, int] = {}
    for item in items:
        by_type[item["type"]] = by_type.get(item["type"], 0) + 1
        by_author[item["author"]] = by_author.get(item["author"], 0) + 1
    return {"count": len(items), "by_type": by_type, "by_author": by_author}


def count_revisions_xml(path: str | Path) -> int:
    total = 0
    with ZipFile(path) as package:
        for name in _revision_parts(package):
            xml = package.read(name).decode("utf-8", "ignore")
            total += len(re.findall(r"<w:(?:ins|del|moveFrom|moveTo)\b", xml))
    return total


def _unwrap(element) -> None:
    parent = element.getparent()
    if parent is None:
        return
    index = parent.index(element)
    for child in list(element):
        parent.insert(index, child)
        index += 1
    parent.remove(element)


def _remove(element) -> None:
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def _paragraph_marked(element, tag: str) -> bool:
    ppr = element.find(f"{W}pPr")
    if ppr is None:
        return False
    rpr = ppr.find(f"{W}rPr")
    return rpr is not None and rpr.find(f"{W}{tag}") is not None


def _merge_into_next(paragraph) -> None:
    parent = paragraph.getparent()
    if parent is None:
        return
    siblings = list(parent)
    try:
        start = siblings.index(paragraph) + 1
    except ValueError:
        return
    next_p = next((s for s in siblings[start:] if s.tag == f"{W}p"), None)
    if next_p is None:
        return
    ppr = next_p.find(f"{W}pPr")
    insert_at = 1 if ppr is not None else 0
    content = [child for child in paragraph if child.tag != f"{W}pPr"]
    for offset, child in enumerate(content):
        next_p.insert(insert_at + offset, child)
    parent.remove(paragraph)


def _transform(data: bytes, mode: str) -> bytes:
    root = etree.fromstring(data)
    merge: list = []

    for paragraph in root.iter(f"{W}p"):
        marker = "del" if mode == "accept" else "ins"
        if _paragraph_marked(paragraph, marker):
            merge.append(paragraph)

    if mode == "accept":
        for element in list(root.iter()):
            if element.tag in (f"{W}ins", f"{W}moveTo"):
                _unwrap(element)
            elif element.tag in (f"{W}del", f"{W}moveFrom"):
                _remove(element)
    else:
        for element in list(root.iter()):
            if element.tag in (f"{W}del", f"{W}moveFrom"):
                for node in element.iter():
                    if node.tag == f"{W}delText":
                        node.tag = f"{W}t"
                    elif node.tag == f"{W}delInstrText":
                        node.tag = f"{W}instrText"
                _unwrap(element)
            elif element.tag in (f"{W}ins", f"{W}moveTo"):
                _remove(element)

    for paragraph in merge:
        if paragraph.getparent() is not None:
            _merge_into_next(paragraph)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _rewrite(source: Path, output: Path, mode: str) -> Path:
    with ZipFile(source) as zin:
        items = [(info, zin.read(info.filename)) for info in zin.infolist()]

    output.parent.mkdir(parents=True, exist_ok=True)
    tmp = output.with_name(output.stem + ".tmp.docx")
    with ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
        for info, payload in items:
            if _xml_part(info.filename):
                payload = _transform(payload, mode)
            new_info = ZipInfo(info.filename, date_time=info.date_time)
            new_info.compress_type = ZIP_DEFLATED
            new_info.external_attr = info.external_attr
            zout.writestr(new_info, payload)
    tmp.replace(output)
    return output


def _check_paths(source: Path, output: Path) -> None:
    if source.resolve() == output.resolve():
        raise RevisionError("refusing in-place operation; choose a new output path")
    if output.exists():
        raise FileExistsError(f"refusing to overwrite {output}")
    if not source.exists():
        raise FileNotFoundError(f"not found: {source}")


def accept_changes(source: str | Path, output: str | Path) -> Path:
    source, output = Path(source), Path(output)
    _check_paths(source, output)
    return _rewrite(source, output, "accept")


def reject_changes(source: str | Path, output: str | Path) -> Path:
    source, output = Path(source), Path(output)
    _check_paths(source, output)
    return _rewrite(source, output, "reject")


def _paragraph_sequence(path: str | Path) -> list[str]:
    texts = []
    for text in region_text(open_document(path)).values():
        normalized = normalize_for_search(text)
        normalized = re.sub(r"\s+", " ", normalized).strip()
        if normalized:
            texts.append(normalized)
    return texts


def audit_changes(
    path: str | Path,
    original: str | Path,
    author: str | None = None,
) -> dict:
    """Reject all revisions and compare the result with the original.

    Any remaining text difference is an edit made without a tracked revision.
    Optionally require every revision to carry a given author.
    """
    path, original = Path(path), Path(original)
    if not path.exists() or not original.exists():
        raise FileNotFoundError("redlined file or original not found")

    items = list_revisions(path)
    authors = sorted({item["author"] for item in items if item["author"]})
    author_mismatches = []
    if author:
        author_mismatches = sorted({item["author"] for item in items if item["author"] != author})

    with tempfile.TemporaryDirectory(prefix="rtl-docx-audit-") as tmp:
        rejected = Path(tmp) / "rejected.docx"
        reject_changes(path, rejected)
        expected = _paragraph_sequence(original)
        observed = _paragraph_sequence(rejected)

    diff = list(difflib.unified_diff(
        expected, observed,
        fromfile=f"original:{original.name}",
        tofile=f"redline-minus-revisions:{path.name}",
        lineterm="",
    ))
    untracked = [line for line in diff
                 if line.startswith(("+", "-")) and not line.startswith(("+++", "---"))]

    return {
        "clean": not untracked and not author_mismatches,
        "revisions": len(items),
        "authors": authors,
        "author_mismatches": author_mismatches,
        "untracked_diff": "\n".join(diff),
        "untracked_count": len(untracked),
    }


def compare_documents(
    original: str | Path,
    modified: str | Path,
    output: str | Path,
    author: str = "Codex",
    detect_moves: bool = False,
    case_insensitive: bool = False,
) -> dict:
    """Produce a native tracked-change redline between two .docx files.

    Requires the optional extra: uv tool install "rtl-docx[redline]".
    """
    original, modified, output = Path(original), Path(modified), Path(output)
    if not original.exists() or not modified.exists():
        raise FileNotFoundError("both input documents must exist")
    if output.exists():
        raise FileExistsError(f"refusing to overwrite {output}")
    if output.resolve() in (original.resolve(), modified.resolve()):
        raise RevisionError("output must differ from both inputs")

    try:
        from python_redlines import DocxodusEngine
    except ImportError as exc:
        raise RevisionError(
            'redlining engine not installed: uv tool install "rtl-docx[redline]"'
        ) from exc

    engine = DocxodusEngine()
    kwargs = {"detect_moves": detect_moves, "case_insensitive": case_insensitive}
    try:
        redline_bytes, _stdout, stderr = engine.run_redline(
            author, original.read_bytes(), modified.read_bytes(), **kwargs
        )
    except TypeError:
        redline_bytes, _stdout, stderr = engine.run_redline(
            author, original.read_bytes(), modified.read_bytes()
        )

    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(redline_bytes)

    from .build import drop_styles_with_effects
    from .validate import repair_required_parts, sanitize_package, validate_docx

    drop_styles_with_effects(output)
    removed = sanitize_package(output)
    repaired = repair_required_parts(output)
    result = validate_docx(output, schema=False)
    if not result["valid"]:
        raise RevisionError(
            "engine produced a structurally invalid document:\n"
            + "\n".join(result["errors"][:5])
        )

    return {
        "output": str(output),
        "revisions": count_revisions_xml(output),
        "sanitized_relationships": removed,
        "repaired_parts": repaired,
        "engine_stdout": (_stdout or "").strip(),
        "engine_stderr": (stderr or "").strip(),
    }

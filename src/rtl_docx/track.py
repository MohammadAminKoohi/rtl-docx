"""Propose edits to an existing .docx as native Word tracked changes."""

from __future__ import annotations

import copy
import random
from datetime import datetime, timezone
from pathlib import Path

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.run import Run

from .common import (
    iter_regions, open_document, paragraph_text,
    paragraph_text_without_insertions, region_text,
)
from .persian import count_occurrences, map_variants

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
UNSAFE_RUN_TAGS = (
    "w:br", "w:tab", "w:cr", "w:drawing", "w:object",
    "w:fldChar", "w:instrText", "w:footnoteReference", "w:endnoteReference",
)

_ids = iter(range(random.randint(100_000, 900_000), 1_000_000))


class EditError(Exception):
    """An edit could not be applied safely."""


def _next_id() -> str:
    return str(next(_ids))


def _assert_plain(run: Run) -> None:
    for tag in UNSAFE_RUN_TAGS:
        if run._r.find(qn(tag)) is not None:
            raise EditError(
                f"run contains <{tag}>; normalise the paragraph before tracking edits"
            )


def _set_run_text(r_el, text: str) -> None:
    ts = r_el.findall(qn("w:t"))
    if not ts:
        t = OxmlElement("w:t")
        r_el.append(t)
        ts = [t]
    for extra in ts[1:]:
        r_el.remove(extra)
    ts[0].text = text
    ts[0].set(XML_SPACE, "preserve")


def _split_run(run: Run, offset: int) -> Run | None:
    if offset <= 0:
        return run
    text = run.text
    if offset >= len(text):
        return None
    _assert_plain(run)
    new_r = copy.deepcopy(run._r)
    run._r.addnext(new_r)
    _set_run_text(run._r, text[:offset])
    _set_run_text(new_r, text[offset:])
    return Run(new_r, run._parent)


def _stamp(el, author: str, date: str) -> None:
    el.set(qn("w:id"), _next_id())
    el.set(qn("w:author"), author)
    el.set(qn("w:date"), date)


def _replace_once(
    paragraph,
    needle: str,
    replacement: str,
    author: str,
    date: str,
    normalize: bool = False,
) -> bool:
    runs = paragraph.runs
    full = "".join(r.text for r in runs)
    haystack = map_variants(full) if normalize else full
    target = map_variants(needle) if normalize else needle
    idx = haystack.find(target)
    if idx < 0:
        return False

    pos = 0
    start_run = None
    start_off = 0
    for run in runs:
        length = len(run.text)
        if pos <= idx < pos + length:
            start_run, start_off = run, idx - pos
            break
        pos += length
    if start_run is None:
        raise EditError(f"could not locate match in paragraph: {full!r}")

    _assert_plain(start_run)
    if start_off:
        start_run = _split_run(start_run, start_off)
    start_el = start_run._r

    runs = paragraph.runs
    si = next(i for i, r in enumerate(runs) if r._r is start_el)

    consumed = 0
    ei = si
    for j in range(si, len(runs)):
        length = len(runs[j].text)
        if consumed + length >= len(needle):
            ei = j
            break
        consumed += length
    else:
        raise EditError(f"match incomplete in paragraph: {needle!r}")

    overflow = consumed + len(runs[ei].text) - len(needle)
    if overflow > 0:
        _assert_plain(runs[ei])
        _split_run(runs[ei], len(runs[ei].text) - overflow)

    selected = runs[si:ei + 1]
    for run in selected:
        _assert_plain(run)

    elements = [r._r for r in selected]
    del_el = OxmlElement("w:del")
    _stamp(del_el, author, date)

    ins_el = None
    if replacement:
        tmpl = copy.deepcopy(elements[0])
        _set_run_text(tmpl, replacement)
        ins_el = OxmlElement("w:ins")
        _stamp(ins_el, author, date)
        ins_el.append(tmpl)

    first = elements[0]
    first.addprevious(del_el)
    for r_el in elements:
        for node in r_el.iter(qn("w:t")):
            node.tag = qn("w:delText")
        del_el.append(r_el)
    if ins_el is not None:
        del_el.addprevious(ins_el)
    return True


def _counts(texts: dict[str, str], needle: str, normalize: bool) -> dict[str, int]:
    if not needle:
        return {}
    return {
        label: count_occurrences(text, needle, normalize)
        for label, text in texts.items()
    }


def apply_edits(
    source: str | Path,
    output: str | Path,
    edits: list[dict[str, str]],
    author: str = "Codex",
    dry_run: bool = False,
    normalize: bool = False,
) -> dict:
    """Apply find/replace edits as w:ins/w:del revisions.

    Returns per-region receipts: where each edit landed, how much old text is
    left visible, and how much new text appeared.
    """
    source, output = Path(source), Path(output)
    if source.resolve() == output.resolve():
        raise FileExistsError("refusing in-place edit; choose a new output path")
    if output.exists():
        raise FileExistsError(f"refusing to overwrite {output}")
    if not source.exists():
        raise FileNotFoundError(f"input not found: {source}")

    for edit in edits:
        if not edit.get("find"):
            raise EditError("empty 'find' is not allowed")
        if edit["find"] == edit["replace"]:
            raise EditError("'find' and 'replace' are identical")

    doc = open_document(source)
    date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    before_regions = len(region_text(doc))

    per_edit: list[tuple[str, int]] = []
    receipts: list[dict] = []
    misses: list[str] = []
    total = 0

    for edit in edits:
        needle, replacement = edit["find"], edit["replace"]
        texts_before = {label: paragraph_text(p) for label, p in iter_regions(doc)}
        old_before = _counts(texts_before, needle, normalize)
        new_before = _counts(texts_before, replacement, normalize)

        found_by_region: dict[str, int] = {}
        found = 0
        for label, paragraph in iter_regions(doc):
            occurrences = count_occurrences(paragraph_text(paragraph), needle, normalize)
            for _ in range(occurrences):
                if not _replace_once(paragraph, needle, replacement, author, date, normalize):
                    raise EditError(f"match disappeared while editing {label}")
                found += 1
                found_by_region[label] = found_by_region.get(label, 0) + 1
        total += found
        per_edit.append((needle, found))
        if found == 0:
            misses.append(needle)

        texts_after = {label: paragraph_text(p) for label, p in iter_regions(doc)}
        old_texts_after = {
            label: paragraph_text_without_insertions(p) for label, p in iter_regions(doc)
        }
        old_after = _counts(old_texts_after, needle, normalize)
        new_after = _counts(texts_after, replacement, normalize)

        regions = {}
        for label, count in found_by_region.items():
            regions[label] = {
                "found": count,
                "old_left": old_after.get(label, 0),
                "new_added": new_after.get(label, 0) - new_before.get(label, 0),
            }
        for label, left in old_after.items():
            if left:
                regions.setdefault(label, {"found": 0, "old_left": left, "new_added": 0})
        receipts.append({"find": needle, "replace": replacement, "found": found, "regions": regions})

    if misses:
        raise EditError("not found: " + ", ".join(repr(m) for m in misses))
    if total == 0:
        raise EditError("no changes produced")

    for receipt in receipts:
        for label, stats in receipt["regions"].items():
            if stats["old_left"]:
                raise EditError(
                    f"verification failed: {stats['old_left']} visible occurrence(s) of "
                    f"{receipt['find']!r} remain in {label}"
                )

    if not dry_run:
        output.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output))
        saved = {
            label: paragraph_text_without_insertions(p)
            for label, p in iter_regions(open_document(output))
        }
        for receipt in receipts:
            leftover = sum(_counts(saved, receipt["find"], normalize).values())
            if leftover:
                raise EditError(
                    f"verification failed after save: {leftover} occurrence(s) of "
                    f"{receipt['find']!r} remain"
                )

    return {
        "replacements": total,
        "per_edit": per_edit,
        "receipts": receipts,
        "dry_run": dry_run,
        "regions_before": before_regions,
    }

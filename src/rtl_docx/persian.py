"""Persian normalization and RTL hygiene checks."""

from __future__ import annotations

import re
import unicodedata

from docx.oxml.ns import qn

from .common import iter_regions, paragraph_text

PERSIAN_RANGE = "\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF"
PERSIAN_RE = re.compile(f"[{PERSIAN_RANGE}]")
CODE_STYLES = ("SourceCode", "Source Code")

# One-to-one mappings only: safe for offset-preserving search/replace.
VARIANTS = str.maketrans({
    "\u064A": "\u06CC",  # Arabic yeh -> Persian yeh
    "\u0649": "\u06CC",  # alef maksura -> Persian yeh
    "\u0643": "\u06A9",  # Arabic kaf -> Persian keheh
    "\u0660": "0", "\u0661": "1", "\u0662": "2", "\u0663": "3", "\u0664": "4",
    "\u0665": "5", "\u0666": "6", "\u0667": "7", "\u0668": "8", "\u0669": "9",
    "\u06F0": "0", "\u06F1": "1", "\u06F2": "2", "\u06F3": "3", "\u06F4": "4",
    "\u06F5": "5", "\u06F6": "6", "\u06F7": "7", "\u06F8": "8", "\u06F9": "9",
})

BIDI_CONTROLS = "\u200e\u200f\u061c\u2066\u2067\u2068\u2069"
ARABIC_YEH = "\u064A"
ARABIC_KAF = "\u0643"
ARABIC_DIGITS = "\u0660-\u0669"


def map_variants(text: str) -> str:
    """Apply length-preserving Persian/Arabic variant mapping."""
    return text.translate(VARIANTS)


def normalize_for_search(text: str, ignore_zwnj: bool = False) -> str:
    text = map_variants(text)
    text = unicodedata.normalize("NFC", text)
    text = "".join(ch for ch in text if ch not in BIDI_CONTROLS)
    if ignore_zwnj:
        text = text.replace("\u200c", "").replace("\u200d", "")
    return text


def count_occurrences(
    haystack: str,
    needle: str,
    normalize: bool = False,
    ignore_zwnj: bool = False,
) -> int:
    if normalize:
        haystack = normalize_for_search(haystack, ignore_zwnj)
        needle = normalize_for_search(needle, ignore_zwnj)
    if not needle:
        return 0
    return haystack.count(needle)


def _style_rtl(style) -> bool:
    seen = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        ppr = style.element.find(qn("w:pPr"))
        if ppr is not None:
            bidi = ppr.find(qn("w:bidi"))
            if bidi is not None:
                return (bidi.get(qn("w:val")) or "1") not in ("0", "false", "off")
        style = style.base_style
    return False


def paragraph_rtl(paragraph) -> bool:
    ppr = paragraph._p.find(qn("w:pPr"))
    if ppr is not None:
        bidi = ppr.find(qn("w:bidi"))
        if bidi is not None:
            return (bidi.get(qn("w:val")) or "1") not in ("0", "false", "off")
    return _style_rtl(paragraph.style)


def _is_code(paragraph) -> bool:
    style = paragraph.style
    return style is not None and style.style_id in CODE_STYLES


def _add(found: dict, kind: str, region: str, sample: str) -> None:
    entry = found.setdefault(kind, {"count": 0, "regions": set(), "samples": []})
    entry["count"] += 1
    entry["regions"].add(region)
    if len(entry["samples"]) < 3:
        entry["samples"].append(sample)


def lint_docx(path) -> dict:
    from .common import open_document

    doc = open_document(path)
    found: dict = {}

    for region, paragraph in iter_regions(doc):
        text = paragraph_text(paragraph)
        if not text.strip():
            continue
        has_persian = bool(PERSIAN_RE.search(text))
        code = _is_code(paragraph)
        rtl = paragraph_rtl(paragraph)
        persian_words = re.findall(f"[{PERSIAN_RANGE}]+", text)
        substantial = len(persian_words) >= 2 and not text.lstrip().startswith("|")

        if has_persian and not rtl and substantial:
            _add(found, "direction_mismatch", region, text[:80])

        for run_el in paragraph._p.iter(qn("w:r")) if substantial else ():
            run_text = "".join(t.text or "" for t in run_el.findall(qn("w:t")))
            if not PERSIAN_RE.search(run_text):
                continue
            rpr = run_el.find(qn("w:rPr"))
            rtl_el = rpr.find(qn("w:rtl")) if rpr is not None else None
            if rtl_el is not None:
                val = rtl_el.get(qn("w:val")) or "1"
                if val in ("0", "false", "off"):
                    _add(found, "run_direction_mismatch", region, run_text[:80])

        if code:
            continue

        for ch, kind in ((ARABIC_YEH, "arabic_yeh"), (ARABIC_KAF, "arabic_kaf")):
            if ch in text:
                for _ in range(text.count(ch)):
                    _add(found, kind, region, text[:80])

        digits = re.findall(f"[{ARABIC_DIGITS}]", text)
        for digit in digits:
            _add(found, "arabic_indic_digits", region, digit)

        for match in re.finditer(f"(?<=[{PERSIAN_RANGE}])\\s*[,;?]", text):
            _add(found, "latin_punctuation", region, text[max(0, match.start() - 20):match.end() + 20])

        for _ in range(sum(text.count(c) for c in BIDI_CONTROLS)):
            _add(found, "bidi_controls", region, text[:80])

        if substantial and not text.lstrip().startswith(("|", "│", "└", "├", "┌")):
            for _ in range(text.count("  ")):
                _add(found, "double_space", region, text[:80])

    findings = []
    for kind, entry in sorted(found.items(), key=lambda kv: (-kv[1]["count"], kv[0])):
        findings.append({
            "kind": kind,
            "count": entry["count"],
            "regions": sorted(entry["regions"]),
            "samples": entry["samples"],
        })
    return {"file": str(path), "findings": findings, "total": sum(f["count"] for f in findings)}

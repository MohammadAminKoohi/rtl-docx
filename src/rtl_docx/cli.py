"""rtl-docx command line interface."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

from . import __version__


def _err(msg: str) -> None:
    print(msg, file=sys.stderr)


def cmd_build(args) -> int:
    from .build import build_docx

    try:
        out = build_docx(args.source, args.output, args.reference, args.pandoc_arg)
    except (FileNotFoundError, FileExistsError, ValueError, RuntimeError) as exc:
        _err(f"error: {exc}")
        return 2
    except subprocess.CalledProcessError as exc:
        _err(f"error: pandoc failed with exit code {exc.returncode}")
        return 1
    print(f"built: {out}")
    print(f"next: rtl-docx verify {out}")
    return 0


def cmd_reference(args) -> int:
    from .build import export_reference

    try:
        out = export_reference(args.output)
    except FileExistsError as exc:
        _err(f"error: {exc}")
        return 2
    print(f"wrote: {out}")
    return 0


def cmd_verify(args) -> int:
    from .verify import VerifyError, verify_docx

    try:
        result = verify_docx(args.file, dpi=args.dpi, timeout=args.timeout)
    except (FileNotFoundError, VerifyError) as exc:
        _err(f"FAIL: {exc}")
        return 1
    print(f"ok: {args.file} -> {result['pdf']} ({len(result['pages'])} pages)")
    print("inspect these PNGs with vision before declaring the task done:")
    for page in result["pages"]:
        print(f"  {page}")
    return 0


def cmd_inventory(args) -> int:
    from .guard import inventory

    try:
        data = inventory(args.file)
    except FileNotFoundError as exc:
        _err(f"error: {exc}")
        return 2
    print(f"file            {args.file}")
    for key, value in data.items():
        print(f"{key:<15} {value}")
    return 0


def cmd_find(args) -> int:
    from .guard import CheckFailed, assert_count, count_matches, group_matches

    try:
        hits = count_matches(args.file, args.text, args.regex)
    except FileNotFoundError as exc:
        _err(f"error: {exc}")
        return 2
    total = sum(hits.values())
    print(f"search: {args.text!r} in {args.file}")
    for group, n in group_matches(hits).items():
        print(f"  {group:<40} {n}")
    print(f"total: {total}")
    try:
        assert_count(total, exact=args.exact, minimum=args.min, allow_zero=args.allow_zero)
    except CheckFailed as exc:
        _err(f"FAIL: {exc}")
        return 1
    return 0


def cmd_dump(args) -> int:
    from .guard import dump_text

    try:
        payload = dump_text(args.file)
    except FileNotFoundError as exc:
        _err(f"error: {exc}")
        return 2
    if args.out:
        Path(args.out).write_text(payload, encoding="utf-8")
        print(f"wrote {args.out}")
    else:
        print(payload)
    return 0


def cmd_diff(args) -> int:
    from .guard import diff_docs

    try:
        payload, changed = diff_docs(args.a, args.b)
    except FileNotFoundError as exc:
        _err(f"error: {exc}")
        return 2
    if args.out:
        Path(args.out).write_text(payload, encoding="utf-8")
        print(f"wrote {args.out}")
    else:
        print(payload)
    print(f"regions changed: {changed}")
    return 0 if changed == 0 else 1


def cmd_track(args) -> int:
    from .track import EditError, apply_edits

    edits = [{"find": old, "replace": new} for old, new in args.replace]
    if args.edits:
        data = json.loads(Path(args.edits).read_text(encoding="utf-8"))
        if not isinstance(data, list):
            _err("error: edits JSON must be a list of {find, replace}")
            return 2
        edits.extend(data)
    if not edits:
        _err("error: provide --replace or --edits")
        return 2

    try:
        result = apply_edits(args.input, args.output, edits, args.author, args.dry_run)
    except (FileNotFoundError, FileExistsError, EditError) as exc:
        _err(f"error: {exc}")
        return 1 if isinstance(exc, EditError) else 2

    for needle, count in result["per_edit"]:
        verb = "would replace" if result["dry_run"] else "replaced"
        print(f"{verb} {count}x: {needle!r}")
    if result["dry_run"]:
        print(f"\ndry run: {result['replacements']} change(s); nothing written")
    else:
        print(
            f"\nwrote {args.output} ({result['replacements']} tracked change(s) "
            f"by {args.author!r})"
        )
    return 0


def cmd_fill(args) -> int:
    from .fill import FillError, fill_template

    try:
        out = fill_template(args.template, args.data, args.output)
    except (FileNotFoundError, FileExistsError, FillError) as exc:
        _err(f"error: {exc}")
        return 1
    print(f"filled: {out}")
    return 0


def cmd_read(args) -> int:
    from .read import ReadError, read_markdown

    try:
        text = read_markdown(args.file)
    except (FileNotFoundError, ReadError) as exc:
        _err(f"error: {exc}")
        return 2
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
        print(f"wrote {args.out}")
    else:
        print(text)
    return 0


def cmd_doctor(_args) -> int:
    from .build import DEFAULT_REFERENCE
    from .verify import VerifyError, find_soffice

    def first_line(cmd: list[str]) -> str:
        try:
            out = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
            return (out.stdout or out.stderr).splitlines()[0].strip()
        except Exception as exc:  # noqa: BLE001
            return f"not available ({exc})"

    print(f"rtl-docx        {__version__}")
    print(f"python          {sys.version.split()[0]}")
    print(f"reference.docx  {'ok' if DEFAULT_REFERENCE.exists() else 'MISSING'}")
    pandoc = shutil.which("pandoc")
    print(f"pandoc          {first_line([pandoc, '--version']) if pandoc else 'NOT FOUND'}")
    try:
        print(f"soffice         {find_soffice()}")
    except VerifyError as exc:
        print(f"soffice         NOT FOUND ({exc})")
    pdftoppm = shutil.which("pdftoppm")
    print(f"pdftoppm        {'ok' if pdftoppm else 'NOT FOUND'}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="rtl-docx",
        description="RTL-first Word workflow: build from Markdown, tracked edits, "
                    "template fills, region-complete verification.",
        epilog="Docs: https://github.com/MohammadAminKoohi/rtl-docx",
    )
    parser.add_argument("--version", action="version", version=f"rtl-docx {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("build", help="Markdown -> styled .docx (never overwrites)")
    p.add_argument("source", type=Path)
    p.add_argument("output", type=Path)
    p.add_argument("--reference", type=Path, help="custom reference.docx")
    p.add_argument("--pandoc-arg", action="append", default=[], help="extra pandoc argument")
    p.set_defaults(func=cmd_build)

    p = sub.add_parser("reference", help="export the bundled RTL reference template")
    p.add_argument("-o", "--output", type=Path, default=Path("reference.docx"))
    p.set_defaults(func=cmd_reference)

    p = sub.add_parser("verify", help="LibreOffice PDF + PNG pages for visual review")
    p.add_argument("file", type=Path)
    p.add_argument("--dpi", type=int, default=150)
    p.add_argument("--timeout", type=int, default=120)
    p.set_defaults(func=cmd_verify)

    p = sub.add_parser("inventory", help="structural counts")
    p.add_argument("file", type=Path)
    p.set_defaults(func=cmd_inventory)

    p = sub.add_parser("find", help="count text across body, tables, headers, footers")
    p.add_argument("file", type=Path)
    p.add_argument("text")
    p.add_argument("--regex", action="store_true")
    p.add_argument("--exact", type=int)
    p.add_argument("--min", type=int)
    p.add_argument("--allow-zero", action="store_true")
    p.set_defaults(func=cmd_find)

    p = sub.add_parser("dump", help="region-labelled text for round-trip diffs")
    p.add_argument("file", type=Path)
    p.add_argument("-o", "--out")
    p.set_defaults(func=cmd_dump)

    p = sub.add_parser("diff", help="unified text diff of two .docx files")
    p.add_argument("a", type=Path)
    p.add_argument("b", type=Path)
    p.add_argument("-o", "--out")
    p.set_defaults(func=cmd_diff)

    p = sub.add_parser("track", help="find/replace as Word tracked changes")
    p.add_argument("input", type=Path)
    p.add_argument("output", type=Path)
    p.add_argument("--author", default="Codex")
    p.add_argument("--replace", nargs=2, action="append", metavar=("OLD", "NEW"), default=[])
    p.add_argument("--edits", type=Path, help="JSON list of {find, replace}")
    p.add_argument("--dry-run", action="store_true")
    p.set_defaults(func=cmd_track)

    p = sub.add_parser("fill", help="render a docxtpl template with JSON")
    p.add_argument("template", type=Path)
    p.add_argument("data", type=Path)
    p.add_argument("output", type=Path)
    p.set_defaults(func=cmd_fill)

    p = sub.add_parser("read", help="convert to Markdown with MarkItDown")
    p.add_argument("file", type=Path)
    p.add_argument("-o", "--out")
    p.set_defaults(func=cmd_read)

    p = sub.add_parser("doctor", help="check external dependencies")
    p.set_defaults(func=cmd_doctor)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())

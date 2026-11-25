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


def _print_json(data) -> None:
    print(json.dumps(data, ensure_ascii=False, indent=2, default=str))


def cmd_build(args) -> int:
    from .build import build_docx

    try:
        out = build_docx(
            args.source, args.output, args.reference, args.pandoc_arg,
            toc=args.toc, toc_depth=args.toc_depth,
            number_sections=args.number_sections,
            page_numbers=args.page_numbers, header=args.header,
            footer=args.footer, page_label=args.page_label, align=args.align,
        )
    except (FileNotFoundError, FileExistsError, ValueError, RuntimeError) as exc:
        _err(f"error: {exc}")
        return 2
    except subprocess.CalledProcessError as exc:
        _err(f"error: pandoc failed with exit code {exc.returncode}")
        return 1
    print(f"built: {out}")
    print(f"next: rtl-docx validate {out}")
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


def cmd_validate(args) -> int:
    from .validate import validate_docx

    try:
        result = validate_docx(args.file, schema=not args.no_schema, max_errors=args.max_errors)
    except FileNotFoundError as exc:
        _err(f"error: {exc}")
        return 2

    if args.json:
        _print_json(result)
    else:
        print(f"{'valid' if result['valid'] else 'INVALID'}: {result['file']}")
        print(
            "schema checks: "
            + ("run" if result["schema_checked"] else "skipped (pip install openxml-audit)")
        )
        for warning in result["warnings"]:
            print(f"warning: {warning}")
        for error in result["errors"]:
            print(f"error: {error}")
    return 0 if result["valid"] else 1


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
        hits = count_matches(
            args.file, args.text, args.regex,
            normalize=args.normalize, ignore_zwnj=args.ignore_zwnj,
        )
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
        result = apply_edits(
            args.input, args.output, edits, args.author,
            args.dry_run, normalize=args.normalize,
        )
    except (FileNotFoundError, FileExistsError, EditError) as exc:
        _err(f"error: {exc}")
        return 1 if isinstance(exc, EditError) else 2

    if args.json:
        _print_json(result)
        return 0

    for needle, count in result["per_edit"]:
        verb = "would replace" if result["dry_run"] else "replaced"
        print(f"{verb} {count}x: {needle!r}")
    for receipt in result["receipts"]:
        touched = len(receipt["regions"])
        print(f"  regions touched: {touched}")
        for label, stats in receipt["regions"].items():
            print(
                f"    {label}: found={stats['found']} "
                f"old_left={stats['old_left']} new_added={stats['new_added']}"
            )
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


def cmd_compare(args) -> int:
    from .revisions import RevisionError, compare_documents

    try:
        result = compare_documents(
            args.original, args.modified, args.output, args.author,
            detect_moves=args.detect_moves, case_insensitive=args.case_insensitive,
        )
    except (FileNotFoundError, FileExistsError, RevisionError) as exc:
        _err(f"error: {exc}")
        return 1
    print(f"redline: {result['output']} ({result['revisions']} revisions by {args.author!r})")
    if result["engine_stderr"]:
        _err(result["engine_stderr"])
    return 0


def cmd_revisions(args) -> int:
    from .revisions import list_revisions, revision_summary

    try:
        items = list_revisions(args.file)
    except FileNotFoundError as exc:
        _err(f"error: {exc}")
        return 2
    if args.json:
        _print_json(items)
        return 0
    summary = revision_summary(args.file)
    print(f"revisions: {summary['count']}  types={summary['by_type']}  authors={summary['by_author']}")
    for item in items:
        print(f"  [{item['type']}] {item['author']} {item['date'] or ''}: {item['text'][:80]!r}")
    return 0


def cmd_accept(args) -> int:
    from .revisions import RevisionError, accept_changes

    try:
        out = accept_changes(args.input, args.output)
    except (FileNotFoundError, FileExistsError, RevisionError) as exc:
        _err(f"error: {exc}")
        return 1
    print(f"accepted: {out}")
    return 0


def cmd_reject(args) -> int:
    from .revisions import RevisionError, reject_changes

    try:
        out = reject_changes(args.input, args.output)
    except (FileNotFoundError, FileExistsError, RevisionError) as exc:
        _err(f"error: {exc}")
        return 1
    print(f"rejected: {out}")
    return 0


def cmd_audit(args) -> int:
    from .revisions import RevisionError, audit_changes

    try:
        result = audit_changes(args.file, args.original, args.author)
    except (FileNotFoundError, RevisionError) as exc:
        _err(f"error: {exc}")
        return 2

    if args.json:
        _print_json(result)
    else:
        print(f"revisions: {result['revisions']}  authors={result['authors']}")
        if result["author_mismatches"]:
            print(f"author mismatch: {result['author_mismatches']}")
        if result["clean"]:
            print("clean: rejecting all revisions reproduces the original")
        else:
            print(f"UNTRACKED CHANGES: {result['untracked_count']} differing line(s)")
            print(result["untracked_diff"])
    return 0 if result["clean"] else 1


def cmd_comments_list(args) -> int:
    from .comments import list_comments

    try:
        items = list_comments(args.file)
    except FileNotFoundError as exc:
        _err(f"error: {exc}")
        return 2
    if args.json:
        _print_json(items)
        return 0
    for item in items:
        print(f"#{item['id']} {item['author']} {item['date'] or ''}")
        print(f"  anchor: {item['anchor'][:100]!r}")
        print(f"  text:   {item['text'][:200]!r}")
    print(f"total: {len(items)}")
    return 0


def cmd_comments_add(args) -> int:
    from .comments import CommentError, add_comment

    try:
        result = add_comment(
            args.input, args.output, args.anchor, args.text,
            args.author, args.initials, args.occurrence,
        )
    except (FileNotFoundError, FileExistsError, CommentError) as exc:
        _err(f"error: {exc}")
        return 1
    print(f"comment #{result['comment_id']} added at {result['region']}: {result['output']}")
    return 0


def cmd_check(args) -> int:
    from .check import CheckError, check_document

    try:
        result = check_document(
            args.docx, args.markdown,
            max_report=args.max_report, min_ratio=args.min_ratio,
        )
    except (FileNotFoundError, CheckError) as exc:
        _err(f"error: {exc}")
        return 2

    if args.json:
        _print_json(result)
    else:
        status = "clean" if result["clean"] else ("PASS" if result["pass"] else "FAIL")
        print(
            f"{status}: ratio={result['ratio']} (min {result['min_ratio']}) "
            f"missing={result['missing_lines']} extra={result['extra_lines']}"
        )
        if not result["clean"]:
            print(result["diff"])
    if args.strict:
        return 0 if result["clean"] else 1
    return 0 if result["pass"] else 1


def cmd_lint(args) -> int:
    from .persian import lint_docx

    try:
        result = lint_docx(args.file)
    except FileNotFoundError as exc:
        _err(f"error: {exc}")
        return 2

    if args.json:
        _print_json(result)
    else:
        print(f"{args.file}: {result['total']} finding(s)")
        for finding in result["findings"]:
            print(f"  {finding['kind']}: {finding['count']}")
            print(f"    regions: {', '.join(finding['regions'][:4])}"
                  + (" ..." if len(finding["regions"]) > 4 else ""))
            for sample in finding["samples"]:
                print(f"    e.g. {sample!r}")
    return 1 if args.strict and result["total"] else 0


def cmd_convert(args) -> int:
    from .convert import ConvertError, convert_document

    to = args.to
    if to is None:
        to = args.output.suffix.lstrip(".") if args.output and args.output.suffix else "docx"
    try:
        out = convert_document(args.source, args.output, to=to, timeout=args.timeout)
    except (FileNotFoundError, FileExistsError, ConvertError) as exc:
        _err(f"error: {exc}")
        return 1
    print(f"converted: {out}")
    return 0


def cmd_merge(args) -> int:
    from .merge import MergeError, merge_documents

    try:
        out = merge_documents(
            args.inputs, args.output, page_break=not args.no_page_break
        )
    except (FileNotFoundError, FileExistsError, MergeError) as exc:
        _err(f"error: {exc}")
        return 1
    print(f"merged: {out}")
    return 0


def cmd_stamp(args) -> int:
    from .stamp import StampError, stamp_copy

    if not (args.header or args.footer or args.page_numbers):
        _err("error: provide --header, --footer or --page-numbers")
        return 2
    try:
        out = stamp_copy(
            args.input, args.output,
            header=args.header, footer=args.footer,
            page_numbers=args.page_numbers, page_label=args.page_label,
            align=args.align,
        )
    except (FileNotFoundError, FileExistsError, StampError) as exc:
        _err(f"error: {exc}")
        return 1
    print(f"stamped: {out}")
    return 0


def cmd_doctor(_args) -> int:
    from .build import DEFAULT_REFERENCE
    from .verify import VerifyError, find_soffice

    def first_line(cmd: list[str]) -> str:
        try:
            out = subprocess.run(cmd, capture_output=True, text=True, timeout=10, check=False)
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
    try:
        import openxml_audit  # noqa: F401
        print("schema checks   ok (openxml-audit)")
    except ImportError:
        print("schema checks   MISSING (pip install openxml-audit)")
    try:
        import python_redlines  # noqa: F401
        print("redline engine  ok (python-redlines)")
    except ImportError:
        print('redline engine  optional: uv tool install "rtl-docx[redline]"')
    try:
        import docxcompose  # noqa: F401
        print("merge support   ok (docxcompose)")
    except ImportError:
        print('merge support   optional: uv tool install "rtl-docx[merge]"')
    try:
        import mcp  # noqa: F401
        print("mcp server      ok (mcp)")
    except ImportError:
        print('mcp server      optional: uv tool install "rtl-docx[mcp]"')
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="rtl-docx",
        description="RTL-first Word workflow: build from Markdown, tracked edits, "
                    "redlines, template fills, region-complete verification.",
        epilog="Docs: https://github.com/MohammadAminKoohi/rtl-docx",
    )
    parser.add_argument("--version", action="version", version=f"rtl-docx {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("build", help="Markdown -> styled .docx (never overwrites)")
    p.add_argument("source", type=Path)
    p.add_argument("output", type=Path)
    p.add_argument("--reference", type=Path, help="custom reference.docx")
    p.add_argument("--pandoc-arg", action="append", default=[], help="extra pandoc argument")
    p.add_argument("--toc", action="store_true", help="insert a table of contents")
    p.add_argument("--toc-depth", type=int)
    p.add_argument("--number-sections", action="store_true")
    p.add_argument("--page-numbers", action="store_true")
    p.add_argument("--header")
    p.add_argument("--footer")
    p.add_argument("--page-label", default="")
    p.add_argument("--align", choices=["left", "center", "right"], default="center")
    p.set_defaults(func=cmd_build)

    p = sub.add_parser("reference", help="export the bundled RTL reference template")
    p.add_argument("-o", "--output", type=Path, default=Path("reference.docx"))
    p.set_defaults(func=cmd_reference)

    p = sub.add_parser("verify", help="LibreOffice PDF + PNG pages for visual review")
    p.add_argument("file", type=Path)
    p.add_argument("--dpi", type=int, default=150)
    p.add_argument("--timeout", type=int, default=120)
    p.set_defaults(func=cmd_verify)

    p = sub.add_parser("validate", help="OOXML package/schema validation")
    p.add_argument("file", type=Path)
    p.add_argument("--no-schema", action="store_true")
    p.add_argument("--json", action="store_true")
    p.add_argument("--max-errors", type=int, default=50)
    p.set_defaults(func=cmd_validate)

    p = sub.add_parser("inventory", help="structural counts")
    p.add_argument("file", type=Path)
    p.set_defaults(func=cmd_inventory)

    p = sub.add_parser("find", help="count text across body, tables, headers, footers")
    p.add_argument("file", type=Path)
    p.add_argument("text")
    p.add_argument("--regex", action="store_true")
    p.add_argument("--normalize-persian", dest="normalize", action="store_true",
                   help="match Arabic/Persian yeh, kaf and digits variants")
    p.add_argument("--ignore-zwnj", action="store_true")
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
    p.add_argument("--normalize-persian", dest="normalize", action="store_true")
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--json", action="store_true", help="print per-region receipts as JSON")
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

    p = sub.add_parser("compare", help="redline two .docx versions as tracked changes")
    p.add_argument("original", type=Path)
    p.add_argument("modified", type=Path)
    p.add_argument("output", type=Path)
    p.add_argument("--author", default="Codex")
    p.add_argument("--detect-moves", action="store_true")
    p.add_argument("--case-insensitive", action="store_true")
    p.set_defaults(func=cmd_compare)

    p = sub.add_parser("revisions", help="list tracked changes")
    p.add_argument("file", type=Path)
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_revisions)

    p = sub.add_parser("accept", help="accept all tracked changes into a new file")
    p.add_argument("input", type=Path)
    p.add_argument("output", type=Path)
    p.set_defaults(func=cmd_accept)

    p = sub.add_parser("reject", help="reject all tracked changes into a new file")
    p.add_argument("input", type=Path)
    p.add_argument("output", type=Path)
    p.set_defaults(func=cmd_reject)

    p = sub.add_parser("audit", help="verify every change in a redline is tracked")
    p.add_argument("file", type=Path)
    p.add_argument("--original", type=Path, required=True)
    p.add_argument("--author", help="require this author on every revision")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_audit)

    p = sub.add_parser("comments", help="list or add Word comments")
    comments_sub = p.add_subparsers(dest="comments_command", required=True)

    c = comments_sub.add_parser("list")
    c.add_argument("file", type=Path)
    c.add_argument("--json", action="store_true")
    c.set_defaults(func=cmd_comments_list)

    c = comments_sub.add_parser("add")
    c.add_argument("input", type=Path)
    c.add_argument("output", type=Path)
    c.add_argument("--anchor", required=True, help="existing text to attach the comment to")
    c.add_argument("--text", required=True)
    c.add_argument("--author", default="Codex")
    c.add_argument("--initials", default="")
    c.add_argument("--occurrence", type=int, default=1)
    c.set_defaults(func=cmd_comments_add)

    p = sub.add_parser("check", help="round-trip text check of a .docx against Markdown")
    p.add_argument("docx", type=Path)
    p.add_argument("markdown", type=Path)
    p.add_argument("--max-report", type=int, default=40)
    p.add_argument("--min-ratio", type=float, default=0.95)
    p.add_argument("--strict", action="store_true", help="require an exact match")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_check)

    p = sub.add_parser("lint", help="Persian/RTL hygiene findings")
    p.add_argument("file", type=Path)
    p.add_argument("--json", action="store_true")
    p.add_argument("--strict", action="store_true", help="exit 1 when findings exist")
    p.set_defaults(func=cmd_lint)

    p = sub.add_parser("convert", help="convert via LibreOffice (doc->docx, docx->pdf)")
    p.add_argument("source", type=Path)
    p.add_argument("output", type=Path, nargs="?")
    p.add_argument("--to", choices=["docx", "pdf", "odt", "rtf", "txt"])
    p.add_argument("--timeout", type=int, default=180)
    p.set_defaults(func=cmd_convert)

    p = sub.add_parser("merge", help="merge several .docx files, preserving styles")
    p.add_argument("inputs", type=Path, nargs="+")
    p.add_argument("-o", "--output", type=Path, required=True)
    p.add_argument("--no-page-break", action="store_true")
    p.set_defaults(func=cmd_merge)

    p = sub.add_parser("stamp", help="add header/footer/page numbers to every section")
    p.add_argument("input", type=Path)
    p.add_argument("output", type=Path)
    p.add_argument("--header")
    p.add_argument("--footer")
    p.add_argument("--page-numbers", action="store_true")
    p.add_argument("--page-label", default="")
    p.add_argument("--align", choices=["left", "center", "right"], default="center")
    p.set_defaults(func=cmd_stamp)

    p = sub.add_parser("doctor", help="check external dependencies")
    p.set_defaults(func=cmd_doctor)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())

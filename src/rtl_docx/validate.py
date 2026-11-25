"""Validate .docx packages: structure, relationships, content types, schema."""

from __future__ import annotations

import json
import posixpath
import zipfile
from pathlib import Path

from lxml import etree

CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

REQUIRED_PARTS = ("[Content_Types].xml", "_rels/.rels", "word/document.xml")
MAX_UNCOMPRESSED = 500 * 1024 * 1024
MAX_RATIO = 200


def _structural(path: Path) -> tuple[list[str], list[str]]:
    errors: list[str] = []
    warnings: list[str] = []

    if not zipfile.is_zipfile(path):
        return ["not a ZIP/OOXML package"], warnings

    with zipfile.ZipFile(path) as package:
        bad = package.testzip()
        if bad:
            errors.append(f"corrupt ZIP member: {bad}")

        infos = package.infolist()
        total = sum(i.file_size for i in infos)
        if total > MAX_UNCOMPRESSED:
            errors.append(f"package expands to {total} bytes (limit {MAX_UNCOMPRESSED})")
        for info in infos:
            if info.compress_size and info.file_size / info.compress_size > MAX_RATIO:
                errors.append(
                    f"suspicious compression ratio for {info.filename} "
                    f"({info.file_size}/{info.compress_size})"
                )
                break

        names = {i.filename for i in infos}
        for required in REQUIRED_PARTS:
            if required not in names:
                errors.append(f"missing required part: {required}")
        if "word/vbaProject.bin" in names:
            warnings.append("document contains VBA macros (word/vbaProject.bin)")

        # XML well-formedness
        for name in sorted(names):
            if name.endswith((".xml", ".rels")):
                try:
                    etree.fromstring(package.read(name))
                except etree.XMLSyntaxError as exc:
                    errors.append(f"malformed XML in {name}: {exc}")

        # content types cover every part, overrides point at real parts
        if "[Content_Types].xml" in names:
            try:
                ct = etree.fromstring(package.read("[Content_Types].xml"))
                defaults = {
                    (el.get("Extension") or "").lower()
                    for el in ct.findall(f"{CT_NS}Default")
                }
                overrides = {
                    el.get("PartName") for el in ct.findall(f"{CT_NS}Override")
                }
                for name in names:
                    if name == "[Content_Types].xml" or name.endswith("/"):
                        continue
                    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
                    if f"/{name}" not in overrides and ext not in defaults:
                        errors.append(f"part has no content type: {name}")
                for part in overrides:
                    if part and part.lstrip("/") not in names:
                        errors.append(f"content-type override targets missing part: {part}")
            except etree.XMLSyntaxError:
                pass  # already reported above

        # relationships resolve to existing internal parts
        for name in sorted(n for n in names if n.endswith(".rels")):
            try:
                rels = etree.fromstring(package.read(name))
            except etree.XMLSyntaxError:
                continue
            source_dir = posixpath.dirname(posixpath.dirname(name))
            for rel in rels.findall(f"{REL_NS}Relationship"):
                if (rel.get("TargetMode") or "").lower() == "external":
                    continue
                target = rel.get("Target") or ""
                if not target or target.startswith("#"):
                    continue
                resolved = posixpath.normpath(posixpath.join(source_dir, target))
                if resolved not in names:
                    errors.append(
                        f"{name}: relationship targets missing part {resolved}"
                    )

    return errors, warnings


def _schema(path: Path, max_errors: int) -> tuple[list[str], bool]:
    try:
        from openxml_audit import OpenXmlValidator
    except ImportError:
        return [], False

    result = OpenXmlValidator(max_errors=max_errors).validate(str(path))
    issues = [
        f"{issue.severity.value.upper()} {issue.part_uri or ''}: {issue.description}"
        for issue in result.errors
    ]
    return issues, True


def validate_docx(
    path: str | Path,
    schema: bool = True,
    max_errors: int = 50,
) -> dict:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"not found: {path}")

    errors, warnings = _structural(path)
    schema_checked = False
    if schema and not errors:
        schema_issues, schema_checked = _schema(path, max_errors)
        errors.extend(schema_issues)

    return {
        "file": str(path),
        "valid": not errors,
        "errors": errors,
        "warnings": warnings,
        "schema_checked": schema_checked,
    }


def main(argv: list[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Validate a .docx package")
    parser.add_argument("file", type=Path)
    parser.add_argument("--no-schema", action="store_true", help="skip OOXML schema checks")
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--max-errors", type=int, default=50)
    args = parser.parse_args(argv)

    result = validate_docx(args.file, schema=not args.no_schema, max_errors=args.max_errors)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"{'valid' if result['valid'] else 'INVALID'}: {result['file']}")
        print(f"schema checks: {'run' if result['schema_checked'] else 'skipped (pip install openxml-audit)'}")
        for warning in result["warnings"]:
            print(f"warning: {warning}")
        for error in result["errors"]:
            print(f"error: {error}")
    return 0 if result["valid"] else 1


if __name__ == "__main__":
    raise SystemExit(main())

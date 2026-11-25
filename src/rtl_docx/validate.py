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


def sanitize_package(path: str | Path) -> int:
    """Remove relationships and content-type overrides that point at missing parts.

    Some editors leave dangling references; Word tolerates them, strict
    validators and other consumers do not.
    """
    path = Path(path)
    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        payloads = {name: zin.read(name) for name in names}

    removed = 0
    for name in list(payloads):
        if not name.endswith(".rels"):
            continue
        try:
            root = etree.fromstring(payloads[name])
        except etree.XMLSyntaxError:
            continue
        source_dir = posixpath.dirname(posixpath.dirname(name))
        changed = False
        for rel in list(root.findall(f"{REL_NS}Relationship")):
            if (rel.get("TargetMode") or "").lower() == "external":
                continue
            target = rel.get("Target") or ""
            if not target or target.startswith("#"):
                continue
            resolved = posixpath.normpath(posixpath.join(source_dir, target))
            if resolved not in names:
                root.remove(rel)
                removed += 1
                changed = True
        if changed:
            payloads[name] = etree.tostring(root, xml_declaration=True,
                                            encoding="UTF-8", standalone=True)

    if "[Content_Types].xml" in payloads:
        try:
            root = etree.fromstring(payloads["[Content_Types].xml"])
        except etree.XMLSyntaxError:
            root = None
        if root is not None:
            changed = False
            for override in list(root.findall(f"{CT_NS}Override")):
                part = (override.get("PartName") or "").lstrip("/")
                if part and part not in names:
                    root.remove(override)
                    removed += 1
                    changed = True
            if changed:
                payloads["[Content_Types].xml"] = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )

    if removed:
        import os
        tmp = path.with_name(path.stem + ".tmp.docx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, payload in payloads.items():
                zout.writestr(name, payload)
        os.replace(tmp, path)
    return removed


REL_TYPES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
def _theme_payload() -> bytes:
    with zipfile.ZipFile(Path(__file__).parent / "assets" / "reference.docx") as package:
        return package.read("word/theme/theme1.xml")


REQUIRED_REL_PARTS = {
    "theme": (
        "word/theme/theme1.xml",
        "application/vnd.openxmlformats-officedocument.theme+xml",
        _theme_payload(),
    ),
    "settings": (
        "word/settings.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
    ),
    "fontTable": (
        "word/fontTable.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
    ),
}


def repair_required_parts(path: str | Path) -> int:
    """Restore settings/fontTable parts when a producer left dangling refs."""
    path = Path(path)
    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        payloads = {name: zin.read(name) for name in names}

    rels_name = "word/_rels/document.xml.rels"
    if rels_name not in payloads:
        return 0
    rels = etree.fromstring(payloads[rels_name])
    existing_types = {rel.get("Type") for rel in rels.findall(f"{REL_NS}Relationship")}
    ct = etree.fromstring(payloads["[Content_Types].xml"])
    overrides = {el.get("PartName") for el in ct.findall(f"{CT_NS}Override")}

    repaired = 0
    for rel_type, (part, content_type, payload) in REQUIRED_REL_PARTS.items():
        type_uri = REL_TYPES + rel_type
        if type_uri in existing_types and part in payloads:
            continue
        if part not in payloads:
            payloads[part] = payload
            repaired += 1
        if type_uri not in existing_types:
            rel = etree.SubElement(rels, f"{REL_NS}Relationship")
            rel.set("Id", f"rIdRtl{rel_type.capitalize()}")
            rel.set("Type", type_uri)
            rel.set("Target", posixpath.relpath(part, "word"))
            repaired += 1
        if f"/{part}" not in overrides:
            override = etree.SubElement(ct, f"{CT_NS}Override")
            override.set("PartName", f"/{part}")
            override.set("ContentType", content_type)
            repaired += 1

    if repaired:
        payloads[rels_name] = etree.tostring(rels, xml_declaration=True,
                                             encoding="UTF-8", standalone=True)
        payloads["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True,
                                                         encoding="UTF-8", standalone=True)
        import os
        tmp = path.with_name(path.stem + ".tmp.docx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, payload in payloads.items():
                zout.writestr(name, payload)
        os.replace(tmp, path)
    return repaired

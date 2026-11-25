from __future__ import annotations

import json
import shutil

import pytest
from docx import Document
from docx.oxml.ns import qn

from rtl_docx.build import DEFAULT_REFERENCE, export_reference
from rtl_docx.fill import FillError, fill_template
from rtl_docx.guard import (
    CheckFailed,
    assert_count,
    count_matches,
    diff_docs,
    dump_text,
    inventory,
)
from rtl_docx.track import EditError, apply_edits


@pytest.fixture
def regions_docx(tmp_path):
    doc = Document()
    doc.add_paragraph("body OLD value")
    table = doc.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "cell OLD value"
    table.cell(0, 1).text = "other"
    doc.sections[0].header.paragraphs[0].text = "header OLD value"
    doc.sections[0].footer.paragraphs[0].text = "footer OLD value"
    path = tmp_path / "regions.docx"
    doc.save(path)
    return path


def test_inventory_counts(regions_docx):
    data = inventory(regions_docx)
    assert data["sections"] == 1
    assert data["tables"] == 1
    assert data["headers"] == 1
    assert data["footers"] == 1
    assert data["words"] > 0


def test_count_covers_every_region(regions_docx):
    hits = count_matches(regions_docx, "OLD")
    assert sum(hits.values()) == 4
    labels = " ".join(hits)
    assert "header" in labels and "footer" in labels and "table" in labels


def test_assert_count():
    with pytest.raises(CheckFailed):
        assert_count(2, exact=3, minimum=None, allow_zero=False)
    with pytest.raises(CheckFailed):
        assert_count(0, exact=None, minimum=1, allow_zero=False)
    assert_count(0, exact=None, minimum=None, allow_zero=True)


def test_track_edits_all_regions(regions_docx, tmp_path):
    out = tmp_path / "out.docx"
    result = apply_edits(regions_docx, out, [{"find": "OLD", "replace": "NEW"}], author="Tester")
    assert result["replacements"] == 4
    assert out.exists()

    assert sum(count_matches(out, "NEW").values()) == 4
    assert sum(count_matches(out, "OLD").values()) == 0

    xml = Document(out).paragraphs[0]._p.xml
    assert "<w:ins" in xml and "<w:del" in xml
    assert 'w:author="Tester"' in xml


def test_track_keeps_source_untouched(regions_docx, tmp_path):
    before = regions_docx.read_bytes()
    apply_edits(regions_docx, tmp_path / "out.docx", [{"find": "OLD", "replace": "NEW"}])
    assert regions_docx.read_bytes() == before


def test_track_refuses_overwrite(regions_docx):
    with pytest.raises(FileExistsError):
        apply_edits(regions_docx, regions_docx, [{"find": "OLD", "replace": "NEW"}])


def test_track_missing_text_raises(regions_docx, tmp_path):
    with pytest.raises(EditError):
        apply_edits(regions_docx, tmp_path / "out.docx", [{"find": "absent", "replace": "x"}])


def test_track_cross_run_match(tmp_path):
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("The quick ")
    p.add_run("brown fox ")
    p.add_run("jumps over the dog.")
    src = tmp_path / "runs.docx"
    doc.save(src)

    out = tmp_path / "runs.out.docx"
    apply_edits(src, out, [{"find": "quick brown fox jumps", "replace": "slow red fox leaps"}])
    text = dump_text(out)
    assert "slow red fox leaps over the dog." in text

    xml = Document(out).paragraphs[0]._p.xml
    assert "<w:ins" in xml and "<w:del" in xml


def test_fill_template(tmp_path):
    doc = Document()
    doc.add_paragraph("Hello {{ name }}, total {{ total }}")
    template = tmp_path / "template.docx"
    doc.save(template)

    data = tmp_path / "data.json"
    data.write_text(json.dumps({"name": "Amin", "total": "1250"}), encoding="utf-8")

    out = tmp_path / "filled.docx"
    fill_template(template, data, out)
    text = dump_text(out)
    assert "Hello Amin, total 1250" in text
    assert "{{" not in text


def test_fill_missing_key_fails(tmp_path):
    doc = Document()
    doc.add_paragraph("{{ missing }}")
    template = tmp_path / "t.docx"
    doc.save(template)
    data = tmp_path / "d.json"
    data.write_text("{}", encoding="utf-8")
    with pytest.raises(FillError):
        fill_template(template, data, tmp_path / "o.docx")


def test_diff_reports_change(tmp_path):
    doc = Document()
    doc.add_paragraph("one")
    a = tmp_path / "a.docx"
    doc.save(a)

    doc = Document()
    doc.add_paragraph("two")
    b = tmp_path / "b.docx"
    doc.save(b)

    payload, changed = diff_docs(a, b)
    assert changed == 1
    assert "-one" in payload and "+two" in payload


def test_reference_export(tmp_path):
    assert DEFAULT_REFERENCE.exists()
    out = export_reference(tmp_path / "ref.docx")
    assert out.exists() and out.stat().st_size > 1000


@pytest.mark.skipif(not shutil.which("pandoc"), reason="pandoc not installed")
def test_build_rtl_docx(tmp_path):
    md = tmp_path / "doc.md"
    md.write_text("# عنوان\n\nمتن فارسی.\n\n```python\nx = 1\n```\n", encoding="utf-8")
    out = tmp_path / "doc.docx"

    from rtl_docx.build import build_docx
    from rtl_docx.validate import validate_docx

    build_docx(md, out)
    doc = Document(out)
    assert any("عنوان" in p.text for p in doc.paragraphs)

    code = next(p for p in doc.paragraphs if "x = 1" in p.text)
    ppr = code._p.find(qn("w:pPr"))
    assert ppr.find(qn("w:bidi")).get(qn("w:val")) == "0"

    # Run-level direction is valid; style-level w:rtl is not allowed by OOXML.
    code_run = code._p.find(qn("w:r"))
    assert code_run.find(qn("w:rPr")).find(qn("w:rtl")).get(qn("w:val")) == "0"
    body = next(p for p in doc.paragraphs if "متن فارسی" in p.text)
    body_run = body._p.find(qn("w:r"))
    assert body_run.find(qn("w:rPr")).find(qn("w:rtl")).get(qn("w:val")) == "1"

    normal = doc.styles["Normal"].element
    assert normal.find(qn("w:rPr")).find(qn("w:rtl")) is None

    result = validate_docx(out)
    assert result["schema_checked"], "openxml-audit should be installed for tests"
    assert result["valid"], result["errors"]


def test_reference_is_schema_valid():
    from rtl_docx.validate import validate_docx

    result = validate_docx(DEFAULT_REFERENCE)
    assert result["schema_checked"]
    assert result["valid"], result["errors"]


def test_validate_rejects_broken_package(tmp_path, regions_docx):
    from rtl_docx.validate import validate_docx

    broken = tmp_path / "broken.docx"
    broken.write_bytes(regions_docx.read_bytes()[:200])
    result = validate_docx(broken)
    assert not result["valid"]
    assert result["errors"]


def test_track_receipts_cover_regions(regions_docx, tmp_path):
    out = tmp_path / "receipts.docx"
    result = apply_edits(regions_docx, out, [{"find": "OLD", "replace": "NEW"}])
    receipt = result["receipts"][0]
    assert receipt["found"] == 4
    labels = " ".join(receipt["regions"])
    assert "header" in labels and "footer" in labels and "table" in labels
    assert all(stats["old_left"] == 0 for stats in receipt["regions"].values())
    assert sum(stats["found"] for stats in receipt["regions"].values()) == 4


def test_track_normalizes_persian_variants(tmp_path):
    doc = Document()
    doc.add_paragraph("علي كريم ١٢٣")
    src = tmp_path / "fa.docx"
    doc.save(src)

    out = tmp_path / "fa_out.docx"
    apply_edits(src, out, [{"find": "علی کریم 123", "replace": "علی کریم 456"}], normalize=True)
    text = dump_text(out)
    assert "456" in text
    assert "123" not in text


def test_find_normalizes_persian_variants(tmp_path):
    from rtl_docx.guard import count_matches

    doc = Document()
    doc.add_paragraph("علي كريم ١٢٣")
    src = tmp_path / "fa.docx"
    doc.save(src)

    hits = count_matches(src, "علی کریم 123", normalize=True)
    assert sum(hits.values()) == 1


def test_comments_add_and_list(regions_docx, tmp_path):
    from rtl_docx.comments import add_comment, list_comments

    out = tmp_path / "commented.docx"
    result = add_comment(
        regions_docx, out, "body OLD", "این متن باید بازنویسی شود",
        author="Amin", initials="A",
    )
    assert result["comment_id"] == 0

    comments = list_comments(out)
    assert len(comments) == 1
    assert comments[0]["text"] == "این متن باید بازنویسی شود"
    assert "body OLD" in comments[0]["anchor"]

    from rtl_docx.validate import validate_docx
    assert validate_docx(out)["valid"]


def test_revisions_accept_reject_audit(regions_docx, tmp_path):
    from rtl_docx.revisions import (
        accept_changes,
        audit_changes,
        list_revisions,
        reject_changes,
    )

    tracked = tmp_path / "tracked.docx"
    apply_edits(regions_docx, tracked, [{"find": "OLD", "replace": "NEW"}])

    items = list_revisions(tracked)
    assert len(items) >= 4
    assert {item["type"] for item in items} == {"insertion", "deletion"}

    accepted = tmp_path / "accepted.docx"
    accept_changes(tracked, accepted)
    accepted_text = dump_text(accepted)
    assert "NEW" in accepted_text and "OLD" not in accepted_text

    rejected = tmp_path / "rejected.docx"
    reject_changes(tracked, rejected)
    rejected_text = dump_text(rejected)
    assert "OLD" in rejected_text and "NEW" not in rejected_text

    audit = audit_changes(tracked, regions_docx)
    assert audit["clean"], audit["untracked_diff"]

    untracked = tmp_path / "untracked.docx"
    doc = Document(regions_docx)
    doc.paragraphs[0].text = "body untracked"
    doc.save(untracked)
    bad = audit_changes(untracked, regions_docx)
    assert not bad["clean"]


def test_compare_documents(tmp_path):
    from rtl_docx.revisions import compare_documents, count_revisions_xml

    a = Document()
    a.add_paragraph("قرارداد نسخه اول است.")
    old = tmp_path / "old.docx"
    a.save(old)

    b = Document()
    b.add_paragraph("قرارداد نسخه دوم است.")
    new = tmp_path / "new.docx"
    b.save(new)

    out = tmp_path / "redline.docx"
    result = compare_documents(old, new, out, author="Amin")
    assert out.exists()
    assert count_revisions_xml(out) > 0
    assert result["revisions"] > 0

    from rtl_docx.validate import validate_docx
    assert validate_docx(out)["valid"]


def test_lint_persian_issues(tmp_path):
    from rtl_docx.persian import lint_docx

    doc = Document()
    doc.add_paragraph("علي كريم, ١٢٣")
    src = tmp_path / "lint.docx"
    doc.save(src)

    result = lint_docx(src)
    kinds = {finding["kind"] for finding in result["findings"]}
    assert "arabic_yeh" in kinds
    assert "arabic_kaf" in kinds
    assert "arabic_indic_digits" in kinds
    assert "latin_punctuation" in kinds


@pytest.mark.skipif(not shutil.which("pandoc"), reason="pandoc not installed")
def test_roundtrip_check(tmp_path):
    from rtl_docx.build import build_docx
    from rtl_docx.check import check_document

    md = tmp_path / "doc.md"
    md.write_text("# عنوان\n\nمتن فارسی برای بررسی.\n", encoding="utf-8")
    out = tmp_path / "doc.docx"
    build_docx(md, out)

    result = check_document(out, md)
    assert result["clean"], result["diff"]


def test_mcp_server(regions_docx):
    import asyncio
    import json

    from rtl_docx.mcp_server import mcp

    async def run():
        names = {tool.name for tool in await mcp.list_tools()}
        assert {"document_inventory", "find_text", "track_edits",
                "compare_documents", "validate_document"} <= names

        result = await mcp.call_tool(
            "document_inventory", {"path": str(regions_docx)}
        )
        data = json.loads(result.content[0].text)
        assert data["tables"] == 1 and data["headers"] == 1

        result = await mcp.call_tool(
            "find_text", {"path": str(regions_docx), "text": "OLD"}
        )
        data = json.loads(result.content[0].text)
        assert data["total"] == 4

    asyncio.run(run())


def test_merge_documents(tmp_path):
    from rtl_docx.merge import merge_documents
    from rtl_docx.validate import validate_docx

    paths = []
    for index, text in enumerate(("سند اول", "سند دوم")):
        doc = Document()
        doc.add_paragraph(text)
        path = tmp_path / f"part{index}.docx"
        doc.save(path)
        paths.append(path)

    out = tmp_path / "merged.docx"
    merge_documents(paths, out)
    text = dump_text(out)
    assert "سند اول" in text and "سند دوم" in text
    assert validate_docx(out)["valid"]


def test_stamp_header_footer_page_numbers(tmp_path):
    from rtl_docx.stamp import stamp_copy
    from rtl_docx.validate import validate_docx

    doc = Document()
    doc.add_paragraph("متن")
    source = tmp_path / "in.docx"
    doc.save(source)

    out = tmp_path / "stamped.docx"
    stamp_copy(
        source, out, header="سربرگ", footer="پانویس",
        page_numbers=True, page_label="صفحه ",
    )
    stamped = Document(out)
    assert stamped.sections[0].header.paragraphs[0].text == "سربرگ"
    footer_xml = stamped.sections[0].footer.paragraphs[0]._p.xml
    assert "PAGE" in footer_xml
    assert validate_docx(out)["valid"]


@pytest.mark.skipif(not shutil.which("pandoc"), reason="pandoc not installed")
def test_build_toc_and_page_numbers(tmp_path):
    from rtl_docx.build import build_docx
    from rtl_docx.validate import validate_docx

    md = tmp_path / "report.md"
    md.write_text(
        "# عنوان\n\n## بخش یک\n\nمتن.\n\n## بخش دو\n\nمتن بیشتر.\n",
        encoding="utf-8",
    )
    out = tmp_path / "report.docx"
    build_docx(
        md, out, toc=True, page_numbers=True,
        header="سربرگ", footer="پانویس",
    )
    doc = Document(out)
    assert doc.sections[0].header.paragraphs[0].text == "سربرگ"
    assert "PAGE" in doc.sections[0].footer.paragraphs[0]._p.xml
    assert validate_docx(out)["valid"]


@pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice not installed")
def test_convert_docx_to_pdf(tmp_path):
    from rtl_docx.convert import convert_document

    doc = Document()
    doc.add_paragraph("PDF")
    source = tmp_path / "in.docx"
    doc.save(source)
    out = convert_document(source, tmp_path / "out.pdf", to="pdf", timeout=90)
    assert out.exists() and out.stat().st_size > 0


def test_validate_rejects_path_traversal(regions_docx, tmp_path):
    import zipfile

    from rtl_docx.validate import validate_docx

    target = tmp_path / "evil.docx"
    with zipfile.ZipFile(regions_docx) as zin, zipfile.ZipFile(target, "w") as zout:
        for info in zin.infolist():
            zout.writestr(info, zin.read(info.filename))
        zout.writestr("../evil.txt", "x")

    result = validate_docx(target, schema=False)
    assert not result["valid"]
    assert any("unsafe path" in error for error in result["errors"])


def test_validate_warns_on_external_relationships(regions_docx, tmp_path):
    import zipfile

    from lxml import etree

    from rtl_docx.validate import REL_NS, validate_docx

    target = tmp_path / "external.docx"
    with zipfile.ZipFile(regions_docx) as zin:
        payloads = {info.filename: zin.read(info.filename) for info in zin.infolist()}

    rels_name = "word/_rels/document.xml.rels"
    root = etree.fromstring(payloads[rels_name])
    rel = etree.SubElement(root, f"{REL_NS}Relationship")
    rel.set("Id", "rIdExternal")
    rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
    rel.set("Target", "https://example.com")
    rel.set("TargetMode", "External")
    payloads[rels_name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8")

    with zipfile.ZipFile(target, "w") as zout:
        for name, payload in payloads.items():
            zout.writestr(name, payload)

    result = validate_docx(target, schema=False)
    assert any("external relationship" in warning for warning in result["warnings"])


def test_plugin_manifest_and_skill_layout():
    import json
    from pathlib import Path

    root = Path(__file__).resolve().parents[1]
    manifest = json.loads((root / ".codex-plugin" / "plugin.json").read_text())
    assert manifest["name"] == "rtl-docx"
    assert manifest["version"] == "0.3.0"
    assert (root / "skills" / "rtl-docx" / "SKILL.md").exists()
    assert manifest["skills"] == "./skills/"

    mcp_config = json.loads((root / ".mcp.json").read_text())
    assert mcp_config["mcpServers"]["rtl-docx"]["command"] == "rtl-docx-mcp"

    skill = (root / "skills" / "rtl-docx" / "SKILL.md").read_text()
    assert skill.startswith("---\nname: rtl-docx\n")
    assert "version: \"0.3.0\"" in skill

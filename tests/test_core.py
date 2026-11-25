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

    build_docx(md, out)
    doc = Document(out)
    assert any("عنوان" in p.text for p in doc.paragraphs)

    code = [p for p in doc.paragraphs if "x = 1" in p.text][0]
    ppr = code._p.find(qn("w:pPr"))
    assert ppr.find(qn("w:bidi")).get(qn("w:val")) == "0"

    normal = doc.styles["Normal"].element.find(qn("w:pPr")).find(qn("w:bidi"))
    assert normal is not None

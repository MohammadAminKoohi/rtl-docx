---
name: rtl-docx
description: Create, read, edit, fill, and visually verify Word .docx files, especially Persian/Farsi/Arabic/Hebrew right-to-left documents. Use for any Word task: Markdown-to-DOCX builds with RTL styling, tracked changes on existing files, Jinja2 template filling, text extraction, region-complete search/replace checks, and LibreOffice render verification. Trigger on .docx, Word, Persian, RTL, tracked changes, redline, template, or report generation.
license: MIT
---

# RTL DOCX

A deterministic CLI plus this skill's rules. The agent writes Markdown/JSON;
`rtl-docx` produces and verifies the Word file. Never hand-write OOXML, never
edit a `.docx` zip, never "fix" a document with ad-hoc python-docx scripts.

## 0. Check the tool

```bash
rtl-docx doctor
```

If `rtl-docx` is missing, install it once:

```bash
uv tool install --editable /path/to/rtl-docx     # local checkout
# or, after publishing:
uv tool install "git+https://github.com/OWNER/rtl-docx"
```

If `pandoc`, `soffice` or `pdftoppm` are missing:

```bash
brew install pandoc librsvg poppler && brew install --cask libreoffice
```

`build` needs pandoc; `verify` needs LibreOffice + poppler. `read`, `find`,
`dump`, `diff`, `track`, `fill` work without either.

## 1. Route the task

| Task | Command |
| --- | --- |
| Read / analyze / summarize | `rtl-docx read FILE.docx` |
| New report or doc | write `.md`, then `rtl-docx build IN.md OUT.docx` |
| Fill a fixed template | `rtl-docx fill TEMPLATE.docx DATA.json OUT.docx` |
| Edit an existing human doc | `rtl-docx track IN.docx OUT.docx --author "Name" --replace OLD NEW` |
| Structural report | `rtl-docx inventory FILE.docx` |
| Verify text everywhere | `rtl-docx find FILE.docx "TEXT" --exact N` |
| Round-trip check | `rtl-docx dump` + `rtl-docx diff` |
| Render for visual review | `rtl-docx verify FILE.docx` |

## 2. Hard rules

1. **Persian/RTL is the default.** The bundled reference is RTL-first:
   `w:bidi` sections, `fa-IR` paragraphs, right-aligned headings, centred
   titles, RTL tables. Persian text needs no `dir` wrappers. Latin identifiers,
   URLs and code stay LTR *inside* the RTL document; never flip the whole
   document to LTR and never wrap Persian in `<div dir="ltr">`.
2. **Never overwrite.** Build/track/fill refuse existing outputs. Write to a new
   path (`report_v2.docx`). Do not delete and rebuild a human's file.
3. **Never edit a human-authored document in place.** Use `track` so every
   change is a native `w:ins`/`w:del` revision with an author, reviewable in
   Word.
4. **Verify every region.** `find` counts body paragraphs, every table cell,
   headers and footers. Assert with `--exact` or `--min`; never check only
   `doc.paragraphs`.
5. **Round-trip every generated file.** Extract the text back (`dump`/`read`)
   and diff it against intent before reporting success.
6. **Render before declaring done.** If LibreOffice fails, the XML is broken;
   if it times out, retry once with `--timeout 300` before calling it corrupt.

## 3. Build a document

Write Markdown as the source of truth, then:

```bash
rtl-docx build report.md output/report.docx
rtl-docx verify output/report.docx
```

- Reference styles come from the packaged `reference.docx`. Export and restyle
  it in Word when needed: `rtl-docx reference -o templates/reference.docx`,
  then `--reference templates/reference.docx`.
- Keep images next to the `.md` (relative paths); pandoc embeds them. SVG needs
  `rsvg-convert` (librsvg).
- Fenced code and inline code are forced LTR automatically.

## 4. Edit an existing document (tracked changes)

```bash
rtl-docx track IN.docx OUT.docx --author "Codex" --replace "old" "new" --dry-run
rtl-docx track IN.docx OUT.docx --author "Codex" --replace "old" "new"
rtl-docx find OUT.docx "new" --exact N
rtl-docx find OUT.docx "old" --allow-zero
```

- Edits apply across body, table cells, headers and footers.
- Each match must lie inside one paragraph. Split multi-paragraph rewrites into
  separate edits.
- Runs containing breaks, tabs, fields, footnotes or drawings are rejected. Do
  not bypass that: normalise the paragraph first.
- Batch edits: `--edits edits.json` with `[{"find": "...", "replace": "..."}]`.

## 5. Fill a template

```bash
rtl-docx fill templates/invoice.docx invoice.json output/invoice_0042.docx
```

- Placeholders are Jinja2: `{{ customer.name }}`, `{%tr for item in items %}`.
- The agent produces **JSON only**; styling stays in Word.
- Missing keys fail loudly (StrictUndefined); unresolved `{{` markers in the
  saved package are a hard failure.

## 6. Visual verification

```bash
rtl-docx verify FILE.docx            # 150-dpi PNGs under .docx-verify/<name>/
```

Read the PNGs with vision and check for:

- tables overflowing margins, clipped cells, invisible text
- awkward page breaks, orphaned headings, split tables
- broken list numbering / bullets
- Persian punctuation and mixed LTR/RTL runs (parentheses, digits, code)
- missing, stretched or oversized images

If vision is unavailable, say so explicitly and report the machine checks
instead of claiming a visual review.

## 7. Done checklist

- [ ] Output is a new path; the input is untouched.
- [ ] `find --exact` passes across body, tables, headers, footers.
- [ ] `diff`/`dump` shows only intended changes.
- [ ] `verify` succeeds; every page PNG was inspected.
- [ ] Report page count, regions changed, and warnings (font substitution, etc.).

---
name: rtl-docx
description: Create, read, edit, fill, and visually verify Word .docx files, especially Persian/Farsi/Arabic/Hebrew right-to-left documents. Use for any Word task: Markdown-to-DOCX builds with RTL styling, tracked changes on existing files, Jinja2 template filling, text extraction, region-complete search/replace checks, and LibreOffice render verification. Trigger on .docx, Word, Persian, RTL, tracked changes, redline, template, or report generation.
license: MIT
compatibility: Requires Python 3.11+ and uv. pandoc is needed for builds; LibreOffice (soffice) and poppler (pdftoppm) for render verification.
metadata:
  author: Mohammad Amin Koohi
  version: "0.3.0"
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
uv tool install "git+https://github.com/MohammadAminKoohi/rtl-docx"

# optional extras:
uv tool install "rtl-docx[redline] @ git+https://github.com/MohammadAminKoohi/rtl-docx"  # compare/audit
uv tool install "rtl-docx[mcp] @ git+https://github.com/MohammadAminKoohi/rtl-docx"      # MCP server
uv tool install "rtl-docx[merge] @ git+https://github.com/MohammadAminKoohi/rtl-docx"    # merge
uv tool install "rtl-docx[all] @ git+https://github.com/MohammadAminKoohi/rtl-docx"      # everything
```

`compare` needs `redline`; `merge` needs `merge`; the MCP server needs `mcp`. The
core install covers build, track, comments, validate, check, lint, convert and
stamp.

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
| Compare two versions | `rtl-docx compare OLD.docx NEW.docx REDLINE.docx --author "Name"` |
| List / accept / reject revisions | `revisions`, `accept`, `reject` |
| Audit a redline | `rtl-docx audit REDLINE.docx --original OLD.docx` |
| Comments | `rtl-docx comments list/add` |
| Structural report | `rtl-docx inventory FILE.docx` |
| OOXML validity | `rtl-docx validate FILE.docx` |
| Verify text everywhere | `rtl-docx find FILE.docx "TEXT" --exact N` |
| Persian variants / lint | `--normalize-persian`, `rtl-docx lint` |
| Round-trip build check | `rtl-docx check OUT.docx IN.md` |
| Convert legacy / PDF | `rtl-docx convert IN.doc OUT.docx` |
| Merge documents | `rtl-docx merge A.docx B.docx -o OUT.docx` |
| Header/footer/page numbers | `rtl-docx stamp IN.docx OUT.docx --page-numbers` |
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
4. **Verify every region.** `track` returns per-region receipts (body, every
   table cell, headers, footers); `find` asserts counts with `--exact`/`--min`.
   Never check only `doc.paragraphs`.
5. **Round-trip every generated file.** `check OUT.docx IN.md` normalizes
   Persian variants and diffs the text; use `dump`/`diff` for DOCX-to-DOCX.
6. **Validate, then render.** `validate` catches schema/relationship errors that
   LibreOffice silently tolerates; `verify` catches layout problems. If
   LibreOffice fails, the XML is broken; if it times out, retry once with
   `--timeout 300` before calling it corrupt.

## 3. Build a document

Write Markdown as the source of truth, then:

```bash
rtl-docx build report.md output/report.docx
rtl-docx validate output/report.docx
rtl-docx verify output/report.docx
```

- Reference styles come from the packaged `reference.docx`. Export and restyle
  it in Word when needed: `rtl-docx reference -o templates/reference.docx`,
  then `--reference templates/reference.docx`.
- Keep images next to the `.md` (relative paths); pandoc embeds them. SVG needs
  `rsvg-convert` (librsvg).
- Fenced code and inline code are forced LTR automatically.
- Report options: `--toc`, `--toc-depth N`, `--number-sections`,
  `--page-numbers`, `--header "..."`, `--footer "..."`, `--page-label "صفحه "`.
  Page-number fields are real `PAGE` fields, centred by default.
- `stamp IN OUT --header ... --footer ... --page-numbers` adds them to an
  existing document.

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

## 5. Compare two versions (native redline)

```bash
rtl-docx compare old.docx new.docx redline.docx --author "Amin"
rtl-docx audit redline.docx --original old.docx --author "Amin"
rtl-docx validate redline.docx
```

- `compare` produces real Word tracked changes (structure-aware: tables, headers,
  formatting; `--detect-moves` optional). It needs the `redline` extra.
- `audit` rejects every revision and compares with the original: any remaining
  difference is an untracked edit. Non-zero exit means do not ship the redline.
- `revisions`, `accept`, `reject` complete the review loop:

```bash
rtl-docx revisions redline.docx
rtl-docx accept redline.docx clean.docx      # accepted view
rtl-docx reject redline.docx original_view.docx
```

## 6. Comments

```bash
rtl-docx comments list FILE.docx
rtl-docx comments add IN.docx OUT.docx --anchor "متن مورد نظر" \
  --text "این بخش بازبینی شود" --author "Amin"
```

- The anchor must lie inside one paragraph; `--occurrence N` picks a later match.
- Comments are written natively; output is a new file and structurally validated.
- Read the anchor in the result — `comments list` shows both anchor and comment.

## 7. Fill a template

```bash
rtl-docx fill templates/invoice.docx invoice.json output/invoice_0042.docx
```

- Placeholders are Jinja2: `{{ customer.name }}`, `{%tr for item in items %}`.
- The agent produces **JSON only**; styling stays in Word.
- Missing keys fail loudly (StrictUndefined); unresolved `{{` markers in the
  saved package are a hard failure.

## 8. Convert, merge, stamp

```bash
rtl-docx convert legacy.doc modern.docx          # LibreOffice conversion
rtl-docx convert report.docx report.pdf --to pdf
rtl-docx merge cover.docx body.docx -o report.docx
rtl-docx stamp report.docx final.docx --header "سربرگ" --footer "پانویس" --page-numbers
```

- Convert/merge/stamp write new files and validate the package structurally.
- Merge preserves styles, numbering and headers via docxcompose; a page break
  separates documents unless `--no-page-break`.
- `stamp` works on every section and emits a real `PAGE` field.

## 9. MCP server (optional)

If the client does not run shell commands, expose the same operations over MCP:

```bash
uv tool install "rtl-docx[mcp] @ git+https://github.com/MohammadAminKoohi/rtl-docx"
rtl-docx-mcp
```

Tools include read, inventory, find, validate, lint, build, track, compare,
revisions, accept/reject, audit, comments, check, fill, verify, convert, merge
and stamp. The bundled plugin manifest (`.codex-plugin/plugin.json`) points at
`skills/` and `.mcp.json`.

## 10. Round-trip and RTL lint

```bash
rtl-docx check OUT.docx IN.md          # PASS at ratio >= 0.95; --strict exact
rtl-docx lint FILE.docx                # --strict to fail on findings
```

`check` extracts expected text with pandoc, normalizes Persian variants and
bidi controls, and diffs against the DOCX text. Small table/footnote rendering
differences are reported but still PASS; inspect the diff, do not ignore it.
`lint` flags Arabic yeh/kaf, Arabic-Indic digits, Latin punctuation in Persian
text, direction mismatches, bidi controls and double spaces.

## 11. Validate and visually verify

```bash
rtl-docx validate FILE.docx         # package, relationships, content types, XSD
rtl-docx verify FILE.docx           # 150-dpi PNGs under .docx-verify/<name>/
```

`validate` checks package structure, relationships, content types and OOXML
schema (pure Python, no .NET via `openxml-audit`). It should be clean before you
spend time looking at pages.

Read the PNGs with vision and check for:

- tables overflowing margins, clipped cells, invisible text
- awkward page breaks, orphaned headings, split tables
- broken list numbering / bullets
- Persian punctuation and mixed LTR/RTL runs (parentheses, digits, code)
- missing, stretched or oversized images

If vision is unavailable, say so explicitly and report the machine checks
instead of claiming a visual review.

## 12. Done checklist

- [ ] Output is a new path; the input is untouched.
- [ ] `validate` is clean.
- [ ] `track` receipts show the expected regions; no `old_left`.
- [ ] A redline passes `audit` with the intended author.
- [ ] `check` PASS (or `--strict` clean for text-only docs).
- [ ] `lint` findings reviewed; `--strict` clean when the document is prose.
- [ ] Converted/merged files validate structurally; merged text is complete.
- [ ] `verify` succeeds; every page PNG was inspected.
- [ ] Report page count, regions changed, and warnings (font substitution, etc.).

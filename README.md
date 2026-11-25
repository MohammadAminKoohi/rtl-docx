# rtl-docx

**A right-to-left-first Word workflow for coding agents.**

`rtl-docx` treats `.docx` as a build artifact: write Markdown or JSON, compile
with a styled Persian/Arabic reference template, propose edits as native tracked
changes, verify text across every region, and render the result before calling it
done.

Built for Persian documentation, where most Word tooling quietly gets it wrong:
RTL paragraphs, right-aligned headings, mirrored tables, Persian punctuation —
while code, URLs and identifiers stay LTR *inside* the RTL document.

```
Markdown / JSON  ->  rtl-docx  ->  styled .docx  ->  PDF + PNG pages  ->  vision check
```

## Why this exists

The usual failure mode with agents and Word files is "write a Python script that
mutates the XML." That produces corrupt packages, dropped formatting, and silent
edits nobody can review. This project encodes the workflows that actually hold
up:

| Task | Approach |
| --- | --- |
| Read / analyze | MarkItDown -> clean Markdown |
| Create | Markdown + pandoc + RTL reference template |
| Fill forms | docxtpl + validated JSON |
| Edit human docs | native `w:ins` / `w:del` tracked changes |
| Trust the result | region-complete counts, round-trip diff, LibreOffice render |

## Install

Requires Python 3.11+, [uv](https://docs.astral.sh/uv/), and for full
functionality `pandoc`, LibreOffice and poppler:

```bash
brew install pandoc librsvg poppler
brew install --cask libreoffice

uv tool install "git+https://github.com/MohammadAminKoohi/rtl-docx"
rtl-docx doctor

# extras, when you need them:
uv tool install "rtl-docx[redline] @ git+https://github.com/MohammadAminKoohi/rtl-docx"  # compare/audit
uv tool install "rtl-docx[merge] @ git+https://github.com/MohammadAminKoohi/rtl-docx"    # merge
uv tool install "rtl-docx[mcp] @ git+https://github.com/MohammadAminKoohi/rtl-docx"      # MCP server
uv tool install "rtl-docx[all] @ git+https://github.com/MohammadAminKoohi/rtl-docx"      # all three
```

Local development:

```bash
git clone https://github.com/MohammadAminKoohi/rtl-docx
cd rtl-docx
./install.sh          # CLI + user skill, editable install
./install.sh --with-redline
./install.sh --with-all
uv sync --dev
uv run pytest -q
```

## Use as an agent skill

Codex CLI and OpenCode both scan `~/.agents/skills`. This repo ships one under
`skills/rtl-docx`:

```bash
git clone https://github.com/MohammadAminKoohi/rtl-docx ~/.local/share/rtl-docx
ln -s ~/.local/share/rtl-docx/skills/rtl-docx ~/.agents/skills/rtl-docx
uv tool install --editable ~/.local/share/rtl-docx
```

Restart Codex if it does not pick the skill up. Invoke explicitly with
`/skills` or `$rtl-docx`, or just ask for a Persian `.docx` and let the
description match.

For a repo-scoped install, copy `skills/rtl-docx` into `.agents/skills/rtl-docx`
and commit it.

### MCP server and plugin

The repo doubles as a Codex plugin: `.codex-plugin/plugin.json` points at
`skills/` and `.mcp.json` launches `rtl-docx-mcp`.

```bash
uv tool install "rtl-docx[mcp] @ git+https://github.com/MohammadAminKoohi/rtl-docx"
rtl-docx-mcp            # stdio MCP server
```

The MCP surface mirrors the CLI: read, inventory, find, validate, lint, build,
track, compare, revisions, accept/reject, audit, comments, check, fill, verify,
convert, merge, stamp.

## Commands

```bash
rtl-docx read report.docx                    # MarkItDown -> Markdown
rtl-docx build report.md report.docx         # RTL-styled Word, never overwrites
rtl-docx reference -o templates/reference.docx

rtl-docx track in.docx out.docx --author "Amin" \
    --replace "متن قدیم" "متن جدید" --dry-run
rtl-docx track in.docx out.docx --author "Amin" --edits edits.json
rtl-docx track in.docx out.docx --author "Amin" --replace "علي" "علی" --normalize-persian

rtl-docx compare old.docx new.docx redline.docx --author "Amin"
rtl-docx audit redline.docx --original old.docx --author "Amin"
rtl-docx revisions redline.docx
rtl-docx accept redline.docx clean.docx
rtl-docx reject redline.docx original_view.docx

rtl-docx comments list report.docx
rtl-docx comments add in.docx out.docx --anchor "متن" --text "بازبینی" --author "Amin"

rtl-docx fill template.docx data.json out.docx

rtl-docx build report.md report.docx --toc --page-numbers --header "سربرگ"
rtl-docx convert legacy.doc modern.docx
rtl-docx convert report.docx report.pdf --to pdf
rtl-docx merge cover.docx body.docx -o report.docx
rtl-docx stamp in.docx out.docx --footer "پانویس" --page-numbers

rtl-docx inventory report.docx               # body/tables/headers/footers/images
rtl-docx validate report.docx                # package + OOXML schema checks
rtl-docx find report.docx "متن" --exact 3    # counts every region, asserts
rtl-docx lint report.docx                    # Persian/RTL hygiene
rtl-docx check report.docx report.md         # round-trip text check
rtl-docx dump report.docx -o text.txt
rtl-docx diff old.docx new.docx

rtl-docx verify report.docx                  # PDF + 150-dpi PNG pages
```

### Guardrails baked in

- **No overwrites.** `build`, `track`, `fill`, `compare`, `accept`, `reject`
  and `comments add` refuse an existing output path.
- **No in-place edits** of human documents; changes are reviewable revisions.
- **No region blind spots.** Search and `track` receipts cover body paragraphs,
  every table cell, headers and footers.
- **Redlines are audited.** `compare` produces the tracked changes; `audit`
  rejects them all and proves the result matches the original, so an untracked
  edit cannot slip through.
- **No unverified claims.** `check` diffs the built DOCX against its Markdown;
  `validate` checks OOXML schema; `verify` renders a real PDF and page PNGs.
- **Persian-aware matching.** `--normalize-persian` treats Arabic/Persian
  yeh/kaf and Arabic-Indic digits as equivalent; `lint` flags mixed direction,
  Arabic variants, Latin punctuation and bidi controls.
- **Untrusted-file checks.** `validate` rejects path traversal, DTD/entity
  declarations, oversized or zip-bomb packages, and warns on external
  relationships and macros. `read` enforces size and compression limits.

## The RTL details

`src/rtl_docx/assets/reference.docx` sets:

- A4, 2.5 cm margins, 1.5 line spacing, Arial
- `w:bidi` section + `w:lang w:bidi="fa-IR"` in styles
- right-aligned headings, centred titles, justified body
- `bidiVisual` tables
- right-aligned table direction and mirrored cell margins
- explicit **LTR** only for `Source Code` / `Verbatim Char`
- `w:rtl` stamped on every **run** after the build — never in style `rPr`,
  where OOXML does not allow it
- code paragraphs stamped `w:bidi=0` and runs `w:rtl=0`

Export it, restyle it in Word, and pass it back with `--reference`.

## Is this novel?

Honestly: **not algorithmically.** It stands on well-known tools — pandoc,
python-docx, docxtpl, MarkItDown, LibreOffice — and similar agent skills and MCP
servers already exist (Anthropic's docx skill, `docx-mcp`, `@superdoc/mcp`).

What is new is the packaging and the defaults:

1. An **RTL-first** reference pipeline, including the LTR-code-inside-RTL-document
   problem, which almost every existing docx tool ignores.
2. **Region-complete verification** (body + table cells + headers + footers) as a
   first-class command with assertions, not an afterthought.
3. A **tracked-change engine across every story part** (body, tables, headers,
   footers), with per-region receipts, native redlining between versions, and
   an `audit` that rejects revisions to prove nothing was edited untracked.
4. **Persian-aware matching and linting** (`--normalize-persian`, `lint`) on top
   of region-complete assertions.
5. **Packaging as a real agent surface**: a Codex/OpenCode skill, a plugin
   manifest, and an MCP server exposing the same deterministic operations.
6. A **single deterministic CLI** plus a skill that encodes the validate,
   round-trip and render-and-look loop, so agents cannot skip it.

That is enough for a genuinely useful skill. It is not a new file format or a
standards contribution — and the README should not pretend otherwise.

## Limitations

- `track` matches within one paragraph; matches spanning paragraphs must be
  split. Text boxes, footnotes and fields are not edited. `--normalize-persian`
  maps one-to-one variants only; ZWNJ-insensitive matching is for `find`.
- `compare` needs the optional `redline` extra (43 MB engine). Its output is
  structurally validated; run `validate` for schema checks. Sources that are
  already schema-invalid may produce a redline that inherits those errors.
- `accept`/`reject` handle run-level insertions/deletions/moves and paragraph
  marks. Exotic table-property or formatting-only revisions are not rewritten.
- `comments add` anchors within one paragraph; it comments the runs overlapping
  the anchor, which may highlight more than the exact characters.
- `check` passes at 0.95 similarity by default because tables and footnotes
  render differently; use `--strict` for exact comparison of text-only docs.
- `convert`, `merge` and `stamp` write new files; conversion fidelity is
  LibreOffice's, not Word's.
- Page numbers use a real `PAGE` field; LibreOffice updates them, Word may show
  the cached value until fields update (print/preview usually does).
- MCP tools run with the user's filesystem permissions; run the server only in
  trusted workspaces.
- Filling requires placeholders docxtpl can see (not split across runs).
- Visual fidelity depends on fonts available to LibreOffice; Word may substitute
  differently. Always eyeball the PNGs.
- Tested on macOS; Windows works if `soffice` is on `PATH` or `SOFFICE` is set.

## Releasing

The repository doubles as the skill.

**Trusted publishing (recommended, no token):** the `publish` workflow uploads
on every `v*` tag using PyPI trusted publishing. Configure a
[pending publisher](https://pypi.org/manage/account/publishing/) with:

- project: `rtl-docx`
- owner: `MohammadAminKoohi`
- repository: `rtl-docx`
- workflow name: `publish.yml`
- environment: `pypi`

Then:

```bash
git tag v0.1.0 && git push origin v0.1.0
```

**Local token upload:**

```bash
uv build
UV_PUBLISH_TOKEN=pypi-your-token uv publish
```

For one-click distribution to ChatGPT/Codex, wrap the same skill in a plugin;
the CLI stays the deterministic core.

## License

MIT

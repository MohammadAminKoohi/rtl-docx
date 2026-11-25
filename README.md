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

Requires Python 3.10+, [uv](https://docs.astral.sh/uv/), and for full
functionality `pandoc`, LibreOffice and poppler:

```bash
brew install pandoc librsvg poppler
brew install --cask libreoffice

uv tool install "git+https://github.com/OWNER/rtl-docx"
rtl-docx doctor
```

Local development:

```bash
git clone https://github.com/OWNER/rtl-docx
cd rtl-docx
./install.sh          # CLI + user skill, editable install
uv sync --dev
uv run pytest -q
```

## Use as an agent skill

Codex CLI and OpenCode both scan `~/.agents/skills`. A skill is just a folder
with `SKILL.md`; this repo is one:

```bash
git clone https://github.com/OWNER/rtl-docx ~/.local/share/rtl-docx
ln -s ~/.local/share/rtl-docx ~/.agents/skills/rtl-docx
uv tool install --editable ~/.local/share/rtl-docx
```

Restart Codex if it does not pick the skill up. Invoke explicitly with
`/skills` or `$rtl-docx`, or just ask for a Persian `.docx` and let the
description match.

For a repo-scoped install, copy the folder into `.agents/skills/rtl-docx` and
commit it.

## Commands

```bash
rtl-docx read report.docx                    # MarkItDown -> Markdown
rtl-docx build report.md report.docx         # RTL-styled Word, never overwrites
rtl-docx reference -o templates/reference.docx

rtl-docx track in.docx out.docx --author "Amin" \
    --replace "متن قدیم" "متن جدید" --dry-run
rtl-docx track in.docx out.docx --author "Amin" --edits edits.json

rtl-docx fill template.docx data.json out.docx

rtl-docx inventory report.docx               # body/tables/headers/footers/images
rtl-docx find report.docx "متن" --exact 3    # counts every region, asserts
rtl-docx dump report.docx -o text.txt
rtl-docx diff old.docx new.docx

rtl-docx verify report.docx                  # PDF + 150-dpi PNG pages
```

### Guardrails baked in

- **No overwrites.** `build`, `track` and `fill` refuse an existing output path.
- **No in-place edits** of human documents; changes are reviewable revisions.
- **No region blind spots.** Search counts body paragraphs, every table cell,
  headers and footers.
- **No unverified claims.** Round-trip text diff plus a real renderer, with a
  timeout instead of a hung LibreOffice process.

## The RTL details

`src/rtl_docx/assets/reference.docx` sets:

- A4, 2.5 cm margins, 1.5 line spacing, Arial
- `w:bidi` section + `w:rtl` runs with `w:lang w:bidi="fa-IR"`
- right-aligned headings, centred titles, justified body
- `bidiVisual` tables
- right-aligned table direction and mirrored cell margins
- explicit **LTR** only for `Source Code` / `Verbatim Char`
- code paragraphs stamped `w:bidi=0` and `w:rtl=0` after the build, because some
  renderers ignore a style-level override on an RTL section

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
3. A **tracked-change engine** that refuses unsafe runs instead of corrupting
   them, and verifies that the old text is no longer visible without accepting
   revisions.
4. A **single deterministic CLI** plus a skill that encodes the render-and-look
   loop, so agents cannot skip it.

That is enough for a genuinely useful skill. It is not a new file format or a
standards contribution — and the README should not pretend otherwise.

## Limitations

- `track` matches within one paragraph; matches spanning paragraphs must be
  split. Text boxes, footnotes and fields are not edited.
- Filling requires template placeholders that docxtpl can see (not split across
  runs).
- Visual fidelity depends on fonts available to LibreOffice; Word may substitute
  differently. Always eyeball the PNGs.
- Tested on macOS; Windows works if `soffice` is on `PATH` or `SOFFICE` is set.

## Publish it

```bash
git init
git add .
git commit -m "rtl-docx: RTL-first Word workflow for agents"
gh repo create rtl-docx --public --source=. --push
```

Then replace `OWNER` in `pyproject.toml`, `README.md` and `SKILL.md`. If you want
one-click distribution to ChatGPT/Codex, wrap the same skill in a plugin; the
CLI stays the deterministic core.

## License

MIT

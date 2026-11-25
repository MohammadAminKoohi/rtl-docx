# Changelog

## 0.3.0

- MCP server (`rtl-docx-mcp`, optional `[mcp]` extra) exposing read, inventory,
  find, validate, lint, build, track, compare, revisions, accept/reject, audit,
  comments, check, fill, verify, convert, merge and stamp.
- Repo is now a Codex plugin: `.codex-plugin/plugin.json`, `.mcp.json`, and the
  skill moved to `skills/rtl-docx/`; installer links that folder.
- `convert` — LibreOffice-backed `.doc`/`.odt`/`.rtf` → `.docx`, `.docx` → PDF.
- `merge` — merge several `.docx` files preserving styles, numbering and headers
  (optional `[merge]` extra, docxcompose).
- `stamp` — headers, footers and real `PAGE` fields on every section.
- `build --toc`, `--toc-depth`, `--number-sections`, `--page-numbers`,
  `--header`, `--footer`, `--page-label`, `--align`.
- Security: ZIP path-traversal rejection, DTD/entity rejection, external
  relationship warnings, read-side size/compression limits.
- Revision engine also resolves move range markers; `accept`/`reject` cover
  body, tables, headers, footers and notes.

## 0.2.0

- `compare` — native tracked-change redline between two .docx files
  (optional `[redline]` extra, structure-aware Docxodus engine).
- `audit` — reject all revisions and prove the result reproduces the original;
  optionally require a single revision author.
- `revisions`, `accept`, `reject` — list and resolve tracked changes across body,
  tables, headers and footers (custom package-wide engine).
- `comments list` / `comments add` — native Word comments via python-docx 1.2.
- `check` — round-trip text check of a built DOCX against its Markdown, with
  Persian normalization and a similarity threshold.
- `lint` — Persian/RTL hygiene: Arabic yeh/kaf, Arabic-Indic digits, Latin
  punctuation, direction mismatches, bidi controls, double spaces.
- `track` now returns per-region receipts (`found`, `old_left`, `new_added`) and
  supports `--normalize-persian`.
- `find` supports `--normalize-persian` and `--ignore-zwnj`.
- Compare output is sanitized (dangling relationships removed, required
  settings/fontTable/theme parts repaired) before validation.
- Version 0.1.1 fixes from the previous release are included.

## 0.1.1

- Fixed invalid style-level `w:rtl` in the reference template (OOXML forbids it
  inside `CT_RPrStyle`); direction is stamped on runs after the build.
- Removed stale `stylesWithEffects.xml`; fixed dangling `TableNormal` basedOn.
- Added `validate` (package, relationships, content types, OOXML schema via
  openxml-audit).
- Generated documents and the bundled reference now validate with zero errors.

## 0.1.0

- Initial release: Markdown → RTL DOCX build, tracked replacements, template
  fills, extraction, region-complete search, structural inventory, Markdown
  dump/diff, LibreOffice PDF/PNG verification, and a Codex/OpenCode skill.

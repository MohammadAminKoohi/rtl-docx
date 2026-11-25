"""MCP server exposing rtl-docx operations to MCP clients.

Install the extra and run over stdio:

    uv tool install "rtl-docx[mcp]"
    rtl-docx-mcp
"""

from __future__ import annotations

from mcp.server.mcpserver import MCPServer

from . import __version__

mcp = MCPServer(
    "rtl-docx",
    version=__version__,
    instructions=(
        "RTL-first Word toolkit. Prefer build_document for new files, track_edits "
        "for tracked changes, compare_documents for redlines, validate_document "
        "before verify_document, and never overwrite existing outputs."
    ),
)


@mcp.tool()
def read_document(path: str) -> str:
    """Read a .docx as Markdown (MarkItDown), preserving headings and tables."""
    from .read import read_markdown

    return read_markdown(path)


@mcp.tool()
def document_inventory(path: str) -> dict:
    """Structural counts: sections, paragraphs, tables, cells, headers, footers, images."""
    from .guard import inventory

    return inventory(path)


@mcp.tool()
def find_text(
    path: str,
    text: str,
    normalize_persian: bool = False,
    ignore_zwnj: bool = False,
) -> dict:
    """Count text across body, table cells, headers and footers."""
    from .guard import count_matches, group_matches

    hits = count_matches(
        path, text, normalize=normalize_persian, ignore_zwnj=ignore_zwnj
    )
    return {"total": sum(hits.values()), "regions": hits, "groups": group_matches(hits)}


@mcp.tool()
def validate_document(path: str, schema: bool = True) -> dict:
    """Validate package structure, relationships, content types and OOXML schema."""
    from .validate import validate_docx

    return validate_docx(path, schema=schema)


@mcp.tool()
def lint_document(path: str) -> dict:
    """Persian/RTL hygiene findings: variants, digits, punctuation, direction."""
    from .persian import lint_docx

    return lint_docx(path)


@mcp.tool()
def build_document(
    markdown_path: str,
    output_path: str,
    toc: bool = False,
    page_numbers: bool = False,
    header: str | None = None,
    footer: str | None = None,
) -> dict:
    """Build a styled RTL .docx from Markdown. Refuses to overwrite."""
    from .build import build_docx

    out = build_docx(
        markdown_path, output_path, toc=toc, page_numbers=page_numbers,
        header=header, footer=footer,
    )
    return {"output": str(out)}


@mcp.tool()
def track_edits(
    input_path: str,
    output_path: str,
    edits: list[dict],
    author: str = "Codex",
    normalize_persian: bool = False,
) -> dict:
    """Apply find/replace edits as native tracked changes with per-region receipts.

    edits: [{"find": "old", "replace": "new"}]
    """
    from .track import apply_edits

    return apply_edits(
        input_path, output_path, edits, author, normalize=normalize_persian
    )


@mcp.tool()
def compare_documents(
    original_path: str,
    modified_path: str,
    output_path: str,
    author: str = "Codex",
    detect_moves: bool = False,
) -> dict:
    """Produce a native tracked-change redline between two .docx versions."""
    from .revisions import compare_documents as _compare

    return _compare(
        original_path, modified_path, output_path, author,
        detect_moves=detect_moves,
    )


@mcp.tool()
def list_revisions(path: str) -> list:
    """List tracked changes across body, tables, headers, footers and notes."""
    from .revisions import list_revisions as _list

    return _list(path)


@mcp.tool()
def resolve_revisions(path: str, output_path: str, mode: str = "accept") -> dict:
    """Accept or reject every tracked change into a new file (mode: accept|reject)."""
    from .revisions import accept_changes, reject_changes

    if mode not in ("accept", "reject"):
        raise ValueError("mode must be 'accept' or 'reject'")
    out = (accept_changes if mode == "accept" else reject_changes)(path, output_path)
    return {"output": str(out), "mode": mode}


@mcp.tool()
def audit_redline(path: str, original_path: str, author: str | None = None) -> dict:
    """Reject revisions and prove the result matches the original; detects untracked edits."""
    from .revisions import audit_changes

    return audit_changes(path, original_path, author)


@mcp.tool()
def list_comments(path: str) -> list:
    """List Word comments with author, date, text and anchor text."""
    from .comments import list_comments as _list

    return _list(path)


@mcp.tool()
def add_comment(
    input_path: str,
    output_path: str,
    anchor: str,
    text: str,
    author: str = "Codex",
    occurrence: int = 1,
) -> dict:
    """Attach a native comment to an existing phrase inside one paragraph."""
    from .comments import add_comment as _add

    return _add(input_path, output_path, anchor, text, author, occurrence=occurrence)


@mcp.tool()
def check_roundtrip(docx_path: str, markdown_path: str) -> dict:
    """Compare extracted DOCX text against the Markdown source."""
    from .check import check_document

    return check_document(docx_path, markdown_path)


@mcp.tool()
def fill_template(template_path: str, data_path: str, output_path: str) -> dict:
    """Render a docxtpl/Jinja2 template with a JSON data file."""
    from .fill import fill_template as _fill

    out = _fill(template_path, data_path, output_path)
    return {"output": str(out)}


@mcp.tool()
def verify_document(path: str, timeout: int = 120) -> dict:
    """Render via LibreOffice and return PDF/page PNG paths for visual review."""
    from .verify import verify_docx

    result = verify_docx(path, timeout=timeout)
    return {
        "pdf": str(result["pdf"]),
        "pages": [str(page) for page in result["pages"]],
    }


@mcp.tool()
def convert_document(source_path: str, output_path: str, to: str = "docx") -> dict:
    """Convert through LibreOffice (legacy .doc -> .docx, docx -> pdf)."""
    from .convert import convert_document as _convert

    out = _convert(source_path, output_path, to=to)
    return {"output": str(out)}


@mcp.tool()
def merge_documents(input_paths: list[str], output_path: str) -> dict:
    """Merge several .docx files, preserving styles and headers."""
    from .merge import merge_documents as _merge

    out = _merge(input_paths, output_path)
    return {"output": str(out)}


@mcp.tool()
def stamp_document(
    input_path: str,
    output_path: str,
    header: str | None = None,
    footer: str | None = None,
    page_numbers: bool = False,
) -> dict:
    """Add header/footer text and page numbers to every section."""
    from .stamp import stamp_copy

    out = stamp_copy(
        input_path, output_path, header=header, footer=footer,
        page_numbers=page_numbers,
    )
    return {"output": str(out)}


def main() -> None:
    mcp.run()

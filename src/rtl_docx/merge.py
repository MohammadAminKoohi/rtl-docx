"""Merge multiple .docx files while preserving styles, numbering and headers."""

from __future__ import annotations

from pathlib import Path


class MergeError(Exception):
    """Merge failed."""


def merge_documents(
    inputs: list[str | Path],
    output: str | Path,
    page_break: bool = True,
) -> Path:
    if len(inputs) < 2:
        raise MergeError("provide at least two input documents")
    output = Path(output)
    sources = [Path(item) for item in inputs]
    for source in sources:
        if not source.exists():
            raise FileNotFoundError(f"not found: {source}")
    if output.exists():
        raise FileExistsError(f"refusing to overwrite {output}")
    if output.resolve() in {source.resolve() for source in sources}:
        raise MergeError("output must differ from every input")

    try:
        from docx import Document
        from docxcompose.composer import Composer
    except ImportError as exc:
        raise MergeError(
            'merge support is not installed: uv tool install "rtl-docx[merge]"'
        ) from exc

    master = Document(str(sources[0]))
    composer = Composer(master)
    for source in sources[1:]:
        if page_break:
            separator = Document()
            separator.add_page_break()
            composer.append(separator)
        composer.append(Document(str(source)))

    output.parent.mkdir(parents=True, exist_ok=True)
    composer.save(str(output))

    from .build import drop_styles_with_effects
    from .validate import repair_required_parts, sanitize_package, validate_docx

    drop_styles_with_effects(output)
    sanitize_package(output)
    repair_required_parts(output)
    result = validate_docx(output, schema=False)
    if not result["valid"]:
        raise MergeError(
            "merged package is structurally invalid:\n"
            + "\n".join(result["errors"][:5])
        )
    return output

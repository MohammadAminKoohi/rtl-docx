"""Convert documents through LibreOffice (doc/odt/rtf -> docx, docx -> pdf)."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from .verify import VerifyError, find_soffice


class ConvertError(Exception):
    """Conversion failed."""


def convert_document(
    source: str | Path,
    output: str | Path | None = None,
    to: str = "docx",
    timeout: int = 180,
) -> Path:
    source = Path(source)
    if not source.exists():
        raise FileNotFoundError(f"not found: {source}")
    to = to.lower().lstrip(".")
    if to not in ("docx", "pdf", "odt", "rtf", "txt"):
        raise ConvertError(f"unsupported target format: {to}")

    target = Path(output) if output else source.with_suffix(f".{to}")
    if target.resolve() == source.resolve():
        raise ConvertError("source and output must differ")
    if target.exists():
        raise FileExistsError(f"refusing to overwrite {target}")

    try:
        soffice = find_soffice()
    except VerifyError as exc:
        raise ConvertError(str(exc)) from exc

    target.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="rtl-docx-convert-") as tmp:
        profile = Path(tmp) / "profile"
        outdir = Path(tmp) / "out"
        outdir.mkdir()
        cmd = [
            soffice, "--headless",
            f"-env:UserInstallation=file://{profile}",
            "--convert-to", to,
            "--outdir", str(outdir),
            str(source.resolve()),
        ]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, check=False)
        except subprocess.TimeoutExpired as exc:
            raise ConvertError(f"LibreOffice timed out after {timeout}s") from exc
        log = (proc.stdout or "") + (proc.stderr or "")
        produced = outdir / f"{source.stem}.{to}"
        if proc.returncode != 0 or not produced.exists():
            raise ConvertError(f"LibreOffice could not convert {source}:\n{log}")

        if to == "docx":
            from .build import drop_styles_with_effects
            drop_styles_with_effects(produced)
            from .validate import repair_required_parts, sanitize_package, validate_docx
            sanitize_package(produced)
            repair_required_parts(produced)
            result = validate_docx(produced, schema=False)
            if not result["valid"]:
                raise ConvertError(
                    "converted package is structurally invalid:\n"
                    + "\n".join(result["errors"][:5])
                )

        shutil.move(str(produced), str(target))
    return target

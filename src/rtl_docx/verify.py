"""Render a .docx through LibreOffice and rasterize pages for visual review."""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path


class VerifyError(Exception):
    """The document could not be rendered."""


def find_soffice() -> str:
    env = os.environ.get("SOFFICE")
    candidates = [
        env,
        shutil.which("soffice"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/lib/libreoffice/program/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)
    raise VerifyError("LibreOffice (soffice) not found; install it or set SOFFICE")


def verify_docx(
    path: str | Path,
    dpi: int = 150,
    timeout: int = 120,
    output_dir: str | Path | None = None,
) -> dict:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"not found: {path}")

    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        raise VerifyError("pdftoppm not found; install poppler")

    outdir = Path(output_dir) if output_dir else Path(".docx-verify") / path.stem
    if outdir.exists():
        shutil.rmtree(outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="rtl-docx-lo-") as profile:
        cmd = [
            find_soffice(), "--headless",
            f"-env:UserInstallation=file://{profile}",
            "--convert-to", "pdf",
            "--outdir", str(outdir.resolve()),
            str(path.resolve()),
        ]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, check=False)
        except subprocess.TimeoutExpired as exc:
            raise VerifyError(
                f"LibreOffice timed out after {timeout}s. The file may be corrupt, "
                f"or too complex for the current timeout. Retry with a larger --timeout."
            ) from exc

    log = (proc.stdout or "") + (proc.stderr or "")
    (outdir / "convert.log").write_text(log, encoding="utf-8")
    pdf = outdir / f"{path.stem}.pdf"
    if proc.returncode != 0 or not pdf.exists() or pdf.stat().st_size == 0:
        raise VerifyError(
            "LibreOffice could not convert the file; the OOXML is likely corrupt.\n" + log
        )

    subprocess.run([pdftoppm, "-png", "-r", str(dpi), str(pdf), str(outdir / "page")], check=True)
    pages = sorted(outdir.glob("page-*.png"))
    if not pages:
        raise VerifyError("no page images produced")
    return {"pdf": pdf, "pages": pages, "log": log, "outdir": outdir}

"""Fill a docxtpl/Jinja2 Word template with JSON data."""

from __future__ import annotations

import json
from pathlib import Path
from zipfile import ZipFile

from docxtpl import DocxTemplate
from jinja2 import Environment, StrictUndefined, UndefinedError


class FillError(Exception):
    """Template rendering failed or left unresolved markers."""


def fill_template(
    template: str | Path,
    data: str | Path | dict,
    output: str | Path,
) -> Path:
    template, output = Path(template), Path(output)
    if not template.exists():
        raise FileNotFoundError(f"template not found: {template}")
    if output.exists():
        raise FileExistsError(f"refusing to overwrite {output}")

    context = json.loads(Path(data).read_text(encoding="utf-8")) if isinstance(data, (str, Path)) else data
    if not isinstance(context, dict):
        raise FillError("data must be a JSON object at the top level")

    tpl = DocxTemplate(str(template))
    env = Environment(autoescape=True, undefined=StrictUndefined)
    try:
        tpl.render(context, jinja_env=env)
    except UndefinedError as exc:
        raise FillError(f"missing data for template variable: {exc}") from exc

    output.parent.mkdir(parents=True, exist_ok=True)
    tpl.save(str(output))

    # Check the raw package, not paragraph text: a tag split across runs still
    # needs to be caught.
    with ZipFile(output) as package:
        xml = "".join(
            package.read(name).decode("utf-8", "ignore")
            for name in package.namelist()
            if name.startswith("word/") and name.endswith(".xml")
        )
    leftovers = [marker for marker in ("{{", "{%", "{#") if marker in xml]
    if leftovers:
        raise FillError(f"unresolved template markers remain: {leftovers}")
    return output

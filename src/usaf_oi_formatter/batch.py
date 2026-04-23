"""Folder iteration. Runs the formatter over every .docx under a path."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from . import formatter, rules


@dataclass
class BatchResult:
    source: Path
    formatted: Path | None
    report: Path | None
    error: str | None

    @property
    def ok(self) -> bool:
        return self.error is None


def run(paths: Iterable[Path],
        output_dir: Path | None = None,
        recurse: bool = False,
        template: Path | None = None,
        log_sink=None) -> list[BatchResult]:
    """Format every .docx under the given paths.

    `paths` may contain files or directories. When `recurse`, directories
    are walked recursively. When `template` is given, that approved OI's
    formatting is cloned onto each draft.
    """
    paths_list = list(paths)
    targets = list(_iter_docx(paths_list, recurse))
    results: list[BatchResult] = []

    owns_sink = False
    if log_sink is None:
        first = Path(paths_list[0]).resolve()
        base_dir = first if first.is_dir() else first.parent
        log_path = base_dir / f"batch_{datetime.now():%Y%m%d_%H%M%S}.log"
        log_sink = log_path.open("w", encoding="utf-8")
        owns_sink = True

    try:
        log_sink.write(f"USAF OI Formatter batch run {datetime.now().isoformat()}\n")
        if template is not None:
            log_sink.write(f"Template: {template}\n")
        log_sink.write(f"Files: {len(targets)}\n\n")

        for target in targets:
            try:
                out_path, report_path = formatter.format_file(
                    target, output_dir=output_dir, template=template)
                results.append(BatchResult(target, out_path, report_path, None))
                log_sink.write(f"OK    {target}  ->  {out_path}\n")
            except Exception as exc:  # noqa: BLE001 - batch keeps going
                results.append(BatchResult(target, None, None, str(exc)))
                log_sink.write(f"FAIL  {target}  {exc!r}\n")
    finally:
        if owns_sink:
            log_sink.close()

    return results


def _iter_docx(paths: Iterable[Path], recurse: bool):
    for raw in paths:
        p = Path(raw)
        if p.is_file() and _is_word_doc(p):
            yield p
        elif p.is_dir():
            for ext in ("docx", "doc"):
                pattern = f"**/*.{ext}" if recurse else f"*.{ext}"
                for found in p.glob(pattern):
                    if _is_word_doc(found) and not _is_formatter_output(found):
                        yield found


def _is_word_doc(p: Path) -> bool:
    return p.suffix.lower() in (".docx", ".docm", ".doc")


def _is_formatter_output(p: Path) -> bool:
    return p.stem.endswith(rules.OUTPUT_SUFFIX)

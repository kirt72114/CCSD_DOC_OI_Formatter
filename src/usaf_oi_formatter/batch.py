"""Folder iteration. Runs the formatter over every .docx under a path."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from . import formatter, rules
from .meta import OIMeta


@dataclass
class BatchResult:
    source: Path
    formatted: Path | None
    report: Path | None
    error: str | None

    @property
    def ok(self) -> bool:
        return self.error is None


def run(paths: Iterable[Path], meta: OIMeta,
        output_dir: Path | None = None,
        recurse: bool = False,
        log_sink=None) -> list[BatchResult]:
    """Format every .docx under the given paths.

    `paths` may contain files or directories. When `recurse`, directories
    are walked recursively. Writes a master log next to the first path if
    `log_sink` is None.
    """
    targets = list(_iter_docx(paths, recurse))
    results: list[BatchResult] = []

    owns_sink = False
    if log_sink is None:
        first = Path(next(iter(paths))).resolve()
        base_dir = first if first.is_dir() else first.parent
        log_path = base_dir / f"batch_{datetime.now():%Y%m%d_%H%M%S}.log"
        log_sink = log_path.open("w", encoding="utf-8")
        owns_sink = True

    try:
        log_sink.write(f"USAF OI Formatter batch run {datetime.now().isoformat()}\n")
        log_sink.write(f"Files: {len(targets)}\n\n")

        for target in targets:
            try:
                out_path, report_path = formatter.format_file(target, meta, output_dir)
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
        if p.is_file() and _is_docx(p):
            yield p
        elif p.is_dir():
            pattern = "**/*.docx" if recurse else "*.docx"
            for found in p.glob(pattern):
                if _is_docx(found) and not _is_formatter_output(found):
                    yield found


def _is_docx(p: Path) -> bool:
    return p.suffix.lower() == ".docx"


def _is_formatter_output(p: Path) -> bool:
    return p.stem.endswith(rules.OUTPUT_SUFFIX)

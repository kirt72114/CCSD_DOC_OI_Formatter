"""argparse entry point: `usaf-oi-formatter ...` or `python -m usaf_oi_formatter`."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import batch as batch_mod
from . import formatter
from .meta import OIMeta


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="usaf-oi-formatter",
        description="Auto-format USAF Operating Instructions to AFH 33-337 / "
                    "DAFMAN 90-161 compliance.",
    )
    p.add_argument("path", type=Path,
                   help="Path to a .docx file or a folder of .docx files.")
    p.add_argument("--recurse", action="store_true",
                   help="When `path` is a folder, recurse into subfolders.")
    p.add_argument("--output-dir", type=Path, default=None,
                   help="Write formatted files here instead of alongside sources.")

    meta = p.add_argument_group("OI metadata (DAFMAN 90-161 title block)")
    meta.add_argument("--unit", default="", help='e.g. "442D MAINTENANCE SQUADRON"')
    meta.add_argument("--unit-short", default="")
    meta.add_argument("--oi-number", default="", help='e.g. "CCSD OI 36-1"')
    meta.add_argument("--date", dest="date_str", default="",
                      help='e.g. "23 April 2026"; defaults to today.')
    meta.add_argument("--category", default="")
    meta.add_argument("--subject", default="")
    meta.add_argument("--opr", default="", help='e.g. "CCSD/CCC"')
    meta.add_argument("--supersedes", default="")
    meta.add_argument("--certified-by", default="")
    meta.add_argument("--pages", default="")
    meta.add_argument("--accessibility", default="")
    meta.add_argument("--releasability", default="")

    p.add_argument("-v", "--verbose", action="store_true")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    meta = OIMeta(
        unit=args.unit,
        unit_short=args.unit_short,
        oi_number=args.oi_number,
        date_str=args.date_str,
        category=args.category,
        subject=args.subject,
        opr=args.opr,
        supersedes=args.supersedes,
        certified_by=args.certified_by,
        pages=args.pages,
        accessibility=args.accessibility,
        releasability=args.releasability,
    )

    path = args.path
    if path.is_file():
        return _run_single(path, meta, args.output_dir, args.verbose)
    if path.is_dir():
        return _run_batch(path, meta, args.output_dir, args.recurse, args.verbose)

    print(f"error: {path} is not a file or directory", file=sys.stderr)
    return 2


def _run_single(path: Path, meta: OIMeta, output_dir: Path | None,
                verbose: bool) -> int:
    try:
        out, report = formatter.format_file(path, meta, output_dir)
    except Exception as exc:  # noqa: BLE001
        print(f"FAIL {path}: {exc}", file=sys.stderr)
        return 1
    print(f"OK   {path} -> {out}")
    if verbose:
        print(f"     report: {report}")
    return 0


def _run_batch(path: Path, meta: OIMeta, output_dir: Path | None,
               recurse: bool, verbose: bool) -> int:
    results = batch_mod.run([path], meta, output_dir=output_dir, recurse=recurse)
    failures = [r for r in results if not r.ok]
    for r in results:
        status = "OK  " if r.ok else "FAIL"
        extra = f" -> {r.formatted}" if r.ok else f" {r.error}"
        print(f"{status} {r.source}{extra}")
    if verbose:
        print(f"Total: {len(results)}  Failed: {len(failures)}")
    return 0 if not failures else 1


if __name__ == "__main__":
    raise SystemExit(main())

"""argparse entry point: `usaf-oi-formatter ...` or `python -m usaf_oi_formatter`."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import batch as batch_mod
from . import formatter


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="usaf-oi-formatter",
        description=(
            "Format USAF Operating Instructions to match an approved OI. "
            "Recommended usage: pass --template pointing at a known-good OI "
            "and the draft inherits its fonts, styles, page layout, running "
            "headers, and numbering scheme."
        ),
    )
    p.add_argument("path", type=Path,
                   help="Path to a .docx file or a folder of .docx files.")
    p.add_argument("--template", type=Path, default=None,
                   help="Path to an approved OI whose formatting should be "
                        "cloned onto the draft.")
    p.add_argument("--recurse", action="store_true",
                   help="When `path` is a folder, recurse into subfolders.")
    p.add_argument("--output-dir", type=Path, default=None,
                   help="Write formatted files here instead of alongside sources.")
    p.add_argument("-v", "--verbose", action="store_true")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    path = args.path
    if args.template is not None and not args.template.is_file():
        print(f"error: --template {args.template} does not exist", file=sys.stderr)
        return 2

    if path.is_file():
        return _run_single(path, args.template, args.output_dir, args.verbose)
    if path.is_dir():
        return _run_batch(path, args.template, args.output_dir,
                          args.recurse, args.verbose)

    print(f"error: {path} is not a file or directory", file=sys.stderr)
    return 2


def _run_single(path: Path, template: Path | None,
                output_dir: Path | None, verbose: bool) -> int:
    try:
        out, report = formatter.format_file(
            path, output_dir=output_dir, template=template)
    except Exception as exc:  # noqa: BLE001
        print(f"FAIL {path}: {exc}", file=sys.stderr)
        if verbose:
            import traceback
            traceback.print_exc()
        return 1
    print(f"OK   {path} -> {out}")
    if verbose:
        print(f"     report: {report}")
    return 0


def _run_batch(path: Path, template: Path | None,
               output_dir: Path | None, recurse: bool,
               verbose: bool) -> int:
    results = batch_mod.run([path],
                            output_dir=output_dir,
                            recurse=recurse,
                            template=template)
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

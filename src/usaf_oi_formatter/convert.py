"""Legacy `.doc` → modern `.docx` conversion via Word COM automation.

Word is the only reliable way to convert the 1997-2003 `.doc` binary
format. LibreOffice's conversion is close but loses some styling; Word
round-trips its own format perfectly. On USAF workstations Word is
always installed, so this is the practical path.

Requires `pywin32`. If pywin32 is missing (e.g. the user is on macOS /
Linux, or hasn't installed the Windows wheel), a clear error is raised
at conversion time — importing this module does not require pywin32.
"""

from __future__ import annotations

import atexit
import os
import tempfile
from pathlib import Path

# Word 2007+ WdSaveFormat constant for .docx
WD_FORMAT_DOCX = 16

_cleanup_paths: list[Path] = []


def ensure_docx(path: Path) -> Path:
    """Return a `.docx` path for `path`.

    If `path` already ends in `.docx`/`.docm`, returned unchanged.
    If it ends in `.doc`, convert to a temporary `.docx` and return the
    temp path. The temp file is deleted at interpreter exit.

    Raises FileNotFoundError if the input doesn't exist, RuntimeError
    if Word COM is unavailable, or an OSError from Word itself if the
    conversion fails.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)

    suffix = path.suffix.lower()
    if suffix in (".docx", ".docm"):
        return path
    if suffix != ".doc":
        raise ValueError(
            f"Unsupported input extension: {path.suffix}. "
            "Expected .doc, .docx, or .docm."
        )

    return _convert_doc_to_docx(path)


def _convert_doc_to_docx(src: Path) -> Path:
    """Run Word, open `src`, save as `.docx`, close. Returns the new path."""
    try:
        import win32com.client as win32  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required to convert .doc files. "
            "Install it with:\n"
            "    pip install --no-index --find-links .\\wheels pywin32\n"
            "…or save the .doc as .docx manually in Word and retry."
        ) from exc

    tmp_dir = Path(tempfile.mkdtemp(prefix="usaf_oi_convert_"))
    dst = tmp_dir / (src.stem + ".docx")
    _cleanup_paths.append(dst)
    _cleanup_paths.append(tmp_dir)

    word = None
    doc = None
    try:
        # DispatchEx forces a fresh Word instance so we don't hijack the
        # user's already-open editor.
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone
        doc = word.Documents.Open(
            FileName=str(src.resolve()),
            ReadOnly=True,
            ConfirmConversions=False,
            AddToRecentFiles=False,
        )
        doc.SaveAs2(FileName=str(dst.resolve()),
                    FileFormat=WD_FORMAT_DOCX)
    except Exception as exc:  # noqa: BLE001 - re-raise with context
        raise OSError(
            f"Word failed to convert {src} to .docx: {exc}"
        ) from exc
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
        except Exception:  # noqa: BLE001 - best effort
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:  # noqa: BLE001 - best effort
            pass

    if not dst.exists():
        raise OSError(f"Word reported success but {dst} is missing.")
    return dst


@atexit.register
def _cleanup_temp_files() -> None:
    for p in _cleanup_paths:
        try:
            if p.is_file():
                p.unlink()
            elif p.is_dir():
                # Directories may still have lingering Word lockfiles; best effort.
                for child in p.iterdir():
                    try:
                        child.unlink()
                    except OSError:
                        pass
                try:
                    p.rmdir()
                except OSError:
                    pass
        except OSError:
            pass

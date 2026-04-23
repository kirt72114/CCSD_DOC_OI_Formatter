"""Change-log collector written next to each formatted document."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from . import rules


class ChangeReport:
    def __init__(self, source_path: Path) -> None:
        self.source_path = Path(source_path)
        self.lines: list[str] = []
        self.started_at = datetime.now()
        self._pre: dict[str, str] = {}
        self.log("=== USAF OI Formatter run ===")
        self.log(f"Document:  {self.source_path}")
        self.log(f"Started:   {self.started_at.isoformat(timespec='seconds')}")

    def log(self, msg: str) -> None:
        self.lines.append(msg)

    def note(self, stage: str, detail: str) -> None:
        self.log(f"[{stage}] {detail}")

    def snapshot_pre(self, data: dict[str, str]) -> None:
        self._pre = dict(data)

    def diff_post(self, data: dict[str, str]) -> None:
        for k, after in data.items():
            before = self._pre.get(k, "(unset)")
            if before != after:
                self.log(f"CHANGED {k}: {before}  =>  {after}")

    def write_sidecar(self, output_docx: Path) -> Path:
        sidecar = self.sidecar_path(output_docx)
        self.log(f"Finished:  {datetime.now().isoformat(timespec='seconds')}")
        sidecar.write_text("\n".join(self.lines), encoding="utf-8")
        return sidecar

    @staticmethod
    def sidecar_path(docx_path: Path) -> Path:
        docx_path = Path(docx_path)
        return docx_path.with_name(docx_path.stem + rules.REPORT_SUFFIX)

    def render(self) -> str:
        return "\n".join(self.lines)

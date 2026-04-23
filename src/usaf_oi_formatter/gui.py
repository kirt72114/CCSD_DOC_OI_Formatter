"""Tkinter GUI. Ship alongside the CLI for analysts who prefer a dialog.

Launch via `usaf-oi-formatter-gui` or `python -m usaf_oi_formatter.gui`.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from tkinter import (BooleanVar, StringVar, Tk, filedialog, messagebox, ttk)

from . import batch as batch_mod
from . import formatter
from .meta import OIMeta


META_FIELDS = [
    ("unit",          "Unit (full name)"),
    ("unit_short",    "Unit (short)"),
    ("oi_number",     "OI number"),
    ("date_str",      "Date"),
    ("category",      "Functional category"),
    ("subject",       "Subject"),
    ("opr",           "OPR"),
    ("supersedes",    "Supersedes"),
    ("certified_by",  "Certified by"),
    ("pages",         "Pages"),
    ("accessibility", "Accessibility"),
    ("releasability", "Releasability"),
]

APP_NAME = "USAF OI Formatter"
CONFIG_PATH = Path.home() / ".usaf_oi_formatter.json"


def main() -> int:
    app = FormatterApp()
    app.mainloop()
    return 0


class FormatterApp(Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("760x720")

        self._vars: dict[str, StringVar] = {}
        self.mode_var = StringVar(value="single")
        self.file_var = StringVar()
        self.folder_var = StringVar()
        self.output_var = StringVar()
        self.recurse_var = BooleanVar(value=False)

        self._build_ui()
        self._load_prefs()

    # ---------- layout ----------------------------------------------

    def _build_ui(self) -> None:
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=12)

        nb.add(self._build_input_tab(nb), text="Input")
        nb.add(self._build_meta_tab(nb), text="OI Metadata")

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(btns, text="Run", command=self._run).pack(side="right")
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(
            side="right", padx=(0, 8))

    def _build_input_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=12)

        mode = ttk.LabelFrame(frame, text="Mode", padding=8)
        mode.pack(fill="x", pady=(0, 8))
        ttk.Radiobutton(mode, text="Single file",
                        variable=self.mode_var, value="single").pack(anchor="w")
        ttk.Radiobutton(mode, text="Folder (batch)",
                        variable=self.mode_var, value="batch").pack(anchor="w")

        file_row = ttk.LabelFrame(frame, text="Single file", padding=8)
        file_row.pack(fill="x", pady=(0, 8))
        ttk.Entry(file_row, textvariable=self.file_var).pack(
            side="left", fill="x", expand=True)
        ttk.Button(file_row, text="Browse...", command=self._browse_file).pack(
            side="left", padx=(8, 0))

        batch_frame = ttk.LabelFrame(frame, text="Batch", padding=8)
        batch_frame.pack(fill="x", pady=(0, 8))
        folder_row = ttk.Frame(batch_frame)
        folder_row.pack(fill="x")
        ttk.Label(folder_row, text="Folder:").pack(side="left")
        ttk.Entry(folder_row, textvariable=self.folder_var).pack(
            side="left", fill="x", expand=True, padx=(4, 0))
        ttk.Button(folder_row, text="Browse...",
                   command=self._browse_folder).pack(side="left", padx=(8, 0))

        ttk.Checkbutton(batch_frame, text="Recurse into subfolders",
                        variable=self.recurse_var).pack(anchor="w", pady=(4, 0))

        out_row = ttk.Frame(frame)
        out_row.pack(fill="x", pady=(8, 0))
        ttk.Label(out_row, text="Output folder (optional):").pack(side="left")
        ttk.Entry(out_row, textvariable=self.output_var).pack(
            side="left", fill="x", expand=True, padx=(4, 0))
        ttk.Button(out_row, text="Browse...",
                   command=self._browse_output).pack(side="left", padx=(8, 0))

        return frame

    def _build_meta_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=12)
        for i, (attr, label) in enumerate(META_FIELDS):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="w",
                                              padx=(0, 8), pady=2)
            var = StringVar()
            ttk.Entry(frame, textvariable=var, width=60).grid(
                row=i, column=1, sticky="ew", pady=2)
            self._vars[attr] = var
        frame.columnconfigure(1, weight=1)
        return frame

    # ---------- actions ---------------------------------------------

    def _browse_file(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("Word documents", "*.docx *.docm"),
                       ("All files", "*.*")])
        if path:
            self.file_var.set(path)

    def _browse_folder(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.folder_var.set(path)

    def _browse_output(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.output_var.set(path)

    def _build_meta(self) -> OIMeta:
        return OIMeta(**{k: v.get() for k, v in self._vars.items()})

    def _run(self) -> None:
        self._save_prefs()
        meta = self._build_meta()
        output_dir = Path(self.output_var.get()) if self.output_var.get() else None

        try:
            if self.mode_var.get() == "single":
                path = Path(self.file_var.get())
                if not path.is_file():
                    messagebox.showerror(APP_NAME, "Pick a .docx file first.")
                    return
                out, report = formatter.format_file(path, meta, output_dir)
                messagebox.showinfo(
                    APP_NAME,
                    f"Formatted:\n{out}\n\nChange report:\n{report}")
            else:
                folder = Path(self.folder_var.get())
                if not folder.is_dir():
                    messagebox.showerror(APP_NAME, "Pick a folder first.")
                    return
                results = batch_mod.run([folder], meta,
                                        output_dir=output_dir,
                                        recurse=self.recurse_var.get())
                failed = sum(1 for r in results if not r.ok)
                messagebox.showinfo(
                    APP_NAME,
                    f"Processed {len(results)} file(s). Failures: {failed}.")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"{type(exc).__name__}: {exc}")

    # ---------- preferences (per-user JSON) -------------------------

    def _load_prefs(self) -> None:
        if not CONFIG_PATH.exists():
            return
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        for attr, var in self._vars.items():
            var.set(data.get(attr, ""))

    def _save_prefs(self) -> None:
        data = {attr: var.get() for attr, var in self._vars.items()}
        try:
            CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
        except OSError:
            pass  # best effort


if __name__ == "__main__":
    raise SystemExit(main())

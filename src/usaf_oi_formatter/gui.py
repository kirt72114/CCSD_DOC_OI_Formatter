"""Tkinter GUI.

Launch with `usaf-oi-formatter-gui`. Point it at a draft `.docx` (or a
folder of drafts) and an approved OI to use as the formatting template.
The template's styles, numbering, headers, footers, and page setup are
cloned onto the draft.
"""

from __future__ import annotations

import json
from pathlib import Path
from tkinter import BooleanVar, StringVar, Tk, filedialog, messagebox, ttk

from . import batch as batch_mod
from . import formatter


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
        self.geometry("720x320")

        self.mode_var = StringVar(value="single")
        self.file_var = StringVar()
        self.folder_var = StringVar()
        self.template_var = StringVar()
        self.output_var = StringVar()
        self.recurse_var = BooleanVar(value=False)

        self._build_ui()
        self._load_prefs()

    # ---------- layout ----------------------------------------------

    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)

        mode = ttk.LabelFrame(frame, text="Mode", padding=8)
        mode.pack(fill="x", pady=(0, 8))
        ttk.Radiobutton(mode, text="Single file",
                        variable=self.mode_var, value="single").pack(anchor="w")
        ttk.Radiobutton(mode, text="Folder (batch)",
                        variable=self.mode_var, value="batch").pack(anchor="w")

        file_row = ttk.Frame(frame)
        file_row.pack(fill="x", pady=2)
        ttk.Label(file_row, text="Draft .docx:", width=18).pack(side="left")
        ttk.Entry(file_row, textvariable=self.file_var).pack(
            side="left", fill="x", expand=True)
        ttk.Button(file_row, text="Browse...", command=self._browse_file).pack(
            side="left", padx=(8, 0))

        folder_row = ttk.Frame(frame)
        folder_row.pack(fill="x", pady=2)
        ttk.Label(folder_row, text="Draft folder:", width=18).pack(side="left")
        ttk.Entry(folder_row, textvariable=self.folder_var).pack(
            side="left", fill="x", expand=True)
        ttk.Button(folder_row, text="Browse...",
                   command=self._browse_folder).pack(side="left", padx=(8, 0))
        ttk.Checkbutton(frame, text="Recurse into subfolders",
                        variable=self.recurse_var).pack(anchor="w", padx=(18 + 4, 0))

        template_row = ttk.Frame(frame)
        template_row.pack(fill="x", pady=(10, 2))
        ttk.Label(template_row, text="Approved template:", width=18).pack(side="left")
        ttk.Entry(template_row, textvariable=self.template_var).pack(
            side="left", fill="x", expand=True)
        ttk.Button(template_row, text="Browse...",
                   command=self._browse_template).pack(side="left", padx=(8, 0))

        out_row = ttk.Frame(frame)
        out_row.pack(fill="x", pady=2)
        ttk.Label(out_row, text="Output folder:", width=18).pack(side="left")
        ttk.Entry(out_row, textvariable=self.output_var).pack(
            side="left", fill="x", expand=True)
        ttk.Button(out_row, text="Browse...",
                   command=self._browse_output).pack(side="left", padx=(8, 0))

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(btns, text="Run", command=self._run).pack(side="right")
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(
            side="right", padx=(0, 8))

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

    def _browse_template(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("Word documents", "*.docx *.docm"),
                       ("All files", "*.*")])
        if path:
            self.template_var.set(path)

    def _browse_output(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.output_var.set(path)

    def _run(self) -> None:
        self._save_prefs()

        template = Path(self.template_var.get()) if self.template_var.get() else None
        output_dir = Path(self.output_var.get()) if self.output_var.get() else None

        if template is not None and not template.is_file():
            messagebox.showerror(APP_NAME,
                                 f"Template not found: {template}")
            return

        try:
            if self.mode_var.get() == "single":
                path = Path(self.file_var.get())
                if not path.is_file():
                    messagebox.showerror(APP_NAME, "Pick a draft .docx first.")
                    return
                out, report = formatter.format_file(
                    path, output_dir=output_dir, template=template)
                messagebox.showinfo(
                    APP_NAME,
                    f"Formatted:\n{out}\n\nChange report:\n{report}")
            else:
                folder = Path(self.folder_var.get())
                if not folder.is_dir():
                    messagebox.showerror(APP_NAME, "Pick a folder first.")
                    return
                results = batch_mod.run(
                    [folder],
                    output_dir=output_dir,
                    recurse=self.recurse_var.get(),
                    template=template)
                failed = sum(1 for r in results if not r.ok)
                messagebox.showinfo(
                    APP_NAME,
                    f"Processed {len(results)} file(s). Failures: {failed}.")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"{type(exc).__name__}: {exc}")

    # ---------- preferences -----------------------------------------

    def _load_prefs(self) -> None:
        if not CONFIG_PATH.exists():
            return
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        self.template_var.set(data.get("template", ""))
        self.output_var.set(data.get("output_dir", ""))

    def _save_prefs(self) -> None:
        data = {
            "template": self.template_var.get(),
            "output_dir": self.output_var.get(),
        }
        try:
            CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
        except OSError:
            pass


if __name__ == "__main__":
    raise SystemExit(main())

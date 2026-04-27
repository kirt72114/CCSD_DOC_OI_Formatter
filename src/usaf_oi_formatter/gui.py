"""Tkinter GUI for the USAF OI Formatter.

Layout:
  - Files tab     : drag-drop zone + Browse buttons + queue list + output dir
  - Settings tab  : template picker + Simple panel + Advanced expander
  - Metadata tab  : DAFMAN 90-161 title-block fields
  - Lint tab      : prose-level T&Q findings, populated after a run

Drag-and-drop uses tkinterdnd2 if it's installed; otherwise the drop
zone is still clickable (it acts as a Browse trigger).

Launch via `usaf-oi-formatter-gui`.
"""

from __future__ import annotations

import json
import threading
from dataclasses import fields
from pathlib import Path
from tkinter import (BooleanVar, DoubleVar, Event, IntVar, StringVar, Tk,
                     filedialog, messagebox, ttk)
from tkinter.scrolledtext import ScrolledText

from . import batch as batch_mod
from . import formatter, lint as lint_mod, templates
from .meta import OIMeta
from .profile import FormattingProfile

# ---- optional drag-drop --------------------------------------------------

try:  # pragma: no cover - optional
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    _HAS_DND = True
    _BaseTk = TkinterDnD.Tk
except Exception:  # noqa: BLE001
    _HAS_DND = False
    _BaseTk = Tk

# ---- constants -----------------------------------------------------------

APP_NAME = "USAF OI Formatter"
CONFIG_PATH = Path.home() / ".usaf_oi_formatter.json"

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

# Settings split: Simple panel = the high-frequency knobs;
# Advanced panel = everything else from FormattingProfile.
_SIMPLE_FIELDS: list[tuple[str, str, str]] = [
    # (attr, label, kind)  kind in {"float", "int", "str", "bool"}
    ("margin_top_in",            "Margin top (in)",         "float"),
    ("margin_bottom_in",         "Margin bottom (in)",      "float"),
    ("margin_left_in",           "Margin left (in)",        "float"),
    ("margin_right_in",          "Margin right (in)",       "float"),
    ("body_font",                "Body font",               "str"),
    ("body_size_pt",             "Body size (pt)",          "float"),
    ("space_after_pt",           "Paragraph space-after",   "float"),
    ("page_numbering_position",  "Page number position",    "str"),
    ("rebuild_title_block",      "Rebuild title block",     "bool"),
    ("seed_glossary",            "Seed Attachment 1 glossary", "bool"),
]

_ADVANCED_FIELDS: list[tuple[str, str, str]] = [
    ("page_width_in",                  "Page width (in)",         "float"),
    ("page_height_in",                 "Page height (in)",        "float"),
    ("heading_font",                   "Heading font",            "str"),
    ("heading_size_pt",                "Heading size (pt)",       "float"),
    ("titleblock_font",                "Title block font",        "str"),
    ("titleblock_size_pt",             "Title block size (pt)",   "float"),
    ("title_size_pt",                  "Title size (pt)",         "float"),
    ("heading_space_before_pt",        "H1 space-before (pt)",    "float"),
    ("sub_heading_space_before_pt",    "H2+ space-before (pt)",   "float"),
    ("bullet_space_after_pt",          "Bullet space-after (pt)", "float"),
    ("bullet_indent_step_in",          "Bullet indent step (in)", "float"),
    ("bullet_l1",                      "Bullet L1 glyph",         "str"),
    ("bullet_l2",                      "Bullet L2 glyph",         "str"),
    ("bullet_l3",                      "Bullet L3 glyph",         "str"),
    ("bullet_l4",                      "Bullet L4 glyph",         "str"),
    ("max_number_depth",               "Numbering max depth",     "int"),
    ("number_indent_step_in",          "Numbering indent (in)",   "float"),
    ("suppress_first_page_number",     "Suppress 1st page #",     "bool"),
    ("apply_hygiene",                  "Apply prose hygiene",     "bool"),
    ("default_accessibility",          "Default accessibility",   "str"),
    ("default_releasability",          "Default releasability",   "str"),
]


# --------------------------------------------------------------------------

def main() -> int:
    app = FormatterApp()
    app.mainloop()
    return 0


class FormatterApp(_BaseTk):  # type: ignore[misc, valid-type]
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("960x780")
        self.minsize(820, 640)

        self._init_style()

        self.queue: list[Path] = []
        self.meta_vars: dict[str, StringVar] = {}
        self.setting_vars: dict[str, object] = {}
        self.template_var = StringVar(value=templates.builtin_names()[0])
        self.output_var = StringVar()
        self.recurse_var = BooleanVar(value=False)
        self.lint_var = BooleanVar(value=True)
        self.advanced_var = BooleanVar(value=False)
        self.status_var = StringVar(value="Ready.")
        self.last_findings: list[lint_mod.LintFinding] = []

        self._build_ui()
        self._load_template_into_settings(self.template_var.get())
        self._load_prefs()
        self._refresh_warnings()

    # ----- chrome --------------------------------------------------

    def _init_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:  # noqa: BLE001
            pass
        style.configure("Header.TLabel", font=("TkDefaultFont", 14, "bold"))
        style.configure("Sub.TLabel", foreground="#555")
        style.configure("Drop.TLabel",
                        background="#f4f6fb",
                        foreground="#3a4467",
                        relief="ridge",
                        borderwidth=1,
                        padding=20,
                        anchor="center",
                        font=("TkDefaultFont", 11))
        style.configure("DropActive.TLabel",
                        background="#dee6ff",
                        foreground="#1d2c63",
                        relief="ridge",
                        borderwidth=2,
                        padding=20,
                        anchor="center",
                        font=("TkDefaultFont", 11, "bold"))
        style.configure("Warn.TLabel", foreground="#a14b00")

    def _build_ui(self) -> None:
        header = ttk.Frame(self, padding=(14, 10, 14, 0))
        header.pack(fill="x")
        ttk.Label(header, text=APP_NAME, style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Format USAF Operating Instructions to AFH 33-337 / DAFMAN 90-161.",
            style="Sub.TLabel",
        ).pack(anchor="w")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=14, pady=10)
        nb.add(self._build_files_tab(nb), text="Files")
        nb.add(self._build_settings_tab(nb), text="Settings")
        nb.add(self._build_meta_tab(nb), text="OI Metadata")
        nb.add(self._build_lint_tab(nb), text="Lint Results")
        self.notebook = nb

        bar = ttk.Frame(self, padding=(14, 0, 14, 10))
        bar.pack(fill="x")
        ttk.Label(bar, textvariable=self.status_var).pack(side="left")
        ttk.Button(bar, text="Run", command=self._run).pack(side="right")
        ttk.Button(bar, text="Clear queue", command=self._clear_queue).pack(
            side="right", padx=(0, 8))

    # ----- Files tab ------------------------------------------------

    def _build_files_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=12)

        drop_text = "Drop .docx files here, or click to browse." if _HAS_DND \
            else "Drag-and-drop unavailable (install tkinterdnd2). Click to browse."
        self.drop_label = ttk.Label(frame, text=drop_text, style="Drop.TLabel")
        self.drop_label.pack(fill="x", pady=(0, 10))
        self.drop_label.bind("<Button-1>", lambda e: self._browse_files())

        if _HAS_DND:
            self.drop_label.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
            self.drop_label.dnd_bind("<<DropEnter>>", self._on_drop_enter)  # type: ignore[attr-defined]
            self.drop_label.dnd_bind("<<DropLeave>>", self._on_drop_leave)  # type: ignore[attr-defined]
            self.drop_label.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore[attr-defined]

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill="x", pady=(0, 10))
        ttk.Button(btn_row, text="Add file(s)…", command=self._browse_files).pack(side="left")
        ttk.Button(btn_row, text="Add folder…", command=self._browse_folder).pack(
            side="left", padx=(8, 0))
        ttk.Checkbutton(btn_row, text="Recurse into subfolders",
                        variable=self.recurse_var).pack(side="left", padx=(16, 0))

        list_frame = ttk.LabelFrame(frame, text="Queue", padding=6)
        list_frame.pack(fill="both", expand=True)

        self.queue_list = ttk.Treeview(list_frame, columns=("path",),
                                       show="headings", height=10)
        self.queue_list.heading("path", text="Path")
        self.queue_list.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(list_frame, orient="vertical",
                               command=self.queue_list.yview)
        scroll.pack(side="right", fill="y")
        self.queue_list.configure(yscrollcommand=scroll.set)

        out_row = ttk.Frame(frame)
        out_row.pack(fill="x", pady=(10, 0))
        ttk.Label(out_row, text="Output folder (optional):").pack(side="left")
        ttk.Entry(out_row, textvariable=self.output_var).pack(
            side="left", fill="x", expand=True, padx=(6, 0))
        ttk.Button(out_row, text="Browse…", command=self._browse_output).pack(
            side="left", padx=(8, 0))
        ttk.Checkbutton(out_row, text="Run T&Q lint",
                        variable=self.lint_var).pack(side="left", padx=(16, 0))
        return frame

    # ----- Settings tab ---------------------------------------------

    def _build_settings_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=12)

        tpl_row = ttk.Frame(frame)
        tpl_row.pack(fill="x", pady=(0, 10))
        ttk.Label(tpl_row, text="Template:").pack(side="left")
        tpl_combo = ttk.Combobox(
            tpl_row, textvariable=self.template_var, state="readonly",
            values=templates.builtin_names() + ["Custom from file…"],
            width=24,
        )
        tpl_combo.pack(side="left", padx=(6, 0))
        tpl_combo.bind("<<ComboboxSelected>>", self._on_template_change)
        ttk.Button(tpl_row, text="Save profile…",
                   command=self._save_profile).pack(side="left", padx=(8, 0))
        ttk.Button(tpl_row, text="Reset",
                   command=lambda: self._load_template_into_settings(
                       self.template_var.get())
                   ).pack(side="left", padx=(8, 0))

        ttk.Label(
            frame,
            text="Template values populate the fields below. "
                 "Override any field; click Run to apply.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(0, 8))

        simple_box = ttk.LabelFrame(frame, text="Common settings", padding=10)
        simple_box.pack(fill="x", pady=(0, 8))
        self._render_setting_grid(simple_box, _SIMPLE_FIELDS)

        adv_toggle = ttk.Checkbutton(
            frame, text="Show advanced settings",
            variable=self.advanced_var,
            command=self._toggle_advanced,
        )
        adv_toggle.pack(anchor="w", pady=(4, 4))

        self.advanced_box = ttk.LabelFrame(
            frame, text="Advanced settings", padding=10)
        self._render_setting_grid(self.advanced_box, _ADVANCED_FIELDS)
        # Hidden until toggled.
        self.advanced_box.pack_forget()

        self.warning_label = ttk.Label(frame, text="", style="Warn.TLabel",
                                       wraplength=820, justify="left")
        self.warning_label.pack(fill="x", pady=(8, 0))
        return frame

    def _render_setting_grid(self, parent: ttk.Frame,
                             specs: list[tuple[str, str, str]]) -> None:
        for row, (attr, label, kind) in enumerate(specs):
            ttk.Label(parent, text=label).grid(
                row=row, column=0, sticky="w", padx=(0, 8), pady=2)
            var = self._make_setting_var(attr, kind)
            self.setting_vars[attr] = var
            if kind == "bool":
                widget = ttk.Checkbutton(parent, variable=var)
            elif attr == "page_numbering_position":
                widget = ttk.Combobox(
                    parent, textvariable=var, state="readonly",
                    values=["bottom-center", "bottom-right", "bottom-left"],
                    width=20)
            else:
                widget = ttk.Entry(parent, textvariable=var, width=42)
            widget.grid(row=row, column=1, sticky="ew", pady=2)
            if hasattr(var, "trace_add"):
                var.trace_add("write", lambda *_: self._refresh_warnings())
        parent.columnconfigure(1, weight=1)

    def _make_setting_var(self, attr: str, kind: str):
        if kind == "bool":
            return BooleanVar()
        if kind == "int":
            return IntVar()
        if kind == "float":
            return DoubleVar()
        return StringVar()

    def _toggle_advanced(self) -> None:
        if self.advanced_var.get():
            self.advanced_box.pack(fill="x", pady=(0, 8))
        else:
            self.advanced_box.pack_forget()

    # ----- Metadata tab ---------------------------------------------

    def _build_meta_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=12)
        ttk.Label(
            frame,
            text="Title-block values per DAFMAN 90-161 Fig A2.2. "
                 "Blanks fall back to template defaults.",
            style="Sub.TLabel",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        for i, (attr, label) in enumerate(META_FIELDS, start=1):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="w",
                                              padx=(0, 8), pady=2)
            var = StringVar()
            ttk.Entry(frame, textvariable=var, width=60).grid(
                row=i, column=1, sticky="ew", pady=2)
            self.meta_vars[attr] = var

        frame.columnconfigure(1, weight=1)
        return frame

    # ----- Lint tab -------------------------------------------------

    def _build_lint_tab(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, padding=12)
        ttk.Label(
            frame,
            text="Prose-level T&Q suggestions appear here after each run "
                 "(when 'Run T&Q lint' is enabled).",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(0, 6))
        self.lint_text = ScrolledText(frame, height=24, wrap="word",
                                      font=("TkFixedFont", 10))
        self.lint_text.pack(fill="both", expand=True)
        self.lint_text.configure(state="disabled")
        return frame

    # ----- queue management -----------------------------------------

    def _browse_files(self) -> None:
        paths = filedialog.askopenfilenames(
            filetypes=[("Word documents", "*.docx *.docm"),
                       ("All files", "*.*")])
        for p in paths:
            self._add_to_queue(Path(p))

    def _browse_folder(self) -> None:
        path = filedialog.askdirectory()
        if not path:
            return
        folder = Path(path)
        pattern = "**/*.docx" if self.recurse_var.get() else "*.docx"
        for found in folder.glob(pattern):
            self._add_to_queue(found)

    def _browse_output(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.output_var.set(path)

    def _clear_queue(self) -> None:
        self.queue.clear()
        for child in self.queue_list.get_children():
            self.queue_list.delete(child)
        self._set_status("Queue cleared.")

    def _add_to_queue(self, path: Path) -> None:
        if not path.is_file():
            return
        if path.suffix.lower() not in {".docx", ".docm"}:
            return
        if path in self.queue:
            return
        self.queue.append(path)
        self.queue_list.insert("", "end", values=(str(path),))
        self._set_status(f"{len(self.queue)} file(s) queued.")

    # ----- drag-drop callbacks --------------------------------------

    def _on_drop_enter(self, _: Event) -> str:  # type: ignore[type-arg]
        self.drop_label.configure(style="DropActive.TLabel")
        return "copy"

    def _on_drop_leave(self, _: Event) -> str:  # type: ignore[type-arg]
        self.drop_label.configure(style="Drop.TLabel")
        return "copy"

    def _on_drop(self, event) -> str:
        self.drop_label.configure(style="Drop.TLabel")
        for raw in self._parse_drop_data(event.data):
            self._add_to_queue(Path(raw))
        return "copy"

    @staticmethod
    def _parse_drop_data(data: str) -> list[str]:
        # tkinterdnd2 emits brace-quoted paths when they contain spaces.
        out: list[str] = []
        buf = []
        in_brace = False
        for ch in data:
            if ch == "{":
                in_brace = True
                continue
            if ch == "}":
                in_brace = False
                out.append("".join(buf))
                buf = []
                continue
            if ch == " " and not in_brace:
                if buf:
                    out.append("".join(buf))
                    buf = []
                continue
            buf.append(ch)
        if buf:
            out.append("".join(buf))
        return [p for p in out if p]

    # ----- profile <-> form -----------------------------------------

    def _on_template_change(self, _event=None) -> None:
        choice = self.template_var.get()
        if choice == "Custom from file…":
            path = filedialog.askopenfilename(
                title="Load FormattingProfile JSON",
                filetypes=[("JSON", "*.json"), ("All files", "*.*")])
            if not path:
                # Roll back to first built-in.
                self.template_var.set(templates.builtin_names()[0])
                self._load_template_into_settings(self.template_var.get())
                return
            try:
                profile = FormattingProfile.load(Path(path))
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror(APP_NAME, f"Could not load profile:\n{exc}")
                self.template_var.set(templates.builtin_names()[0])
                self._load_template_into_settings(self.template_var.get())
                return
            self._populate_settings(profile)
            self._set_status(f"Loaded profile from {path}.")
            return

        self._load_template_into_settings(choice)

    def _load_template_into_settings(self, name: str) -> None:
        try:
            profile = templates.get_builtin(name)
        except KeyError:
            profile = templates.default()
        self._populate_settings(profile)
        self._set_status(f"Template '{profile.name}' loaded.")

    def _populate_settings(self, profile: FormattingProfile) -> None:
        for attr, var in self.setting_vars.items():
            value = getattr(profile, attr, None)
            if value is None:
                continue
            try:
                var.set(value)
            except Exception:  # noqa: BLE001
                # IntVar/DoubleVar may reject blank strings; coerce.
                if isinstance(var, (IntVar, DoubleVar)):
                    var.set(0)
                else:
                    var.set(str(value))

    def _profile_from_form(self) -> FormattingProfile:
        # Start from the currently selected built-in (gives us defaults
        # for any field not exposed in either grid), then override.
        choice = self.template_var.get()
        try:
            profile = templates.get_builtin(choice)
        except KeyError:
            profile = templates.default()

        valid = {f.name for f in fields(FormattingProfile)}
        overrides: dict[str, object] = {}
        for attr, var in self.setting_vars.items():
            if attr not in valid:
                continue
            try:
                overrides[attr] = var.get()
            except Exception:  # noqa: BLE001
                continue
        return profile.copy(**overrides)

    def _refresh_warnings(self) -> None:
        try:
            profile = self._profile_from_form()
        except Exception:  # noqa: BLE001
            self.warning_label.configure(text="")
            return
        warnings = profile.validate()
        self.warning_label.configure(
            text=("\n".join("⚠  " + w for w in warnings) if warnings else ""))

    def _save_profile(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            title="Save formatting profile")
        if not path:
            return
        try:
            self._profile_from_form().save(Path(path))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"Save failed:\n{exc}")
            return
        self._set_status(f"Profile saved → {path}")

    # ----- run pipeline --------------------------------------------

    def _build_meta(self) -> OIMeta:
        return OIMeta(**{k: v.get() for k, v in self.meta_vars.items()})

    def _run(self) -> None:
        if not self.queue:
            messagebox.showwarning(APP_NAME, "Add at least one file to the queue.")
            return
        try:
            profile = self._profile_from_form()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"Settings error: {exc}")
            return

        meta = self._build_meta()
        output_dir = Path(self.output_var.get()) if self.output_var.get() else None
        run_lint = self.lint_var.get()
        files = list(self.queue)

        self._save_prefs(profile)
        self._set_status(f"Formatting {len(files)} file(s)…")

        thread = threading.Thread(
            target=self._do_run, args=(files, meta, output_dir, profile, run_lint),
            daemon=True,
        )
        thread.start()

    def _do_run(self, files: list[Path], meta: OIMeta,
                output_dir: Path | None, profile: FormattingProfile,
                run_lint: bool) -> None:
        results = []
        try:
            results = batch_mod.run(
                files, meta,
                output_dir=output_dir,
                recurse=False,
                profile=profile,
                run_lint=run_lint,
            )
        except Exception as exc:  # noqa: BLE001
            self.after(0, lambda: messagebox.showerror(APP_NAME, str(exc)))
            self.after(0, lambda: self._set_status("Run failed."))
            return

        ok = sum(1 for r in results if r.ok)
        fail = len(results) - ok
        msg = f"Done: {ok} OK, {fail} failed."

        # Aggregate lint findings from successful runs by re-reading the
        # written lint sidecars (simpler than re-opening the doc).
        all_findings_text = ""
        if run_lint:
            chunks = []
            for r in results:
                if r.ok and r.lint and Path(r.lint).exists():
                    chunks.append(f"=== {r.source.name} ===")
                    chunks.append(Path(r.lint).read_text(encoding="utf-8"))
                    chunks.append("")
            all_findings_text = "\n".join(chunks)

        def finish() -> None:
            self._set_status(msg)
            if all_findings_text:
                self._set_lint_text(all_findings_text)
                self.notebook.select(3)
            messagebox.showinfo(APP_NAME, msg)

        self.after(0, finish)

    def _set_lint_text(self, text: str) -> None:
        self.lint_text.configure(state="normal")
        self.lint_text.delete("1.0", "end")
        self.lint_text.insert("1.0", text)
        self.lint_text.configure(state="disabled")

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)

    # ----- preferences ---------------------------------------------

    def _load_prefs(self) -> None:
        if not CONFIG_PATH.exists():
            return
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        for attr, var in self.meta_vars.items():
            var.set(data.get("meta", {}).get(attr, ""))
        last_template = data.get("template")
        if last_template and last_template in templates.builtin_names():
            self.template_var.set(last_template)
            self._load_template_into_settings(last_template)
        if "output_dir" in data:
            self.output_var.set(data["output_dir"])
        if "lint" in data:
            self.lint_var.set(bool(data["lint"]))
        # Apply any saved profile overrides on top of the template.
        prof_data = data.get("profile_overrides")
        if isinstance(prof_data, dict):
            for attr, var in self.setting_vars.items():
                if attr in prof_data:
                    try:
                        var.set(prof_data[attr])
                    except Exception:  # noqa: BLE001
                        continue

    def _save_prefs(self, profile: FormattingProfile) -> None:
        overrides = {a: _safe_get(v) for a, v in self.setting_vars.items()}
        data = {
            "meta": {k: v.get() for k, v in self.meta_vars.items()},
            "template": self.template_var.get(),
            "output_dir": self.output_var.get(),
            "lint": self.lint_var.get(),
            "profile_overrides": {k: v for k, v in overrides.items() if v is not None},
        }
        try:
            CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
        except OSError:
            pass


def _safe_get(var):
    try:
        return var.get()
    except Exception:  # noqa: BLE001
        return None


if __name__ == "__main__":
    raise SystemExit(main())

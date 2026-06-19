#!/usr/bin/env python3
"""
Smart Budget ↔ GFC Mapper — GUI entry point.
Run with: python gui.py
"""
import sys
import os
import json
import threading
import traceback
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

PREFS_FILE = ROOT / ".mapper_prefs.json"

NAVY  = "#1e3a5f"
NAVY2 = "#2d5a8e"
BG    = "#f2f3f5"
WHITE = "#ffffff"
DARK_BG   = "#1e1e1e"
DARK_FG   = "#d4d4d4"
COL_OK    = "#4ec9b0"
COL_ERR   = "#f48771"
COL_STAT  = "#dcdcaa"
COL_DIM   = "#6a6a6a"


def _load_prefs() -> dict:
    try:
        if PREFS_FILE.exists():
            return json.loads(PREFS_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def _save_prefs(prefs: dict):
    try:
        PREFS_FILE.write_text(json.dumps(prefs, indent=2), encoding="utf-8")
    except Exception:
        pass


class MapperApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Smart Budget ↔ GFC Mapper")
        self.configure(bg=BG)
        self.minsize(720, 560)
        self.resizable(True, True)

        self._prefs = _load_prefs()
        self._build_ui()
        self._apply_prefs()
        self._center()

    def _center(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    # ─────────────────────────────────────────────────────────
    # UI construction
    # ─────────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_header()
        self._build_form()
        self._build_thresholds()
        self._build_run_btn()
        self._build_log()

    def _build_header(self):
        bar = tk.Frame(self, bg=NAVY, padx=18, pady=14)
        bar.pack(fill="x")
        tk.Label(
            bar,
            text="Smart Budget  ↔  GFC Mapper",
            font=("Segoe UI", 14, "bold"),
            bg=NAVY, fg=WHITE,
        ).pack(side="left")

    def _build_form(self):
        outer = tk.Frame(self, bg=BG, padx=18, pady=14)
        outer.pack(fill="x")

        self.gfc_var    = tk.StringVar()
        self.boq_var    = tk.StringVar()
        self.master_var = tk.StringVar()
        self.out_var    = tk.StringVar()

        rows = [
            ("GFC File",        self.gfc_var,    self._browse_gfc),
            ("BOQ / Smart Budget", self.boq_var, self._browse_boq),
            ("Category Master", self.master_var, self._browse_master),
            ("Output File",     self.out_var,    self._browse_output),
        ]

        for i, (label, var, cmd) in enumerate(rows):
            tk.Label(
                outer, text=label + ":", anchor="e", width=18,
                bg=BG, font=("Segoe UI", 9),
            ).grid(row=i, column=0, padx=(0, 10), pady=5, sticky="e")

            entry = tk.Entry(
                outer, textvariable=var, font=("Segoe UI", 9),
                relief="solid", bd=1, bg=WHITE,
            )
            entry.grid(row=i, column=1, pady=5, sticky="ew", ipady=3)

            tk.Button(
                outer, text="Browse…", command=cmd,
                font=("Segoe UI", 8), relief="solid", bd=1,
                bg=WHITE, padx=10, pady=2, cursor="hand2",
            ).grid(row=i, column=2, padx=(10, 0), pady=5)

        outer.columnconfigure(1, weight=1)

    def _build_thresholds(self):
        row = tk.Frame(self, bg=BG, padx=18, pady=2)
        row.pack(fill="x")

        def _lbl(text):
            tk.Label(row, text=text, bg=BG, font=("Segoe UI", 9)).pack(side="left")

        _lbl("Match thresholds:")
        _lbl("   Auto ≥")
        self.auto_thresh = tk.IntVar(value=self._prefs.get("auto_thresh", 75))
        tk.Spinbox(
            row, from_=50, to=95, textvariable=self.auto_thresh,
            width=4, font=("Segoe UI", 9), relief="solid", bd=1,
        ).pack(side="left")
        _lbl("     Suggest ≥")
        self.suggest_thresh = tk.IntVar(value=self._prefs.get("suggest_thresh", 55))
        tk.Spinbox(
            row, from_=30, to=74, textvariable=self.suggest_thresh,
            width=4, font=("Segoe UI", 9), relief="solid", bd=1,
        ).pack(side="left")

    def _build_run_btn(self):
        frame = tk.Frame(self, bg=BG, pady=12)
        frame.pack()

        self.run_btn = tk.Button(
            frame,
            text="▶   Run Mapping",
            font=("Segoe UI", 11, "bold"),
            bg=NAVY, fg=WHITE,
            activebackground=NAVY2, activeforeground=WHITE,
            relief="flat", padx=28, pady=9,
            cursor="hand2",
            command=self._run,
        )
        self.run_btn.pack()

    def _build_log(self):
        outer = tk.Frame(self, bg=BG, padx=18, pady=(0, 14))
        outer.pack(fill="both", expand=True)

        tk.Label(outer, text="Log", bg=BG, font=("Segoe UI", 8, "bold"),
                 anchor="w").pack(fill="x", pady=(0, 3))

        inner = tk.Frame(outer, bg=DARK_BG, bd=0)
        inner.pack(fill="both", expand=True)

        self.log = tk.Text(
            inner,
            font=("Consolas", 9),
            bg=DARK_BG, fg=DARK_FG,
            insertbackground=WHITE,
            relief="flat", wrap="word",
            state="disabled",
            padx=8, pady=6,
        )
        scroll = tk.Scrollbar(inner, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.log.tag_configure("ok",   foreground=COL_OK)
        self.log.tag_configure("err",  foreground=COL_ERR)
        self.log.tag_configure("stat", foreground=COL_STAT)
        self.log.tag_configure("dim",  foreground=COL_DIM)

    # ─────────────────────────────────────────────────────────
    # Prefs
    # ─────────────────────────────────────────────────────────

    def _apply_prefs(self):
        self.gfc_var.set(self._prefs.get("gfc_path", ""))
        self.boq_var.set(self._prefs.get("boq_path", ""))
        self.master_var.set(self._prefs.get("master_path", ""))
        self.out_var.set(self._prefs.get("out_path", ""))

    def _persist_prefs(self):
        self._prefs.update({
            "gfc_path":      self.gfc_var.get(),
            "boq_path":      self.boq_var.get(),
            "master_path":   self.master_var.get(),
            "out_path":      self.out_var.get(),
            "auto_thresh":   self.auto_thresh.get(),
            "suggest_thresh": self.suggest_thresh.get(),
        })
        _save_prefs(self._prefs)

    # ─────────────────────────────────────────────────────────
    # File pickers
    # ─────────────────────────────────────────────────────────

    def _browse_gfc(self):
        p = filedialog.askopenfilename(
            title="Select GFC File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if p:
            self.gfc_var.set(p)

    def _browse_boq(self):
        p = filedialog.askopenfilename(
            title="Select BOQ / Smart Budget File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if p:
            self.boq_var.set(p)

    def _browse_master(self):
        p = filedialog.askopenfilename(
            title="Select Category Master File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if p:
            self.master_var.set(p)

    def _browse_output(self):
        p = filedialog.asksaveasfilename(
            title="Save Enriched Output As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="Enriched_Output.xlsx",
        )
        if p:
            self.out_var.set(p)

    # ─────────────────────────────────────────────────────────
    # Logging
    # ─────────────────────────────────────────────────────────

    def _log(self, msg: str, tag: str = ""):
        def _write():
            self.log.configure(state="normal")
            if tag:
                self.log.insert("end", msg + "\n", tag)
            else:
                self.log.insert("end", msg + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")
        self.after(0, _write)

    # ─────────────────────────────────────────────────────────
    # Run
    # ─────────────────────────────────────────────────────────

    def _run(self):
        gfc    = self.gfc_var.get().strip()
        boq    = self.boq_var.get().strip()
        master = self.master_var.get().strip()
        out    = self.out_var.get().strip()

        errors = []
        if not gfc:
            errors.append("GFC file is required.")
        elif not Path(gfc).exists():
            errors.append(f"GFC file not found:\n{gfc}")
        if not boq:
            errors.append("BOQ file is required.")
        elif not Path(boq).exists():
            errors.append(f"BOQ file not found:\n{boq}")
        if not master:
            errors.append("Category Master file is required.")
        elif not Path(master).exists():
            errors.append(f"Category Master not found:\n{master}")
        if not out:
            errors.append("Output file path is required.")

        if errors:
            messagebox.showerror("Missing inputs", "\n\n".join(errors))
            return

        self._persist_prefs()
        self.run_btn.configure(state="disabled", text="Running…")

        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

        threading.Thread(
            target=self._worker,
            args=(gfc, boq, master, out),
            daemon=True,
        ).start()

    def _worker(self, gfc: str, boq: str, master: str, out: str):
        try:
            from engine.gfc_mapping_engine import (
                process_gfc, load_category_master, MATCH_THRESHOLDS,
            )
            from runners.output_builder import build_output_workbook
            from engine.boq_loaders.auto import load_boq

            MATCH_THRESHOLDS["auto"]      = self.auto_thresh.get()
            MATCH_THRESHOLDS["suggested"] = self.suggest_thresh.get()

            self._log("Loading BOQ …")
            budget = load_boq(boq)
            self._log(f"  {len(budget)} BOQ line items loaded", "ok")

            self._log("Loading Category Master …")
            master_data = load_category_master(master)
            self._log(f"  {len(master_data)} master items", "ok")

            self._log("Processing GFC …")
            results = process_gfc(gfc, budget, master_data)
            self._log(f"  {len(results)} GFC line items enriched", "ok")

            self._log("Building output workbook …")
            Path(out).parent.mkdir(parents=True, exist_ok=True)
            project_name = Path(gfc).stem
            stats = build_output_workbook(
                results, budget, master_data, gfc, project_name, out,
            )

            n = max(len(results), 1)
            auto_pct = 100 * stats["auto"] / n
            sug_pct  = 100 * stats["suggested"] / n

            self._log(f"\n✅  Saved → {out}", "ok")
            self._log("─" * 56, "dim")
            self._log(
                f"  Sheets:       {stats['total_sheets']} total  ·  "
                f"{stats['mapped_sheets']} mapped  ·  "
                f"{stats['skipped_sheets']} skipped",
                "stat",
            )
            self._log(
                f"  Items:        {len(results)} extracted  ·  {len(budget)} in BOQ",
                "stat",
            )
            self._log(f"  🟢 Auto       {stats['auto']}  ({auto_pct:.1f}%)", "stat")
            self._log(f"  🟡 Suggested  {stats['suggested']}  ({sug_pct:.1f}%)", "stat")
            self._log(f"  🔴 Not in BOQ {stats['not_in_budget']}", "stat")
            self._log(f"  ✅ PR-ready    {stats['pr_ready']}", "stat")
            if stats.get("new_scope"):
                self._log(f"  ⚠️  New scope: {', '.join(stats['new_scope'])}", "stat")
            self._log("─" * 56, "dim")

            subprocess.Popen(f'explorer "{Path(out).parent}"')

            self.after(0, lambda: messagebox.showinfo(
                "Complete",
                f"Mapping complete!\n\n"
                f"🟢 Auto-matched:  {stats['auto']}  ({auto_pct:.1f}%)\n"
                f"🟡 Suggested:     {stats['suggested']}  ({sug_pct:.1f}%)\n"
                f"🔴 Not in BOQ:    {stats['not_in_budget']}\n"
                f"✅ PR-ready:       {stats['pr_ready']}\n\n"
                f"Saved to:\n{out}",
            ))

        except Exception as exc:
            self._log(f"\n❌  Error: {exc}", "err")
            self._log(traceback.format_exc(), "err")
            self.after(0, lambda e=str(exc): messagebox.showerror("Error", e))

        finally:
            self.after(0, lambda: self.run_btn.configure(
                state="normal", text="▶   Run Mapping",
            ))


if __name__ == "__main__":
    app = MapperApp()
    app.mainloop()

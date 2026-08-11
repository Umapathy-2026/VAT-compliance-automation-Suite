"""
VAT-Returns page
=================
The sidebar page added to VAT_Compliance_App_V6: pick an EU member state
from a dropdown. Spain opens the VAT Transaction vs Sub-ledger reconciliation
tool (vat_reconciliation_core.run_reconciliation), rebuilt with this app's
card/button/color theme (app_theme.py) so it looks like a native page rather
than an embedded tab. Every other country shows an "Under progress" message -
the reconciliation logic itself is Spain-specific (see vat_reconciliation_core's
TAX_REG_NO / TAX_CODES / account numbers), so there is nothing to run yet for
the rest of the EU.

vat_reconciliation_tab.VatReconciliationTab (the plain-ttk version used by
the standalone VAT_Reconciliation.pyw / VAT_Reconciliation_Standalone.pyw) is
untouched - SpainReconciliationPanel below is a separate, re-skinned UI over
the same vat_reconciliation_core engine, not a replacement for it.
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from app_theme import (
    C, FONT_BODY, FONT_BODY_BOLD, make_button, make_card, make_log_box,
    append_log, clear_log, card_header_row,
)
from vat_reconciliation_core import (
    EU_MEMBER_STATES,
    ReconciliationCancelled,
    ReconciliationError,
    run_reconciliation,
)

VAT_FILETYPES = [("Excel / report exports", "*.xlsx *.xls *.csv *.txt"), ("All files", "*.*")]
XLSX_FILETYPES = [("Excel workbook", "*.xlsx"), ("All files", "*.*")]
EXPECTED_VAT_PARTS = 3  # VAT Transaction exports per quarter (Sub-ledger is one file)

# Display names for the dropdown. Driven by vat_reconciliation_core's own
# EU_MEMBER_STATES set, so the two stay in sync if that set ever changes.
_COUNTRY_NAMES = {
    "AT": "Austria", "BE": "Belgium", "BG": "Bulgaria", "CY": "Cyprus",
    "CZ": "Czechia", "DE": "Germany", "DK": "Denmark", "EE": "Estonia",
    "ES": "Spain", "FI": "Finland", "FR": "France", "GR": "Greece",
    "HR": "Croatia", "HU": "Hungary", "IE": "Ireland", "IT": "Italy",
    "LT": "Lithuania", "LU": "Luxembourg", "LV": "Latvia", "MT": "Malta",
    "NL": "Netherlands", "PL": "Poland", "PT": "Portugal", "RO": "Romania",
    "SE": "Sweden", "SI": "Slovenia", "SK": "Slovakia",
    "XI": "Northern Ireland (UK)",
}
COUNTRIES = sorted(
    ((_COUNTRY_NAMES[code], code) for code in EU_MEMBER_STATES if code in _COUNTRY_NAMES),
    key=lambda pair: pair[0],
)
_NAME_TO_CODE = dict(COUNTRIES)
_SPAIN_CODE = "ES"


# ════════════════════════════════════════════════════════════
#  THEMED FILE PICKERS
#  (same behaviour as vat_reconciliation_tab.py's pickers, rebuilt with
#  app_theme colors/buttons instead of plain ttk)
# ════════════════════════════════════════════════════════════
class _ThemedFilePicker(tk.Frame):
    """Label + read-only path entry + Browse button, on one row."""

    def __init__(self, parent, label, filetypes, save=False, optional=False):
        super().__init__(parent, bg=C.CARD_BG)
        self._filetypes = filetypes
        self._save = save
        self.optional = optional
        self.var = tk.StringVar()

        self.columnconfigure(1, weight=1)
        text = f"{label} (optional):" if optional else f"{label}:"
        tk.Label(self, text=text, bg=C.CARD_BG, fg=C.TEXT_DARK, font=FONT_BODY_BOLD,
                  width=24, anchor="w").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.entry = ttk.Entry(self, textvariable=self.var)
        self.entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.button = make_button(self, "Browse...", self._browse, "secondary", width=10)
        self.button.grid(row=0, column=2)

    def _browse(self):
        current = self.var.get().strip()
        initial_dir = os.path.dirname(current) if current else os.getcwd()
        if self._save:
            path = filedialog.asksaveasfilename(
                title="Save reconciliation output as",
                defaultextension=".xlsx",
                initialdir=initial_dir,
                initialfile=os.path.basename(current) or "VAT_Reconciliation_Output.xlsx",
                filetypes=self._filetypes,
            )
        else:
            path = filedialog.askopenfilename(
                title="Select file", initialdir=initial_dir, filetypes=self._filetypes,
            )
        if path:
            self.var.set(os.path.normpath(path))

    def get(self):
        return self.var.get().strip()

    def set_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        self.entry.configure(state="normal" if enabled else "readonly")
        self.button.configure(state=state)


class _ThemedMultiFilePicker(tk.Frame):
    """A labelled list of input files, with Add / Remove / Clear."""

    def __init__(self, parent, label, filetypes, expected=3):
        super().__init__(parent, bg=C.CARD_BG)
        self._filetypes = filetypes
        self._label = label
        self._expected = expected
        self._paths = []

        self.columnconfigure(1, weight=1)

        head = tk.Frame(self, bg=C.CARD_BG)
        head.grid(row=0, column=0, sticky="nw", padx=(0, 8))
        tk.Label(head, text=f"{label}:", bg=C.CARD_BG, fg=C.TEXT_DARK, font=FONT_BODY_BOLD,
                  width=24, anchor="w").grid(row=0, column=0, sticky="w")
        self.count_var = tk.StringVar()
        self.count_label = tk.Label(head, textvariable=self.count_var, bg=C.CARD_BG,
                                     font=FONT_BODY, anchor="w")
        self.count_label.grid(row=1, column=0, sticky="w", pady=(2, 0))

        list_frame = tk.Frame(self, bg="white", highlightbackground=C.CARD_BORDER,
                               highlightthickness=1)
        list_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 8))
        self.listbox = tk.Listbox(list_frame, height=3, activestyle="none", exportselection=False,
                                   bg="white", fg=C.TEXT_DARK, font=FONT_BODY, bd=0,
                                   highlightthickness=0)
        self.listbox.pack(side="left", fill="both", expand=True, padx=6, pady=4)
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scroll.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=scroll.set)

        btns = tk.Frame(self, bg=C.CARD_BG)
        btns.grid(row=0, column=2, sticky="n")
        self.add_button = make_button(btns, "Add files...", self._add, "secondary", width=11)
        self.add_button.pack(pady=(0, 4))
        self.remove_button = make_button(btns, "Remove", self._remove, "secondary", width=11)
        self.remove_button.pack(pady=(0, 4))
        self.clear_button = make_button(btns, "Clear", self._clear, "secondary", width=11)
        self.clear_button.pack()

        self._refresh()

    def get(self):
        return list(self._paths)

    def _refresh(self):
        self.listbox.delete(0, "end")
        for p in self._paths:
            self.listbox.insert("end", f"  {os.path.basename(p)}")
        n = len(self._paths)
        if n == 0:
            self.count_var.set(f"no files ({self._expected} expected)")
            colour = C.ERROR_TEXT
        elif n == self._expected:
            self.count_var.set(f"{n} files")
            colour = C.SUCCESS_TEXT
        else:
            self.count_var.set(f"{n} file{'s' if n != 1 else ''} ({self._expected} expected)")
            colour = C.WARN_TEXT
        self.count_label.configure(fg=colour)
        self.event_generate("<<FilesChanged>>")

    def _add(self):
        initial = os.path.dirname(self._paths[-1]) if self._paths else os.getcwd()
        chosen = filedialog.askopenfilenames(
            title=f"Select {self._label} (you can select several at once)",
            initialdir=initial, filetypes=self._filetypes,
        )
        if not chosen:
            return
        added = 0
        for p in chosen:
            p = os.path.normpath(p)
            if p not in self._paths:
                self._paths.append(p)
                added += 1
        if added < len(chosen):
            messagebox.showinfo(
                "Duplicate files skipped",
                f"{len(chosen) - added} file(s) were already in the list and were not added again.",
                parent=self,
            )
        self._refresh()

    def _remove(self):
        for i in sorted(self.listbox.curselection(), reverse=True):
            del self._paths[i]
        self._refresh()

    def _clear(self):
        self._paths = []
        self._refresh()

    def set_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        for b in (self.add_button, self.remove_button, self.clear_button):
            b.configure(state=state)
        self.listbox.configure(state=state)


# ════════════════════════════════════════════════════════════
#  SPAIN RECONCILIATION PANEL
#  Same engine as vat_reconciliation_tab.VatReconciliationTab
#  (vat_reconciliation_core.run_reconciliation), themed to match this app.
# ════════════════════════════════════════════════════════════
class SpainReconciliationPanel(tk.Frame):
    POLL_MS = 100

    def __init__(self, parent):
        super().__init__(parent, bg=C.PAGE_BG)

        self._queue = queue.Queue()
        self._worker = None
        self._cancel_requested = threading.Event()
        self._last_output = None
        self._poll_id = None
        self._destroyed = False

        self._build_inputs()
        self._build_progress()
        self._build_log()

        self._set_running(False)
        self.bind("<Configure>", self._on_configure)
        self.bind("<Destroy>", self._on_destroy)
        self._poll_id = self.after(self.POLL_MS, self._drain_queue)

    def _on_destroy(self, event):
        if event.widget is not self:
            return  # child widget teardown, not ours
        self._destroyed = True
        if self._poll_id is not None:
            try:
                self.after_cancel(self._poll_id)
            except tk.TclError:
                pass
            self._poll_id = None

    def _on_configure(self, event):
        wrap = max(200, event.width - 60)
        for label in (getattr(self, "status_label", None), getattr(self, "hint_label", None)):
            if label is not None and label.cget("wraplength") != wrap:
                label.configure(wraplength=wrap)

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------
    def _build_inputs(self):
        body = make_card(self, "1. Input Files", "\U0001F4CB")

        self.vat_picker = _ThemedMultiFilePicker(body, "VAT Transaction files", VAT_FILETYPES)
        self.vat_picker.pack(fill="x", pady=(0, 10))

        self.sub_picker = _ThemedFilePicker(body, "Sub-ledger file", VAT_FILETYPES)
        self.sub_picker.pack(fill="x", pady=(0, 10))

        # Local Purchase handling is still to be decided - the sheet is copied
        # in verbatim when a file is given, and simply omitted when it isn't.
        self.local_picker = _ThemedFilePicker(body, "Local Purchase file", XLSX_FILETYPES, optional=True)
        self.local_picker.pack(fill="x", pady=(0, 10))

        tk.Frame(body, bg=C.CARD_BORDER, height=1).pack(fill="x", pady=(0, 10))

        self.out_picker = _ThemedFilePicker(body, "Save output as", XLSX_FILETYPES, save=True)
        self.out_picker.pack(fill="x")

        self.hint_label = tk.Label(
            body,
            text=("Add all the quarter's VAT Transaction parts - they are consolidated "
                  "automatically. Report preamble and totals rows are stripped from every "
                  "file. The VAT Summary layout is built in."),
            bg=C.CARD_BG, fg=C.TEXT_MUTED, font=FONT_BODY, anchor="w", justify="left",
        )
        self.hint_label.pack(fill="x", pady=(10, 0))

        btn_row = tk.Frame(body, bg=C.CARD_BG)
        btn_row.pack(fill="x", pady=(14, 0))
        self.run_button = make_button(btn_row, "Run Reconciliation", self._on_run, "primary")
        self.run_button.pack(side="left")
        self.cancel_button = make_button(btn_row, "Cancel", self._on_cancel, "muted")
        self.cancel_button.pack(side="left", padx=(10, 0))

        # Default the output next to the first VAT file the moment one is
        # chosen, so the common case needs no interaction with the save picker.
        self.vat_picker.bind("<<FilesChanged>>", self._suggest_output_path)

    def _build_progress(self):
        body = make_card(self, "2. Progress", "⏳")
        self.status = tk.StringVar(value="Select the VAT Transaction and Sub-ledger files to begin.")
        self.status_label = tk.Label(
            body, textvariable=self.status, bg=C.CARD_BG, fg=C.TEXT_DARK, font=FONT_BODY,
            anchor="w", justify="left",
        )
        self.status_label.pack(fill="x", pady=(0, 8))
        self.progress = ttk.Progressbar(
            body, style="Accent.Horizontal.TProgressbar", orient="horizontal",
            mode="determinate", maximum=1000,
        )
        self.progress.pack(fill="x")

    def _build_log(self):
        body, actions = card_header_row(self, "3. Activity Log", "\U0001F4DD")
        self.open_button = make_button(actions, "Open Output", self._open_output,
                                        "secondary", state="disabled")
        self.open_button.pack(side="right")
        self.log_text = make_log_box(body, height=10)
        self.log_text.tag_configure("note", foreground=C.WARN_TEXT)
        self.log_text.tag_configure("error", foreground=C.ERROR_TEXT)
        self.log_text.tag_configure("success", foreground=C.SUCCESS_TEXT)
        self.log_text.tag_configure("heading", font=("Consolas", 9, "bold"))

    # ------------------------------------------------------------------
    # Small UI helpers
    # ------------------------------------------------------------------
    def _suggest_output_path(self, *_):
        vat_paths = self.vat_picker.get()
        if not vat_paths:
            return
        suggested = os.path.join(
            os.path.dirname(vat_paths[0]), "VAT_Reconciliation_Output.xlsx"
        )
        if not self.out_picker.get():
            self.out_picker.var.set(os.path.normpath(suggested))

    def _set_running(self, running):
        for picker in (self.vat_picker, self.sub_picker, self.local_picker, self.out_picker):
            picker.set_enabled(not running)
        self.run_button.configure(state="disabled" if running else "normal")
        self.cancel_button.configure(state="normal" if running else "disabled")
        if not running:
            self.open_button.configure(
                state="normal" if self._last_output and os.path.exists(self._last_output) else "disabled"
            )

    def _open_output(self):
        if not self._last_output or not os.path.exists(self._last_output):
            return
        try:
            if sys.platform == "win32":
                os.startfile(self._last_output)  # noqa: S606 - opening the file we just wrote
            elif sys.platform == "darwin":
                subprocess.Popen(["open", self._last_output])
            else:
                subprocess.Popen(["xdg-open", self._last_output])
        except OSError as exc:
            messagebox.showerror("Could not open file", str(exc), parent=self)

    # ------------------------------------------------------------------
    # Validation and run
    # ------------------------------------------------------------------
    def _validate(self):
        """Return (vat_list, sub_path, local_or_None, out) or None after showing
        an error."""
        vat = self.vat_picker.get()
        sub = self.sub_picker.get()
        local = self.local_picker.get()
        out = self.out_picker.get()

        if not vat:
            messagebox.showwarning(
                "Files required",
                "Please add the VAT Transaction files for the quarter.",
                parent=self,
            )
            return None
        missing = [p for p in vat if not os.path.exists(p)]
        if missing:
            messagebox.showerror(
                "Files not found",
                "These VAT Transaction files no longer exist:\n\n"
                + "\n".join(os.path.basename(p) for p in missing),
                parent=self,
            )
            return None

        if not sub:
            messagebox.showwarning(
                "File required",
                "Please select the Sub-ledger file for the quarter.",
                parent=self,
            )
            return None
        if not os.path.exists(sub):
            messagebox.showerror(
                "File not found",
                f"The Sub-ledger file no longer exists:\n\n{sub}",
                parent=self,
            )
            return None

        # The same file on both sides would be reconciled against itself.
        if os.path.normcase(sub) in {os.path.normcase(p) for p in vat}:
            messagebox.showerror(
                "Same file used twice",
                "The Sub-ledger file is also in the VAT Transaction list:\n\n"
                f"{os.path.basename(sub)}",
                parent=self,
            )
            return None

        # An unexpected VAT file count usually means an export was forgotten.
        # Warn, but let the preparer proceed - some quarters legitimately differ.
        if len(vat) != EXPECTED_VAT_PARTS and not messagebox.askyesno(
            "Unexpected number of files",
            f"The VAT Transaction report is normally exported in "
            f"{EXPECTED_VAT_PARTS} parts per quarter, but {len(vat)} "
            f"file(s) are selected.\n\nContinue anyway?",
            parent=self,
        ):
            return None

        if local and not os.path.exists(local):
            messagebox.showerror(
                "File not found", f"The Local Purchase file no longer exists:\n\n{local}", parent=self
            )
            return None

        if not out:
            messagebox.showwarning(
                "Output required", "Please choose where to save the output file.", parent=self
            )
            return None

        out_dir = os.path.dirname(out) or "."
        if not os.path.isdir(out_dir):
            messagebox.showerror(
                "Folder not found", f"The output folder does not exist:\n\n{out_dir}", parent=self
            )
            return None

        # Writing over an input would destroy it - refuse rather than warn.
        all_inputs = list(vat) + [sub] + ([local] if local else [])
        if os.path.normcase(os.path.abspath(out)) in {
            os.path.normcase(os.path.abspath(p)) for p in all_inputs
        }:
            messagebox.showerror(
                "Invalid output file",
                "The output file cannot be one of the input files.\n\n"
                "Please choose a different output location.",
                parent=self,
            )
            return None

        if os.path.exists(out) and not messagebox.askyesno(
            "Overwrite existing file?",
            f"{os.path.basename(out)} already exists in that folder.\n\nReplace it?",
            parent=self,
        ):
            return None

        return vat, sub, (local or None), out

    def _on_run(self):
        if self._worker and self._worker.is_alive():
            return
        paths = self._validate()
        if not paths:
            return
        vat, sub, local, out = paths

        clear_log(self.log_text)
        self._last_output = None
        self._cancel_requested.clear()
        self.progress.configure(value=0)
        append_log(self.log_text, f"VAT reconciliation started {datetime.now():%Y-%m-%d %H:%M:%S}")
        append_log(
            self.log_text,
            f"Consolidating {len(vat)} VAT Transaction file(s) and reading "
            f"1 Sub-ledger file. Large exports take a few minutes to read.",
        )
        append_log(
            self.log_text,
            "Note: VIES validation contacts the EU service once per unique VAT "
            "number, so this step takes about a second each.",
        )
        self._set_running(True)
        self.status.set("Starting ...")

        self._worker = threading.Thread(
            target=self._work, args=(vat, sub, local, out), daemon=True
        )
        self._worker.start()

    def _on_cancel(self):
        if self._worker and self._worker.is_alive():
            self._cancel_requested.set()
            self.status.set("Cancelling ...")
            self.cancel_button.configure(state="disabled")

    # ------------------------------------------------------------------
    # Worker thread - must only communicate via self._queue
    # ------------------------------------------------------------------
    def _work(self, vat, sub, local, out):
        def log(message, level="info"):
            self._queue.put(("log", message, level))

        def progress(fraction, message):
            self._queue.put(("progress", fraction, message))

        try:
            result = run_reconciliation(
                vat_files=vat,
                subledger_file=sub,
                output_file=out,
                local_purchase_file=local,
                log=log,
                progress=progress,
                should_cancel=self._cancel_requested.is_set,
            )
            self._queue.put(("done", result, None))
        except ReconciliationCancelled:
            self._queue.put(("cancelled", None, None))
        except ReconciliationError as exc:
            self._queue.put(("failed", str(exc), None))
        except Exception as exc:  # unexpected - show the type so it's diagnosable
            self._queue.put(("failed", f"Unexpected {type(exc).__name__}: {exc}", None))

    # ------------------------------------------------------------------
    # UI thread - drains worker messages
    # ------------------------------------------------------------------
    def _drain_queue(self):
        try:
            while True:
                kind, a, b = self._queue.get_nowait()
                if kind == "log":
                    append_log(self.log_text, a) if b not in ("note", "error") else \
                        self._append_log_tagged(a, b)
                elif kind == "progress":
                    self.progress.configure(value=a * 1000)
                    if b:
                        self.status.set(b)
                elif kind == "done":
                    self._on_finished(a)
                elif kind == "cancelled":
                    self._on_cancelled()
                elif kind == "failed":
                    self._on_failed(a)
        except queue.Empty:
            pass
        if not self._destroyed:
            self._poll_id = self.after(self.POLL_MS, self._drain_queue)

    def _append_log_tagged(self, message, tag):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _on_finished(self, result):
        self._last_output = result.output_file
        self.progress.configure(value=1000)
        self.status.set(f"Finished - {result.row_count} record(s) written.")

        append_log(self.log_text, "")
        self._append_log_tagged("Reconciliation complete.", "success")
        append_log(self.log_text, f"  Records written : {result.row_count}")
        if result.vies_counts:
            summary = ", ".join(
                f"{count} {status.lower()}" for status, count in sorted(result.vies_counts.items())
            )
            append_log(self.log_text, f"  VIES results    : {summary}")
        if result.notes:
            append_log(self.log_text, "  Warnings:")
            for note in result.notes:
                self._append_log_tagged(f"    - {note}", "note")
        append_log(self.log_text, f"  Output          : {result.output_file}")

        self._set_running(False)
        messagebox.showinfo(
            "Reconciliation complete",
            f"{result.row_count} record(s) written to:\n\n{result.output_file}",
            parent=self,
        )

    def _on_cancelled(self):
        self.status.set("Cancelled.")
        append_log(self.log_text, "")
        self._append_log_tagged("Cancelled - no output file was written.", "error")
        self.progress.configure(value=0)
        self._set_running(False)

    def _on_failed(self, message):
        self.status.set("Failed.")
        append_log(self.log_text, "")
        self._append_log_tagged(f"Failed: {message}", "error")
        self.progress.configure(value=0)
        self._set_running(False)
        messagebox.showerror("Reconciliation failed", message, parent=self)


# ════════════════════════════════════════════════════════════
#  VAT-RETURNS PAGE (country picker + content switching)
# ════════════════════════════════════════════════════════════
class VatReturnsPage(tk.Frame):
    TITLE = "VAT-Returns"
    SUBTITLE = "Select an EU member state to open its VAT return / reconciliation workflow."
    ICON = "\U0001F9FE"  # receipt

    def __init__(self, parent):
        super().__init__(parent, bg=C.PAGE_BG)

        picker_body = make_card(self, "Country", "\U0001F30D")
        picker_body.columnconfigure(1, weight=0)
        tk.Label(picker_body, text="Country:", bg=C.CARD_BG, fg=C.TEXT_DARK,
                 font=FONT_BODY_BOLD).grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.country_var = tk.StringVar()
        self.country_box = ttk.Combobox(
            picker_body, textvariable=self.country_var, state="readonly",
            values=[name for name, _ in COUNTRIES], width=35,
        )
        self.country_box.grid(row=0, column=1, sticky="w")
        self.country_box.bind("<<ComboboxSelected>>", self._on_country_change)

        self.content = tk.Frame(self, bg=C.PAGE_BG)
        self.content.pack(fill="both", expand=True)

        self._spain_panel = None       # built lazily, first time Spain is picked
        self._progress_panel = None    # built lazily, reused for every other country
        self._progress_label = None
        self._current = None

        self._placeholder = self._build_placeholder()
        self._show(self._placeholder)

    # -- content builders -------------------------------------------------
    def _build_placeholder(self):
        frame = tk.Frame(self.content, bg=C.PAGE_BG)
        body = make_card(frame, "Get started", "\U0001F447")
        tk.Label(
            body, text="Select a country above to open its VAT return workflow.",
            bg=C.CARD_BG, fg=C.TEXT_MUTED, font=FONT_BODY, anchor="w",
        ).pack(anchor="w")
        return frame

    def _build_progress_panel(self):
        frame = tk.Frame(self.content, bg=C.PAGE_BG)
        body = make_card(frame, "Under progress", "\U0001F6A7")
        self._progress_label = tk.Label(
            body, text="", bg=C.CARD_BG, fg=C.TEXT_MUTED, font=FONT_BODY,
            anchor="w", justify="left", wraplength=760,
        )
        self._progress_label.pack(anchor="w")
        return frame

    # -- selection handling -------------------------------------------------
    def _on_country_change(self, _event=None):
        name = self.country_var.get()
        code = _NAME_TO_CODE.get(name)

        if code == _SPAIN_CODE:
            if self._spain_panel is None:
                self._spain_panel = SpainReconciliationPanel(self.content)
            self._show(self._spain_panel)
        elif code:
            if self._progress_panel is None:
                self._progress_panel = self._build_progress_panel()
            self._progress_label.configure(
                text=f"VAT returns for {name} are Under Progress. This workflow "
                     "isn't available yet - check back in a future release."
            )
            self._show(self._progress_panel)
        else:
            self._show(self._placeholder)

    def _show(self, widget):
        if self._current is widget:
            return
        if self._current is not None:
            self._current.pack_forget()
        widget.pack(fill="both", expand=True)
        self._current = widget

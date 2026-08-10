"""
VAT Reconciliation - desktop UI tab
===================================
`VatReconciliationTab` is a plain `ttk.Frame`, so it drops straight into the
main application's notebook:

    from vat_reconciliation_tab import VatReconciliationTab

    notebook.add(VatReconciliationTab(notebook), text="VAT Reconciliation")

It owns no toplevel window and creates no styles that would affect other tabs.
Run VAT_Reconciliation.pyw to use it stand-alone.

The reconciliation itself runs on a worker thread; the worker never touches
Tk widgets directly - it pushes messages onto a queue that the UI thread
drains on a timer. That keeps the window responsive (and the Cancel button
live) during the slow VIES lookups.
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from vat_reconciliation_core import (
    ReconciliationCancelled,
    ReconciliationError,
    run_reconciliation,
)

PAD = 8
LABEL_WIDTH = 30  # characters; sized to the longest picker label
EXPECTED_VAT_PARTS = 3  # VAT Transaction exports per quarter (Sub-ledger is one file)

VAT_FILETYPES = [("Excel / report exports", "*.xlsx *.xls *.csv *.txt"), ("All files", "*.*")]
XLSX_FILETYPES = [("Excel workbook", "*.xlsx"), ("All files", "*.*")]


class _MultiFilePicker(ttk.Frame):
    """A labelled list of input files, with Add / Remove / Clear.

    Used for the report sets that arrive as several exports per quarter
    (normally three). Any number is accepted - the count is shown so a preparer
    can see at a glance whether the expected number of files is loaded.
    """

    def __init__(self, parent, label, filetypes, expected=3):
        super().__init__(parent)
        self._filetypes = filetypes
        self._label = label
        self._expected = expected
        self._paths = []

        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        head = ttk.Frame(self)
        head.grid(row=0, column=0, sticky="nw", padx=(0, 4))
        ttk.Label(head, text=f"{label}:", width=LABEL_WIDTH, anchor="w").grid(row=0, column=0, sticky="w")
        self.count_var = tk.StringVar()
        self.count_label = ttk.Label(head, textvariable=self.count_var, anchor="w")
        self.count_label.grid(row=1, column=0, sticky="w", pady=(2, 0))

        self.listbox = tk.Listbox(self, height=3, activestyle="none", exportselection=False)
        self.listbox.grid(row=0, column=1, sticky="nsew")
        scroll = ttk.Scrollbar(self, orient="vertical", command=self.listbox.yview)
        scroll.grid(row=0, column=2, sticky="ns", padx=(0, PAD))
        self.listbox.configure(yscrollcommand=scroll.set)

        btns = ttk.Frame(self)
        btns.grid(row=0, column=3, sticky="n")
        self.add_button = ttk.Button(btns, text="Add files...", width=12, command=self._add)
        self.add_button.grid(row=0, column=0, pady=(0, 2))
        self.remove_button = ttk.Button(btns, text="Remove", width=12, command=self._remove)
        self.remove_button.grid(row=1, column=0, pady=(0, 2))
        self.clear_button = ttk.Button(btns, text="Clear", width=12, command=self._clear)
        self.clear_button.grid(row=2, column=0)

        self._refresh()

    # -- state ------------------------------------------------------------
    def get(self):
        return list(self._paths)

    def set_paths(self, paths):
        self._paths = list(paths)
        self._refresh()

    def _refresh(self):
        self.listbox.delete(0, "end")
        for p in self._paths:
            self.listbox.insert("end", f"  {os.path.basename(p)}")
        n = len(self._paths)
        if n == 0:
            self.count_var.set(f"no files ({self._expected} expected)")
            colour = "#b3261e"
        elif n == self._expected:
            self.count_var.set(f"{n} files")
            colour = "#1a7f37"
        else:
            self.count_var.set(f"{n} file{'s' if n != 1 else ''} ({self._expected} expected)")
            colour = "#9a6700"
        self.count_label.configure(foreground=colour)
        self.event_generate("<<FilesChanged>>")

    # -- actions ----------------------------------------------------------
    def _add(self):
        initial = os.path.dirname(self._paths[-1]) if self._paths else os.getcwd()
        chosen = filedialog.askopenfilenames(
            title=f"Select {self._label} (you can select several at once)",
            initialdir=initial,
            filetypes=self._filetypes,
        )
        if not chosen:
            return
        added = 0
        for p in chosen:
            p = os.path.normpath(p)
            # Silently ignore a file already in the list - adding it twice would
            # double every amount it contributes.
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


class _FilePicker(ttk.Frame):
    """Label + read-only path entry + Browse button, on one row."""

    def __init__(self, parent, label, filetypes, save=False, optional=False):
        super().__init__(parent)
        self._filetypes = filetypes
        self._save = save
        self.optional = optional
        self.var = tk.StringVar()

        self.columnconfigure(1, weight=1)
        # LABEL_WIDTH must fit the longest label in full ("Local Purchase file
        # (optional):"), otherwise ttk clips it rather than growing the column.
        text = f"{label} (optional):" if optional else f"{label}:"
        ttk.Label(self, text=text, width=LABEL_WIDTH, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 4)
        )
        self.entry = ttk.Entry(self, textvariable=self.var)
        self.entry.grid(row=0, column=1, sticky="ew", padx=(0, PAD))
        self.button = ttk.Button(self, text="Browse...", width=12, command=self._browse)
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
                title="Select file",
                initialdir=initial_dir,
                filetypes=self._filetypes,
            )
        if path:
            self.var.set(os.path.normpath(path))

    def get(self):
        return self.var.get().strip()

    def set_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        self.entry.configure(state="normal" if enabled else "readonly")
        self.button.configure(state=state)


class VatReconciliationTab(ttk.Frame):
    """The VAT reconciliation screen. Embed as a notebook tab or pack alone."""

    POLL_MS = 100  # how often the UI thread drains the worker's message queue

    def __init__(self, parent, **kwargs):
        super().__init__(parent, padding=PAD, **kwargs)

        self._queue = queue.Queue()
        self._worker = None
        self._cancel_requested = threading.Event()
        self._last_output = None
        self._poll_id = None
        self._destroyed = False

        # Row layout: 0 inputs | 1 action bar | 2 status line | 3 log (grows).
        # Each of these gets its own row - putting the status label in the same
        # cell as the action bar makes it paint over the buttons.
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        self._build_inputs()
        self._build_actions()
        self._build_log()

        self._set_running(False)
        self.bind("<Configure>", self._on_configure)
        self.bind("<Destroy>", self._on_destroy)
        self._poll_id = self.after(self.POLL_MS, self._drain_queue)

    def _on_destroy(self, event):
        """Stop the polling loop when the tab goes away, otherwise the pending
        `after` callback fires against a dead widget and Tk raises
        'invalid command name ..._drain_queue' as the host app shuts down."""
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
        """Keep the wrapping text fitted to the tab's current width.

        Wrapping to the actual width rather than a fixed number of pixels keeps
        these labels one line on a wide window, which matters: every extra line
        raises the tab's minimum height and squeezes the log.
        """
        wrap = max(200, event.width - 2 * PAD)
        for label in (self.status_label, self.hint_label):
            if label.cget("wraplength") != wrap:
                label.configure(wraplength=wrap)

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------
    def _build_inputs(self):
        box = ttk.LabelFrame(self, text="Input files", padding=PAD)
        box.grid(row=0, column=0, sticky="ew")
        box.columnconfigure(0, weight=1)

        # The VAT Transaction report is exported in several parts per quarter
        # and is consolidated by the tool, so the user no longer merges them by
        # hand. The Sub-ledger report can be pulled for the whole quarter in one
        # go, so it is a single file - the tool still strips its report preamble
        # and trailer.
        self.vat_picker = _MultiFilePicker(box, "VAT Transaction files", VAT_FILETYPES)
        self.vat_picker.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        self.sub_picker = _FilePicker(box, "Sub-ledger file", VAT_FILETYPES)
        self.sub_picker.grid(row=1, column=0, sticky="ew", pady=(0, 6))

        # Local Purchase handling is still to be decided - the sheet is copied
        # in verbatim when a file is given, and simply omitted when it isn't.
        self.local_picker = _FilePicker(
            box, "Local Purchase file", XLSX_FILETYPES, optional=True
        )
        self.local_picker.grid(row=2, column=0, sticky="ew", pady=(0, 4))

        ttk.Separator(box, orient="horizontal").grid(row=3, column=0, sticky="ew", pady=PAD)

        self.out_picker = _FilePicker(box, "Save output as", XLSX_FILETYPES, save=True)
        self.out_picker.grid(row=4, column=0, sticky="ew")

        self.hint_label = ttk.Label(
            box,
            text=("Add all the quarter's VAT Transaction parts - they are consolidated "
                  "automatically. Report preamble and totals rows are stripped from every "
                  "file. The VAT Summary layout is built in."),
            foreground="#666666",
            justify="left",
        )
        self.hint_label.grid(row=5, column=0, sticky="w", pady=(6, 0))

        # Default the output next to the first VAT file the moment one is
        # chosen, so the common case needs no interaction with the save picker.
        self.vat_picker.bind("<<FilesChanged>>", self._suggest_output_path)

    def _build_actions(self):
        bar = ttk.Frame(self)
        bar.grid(row=1, column=0, sticky="ew", pady=(PAD, 2))
        bar.columnconfigure(2, weight=1)

        self.run_button = ttk.Button(bar, text="Run Reconciliation", command=self._on_run)
        self.run_button.grid(row=0, column=0, sticky="ns")

        self.cancel_button = ttk.Button(bar, text="Cancel", command=self._on_cancel)
        self.cancel_button.grid(row=0, column=1, padx=(6, 0), sticky="ns")

        self.progress = ttk.Progressbar(bar, mode="determinate", maximum=1000)
        self.progress.grid(row=0, column=2, sticky="ew", padx=PAD)

        self.open_button = ttk.Button(
            bar, text="Open Output", command=self._open_output, state="disabled"
        )
        self.open_button.grid(row=0, column=3, sticky="ns")

        self.status = tk.StringVar(value="Select the VAT Transaction and Sub-ledger files to begin.")
        # Own row, below the buttons. wraplength is set from the current width
        # in _on_configure so a long status/error never widens the whole tab.
        self.status_label = ttk.Label(
            self, textvariable=self.status, anchor="w", justify="left", foreground="#444444"
        )
        self.status_label.grid(row=2, column=0, sticky="ew", pady=(0, 2))

    def _build_log(self):
        box = ttk.LabelFrame(self, text="Progress log", padding=4)
        box.grid(row=3, column=0, sticky="nsew", pady=(4, 0))
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)

        # height=1, not a comfortable reading height: this is the only row that
        # expands, so it takes all the leftover space anyway. Requesting more
        # would raise the tab's minimum height above what a short host window
        # can give, and Tk then squashes this widget to a negative offset
        # instead of simply making it small.
        self.log_text = tk.Text(box, height=1, wrap="none", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        yscroll = ttk.Scrollbar(box, orient="vertical", command=self.log_text.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll = ttk.Scrollbar(box, orient="horizontal", command=self.log_text.xview)
        xscroll.grid(row=1, column=0, sticky="ew")
        self.log_text.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.log_text.tag_configure("note", foreground="#9a6700")
        self.log_text.tag_configure("error", foreground="#b3261e")
        self.log_text.tag_configure("success", foreground="#1a7f37")
        self.log_text.tag_configure("heading", font=("TkDefaultFont", 9, "bold"))

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
        # Only fill a blank box - never overwrite a location the user picked.
        if not self.out_picker.get():
            self.out_picker.var.set(os.path.normpath(suggested))

    def _append_log(self, message, tag=None):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n", tag or ())
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

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

        self._clear_log()
        self._last_output = None
        self._cancel_requested.clear()
        self.progress.configure(value=0)
        self._append_log(
            f"VAT reconciliation started {datetime.now():%Y-%m-%d %H:%M:%S}", "heading"
        )
        self._append_log(
            f"Consolidating {len(vat)} VAT Transaction file(s) and reading "
            f"1 Sub-ledger file. Large exports take a few minutes to read."
        )
        self._append_log(
            "Note: VIES validation contacts the EU service once per unique VAT "
            "number, so this step takes about a second each."
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
                    self._append_log(a, b if b in ("note", "error") else None)
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

    def _on_finished(self, result):
        self._last_output = result.output_file
        self.progress.configure(value=1000)
        self.status.set(f"Finished - {result.row_count} record(s) written.")

        self._append_log("")
        self._append_log("Reconciliation complete.", "success")
        self._append_log(f"  Records written : {result.row_count}")
        if result.vies_counts:
            summary = ", ".join(
                f"{count} {status.lower()}" for status, count in sorted(result.vies_counts.items())
            )
            self._append_log(f"  VIES results    : {summary}")
        if result.notes:
            self._append_log("  Warnings:")
            for note in result.notes:
                self._append_log(f"    - {note}", "note")
        self._append_log(f"  Output          : {result.output_file}")

        self._set_running(False)
        messagebox.showinfo(
            "Reconciliation complete",
            f"{result.row_count} record(s) written to:\n\n{result.output_file}",
            parent=self,
        )

    def _on_cancelled(self):
        self.status.set("Cancelled.")
        self._append_log("")
        self._append_log("Cancelled - no output file was written.", "error")
        self.progress.configure(value=0)
        self._set_running(False)

    def _on_failed(self, message):
        self.status.set("Failed.")
        self._append_log("")
        self._append_log(f"Failed: {message}", "error")
        self.progress.configure(value=0)
        self._set_running(False)
        messagebox.showerror("Reconciliation failed", message, parent=self)

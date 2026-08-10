"""
VAT Reconciliation - stand-alone launcher
=========================================
Double-click this file to run the reconciliation tool on its own (the .pyw
extension launches it without a console window).

IMPORTANT - this file is not self-contained. It needs the other program files
listed in REQUIRED_MODULES below sitting in the same folder. Copy the whole
folder when sharing the tool, not just this one file.

The screen itself lives in vat_reconciliation_tab.VatReconciliationTab, which
is a plain ttk.Frame. To fold it into the main application instead, drop the
tab straight into that app's notebook:

    from vat_reconciliation_tab import VatReconciliationTab

    notebook.add(VatReconciliationTab(notebook), text="VAT Reconciliation")

This launcher is then no longer needed.
"""

import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk

APP_DIR = os.path.dirname(os.path.abspath(__file__))

# Allow launching from any working directory (e.g. a desktop shortcut).
sys.path.insert(0, APP_DIR)

# The program's own files, which must sit next to this launcher.
REQUIRED_MODULES = [
    "vat_reconciliation_tab.py",
    "vat_reconciliation_core.py",
    "vat_source_loader.py",
    "vat_summary_template_data.py",
]

# Third-party packages, mapped to the name used to install them.
REQUIRED_PACKAGES = {
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "requests": "requests",
    "xlrd": "xlrd",
}


def find_missing_files():
    return [name for name in REQUIRED_MODULES
            if not os.path.exists(os.path.join(APP_DIR, name))]


def find_missing_packages():
    import importlib.util

    missing = []
    for module_name, install_name in REQUIRED_PACKAGES.items():
        if importlib.util.find_spec(module_name) is None:
            missing.append(install_name)
    return missing


def show_error(title, message):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(title, message)
    root.destroy()


def main():
    # Check the program's own files first. Copying only the .pyw is the most
    # common way to break this, and the resulting ImportError looks exactly
    # like a missing third-party package unless the two are told apart.
    missing_files = find_missing_files()
    if missing_files:
        show_error(
            "Program files missing",
            "This launcher needs the rest of the program's files in the same "
            "folder, and these are missing:\n\n"
            + "\n".join(f"    {name}" for name in missing_files)
            + f"\n\nLooked in:\n    {APP_DIR}\n\n"
            "Copy the whole application folder, not just VAT_Reconciliation.pyw.",
        )
        return

    missing_packages = find_missing_packages()
    if missing_packages:
        show_error(
            "Missing Python packages",
            "These Python packages need to be installed:\n\n"
            + "\n".join(f"    {name}" for name in missing_packages)
            + "\n\nInstall them by running this at a command prompt:\n\n"
            f"    pip install {' '.join(missing_packages)}",
        )
        return

    try:
        from vat_reconciliation_tab import VatReconciliationTab
    except Exception as exc:
        show_error(
            "Could not start",
            f"The application failed to load:\n\n{type(exc).__name__}: {exc}",
        )
        return

    root = tk.Tk()
    root.title("VAT Reconciliation")
    root.geometry("880x620")
    root.minsize(720, 520)

    # A notebook even with one tab, so the layout matches how this screen will
    # sit inside the main application.
    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True, padx=6, pady=6)
    notebook.add(VatReconciliationTab(notebook), text="VAT Reconciliation")

    root.mainloop()


if __name__ == "__main__":
    main()

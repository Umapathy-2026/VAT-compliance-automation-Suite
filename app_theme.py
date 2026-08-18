"""
VAT Compliance Automation - shared UI theme
============================================
The color palette, fonts and widget-building helpers used across every page
of VAT_Compliance_App_V6.pyw, factored out so other modules (e.g. the VAT
Returns / Spain reconciliation page) can build widgets that look identical
without duplicating this code.

Import `configure_style()` once, after the Tk root is created, before
building any pages - it registers the ttk styles (Combobox, Entry, Treeview,
Progressbar, Radiobutton) that the helpers below and any themed ttk widgets
rely on.
"""

import tkinter as tk
from tkinter import ttk


class C:
    # Brand orange, unchanged from the Johnson Electric mark - every other
    # token here is chosen to make this color read as sharp and deliberate
    # rather than the whole UI leaning on it for contrast.
    ACCENT = "#F2711C"
    ACCENT_DARK = "#D35400"
    ACCENT_DEEP = "#B84A0F"       # pressed/emphasis state, one step past hover
    ACCENT_LIGHT = "#FDE9DA"
    SIDEBAR_BG = "#F2711C"
    SIDEBAR_BG_HOVER = "#E8660F"  # inactive-row hover, one step darker than the sidebar
    SIDEBAR_ACTIVE = "#FFFFFF"
    SIDEBAR_TEXT = "#FBDCC7"
    SIDEBAR_TEXT_ACTIVE = "#1C1917"
    PAGE_BG = "#FAF8F6"    # a hair off white, so white cards read as raised
    CARD_BG = "#FFFFFF"
    CARD_BORDER = "#E9E2DB"
    # Warm near-black (echoes the logo's wordmark) instead of a cool gray -
    # pairs more deliberately with the orange than a generic slate would.
    TEXT_DARK = "#1C1917"
    TEXT_MUTED = "#78716C"
    SUCCESS_BG = "#EAF7ED"
    SUCCESS_TEXT = "#1E7B34"
    ERROR_BG = "#FDEDEC"
    ERROR_TEXT = "#B3261E"
    WARN_BG = "#FFF7E6"
    WARN_TEXT = "#8A5B00"
    LOG_BG = "#FAFAF9"
    LOG_TEXT = "#3A3532"


FONT_TITLE = ("Segoe UI Semibold", 22)
FONT_SUBTITLE = ("Segoe UI", 10)
FONT_CARD_TITLE = ("Segoe UI Semibold", 12)
FONT_BODY = ("Segoe UI", 10)
FONT_BODY_BOLD = ("Segoe UI Semibold", 10)
FONT_NAV = ("Segoe UI", 11)
FONT_NAV_ACTIVE = ("Segoe UI Semibold", 11)
FONT_LOGO = ("Segoe UI Semibold", 13)
FONT_MONO = ("Consolas", 9)


def configure_style():
    style = ttk.Style()
    style.theme_use("clam")

    style.configure("TCombobox", fieldbackground="white", background="white",
                     bordercolor=C.CARD_BORDER, arrowsize=14, arrowcolor=C.TEXT_MUTED)
    style.map("TCombobox", bordercolor=[("focus", C.ACCENT)],
              arrowcolor=[("active", C.ACCENT)])
    style.configure("TEntry", fieldbackground="white", bordercolor=C.CARD_BORDER)
    style.map("TEntry", bordercolor=[("focus", C.ACCENT)])

    style.configure("Accent.Horizontal.TProgressbar",
                     troughcolor="#F1F1EF", background=C.ACCENT,
                     bordercolor="#F1F1EF", lightcolor=C.ACCENT, darkcolor=C.ACCENT,
                     thickness=10)

    style.configure("Card.TRadiobutton", background=C.CARD_BG, font=FONT_BODY)
    style.map("Card.TRadiobutton", background=[("active", C.CARD_BG)])

    style.configure("Treeview", rowheight=26, font=FONT_BODY, fieldbackground="white",
                     background="white", bordercolor=C.CARD_BORDER, borderwidth=1)
    style.configure("Treeview.Heading", font=FONT_BODY_BOLD, background=C.ACCENT_LIGHT,
                     foreground=C.ACCENT_DARK, relief="flat")
    style.map("Treeview.Heading", background=[("active", C.ACCENT_LIGHT)])
    style.map("Treeview", background=[("selected", C.ACCENT_LIGHT)],
              foreground=[("selected", C.TEXT_DARK)])


def make_button(parent, text, command, kind="primary", state="normal", width=None):
    """A flat, hover-aware button styled to match the reference UI (ttk buttons can't
    take arbitrary flat colors reliably across platforms, so a tk.Button is used)."""
    palette = {
        "primary":   dict(bg=C.ACCENT, fg="white", hover=C.ACCENT_DARK,
                           press=C.ACCENT_DEEP, border=C.ACCENT_DARK),
        "secondary": dict(bg="white", fg=C.ACCENT, hover=C.ACCENT_LIGHT,
                           press=C.ACCENT_LIGHT, border=C.ACCENT),
        "muted":     dict(bg="#F3F1EF", fg=C.TEXT_DARK, hover="#E8E5E1",
                           press="#DEDAD5", border=C.CARD_BORDER),
    }[kind]
    btn = tk.Button(
        parent, text=text, command=command, state=state,
        bg=palette["bg"], fg=palette["fg"], activebackground=palette["press"],
        activeforeground=palette["fg"], relief="flat", bd=0,
        font=FONT_BODY_BOLD, padx=16, pady=8, cursor="hand2",
        highlightthickness=1, highlightbackground=C.CARD_BORDER,
        disabledforeground="#B0AAA3",
    )
    if width:
        btn.configure(width=width)

    def on_enter(_e):
        if btn["state"] != "disabled":
            btn.configure(bg=palette["hover"], highlightbackground=palette["border"])

    def on_leave(_e):
        if btn["state"] != "disabled":
            btn.configure(bg=palette["bg"], highlightbackground=C.CARD_BORDER)

    def on_press(_e):
        if btn["state"] != "disabled":
            btn.configure(bg=palette["press"])

    def on_release(_e):
        if btn["state"] != "disabled":
            btn.configure(bg=palette["hover"])

    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    btn.bind("<ButtonPress-1>", on_press)
    btn.bind("<ButtonRelease-1>", on_release)
    return btn


def _card_shell(parent):
    """The bordered white panel every card sits in, with a thin accent strip
    across its top edge - the one piece of brand color that recurs through
    every content page, not just the sidebar chrome."""
    outer = tk.Frame(parent, bg=C.CARD_BG, highlightbackground=C.CARD_BORDER,
                      highlightthickness=1, bd=0)
    outer.pack(fill="x", pady=(0, 16))
    tk.Frame(outer, bg=C.ACCENT, height=3).pack(fill="x", side="top")
    return outer


def make_card(parent, title, icon=""):
    """A white bordered panel with a title row. Returns the inner content frame."""
    outer = _card_shell(parent)

    header = tk.Frame(outer, bg=C.CARD_BG)
    header.pack(fill="x", padx=20, pady=(14, 8))
    tk.Label(header, text=f"{icon}  {title}".strip(), bg=C.CARD_BG, fg=C.TEXT_DARK,
              font=FONT_CARD_TITLE).pack(anchor="w")

    body = tk.Frame(outer, bg=C.CARD_BG)
    body.pack(fill="both", expand=True, padx=20, pady=(0, 18))
    return body


def make_paste_box(parent, height=7, hint=""):
    if hint:
        tk.Label(parent, text=hint, bg=C.CARD_BG, fg=C.TEXT_MUTED, font=FONT_BODY,
                  anchor="w", justify="left").pack(fill="x", pady=(0, 6))
    frame = tk.Frame(parent, bg="white", highlightbackground=C.CARD_BORDER, highlightthickness=1)
    frame.pack(fill="x")
    text = tk.Text(frame, height=height, bg="white", fg=C.TEXT_DARK, font=FONT_MONO,
                    bd=0, padx=10, pady=8, wrap="none")
    scrollbar = ttk.Scrollbar(frame, command=text.yview)
    text.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    text.pack(side="left", fill="both", expand=True)
    return text


def make_log_box(parent, height=10):
    frame = tk.Frame(parent, bg=C.LOG_BG, highlightbackground=C.CARD_BORDER, highlightthickness=1)
    frame.pack(fill="both", expand=True)
    text = tk.Text(frame, height=height, state="disabled", bg=C.LOG_BG, fg=C.LOG_TEXT,
                    font=FONT_MONO, bd=0, padx=10, pady=8, wrap="word")
    scrollbar = ttk.Scrollbar(frame, command=text.yview)
    text.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    text.pack(side="left", fill="both", expand=True)
    return text


def append_log(text_widget, line):
    text_widget.configure(state="normal")
    text_widget.insert("end", line + "\n")
    text_widget.see("end")
    text_widget.configure(state="disabled")


def clear_log(text_widget):
    text_widget.configure(state="normal")
    text_widget.delete("1.0", "end")
    text_widget.configure(state="disabled")


def make_table(parent, columns, widths=None, height=10):
    """A styled Treeview results table with color-coded row tags."""
    frame = tk.Frame(parent, bg=C.CARD_BG)
    frame.pack(fill="both", expand=True)
    tree = ttk.Treeview(frame, columns=columns, show="headings", height=height)
    for i, col in enumerate(columns):
        tree.heading(col, text=col)
        width = widths[i] if widths and i < len(widths) else 140
        tree.column(col, width=width, anchor="w")
    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    vsb.pack(side="right", fill="y")
    hsb.pack(side="bottom", fill="x")
    tree.pack(side="left", fill="both", expand=True)

    tree.tag_configure("good", background=C.SUCCESS_BG, foreground=C.SUCCESS_TEXT)
    tree.tag_configure("bad", background=C.ERROR_BG, foreground=C.ERROR_TEXT)
    tree.tag_configure("warn", background=C.WARN_BG, foreground=C.WARN_TEXT)
    return tree


def clear_table(tree):
    for item in tree.get_children():
        tree.delete(item)


def labeled_field(parent, label_text, row):
    tk.Label(parent, text=label_text, bg=C.CARD_BG, fg=C.TEXT_DARK,
              font=FONT_BODY_BOLD).grid(row=row, column=0, sticky="w", pady=6, padx=(0, 12))


def card_header_row(parent, title, icon=""):
    """Card title on the left, action buttons area on the right — returns (card_body, actions_frame)."""
    outer = _card_shell(parent)
    outer.pack_configure(fill="both", expand=True)

    header = tk.Frame(outer, bg=C.CARD_BG)
    header.pack(fill="x", padx=20, pady=(14, 8))
    tk.Label(header, text=f"{icon}  {title}".strip(), bg=C.CARD_BG, fg=C.TEXT_DARK,
              font=FONT_CARD_TITLE).pack(side="left")
    actions = tk.Frame(header, bg=C.CARD_BG)
    actions.pack(side="right")

    body = tk.Frame(outer, bg=C.CARD_BG)
    body.pack(fill="both", expand=True, padx=20, pady=(0, 18))
    return body, actions


def open_path(path):
    """Open a file or folder with the OS default handler (Windows-focused, cross-platform safe)."""
    import os
    path = str(path)
    try:
        os.startfile(path)  # type: ignore[attr-defined]
    except AttributeError:
        import subprocess
        import sys
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, path])

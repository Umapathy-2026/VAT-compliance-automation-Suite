# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for VAT Compliance Automation. Freezes
VAT_Compliance_App_V8.pyw (and its sibling modules app_theme.py,
vat_returns_page.py, vat_reconciliation_core.py, vat_source_loader.py,
vat_summary_template_data.py) plus every third-party dependency - openpyxl,
requests, requests-ntlm, urllib3, pandas, numpy, xlrd, and optionally
Pillow - into a single self-contained VATComplianceAutomation.exe.

Built onefile rather than onedir by explicit request, since the plan to
wrap this in an MSI via WiX was dropped - this build machine itself has the
same Application Control policy blocking execution of any newly-created/
unsigned binary (confirmed against both wix.exe and an onedir build of this
same app), so authoring an MSI here isn't possible either. The .exe is
instead handed directly to IT, who will install/trust it through their own
channel. Onefile's trade-off is a slower launch (it self-extracts to a temp
folder each run) in exchange for being a single file to hand off - not
relevant to Application Control either way once IT's process is what
grants trust, not where these bytes happen to sit.

Build with (run from the repo root):
    python -m PyInstaller packaging/VATComplianceAutomation.spec --distpath packaging/dist --workpath packaging/build --noconfirm
"""

import os

APP_ROOT = os.path.abspath(os.path.join(SPECPATH, ".."))
ICON_PATH = os.path.join(APP_ROOT, "tools", "build_assets", "app_icon.ico")
APP_NAME = "VATComplianceAutomation"

block_cipher = None

a = Analysis(
    [os.path.join(APP_ROOT, "VAT_Compliance_App_V8.pyw")],
    pathex=[APP_ROOT],
    binaries=[],
    datas=[
        (os.path.join(APP_ROOT, "logo.png"), "."),
    ],
    # PIL is imported inside a try/except ImportError in load_sidebar_logo()
    # (an intentionally optional dependency in the source) - forced in here
    # explicitly since PyInstaller's static analysis is less reliable for
    # imports that only happen conditionally at runtime.
    hiddenimports=["PIL", "PIL.Image", "PIL.ImageTk"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# Onefile: binaries/zipfiles/datas are passed straight into EXE() (and
# exclude_binaries left False) instead of going through a separate COLLECT
# step - that's what folds everything into the one .exe.
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPX-compressed binaries are *more* likely to trip AV/EDR heuristics,
    # not less - counterproductive for an app that already hit an
    # Application Control block. Left off deliberately.
    upx=False,
    console=False,
    icon=ICON_PATH,
)

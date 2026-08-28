# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all


hiddenimports = []
datas = [("reference", "reference")]
binaries = []

# app.py stages these imports after the window is visible, so PyInstaller cannot
# discover all of their modules through ordinary static import analysis.
for package in (
    "cv2",
    "google.auth",
    "google_auth_oauthlib",
    "gspread",
    "numpy",
    "openpyxl",
    "pandas",
    "pdf2image",
    "PIL",
    "pytesseract",
    "rapidfuzz",
):
    package_datas, package_binaries, package_hiddenimports = collect_all(package)
    datas += package_datas
    binaries += package_binaries
    hiddenimports += package_hiddenimports


analysis = Analysis(
    ["app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(analysis.pure)
exe = EXE(
    pyz,
    analysis.scripts,
    [],
    exclude_binaries=True,
    name="Quiz Processing System",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=None,
)
collection = COLLECT(
    exe,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="Quiz Processing System",
)

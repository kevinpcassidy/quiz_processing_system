# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
import os


hiddenimports = []
required_release_paths = (
    "google_oauth_client.json",
    "reference",
    "vendor/tesseract/tesseract.exe",
    "vendor/tesseract/tessdata/eng.traineddata",
    "vendor/poppler/Library/bin/pdfinfo.exe",
    "vendor/poppler/Library/bin/pdftoppm.exe",
    "LICENSE",
    "THIRD_PARTY_LICENSES.txt",
)
missing_release_paths = [path for path in required_release_paths if not os.path.exists(path)]
if missing_release_paths:
    raise SystemExit(
        "Release resources are missing:\n- " + "\n- ".join(missing_release_paths)
    )

datas = [
    ("reference", "reference"),
    ("google_oauth_client.json", "."),
    ("vendor/tesseract", "vendor/tesseract"),
    ("vendor/poppler", "vendor/poppler"),
    ("vendor_docs", "vendor_docs"),
    ("LICENSE", "."),
    ("THIRD_PARTY_LICENSES.txt", "."),
]
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
    upx=False,
    console=False,
    contents_directory=".",
    icon=None,
)
collection = COLLECT(
    exe,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="Quiz Processing System",
)

# Quiz Processing System

Quiz Processing System extracts topic scores from quiz PDFs and exports CSV or Excel results. It supports optional local gradebooks, an optional Google Sheets gradebook created by the application, and a tabbed Excel export of current rosters for Microsoft Word mail merge.

Google Sheets developer setup and the complete test checklist are documented in [`developer_implementation.md`](developer_implementation.md). Do not use or distribute a service-account key; the application uses per-teacher Desktop OAuth with the `drive.file` scope.

## Run and test from source

Python 3.11.x is the supported interpreter for source development and Windows
release builds. Create an isolated environment and install the direct development
dependencies from `requirements.txt`:

```powershell
py -3.11 -m venv .venv
.venv\Scripts\python -m pip install --upgrade pip
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python app.py
```

`requirements.txt` intentionally lists direct dependencies without version
pins so contributors can test compatible updates. To reproduce the Python
application dependencies shipped in release 1.0.0, install
`release-requirements-1.0.0.txt` instead. That release snapshot does not include
the separately installed PyInstaller build tool, which is pinned in
`build-requirements.txt`.

PDF conversion requires [Poppler](https://poppler.freedesktop.org/) and OCR
requires [Tesseract](https://github.com/tesseract-ocr/tesseract). Their
command-line programs must be available on `PATH`, unless the checked-in
`vendor/` release trees are present; those trees take precedence. On Windows,
the required programs are `pdfinfo.exe` and `pdftoppm.exe` from Poppler and
`tesseract.exe` from Tesseract.

The source entry point is `app.py`. Run the automated checks with:

```text
python -m py_compile app.py
python -m unittest discover -s tests -v
```

## Build the Windows release

The supported packaged release target is 64-bit Windows 10 or 11. Prepare the
following untracked, release-only files in the repository root:

```text
google_oauth_client.json
vendor/tesseract/tesseract.exe
vendor/tesseract/tessdata/eng.traineddata
vendor/poppler/Library/bin/pdfinfo.exe
vendor/poppler/Library/bin/pdftoppm.exe
```

Keep the complete downloaded `vendor/tesseract/` and `vendor/poppler/` trees;
do not copy only the executables, because their DLLs and license materials are
required. The OAuth file must be a Google Desktop application client. It is
embedded in the release but remains excluded from Git.

Build the one-directory application from a 64-bit Windows Python 3.11.x
environment. For release 1.0.0, create a clean environment and install its
versioned application dependency snapshot before installing PyInstaller:

```powershell
py -3.11 -m venv .venv-release
.venv-release\Scripts\python -m pip install --upgrade pip
.venv-release\Scripts\python -m pip install -r release-requirements-1.0.0.txt
.venv-release\Scripts\python -m pip install -r build-requirements.txt
.venv-release\Scripts\python -m PyInstaller --clean --noconfirm quiz_processing_system.spec
```

See [`BUILDING.md`](BUILDING.md) for local and GitHub Actions build procedures,
validation details, artifact download instructions, and the portable-system
test checklist.

The spec validates release inputs, bundles the complete `reference/` directory
automatically (including `reference/SAMPLE.pdf` when present), and produces a
directory named `Quiz Processing System`. Before publishing its ZIP, follow the
release-builder checklist in `THIRD_PARTY_LICENSES.txt`. The project is
MIT-licensed; bundled components remain subject to their own licenses.

## Versions and updates

The Home panel displays the installed application version and checks the public
[GitHub Releases page](https://github.com/kevinpcassidy/quiz_processing_system/releases)
for the latest stable release in the background. When a newer version is available,
the download action prefers the versioned Windows ZIP attached to the release and
falls back to the release page.

Mutable user data is stored separately under
`%LOCALAPPDATA%\quiz_processing_system`, so the packaged program directory can be
replaced during an update without deleting classes, settings, grading scales,
roster caches, or Google authorization. See the release procedure in
[`developer_implementation.md`](developer_implementation.md).

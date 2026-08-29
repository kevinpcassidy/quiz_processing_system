# Quiz Processing System

Quiz Processing System extracts topic scores from quiz PDFs and exports CSV or Excel results. It supports optional local gradebooks, an optional Google Sheets gradebook created by the application, and a tabbed Excel export of current rosters for Microsoft Word mail merge.

Google Sheets developer setup and the complete test checklist are documented in [`developer_implementation.md`](developer_implementation.md). Do not use or distribute a service-account key; the application uses per-teacher Desktop OAuth with the `drive.file` scope.

## Run and test from source

Install the dependencies from `requirements.txt`. PDF conversion requires
[Poppler](https://poppler.freedesktop.org/) and OCR requires
[Tesseract](https://github.com/tesseract-ocr/tesseract); install both separately
and make their command-line programs available on `PATH`. On Windows, this
means `pdfinfo.exe` and `pdftoppm.exe` from Poppler and `tesseract.exe` from
Tesseract. A complete local `vendor/` release tree takes precedence when it is
present, but it is intentionally not stored in Git.

Then run:

```text
python app.py
```

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

Build the one-directory application from a 64-bit Windows Python environment:

```text
pyinstaller --clean --noconfirm quiz_processing_system.spec
```

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

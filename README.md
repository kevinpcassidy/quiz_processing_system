# Quiz Processing System

Quiz Processing System extracts topic scores from quiz PDFs and exports CSV or Excel results. It supports optional local gradebooks, an optional Google Sheets gradebook created by the application, and a tabbed Excel export of current rosters for Microsoft Word mail merge.

Google Sheets developer setup and the complete test checklist are documented in [`developer_implementation.md`](developer_implementation.md). Do not use or distribute a service-account key; the application uses per-teacher Desktop OAuth with the `drive.file` scope.

## Run and test from source

Install the dependencies from `requirements.txt`, then run:

```text
python app.py
```

The source entry point is `app.py`. Run the automated checks with:

```text
python -m py_compile app.py
python -m unittest discover -s tests -v
```

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

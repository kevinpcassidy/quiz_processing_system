# Quiz Processing System: Google Sheets implementation guide

## Architecture

Quiz Processing System supports two independent workflows:

- **Local classes and Local Gradebooks** use CSV roster caches on the computer.
- **Google Sheets classes** use an application-created spreadsheet as the authority. Each class maps to one worksheet tab. At launch (and when **Refresh Rosters Now** is clicked), column A is copied to a local CSV cache so extraction does not repeatedly call Google.

The OAuth grant uses only `https://www.googleapis.com/auth/drive.file`. The program can access files it creates; it does not enumerate or read unrelated Drive files. Creating a replacement gradebook never deletes the former file.

## Google Cloud setup

1. Create a dedicated project in [Google Cloud Console](https://console.cloud.google.com/).
2. Enable **Google Sheets API** and **Google Drive API**.
3. Configure the Google Auth Platform/OAuth consent screen:
   - Application name: **Quiz Processing System**.
   - Audience: **External** for teachers outside your Workspace organization.
   - Add support and developer-contact email addresses.
   - Request only `https://www.googleapis.com/auth/drive.file`.
4. During development, add each tester as a test user. Before conference distribution, move the app to Production and complete any Google publishing requirements shown by the console.
5. Create an OAuth client with application type **Desktop app**.
6. Download the client JSON, rename it exactly `google_oauth_client.json`, and place it beside `app.py` (or beside the packaged executable/resource bundle).
7. Never distribute a service-account key or any user's `google_token.json`.

A Desktop OAuth client configuration identifies the distributed application; it is not a protected server credential. Each teacher authorizes their own account and receives a private token on their computer.

## Windows data locations

Mutable user data is stored under:

```text
%LOCALAPPDATA%\quiz_processing_system\
├── google_token.json
├── quiz_settings.json
├── saved_classes.json
└── rosters\
```

`google_token.json` is sensitive. Never commit, email, log, or package it. Reconnection preserves settings and spreadsheet/class IDs. The application copies legacy settings, classes, and roster CSVs into this directory on first launch without deleting the originals.

## Expected spreadsheet format

The application creates `CURRENTYEAR-NEXTYEAR Quiz Processing System`, adding `_1`, `_2`, and so on when an app-created file already uses that title. Its initial tab is `Roster 1`, with `Name` in A1, followed by a `SAMPLE` tab copied from the bundled example workbook. The sample can be mapped as a practice class. If the sample cannot be copied, gradebook creation still succeeds and displays a warning. Local roster CSV files use the same one-column format.

- Create one worksheet tab per roster.
- Put `Name` in A1 and one student per row in column A.
- Use a unique, non-empty header for every used score column.
- Do not merge cells in the roster/grade area.
- Teachers may rename the spreadsheet or tabs; stable Google IDs preserve the link.
- Existing topic columns are updated. New topics are appended after the last used header.
- Edit Google-backed students in column A, then refresh rosters or restart.

## Sample walkthrough

Every launch shows a suppressible sample walkthrough unless the user disables it. Advanced Settings can re-enable it. The walkthrough links the bundled `SAMPLE` worksheet, the `SAMPLE` grading scale (`5, 6, 7, 8, 9, 10`), and the forthcoming `reference/SAMPLE.pdf`. Teachers who do not use Google Sheets can instead import the forthcoming `reference/SAMPLE.csv` as a local roster. The **Use SAMPLE.pdf** action reports a clear warning while that file is absent.

## Development checks

Run from the repository root:

```powershell
python -m py_compile app.py
python -m unittest discover -s tests -v
```

For a clean Windows integration test, temporarily move `%LOCALAPPDATA%\quiz_processing_system` aside and restore it afterward.

## Manual testing checklist

### Installation and migration

- [ ] Launch as a new Windows user; `%LOCALAPPDATA%\quiz_processing_system` is created.
- [ ] Launch with legacy `quiz_settings.json`, `saved_classes.json`, and `rosters`; data is copied without deleting originals.
- [ ] Missing `google_oauth_client.json` produces a clear error rather than a crash.
- [ ] No OAuth token appears in the repo or application logs.

### Opt-in and OAuth

- [ ] Selecting Google integration opens the warning and 500×500 help image.
- [ ] **Cancel** and the window close button both clear the checkbox.
- [ ] **Proceed** reveals authorization controls and persists the preference.
- [ ] OAuth denial leaves settings intact and reports the error.
- [ ] Successful OAuth writes the token only under `%LOCALAPPDATA%\quiz_processing_system`.
- [ ] A refreshable expired token refreshes silently at launch.
- [ ] A revoked/non-refreshable token offers reconnection while preserving settings.
- [ ] **Reconnect / Change Google Account** works.

### Gradebook creation

- [ ] The first title follows `CURRENTYEAR-NEXTYEAR Quiz Processing System`.
- [ ] A duplicate app-created title receives `_1`, then `_2`.
- [ ] The initial tab is `Roster 1` with `Name` in A1, followed by the formatted `SAMPLE` tab.
- [ ] A missing or invalid sample workbook warns the user without failing gradebook creation.
- [ ] Creation opens setup guidance; **About Google Sheets** reopens it.
- [ ] Renaming the spreadsheet is detected at next launch and does not break its link.
- [ ] **Open Google Sheets Gradebook** opens the correct stable ID.
- [ ] Replacement warning explains that the old file is retained.
- [ ] Canceling replacement preserves the current link.
- [ ] Failed replacement preserves the current link and mappings.
- [ ] Successful replacement marks old Google class mappings for remapping and leaves local classes untouched.

### Sample walkthrough

- [ ] `SAMPLE` appears in the grading-scale selector with scores `5, 6, 7, 8, 9, 10`.
- [ ] An existing grading scale named `SAMPLE` is replaced with the canonical sample scores at launch.
- [ ] The startup walkthrough describes both the Google Sheets sample and local `SAMPLE.csv` alternative.
- [ ] **Use SAMPLE.pdf** selects `reference/SAMPLE.pdf` when present and warns clearly when absent.
- [ ] **Don't show this walkthrough again** persists after restarting.
- [ ] Advanced Settings can re-enable the startup walkthrough, and **Restore Defaults** enables it.

### Classes and roster cache

- [ ] Local class import accepts CSV/Excel.
- [ ] **Add or Remove Students** remains available for a local class.
- [ ] A Google class maps a course name to a worksheet tab.
- [ ] Attempting student editing for a Google class directs the teacher to column A.
- [ ] Deleting a Google class does not delete its tab or spreadsheet.
- [ ] Double-click preview works for both local and cached Google rosters.
- [ ] Startup refresh reads column A and row 1 in the background.
- [ ] Manual refresh updates the local CSV atomically.
- [ ] Missing `Name` in A1 rejects refresh and preserves the previous cache.
- [ ] Offline/API failure preserves the previous cache.
- [ ] A renamed tab remains mapped by worksheet ID and its displayed title updates.
- [ ] A deleted tab reports failure without deleting the cache.

### Extraction, export, and synchronization

- [ ] Extraction uses the refreshed local CSV without per-student Google calls.
- [ ] CSV export succeeds.
- [ ] Excel `.xlsx` export succeeds.
- [ ] Local Gradebook can be enabled and updated independently.
- [ ] Google Sheets can be enabled independently.
- [ ] Both update actions can be shown and used in the same extraction.
- [ ] Sync rereads row 1 immediately before writing.
- [ ] Existing topic headers are reused case-insensitively.
- [ ] New topic headers append after the last used header.
- [ ] Only mapped roster names receive scores; extraction-only names are not silently added.
- [ ] Unrelated columns and rows are not overwritten.
- [ ] Edited preview values, blank values, and `skip` synchronize correctly.
- [ ] Failed sync leaves the button available for retry.


## Home panel and mail merge

The Home panel provides **Export Current Rosters for Mail Merge**. It creates one Excel workbook with one sanitized worksheet per current class and a single `Name` column. Both the left workflow panel and center Home panel use auto-hiding vertical scrollbars: mouse-wheel input is accepted only while the pointer is over overflowing content, and movement is clamped at the first and final content rows.

## Extraction safety

For Google-backed classes, the program snapshots the exact cached roster sequence when calibration begins. It rereads column A before any score or header write and refuses synchronization if the sequence changed or contains duplicate names. A suppressible pre-extraction confirmation is stored in preferences; the active extraction warning is always visible. After a complete roster refresh, the UI stores a timezone-aware timestamp and displays it using the Windows computer's local time.

Additional manual checks:

- [ ] The left scrollbar appears only after enough topics are added to overflow.
- [ ] The center Home scrollbar appears only when content overflows.
- [ ] Mouse-wheel scrolling stops exactly at the top and bottom and does not move unrelated Treeviews.
- [ ] Mail-merge export contains one `Name` worksheet per local and Google-backed class.
- [ ] Invalid/duplicate/overlong class names become unique valid Excel worksheet names.
- [ ] Local imports reject files without `Name` in A1 and save only the first-column names.
- [ ] The OAuth dialog exposes the exact temporary URL with Copy Link and Open buttons.
- [ ] Suppressing the safety confirmation persists, and Advanced Settings can restore it.
- [ ] Changing Google column A after extraction starts prevents every Sheets write.
- [ ] Complete refreshes update the local timestamp; partial refreshes preserve the prior complete time.

## Application releases

`app.py` is the source entry point. `APP_VERSION` in that file is the authoritative
application version and must match the GitHub Release tag (for example, application
version `1.0.0` uses tag `v1.0.0`). Normal update checks use only the latest stable,
non-draft GitHub Release.

Build the Windows one-directory distribution from the repository root:

```powershell
python -m pip install pyinstaller
pyinstaller --clean --noconfirm quiz_processing_system.spec
Compress-Archive -Path "dist\Quiz Processing System\*" -DestinationPath "Quiz-Processing-System-1.0.0-Windows.zip"
```

Attach the versioned ZIP to the matching GitHub Release. The application prefers a
`.zip` release asset containing `Windows` in its filename and falls back to the
release page if no suitable asset exists. Before publishing, launch the packaged
executable, exercise PDF/OCR and export features, and verify that the Home panel
reports the released version.

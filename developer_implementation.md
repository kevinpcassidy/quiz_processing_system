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
6. Download the client JSON, rename it exactly `google_oauth_client.json`, and place it beside `quiz_pipeline_gui_v6_personal.py` (or beside the packaged executable/resource bundle).
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

The application creates `CURRENTYEAR-NEXTYEAR Topic Quiz Grades`, adding `_1`, `_2`, and so on when an app-created file already uses that title. Its initial tab is `Roster 1`, with `Name` in A1.

- Create one worksheet tab per roster.
- Put `Name` in A1 and one student per row in column A.
- Use a unique, non-empty header for every used score column.
- Do not merge cells in the roster/grade area.
- Teachers may rename the spreadsheet or tabs; stable Google IDs preserve the link.
- Existing topic columns are updated. New topics are appended after the last used header.
- Edit Google-backed students in column A, then refresh rosters or restart.

## Development checks

Run from the repository root:

```powershell
python -m py_compile quiz_pipeline_gui_v6_personal.py
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

- [ ] The first title follows `CURRENTYEAR-NEXTYEAR Topic Quiz Grades`.
- [ ] A duplicate app-created title receives `_1`, then `_2`.
- [ ] The initial tab is `Roster 1` and A1 is `Name`.
- [ ] Creation opens setup guidance; **About Google Sheets** reopens it.
- [ ] Renaming the spreadsheet is detected at next launch and does not break its link.
- [ ] **Open Google Sheets Gradebook** opens the correct stable ID.
- [ ] Replacement warning explains that the old file is retained.
- [ ] Canceling replacement preserves the current link.
- [ ] Failed replacement preserves the current link and mappings.
- [ ] Successful replacement marks old Google class mappings for remapping and leaves local classes untouched.

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

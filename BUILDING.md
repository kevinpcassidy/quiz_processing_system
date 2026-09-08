# Building Quiz Processing System

## Build inputs

QPS uses 64-bit Python 3.11.x and the existing
`quiz_processing_system.spec` PyInstaller onedir configuration. Runtime Python
packages for release 1.0.0 are pinned in `release-requirements-1.0.0.txt`.
PyInstaller is a build-only dependency pinned in `build-requirements.txt`.

The complete `vendor/tesseract` and `vendor/poppler` distributions are
intentionally committed. Builds must use these files; do not install or
download replacement Tesseract or Poppler versions. The spec also bundles the
complete `reference` and `vendor_docs` trees plus the QPS and third-party
license documents.

## Local Windows build

From PowerShell in the repository root:

```powershell
py -3.11 -m venv .venv-release
.venv-release\Scripts\python -m pip install --upgrade pip
.venv-release\Scripts\python -m pip install -r release-requirements-1.0.0.txt
.venv-release\Scripts\python -m pip install -r build-requirements.txt
```

Place the Desktop OAuth client configuration at
`google_oauth_client.json`, then run:

```powershell
.venv-release\Scripts\python -m pip check
.venv-release\Scripts\python -m unittest discover -s tests -v
.venv-release\Scripts\python -m PyInstaller --clean --noconfirm quiz_processing_system.spec
```

The result is the complete portable directory
`dist\Quiz Processing System`. Do not distribute only its executable.
`google_oauth_client.json` is a required build input and remains ignored by
Git. Never place an end-user `google_token.json`, refresh token, or other user
credential in the build.

## GitHub Actions build

`.github/workflows/build-windows.yml` runs only when manually dispatched. It:

1. checks out the committed source and vendor trees on `windows-latest`;
2. selects Python 3.11 with `actions/setup-python`;
3. installs the pinned release and build requirements;
4. reads the existing `GOOGLE_OAUTH_CLIENT_JSON` repository secret from an
   environment variable, validates it as JSON in memory, and writes
   `google_oauth_client.json` without printing its contents;
5. validates source resources and runs the unit tests;
6. builds the existing PyInstaller onedir spec;
7. checks required packaged resources, compares every packaged vendor,
   `vendor_docs`, and `reference` file to its committed source by SHA-256, and
   executes the packaged copies of Tesseract, `pdftoppm`, and `pdfinfo` by
   explicit path;
8. creates `Quiz_Processing_System_Windows.zip`; and
9. uploads it as the `Quiz_Processing_System_Windows` Actions artifact.

The ZIP contains the `Quiz Processing System` directory at its root. Extract
that directory and run `Quiz Processing System.exe` inside it.

### Run and download

1. Open the repository's **Actions** tab on GitHub.
2. Select **Build Windows portable application**.
3. Choose **Run workflow** and wait for the build job to finish.
4. In the completed workflow run, download the
   **Quiz_Processing_System_Windows** artifact.
5. Extract the downloaded Actions artifact, then extract
   `Quiz_Processing_System_Windows.zip` to obtain the application directory.

The repository secret `GOOGLE_OAUTH_CLIENT_JSON` is already expected to exist;
its value must be the complete working Desktop OAuth client JSON. The workflow
does not create or upload a GitHub Release.

## Portable-distribution test

Test the downloaded ZIP on a Windows computer without globally available
Tesseract or Poppler:

1. Remove their external installation directories from `PATH`, or temporarily
   rename those directories.
2. Confirm `tesseract --version`, `pdftoppm -v`, and `pdfinfo -v` do not work in
   a new terminal.
3. Extract the GitHub-built ZIP into a new directory.
4. Launch QPS and verify PDF import/conversion, OCR, student-name recognition,
   score/circle processing, and CSV/Excel export.
5. Verify Desktop OAuth and Google Sheets integration with a test account.

Passing this test demonstrates that QPS is using its packaged vendor tools
rather than software installed elsewhere on the computer.

## Signing status

SignPath is intentionally not integrated yet. This workflow stops after
producing and uploading the unsigned portable ZIP. A future signing stage can
be inserted after validation and before ZIP creation and release publication.

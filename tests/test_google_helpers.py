import ast
import csv
import json
import os
import re
import tempfile
import unittest
from types import SimpleNamespace
from datetime import datetime
from pathlib import Path

source = Path("app.py").read_text(encoding="utf-8")
module = ast.parse(source)
helper_names = {
    "atomic_write_json",
    "excel_sheet_title",
    "format_google_progress_status",
    "format_local_timestamp",
    "install_sample_grading_scale",
    "normalize_score_row",
    "normalize_score_value",
    "read_roster_names",
    "release_download_url",
    "unique_gradebook_title",
    "version_tuple",
    "write_roster_names",
}
wanted = [node for node in module.body if isinstance(node, ast.FunctionDef) and node.name in helper_names]
namespace = {
    "os": os,
    "json": json,
    "datetime": datetime,
    "csv": csv,
    "re": re,
    "GITHUB_RELEASES_URL": "https://github.com/kevinpcassidy/quiz_processing_system/releases",
    "SAMPLE_GRADING_SCALE": [5, 6, 7, 8, 9, 10],
}
exec(compile(ast.Module(body=wanted, type_ignores=[]), "app.py", "exec"), namespace)
atomic_write_json = namespace["atomic_write_json"]
excel_sheet_title = namespace["excel_sheet_title"]
format_local_timestamp = namespace["format_local_timestamp"]
format_google_progress_status = namespace["format_google_progress_status"]
install_sample_grading_scale = namespace["install_sample_grading_scale"]
normalize_score_row = namespace["normalize_score_row"]
normalize_score_value = namespace["normalize_score_value"]
read_roster_names = namespace["read_roster_names"]
release_download_url = namespace["release_download_url"]
unique_gradebook_title = namespace["unique_gradebook_title"]
version_tuple = namespace["version_tuple"]
write_roster_names = namespace["write_roster_names"]


class GoogleHelperTests(unittest.TestCase):
    def test_score_normalization_preserves_numeric_types(self):
        self.assertEqual(normalize_score_value("10.0"), 10)
        self.assertIsInstance(normalize_score_value("10.0"), int)
        self.assertEqual(normalize_score_value(10.0), 10)
        self.assertEqual(normalize_score_value("7.5"), 7.5)
        self.assertIsInstance(normalize_score_value("7.5"), float)

    def test_score_normalization_preserves_non_scores_and_blanks(self):
        self.assertEqual(normalize_score_value(None), "")
        self.assertEqual(normalize_score_value(" Skip "), "Skip")
        self.assertEqual(normalize_score_row(["00123", "10.0", "7.5"]), ["00123", 10, 7.5])

    def test_google_progress_status_animates_and_reports_slow_connections(self):
        self.assertEqual(
            format_google_progress_status("preparing", 2, 1),
            "Google Sheets: Preparing.",
        )
        self.assertEqual(
            format_google_progress_status("connecting", 5, 3),
            "Google Sheets: Connecting...",
        )
        self.assertEqual(
            format_google_progress_status("connecting", 15, 2),
            "Google Sheets: Still connecting..",
        )

    def test_heavy_dependencies_are_not_imported_at_module_startup(self):
        heavy_packages = {
            "cv2", "google", "google_auth_oauthlib", "gspread", "numpy",
            "openpyxl", "pandas", "pdf2image", "PIL", "pytesseract", "rapidfuzz",
        }
        imported = []
        for node in module.body:
            if isinstance(node, ast.Import):
                imported.extend(alias.name for alias in node.names)
            elif isinstance(node, ast.ImportFrom):
                imported.append(node.module or "")
        eager_heavy_imports = [
            name for name in imported if name.split(".", 1)[0] in heavy_packages
        ]
        self.assertEqual(eager_heavy_imports, [])

    def test_dependency_groups_have_the_expected_preparation_order(self):
        assignment = next(
            node for node in module.body
            if isinstance(node, ast.Assign)
            and any(isinstance(target, ast.Name) and target.id == "DEPENDENCY_ORDER"
                    for target in node.targets)
        )
        self.assertEqual(
            ast.literal_eval(assignment.value),
            ("google", "excel", "pdf", "ocr"),
        )

    def test_unique_gradebook_title_uses_year_and_suffix(self):
        base = "2026-2027 Quiz Processing System"
        self.assertEqual(unique_gradebook_title([], 2026), base)
        self.assertEqual(unique_gradebook_title([base], 2026), f"{base}_1")
        self.assertEqual(unique_gradebook_title([base, f"{base}_1"], 2026), f"{base}_2")

    def test_atomic_write_json_replaces_content_and_cleans_temp_file(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "settings.json")
            atomic_write_json(path, {"old": True})
            atomic_write_json(path, {"new": True})
            with open(path, encoding="utf-8") as handle:
                self.assertEqual(json.load(handle), {"new": True})
            self.assertFalse(os.path.exists(f"{path}.tmp"))

    def test_sample_grading_scale_is_saved_and_replaces_same_named_scale(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "saved_grading_scales.json")
            scales = {"Existing": [1, 2], "SAMPLE": [99]}
            install_sample_grading_scale(path, scales)
            with open(path, encoding="utf-8") as handle:
                saved = json.load(handle)
            self.assertEqual(saved["Existing"], [1, 2])
            self.assertEqual(saved["SAMPLE"], [5, 6, 7, 8, 9, 10])

    def test_roster_round_trip_requires_name_header(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "roster.csv")
            write_roster_names(path, ["Student One", "", " Student Two "])
            self.assertEqual(read_roster_names(path), ["Student One", "Student Two"])
            with open(path, encoding="utf-8-sig") as handle:
                self.assertEqual(next(csv.reader(handle)), ["Name"])

    def test_roster_rejects_missing_name_header(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "roster.csv")
            Path(path).write_text("Student One\nStudent Two\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "A1"):
                read_roster_names(path)

    def test_excel_titles_are_valid_unique_and_limited(self):
        first = excel_sheet_title("Period 1: Math/Support [Long Class Name]", [])
        second = excel_sheet_title("Period 1: Math/Support [Long Class Name]", [first])
        self.assertLessEqual(len(first), 31)
        self.assertNotRegex(first, r"[:\\/?*\[\]]")
        self.assertNotEqual(first.casefold(), second.casefold())

    def test_local_timestamp_has_date_and_time(self):
        formatted = format_local_timestamp("2026-08-25T15:42:18+00:00")
        self.assertIn("August 25, 2026", formatted)
        self.assertRegex(formatted, r"\d{2}:\d{2}:\d{2} [AP]M")

    def test_versions_accept_v_prefix_and_compare_numerically(self):
        self.assertEqual(version_tuple("v1.2.3"), (1, 2, 3))
        self.assertGreater(version_tuple("1.10.0"), version_tuple("1.9.9"))
        with self.assertRaises(ValueError):
            version_tuple("1.2")

    def test_release_download_prefers_windows_zip(self):
        release = {
            "html_url": "https://example.invalid/release",
            "assets": [
                {"name": "source.zip", "browser_download_url": "https://example.invalid/source.zip"},
                {"name": "Quiz-Processing-System-1.1.0-Windows.zip", "browser_download_url": "https://example.invalid/windows.zip"},
            ],
        }
        self.assertEqual(release_download_url(release), "https://example.invalid/windows.zip")

    def test_release_download_falls_back_to_release_page(self):
        release = {"html_url": "https://example.invalid/release", "assets": []}
        self.assertEqual(release_download_url(release), "https://example.invalid/release")


class SampleWorksheetTests(unittest.TestCase):
    class FakeSpreadsheet:
        def __init__(self):
            self.worksheet = SimpleNamespace(id=2468, update_calls=[])
            self.add_calls = []
            self.batch_bodies = []

        def add_worksheet(self, **kwargs):
            self.add_calls.append(kwargs)
            self.worksheet.update = lambda **call: self.worksheet.update_calls.append(call)
            return self.worksheet

        def batch_update(self, body):
            self.batch_bodies.append(body)

    def test_bundled_sample_is_copied_with_values_and_formatting(self):
        from app import QuizAppGUI

        app = QuizAppGUI.__new__(QuizAppGUI)
        app.project_root = os.getcwd()
        spreadsheet = self.FakeSpreadsheet()
        app._add_google_sample_worksheet(spreadsheet)

        self.assertEqual(
            spreadsheet.add_calls,
            [{"title": "SAMPLE", "rows": 37, "cols": 5, "index": 1}],
        )
        values = spreadsheet.worksheet.update_calls[0]["values"]
        self.assertEqual(values[0][0], "Name")
        self.assertEqual(values[3][4], "SAMPLE ROSTER")
        self.assertEqual(values[6][4][:10], "MANDATORY:")
        requests = spreadsheet.batch_bodies[0]["requests"]
        self.assertTrue(any("repeatCell" in request for request in requests))
        self.assertEqual(
            sum("updateDimensionProperties" in request for request in requests), 5
        )

    def test_sample_workbook_is_bundled_from_reference_directory(self):
        self.assertTrue(Path("reference/SAMPLE.xlsx").is_file())
        spec = Path("quiz_processing_system.spec").read_text(encoding="utf-8")
        self.assertIn('("reference", "reference")', spec)


if __name__ == "__main__":
    unittest.main()

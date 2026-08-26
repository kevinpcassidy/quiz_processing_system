import ast
import csv
import json
import os
import re
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

source = Path("quiz_pipeline_gui_v6_personal.py").read_text(encoding="utf-8")
module = ast.parse(source)
helper_names = {
    "atomic_write_json",
    "excel_sheet_title",
    "format_local_timestamp",
    "read_roster_names",
    "unique_gradebook_title",
    "write_roster_names",
}
wanted = [node for node in module.body if isinstance(node, ast.FunctionDef) and node.name in helper_names]
namespace = {"os": os, "json": json, "datetime": datetime, "csv": csv, "re": re}
exec(compile(ast.Module(body=wanted, type_ignores=[]), "quiz_pipeline_gui_v6_personal.py", "exec"), namespace)
atomic_write_json = namespace["atomic_write_json"]
excel_sheet_title = namespace["excel_sheet_title"]
format_local_timestamp = namespace["format_local_timestamp"]
read_roster_names = namespace["read_roster_names"]
unique_gradebook_title = namespace["unique_gradebook_title"]
write_roster_names = namespace["write_roster_names"]


class GoogleHelperTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()

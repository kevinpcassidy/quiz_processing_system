import json
import os
import tempfile
import unittest

import ast
from datetime import datetime
from pathlib import Path

source = Path("quiz_pipeline_gui_v6_personal.py").read_text(encoding="utf-8")
module = ast.parse(source)
wanted = [node for node in module.body if isinstance(node, ast.FunctionDef) and node.name in {"atomic_write_json", "unique_gradebook_title"}]
namespace = {"os": os, "json": json, "datetime": datetime}
exec(compile(ast.Module(body=wanted, type_ignores=[]), "quiz_pipeline_gui_v6_personal.py", "exec"), namespace)
atomic_write_json = namespace["atomic_write_json"]
unique_gradebook_title = namespace["unique_gradebook_title"]


class GoogleHelperTests(unittest.TestCase):
    def test_unique_gradebook_title_uses_year_and_suffix(self):
        base = "2026-2027 Topic Quiz Grades"
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


if __name__ == "__main__":
    unittest.main()

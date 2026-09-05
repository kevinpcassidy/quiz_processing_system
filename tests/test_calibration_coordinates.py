import ast
import math
import unittest
from pathlib import Path


source = Path("app.py").read_text(encoding="utf-8")
module = ast.parse(source)
helper_names = {"display_rectangle_to_source", "display_x_to_source"}
wanted = [
    node for node in module.body
    if isinstance(node, ast.FunctionDef) and node.name in helper_names
]
namespace = {"math": math}
exec(compile(ast.Module(body=wanted, type_ignores=[]), "app.py", "exec"), namespace)
display_rectangle_to_source = namespace["display_rectangle_to_source"]
display_x_to_source = namespace["display_x_to_source"]


class CalibrationCoordinateTests(unittest.TestCase):
    def test_display_rectangle_round_trips_with_outward_pixel_tolerance(self):
        source_size = (1700, 2200)
        displayed_size = (503, 651)
        source_box = (211, 307, 1419, 1803)
        scale_x = displayed_size[0] / source_size[0]
        scale_y = displayed_size[1] / source_size[1]
        displayed_box = (
            source_box[0] * scale_x,
            source_box[1] * scale_y,
            source_box[2] * scale_x,
            source_box[3] * scale_y,
        )

        converted = display_rectangle_to_source(
            displayed_box, source_size, displayed_size
        )

        self.assertLessEqual(source_box[0] - converted[0], 1)
        self.assertLessEqual(source_box[1] - converted[1], 1)
        self.assertLessEqual(converted[2] - source_box[2], 1)
        self.assertLessEqual(converted[3] - source_box[3], 1)

    def test_maximum_boundaries_are_never_contracted(self):
        converted = display_rectangle_to_source(
            (10.2, 20.4, 99.1, 149.2), (1700, 2200), (503, 651)
        )
        exact_right = 99.1 * 1700 / 503
        exact_bottom = 149.2 * 2200 / 651

        self.assertGreaterEqual(converted[2], exact_right)
        self.assertGreaterEqual(converted[3], exact_bottom)

    def test_reverse_drag_produces_the_same_rectangle(self):
        forward = display_rectangle_to_source(
            (10.2, 20.4, 99.1, 149.2), (1700, 2200), (503, 651)
        )
        reverse = display_rectangle_to_source(
            (99.1, 149.2, 10.2, 20.4), (1700, 2200), (503, 651)
        )
        self.assertEqual(forward, reverse)

    def test_resize_does_not_change_a_stored_source_rectangle(self):
        source_size = (1700, 2200)
        stored_box = (200, 300, 1400, 1800)

        for displayed_size in ((503, 651), (710, 919)):
            scale_x = displayed_size[0] / source_size[0]
            scale_y = displayed_size[1] / source_size[1]
            displayed_box = (
                stored_box[0] * scale_x,
                stored_box[1] * scale_y,
                stored_box[2] * scale_x,
                stored_box[3] * scale_y,
            )
            converted = display_rectangle_to_source(
                displayed_box, source_size, displayed_size
            )
            for actual, expected in zip(converted, stored_box):
                self.assertLessEqual(abs(actual - expected), 1)

    def test_full_page_and_crop_use_their_own_display_dimensions(self):
        full_page_x = display_x_to_source(250, 1700, 500)
        crop_x = display_x_to_source(250, 400, 800)

        self.assertEqual(full_page_x, 850)
        self.assertEqual(crop_x, 125)

    def test_crop_width_matches_stored_exclusive_boundaries(self):
        box = display_rectangle_to_source(
            (10.2, 20.4, 99.1, 149.2), (1700, 2200), (503, 651)
        )
        simulated_source_row = list(range(1700))
        simulated_crop = simulated_source_row[box[0]:box[2]]

        self.assertEqual(box, (34, 68, 335, 505))
        self.assertEqual(len(simulated_crop), box[2] - box[0])

    def test_coordinates_are_clamped_to_source_bounds(self):
        self.assertEqual(
            display_rectangle_to_source(
                (-5, -10, 510, 660), (1700, 2200), (503, 651)
            ),
            (0, 0, 1700, 2200),
        )

    def test_invalid_display_dimensions_are_rejected(self):
        with self.assertRaises(ValueError):
            display_rectangle_to_source((0, 0, 1, 1), (10, 10), (0, 10))
        with self.assertRaises(ValueError):
            display_x_to_source(1, 10, 0)


if __name__ == "__main__":
    unittest.main()

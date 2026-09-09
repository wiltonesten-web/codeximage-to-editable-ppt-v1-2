import unittest

from validate_layout_powerpoint import _is_minor_decorative_marker_bounds_difference


class DecorativeMarkerBoundsTests(unittest.TestCase):
    def test_actual_right_arrow_false_positive_is_ignored(self):
        self.assertTrue(_is_minor_decorative_marker_bounds_difference(
            "→", 26.875, 23.185, 25.781, 29.531, 1.0424, 0.7851
        ))

    def test_actual_down_arrow_false_positive_is_ignored(self):
        self.assertTrue(_is_minor_decorative_marker_bounds_difference(
            "↓", 9.375, 13.25, 18.75, 12.656, 0.5, 1.0469
        ))

    def test_actual_bullet_false_positive_is_ignored(self):
        self.assertTrue(_is_minor_decorative_marker_bounds_difference(
            "•", 9.625, 18.6, 9.375, 16.406, 1.0267, 1.1337
        ))

    def test_single_letter_is_not_ignored(self):
        self.assertFalse(_is_minor_decorative_marker_bounds_difference(
            "A", 10.0, 18.6, 9.375, 16.406, 1.0667, 1.1337
        ))

    def test_multiple_characters_are_not_ignored(self):
        self.assertFalse(_is_minor_decorative_marker_bounds_difference(
            "普通文本", 108.0, 16.0, 100.0, 16.4, 1.08, 0.9756
        ))

    def test_large_ratio_marker_overflow_is_not_ignored(self):
        self.assertFalse(_is_minor_decorative_marker_bounds_difference(
            "•", 9.625, 22.0, 9.375, 16.406, 1.0267, 1.3410
        ))

    def test_large_absolute_marker_overflow_is_not_ignored(self):
        self.assertFalse(_is_minor_decorative_marker_bounds_difference(
            "→", 29.0, 23.0, 25.781, 29.531, 1.1249, 0.7788
        ))

    def test_marker_in_large_text_box_is_not_ignored(self):
        self.assertFalse(_is_minor_decorative_marker_bounds_difference(
            "★", 41.5, 20.0, 40.2, 22.0, 1.0323, 0.9091
        ))


if __name__ == "__main__":
    unittest.main()

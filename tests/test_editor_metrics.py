import unittest

from core.editor import Editor


class EditorMetricTests(unittest.TestCase):
    def test_metric_names_include_default_first(self):
        self.assertEqual(Editor.get_metric_names()[0], "Libre")

    def test_free_metric_has_no_targets(self):
        self.assertEqual(
            Editor.get_line_metric_targets("Libre", [8, 10, 12]),
            [None, None, None],
        )

    def test_fixed_metric_targets_each_non_empty_line(self):
        self.assertEqual(
            Editor.get_line_metric_targets("Alexandrin 12", [12, 0, 10]),
            [12, None, 12],
        )

    def test_haiku_targets_ignore_empty_lines(self):
        self.assertEqual(
            Editor.get_line_metric_targets("Haiku 5/7/5", [5, 0, 7, 5]),
            [5, None, 7, 5],
        )

    def test_haiku_extra_non_empty_line_is_flagged(self):
        self.assertEqual(
            Editor.get_line_metric_targets("Haiku 5/7/5", [5, 7, 5, 4]),
            [5, 7, 5, 0],
        )

    def test_metric_progress_counts_matching_checked_lines(self):
        self.assertEqual(
            Editor.summarize_metric_progress("Haiku 5/7/5", [5, 6, 5, 4]),
            (2, 4),
        )


if __name__ == "__main__":
    unittest.main()

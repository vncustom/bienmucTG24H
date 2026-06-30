import unittest

from app_bien_muc_tg24h import (
    find_crew_rtf_file,
    is_main_news_name,
    is_valid_video_id,
    should_include_news_row,
)


class ListRuleTests(unittest.TestCase):
    def test_video_id_must_start_with_digit(self):
        self.assertTrue(is_valid_video_id("260611056"))
        self.assertTrue(is_valid_video_id("260611056a"))
        self.assertFalse(is_valid_video_id(""))
        self.assertFalse(is_valid_video_id("qc123"))

    def test_main_news_name_prefixes(self):
        self.assertTrue(is_main_news_name("24H-Iran/Israel"))
        self.assertTrue(is_main_news_name("24h-Iran/Israel"))
        self.assertTrue(is_main_news_name("24 Ukraine"))
        self.assertTrue(is_main_news_name("GAT24H-Israel airstrike"))
        self.assertTrue(is_main_news_name("GAT24h-Israel airstrike"))
        self.assertTrue(is_main_news_name("GAT abc"))
        self.assertFalse(is_main_news_name("QC123"))

    def test_news_row_ignores_column_d_status(self):
        for status in ("ONLINE", "PLAYED", "OFFLINE", "", None):
            with self.subTest(status=status):
                self.assertTrue(
                    should_include_news_row(
                        "GAT 03 2906 - AFP - WHO 1300 nguoi tu vong do nang nong",
                        "260629183",
                        status,
                    )
                )

    def test_news_row_still_requires_column_a_and_c_rules(self):
        self.assertFalse(should_include_news_row("QC123", "260629183", "ONLINE"))
        self.assertFalse(should_include_news_row("GAT abc", "qc123", "ONLINE"))


class CrewRtfRuleTests(unittest.TestCase):
    def test_finds_any_rtf_containing_crew_name(self):
        files = [
            "24H-news.rtf",
            "NHUNG NGUOI THUC HIEN abc.rtf",
            "NHUNG NGUOI THUC HIEN.txt",
        ]

        self.assertEqual(
            find_crew_rtf_file(files),
            "NHUNG NGUOI THUC HIEN abc.rtf",
        )

    def test_prefers_exact_crew_rtf_name_when_available(self):
        files = [
            "NHUNG NGUOI THUC HIEN abc.rtf",
            "NHUNG NGUOI THUC HIEN.rtf",
        ]

        self.assertEqual(
            find_crew_rtf_file(files),
            "NHUNG NGUOI THUC HIEN.rtf",
        )


if __name__ == "__main__":
    unittest.main()

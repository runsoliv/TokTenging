import tempfile
import unittest
from datetime import date
from pathlib import Path

from tok.email_downloader import (
    imap_date_bounds,
    new_output_directory,
    parse_date_range,
    safe_filename,
    unused_path,
)


class EmailDownloaderTests(unittest.TestCase):
    def test_date_bounds_include_both_user_dates(self):
        start, end = parse_date_range("2026-07-01", "2026-08-31")
        self.assertEqual((start, end), (date(2026, 7, 1), date(2026, 8, 31)))
        self.assertEqual(imap_date_bounds(start, end), ("01-Jul-2026", "01-Sep-2026"))

    def test_same_day_range_includes_exact_day(self):
        start, end = parse_date_range("2026-07-01", "2026-07-01")
        self.assertEqual(imap_date_bounds(start, end), ("01-Jul-2026", "02-Jul-2026"))

    def test_reversed_and_invalid_ranges_are_rejected(self):
        with self.assertRaises(ValueError):
            parse_date_range("2026-08-31", "2026-07-01")
        with self.assertRaises(ValueError):
            parse_date_range("01/07/2026", "2026-08-31")

    def test_output_folders_are_unique_and_date_labelled(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            first = new_output_directory(temp_dir, date(2026, 7, 1), date(2026, 8, 31))
            second = new_output_directory(temp_dir, date(2026, 7, 1), date(2026, 8, 31))
            self.assertEqual(first.name, "download_2026-07-01_to_2026-08-31")
            self.assertEqual(second.name, "download_2026-07-01_to_2026-08-31_2")

    def test_windows_unsafe_and_long_names_are_bounded(self):
        cleaned = safe_filename("folder/" + "x" * 300 + ":invoice?.pdf")
        self.assertLessEqual(len(cleaned), 140)
        self.assertEqual(Path(cleaned).suffix, ".pdf")
        self.assertNotRegex(cleaned, r'[<>:"/\\|?*]')

    def test_duplicate_suffix_does_not_grow_repeatedly(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            (directory / "invoice.pdf").write_bytes(b"one")
            (directory / "invoice_2.pdf").write_bytes(b"two")
            self.assertEqual(unused_path(directory, "invoice.pdf").name, "invoice_3.pdf")


if __name__ == "__main__":
    unittest.main()

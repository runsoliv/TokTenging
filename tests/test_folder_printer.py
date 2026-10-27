import tempfile
import unittest
from pathlib import Path

from tok.folder_printer import PdfPrintUnavailable, printable_files, print_folder_files


class FolderPrinterTests(unittest.TestCase):
    def test_lists_only_top_level_files(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "one.pdf").write_bytes(b"one")
            (root / "two.txt").write_text("two", encoding="utf-8")
            nested = root / "nested"
            nested.mkdir()
            (nested / "three.pdf").write_bytes(b"three")

            self.assertEqual(
                {path.name for path in printable_files(root)},
                {"one.pdf", "two.txt"},
            )

    def test_prints_every_file_and_continues_after_failure(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "one.pdf").write_bytes(b"one")
            (root / "two.txt").write_text("two", encoding="utf-8")
            attempted = []
            output = []

            def fake_print(path):
                attempted.append(Path(path).name)
                if Path(path).suffix == ".txt":
                    raise OSError("no print association")

            result = print_folder_files(
                root,
                emit=output.append,
                print_file=fake_print,
                delay_seconds=0,
            )

            self.assertEqual(set(attempted), {"one.pdf", "two.txt"})
            self.assertEqual(result["total"], 2)
            self.assertEqual(result["printed"], 1)
            self.assertEqual(len(result["failed"]), 1)
            self.assertTrue(any("Could not print" in line for line in output))

    def test_missing_and_empty_folders_are_rejected(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            with self.assertRaises(ValueError):
                printable_files(Path(temp_dir) / "missing")
            with self.assertRaises(ValueError):
                print_folder_files(temp_dir, print_file=lambda _path: None, delay_seconds=0)

    def test_cancel_stops_before_submitting_the_next_file(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            for number in range(3):
                (root / f"{number}.pdf").write_bytes(b"pdf")
            attempted = []

            result = print_folder_files(
                root,
                print_file=lambda path: attempted.append(path),
                is_cancelled=lambda: len(attempted) >= 1,
                delay_seconds=0,
            )

            self.assertEqual(len(attempted), 1)
            self.assertTrue(result["cancelled"])
            self.assertEqual(result["printed"], 1)

    def test_missing_pdf_backend_stops_the_batch_immediately(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "one.pdf").write_bytes(b"pdf")
            (root / "two.pdf").write_bytes(b"pdf")
            attempted = []

            def unavailable(path):
                attempted.append(path)
                raise PdfPrintUnavailable("PDF printer missing")

            with self.assertRaises(PdfPrintUnavailable):
                print_folder_files(root, print_file=unavailable, delay_seconds=0)
            self.assertEqual(len(attempted), 1)


if __name__ == "__main__":
    unittest.main()

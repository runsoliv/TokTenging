import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd

from tok.auto_coder import TransactionAutoCoder, get_auto_coder, merchant_key


class AutoCoderCleanRebuildTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.training_dir = self.root / "trainingCoded"
        self.training_data_dir = self.root / "trainingData"
        self.key_dir = self.root / "keys"
        self.memory_path = self.training_dir / "auto_code_memory.json"
        self.training_dir.mkdir()

    def tearDown(self):
        self.temp_dir.cleanup()

    def _write_training_file(self, rows):
        path = self.training_dir / "training.xlsx"
        pd.DataFrame(rows).to_excel(path, index=False)
        return path

    def _write_archive_session(self, coded_rows, paired_rows):
        session = self.training_data_dir / "imports" / "test_session"
        session.mkdir(parents=True)
        coded_path = session / "test_coded.xlsx"
        paired_path = session / "paired_rows.csv"
        pd.DataFrame(coded_rows).to_excel(coded_path, index=False)
        pd.DataFrame(paired_rows).to_csv(paired_path, index=False)
        manifest = {
            "archive_files": {"coded": coded_path.name},
            "output_coded_filename": coded_path.name,
        }
        (session / "manifest.json").write_text(
            json.dumps(manifest),
            encoding="utf-8",
        )
        return coded_path, paired_path

    def _coder(self, force_rebuild=False):
        return TransactionAutoCoder(
            training_dir=self.training_dir,
            training_data_dir=self.training_data_dir,
            key_dir=self.key_dir,
            memory_path=self.memory_path,
            force_rebuild=force_rebuild,
        )

    @staticmethod
    def _row(text, debit):
        return {
            "DATE": "01.01.2026",
            "TEXT": text,
            "DEBIT": debit,
            "AMOUNT": 1000,
            "ID": "",
            "CREDIT": 7810,
            "Positive/Negative": "-",
        }

    def test_changed_file_replaces_old_counts_instead_of_merging(self):
        self._write_training_file([self._row("Clean Rebuild Merchant", 4100)])
        first = self._coder()
        key = merchant_key("Clean Rebuild Merchant")
        self.assertEqual(first.global_exact[key], {"4100": 1})

        self._write_training_file([self._row("Clean Rebuild Merchant", 6200)])
        rebuilt = self._coder(force_rebuild=True)

        self.assertEqual(rebuilt.global_exact[key], {"6200": 1})
        self.assertNotIn("4100", rebuilt.global_exact[key])
        self.assertEqual(rebuilt.training_rows, 1)

    def test_deleted_sources_clear_model_and_json(self):
        source_path = self._write_training_file(
            [self._row("Merchant To Delete", 4200)]
        )
        self._coder()
        source_path.unlink()

        rebuilt = self._coder(force_rebuild=True)
        payload = json.loads(self.memory_path.read_text(encoding="utf-8"))

        self.assertFalse(rebuilt.ready())
        self.assertEqual(rebuilt.training_rows, 0)
        self.assertEqual(dict(rebuilt.global_exact), {})
        self.assertEqual(payload["training_rows"], 0)
        self.assertEqual(payload["global_exact"], {})
        self.assertEqual(payload["source_fingerprint"], {})

    def test_normal_auto_coding_uses_memory_until_explicit_rebuild(self):
        self._write_training_file([self._row("Button Rebuild Merchant", 4050)])
        first = get_auto_coder(
            force_reload=True,
            force_rebuild=True,
            training_dir=self.training_dir,
            training_data_dir=self.training_data_dir,
            key_dir=self.key_dir,
            memory_path=self.memory_path,
        )
        key = merchant_key("Button Rebuild Merchant")
        self._write_training_file(
            [self._row("Button Rebuild Merchant", 6200)]
        )

        cached = get_auto_coder(
            training_dir=self.training_dir,
            training_data_dir=self.training_data_dir,
            key_dir=self.key_dir,
            memory_path=self.memory_path,
        )
        reloaded = get_auto_coder(
            force_reload=True,
            training_dir=self.training_dir,
            training_data_dir=self.training_data_dir,
            key_dir=self.key_dir,
            memory_path=self.memory_path,
        )
        rebuilt = get_auto_coder(
            force_reload=True,
            force_rebuild=True,
            training_dir=self.training_dir,
            training_data_dir=self.training_data_dir,
            key_dir=self.key_dir,
            memory_path=self.memory_path,
        )

        self.assertIs(first, cached)
        self.assertIsNot(first, reloaded)
        self.assertTrue(reloaded.loaded_from_memory)
        self.assertEqual(reloaded.global_exact[key], {"4050": 1})
        self.assertEqual(rebuilt.global_exact[key], {"6200": 1})

    def test_archive_rebuild_uses_existing_coded_excel_not_paired_csv(self):
        coded_path, paired_path = self._write_archive_session(
            [self._row("Current Excel Merchant", 4100)],
            [self._row("Stale CSV Merchant", 9999)],
        )

        rebuilt = self._coder(force_rebuild=True)

        self.assertEqual(
            rebuilt.global_exact[merchant_key("Current Excel Merchant")],
            {"4100": 1},
        )
        self.assertNotIn(merchant_key("Stale CSV Merchant"), rebuilt.global_exact)
        self.assertTrue(paired_path.exists())

        coded_path.unlink()
        after_delete = self._coder(force_rebuild=True)
        payload = json.loads(self.memory_path.read_text(encoding="utf-8"))

        self.assertFalse(after_delete.ready())
        self.assertEqual(dict(after_delete.global_exact), {})
        self.assertEqual(payload["training_rows"], 0)

    def test_memory_json_is_written_compactly(self):
        self._write_training_file([self._row("Compact JSON Merchant", 4050)])

        self._coder(force_rebuild=True)
        memory_text = self.memory_path.read_text(encoding="utf-8")

        self.assertNotIn("\n", memory_text)
        self.assertLess(len(memory_text), 100000)

    def test_failed_memory_write_keeps_last_valid_file(self):
        self._write_training_file([self._row("Atomic Memory Merchant", 4050)])
        coder = self._coder(force_rebuild=True)
        before = self.memory_path.read_bytes()

        with patch("tok.auto_coder.json.dump", side_effect=OSError("disk full")):
            saved = coder.save_memory()

        self.assertFalse(saved)
        self.assertEqual(self.memory_path.read_bytes(), before)
        self.assertEqual(list(self.memory_path.parent.glob(".auto_code_memory.json.*.tmp")), [])


if __name__ == "__main__":
    unittest.main()

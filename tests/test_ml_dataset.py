import hashlib
import json
from pathlib import Path
import shutil
import tempfile
import unittest

import pandas as pd

from tok.ml_dataset import (
    TrainingDataLoader,
    VEHICLE_EXPENSE_FAMILY,
    build_runtime_feature_record,
    normalize_kennitala,
)


class TrainingDatasetTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.training_data = self.root / "trainingData"
        self.imports = self.training_data / "imports"
        self.imports.mkdir(parents=True)
        self.legacy = self.root / "trainingCoded"
        self.legacy.mkdir()

    def tearDown(self):
        self.temp_dir.cleanup()

    @staticmethod
    def _row(
        text="Merchant",
        debit="4061",
        credit="7810",
        sign="-",
        amount=1000,
        date="01.01.2026",
        id_value="",
    ):
        return {
            "DATE": date,
            "TEXT": text,
            "DEBIT": debit,
            "ID": id_value,
            "AMOUNT": amount,
            "CREDIT": credit,
            "Positive/Negative": sign,
        }

    def _write_session(
        self,
        name,
        coded_rows,
        *,
        coded_name="verified_coded.xlsx",
        paired_rows=None,
        raw_bytes=None,
        raw_column_map=None,
        sync_mode="direct",
        compression_map=None,
        bank_type="landsbanki",
        client="client-a",
    ):
        session = self.imports / name
        session.mkdir()
        coded_path = session / coded_name
        pd.DataFrame(coded_rows).to_excel(coded_path, index=False)
        archive_files = {"coded": coded_name}
        if paired_rows is not None:
            paired_path = session / "paired_rows.csv"
            pd.DataFrame(paired_rows).to_csv(paired_path, index=False, encoding="utf-8-sig")
            archive_files["paired_rows"] = paired_path.name
        if raw_bytes is not None:
            raw_path = session / "raw.xlsx"
            raw_path.write_bytes(raw_bytes)
            archive_files["raw"] = raw_path.name
        if compression_map is not None:
            map_path = session / "compression_map.json"
            map_path.write_text(json.dumps(compression_map), encoding="utf-8")
            archive_files["compression_map"] = map_path.name
        manifest = {
            "schema_version": "1.0",
            "archive_files": archive_files,
            "output_coded_filename": coded_name,
            "detected_bank_type": bank_type,
            "client_hint": client,
            "columns": {"raw_column_map": raw_column_map or {}},
        }
        if sync_mode is not None:
            manifest["paired_rows_sync_mode"] = sync_mode
        (session / "manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False),
            encoding="utf-8",
        )
        return session, coded_path

    def _load(self, include_legacy=False):
        return TrainingDataLoader(
            training_data_dir=self.training_data,
            legacy_training_dir=self.legacy,
            include_legacy=include_legacy,
        ).load()

    @staticmethod
    def _hash(path):
        return hashlib.sha256(Path(path).read_bytes()).hexdigest()

    def test_manifest_designates_only_the_verified_coded_workbook(self):
        session, _ = self._write_session(
            "session-a",
            [self._row()],
            coded_name="chosen_coded_compressed_txs.xlsx",
        )
        pd.DataFrame([self._row(text="Unrelated", debit="9999")]).to_excel(
            session / "unrelated.xlsx",
            index=False,
        )
        unmanifested = self.imports / "no-manifest"
        unmanifested.mkdir()
        pd.DataFrame([self._row(text="No manifest")]).to_excel(
            unmanifested / "orphan.xlsx",
            index=False,
        )

        result = self._load()

        self.assertEqual(len(result.frame), 1)
        self.assertEqual(result.frame.iloc[0]["text"], "Merchant")
        self.assertEqual(result.report["counts"]["import_sessions_discovered"], 2)
        self.assertEqual(result.report["counts"]["import_workbooks_selected"], 1)
        self.assertIn("missing_manifest", {issue.reason for issue in result.issues})

    def test_direction_selects_the_non_bank_side_and_excludes_transfers(self):
        self._write_session(
            "directions",
            [
                self._row(text="Expense", debit="4061", credit="7810", sign="-"),
                self._row(text="Income", debit="7810", credit="3000", sign="+"),
                self._row(text="Transfer", debit="7820", credit="7810", sign="-"),
                self._row(text="Vehicle A", debit="4635", credit="7810", sign="-"),
                self._row(text="Vehicle B", debit="4640", credit="7810", sign="-"),
            ],
        )

        result = self._load()

        self.assertEqual(result.frame["target_code"].tolist(), ["4061", "3000", "4635", "4640"])
        self.assertEqual(result.frame["direction"].tolist(), ["outgoing", "incoming", "outgoing", "outgoing"])
        self.assertEqual(
            result.frame.loc[result.frame["target_code"].isin(["4635", "4640"]), "model_target"].tolist(),
            [VEHICLE_EXPENSE_FAMILY, VEHICLE_EXPENSE_FAMILY],
        )
        self.assertEqual(result.report["counts"]["rows_skipped_control_account_target"], 1)

    def test_fixed_or_skipped_formats_are_reported_but_excluded_from_model_frame(self):
        self._write_session(
            "ordinary",
            [self._row(text="Ordinary")],
            bank_type="landsbanki",
        )
        self._write_session(
            "special",
            [self._row(text="Special")],
            bank_type="islandsbanki_special",
        )

        result = self._load()

        self.assertEqual(len(result.frame), 2)
        self.assertEqual(result.model_frame()["text"].tolist(), ["Ordinary"])
        self.assertEqual(result.report["counts"]["production_excluded_rows"], 1)
        self.assertEqual(result.report["counts"]["production_model_rows"], 1)

    def test_coded_header_aliases_are_normalized(self):
        self._write_session(
            "aliases",
            [{
                "Dagsetning": "01.01.2026",
                "Texti": "Alias Merchant",
                "Debet": "4061",
                "Kennitala": "123456-7890",
                "Upphæð": 500,
                "Kredit": "7810",
                "Stefna": "-",
            }],
        )

        result = self._load()

        self.assertEqual(len(result.frame), 1)
        self.assertEqual(result.frame.iloc[0]["target_code"], "4061")
        self.assertEqual(result.frame.iloc[0]["valid_kennitala"], "1234567890")

    def test_validated_raw_pair_extracts_aliases_and_strict_kennitala(self):
        coded = self._row(text="Generic", id_value="not-an-id")
        paired = {
            **coded,
            "raw_index": 7,
            "coded_index": 0,
            "raw_skyring": "Nánari skýring frá banka",
            "raw_tilvisun": "123456-7890",
        }
        self._write_session(
            "paired",
            [coded],
            paired_rows=[paired],
            raw_bytes=b"read-only raw source",
            raw_column_map={
                "raw_skyring": "Skýring",
                "raw_tilvisun": "Tilvísun",
            },
        )

        result = self._load()
        row = result.frame.iloc[0]

        self.assertEqual(row["raw_pair_status"], "validated_pair")
        self.assertIn("Nánari skýring", row["raw_explanation"])
        self.assertEqual(row["valid_kennitala"], "1234567890")
        self.assertTrue(row["transaction_reference"].startswith("tilvisun:"))

    def test_cardholder_kennitala_is_excluded_and_merchant_is_preferred(self):
        coded = self._row(text="Card purchase", id_value="")
        paired = {
            **coded,
            "raw_index": 0,
            "raw_cardholder_id": "111111-1111",
            "raw_merchant_id": "222222-2222",
        }
        self._write_session(
            "landsbankinn-card-identifiers",
            [coded],
            paired_rows=[paired],
            raw_bytes=b"card raw",
            raw_column_map={
                "raw_cardholder_id": "Kennitala korthafa",
                "raw_merchant_id": "Kennitala sÃ¶luaÃ°ila",
            },
            bank_type="landsbankinn_kort",
        )

        row = self._load().frame.iloc[0]

        self.assertEqual(row["valid_kennitala"], "2222222222")
        self.assertNotIn("1111111111", row["raw_explanation"])

    def test_reference_identifiers_do_not_enter_model_explanation_text(self):
        coded = self._row(text="Generic", id_value="")
        paired = {
            **coded,
            "raw_index": 0,
            "raw_reference": "CLAIM-ABC-12345",
        }
        self._write_session(
            "reference-only",
            [coded],
            paired_rows=[paired],
            raw_bytes=b"raw",
            raw_column_map={"raw_reference": "Tilvisun"},
        )

        row = self._load().frame.iloc[0]

        self.assertEqual(row["raw_explanation"], "")
        self.assertTrue(row["transaction_reference"].startswith("tilvisun:"))

    def test_pair_identity_mismatch_omits_all_raw_features(self):
        coded = self._row(date="01.01.2026")
        paired = {
            **coded,
            "DATE": "02.01.2026",
            "raw_index": 0,
            "raw_skyring": "Must not be paired",
        }
        self._write_session(
            "mismatch",
            [coded],
            paired_rows=[paired],
            raw_bytes=b"raw",
            raw_column_map={"raw_skyring": "Skýring"},
        )

        result = self._load()
        row = result.frame.iloc[0]

        self.assertEqual(row["raw_pair_status"], "identity_unmatched")
        self.assertEqual(row["raw_explanation"], "")
        self.assertEqual(row["valid_kennitala"], "")

    def test_sync_error_and_coded_only_rebuild_are_missing_safe(self):
        session, _ = self._write_session(
            "sync-error",
            [self._row()],
            paired_rows=[{**self._row(), "raw_skyring": "Unsafe"}],
            raw_column_map={"raw_skyring": "Skýring"},
        )
        (session / "paired_rows_sync_error.txt").write_text("unsafe", encoding="utf-8")
        self._write_session(
            "coded-only",
            [self._row(text="Coded only")],
            paired_rows=[self._row(text="Coded only")],
            sync_mode="coded_only_rebuild",
        )

        result = self._load()

        self.assertEqual(
            set(result.frame["raw_pair_status"]),
            {"unsafe_sync_error", "coded_only_rebuild"},
        )
        self.assertTrue((result.frame["raw_explanation"] == "").all())

    def test_missing_raw_columns_do_not_reject_a_coded_row(self):
        self._write_session("no-pair", [self._row()], paired_rows=None, raw_bytes=None)

        result = self._load()

        self.assertEqual(len(result.frame), 1)
        self.assertEqual(result.frame.iloc[0]["raw_pair_status"], "paired_rows_missing")
        self.assertEqual(result.frame.iloc[0]["raw_explanation"], "")

    def test_multirow_compression_pair_is_intentionally_ambiguous(self):
        coded = self._row(amount=2000)
        paired_rows = [
            {**self._row(amount=1000), "raw_index": 0, "raw_skyring": "First"},
            {**self._row(amount=1000), "raw_index": 1, "raw_skyring": "Second"},
        ]
        compression_map = {
            "rows": [{
                "coded_index": 0,
                "paired_indices": [0, 1],
                "identity": {
                    "DATE": coded["DATE"],
                    "AMOUNT": coded["AMOUNT"],
                    "Positive/Negative": coded["Positive/Negative"],
                },
            }],
        }
        self._write_session(
            "compressed",
            [coded],
            paired_rows=paired_rows,
            raw_bytes=b"raw",
            raw_column_map={"raw_skyring": "Skýring"},
            sync_mode="compressed_map",
            compression_map=compression_map,
        )

        result = self._load()

        self.assertEqual(result.frame.iloc[0]["raw_pair_status"], "ambiguous_compressed_group")
        self.assertEqual(result.frame.iloc[0]["raw_explanation"], "")

    def test_swapped_singleton_compression_map_does_not_attach_raw_features(self):
        coded_rows = [self._row(text="Merchant A"), self._row(text="Merchant B")]
        paired_rows = [
            {**coded_rows[1], "raw_index": 0, "raw_text": "Raw B"},
            {**coded_rows[0], "raw_index": 1, "raw_text": "Raw A"},
        ]
        compression_map = {"rows": [
            {
                "coded_index": 0,
                "paired_indices": [0],
                "identity": {"DATE": coded_rows[0]["DATE"]},
            },
            {
                "coded_index": 1,
                "paired_indices": [1],
                "identity": {"DATE": coded_rows[1]["DATE"]},
            },
        ]}
        self._write_session(
            "swapped-compression",
            coded_rows,
            paired_rows=paired_rows,
            raw_bytes=b"raw",
            raw_column_map={"raw_text": "Texti"},
            compression_map=compression_map,
        )

        result = self._load()

        self.assertEqual(set(result.frame["raw_pair_status"]), {"compression_identity_mismatch"})
        self.assertTrue((result.frame["raw_explanation"] == "").all())

    def test_repeated_direct_identity_is_not_paired_by_position(self):
        coded = self._row(text="Repeated")
        self._write_session(
            "repeated-direct",
            [coded, coded],
            paired_rows=[
                {**coded, "raw_index": 0, "raw_text": "First raw detail"},
                {**coded, "raw_index": 1, "raw_text": "Second raw detail"},
            ],
            raw_bytes=b"raw",
            raw_column_map={"raw_text": "Texti"},
        )

        result = self._load()

        self.assertEqual(set(result.frame["raw_pair_status"]), {"identity_ambiguous"})
        self.assertTrue((result.frame["raw_explanation"] == "").all())

    def test_exact_duplicate_workbook_is_selected_once(self):
        first_session, first_path = self._write_session("first", [self._row()])
        second_session = self.imports / "second"
        second_session.mkdir()
        second_path = second_session / "copy_coded.xlsx"
        shutil.copy2(first_path, second_path)
        manifest = {
            "archive_files": {"coded": second_path.name},
            "output_coded_filename": second_path.name,
            "detected_bank_type": "landsbanki",
            "client_hint": "client-a",
        }
        (second_session / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")

        result = self._load()

        self.assertEqual(len(result.frame), 1)
        self.assertEqual(result.report["counts"]["exact_duplicate_workbooks"], 1)
        self.assertEqual(result.report["counts"]["import_workbooks_selected"], 1)
        self.assertTrue(first_session.exists())

    def test_same_raw_transaction_is_deduplicated_across_rebuilt_sessions(self):
        coded = self._row()
        paired = {**coded, "raw_index": 0, "raw_skyring": "Same transaction"}
        self._write_session(
            "older",
            [coded],
            paired_rows=[paired],
            raw_bytes=b"identical raw workbook",
            raw_column_map={"raw_skyring": "Skýring"},
        )
        # An extra harmless column makes the coded workbook bytes different;
        # the raw hash + validated raw index still identifies the transaction.
        newer = dict(coded)
        newer["Cnf"] = "MN"
        newer_paired = dict(paired)
        newer_paired["Cnf"] = "MN"
        self._write_session(
            "newer",
            [newer],
            paired_rows=[newer_paired],
            raw_bytes=b"identical raw workbook",
            raw_column_map={"raw_skyring": "Skýring"},
        )

        result = self._load()

        self.assertEqual(len(result.frame), 1)
        self.assertEqual(result.report["counts"]["duplicate_rows_removed"], 1)

    def test_conflicting_duplicate_transaction_is_removed_not_majority_voted(self):
        first = self._row(debit="4061")
        second = self._row(debit="6220")
        self._write_session(
            "conflict-a",
            [first],
            paired_rows=[{**first, "raw_index": 0}],
            raw_bytes=b"same raw",
        )
        self._write_session(
            "conflict-b",
            [second],
            paired_rows=[{**second, "raw_index": 0}],
            raw_bytes=b"same raw",
        )

        result = self._load()

        self.assertTrue(result.frame.empty)
        self.assertEqual(result.report["counts"]["duplicate_conflict_rows_removed"], 2)
        self.assertIn("duplicate_label_conflict", {issue.reason for issue in result.issues})

    def test_different_stable_references_preserve_legitimate_repeat_payments(self):
        coded = self._row(text="Same merchant", id_value="1234567890")
        for name, reference in (("repeat-a", "ARN-A"), ("repeat-b", "ARN-B")):
            session_coded = {**coded, "Cnf": name}
            self._write_session(
                name,
                [session_coded],
                paired_rows=[{**session_coded, "raw_index": 0, "raw_arn": reference}],
                raw_bytes=name.encode("ascii"),
                raw_column_map={"raw_arn": "ARN"},
            )

        result = self._load()

        self.assertEqual(len(result.frame), 2)
        self.assertEqual(result.report["counts"].get("duplicate_rows_removed", 0), 0)

    def test_legacy_discovery_excludes_generated_variants(self):
        pd.DataFrame([self._row(text="Legacy")]).to_excel(self.legacy / "verified.xlsx", index=False)
        pd.DataFrame([self._row(text="Combined")]).to_excel(self.legacy / "old_combined.xlsx", index=False)
        key_dir = self.legacy / "coded keys"
        key_dir.mkdir()
        pd.DataFrame([self._row(text="Key file")]).to_excel(key_dir / "keys.xlsx", index=False)

        result = self._load(include_legacy=True)

        self.assertEqual(len(result.frame), 1)
        self.assertEqual(result.frame.iloc[0]["text"], "Legacy")
        self.assertEqual(result.report["counts"]["legacy_workbooks_selected"], 1)

    def test_loading_is_reproducible_and_never_changes_sources(self):
        session, coded_path = self._write_session("stable", [self._row()])
        manifest_path = session / "manifest.json"
        before = {path: self._hash(path) for path in (coded_path, manifest_path)}

        first = self._load()
        second = self._load()
        after = {path: self._hash(path) for path in before}

        self.assertEqual(first.fingerprint, second.fingerprint)
        self.assertEqual(first.frame.to_dict("records"), second.frame.to_dict("records"))
        self.assertEqual(before, after)

    def test_fingerprint_changes_when_paired_raw_features_change(self):
        coded = self._row()
        session, _ = self._write_session(
            "feature-fingerprint",
            [coded],
            paired_rows=[{**coded, "raw_index": 0, "raw_skyring": "First explanation"}],
            raw_bytes=b"same raw bytes",
            raw_column_map={"raw_skyring": "Texti"},
        )
        first = self._load()

        paired_path = session / "paired_rows.csv"
        changed = pd.read_csv(paired_path, dtype=object, keep_default_na=False)
        changed.loc[0, "raw_skyring"] = "Corrected explanation"
        changed.to_csv(paired_path, index=False, encoding="utf-8-sig")
        second = self._load()

        self.assertNotEqual(first.fingerprint, second.fingerprint)
        self.assertNotEqual(
            first.frame.iloc[0]["raw_explanation"],
            second.frame.iloc[0]["raw_explanation"],
        )


class KennitalaValidationTests(unittest.TestCase):
    def test_only_exactly_ten_digits_with_optional_hyphen_are_valid(self):
        self.assertEqual(normalize_kennitala("123456-7890"), "1234567890")
        self.assertEqual(normalize_kennitala("1234567890"), "1234567890")
        self.assertEqual(normalize_kennitala(1234567890.0), "1234567890")
        for invalid in ("123456789", "012345678", "12345678901", "12345-67890", "A234567890", ""):
            with self.subTest(invalid=invalid):
                self.assertEqual(normalize_kennitala(invalid), "")

    def test_runtime_feature_record_matches_missing_safe_raw_rules(self):
        coded = pd.Series({
            "DATE": "31.08.2026",
            "TEXT": "Verslun **-8748",
            "ID": "",
            "AMOUNT": 1250,
            "Positive/Negative": "-",
        })
        raw = pd.Series({
            "Skýring": "Nánari skýring",
            "Kennitala mótaðila": "123456-7890",
            "Kennitala korthafa": "999999-9999",
            "Cnf": "AD",
            "STATUS": "coded",
        })

        record = build_runtime_feature_record(
            coded,
            source_row=raw,
            bank_type="landsbankinn_kort",
        )
        coded_only = build_runtime_feature_record(coded, bank_type="landsbankinn_kort")

        self.assertEqual(record["direction"], "outgoing")
        self.assertEqual(record["valid_kennitala"], "1234567890")
        self.assertEqual(record["raw_explanation"], "Nánari skýring")
        self.assertNotIn("AD", record["raw_explanation"])
        self.assertEqual(coded_only["raw_explanation"], "")
        self.assertEqual(coded_only["raw_pair_status"], "runtime_coded_only")


if __name__ == "__main__":
    unittest.main()

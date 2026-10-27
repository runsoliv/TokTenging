import os
from pathlib import Path
import tempfile
import unittest

import pandas as pd

from tok.bank_formatter import (
    _append_reference_rows,
    _compressed_reference_path,
    _compression_file_signature,
    _compression_identity_date,
    _compression_store_identity,
    _drag_positions_between,
    _find_compressible_transaction_clusters,
    _read_transaction_compression_workbook,
    _remember_compressed_reference_path,
    _save_transaction_compressions,
    _source_groups_for_current_output,
    _top_selected_transaction_text,
)
from tok.coded_file_tools import (
    CodedWorkbookError,
    combined_output_path,
    combine_coded_workbooks,
    find_recent_coded_files,
)
from tok.ui_components import _parse_drop_paths


COLUMNS = [
    "DATE",
    "TEXT",
    "DEBIT",
    "Cnf",
    "ID",
    "AMOUNT",
    "CREDIT",
    "Positive/Negative",
    "STATUS",
]

REFERENCE_PRINT_COLUMNS = ["DATE", "TEXT", "DEBIT", "ID", "AMOUNT"]


class ManualCompressionSelectionTests(unittest.TestCase):
    def test_drag_range_includes_rows_skipped_between_pointer_events(self):
        self.assertEqual(list(_drag_positions_between(None, 40)), [40])
        self.assertEqual(list(_drag_positions_between(40, 44)), [40, 41, 42, 43, 44])
        self.assertEqual(list(_drag_positions_between(44, 40)), [44, 43, 42, 41, 40])

    def test_compression_text_uses_topmost_selected_transaction(self):
        frame = pd.DataFrame(
            {"TEXT": ["First transaction", "Second transaction", "Third transaction"]},
            index=[10, 20, 30],
        )
        self.assertEqual(
            _top_selected_transaction_text(frame, [10, 20, 30], {2, 1}),
            "Second transaction",
        )
        self.assertEqual(
            _top_selected_transaction_text(frame, [10, 20, 30], set()),
            "",
        )

    def test_compression_map_dates_normalize_excel_and_iso_values(self):
        self.assertEqual(
            _compression_identity_date("2026-07-17T00:00:00"),
            "17.07.2026",
        )
        self.assertEqual(_compression_identity_date("17.07.2026"), "17.07.2026")


def coded_frame(rows):
    return pd.DataFrame(rows, columns=COLUMNS)


def tx(text, debit, amount, identifier=""):
    return ["17.07.2026", text, debit, "", identifier, amount, 7810, "+", ""]


def expense_tx(text, debit, amount, counter=7820, identifier=""):
    return ["17.07.2026", text, debit, "", identifier, amount, counter, "-", ""]


def income_tx(text, credit, amount, counter=1000, identifier=""):
    return ["17.07.2026", text, counter, "", identifier, amount, credit, "+", ""]


def bonus_expense_rows(counter=7820):
    return [
        expense_tx("BONUS KEFLAVIK, **-8748", 4061, 100, counter),
        expense_tx("BÓNUS NJARÐVÍK, **-8748", 9650, 200, counter),
        expense_tx("Bonus Helluhraun", 2107, 300, counter),
        expense_tx("BONUS MIDHRAUN - Sölunóta", 4061, 400, counter),
        expense_tx("BONUS KAUPTUN", 9650, 500, counter),
    ]


def cluster(indices, total, text):
    return {
        "row_indices": indices,
        "count": len(indices),
        "total": total,
        "summary_text": text,
    }


class AutomaticCompressionGroupingTests(unittest.TestCase):
    def test_groups_five_store_locations_and_flags_mixed_codes(self):
        clusters = _find_compressible_transaction_clusters(
            coded_frame(bonus_expense_rows()),
            minimum_count=5,
        )

        self.assertEqual(len(clusters), 1)
        self.assertEqual(clusters[0]["count"], 5)
        self.assertEqual(clusters[0]["summary_text"], "BONUS")
        self.assertEqual(clusters[0]["counter_account"], "7820")
        self.assertTrue(clusters[0]["mixed_codes"])

    def test_requires_five_rows_and_separates_sign_and_counter_account(self):
        self.assertEqual(
            _find_compressible_transaction_clusters(
                coded_frame(bonus_expense_rows()[:4]),
                minimum_count=5,
            ),
            [],
        )

        rows = bonus_expense_rows(counter=7820)
        rows.extend(bonus_expense_rows(counter=7810))
        rows.extend([
            income_tx("BONUS KEFLAVIK", 4000, 10),
            income_tx("BÓNUS NJARÐVÍK", 4000, 20),
            income_tx("BONUS HELLUHRAUN", 4000, 30),
            income_tx("BONUS MIDHRAUN", 4000, 40),
            income_tx("BONUS KAUPTUN", 4000, 50),
        ])
        clusters = _find_compressible_transaction_clusters(coded_frame(rows), minimum_count=5)

        self.assertEqual(len(clusters), 3)
        self.assertEqual(
            {(item["sign"], item["counter_account"]) for item in clusters},
            {("-", "7820"), ("-", "7810"), ("+", "1000")},
        )

    def test_store_aliases_strip_locations_but_processors_stay_separate(self):
        self.assertEqual(
            _compression_store_identity("Ískraft - Selfossi - Sölunóta")[0],
            "store:iskraft",
        )
        self.assertEqual(
            _compression_store_identity("STR*Parka, **-1234")[0],
            "store:parka",
        )
        self.assertNotEqual(
            _compression_store_identity("PAYPAL *NETFLIX.COM")[0],
            _compression_store_identity("PAYPAL *TEMU")[0],
        )
        self.assertNotEqual(
            _compression_store_identity("Bjarni Ágústsson")[0],
            _compression_store_identity("Bjarni Einarsson")[0],
        )
        self.assertNotEqual(
            _compression_store_identity("COSTCO ICELAND FUEL 401")[0],
            _compression_store_identity("COSTCO ICELAND WAREHOUSE")[0],
        )

    def test_machine_references_and_facebook_aliases_share_a_merchant(self):
        facebook_keys = {
            _compression_store_identity(value)[0]
            for value in (
                "FACEBK 2WLNWU57H2",
                "FACEBK QPPV3TM6H2, **-9248",
                "FACEBK REYKJAVIK",
                "FACEBOOK KEFLAVIK",
            )
        }
        spotify_keys = {
            _compression_store_identity(value)[0]
            for value in (
                "Spotify",
                "Spotify P42C61FD2F",
                "Spotify P43DC96916",
            )
        }

        self.assertEqual(facebook_keys, {"store:facebook"})
        self.assertEqual(len(spotify_keys), 1)

    def test_reference_grouping_does_not_use_unsafe_first_five_matching(self):
        self.assertNotEqual(
            _compression_store_identity("FACEBIRD A1B2C3D4")[0],
            _compression_store_identity("FACEBK A1B2C3D4")[0],
        )
        self.assertNotEqual(
            _compression_store_identity("Bjarni Agustsson A1B2C3D4")[0],
            _compression_store_identity("Bjarni Einarsson E5F6G7H8")[0],
        )
        self.assertNotEqual(
            _compression_store_identity("PAYPAL *NETFLIX A1B2C3D4")[0],
            _compression_store_identity("PAYPAL *TEMU E5F6G7H8")[0],
        )

    def test_facebk_references_form_one_mixed_code_cluster(self):
        rows = [
            expense_tx("FACEBK 2WLNWU57H2", 4055, 100),
            expense_tx("FACEBK 4GLAWSZ6H2", 4050, 200),
            expense_tx("FACEBK 5L2N4V56H2", 4055, 300),
            expense_tx("FACEBK ABBJVUR6H2", 4050, 400),
            expense_tx("FACEBK QPPV3TM6H2, **-9248", 4055, 500),
        ]

        clusters = _find_compressible_transaction_clusters(
            coded_frame(rows),
            minimum_count=5,
        )

        self.assertEqual(len(clusters), 1)
        self.assertEqual(clusters[0]["summary_text"], "FACEBOOK")
        self.assertEqual(clusters[0]["count"], 5)
        self.assertTrue(clusters[0]["mixed_codes"])

    def test_facebk_6220_vetoes_the_canonical_merchant(self):
        rows = [
            expense_tx(f"FACEBK REF{index}A1B2C3", 4055, index * 100)
            for index in range(1, 6)
        ]
        rows.append(expense_tx("FACEBOOK REYKJAVIK", 6220, 75))

        clusters = _find_compressible_transaction_clusters(
            coded_frame(rows),
            minimum_count=5,
        )

        self.assertEqual(clusters, [])

    def test_bare_6220_merchant_vetoes_its_referenced_variants(self):
        rows = [
            expense_tx(f"Spotify P4{index}C61FD2F", 4055, index * 100)
            for index in range(1, 6)
        ]
        rows.append(expense_tx("Spotify", 6220, 75))

        clusters = _find_compressible_transaction_clusters(
            coded_frame(rows),
            minimum_count=5,
        )

        self.assertEqual(clusters, [])

    def test_any_6220_transaction_blocks_the_entire_merchant(self):
        rows = bonus_expense_rows(counter=7820)
        rows.insert(
            2,
            expense_tx("BONUS SELFOSSI, **-8748", 6220, 75, counter=7810),
        )

        clusters = _find_compressible_transaction_clusters(
            coded_frame(rows),
            minimum_count=5,
        )

        self.assertEqual(clusters, [])

    def test_6220_blocks_only_its_merchant_not_a_counter_account(self):
        rows = bonus_expense_rows(counter=6220)
        rows.append(expense_tx("PARKA BORGARTUN", 6220, 75, counter=7820))

        clusters = _find_compressible_transaction_clusters(
            coded_frame(rows),
            minimum_count=5,
        )

        self.assertEqual(len(clusters), 1)
        self.assertEqual(clusters[0]["summary_text"], "BONUS")
        self.assertEqual(clusters[0]["counter_account"], "6220")
        self.assertTrue(clusters[0]["mixed_codes"])

    def test_positive_6220_category_also_blocks_the_merchant(self):
        rows = [
            income_tx(f"BONUS LOCATION {index}", 4000, index * 10)
            for index in range(1, 6)
        ]
        rows.append(income_tx("BONUS OTHER LOCATION", 6220, 75, counter=1007))

        clusters = _find_compressible_transaction_clusters(
            coded_frame(rows),
            minimum_count=5,
        )

        self.assertEqual(clusters, [])

    def test_excluded_6220_row_still_vetoes_the_merchant(self):
        rows = bonus_expense_rows()
        rows.append(expense_tx("BONUS SELFOSSI", 6220, 75))

        clusters = _find_compressible_transaction_clusters(
            coded_frame(rows),
            minimum_count=5,
            excluded_indices={5},
        )

        self.assertEqual(clusters, [])


class MixedCompressionModeTests(unittest.TestCase):
    def _run_two_modes(self, folder, automatic_first):
        coded_path = Path(folder) / "mixed_coded.xlsx"
        rows = bonus_expense_rows()
        rows.extend([
            expense_tx("Manual row one", 6220, 60),
            expense_tx("Manual row two", 6220, 70),
        ])
        coded_frame(rows).to_excel(coded_path, index=False)

        current, signature, error = _read_transaction_compression_workbook(coded_path)
        self.assertFalse(error)
        bonus_cluster = _find_compressible_transaction_clusters(current, minimum_count=5)[0]
        manual_cluster = cluster([5, 6], 130, "Handvirk þjöppun")

        if automatic_first:
            first_selection = [bonus_cluster]
        else:
            first_selection = [manual_cluster]
        first_result = _save_transaction_compressions(
            coded_path,
            current,
            first_selection,
            expected_signature=signature,
        )
        self.assertTrue(first_result["ok"], first_result.get("message"))

        refreshed, refreshed_signature, error = _read_transaction_compression_workbook(coded_path)
        self.assertFalse(error)
        source_groups = _source_groups_for_current_output(coded_path, refreshed)
        compressed_indices = {
            idx for idx, paired_indices in source_groups.items() if len(paired_indices) > 1
        }

        if automatic_first:
            remaining = [idx for idx in refreshed.index if idx not in compressed_indices]
            second_selection = [cluster(remaining, 130, "Handvirk þjöppun")]
        else:
            auto_clusters = _find_compressible_transaction_clusters(
                refreshed,
                minimum_count=5,
                excluded_indices=compressed_indices,
            )
            self.assertEqual(len(auto_clusters), 1)
            second_selection = auto_clusters

        second_result = _save_transaction_compressions(
            coded_path,
            refreshed,
            second_selection,
            expected_signature=refreshed_signature,
        )
        self.assertTrue(second_result["ok"], second_result.get("message"))
        self.assertEqual(first_result["reference_path"], second_result["reference_path"])

        final, _signature, error = _read_transaction_compression_workbook(coded_path)
        self.assertFalse(error)
        self.assertEqual(len(final), 2)
        self.assertEqual(set(final["TEXT"]), {"BONUS", "Handvirk þjöppun"})
        self.assertEqual(set(final["STATUS"]), {"Þjöppuð færsla"})

        reference = pd.read_excel(second_result["reference_path"], dtype=object, keep_default_na=False)
        self.assertEqual(list(reference.columns), REFERENCE_PRINT_COLUMNS)
        summaries = reference.loc[
            reference["TEXT"].isin({"BONUS", "Handvirk þjöppun"}),
            ["TEXT", "AMOUNT"],
        ]
        self.assertEqual(
            dict(zip(summaries["TEXT"], summaries["AMOUNT"])),
            {"BONUS": 1500, "Handvirk þjöppun": 130},
        )

        final_groups = _source_groups_for_current_output(coded_path, final)
        flattened = [index for group in final_groups.values() for index in group]
        self.assertEqual(sorted(flattened), list(range(7)))
        self.assertEqual(len(flattened), len(set(flattened)))

    def test_manual_and_automatic_modes_merge_in_both_orders(self):
        for automatic_first in (True, False):
            with self.subTest(automatic_first=automatic_first):
                with tempfile.TemporaryDirectory() as folder:
                    self._run_two_modes(folder, automatic_first)


class CombineCodedWorkbooksTests(unittest.TestCase):
    def test_recent_scan_returns_only_six_newest_coded_outputs(self):
        with tempfile.TemporaryDirectory() as folder:
            imports = Path(folder) / "imports"
            created = []
            for index in range(8):
                session = imports / f"session_{index}"
                session.mkdir(parents=True)
                name = f"client_{index}_coded.xlsx" if index != 3 else "client_3_coded1.xlsx"
                path = session / name
                path.write_bytes(b"coded")
                os.utime(path, ns=(index + 1, index + 1))
                created.append(str(path))

            raw = imports / "session_raw" / "raw.xlsx"
            raw.parent.mkdir()
            raw.write_bytes(b"raw")
            compressed = imports / "session_raw" / "client_coded_compressed_txs.xlsx"
            compressed.write_bytes(b"reference")
            os.utime(raw, ns=(100, 100))
            os.utime(compressed, ns=(101, 101))

            recent = find_recent_coded_files(folder)

            self.assertEqual(recent, list(reversed(created[-6:])))

    def test_multi_file_drop_data_keeps_paths_with_spaces(self):
        paths = _parse_drop_paths(r"{C:\Bank Files\first coded.xlsx} C:\second.xlsx")
        self.assertEqual(paths, [r"C:\Bank Files\first coded.xlsx", r"C:\second.xlsx"])

    def test_stacks_rows_into_combined_name_without_changing_sources(self):
        with tempfile.TemporaryDirectory() as folder:
            first = Path(folder) / "first_coded.xlsx"
            second = Path(folder) / "second_coded.xlsx"
            coded_frame([tx("first", 1000, 10)]).to_excel(first, index=False)
            coded_frame([tx("second", 2000, 20), tx("third", 3000, 30)]).to_excel(second, index=False)

            result = combine_coded_workbooks([first, second])

            output = Path(folder) / "first_coded_combined.xlsx"
            self.assertEqual(os.path.normcase(result["output_path"]), os.path.normcase(str(output)))
            self.assertEqual(result["file_count"], 2)
            self.assertEqual(result["row_count"], 3)
            combined = pd.read_excel(output, dtype=object, keep_default_na=False)
            self.assertEqual(combined["TEXT"].tolist(), ["first", "second", "third"])
            self.assertEqual(list(combined.columns), COLUMNS)
            unchanged_first = pd.read_excel(first, dtype=object, keep_default_na=False)
            unchanged_second = pd.read_excel(second, dtype=object, keep_default_na=False)
            self.assertEqual(unchanged_first["TEXT"].tolist(), ["first"])
            self.assertEqual(unchanged_second["TEXT"].tolist(), ["second", "third"])

    def test_combined_name_is_only_added_once(self):
        path = Path("folder") / "client_coded_combined.xlsx"
        self.assertEqual(Path(combined_output_path(path)).name, "client_coded_combined.xlsx")

    def test_rejects_a_different_column_layout_without_changing_first(self):
        with tempfile.TemporaryDirectory() as folder:
            first = Path(folder) / "first_coded.xlsx"
            second = Path(folder) / "second_coded.xlsx"
            coded_frame([tx("first", 1000, 10)]).to_excel(first, index=False)
            mismatched = coded_frame([tx("second", 2000, 20)]).drop(columns=["STATUS"])
            mismatched.to_excel(second, index=False)

            with self.assertRaises(CodedWorkbookError):
                combine_coded_workbooks([first, second])

            unchanged = pd.read_excel(first, dtype=object, keep_default_na=False)
            self.assertEqual(unchanged["TEXT"].tolist(), ["first"])

    def test_ignores_unnamed_side_notes_without_changing_sources(self):
        with tempfile.TemporaryDirectory() as folder:
            first = Path(folder) / "first_coded.xlsx"
            second = Path(folder) / "second_coded.xlsx"
            first_data = coded_frame([tx("first", 1000, 10)])
            first_data["Unnamed: 9"] = ["Kostnaður"]
            first_data["Unnamed: 10"] = [""]
            first_data["Unnamed: 11"] = ["6220"]
            first_data.to_excel(first, index=False)
            coded_frame([tx("second", 2000, 20)]).to_excel(second, index=False)

            result = combine_coded_workbooks([first, second])

            combined = pd.read_excel(
                result["output_path"],
                dtype=object,
                keep_default_na=False,
            )
            unchanged_first = pd.read_excel(
                first,
                dtype=object,
                keep_default_na=False,
            )
            self.assertEqual(list(combined.columns), COLUMNS)
            self.assertEqual(combined["TEXT"].tolist(), ["first", "second"])
            self.assertIn("Unnamed: 9", unchanged_first.columns)
            self.assertEqual(unchanged_first.loc[0, "Unnamed: 9"], "Kostnaður")


class RepeatedCompressionTests(unittest.TestCase):
    def test_refreshes_saved_codes_and_appends_reference_groups(self):
        with tempfile.TemporaryDirectory() as folder:
            coded_path = Path(folder) / "client_coded.xlsx"
            coded_frame([
                tx("one", 1000, 10, "1"),
                tx("two", 1000, 20, "2"),
                tx("three", 2000, 30, "3"),
                tx("four", 2000, 40, "4"),
            ]).to_excel(coded_path, index=False)

            current, signature, error = _read_transaction_compression_workbook(coded_path)
            self.assertFalse(error)
            first_result = _save_transaction_compressions(
                coded_path,
                current,
                [cluster([0, 1], 30, "first compression")],
                expected_signature=signature,
            )
            self.assertTrue(first_result["ok"], first_result.get("message"))

            renamed_reference = Path(folder) / "renamed_compression_reference.xlsx"
            Path(first_result["reference_path"]).rename(renamed_reference)
            self.assertTrue(
                _remember_compressed_reference_path(coded_path, renamed_reference)
            )

            # Simulate coding and saving the remaining rows between compression runs.
            edited = pd.read_excel(coded_path, dtype=object, keep_default_na=False)
            edited.loc[[1, 2], "DEBIT"] = 9999
            edited.loc[1, "TEXT"] = "edited source text"
            edited.to_excel(coded_path, index=False)

            refreshed, refreshed_signature, error = _read_transaction_compression_workbook(coded_path)
            self.assertFalse(error)
            self.assertEqual(refreshed.loc[1, "DEBIT"], 9999)
            self.assertEqual(refreshed.loc[1, "TEXT"], "edited source text")
            second_result = _save_transaction_compressions(
                coded_path,
                refreshed,
                [cluster([1, 2], 70, "second compression")],
                expected_signature=refreshed_signature,
            )
            self.assertTrue(second_result["ok"], second_result.get("message"))
            self.assertEqual(
                os.path.normcase(second_result["reference_path"]),
                os.path.normcase(str(renamed_reference)),
            )

            final_coded = pd.read_excel(coded_path, dtype=object, keep_default_na=False)
            self.assertEqual(len(final_coded), 2)
            self.assertEqual(final_coded.loc[1, "DEBIT"], 9999)
            self.assertEqual(final_coded.loc[1, "TEXT"], "second compression")

            reference = pd.read_excel(second_result["reference_path"], dtype=object, keep_default_na=False)
            self.assertEqual(list(reference.columns), REFERENCE_PRINT_COLUMNS)
            summaries = reference.loc[
                reference["TEXT"].isin({"first compression", "second compression"}),
                ["TEXT", "AMOUNT"],
            ]
            self.assertEqual(
                dict(zip(summaries["TEXT"], summaries["AMOUNT"])),
                {"first compression": 30, "second compression": 70},
            )

    def test_legacy_reference_columns_are_removed_when_rows_are_appended(self):
        with tempfile.TemporaryDirectory() as folder:
            reference_path = Path(folder) / "legacy_compressed_txs.xlsx"
            legacy = pd.DataFrame(
                [
                    {
                        "DATE": "01.01.2026",
                        "TEXT": "old row",
                        "DEBIT": 4055,
                        "Cnf": 90,
                        "ID": "",
                        "AMOUNT": 100,
                        "CREDIT": 7810,
                        "Positive/Negative": "-",
                        "STATUS": "coded",
                        "COMPRESSED_GROUP": 1,
                        "REFERENCE_TYPE": "deleted_tx",
                        "COMPRESSED_COUNT": 5,
                        "COMPRESSED_SUM_AMOUNT": 500,
                    }
                ]
            )
            legacy.to_excel(reference_path, index=False)
            new_rows = pd.DataFrame(
                [
                    {
                        "DATE": "02.01.2026",
                        "TEXT": "new row",
                        "DEBIT": 4061,
                        "ID": "",
                        "AMOUNT": 200,
                    }
                ]
            )

            combined = _append_reference_rows(reference_path, new_rows)

            self.assertEqual(list(combined.columns), REFERENCE_PRINT_COLUMNS)
            self.assertEqual(combined["TEXT"].tolist(), ["old row", "new row"])
            self.assertEqual(combined["DATE"].tolist(), ["01.01.2026", "02.01.2026"])

    def test_does_not_reuse_a_neighbor_workbooks_map_or_reference(self):
        with tempfile.TemporaryDirectory() as folder:
            first_path = Path(folder) / "first_coded.xlsx"
            coded_frame([
                tx("one", 1000, 10),
                tx("two", 1000, 20),
            ]).to_excel(first_path, index=False)
            current, signature, error = _read_transaction_compression_workbook(first_path)
            self.assertFalse(error)
            result = _save_transaction_compressions(
                first_path,
                current,
                [cluster([0, 1], 30, "first compression")],
                expected_signature=signature,
            )
            self.assertTrue(result["ok"], result.get("message"))

            second_path = Path(folder) / "second_coded.xlsx"
            coded_frame([tx("unrelated", 3000, 99)]).to_excel(second_path, index=False)
            second = pd.read_excel(second_path, dtype=object, keep_default_na=False)

            self.assertEqual(_source_groups_for_current_output(second_path, second), {})
            self.assertEqual(
                Path(_compressed_reference_path(second_path)).name,
                "second_coded_compressed_txs.xlsx",
            )

    def test_stable_map_mismatch_refuses_to_overwrite_source_linkage(self):
        with tempfile.TemporaryDirectory() as folder:
            coded_path = Path(folder) / "client_coded.xlsx"
            coded_frame([
                tx("one", 1000, 10),
                tx("two", 1000, 20),
                tx("three", 2000, 30),
            ]).to_excel(coded_path, index=False)
            current, signature, error = _read_transaction_compression_workbook(coded_path)
            self.assertFalse(error)
            first_result = _save_transaction_compressions(
                coded_path,
                current,
                [cluster([0, 1], 30, "first compression")],
                expected_signature=signature,
            )
            self.assertTrue(first_result["ok"], first_result.get("message"))

            changed = pd.read_excel(coded_path, dtype=object, keep_default_na=False)
            changed.loc[1, "AMOUNT"] = 999
            changed.to_excel(coded_path, index=False)
            refreshed, refreshed_signature, error = _read_transaction_compression_workbook(coded_path)
            self.assertFalse(error)
            map_path = Path(folder) / "compression_map.json"
            original_map = map_path.read_bytes()

            result = _save_transaction_compressions(
                coded_path,
                refreshed,
                [cluster([0, 1], 1029, "unsafe compression")],
                expected_signature=refreshed_signature,
            )

            self.assertFalse(result["ok"])
            self.assertIn("Upphæðir", result["message"])
            self.assertEqual(map_path.read_bytes(), original_map)
            unchanged = pd.read_excel(coded_path, dtype=object, keep_default_na=False)
            self.assertEqual(len(unchanged), 2)
            self.assertEqual(unchanged.loc[1, "AMOUNT"], 999)

    def test_refuses_to_overwrite_a_file_changed_after_refresh(self):
        with tempfile.TemporaryDirectory() as folder:
            coded_path = Path(folder) / "client_coded.xlsx"
            coded_frame([tx("one", 1000, 10), tx("two", 1000, 20)]).to_excel(coded_path, index=False)
            stale = pd.read_excel(coded_path, dtype=object, keep_default_na=False)
            stale_signature = _compression_file_signature(coded_path)

            changed = stale.copy()
            changed.loc[0, "DEBIT"] = 5555
            changed.to_excel(coded_path, index=False)

            result = _save_transaction_compressions(
                coded_path,
                stale,
                [cluster([0, 1], 30, "stale compression")],
                expected_signature=stale_signature,
            )

            self.assertFalse(result["ok"])
            current = pd.read_excel(coded_path, dtype=object, keep_default_na=False)
            self.assertEqual(current.loc[0, "DEBIT"], 5555)
            self.assertEqual(len(current), 2)


if __name__ == "__main__":
    unittest.main()

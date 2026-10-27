import unittest
from unittest.mock import patch

import pandas as pd

from tok.auto_coder import AutoCodeResult, apply_auto_debit_codes


class _FixedHistoryCoder:
    def ready(self):
        return True

    def predict(self, *_args, **_kwargs):
        return AutoCodeResult(
            code="9320",
            confidence=94,
            status="coded",
            reason="fixed merchant history",
            source="fixed_history",
        )


class _Fixed6220HistoryCoder:
    def ready(self):
        return True

    def predict(self, *_args, **_kwargs):
        return AutoCodeResult(
            code="6220",
            confidence=94,
            status="coded",
            reason="learned 6220 merchant history",
            source="fixed_history",
        )


def _output(rows):
    return pd.DataFrame(
        [
            {
                "DATE": date,
                "TEXT": text,
                "DEBIT": "",
                "ID": id_value,
                "AMOUNT": amount,
                "CREDIT": "7810",
                "Positive/Negative": "-",
                "_TRAINING_RAW_INDEX": raw_index,
            }
            for raw_index, date, text, id_value, amount in rows
        ]
    )


def _apply(output, source):
    with patch("tok.auto_coder.get_auto_coder", return_value=_FixedHistoryCoder()):
        return apply_auto_debit_codes(output, source_df=source)


class CollectionCostRuleTests(unittest.TestCase):
    def test_explicit_cost_wording_overrides_history_in_all_bank_layouts(self):
        output = _output(
            [
                (101, "16.06.2026", "Höldur ehf.", "6511740239", 1306),
                (102, "16.06.2026", "Höldur ehf.", "6511740239", 183521),
                (201, "03.06.2026", "Gildi", "5611952779", 457),
                (202, "03.06.2026", "Gildi", "5611952779", 46865),
                (301, "18.06.2026", "Sýn hf.", "4709051740", 30041),
                (302, "18.06.2026", "Sýn hf.", "4709051740", 12539),
                (401, "18.05.2026", "Birta", "4302696589", 8042),
                (402, "18.05.2026", "Birta", "4302696589", 8789),
            ]
        )
        source = pd.DataFrame(
            [
                {
                    "Texti": "Höldur ehf.",
                    "Textalykill": "Innheimtukrafa - kostnaður",
                    "Skýring greiðslu": "Innheimtukrafa - kostnaður",
                    "Tnr/Seðilnr.": "0050626",
                },
                {
                    "Texti": "Höldur ehf.",
                    "Textalykill": "Reikningur",
                    "Skýring greiðslu": "Reikningur",
                    "Tnr/Seðilnr.": "0050626",
                },
                {
                    "Skýring": "Gildi",
                    "Texti": "Innheimtukrafa - kostnaður",
                    "Seðilnúmer": "0100526",
                },
                {
                    "Skýring": "Gildi",
                    "Texti": "Iðgjöld",
                    "Seðilnúmer": "0100526",
                },
                {
                    "Mótaðili": "Sýn hf.",
                    "Tegund": "Kostnaður",
                    "Textalykill": "Kostnaður",
                    "Upplýsingar": "Kostnaður vegna kröfu",
                    "Seðilnr.": "260126",
                },
                {
                    "Mótaðili": "Sýn hf.",
                    "Tegund": "Innheimtuþjónusta",
                    "Textalykill": "Innheimtuþjónusta",
                    "Seðilnr.": "260126",
                },
                {
                    "Mótaðili": "Birta",
                    "Kennitala mótaðila": "4302696589",
                    "Tilvísun": "010526",
                    "Texti": "Kostnaður",
                },
                {
                    "Mótaðili": "Birta",
                    "Kennitala mótaðila": "4302696589",
                    "Tilvísun": "010526",
                    "Texti": "Iðgjöld",
                },
            ],
            index=[101, 102, 201, 202, 301, 302, 401, 402],
        )

        result = _apply(output, source)

        self.assertEqual(
            result["DEBIT"].tolist(),
            ["6220", "9320", "6220", "9320", "6220", "9320", "6220", "9320"],
        )
        self.assertEqual(
            result.loc[[0, 2, 4, 6], "_AUTO_CODE_SOURCE"].tolist(),
            ["hardcoded_collection_cost"] * 4,
        )
        self.assertEqual(result.loc[[0, 2, 4, 6], "CONFIDENCE"].tolist(), ["AD"] * 4)
        self.assertEqual(result.loc[[1, 3, 5, 7], "CONFIDENCE"].tolist(), [94] * 4)

    def test_three_and_four_row_dates_are_partitioned_without_contamination(self):
        rows = [
            (10, "18.06.2026", "DS lausnir ehf.", "5312023840", 91183),
            (11, "18.06.2026", "DS lausnir ehf.", "5312023840", 1563),
            (12, "18.06.2026", "DS lausnir ehf.", "5312023840", 173362),
            (20, "06.05.2026", "Tæki.is ehf.", "6503992859", 14137),
            (21, "06.05.2026", "Tæki.is ehf.", "6503992859", 125819),
            (22, "06.05.2026", "Tæki.is ehf.", "6503992859", 115148),
            (23, "06.05.2026", "Tæki.is ehf.", "6503992859", 271587),
        ]
        output = _output(rows)
        source_rows = []
        for raw_index, _date, _text, _id, _amount in rows:
            reference = {
                10: "0290526",
                11: "0270526",
                12: "0270526",
                20: "0071125",
                21: "0071125",
                22: "0101225",
                23: "0101225",
            }[raw_index]
            is_cost = raw_index in {11, 20, 22}
            source_rows.append(
                {
                    "Textalykill": "Innheimtukrafa - kostnaður" if is_cost else "Reikningur",
                    "Skýring greiðslu": "Innheimtukrafa - kostnaður" if is_cost else "Reikningur",
                    "Tnr/Seðilnr.": reference,
                }
            )
        source = pd.DataFrame(source_rows, index=[row[0] for row in rows])

        result = _apply(output, source)

        self.assertEqual(
            result["DEBIT"].tolist(),
            ["9320", "6220", "9320", "6220", "9320", "6220", "9320"],
        )

    def test_amount_difference_alone_never_creates_a_cost_code(self):
        output = _output(
            [
                (501, "25.06.2026", "Marpól ehf.", "5401942629", 7345),
                (502, "25.06.2026", "Marpól ehf.", "5401942629", 27016),
            ]
        )
        source = pd.DataFrame(
            [
                {"Textalykill": "Millifært", "Tnr/Seðilnr.": "0100626"},
                {"Textalykill": "Millifært", "Tnr/Seðilnr.": "0100626"},
            ],
            index=[501, 502],
        )

        result = _apply(output, source)

        self.assertEqual(result["DEBIT"].tolist(), ["9320", "9320"])
        self.assertEqual(result["_AUTO_CODE_SOURCE"].tolist(), ["fixed_history", "fixed_history"])

    def test_learned_6220_without_raw_cost_text_keeps_numeric_confidence(self):
        output = _output(
            [(551, "25.06.2026", "Venjulegur mótaðili", "5401942629", 7345)]
        )
        source = pd.DataFrame(
            [{"Skýring": "Reikningur", "Texti": "Venjulegur mótaðili"}],
            index=[551],
        )

        with patch("tok.auto_coder.get_auto_coder", return_value=_Fixed6220HistoryCoder()):
            result = apply_auto_debit_codes(output, source_df=source)

        self.assertEqual(result.loc[0, "DEBIT"], "6220")
        self.assertEqual(result.loc[0, "CONFIDENCE"], 94)
        self.assertEqual(result.loc[0, "_AUTO_CODE_SOURCE"], "fixed_history")

    def test_bare_cost_requires_a_matching_claim_reference(self):
        output = _output(
            [
                (601, "18.05.2026", "Birta", "4302696589", 8042),
                (602, "18.05.2026", "Birta", "4302696589", 8789),
            ]
        )
        source = pd.DataFrame(
            [
                {
                    "Kennitala mótaðila": "4302696589",
                    "Tilvísun": "A-100",
                    "Texti": "Kostnaður",
                },
                {
                    "Kennitala mótaðila": "4302696589",
                    "Tilvísun": "B-200",
                    "Texti": "Iðgjöld",
                },
            ],
            index=[601, 602],
        )

        result = _apply(output, source)

        self.assertEqual(result["DEBIT"].tolist(), ["9320", "9320"])


if __name__ == "__main__":
    unittest.main()

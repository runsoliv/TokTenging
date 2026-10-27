import unittest
from unittest.mock import patch
from pathlib import Path
import tempfile

import openpyxl
import pandas as pd

from tok.auto_coder import (
    AutoCodeResult,
    _hardcoded_bank_signal_rule,
    apply_auto_debit_codes,
)
from tok.bank_utils import autofit_excel_columns


class _UnavailableCoder:
    def ready(self):
        return False


class _CertainCoder:
    def ready(self):
        return True

    def predict(self, *_args, **_kwargs):
        return AutoCodeResult(
            code="9615",
            confidence=99,
            status="coded",
            reason="certain learned merchant history",
            source="global_exact_consistent",
        )


class TreasuryCollectionRuleTests(unittest.TestCase):
    def test_unknown_treasury_explanation_is_capped_at_seventy(self):
        output = pd.DataFrame({
            "TEXT": ["Ríkissjóðsinnheimtur", "Ríkissjóðsinnheimtur", "Annar mótaðili"],
            "DEBIT": ["", "", ""],
            "AMOUNT": [100, 100, 100],
            "Positive/Negative": ["-", "-", "-"],
        })
        source = pd.DataFrame({
            "Texti": ["Ný óþekkt skýring", "Bifreiðagjöld", "Ný óþekkt skýring"],
        })

        with patch("tok.auto_coder.get_auto_coder", return_value=_CertainCoder()):
            result = apply_auto_debit_codes(output, source_df=source)

        self.assertEqual(result["CONFIDENCE"].tolist(), [70, "AD", 99])
        self.assertEqual(result["DEBIT"].tolist(), ["9615", "4610", "9615"])
        self.assertEqual(
            result["_AUTO_CODE_SOURCE"].tolist(),
            ["global_exact_consistent", "hardcoded_bank_signal", "global_exact_consistent"],
        )

    def test_all_requested_explanations_have_mandatory_codes(self):
        cases = {
            "Gjöld utan staðgreiðslu": "9610",
            "Kílómetragjald": "4610",
            "Kílómetrjagjald": "4610",
            "Bifreiðagjöld": "4610",
            "Bifreiðagjald": "4610",
            "Gr.frestur aðflutningsgjalda": "9540",
            "Greiðslufrestur aðflutningsgjalda": "9540",
            "Tryggingagjald": "9615",
            "Staðgreiðsluskattur": "9615",
            "Kostnaður": "6220",
        }
        for explanation, expected in cases.items():
            with self.subTest(explanation=explanation):
                result = _hardcoded_bank_signal_rule(
                    "Ríkissjóðsinnheimtur",
                    f"Skýring: {explanation}",
                )
                self.assertIsNotNone(result)
                self.assertEqual(result.code, expected)
                self.assertEqual(result.confidence, 100)
                self.assertEqual(result.source, "hardcoded_bank_signal")

    def test_generic_cost_is_not_hardcoded_for_other_payees(self):
        result = _hardcoded_bank_signal_rule("Annar mótaðili", "Kostnaður")
        self.assertIsNone(result)

    def test_merchant_and_explanation_can_come_from_either_side(self):
        result = _hardcoded_bank_signal_rule(
            "Innheimtukrafa - kostnaður",
            "Ríkissjóðsinnheimtur",
        )
        self.assertIsNotNone(result)
        self.assertEqual(result.code, "6220")

    def test_rules_apply_even_when_learned_coder_is_unavailable(self):
        output = pd.DataFrame({
            "TEXT": [
                "Ríkissjóðsinnheimtur",
                "Ríkissjóðsinnheimtur",
                "Ríkissjóðsinnheimtur",
                "Ríkissjóðsinnheimtur",
                "Ríkissjóðsinnheimtur",
                "Ríkissjóðsinnheimtur",
                "Annar mótaðili",
            ],
            "DEBIT": ["", "", "", "", "", "", ""],
            "AMOUNT": [100] * 7,
            "Positive/Negative": ["-"] * 7,
        })
        source = pd.DataFrame({
            "Texti": [
                "Gjöld utan staðgreiðslu",
                "Kílómetragjald",
                "Tryggingagjald",
                "Staðgreiðsluskattur",
                "Kostnaður",
                "Eftirstöðvar innheimta",
                "Kostnaður",
            ],
        })

        with patch("tok.auto_coder.get_auto_coder", return_value=_UnavailableCoder()):
            result = apply_auto_debit_codes(output, source_df=source)

        self.assertEqual(
            result["DEBIT"].tolist(),
            ["9610", "4610", "9615", "9615", "6220", "", ""],
        )
        self.assertEqual(
            result["_AUTO_CODE_SOURCE"].tolist(),
            [
                "hardcoded_bank_signal",
                "hardcoded_bank_signal",
                "hardcoded_bank_signal",
                "hardcoded_bank_signal",
                "hardcoded_bank_signal",
                "unavailable",
                "unavailable",
            ],
        )
        self.assertEqual(
            result["CONFIDENCE"].tolist(),
            ["AD", "AD", "AD", "AD", "AD", "", ""],
        )

    def test_existing_and_manual_codes_are_protected_before_mandatory_rules(self):
        output = pd.DataFrame({
            "TEXT": ["Ríkissjóðsinnheimtur", "Ríkissjóðsinnheimtur"],
            "DEBIT": ["9999", "9999"],
            "AMOUNT": [100, 100],
            "Positive/Negative": ["-", "-"],
            "CONFIDENCE": ["MN", ""],
            "STATUS": ["coded", "coded"],
            "_AUTO_CODE_SOURCE": ["manual_review", "existing"],
        })
        source = pd.DataFrame({"Texti": ["Gjöld utan staðgreiðslu"] * 2})

        with patch("tok.auto_coder.get_auto_coder", return_value=_UnavailableCoder()):
            default_result = apply_auto_debit_codes(output, source_df=source)
            overwrite_result = apply_auto_debit_codes(
                output,
                source_df=source,
                overwrite_existing=True,
            )

        self.assertEqual(default_result["DEBIT"].tolist(), ["9999", "9999"])
        self.assertEqual(overwrite_result["DEBIT"].tolist(), ["9999", "9610"])
        self.assertEqual(overwrite_result.loc[0, "CONFIDENCE"], "MN")
        self.assertEqual(overwrite_result.loc[0, "_AUTO_CODE_SOURCE"], "manual_review")
        self.assertEqual(overwrite_result.loc[1, "CONFIDENCE"], "AD")

    def test_rules_cover_all_training_raw_bank_layouts(self):
        output = pd.DataFrame(
            {
                "TEXT": [
                    "Ríkissjóðsinnheimtur",
                    "Ríkissjóðsinnheimtur",
                    "Ríkissjóðsinnheimtur",
                    "Staðgreiðsluskattur",
                    "Ríkissjóðsinnheimtur",
                    "Ríkissjóðsinnheimtur",
                    "Ríkissjóðsinnheimtur",
                ],
                "DEBIT": ["", "", "", "", "", "", ""],
                "AMOUNT": [100] * 7,
                "Positive/Negative": ["-"] * 7,
                "_TRAINING_RAW_INDEX": [101, 202, 303, 404, 505, 606, 707],
            },
            index=[70, 60, 50, 40, 30, 20, 10],
        )
        source = pd.DataFrame(
            [
                {
                    "Mótaðili": "Ríkissjóðsinnheimtur",
                    "Tegund": "Kostnaður",
                    "Textalykill": "Kostnaður",
                },
                {
                    "Texti": "Ríkissjóðsinnheimtur",
                    "Skýring greiðslu": "Kílómetragjald",
                },
                {
                    "Skýring": "Ríkissjóðsinnheimtur",
                    "Nafn viðtakanda eða greiðanda": "Ríkissjóðsinnheimtur",
                    "Texti": "Tryggingagjald",
                },
                {
                    "Greiðslumóttakandi": "Ríkissjóðsinnheimtur",
                    "Lýsing færslu": "Staðgreiðsluskattur",
                },
                {
                    "Skýring": "Ríkissjóðsinnheimtur",
                    "Nafn viðtakanda eða greiðanda": "Ríkissjóðsinnheimtur",
                    "Texti": "Gjöld utan staðgreiðslu",
                },
                {
                    "Mótaðili": "Ríkissjóðsinnheimtur",
                    "Tegund": "Gjöld utan staðgreiðslu",
                    "Textalykill": "Gjöld utan staðgreiðslu",
                },
                {
                    "Texti": "Ríkissjóðsinnheimtur",
                    "Textalykill": "Gjöld utan staðgreiðslu",
                    "Skýring greiðslu": "Gjöld utan staðgreiðslu",
                },
            ],
            index=[101, 202, 303, 404, 505, 606, 707],
        )

        with patch("tok.auto_coder.get_auto_coder", return_value=_UnavailableCoder()):
            result = apply_auto_debit_codes(output, source_df=source)

        self.assertEqual(
            result["DEBIT"].tolist(),
            ["6220", "4610", "9615", "9615", "9610", "9610", "9610"],
        )
        self.assertEqual(
            result["_AUTO_CODE_SOURCE"].tolist(),
            ["hardcoded_bank_signal"] * 7,
        )
        self.assertEqual(result["CONFIDENCE"].tolist(), ["AD"] * 7)

    def test_ad_confidence_is_formatted_purple(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "treasury_ad.xlsx"
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(["TEXT", "Cnf"])
            sheet.append(["Ríkissjóðsinnheimtur", "AD"])
            sheet.append(["Ríkissjóðsinnheimtur", 70])
            workbook.save(path)

            autofit_excel_columns(path)

            formatted = openpyxl.load_workbook(path)
            formatted_sheet = formatted.active
            self.assertEqual(formatted_sheet["B2"].fill.fgColor.rgb, "00B4A7D6")
            self.assertNotEqual(
                formatted_sheet["B3"].fill.fgColor.rgb,
                formatted_sheet["B2"].fill.fgColor.rgb,
            )


if __name__ == "__main__":
    unittest.main()

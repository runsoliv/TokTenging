import unittest
from unittest.mock import patch

import pandas as pd

from tok import bank_formatter
from tok.auto_coder import AutoCodeResult, apply_auto_debit_codes


class _VehiclePredictionCoder:
    def ready(self):
        return True

    def predict(self, text, *_args, **_kwargs):
        code = {
            "Spáð 4635": "4635",
            "Spáð 4640": "4640",
            "Annar bifreiðalykill": "4620",
        }.get(str(text), "")
        return AutoCodeResult(
            code=code,
            confidence=91 if code else 0,
            status="coded" if code else "review",
            reason="learned vehicle merchant",
            source="context_model",
        )


class _SettingValue:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value


class VehicleExpenseCodeSettingTests(unittest.TestCase):
    def _output(self):
        return pd.DataFrame(
            {
                "TEXT": [
                    "Spáð 4635",
                    "Spáð 4640",
                    "Annar bifreiðalykill",
                    "Handvirk færsla",
                ],
                "DEBIT": ["", "", "", "4640"],
                "AMOUNT": [100, 200, 300, 400],
                "Positive/Negative": ["-", "-", "-", "-"],
            }
        )

    def test_selected_4635_normalizes_only_auto_predictions_in_pair(self):
        with patch(
            "tok.auto_coder.get_auto_coder",
            return_value=_VehiclePredictionCoder(),
        ):
            result = apply_auto_debit_codes(
                self._output(),
                vehicle_expense_code="4635",
            )

        self.assertEqual(result["DEBIT"].tolist(), ["4635", "4635", "4620", "4640"])
        self.assertEqual(
            result["_AUTO_CODE_SOURCE"].tolist(),
            ["context_model", "context_model", "context_model", "existing"],
        )
        self.assertEqual(result["CONFIDENCE"].tolist(), [91, 91, 91, ""])

    def test_selected_4640_normalizes_4635_even_in_restaurant_mode(self):
        with patch(
            "tok.auto_coder.get_auto_coder",
            return_value=_VehiclePredictionCoder(),
        ):
            result = apply_auto_debit_codes(
                self._output().iloc[[0]].copy(),
                industry_context="restaurant",
                vehicle_expense_code="4640",
            )

        self.assertEqual(result.iloc[0]["DEBIT"], "4640")
        self.assertEqual(result.iloc[0]["_AUTO_CODE_SOURCE"], "context_model")
        self.assertEqual(result.iloc[0]["CONFIDENCE"], 91)

    def test_missing_or_invalid_choice_preserves_prediction(self):
        for selected_code in ("", "9999", None):
            with self.subTest(selected_code=selected_code), patch(
                "tok.auto_coder.get_auto_coder",
                return_value=_VehiclePredictionCoder(),
            ):
                result = apply_auto_debit_codes(
                    self._output().iloc[[0]].copy(),
                    vehicle_expense_code=selected_code,
                )

            self.assertEqual(result.iloc[0]["DEBIT"], "4635")

    def test_bank_formatter_reads_and_forwards_selected_code(self):
        output = self._output().iloc[[0]].copy()
        expected = output.copy()
        with (
            patch.object(
                bank_formatter,
                "bank_vehicle_expense_code_var",
                _SettingValue("4635"),
                create=True,
            ),
            patch.object(bank_formatter, "AUTO_CODER_AVAILABLE", True),
            patch.object(
                bank_formatter,
                "apply_auto_debit_codes",
                return_value=expected,
            ) as apply_codes,
            patch.object(bank_formatter, "_auto_code_training_dir", return_value="training"),
            patch.object(bank_formatter, "_auto_code_key_dir", return_value="keys"),
            patch.object(bank_formatter, "_auto_code_memory_path", return_value="memory.json"),
            patch.object(bank_formatter, "_training_data_dir", return_value="training-data"),
            patch.object(bank_formatter, "_auto_code_mode", return_value="legacy"),
        ):
            selected_code = bank_formatter._bank_vehicle_expense_code()
            result = bank_formatter._apply_bank_auto_coding(
                output,
                None,
                "raw.xlsx",
                vehicle_expense_code=selected_code,
            )

        self.assertIs(result, expected)
        self.assertEqual(selected_code, "4635")
        self.assertEqual(apply_codes.call_args.kwargs["vehicle_expense_code"], "4635")

    def test_bank_formatter_hybrid_switch_runs_ml_after_legacy(self):
        output = self._output().iloc[[0]].copy()
        legacy = output.copy()
        hybrid = output.copy()
        hybrid.loc[hybrid.index[0], "DEBIT"] = "4200"
        with (
            patch.object(bank_formatter, "AUTO_CODER_AVAILABLE", True),
            patch.object(bank_formatter, "ML_RUNTIME_AVAILABLE", True),
            patch.object(bank_formatter, "apply_auto_debit_codes", return_value=legacy),
            patch.object(bank_formatter, "apply_ml_hybrid_codes", return_value=hybrid) as apply_ml,
            patch.object(bank_formatter, "_auto_code_training_dir", return_value="training"),
            patch.object(bank_formatter, "_auto_code_key_dir", return_value="keys"),
            patch.object(bank_formatter, "_auto_code_memory_path", return_value="memory.json"),
            patch.object(bank_formatter, "_training_data_dir", return_value="training-data"),
            patch.object(bank_formatter, "_auto_code_mode", return_value="ml_hybrid"),
            patch.object(bank_formatter, "_ml_artifact_dir", return_value="ml-artifacts"),
        ):
            result = bank_formatter._apply_bank_auto_coding(
                output,
                None,
                "raw.xlsx",
                vehicle_expense_code="4635",
                bank_type="arion_kort",
            )

        self.assertIs(result, hybrid)
        self.assertIs(apply_ml.call_args.args[0], legacy)
        self.assertEqual(apply_ml.call_args.kwargs["artifact_dir"], "ml-artifacts")
        self.assertEqual(apply_ml.call_args.kwargs["bank_type"], "arion_kort")

    def test_bank_formatter_hybrid_exception_keeps_legacy_result(self):
        output = self._output().iloc[[0]].copy()
        legacy = output.copy()
        with (
            patch.object(bank_formatter, "AUTO_CODER_AVAILABLE", True),
            patch.object(bank_formatter, "ML_RUNTIME_AVAILABLE", True),
            patch.object(bank_formatter, "apply_auto_debit_codes", return_value=legacy),
            patch.object(bank_formatter, "apply_ml_hybrid_codes", side_effect=RuntimeError("broken")),
            patch.object(bank_formatter, "_auto_code_training_dir", return_value="training"),
            patch.object(bank_formatter, "_auto_code_key_dir", return_value="keys"),
            patch.object(bank_formatter, "_auto_code_memory_path", return_value="memory.json"),
            patch.object(bank_formatter, "_training_data_dir", return_value="training-data"),
            patch.object(bank_formatter, "_auto_code_mode", return_value="ml_hybrid"),
            patch.object(bank_formatter, "_ml_artifact_dir", return_value="ml-artifacts"),
        ):
            result = bank_formatter._apply_bank_auto_coding(output, None, "raw.xlsx")

        self.assertIs(result, legacy)


if __name__ == "__main__":
    unittest.main()

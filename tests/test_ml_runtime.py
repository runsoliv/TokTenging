import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pandas as pd

from tok.ml_dataset import VEHICLE_EXPENSE_FAMILY
from tok.ml_model import ModelPrediction
from tok.ml_runtime import apply_ml_hybrid_codes, ml_runtime_status


def _policy(**overrides):
    payload = {
        "probability_threshold": 0.90,
        "margin_threshold": 0.10,
        "min_class_support": 5,
        "unseen_merchant_threshold": 0.94,
        "weak_text_threshold": 0.97,
        "sensitive_6220_threshold": 0.98,
        "target_precision": 0.95,
        "validation_precision": 0.96,
        "validation_coverage": 0.20,
        "validation_accepted": 20,
        "production_enabled": True,
    }
    payload.update(overrides)
    return payload


def _loaded(model, **policy_overrides):
    metadata = SimpleNamespace(
        model_version="fixture-v1",
        training_row_count=100,
        evaluation_summary={
            "gate_passed": True,
            "confidence_policy": _policy(**policy_overrides),
        },
    )
    return SimpleNamespace(model=model, metadata=metadata)


class _Model:
    def __init__(self, predictions):
        self.predictions = list(predictions)
        self.records = None

    def predict_records(self, records, top_k=3):
        self.records = list(records)
        return self.predictions


class MlRuntimeTests(unittest.TestCase):
    @staticmethod
    def _legacy(rows=None):
        rows = rows or [{
            "DATE": "01.08.2026",
            "TEXT": "Nýr söluaðili",
            "DEBIT": "4510",
            "ID": "",
            "AMOUNT": 1000,
            "CREDIT": "7810",
            "Positive/Negative": "-",
            "STATUS": "review needed",
            "CONFIDENCE": 20,
            "_AUTO_CODE_SOURCE": "fallback",
        }]
        return pd.DataFrame(rows)

    def _run_with_model(self, frame, model, **kwargs):
        store = Mock()
        store.load_active.return_value = _loaded(model)
        with patch("tok.ml_runtime.ArtifactStore", return_value=store):
            return apply_ml_hybrid_codes(frame, bank_type="arion", **kwargs)

    def test_missing_artifact_returns_legacy_values_unchanged(self):
        frame = self._legacy()
        with tempfile.TemporaryDirectory() as temp_dir:
            result = apply_ml_hybrid_codes(frame, artifact_dir=Path(temp_dir))
            status = ml_runtime_status(Path(temp_dir))
        pd.testing.assert_frame_equal(result, frame)
        self.assertFalse(status["available"])

    def test_ungated_policy_never_reaches_model_inference(self):
        frame = self._legacy()
        model = _Model([
            ModelPrediction("4200", 0.999, 0.90, (("4200", 0.999),), 100)
        ])
        loaded = _loaded(model)
        loaded.metadata.evaluation_summary["gate_passed"] = False
        store = Mock()
        store.load_active.return_value = loaded

        with patch("tok.ml_runtime.ArtifactStore", return_value=store):
            result = apply_ml_hybrid_codes(frame, bank_type="arion")

        pd.testing.assert_frame_equal(result, frame)
        self.assertIsNone(model.records)

    def test_only_review_row_is_replaced_and_internal_provenance_is_recorded(self):
        rows = self._legacy().to_dict("records")
        protected = dict(rows[0])
        protected.update({
            "TEXT": "Handvirkt",
            "DEBIT": "9999",
            "STATUS": "coded",
            "CONFIDENCE": "MN",
            "_AUTO_CODE_SOURCE": "manual_review",
        })
        deterministic = dict(rows[0])
        deterministic.update({
            "TEXT": "Ríkissjóðsinnheimtur",
            "DEBIT": "9615",
            "STATUS": "coded",
            "CONFIDENCE": "AD",
            "_AUTO_CODE_SOURCE": "hardcoded_bank_signal",
        })
        frame = self._legacy(rows + [protected, deterministic])
        prediction = ModelPrediction("4200", 0.99, 0.70, (("4200", 0.99),), 30)
        model = _Model([prediction])

        result = self._run_with_model(frame, model)

        self.assertEqual(len(model.records), 1)
        self.assertEqual(result.loc[0, "DEBIT"], "4200")
        self.assertEqual(result.loc[0, "CONFIDENCE"], 99)
        self.assertEqual(result.loc[0, "_AUTO_CODE_SOURCE"], "ml")
        self.assertEqual(result.loc[0, "_AUTO_CODE_ML_VERSION"], "fixture-v1")
        self.assertEqual(result.loc[1, "DEBIT"], "9999")
        self.assertEqual(result.loc[1, "CONFIDENCE"], "MN")
        self.assertEqual(result.loc[2, "DEBIT"], "9615")
        self.assertEqual(result.loc[2, "CONFIDENCE"], "AD")

    def test_vehicle_family_is_resolved_only_through_selected_setting(self):
        prediction = ModelPrediction(
            VEHICLE_EXPENSE_FAMILY,
            0.99,
            0.70,
            ((VEHICLE_EXPENSE_FAMILY, 0.99),),
            30,
        )
        result = self._run_with_model(
            self._legacy(),
            _Model([prediction]),
            vehicle_expense_code="4635",
        )
        self.assertEqual(result.loc[0, "DEBIT"], "4635")

    def test_unknown_treasury_description_is_never_ml_accepted(self):
        frame = self._legacy()
        frame.loc[0, "TEXT"] = "Ríkissjóðsinnheimtur"
        source = pd.DataFrame({"Skýring": ["Ný óþekkt skatttegund"]})
        prediction = ModelPrediction("9615", 0.999, 0.90, (("9615", 0.999),), 100)

        result = self._run_with_model(frame, _Model([prediction]), source_df=source)

        pd.testing.assert_frame_equal(result, frame)

    def test_model_failure_returns_entire_legacy_frame_without_partial_changes(self):
        class BrokenModel:
            def predict_records(self, _records, top_k=3):
                raise RuntimeError("broken fixture")

        frame = self._legacy()
        result = self._run_with_model(frame, BrokenModel())
        pd.testing.assert_frame_equal(result, frame)


if __name__ == "__main__":
    unittest.main()

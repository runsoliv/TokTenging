import unittest
import tempfile
from pathlib import Path

import pandas as pd

from tok.ml_evaluation import ConfidencePolicy, make_grouped_folds
from tok.ml_model import LocalNgramClassifier, NgramModelConfig
from tok.ml_training import (
    _deterministic_prediction,
    _legacy_prediction,
    activate_evaluated_model,
    evaluate_ml_config,
    fit_legacy_evaluation_coder,
    hybrid_predictions,
)


def _row(example_id, session, merchant, target, *, raw="", direction="outgoing"):
    return {
        "example_id": example_id,
        "semantic_identity": f"identity-{example_id}",
        "session_id": session,
        "client_id": session.split("-")[0],
        "bank_type": "landsbanki",
        "date": "2026-07-01",
        "text": merchant,
        "normalized_text": merchant.lower(),
        "merchant_family": merchant.lower(),
        "raw_explanation": raw,
        "valid_kennitala": "",
        "amount": 1000,
        "signed_amount": -1000 if direction == "outgoing" else 1000,
        "direction": direction,
        "target_code": target,
        "model_target": target,
        "raw_pair_status": "coded_only",
        "weak_text": False,
    }


class LegacyBaselineTests(unittest.TestCase):
    def test_fold_fitted_legacy_history_matches_without_writing_memory(self):
        train = pd.DataFrame([
            _row("a1", "client-a-1", "Bonus", "4061"),
            _row("a2", "client-a-2", "Bonus", "4061"),
            _row("b1", "client-a-3", "Vodafone", "4030"),
        ])
        coder = fit_legacy_evaluation_coder(train)

        result = _legacy_prediction(coder, pd.Series(_row("test", "client-z-1", "Bonus", "4061")))

        self.assertEqual(result["predicted"], "4061")
        self.assertTrue(result["accepted"])
        self.assertIn("consistent", result["source"])

    def test_deterministic_treasury_rule_is_authoritative(self):
        row = pd.Series(_row(
            "rsk",
            "client-a-1",
            "Rikissjodsinnheimtur",
            "4610",
            raw="Kilometragjald",
        ))

        result = _deterministic_prediction(row)

        self.assertEqual(result["predicted"], "4610")
        self.assertTrue(result["accepted"])


class GroupedModelHarnessTests(unittest.TestCase):
    def test_small_grouped_model_evaluation_returns_each_row_once(self):
        rows = []
        for session in range(6):
            rows.append(_row(f"food-{session}", f"client-{session}", "Bonus", "4061"))
            rows.append(_row(f"phone-{session}", f"client-{session}", "Vodafone", "4030"))
        frame = pd.DataFrame(rows)
        folds = make_grouped_folds(
            frame,
            name="unseen_session",
            group_columns=("session_id",),
            n_splits=3,
        )
        config = NgramModelConfig(
            name="test-centroid",
            algorithm="centroid",
            dimensions=1024,
            char_ngram_max=4,
        )

        predictions = evaluate_ml_config(frame, folds, config)

        self.assertEqual(len(predictions), len(frame))
        self.assertEqual(set(predictions["example_id"]), set(frame["example_id"]))
        self.assertTrue(predictions["probability"].between(0, 1).all())

    def test_hybrid_keeps_accepted_legacy_and_uses_only_gated_ml_for_review(self):
        frame = pd.DataFrame([
            _row("legacy", "s1", "Bonus", "4061"),
            _row("ml", "s2", "Unknown", "4030"),
        ])
        legacy = pd.DataFrame([
            {"example_id": "legacy", "predicted": "4061", "accepted": True, "source": "global_exact_consistent"},
            {"example_id": "ml", "predicted": "4510", "accepted": False, "source": "fallback"},
        ])
        ml = pd.DataFrame([
            {
                "example_id": "legacy", "predicted": "9999", "probability": 0.99,
                "margin": 0.8, "class_support": 50, "merchant_familiar": True,
                "weak_text": False, "unknown_treasury": False,
            },
            {
                "example_id": "ml", "predicted": "4030", "probability": 0.99,
                "margin": 0.8, "class_support": 50, "merchant_familiar": True,
                "weak_text": False, "unknown_treasury": False,
            },
        ])
        policy = ConfidencePolicy(
            probability_threshold=0.9,
            margin_threshold=0.2,
            min_class_support=5,
            unseen_merchant_threshold=0.95,
            weak_text_threshold=0.97,
            sensitive_6220_threshold=0.99,
            production_enabled=True,
        )

        hybrid = hybrid_predictions(frame, legacy, ml, policy).set_index("example_id")

        self.assertEqual(hybrid.loc["legacy", "predicted"], "4061")
        self.assertEqual(hybrid.loc["legacy", "source"], "legacy:global_exact_consistent")
        self.assertEqual(hybrid.loc["ml", "predicted"], "4030")
        self.assertEqual(hybrid.loc["ml", "source"], "ml")


class ArtifactActivationGateTests(unittest.TestCase):
    @staticmethod
    def _report(gate_passed):
        check = {
            "meets_95_percent_ml_precision": bool(gate_passed),
            "no_accuracy_regression": True,
        }
        return {
            "evaluation_schema_version": "fixture-v1",
            "created_at": "2026-08-31T12:00:00+00:00",
            "dataset": {"fingerprint": "a" * 64},
            "production_model_rows": 2,
            "production_class_count": 2,
            "selected_model": {"config": {"name": "fixture"}},
            "gate_checks": {
                "unseen_session": dict(check),
                "unseen_client": dict(check),
                "unseen_merchant": dict(check),
            },
            "gate_passed": bool(gate_passed),
        }

    @staticmethod
    def _policy(enabled):
        return ConfidencePolicy(
            probability_threshold=0.9,
            margin_threshold=0.2,
            min_class_support=1,
            unseen_merchant_threshold=0.95,
            weak_text_threshold=0.97,
            sensitive_6220_threshold=0.99,
            target_precision=0.95,
            validation_precision=0.96 if enabled else 0.50,
            validation_coverage=0.10,
            validation_accepted=30 if enabled else 0,
            production_enabled=bool(enabled),
        )

    def test_failed_gate_does_not_create_or_replace_artifact(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            artifact_dir = Path(temp_dir) / "models"
            result = activate_evaluated_model(
                self._report(False),
                object(),
                self._policy(False),
                artifact_dir=artifact_dir,
            )
            self.assertFalse(result["activated"])
            self.assertFalse(artifact_dir.exists())

    def test_passed_gate_atomically_activates_reloadable_model(self):
        model = LocalNgramClassifier(
            NgramModelConfig(
                name="activation_test_centroid",
                algorithm="centroid",
                dimensions=1024,
            )
        ).fit(
            ["merchant alpha", "merchant beta"],
            ["4000", "4100"],
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            result = activate_evaluated_model(
                self._report(True),
                model,
                self._policy(True),
                artifact_dir=temp_dir,
            )
            self.assertTrue(result["activated"])
            self.assertTrue(Path(result["artifact_path"]).is_file())
            self.assertTrue((Path(temp_dir) / "active_model.json").is_file())


if __name__ == "__main__":
    unittest.main()

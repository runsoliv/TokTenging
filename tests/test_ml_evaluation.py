import unittest

import pandas as pd

from tok.ml_evaluation import (
    ConfidencePolicy,
    _merchant_evaluation_components,
    audit_fold_leakage,
    calibration_summary,
    classification_metrics,
    make_grouped_folds,
    prediction_is_accepted,
    select_confidence_policy,
    top_k_accuracy,
)


class GroupedSplitTests(unittest.TestCase):
    @staticmethod
    def _frame():
        rows = []
        for session in range(10):
            for position in range(4):
                rows.append({
                    "session_id": f"session-{session}",
                    "client_id": f"client-{session // 2}",
                    "merchant_family": f"merchant-{position}",
                    "semantic_identity": f"transaction-{session}-{position}",
                    "model_target": "4061" if position < 3 else "6220",
                })
        return pd.DataFrame(rows)

    def test_groups_and_semantic_identities_never_cross_a_fold(self):
        frame = self._frame()
        folds = make_grouped_folds(
            frame,
            name="session",
            group_columns=("session_id",),
            n_splits=5,
        )

        self.assertEqual(len(folds), 5)
        all_test = []
        for fold in folds:
            audit = audit_fold_leakage(frame, fold)
            self.assertTrue(audit["ok"], audit)
            all_test.extend(fold.test_indices)
        self.assertEqual(sorted(all_test), list(range(len(frame))))

    def test_duplicate_identity_connects_otherwise_separate_groups(self):
        frame = self._frame()
        frame.loc[0, "semantic_identity"] = "shared"
        frame.loc[4, "semantic_identity"] = "shared"

        folds = make_grouped_folds(
            frame,
            name="session",
            group_columns=("session_id",),
            n_splits=4,
        )

        containing = [fold for fold in folds if 0 in fold.test_indices or 4 in fold.test_indices]
        self.assertEqual(len(containing), 1)
        self.assertIn(0, containing[0].test_indices)
        self.assertIn(4, containing[0].test_indices)

    def test_multiple_group_dimensions_are_each_protected(self):
        frame = self._frame()
        folds = make_grouped_folds(
            frame,
            name="client_and_session",
            group_columns=("client_id", "session_id"),
            n_splits=4,
        )

        for fold in folds:
            audit = audit_fold_leakage(frame, fold)
            self.assertTrue(audit["ok"], audit)
            self.assertFalse(audit["group_overlap"]["client_id"])
            self.assertFalse(audit["group_overlap"]["session_id"])

    def test_split_is_reproducible(self):
        frame = self._frame()
        first = make_grouped_folds(frame, name="client", group_columns=("client_id",), n_splits=5)
        second = make_grouped_folds(frame, name="client", group_columns=("client_id",), n_splits=5)
        self.assertEqual(first, second)

    def test_unseen_merchant_components_link_family_and_kennitala(self):
        frame = pd.DataFrame([
            {"merchant_family": "", "valid_kennitala": "1234567890", "normalized_text": "generic"},
            {"merchant_family": "merchant-a", "valid_kennitala": "1234567890", "normalized_text": "merchant a"},
            {"merchant_family": "merchant-a", "valid_kennitala": "", "normalized_text": "merchant a branch"},
            {"merchant_family": "merchant-b", "valid_kennitala": "9999999999", "normalized_text": "merchant b"},
        ])

        components = _merchant_evaluation_components(frame)

        self.assertEqual(components[0], components[1])
        self.assertEqual(components[1], components[2])
        self.assertNotEqual(components[2], components[3])


class EvaluationMetricTests(unittest.TestCase):
    def test_classification_and_top_k_metrics(self):
        metrics = classification_metrics(["A", "A", "B"], ["A", "B", "B"])
        self.assertAlmostEqual(metrics["accuracy"], 2 / 3)
        self.assertGreater(metrics["macro_f1"], 0)
        self.assertAlmostEqual(
            top_k_accuracy(["A", "B"], [["C", "A"], ["B", "A"]], 2),
            1.0,
        )

    def test_calibration_is_reported_in_bins(self):
        summary = calibration_summary([0.9, 0.8, 0.2], [True, False, False], bins=5)
        self.assertGreaterEqual(summary["ece"], 0)
        self.assertGreaterEqual(summary["brier_top_label"], 0)
        self.assertTrue(summary["bins"])


class ConfidencePolicyTests(unittest.TestCase):
    @staticmethod
    def _policy():
        return ConfidencePolicy(
            probability_threshold=0.8,
            margin_threshold=0.2,
            min_class_support=5,
            unseen_merchant_threshold=0.9,
            weak_text_threshold=0.92,
            sensitive_6220_threshold=0.95,
            production_enabled=True,
        )

    def test_rare_unknown_and_sensitive_predictions_abstain(self):
        base = {
            "predicted": "4061",
            "probability": 0.9,
            "margin": 0.4,
            "class_support": 20,
            "merchant_familiar": True,
            "weak_text": False,
        }
        self.assertEqual(prediction_is_accepted(base, self._policy()), (True, "accepted"))
        self.assertEqual(
            prediction_is_accepted({**base, "class_support": 1}, self._policy())[1],
            "rare_or_unsupported_class",
        )
        self.assertEqual(
            prediction_is_accepted({**base, "unknown_treasury": True}, self._policy())[1],
            "unknown_treasury",
        )
        self.assertFalse(prediction_is_accepted({**base, "predicted": "6220", "probability": 0.94}, self._policy())[0])
        self.assertFalse(prediction_is_accepted({**base, "weak_text": True, "probability": 0.91}, self._policy())[0])

    def test_non_finite_or_malformed_predictions_fail_closed(self):
        base = {
            "predicted": "4061",
            "probability": 0.99,
            "margin": 0.5,
            "class_support": 20,
            "merchant_familiar": True,
            "weak_text": False,
        }
        for mutation in (
            {"probability": float("nan")},
            {"probability": float("inf")},
            {"probability": 1.1},
            {"margin": float("nan")},
            {"margin": -0.1},
            {"class_support": "invalid"},
            {"predicted": ""},
        ):
            with self.subTest(mutation=mutation):
                accepted, reason = prediction_is_accepted({**base, **mutation}, self._policy())
                self.assertFalse(accepted)
                self.assertEqual(reason, "invalid_prediction_values")

    def test_disabled_or_invalid_policy_fails_closed(self):
        record = {
            "predicted": "4061",
            "probability": 0.99,
            "margin": 0.5,
            "class_support": 20,
            "merchant_familiar": True,
            "weak_text": False,
        }
        disabled = ConfidencePolicy(
            probability_threshold=0.8,
            margin_threshold=0.2,
            min_class_support=5,
            unseen_merchant_threshold=0.9,
            weak_text_threshold=0.92,
            sensitive_6220_threshold=0.95,
        )
        self.assertEqual(
            prediction_is_accepted(record, disabled),
            (False, "policy_not_enabled"),
        )
        with self.assertRaises(ValueError):
            ConfidencePolicy.from_dict({
                **disabled.to_dict(),
                "probability_threshold": float("nan"),
            })

    def test_policy_is_selected_from_held_out_precision_not_raw_probability(self):
        rows = []
        for index in range(100):
            rows.append({
                "correct": index < 97,
                "probability": 0.9,
                "margin": 0.5,
                "class_support": 50,
                "merchant_familiar": True,
                "weak_text": False,
                "predicted": "4061",
                "unknown_treasury": False,
            })
        policy = select_confidence_policy(pd.DataFrame(rows), minimum_accepted=25)
        self.assertTrue(policy.production_enabled)
        self.assertGreaterEqual(policy.validation_precision, 0.95)
        self.assertEqual(policy.validation_accepted, 100)

    def test_policy_stays_disabled_when_precision_gate_is_not_met(self):
        rows = [{
            "correct": index % 2 == 0,
            "probability": 0.9,
            "margin": 0.5,
            "class_support": 50,
            "merchant_familiar": True,
            "weak_text": False,
            "predicted": "4061",
            "unknown_treasury": False,
        } for index in range(100)]
        policy = select_confidence_policy(pd.DataFrame(rows), minimum_accepted=25)
        self.assertFalse(policy.production_enabled)


if __name__ == "__main__":
    unittest.main()

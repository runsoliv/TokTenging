import unittest

from tok.ml_dataset import VEHICLE_EXPENSE_FAMILY
from tok.ml_features import (
    amount_bucket,
    build_feature_document,
    model_target_code,
)


class MlFeatureTests(unittest.TestCase):
    def test_document_contains_only_controlled_non_leaking_context(self):
        row = {
            "text": "FACEBK *AB12345678 **-8748",
            "raw_explanation": "Auglýsing",
            "merchant_family": "FACEBK",
            "bank_type": "islandsbanki",
            "direction": "outgoing",
            "amount": 12_345,
            "date": "2026-08-31",
            "valid_kennitala": "1234567890",
            "weak_text": False,
            "client_id": "secret-client-shortcut",
            "session_id": "secret-session",
            "target_code": "4050",
            "Cnf": "AD",
            "STATUS": "coded",
            "_AUTO_CODE_SOURCE": "manual_review",
        }

        document = build_feature_document(row)

        self.assertIn("text facebk ab", document)
        self.assertNotIn("ab12345678", document)
        self.assertNotIn("text facebk ab 8748", document)
        self.assertIn("raw_explanation auglysing", document)
        self.assertIn("kt1234567890", document)
        self.assertNotIn("secret-client", document)
        self.assertNotIn("secret-session", document)
        self.assertNotIn("4050", document)
        self.assertNotIn("manual_review", document)
        self.assertNotIn(" coded", document)

    def test_missing_raw_and_invalid_kennitala_are_explicit_but_safe(self):
        document = build_feature_document({
            "text": "Reikningur",
            "raw_explanation": "",
            "bank_type": "landsbanki",
            "direction": "outgoing",
            "amount": 99,
            "date": "",
            "valid_kennitala": "123456789",
            "weak_text": True,
        })

        self.assertIn("raw_missing", document)
        self.assertIn("kennitala_missing", document)
        self.assertIn("month_missing", document)
        self.assertIn("weak_text", document)
        self.assertNotIn("123456789", document)

    def test_vehicle_codes_share_one_model_family(self):
        self.assertEqual(model_target_code("4635"), VEHICLE_EXPENSE_FAMILY)
        self.assertEqual(model_target_code("4640"), VEHICLE_EXPENSE_FAMILY)
        self.assertEqual(model_target_code("4061"), "4061")

    def test_amount_buckets_are_conservative_and_stable(self):
        self.assertEqual(amount_bucket(None), "amount_missing")
        self.assertEqual(amount_bucket(1000), "amount_0000_001k")
        self.assertEqual(amount_bucket(1001), "amount_001k_003k")
        self.assertEqual(amount_bucket(-800000), "amount_750k_plus")


if __name__ == "__main__":
    unittest.main()

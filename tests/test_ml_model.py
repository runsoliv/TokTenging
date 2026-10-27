import pickle
import unittest

from tok.ml_model import (
    DEFAULT_CANDIDATE_CONFIGS,
    LocalNgramClassifier,
    NgramModelConfig,
    hashed_document,
)


class LocalNgramClassifierTests(unittest.TestCase):
    @staticmethod
    def _documents():
        return [
            "bank_arion direction_outgoing text bonus skeifan merchant_family bonus",
            "bank_landsbanki direction_outgoing text bonus smaralind merchant_family bonus",
            "bank_arion direction_outgoing text kronan merchant_family kronan",
            "bank_landsbanki direction_outgoing text vodafone merchant_family vodafone",
            "bank_arion direction_outgoing text siminn merchant_family siminn",
            "bank_landsbanki direction_outgoing text nova merchant_family nova",
        ]

    @staticmethod
    def _labels():
        return ["4061", "4061", "4061", "4030", "4030", "4030"]

    def test_feature_hashing_is_deterministic_and_bounded(self):
        config = DEFAULT_CANDIDATE_CONFIGS[0]
        first_indices, first_values = hashed_document(self._documents()[0], config)
        second_indices, second_values = hashed_document(self._documents()[0], config)

        self.assertEqual(first_indices.tolist(), second_indices.tolist())
        self.assertEqual(first_values.tolist(), second_values.tolist())
        self.assertTrue(all(0 <= value < config.dimensions for value in first_indices))

    def test_both_candidate_algorithms_fit_and_return_ranked_probabilities(self):
        for config in DEFAULT_CANDIDATE_CONFIGS:
            with self.subTest(config=config.name):
                model = LocalNgramClassifier(config).fit(self._documents(), self._labels())
                predictions = model.predict([
                    "bank_arion direction_outgoing text bonus akureyri merchant_family bonus",
                    "bank_arion direction_outgoing text vodafone merchant_family vodafone",
                ])

                self.assertEqual([prediction.predicted for prediction in predictions], ["4061", "4030"])
                for prediction in predictions:
                    self.assertGreaterEqual(prediction.probability, 0.0)
                    self.assertLessEqual(prediction.probability, 1.0)
                    self.assertGreaterEqual(prediction.margin, 0.0)
                    self.assertEqual(len(prediction.candidates), 2)

    def test_probability_temperature_is_fitted_on_known_calibration_labels(self):
        model = LocalNgramClassifier(DEFAULT_CANDIDATE_CONFIGS[0]).fit(
            self._documents(),
            self._labels(),
        )

        temperature = model.calibrate_temperature(self._documents(), self._labels())

        self.assertGreater(temperature, 0)
        self.assertEqual(model.temperature_, temperature)

    def test_model_pickle_round_trip_has_fixed_inference(self):
        config = NgramModelConfig(
            name="fixture",
            algorithm="centroid",
            dimensions=4096,
            char_ngram_max=4,
        )
        model = LocalNgramClassifier(config).fit(self._documents(), self._labels())
        query = ["direction_outgoing text bonus merchant_family bonus"]
        before = model.predict(query)[0]

        restored = pickle.loads(pickle.dumps(model))
        after = restored.predict(query)[0]

        self.assertEqual(before.predicted, after.predicted)
        self.assertAlmostEqual(before.probability, after.probability, places=7)
        self.assertEqual(restored.metadata()["schema_version"], model.metadata()["schema_version"])


if __name__ == "__main__":
    unittest.main()

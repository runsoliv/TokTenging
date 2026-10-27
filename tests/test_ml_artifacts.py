import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd

from tok.ml_artifacts import (
    ARTIFACT_SCHEMA_VERSION,
    ArtifactCorrupt,
    ArtifactError,
    ArtifactIncompatible,
    ArtifactMetadata,
    ArtifactStore,
    ArtifactUnavailable,
)
from tok.ml_features import FEATURE_VERSION
from tok.ml_model import LocalNgramClassifier, NgramModelConfig


class FeatureTaggedModel:
    feature_version = "different-features"


class ArtifactStoreTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.store = ArtifactStore(self.root)

    def tearDown(self):
        self.temp_dir.cleanup()

    @staticmethod
    def _metadata(version="model-v1", feature=FEATURE_VERSION):
        return ArtifactMetadata(
            model_version=version,
            training_data_fingerprint="a" * 64,
            training_row_count=100,
            class_count=3,
            feature_version=feature,
            evaluation_summary={"precision": 0.97, "coverage": 0.25},
            production_enabled=True,
            pandas_version=pd.__version__,
        )

    def test_model_serializes_reloads_and_records_required_metadata(self):
        model = {"classes": ["4061", "6220", "9410"]}

        activated = self.store.activate(model, self._metadata())
        loaded = self.store.load_active(expected_feature_version=FEATURE_VERSION)

        self.assertEqual(loaded.model, model)
        self.assertEqual(loaded.metadata.training_row_count, 100)
        self.assertEqual(loaded.metadata.class_count, 3)
        self.assertTrue(loaded.metadata.created_at)
        self.assertEqual(loaded.checksum, activated.checksum)
        self.assertTrue(loaded.artifact_path.is_file())

    def test_missing_corrupt_and_incompatible_artifacts_fail_safely(self):
        with self.assertRaises(ArtifactUnavailable):
            self.store.load_active(expected_feature_version=FEATURE_VERSION)

        self.store.activate({"ok": True}, self._metadata())
        with self.assertRaises(ArtifactIncompatible):
            self.store.load_active(expected_feature_version="features-v2")

        pointer = json.loads(self.store.pointer_path.read_text(encoding="utf-8"))
        artifact = self.root / pointer["artifact_file"]
        artifact.write_bytes(b"corrupt")
        with self.assertRaises(ArtifactCorrupt):
            self.store.load_active(expected_feature_version=FEATURE_VERSION)

        status = self.store.status(expected_feature_version=FEATURE_VERSION)
        self.assertFalse(status["available"])
        self.assertEqual(status["reason"], "ArtifactCorrupt")

    def test_failed_pointer_activation_retains_last_known_valid_model(self):
        self.store.activate({"version": 1}, self._metadata(version="v1"))
        original_pointer = self.store.pointer_path.read_bytes()
        real_replace = os.replace

        def fail_pointer(source, destination):
            if Path(destination) == self.store.pointer_path:
                raise OSError("simulated pointer failure")
            return real_replace(source, destination)

        with patch("tok.ml_artifacts.os.replace", side_effect=fail_pointer):
            with self.assertRaises(ArtifactError):
                self.store.activate({"version": 2}, self._metadata(version="v2"))

        self.assertEqual(self.store.pointer_path.read_bytes(), original_pointer)
        loaded = self.store.load_active(expected_feature_version=FEATURE_VERSION)
        self.assertEqual(loaded.model, {"version": 1})

    def test_pointer_path_traversal_is_rejected(self):
        self.root.mkdir(exist_ok=True)
        self.store.pointer_path.write_text(json.dumps({
            "artifact_schema_version": ARTIFACT_SCHEMA_VERSION,
            "artifact_file": "../outside.pkl",
            "artifact_sha256": "x",
            "metadata": {},
        }), encoding="utf-8")

        with self.assertRaises(ArtifactCorrupt):
            self.store.load_active()

    def test_runtime_and_model_feature_incompatibility_are_rejected(self):
        incompatible = ArtifactMetadata(
            **{
                **self._metadata().__dict__,
                "python_version": "2.7.18",
            }
        )
        with self.assertRaises(ArtifactIncompatible):
            self.store.activate({"ok": True}, incompatible)

        with self.assertRaises(ArtifactCorrupt):
            self.store.activate(FeatureTaggedModel(), self._metadata())

    def test_ungated_or_unfitted_model_cannot_replace_active_pointer(self):
        self.store.activate({"version": 1}, self._metadata(version="v1"))
        original_pointer = self.store.pointer_path.read_bytes()
        ungated = ArtifactMetadata(**{
            **self._metadata(version="v2").__dict__,
            "production_enabled": False,
        })
        with self.assertRaises(ArtifactIncompatible):
            self.store.activate({"version": 2}, ungated)

        unfitted = LocalNgramClassifier(NgramModelConfig(
            name="unfitted",
            algorithm="centroid",
            dimensions=1024,
        ))
        with self.assertRaises(ArtifactCorrupt):
            self.store.activate(unfitted, self._metadata(version="v3"))
        self.assertEqual(self.store.pointer_path.read_bytes(), original_pointer)


if __name__ == "__main__":
    unittest.main()

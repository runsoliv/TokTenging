"""Versioned, atomic activation for local Tok Tenging ML artifacts."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from hashlib import sha256
from pathlib import Path
from typing import Any
import datetime as _datetime
import json
import os
import pickle
import re
import sys
import uuid

try:
    from .ml_features import FEATURE_VERSION
except ImportError:  # pragma: no cover - direct script compatibility
    from ml_features import FEATURE_VERSION


ARTIFACT_SCHEMA_VERSION = 1
COMPATIBILITY_VERSION = "tok-ml-runtime-v1"
ACTIVE_POINTER_FILENAME = "active_model.json"


class ArtifactError(RuntimeError):
    pass


class ArtifactUnavailable(ArtifactError):
    pass


class ArtifactCorrupt(ArtifactError):
    pass


class ArtifactIncompatible(ArtifactError):
    pass


@dataclass(frozen=True)
class ArtifactMetadata:
    model_version: str
    training_data_fingerprint: str
    training_row_count: int
    class_count: int
    feature_version: str
    evaluation_summary: dict[str, Any]
    production_enabled: bool
    created_at: str = ""
    compatibility_version: str = COMPATIBILITY_VERSION
    artifact_schema_version: int = ARTIFACT_SCHEMA_VERSION
    python_version: str = ""
    sklearn_version: str = ""
    pandas_version: str = ""

    def normalized(self) -> "ArtifactMetadata":
        if type(self.production_enabled) is not bool:
            raise ArtifactCorrupt("production_enabled must be a boolean")
        return ArtifactMetadata(
            model_version=str(self.model_version),
            training_data_fingerprint=str(self.training_data_fingerprint),
            training_row_count=int(self.training_row_count),
            class_count=int(self.class_count),
            feature_version=str(self.feature_version),
            evaluation_summary=dict(self.evaluation_summary or {}),
            production_enabled=self.production_enabled,
            created_at=self.created_at or _datetime.datetime.now(_datetime.timezone.utc).isoformat(timespec="seconds"),
            compatibility_version=str(self.compatibility_version),
            artifact_schema_version=int(self.artifact_schema_version),
            python_version=self.python_version or sys.version.split()[0],
            sklearn_version=str(self.sklearn_version or ""),
            pandas_version=str(self.pandas_version or ""),
        )

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ArtifactMetadata":
        if not isinstance(payload, dict):
            raise ArtifactCorrupt("artifact metadata is not an object")
        required = {
            "model_version",
            "training_data_fingerprint",
            "training_row_count",
            "class_count",
            "feature_version",
            "evaluation_summary",
            "production_enabled",
        }
        missing = sorted(required - set(payload))
        if missing:
            raise ArtifactCorrupt(f"artifact metadata missing: {', '.join(missing)}")
        allowed = set(cls.__dataclass_fields__)
        try:
            if type(payload.get("production_enabled")) is not bool:
                raise ValueError("production_enabled must be a boolean")
            return cls(**{key: payload[key] for key in allowed if key in payload}).normalized()
        except Exception as exc:
            raise ArtifactCorrupt(f"invalid artifact metadata: {exc}") from exc


@dataclass(frozen=True)
class LoadedArtifact:
    model: Any
    metadata: ArtifactMetadata
    artifact_path: Path
    checksum: str


def _sha256_file(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _fsync_file(handle) -> None:
    handle.flush()
    os.fsync(handle.fileno())


class ArtifactStore:
    """Store immutable model generations behind one atomically replaced pointer."""

    def __init__(self, root: str | os.PathLike, pointer_name: str = ACTIVE_POINTER_FILENAME):
        self.root = Path(root)
        self.pointer_path = self.root / Path(pointer_name).name

    def activate(self, model: Any, metadata: ArtifactMetadata) -> LoadedArtifact:
        metadata = metadata.normalized()
        self._validate_metadata(metadata, expected_feature_version=FEATURE_VERSION)
        self.root.mkdir(parents=True, exist_ok=True)
        generation = self._generation_name(metadata)
        artifact_name = f"model_{generation}.pkl"
        artifact_path = self.root / artifact_name
        artifact_temp = self.root / f".{artifact_name}.{uuid.uuid4().hex}.tmp"
        pointer_temp = self.root / f".{self.pointer_path.name}.{uuid.uuid4().hex}.tmp"
        payload = {
            "artifact_schema_version": ARTIFACT_SCHEMA_VERSION,
            "metadata": asdict(metadata),
            "model": model,
        }
        try:
            with artifact_temp.open("wb") as handle:
                pickle.dump(payload, handle, protocol=pickle.HIGHEST_PROTOCOL)
                _fsync_file(handle)
            # Validate serialization and compatibility before making the blob
            # visible or touching the active pointer.
            self._load_blob(artifact_temp, expected_feature_version=FEATURE_VERSION)
            os.replace(artifact_temp, artifact_path)
            checksum = _sha256_file(artifact_path)
            pointer = {
                "artifact_schema_version": ARTIFACT_SCHEMA_VERSION,
                "artifact_file": artifact_name,
                "artifact_sha256": checksum,
                "metadata": asdict(metadata),
            }
            with pointer_temp.open("w", encoding="utf-8") as handle:
                json.dump(pointer, handle, ensure_ascii=False, indent=2, sort_keys=True)
                _fsync_file(handle)
            os.replace(pointer_temp, self.pointer_path)
            return LoadedArtifact(model=model, metadata=metadata, artifact_path=artifact_path, checksum=checksum)
        except Exception as exc:
            for temp_path in (artifact_temp, pointer_temp):
                try:
                    temp_path.unlink(missing_ok=True)
                except OSError:
                    pass
            if isinstance(exc, ArtifactError):
                raise
            raise ArtifactError(f"artifact activation failed: {exc}") from exc

    def load_active(self, *, expected_feature_version: str | None = FEATURE_VERSION) -> LoadedArtifact:
        if not self.pointer_path.is_file():
            raise ArtifactUnavailable("active ML artifact pointer is missing")
        try:
            with self.pointer_path.open("r", encoding="utf-8") as handle:
                pointer = json.load(handle)
        except Exception as exc:
            raise ArtifactCorrupt(f"active ML pointer is unreadable: {exc}") from exc
        if not isinstance(pointer, dict):
            raise ArtifactCorrupt("active ML pointer is not an object")
        if pointer.get("artifact_schema_version") != ARTIFACT_SCHEMA_VERSION:
            raise ArtifactIncompatible("active ML pointer schema is incompatible")
        artifact_name = pointer.get("artifact_file")
        if not isinstance(artifact_name, str) or not artifact_name:
            raise ArtifactCorrupt("active ML pointer has no artifact filename")
        if Path(artifact_name).name != artifact_name:
            raise ArtifactCorrupt("active ML pointer contains an unsafe artifact path")
        artifact_path = self.root / artifact_name
        try:
            if artifact_path.resolve().parent != self.root.resolve():
                raise ArtifactCorrupt("active ML artifact escapes its store")
        except OSError as exc:
            raise ArtifactCorrupt(f"active ML artifact path is invalid: {exc}") from exc
        if not artifact_path.is_file():
            raise ArtifactUnavailable("active ML model file is missing")
        expected_checksum = pointer.get("artifact_sha256")
        actual_checksum = _sha256_file(artifact_path)
        if not isinstance(expected_checksum, str) or expected_checksum != actual_checksum:
            raise ArtifactCorrupt("active ML model checksum does not match")
        loaded = self._load_blob(artifact_path, expected_feature_version=expected_feature_version)
        pointer_metadata = ArtifactMetadata.from_dict(pointer.get("metadata"))
        if pointer_metadata != loaded.metadata:
            raise ArtifactCorrupt("pointer and model metadata do not match")
        return LoadedArtifact(
            model=loaded.model,
            metadata=loaded.metadata,
            artifact_path=artifact_path,
            checksum=actual_checksum,
        )

    def status(self, *, expected_feature_version: str | None = FEATURE_VERSION) -> dict[str, Any]:
        try:
            loaded = self.load_active(expected_feature_version=expected_feature_version)
        except ArtifactError as exc:
            return {
                "available": False,
                "reason": exc.__class__.__name__,
                "message": str(exc),
            }
        return {
            "available": True,
            "reason": "ok",
            "artifact_path": str(loaded.artifact_path),
            "checksum": loaded.checksum,
            "metadata": asdict(loaded.metadata),
        }

    def _load_blob(
        self,
        path: Path,
        *,
        expected_feature_version: str | None,
    ) -> LoadedArtifact:
        try:
            with path.open("rb") as handle:
                payload = pickle.load(handle)
        except Exception as exc:
            raise ArtifactCorrupt(f"ML model file is unreadable: {exc}") from exc
        if not isinstance(payload, dict):
            raise ArtifactCorrupt("ML model payload is not an object")
        if payload.get("artifact_schema_version") != ARTIFACT_SCHEMA_VERSION:
            raise ArtifactIncompatible("ML model schema is incompatible")
        if "model" not in payload:
            raise ArtifactCorrupt("ML model payload is missing its model")
        metadata = ArtifactMetadata.from_dict(payload.get("metadata"))
        self._validate_metadata(metadata, expected_feature_version=expected_feature_version)
        model = payload["model"]
        try:
            model_feature_version = getattr(model, "feature_version", metadata.feature_version)
            if str(model_feature_version) != metadata.feature_version:
                raise ArtifactCorrupt("ML model and metadata feature versions do not match")
            ready = getattr(model, "ready", None)
            if callable(ready):
                if not bool(ready()):
                    raise ArtifactCorrupt("ML model is not fitted and ready")
            classes = getattr(model, "classes_", None)
            if classes is not None and len(classes) != metadata.class_count:
                raise ArtifactCorrupt("ML model class count does not match metadata")
            support = getattr(model, "class_support_", None)
            if support is not None:
                support_total = int(sum(int(value) for value in support))
                if support_total != metadata.training_row_count:
                    raise ArtifactCorrupt("ML model training row count does not match metadata")
        except ArtifactError:
            raise
        except Exception as exc:
            raise ArtifactCorrupt(f"ML model logical validation failed: {exc}") from exc
        return LoadedArtifact(
            model=model,
            metadata=metadata,
            artifact_path=path,
            checksum=_sha256_file(path),
        )

    @staticmethod
    def _validate_metadata(
        metadata: ArtifactMetadata,
        *,
        expected_feature_version: str | None = None,
    ) -> None:
        if metadata.artifact_schema_version != ARTIFACT_SCHEMA_VERSION:
            raise ArtifactIncompatible("ML artifact schema version is incompatible")
        if metadata.compatibility_version != COMPATIBILITY_VERSION:
            raise ArtifactIncompatible("ML runtime compatibility version is incompatible")
        if expected_feature_version and metadata.feature_version != expected_feature_version:
            raise ArtifactIncompatible("ML feature schema version is incompatible")
        if metadata.training_row_count < 1:
            raise ArtifactCorrupt("ML artifact has no training rows")
        if metadata.class_count < 2:
            raise ArtifactCorrupt("ML artifact needs at least two classes")
        if metadata.production_enabled is not True:
            raise ArtifactIncompatible("ML artifact has not passed the production gate")
        if not metadata.model_version or not metadata.training_data_fingerprint:
            raise ArtifactCorrupt("ML artifact version/fingerprint is missing")
        if not ArtifactStore._runtime_version_compatible(metadata.python_version, sys.version.split()[0], parts=2):
            raise ArtifactIncompatible("ML artifact Python version is incompatible")
        if metadata.pandas_version:
            try:
                import pandas as pd
            except Exception as exc:
                raise ArtifactIncompatible("ML artifact requires pandas") from exc
            if not ArtifactStore._runtime_version_compatible(metadata.pandas_version, pd.__version__, parts=1):
                raise ArtifactIncompatible("ML artifact pandas version is incompatible")
        if metadata.sklearn_version:
            try:
                import sklearn
            except Exception as exc:
                raise ArtifactIncompatible("ML artifact requires scikit-learn") from exc
            if not ArtifactStore._runtime_version_compatible(metadata.sklearn_version, sklearn.__version__, parts=2):
                raise ArtifactIncompatible("ML artifact scikit-learn version is incompatible")

    @staticmethod
    def _runtime_version_compatible(expected: str, actual: str, *, parts: int) -> bool:
        def numeric_prefix(value: str) -> tuple[int, ...] | None:
            match = re.match(r"^(\d+)(?:\.(\d+))?(?:\.(\d+))?", str(value or ""))
            if not match:
                return None
            numbers = tuple(int(item) for item in match.groups(default="0"))
            return numbers[:parts]

        expected_prefix = numeric_prefix(expected)
        actual_prefix = numeric_prefix(actual)
        return bool(expected_prefix and actual_prefix and expected_prefix == actual_prefix)

    @staticmethod
    def _generation_name(metadata: ArtifactMetadata) -> str:
        version = re.sub(r"[^A-Za-z0-9_.-]", "_", metadata.model_version)[:32] or "model"
        fingerprint = re.sub(r"[^a-fA-F0-9]", "", metadata.training_data_fingerprint)[:16] or "data"
        return f"{version}_{fingerprint}_{uuid.uuid4().hex[:12]}"


__all__ = [
    "ACTIVE_POINTER_FILENAME",
    "ARTIFACT_SCHEMA_VERSION",
    "ArtifactCorrupt",
    "ArtifactError",
    "ArtifactIncompatible",
    "ArtifactMetadata",
    "ArtifactStore",
    "ArtifactUnavailable",
    "COMPATIBILITY_VERSION",
    "LoadedArtifact",
]

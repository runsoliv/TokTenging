"""Fail-closed runtime adapter for the optional local ML autocoder.

The historical autocoder always runs first.  This module may replace only
unresolved/review targets, and only with an atomically activated artifact whose
metadata and confidence policy both passed the production gate.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any
import json
import logging

import pandas as pd

try:
    from .auto_coder import (
        _hardcoded_bank_signal_rule,
        _source_row_for_output,
        _source_rule_text,
        clean_code,
        normalize_text,
    )
    from .ml_artifacts import ArtifactError, ArtifactStore
    from .ml_dataset import (
        NON_CLASSIFICATION_ACCOUNT_CODES,
        VEHICLE_EXPENSE_FAMILY,
        build_runtime_feature_record,
    )
    from .ml_evaluation import ConfidencePolicy, prediction_is_accepted
    from .ml_features import FEATURE_VERSION
except ImportError:  # pragma: no cover - direct script compatibility
    from auto_coder import (
        _hardcoded_bank_signal_rule,
        _source_row_for_output,
        _source_rule_text,
        clean_code,
        normalize_text,
    )
    from ml_artifacts import ArtifactError, ArtifactStore
    from ml_dataset import (
        NON_CLASSIFICATION_ACCOUNT_CODES,
        VEHICLE_EXPENSE_FAMILY,
        build_runtime_feature_record,
    )
    from ml_evaluation import ConfidencePolicy, prediction_is_accepted
    from ml_features import FEATURE_VERSION


LOGGER = logging.getLogger(__name__)
REVIEW_STATUSES = frozenset({"", "review", "review needed", "fallback review", "fallback_review"})
PROTECTED_SOURCES = frozenset({
    "existing",
    "manual_review",
    "hardcoded_bank_signal",
    "hardcoded_collection_cost",
    "client_exact_consistent",
    "global_exact_consistent",
    "bank_signal_rule",
    "amount_rule",
    "restaurant_context",
    "general_context",
    "keyword",
})
MANUAL_OR_DETERMINISTIC_MARKERS = frozenset({"AD", "MN"})
VEHICLE_CODES = frozenset({"4635", "4640"})


def _default_artifact_dir() -> Path:
    return Path(__file__).resolve().parent / "trainingData" / "ml_artifacts"


def _confidence_policy(evaluation_summary: Any) -> ConfidencePolicy:
    if not isinstance(evaluation_summary, dict):
        raise ValueError("artifact evaluation summary is missing")
    if evaluation_summary.get("gate_passed") is not True:
        raise ValueError("artifact evaluation did not pass every production gate")
    gate_checks = evaluation_summary.get("gate_checks")
    if gate_checks is not None:
        if not isinstance(gate_checks, dict) or not gate_checks:
            raise ValueError("artifact production gate checks are invalid")
        if any(
            not isinstance(check, dict)
            or check.get("meets_95_percent_ml_precision") is not True
            or check.get("no_accuracy_regression") is not True
            for check in gate_checks.values()
        ):
            raise ValueError("artifact failed a grouped production gate")
    payload = evaluation_summary.get("confidence_policy")
    if payload is None:
        payload = evaluation_summary.get("policy")
    if payload is None:
        payload = evaluation_summary
    policy = ConfidencePolicy.from_dict(payload)
    if policy.production_enabled is not True:
        raise ValueError("artifact confidence policy has not passed the production gate")
    if (
        policy.validation_precision < policy.target_precision
        or policy.validation_accepted < 1
    ):
        raise ValueError("artifact confidence evidence is below its production target")
    return policy


def _marker(row: Any) -> str:
    getter = row.get if hasattr(row, "get") else lambda _key, default="": default
    value = getter("CONFIDENCE", "")
    if value in (None, ""):
        value = getter("Cnf", "")
    return str(value or "").strip().upper()


def _target_column(row: Any, record: dict[str, Any]) -> str:
    getter = row.get if hasattr(row, "get") else lambda _key, default="": default
    source = str(getter("_AUTO_CODE_SOURCE", "") or "").strip().lower()
    status = str(getter("STATUS", "") or "").strip().lower()
    if _marker(row) in MANUAL_OR_DETERMINISTIC_MARKERS:
        return ""
    if source in PROTECTED_SOURCES or source.startswith("hardcoded_"):
        return ""
    if status not in REVIEW_STATUSES:
        return ""

    direction = str(record.get("direction", "")).strip().lower()
    target = "DEBIT" if direction == "outgoing" else "CREDIT" if direction == "incoming" else ""
    if not target:
        return ""
    existing = clean_code(getter(target, ""))
    # A nonblank value with no autocoder provenance is manual/external work.
    if existing and not source:
        return ""
    return target


def _unknown_treasury(text: Any, source_row: Any) -> bool:
    rule_text = _source_rule_text(source_row)
    combined = " ".join((normalize_text(text), normalize_text(rule_text))).strip()
    return (
        "rikissjodsinnheimt" in combined
        and _hardcoded_bank_signal_rule(str(text or ""), rule_text) is None
    )


def _merchant_familiar(model: Any, record: dict[str, Any]) -> bool:
    family = normalize_text(record.get("merchant_family", ""))
    known = getattr(model, "merchant_families_", None)
    if not family or known is None:
        # Absence of an explicit training-set signal is treated as unseen.  This
        # intentionally selects the stricter held-out threshold.
        return False
    try:
        return family in {normalize_text(value) for value in known}
    except Exception:
        return False


def _prediction_payload(prediction: Any) -> dict[str, Any]:
    if hasattr(prediction, "to_dict"):
        payload = prediction.to_dict()
    elif isinstance(prediction, dict):
        payload = dict(prediction)
    else:
        payload = {
            key: getattr(prediction, key)
            for key in ("predicted", "probability", "margin", "candidates", "class_support")
        }
    if not isinstance(payload, dict):
        raise ValueError("ML prediction is not an object")
    return payload


def _resolved_code(predicted: Any, vehicle_expense_code: Any) -> str:
    value = str(predicted or "").strip()
    if value == VEHICLE_EXPENSE_FAMILY:
        selected = clean_code(vehicle_expense_code)
        return selected if selected in VEHICLE_CODES else ""
    code = clean_code(value)
    if not code or code in NON_CLASSIFICATION_ACCOUNT_CODES:
        return ""
    return code


def apply_ml_hybrid_codes(
    legacy_df: pd.DataFrame,
    *,
    source_df: pd.DataFrame | None = None,
    input_file_path: str = "",
    bank_type: str = "",
    vehicle_expense_code: str = "",
    artifact_dir: str | Path | None = None,
) -> pd.DataFrame:
    """Apply accepted ML predictions without ever weakening legacy output.

    Missing/corrupt/incompatible artifacts, invalid policies, inference errors,
    or malformed predictions return a deep copy of ``legacy_df`` unchanged.
    ``input_file_path`` is accepted for API symmetry but deliberately excluded
    from features so client identity cannot become a shortcut.
    """
    del input_file_path
    unchanged = legacy_df.copy(deep=True)
    try:
        store = ArtifactStore(artifact_dir or _default_artifact_dir())
        loaded = store.load_active(expected_feature_version=FEATURE_VERSION)
        policy = _confidence_policy(loaded.metadata.evaluation_summary)

        candidate_indices: list[Any] = []
        candidate_records: list[dict[str, Any]] = []
        target_columns: list[str] = []
        treasury_flags: list[bool] = []
        for index, row in unchanged.iterrows():
            source_row = _source_row_for_output(source_df, row, index)
            record = build_runtime_feature_record(
                row,
                source_row=source_row,
                bank_type=bank_type,
            )
            target = _target_column(row, record)
            if not target:
                continue
            candidate_indices.append(index)
            candidate_records.append(record)
            target_columns.append(target)
            treasury_flags.append(_unknown_treasury(row.get("TEXT", ""), source_row))

        if not candidate_records:
            return unchanged
        predictions = loaded.model.predict_records(candidate_records, top_k=3)
        if len(predictions) != len(candidate_records):
            raise ValueError("ML model returned the wrong number of predictions")

        updates: list[dict[str, Any]] = []
        for index, target, record, unknown_treasury, prediction in zip(
            candidate_indices,
            target_columns,
            candidate_records,
            treasury_flags,
            predictions,
        ):
            payload = _prediction_payload(prediction)
            payload["merchant_familiar"] = _merchant_familiar(loaded.model, record)
            payload["weak_text"] = bool(record.get("weak_text", False))
            payload["unknown_treasury"] = bool(unknown_treasury)
            accepted, _reason = prediction_is_accepted(payload, policy)
            code = _resolved_code(payload.get("predicted"), vehicle_expense_code)
            if not accepted or not code:
                continue
            probability = float(payload.get("probability", 0.0))
            confidence = max(1, min(99, int(round(probability * 100))))
            updates.append({
                "index": index,
                "target": target,
                "code": code,
                "confidence": confidence,
                "candidates": payload.get("candidates", ()),
            })

        result = unchanged.copy(deep=True)
        confidence_column = "CONFIDENCE" if "CONFIDENCE" in result.columns else "Cnf" if "Cnf" in result.columns else "CONFIDENCE"
        for update in updates:
            index = update["index"]
            result.at[index, update["target"]] = update["code"]
            result.at[index, "STATUS"] = "coded"
            result.at[index, confidence_column] = update["confidence"]
            result.at[index, "_AUTO_CODE_SOURCE"] = "ml"
            result.at[index, "_AUTO_CODE_ML_VERSION"] = loaded.metadata.model_version
            result.at[index, "_AUTO_CODE_ML_CANDIDATES"] = json.dumps(
                update["candidates"],
                ensure_ascii=False,
                separators=(",", ":"),
            )
        return result
    except (ArtifactError, OSError, ValueError, TypeError, AttributeError) as exc:
        LOGGER.info("ML autocoding unavailable; legacy result retained: %s", exc)
        return unchanged
    except Exception as exc:
        LOGGER.warning("ML autocoding failed; legacy result retained: %s", exc)
        return unchanged


def ml_runtime_status(artifact_dir: str | Path | None = None) -> dict[str, Any]:
    """Return a non-throwing health summary for the settings UI."""
    try:
        loaded = ArtifactStore(artifact_dir or _default_artifact_dir()).load_active(
            expected_feature_version=FEATURE_VERSION
        )
        policy = _confidence_policy(loaded.metadata.evaluation_summary)
        return {
            "available": True,
            "reason": "ok",
            "model_version": loaded.metadata.model_version,
            "training_rows": loaded.metadata.training_row_count,
            "validation_precision": policy.validation_precision,
            "validation_coverage": policy.validation_coverage,
        }
    except Exception as exc:
        return {
            "available": False,
            "reason": exc.__class__.__name__,
            "message": str(exc),
        }


__all__ = ["apply_ml_hybrid_codes", "ml_runtime_status"]

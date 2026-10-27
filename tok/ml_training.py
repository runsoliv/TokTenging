"""Reproducible grouped evaluation and training entry point for Tok Tenging ML.

Run Gate 1 without changing any production artifact::

    python -m tok.ml_training evaluate --training-data-dir tok/trainingData

The command reads only manifest-designated workbooks.  It writes an atomic JSON
report to ``artifacts/ml/evaluation_report.json`` by default and never modifies
the source workbooks or the legacy autocoder memory.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import asdict
from pathlib import Path
from typing import Any, Iterable, Sequence
import argparse
import datetime as _datetime
import json
import math
import os
import re
import tempfile

import numpy as np
import pandas as pd

try:
    from .auto_coder import (
        RAW_DETERMINISTIC_CONFIDENCE_SOURCES,
        TransactionAutoCoder,
        _cap_unrecognized_treasury_confidence,
        _context_features,
        _hardcoded_bank_signal_rule,
        _is_explicit_collection_cost_signal,
        _is_generic_training_text,
        clean_code,
        client_key_from_path,
        merchant_family,
        merchant_key,
        normalize_text,
        tokenize,
    )
    from .ml_dataset import TrainingDataset, load_training_dataset
    from .ml_artifacts import ArtifactMetadata, ArtifactStore
    from .ml_evaluation import (
        ConfidencePolicy,
        audit_fold_leakage,
        calibration_summary,
        classification_metrics,
        default_validation_schemes,
        make_grouped_folds,
        prediction_is_accepted,
        select_confidence_policy,
        top_k_accuracy,
    )
    from .ml_features import FEATURE_VERSION, feature_documents, model_target_code
    from .ml_model import (
        DEFAULT_CANDIDATE_CONFIGS,
        LocalModelEnsemble,
        LocalNgramClassifier,
        NgramModelConfig,
    )
except ImportError:  # pragma: no cover - direct script compatibility
    from auto_coder import (
        RAW_DETERMINISTIC_CONFIDENCE_SOURCES,
        TransactionAutoCoder,
        _cap_unrecognized_treasury_confidence,
        _context_features,
        _hardcoded_bank_signal_rule,
        _is_explicit_collection_cost_signal,
        _is_generic_training_text,
        clean_code,
        client_key_from_path,
        merchant_family,
        merchant_key,
        normalize_text,
        tokenize,
    )
    from ml_dataset import TrainingDataset, load_training_dataset
    from ml_artifacts import ArtifactMetadata, ArtifactStore
    from ml_evaluation import (
        ConfidencePolicy,
        audit_fold_leakage,
        calibration_summary,
        classification_metrics,
        default_validation_schemes,
        make_grouped_folds,
        prediction_is_accepted,
        select_confidence_policy,
        top_k_accuracy,
    )
    from ml_features import FEATURE_VERSION, feature_documents, model_target_code
    from ml_model import DEFAULT_CANDIDATE_CONFIGS, LocalModelEnsemble, LocalNgramClassifier, NgramModelConfig


EVALUATION_SCHEMA_VERSION = "tok-ml-evaluation-v1"
AUTHORITATIVE_LEGACY_SOURCES = frozenset({
    *RAW_DETERMINISTIC_CONFIDENCE_SOURCES,
    "client_exact_consistent",
    "global_exact_consistent",
})
UNSAFE_LEGACY_REVIEW_SOURCES = frozenset({
    "",
    "fallback",
    "person_default",
    "unavailable",
})


def _json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): _json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(item) for item in value]
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating,)):
        return float(value) if math.isfinite(float(value)) else None
    if isinstance(value, (np.bool_,)):
        return bool(value)
    if isinstance(value, Path):
        return str(value)
    return value


def _atomic_json_write(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temp_name = tempfile.mkstemp(
        prefix=f".{path.name}.",
        suffix=".tmp",
        dir=path.parent,
    )
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
            json.dump(_json_safe(payload), handle, ensure_ascii=False, indent=2, sort_keys=True)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_name, path)
    except Exception:
        try:
            os.unlink(temp_name)
        except OSError:
            pass
        raise


def _empty_legacy_coder() -> TransactionAutoCoder:
    coder = TransactionAutoCoder.__new__(TransactionAutoCoder)
    coder.global_exact = defaultdict(Counter)
    coder.global_family = defaultdict(Counter)
    coder.client_exact = defaultdict(Counter)
    coder.client_family = defaultdict(Counter)
    coder.context_counts = defaultdict(Counter)
    coder.client_context_counts = defaultdict(Counter)
    coder.token_counts = defaultdict(Counter)
    coder.bigram_counts = defaultdict(Counter)
    coder.code_counts = Counter()
    coder.account_descriptions = {}
    coder.training_rows = 0
    coder.training_files = 0
    coder.source_fingerprint = {}
    coder.loaded_from_memory = False
    coder.memory_saved = False
    return coder


def fit_legacy_evaluation_coder(frame: pd.DataFrame) -> TransactionAutoCoder:
    """Fit the current rule/history implementation in memory for one fold."""
    coder = _empty_legacy_coder()
    records = []
    for _, row in frame.iterrows():
        if str(row.get("direction", "")) != "outgoing":
            continue
        target = clean_code(row.get("target_code"))
        text = row.get("text", "")
        if not target or _is_generic_training_text(text):
            continue
        key = merchant_key(text)
        family = merchant_family(text)
        if not key:
            continue
        client = client_key_from_path(row.get("client_id", ""))
        records.append((row, target, text, key, family, client))

    key_counts = Counter(item[3] for item in records)
    family_counts = Counter(item[4] for item in records if item[4])
    for row, target, text, key, family, client in records:
        coder.training_rows += 1
        coder.code_counts[target] += 1
        coder.global_exact[key][target] += 1
        if family:
            coder.global_family[family][target] += 1
        if client:
            coder.client_exact[(client, key)][target] += 1
            if family:
                coder.client_family[(client, family)][target] += 1
        tokens = tokenize(text, drop_processors=True)
        for token in set(tokens):
            coder.token_counts[token][target] += 1
        for bigram in set(" ".join(tokens[index:index + 2]) for index in range(len(tokens) - 1)):
            coder.bigram_counts[bigram][target] += 1
        repeat_count = key_counts.get(key) or family_counts.get(family) or 1
        for feature in _context_features(
            text,
            amount=row.get("amount"),
            date=row.get("date"),
            repeat_count=repeat_count,
            id_value=row.get("valid_kennitala"),
        ):
            coder.context_counts[feature][target] += 1
            if client:
                coder.client_context_counts[(client, feature)][target] += 1
    return coder


def _legacy_prediction(coder: TransactionAutoCoder, row: pd.Series) -> dict[str, Any]:
    if str(row.get("direction", "")) != "outgoing":
        return {
            "predicted": "",
            "confidence": 0,
            "source": "incoming_unhandled",
            "status": "review",
            "accepted": False,
        }
    raw_text = str(row.get("raw_explanation", "") or "")
    context_flags = {"explicit_collection_cost"} if _is_explicit_collection_cost_signal(raw_text) else set()
    result = coder.predict(
        row.get("text", ""),
        signal_text=raw_text,
        amount=row.get("amount"),
        date=row.get("date"),
        id_value=row.get("valid_kennitala"),
        repeat_count=1,
        client_hint=row.get("client_id", ""),
        context_flags=context_flags,
    )
    result = _cap_unrecognized_treasury_confidence(
        result,
        str(row.get("text", "") or ""),
        raw_text,
    )
    predicted = model_target_code(result.code)
    accepted = bool(
        predicted
        and (
            result.source in AUTHORITATIVE_LEGACY_SOURCES
            or (
                result.source not in UNSAFE_LEGACY_REVIEW_SOURCES
                and result.status == "coded"
                and int(result.confidence or 0) >= 55
            )
        )
    )
    return {
        "predicted": predicted,
        "confidence": int(result.confidence or 0),
        "source": str(result.source or ""),
        "status": str(result.status or ""),
        "accepted": accepted,
    }


def _deterministic_prediction(row: pd.Series) -> dict[str, Any]:
    text = str(row.get("text", "") or "")
    raw = str(row.get("raw_explanation", "") or "")
    result = _hardcoded_bank_signal_rule(text, raw)
    if result is None and _is_explicit_collection_cost_signal(raw):
        # The production row-context rule assigns the same authoritative code;
        # this row-level evidence is sufficient for held-out evaluation.
        return {"predicted": "6220", "accepted": True, "source": "hardcoded_collection_cost"}
    if result is None:
        return {"predicted": "", "accepted": False, "source": "no_deterministic_rule"}
    return {
        "predicted": model_target_code(result.code),
        "accepted": True,
        "source": result.source,
    }


def _exact_history_predictions(train: pd.DataFrame, test: pd.DataFrame) -> list[dict[str, Any]]:
    counters: dict[str, Counter] = defaultdict(Counter)
    for _, row in train.iterrows():
        key = merchant_key(row.get("text", ""))
        if key:
            counters[key][str(row.get("model_target", ""))] += 1
    output = []
    for _, row in test.iterrows():
        key = merchant_key(row.get("text", ""))
        counter = counters.get(key)
        if not counter:
            output.append({"predicted": "", "accepted": False, "source": "no_exact_history"})
            continue
        predicted, count = counter.most_common(1)[0]
        total = sum(counter.values())
        ratio = count / total if total else 0.0
        consistent = total >= 2 and (ratio >= 0.98 or (total >= 5 and ratio >= 0.94))
        output.append({
            "predicted": predicted if consistent else "",
            "accepted": bool(consistent),
            "source": "exact_consistent" if consistent else "exact_conflict",
        })
    return output


def _merchant_signals(row: pd.Series) -> set[str]:
    signals = set()
    for column, prefix in (
        ("merchant_family", "family"),
        ("valid_kennitala", "kt"),
        ("normalized_text", "text"),
    ):
        value = str(row.get(column, "") or "").strip()
        if value:
            signals.add(f"{prefix}:{value}")
    return signals


def _prediction_context(train: pd.DataFrame, test: pd.DataFrame) -> list[dict[str, Any]]:
    familiar_signals = set()
    for _, row in train.iterrows():
        familiar_signals.update(_merchant_signals(row))
    output = []
    for _, row in test.iterrows():
        signals = _merchant_signals(row)
        text = " ".join((str(row.get("text", "")), str(row.get("raw_explanation", ""))))
        normalized = normalize_text(text)
        unknown_treasury = (
            "rikissjodsinnheimt" in normalized
            and _hardcoded_bank_signal_rule(row.get("text", ""), row.get("raw_explanation", "")) is None
        )
        output.append({
            "merchant_familiar": bool(signals & familiar_signals),
            "weak_text": bool(row.get("weak_text", False)),
            "unknown_treasury": bool(unknown_treasury),
        })
    return output


def _calibrated_fold_model(
    config: NgramModelConfig,
    train: pd.DataFrame,
) -> LocalNgramClassifier:
    train = train.reset_index(drop=True)
    temperature = 1.0
    if config.algorithm == "similarity":
        return LocalNgramClassifier(config).fit(
            feature_documents(train),
            train["model_target"].astype(str).tolist(),
            temperature=temperature,
        )
    try:
        inner_folds = make_grouped_folds(
            train,
            name="inner_calibration",
            group_columns=("session_id",),
            n_splits=min(4, max(2, train["session_id"].nunique())),
        )
        calibration_fold = inner_folds[0]
        fit_frame = train.iloc[list(calibration_fold.train_indices)]
        calibration_frame = train.iloc[list(calibration_fold.test_indices)]
        inner_model = LocalNgramClassifier(config).fit(
            feature_documents(fit_frame),
            fit_frame["model_target"].astype(str).tolist(),
        )
        temperature = inner_model.calibrate_temperature(
            feature_documents(calibration_frame),
            calibration_frame["model_target"].astype(str).tolist(),
        )
    except (ValueError, RuntimeError):
        temperature = 1.0
    return LocalNgramClassifier(config).fit(
        feature_documents(train),
        train["model_target"].astype(str).tolist(),
        temperature=temperature,
    )


def evaluate_ml_config(
    frame: pd.DataFrame,
    folds,
    config: NgramModelConfig,
) -> pd.DataFrame:
    frame = frame.reset_index(drop=True)
    records = []
    for fold in folds:
        leakage = audit_fold_leakage(frame, fold)
        if not leakage["ok"]:
            raise RuntimeError(f"group leakage in {fold.name}/{fold.fold}: {leakage}")
        train = frame.iloc[list(fold.train_indices)].reset_index(drop=True)
        test = frame.iloc[list(fold.test_indices)].reset_index(drop=True)
        model = _calibrated_fold_model(config, train)
        predictions = model.predict(feature_documents(test), top_k=3)
        contexts = _prediction_context(train, test)
        for position, (prediction, context) in enumerate(zip(predictions, contexts)):
            row = test.iloc[position]
            records.append({
                "example_id": row["example_id"],
                "fold": int(fold.fold),
                "scheme": fold.name,
                "actual": str(row["model_target"]),
                "predicted": prediction.predicted,
                "probability": float(prediction.probability),
                "margin": float(prediction.margin),
                "candidates": list(prediction.candidates),
                "ranked": [candidate[0] for candidate in prediction.candidates],
                "class_support": int(prediction.class_support),
                "correct": prediction.predicted == str(row["model_target"]),
                "bank_type": str(row["bank_type"]),
                "direction": str(row["direction"]),
                "merchant_family": str(row.get("merchant_family", "")),
                "raw_pair_status": str(row.get("raw_pair_status", "")),
                **context,
            })
    return pd.DataFrame(records).sort_values("example_id", kind="stable").reset_index(drop=True)


def ensemble_prediction_frames(
    prediction_frames: Sequence[pd.DataFrame],
    *,
    name: str = "unanimous_ensemble",
) -> pd.DataFrame:
    if len(prediction_frames) < 2:
        raise ValueError("at least two prediction frames are required")
    indexed = [frame.set_index("example_id", drop=False) for frame in prediction_frames]
    expected_ids = set(indexed[0].index)
    if any(set(frame.index) != expected_ids for frame in indexed[1:]):
        raise ValueError("ensemble prediction frames do not contain identical examples")
    records = []
    for example_id in sorted(expected_ids):
        rows = [frame.loc[example_id] for frame in indexed]
        votes = Counter(str(row["predicted"]) for row in rows)
        predicted, vote_count = votes.most_common(1)[0]
        if vote_count == len(rows):
            probability, margin = 0.999, 0.90
        elif vote_count >= math.ceil(len(rows) / 2):
            probability, margin = 0.85, 0.40
        else:
            strongest = max(rows, key=lambda row: float(row["probability"]))
            predicted = str(strongest["predicted"])
            probability = min(0.49, float(strongest["probability"]))
            margin = min(0.10, float(strongest["margin"]))
        first = rows[0]
        ranked_votes = Counter()
        for row in rows:
            for rank, code in enumerate(row["ranked"]):
                ranked_votes[str(code)] += max(1, 3 - rank)
        ranked = [predicted] + [
            code for code, _score in ranked_votes.most_common()
            if code != predicted
        ]
        actual = str(first["actual"])
        records.append({
            "example_id": example_id,
            "fold": int(first["fold"]),
            "scheme": str(first["scheme"]),
            "actual": actual,
            "predicted": predicted,
            "probability": probability,
            "margin": margin,
            "candidates": [(code, probability if index == 0 else 0.0) for index, code in enumerate(ranked[:3])],
            "ranked": ranked[:3],
            "class_support": min(int(row["class_support"]) for row in rows if str(row["predicted"]) == predicted),
            "correct": predicted == actual,
            "bank_type": str(first["bank_type"]),
            "direction": str(first["direction"]),
            "merchant_family": str(first["merchant_family"]),
            "raw_pair_status": str(first["raw_pair_status"]),
            "merchant_familiar": bool(first["merchant_familiar"]),
            "weak_text": bool(first["weak_text"]),
            "unknown_treasury": bool(first["unknown_treasury"]),
            "ensemble_name": name,
            "agreement_count": int(vote_count),
        })
    return pd.DataFrame(records)


def evaluate_baselines(frame: pd.DataFrame, folds) -> dict[str, pd.DataFrame]:
    frame = frame.reset_index(drop=True)
    output = {name: [] for name in ("deterministic_only", "exact_history_only", "legacy_complete")}
    for fold in folds:
        train = frame.iloc[list(fold.train_indices)].reset_index(drop=True)
        test = frame.iloc[list(fold.test_indices)].reset_index(drop=True)
        legacy = fit_legacy_evaluation_coder(train)
        exact = _exact_history_predictions(train, test)
        deterministic = [_deterministic_prediction(row) for _, row in test.iterrows()]
        legacy_predictions = [_legacy_prediction(legacy, row) for _, row in test.iterrows()]
        for name, predictions in (
            ("deterministic_only", deterministic),
            ("exact_history_only", exact),
            ("legacy_complete", legacy_predictions),
        ):
            for position, prediction in enumerate(predictions):
                row = test.iloc[position]
                predicted = str(prediction.get("predicted", "") or "")
                output[name].append({
                    "example_id": row["example_id"],
                    "fold": int(fold.fold),
                    "scheme": fold.name,
                    "actual": str(row["model_target"]),
                    "predicted": predicted,
                    "correct": predicted == str(row["model_target"]),
                    "accepted": bool(prediction.get("accepted", False)),
                    "source": str(prediction.get("source", "")),
                    "confidence": int(prediction.get("confidence", 0) or 0),
                    "bank_type": str(row["bank_type"]),
                    "direction": str(row["direction"]),
                })
    return {
        name: pd.DataFrame(records).sort_values("example_id", kind="stable").reset_index(drop=True)
        for name, records in output.items()
    }


def _coverage_table(predictions: pd.DataFrame) -> list[dict[str, Any]]:
    rows = []
    for threshold in (0.50, 0.60, 0.70, 0.80, 0.85, 0.90, 0.95, 0.97, 0.99):
        accepted = predictions["probability"] >= threshold
        count = int(accepted.sum())
        rows.append({
            "threshold": threshold,
            "accepted": count,
            "coverage": count / len(predictions) if len(predictions) else 0.0,
            "precision": float(predictions.loc[accepted, "correct"].mean()) if count else 0.0,
        })
    return rows


def _important_confusions(predictions: pd.DataFrame, limit: int = 30) -> list[dict[str, Any]]:
    errors = predictions[predictions["actual"] != predictions["predicted"]]
    counts = (
        errors.groupby(["actual", "predicted"], dropna=False)
        .size()
        .sort_values(ascending=False)
    )
    return [
        {"actual": str(actual), "predicted": str(predicted), "count": int(count)}
        for (actual, predicted), count in counts.head(limit).items()
    ]


def _segment_metrics(predictions: pd.DataFrame, column: str) -> dict[str, Any]:
    output = {}
    for value, subset in predictions.groupby(column, dropna=False):
        output[str(value)] = classification_metrics(
            subset["actual"].astype(str).tolist(),
            subset["predicted"].astype(str).tolist(),
        )
    return output


def summarize_predictions(
    predictions: pd.DataFrame,
    *,
    include_probability: bool,
) -> dict[str, Any]:
    metrics = classification_metrics(
        predictions["actual"].astype(str).tolist(),
        predictions["predicted"].astype(str).tolist(),
    )
    summary: dict[str, Any] = {
        **metrics,
        "by_bank": _segment_metrics(predictions, "bank_type"),
        "by_direction": _segment_metrics(predictions, "direction"),
        "important_confusions": _important_confusions(predictions),
    }
    if "accepted" in predictions.columns:
        accepted = predictions["accepted"].astype(bool)
        count = int(accepted.sum())
        summary["automatic_acceptance"] = {
            "accepted": count,
            "coverage": count / len(predictions) if len(predictions) else 0.0,
            "precision": float(predictions.loc[accepted, "correct"].mean()) if count else 0.0,
            "review_rate": 1.0 - (count / len(predictions) if len(predictions) else 0.0),
        }
    if include_probability:
        summary["top_2_accuracy"] = top_k_accuracy(
            predictions["actual"].astype(str).tolist(),
            predictions["ranked"].tolist(),
            2,
        )
        summary["top_3_accuracy"] = top_k_accuracy(
            predictions["actual"].astype(str).tolist(),
            predictions["ranked"].tolist(),
            3,
        )
        summary["calibration"] = calibration_summary(
            predictions["probability"].astype(float).tolist(),
            predictions["correct"].astype(bool).tolist(),
        )
        summary["coverage_by_probability"] = _coverage_table(predictions)
        for segment_name, mask in (
            ("known_merchant", predictions["merchant_familiar"].astype(bool)),
            ("unknown_merchant", ~predictions["merchant_familiar"].astype(bool)),
            ("weak_text", predictions["weak_text"].astype(bool)),
        ):
            subset = predictions.loc[mask]
            summary[segment_name] = classification_metrics(
                subset["actual"].astype(str).tolist(),
                subset["predicted"].astype(str).tolist(),
            )
    return summary


def apply_policy(predictions: pd.DataFrame, policy: ConfidencePolicy) -> pd.DataFrame:
    output = predictions.copy()
    decisions = [prediction_is_accepted(row, policy) for _, row in output.iterrows()]
    output["accepted"] = [decision[0] for decision in decisions]
    output["acceptance_reason"] = [decision[1] for decision in decisions]
    return output


def unresolved_ml_predictions(ml: pd.DataFrame, legacy: pd.DataFrame) -> pd.DataFrame:
    """Restrict policy fitting/evaluation to rows ML can affect in production."""
    legacy_flags = legacy.set_index("example_id")["accepted"].astype(bool)
    eligible_ids = [
        example_id
        for example_id in ml["example_id"].astype(str)
        if not bool(legacy_flags.get(example_id, False))
    ]
    return ml.loc[ml["example_id"].astype(str).isin(eligible_ids)].copy().reset_index(drop=True)


def hybrid_predictions(
    frame: pd.DataFrame,
    legacy: pd.DataFrame,
    ml: pd.DataFrame,
    policy: ConfidencePolicy,
) -> pd.DataFrame:
    frame_by_id = frame.set_index("example_id", drop=False)
    legacy_by_id = legacy.set_index("example_id", drop=False)
    ml_by_id = apply_policy(ml, policy).set_index("example_id", drop=False)
    records = []
    for example_id in sorted(frame_by_id.index):
        row = frame_by_id.loc[example_id]
        old = legacy_by_id.loc[example_id]
        learned = ml_by_id.loc[example_id]
        if bool(old["accepted"]):
            predicted = str(old["predicted"])
            source = f"legacy:{old['source']}"
            accepted = True
        elif bool(learned["accepted"]):
            predicted = str(learned["predicted"])
            source = "ml"
            accepted = True
        else:
            # Production retains the old fallback value for compatibility,
            # while the acceptance flag correctly records required review.
            predicted = str(old["predicted"])
            source = f"review:{learned['acceptance_reason']}"
            accepted = False
        records.append({
            "example_id": example_id,
            "actual": str(row["model_target"]),
            "predicted": predicted,
            "correct": predicted == str(row["model_target"]),
            "accepted": accepted,
            "source": source,
            "bank_type": str(row["bank_type"]),
            "direction": str(row["direction"]),
        })
    return pd.DataFrame(records)


def run_evaluation(
    *,
    training_data_dir: str | Path,
    legacy_training_dir: str | Path | None = None,
    include_legacy: bool = False,
    n_splits: int = 3,
) -> tuple[dict[str, Any], Any, ConfidencePolicy]:
    dataset: TrainingDataset = load_training_dataset(
        training_data_dir=training_data_dir,
        legacy_training_dir=legacy_training_dir,
        include_legacy=include_legacy,
    )
    frame = dataset.model_frame().reset_index(drop=True)
    if len(frame) < n_splits:
        raise RuntimeError("not enough production-eligible rows for grouped evaluation")
    schemes = default_validation_schemes(frame, n_splits=n_splits)
    leakage_audit = {
        name: [audit_fold_leakage(frame, fold) for fold in folds]
        for name, folds in schemes.items()
    }
    if not all(item["ok"] for audits in leakage_audit.values() for item in audits):
        raise RuntimeError(f"group leakage audit failed: {leakage_audit}")

    primary_folds = schemes["unseen_session"]
    baselines = evaluate_baselines(frame, primary_folds)
    candidate_predictions = {}
    candidate_summaries = {}
    for config in DEFAULT_CANDIDATE_CONFIGS:
        predictions = evaluate_ml_config(frame, primary_folds, config)
        candidate_predictions[config.name] = predictions
        candidate_summaries[config.name] = summarize_predictions(
            predictions,
            include_probability=True,
        )
    ensemble_name = "unanimous_ensemble"
    candidate_predictions[ensemble_name] = ensemble_prediction_frames(
        [candidate_predictions[config.name] for config in DEFAULT_CANDIDATE_CONFIGS],
        name=ensemble_name,
    )
    candidate_summaries[ensemble_name] = summarize_predictions(
        candidate_predictions[ensemble_name],
        include_probability=True,
    )
    candidate_policies = {}
    candidate_incremental = {}
    candidate_names = [config.name for config in DEFAULT_CANDIDATE_CONFIGS] + [ensemble_name]
    for candidate_name in candidate_names:
        unresolved = unresolved_ml_predictions(
            candidate_predictions[candidate_name],
            baselines["legacy_complete"],
        )
        candidate_policy = select_confidence_policy(unresolved, target_precision=0.95)
        candidate_policies[candidate_name] = candidate_policy
        candidate_incremental[candidate_name] = summarize_predictions(
            apply_policy(unresolved, candidate_policy),
            include_probability=True,
        )
    selected_name = max(
        candidate_names,
        key=lambda candidate_name: (
            candidate_policies[candidate_name].production_enabled,
            candidate_policies[candidate_name].validation_coverage,
            candidate_policies[candidate_name].validation_precision,
            candidate_summaries[candidate_name]["weighted_f1"],
            candidate_summaries[candidate_name]["accuracy"],
            candidate_summaries[candidate_name]["macro_f1"],
        ),
    )
    selected_config = next(
        (config for config in DEFAULT_CANDIDATE_CONFIGS if config.name == selected_name),
        None,
    )
    primary_ml = candidate_predictions[selected_name]
    policy = candidate_policies[selected_name]

    selected_by_scheme = {"unseen_session": primary_ml}
    for name in ("unseen_client", "unseen_merchant"):
        if selected_config is not None:
            selected_by_scheme[name] = evaluate_ml_config(frame, schemes[name], selected_config)
        else:
            components = [
                evaluate_ml_config(frame, schemes[name], config)
                for config in DEFAULT_CANDIDATE_CONFIGS
            ]
            selected_by_scheme[name] = ensemble_prediction_frames(components, name=ensemble_name)

    baseline_summaries = {
        name: summarize_predictions(predictions, include_probability=False)
        for name, predictions in baselines.items()
    }
    scheme_summaries = {}
    hybrid_summaries = {}
    incremental_ml_summaries = {}
    legacy_by_scheme = {}
    hybrid_frames = {}
    for name, predictions in selected_by_scheme.items():
        gated = apply_policy(predictions, policy)
        scheme_summaries[name] = summarize_predictions(gated, include_probability=True)
        if name == "unseen_session":
            scheme_legacy = baselines["legacy_complete"]
        else:
            scheme_legacy = evaluate_baselines(frame, schemes[name])["legacy_complete"]
        legacy_by_scheme[name] = scheme_legacy
        unresolved = unresolved_ml_predictions(predictions, scheme_legacy)
        incremental_ml_summaries[name] = summarize_predictions(
            apply_policy(unresolved, policy),
            include_probability=True,
        )
        hybrid = hybrid_predictions(frame, scheme_legacy, predictions, policy)
        hybrid_frames[name] = hybrid
        hybrid_summaries[name] = summarize_predictions(hybrid, include_probability=False)

    gate_checks = {}
    for name, summary in incremental_ml_summaries.items():
        acceptance = summary["automatic_acceptance"]
        unresolved_rows = int(summary["rows"])
        minimum_accepted = max(25, math.ceil(unresolved_rows * 0.01))
        legacy_metrics = summarize_predictions(
            legacy_by_scheme[name],
            include_probability=False,
        )
        hybrid_metrics = hybrid_summaries[name]
        gate_checks[name] = {
            "unresolved_rows": unresolved_rows,
            "ml_accepted": acceptance["accepted"],
            "ml_precision_at_policy": acceptance["precision"],
            "ml_coverage_of_unresolved": acceptance["coverage"],
            "minimum_accepted": minimum_accepted,
            "meets_95_percent_ml_precision": (
                acceptance["precision"] >= 0.95
                and acceptance["accepted"] >= minimum_accepted
            ),
            "legacy_overall_accuracy": legacy_metrics["accuracy"],
            "hybrid_overall_accuracy": hybrid_metrics["accuracy"],
            "no_accuracy_regression": hybrid_metrics["accuracy"] >= legacy_metrics["accuracy"],
        }
    gate_passed = bool(
        policy.production_enabled
        and all(check["meets_95_percent_ml_precision"] for check in gate_checks.values())
        and all(check["no_accuracy_regression"] for check in gate_checks.values())
    )
    if policy.production_enabled != gate_passed:
        policy = ConfidencePolicy(**{
            **policy.to_dict(),
            "production_enabled": gate_passed,
        })

    if selected_config is not None:
        final_model = _calibrated_fold_model(selected_config, frame)
        selected_config_payload: dict[str, Any] = asdict(selected_config)
    else:
        final_model = LocalModelEnsemble([
            _calibrated_fold_model(config, frame)
            for config in DEFAULT_CANDIDATE_CONFIGS
        ])
        selected_config_payload = {
            "name": ensemble_name,
            "algorithm": "agreement_ensemble",
            "components": [asdict(config) for config in DEFAULT_CANDIDATE_CONFIGS],
        }
    report = {
        "evaluation_schema_version": EVALUATION_SCHEMA_VERSION,
        "created_at": _datetime.datetime.now(_datetime.timezone.utc).isoformat(timespec="seconds"),
        "command_scope": {
            "include_legacy_directory": bool(include_legacy),
            "primary_dataset": "manifest-designated imports",
            "n_splits": int(n_splits),
            "selection_scheme": "unseen_session",
            "gate_schemes": ["unseen_session", "unseen_client", "unseen_merchant"],
        },
        "dataset": dataset.report_payload(),
        "production_model_rows": len(frame),
        "production_class_count": int(frame["model_target"].nunique()),
        "leakage_audit": leakage_audit,
        "baselines_unseen_session": baseline_summaries,
        "candidate_models_unseen_session": candidate_summaries,
        "candidate_incremental_ml_unseen_session": candidate_incremental,
        "candidate_confidence_policies": {
            name: candidate_policy.to_dict()
            for name, candidate_policy in candidate_policies.items()
        },
        "selected_model": {
            "config": selected_config_payload,
            "metadata": final_model.metadata(),
        },
        "selected_model_by_scheme": scheme_summaries,
        "incremental_ml_by_scheme": incremental_ml_summaries,
        "hybrid_by_scheme": hybrid_summaries,
        "confidence_policy": policy.to_dict(),
        "gate_checks": gate_checks,
        "gate_passed": gate_passed,
        "limitations": [
            "The primary training set excludes fixed/skip formats that production autocoding never handles.",
            "Client identity is used only for grouping and legacy-baseline isolation, never as an ML feature.",
            "The 4635/4640 labels are evaluated as one vehicle-expense family and resolved by user setting at runtime.",
            "Aggregate compression rows remain coded-only and are never expanded into duplicate labels.",
        ],
    }
    return report, final_model, policy


def activate_evaluated_model(
    report: dict[str, Any],
    model: Any,
    policy: ConfidencePolicy,
    *,
    artifact_dir: str | Path,
) -> dict[str, Any]:
    """Atomically activate only a model that passed every grouped gate."""
    if report.get("gate_passed") is not True or policy.production_enabled is not True:
        return {
            "activated": False,
            "reason": "production_gate_failed",
            "artifact_dir": str(Path(artifact_dir)),
        }
    gate_checks = report.get("gate_checks")
    if not isinstance(gate_checks, dict) or not gate_checks or any(
        not isinstance(check, dict)
        or check.get("meets_95_percent_ml_precision") is not True
        or check.get("no_accuracy_regression") is not True
        for check in gate_checks.values()
    ):
        raise ValueError("cannot activate a model with incomplete or failed grouped gates")
    dataset = report.get("dataset")
    if not isinstance(dataset, dict) or not dataset.get("fingerprint"):
        raise ValueError("cannot activate a model without a training-data fingerprint")
    created_at = str(report.get("created_at") or _datetime.datetime.now(_datetime.timezone.utc).isoformat(timespec="seconds"))
    stamp = re.sub(r"[^0-9]", "", created_at)[:14] or "model"
    fingerprint = str(dataset["fingerprint"])
    metadata = ArtifactMetadata(
        model_version=f"{stamp}-{fingerprint[:12]}",
        training_data_fingerprint=fingerprint,
        training_row_count=int(report.get("production_model_rows", 0)),
        class_count=int(report.get("production_class_count", 0)),
        feature_version=str(getattr(model, "feature_version", FEATURE_VERSION)),
        evaluation_summary={
            "gate_passed": True,
            "gate_checks": gate_checks,
            "confidence_policy": policy.to_dict(),
            "selected_model": report.get("selected_model", {}),
            "evaluation_schema_version": report.get("evaluation_schema_version", ""),
        },
        production_enabled=True,
        created_at=created_at,
        pandas_version=pd.__version__,
    )
    activated = ArtifactStore(artifact_dir).activate(model, metadata)
    return {
        "activated": True,
        "reason": "ok",
        "artifact_dir": str(Path(artifact_dir)),
        "artifact_path": str(activated.artifact_path),
        "model_version": activated.metadata.model_version,
        "checksum": activated.checksum,
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Tok Tenging local ML evaluation")
    subparsers = parser.add_subparsers(dest="command", required=True)
    evaluate = subparsers.add_parser("evaluate", help="run grouped Gate-1 evaluation")
    evaluate.add_argument("--training-data-dir", default="tok/trainingData")
    evaluate.add_argument("--legacy-training-dir", default="tok/trainingCoded")
    evaluate.add_argument("--include-legacy", action="store_true")
    evaluate.add_argument("--folds", type=int, default=3)
    evaluate.add_argument("--output", default="artifacts/ml/evaluation_report.json")
    rebuild = subparsers.add_parser(
        "rebuild",
        help="evaluate from scratch and atomically activate only if every gate passes",
    )
    rebuild.add_argument("--training-data-dir", default="tok/trainingData")
    rebuild.add_argument("--legacy-training-dir", default="tok/trainingCoded")
    rebuild.add_argument("--include-legacy", action="store_true")
    rebuild.add_argument("--folds", type=int, default=3)
    rebuild.add_argument("--output", default="artifacts/ml/evaluation_report.json")
    rebuild.add_argument("--artifact-dir", default="tok/trainingData/ml_artifacts")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if args.command in {"evaluate", "rebuild"}:
        report, model, policy = run_evaluation(
            training_data_dir=args.training_data_dir,
            legacy_training_dir=args.legacy_training_dir,
            include_legacy=args.include_legacy,
            n_splits=args.folds,
        )
        output_path = Path(args.output)
        _atomic_json_write(output_path, report)
        activation = {
            "activated": False,
            "reason": "evaluation_only",
        }
        if args.command == "rebuild":
            activation = activate_evaluated_model(
                report,
                model,
                policy,
                artifact_dir=args.artifact_dir,
            )
        print(json.dumps({
            "report": str(output_path),
            "gate_passed": report["gate_passed"],
            "selected_model": report["selected_model"]["config"]["name"],
            "policy": report["confidence_policy"],
            "activation": activation,
        }, ensure_ascii=False, indent=2))
        return 0
    return 2


if __name__ == "__main__":  # pragma: no cover - exercised by smoke command
    raise SystemExit(main())


__all__ = [
    "EVALUATION_SCHEMA_VERSION",
    "apply_policy",
    "activate_evaluated_model",
    "evaluate_baselines",
    "evaluate_ml_config",
    "fit_legacy_evaluation_coder",
    "hybrid_predictions",
    "run_evaluation",
    "summarize_predictions",
]

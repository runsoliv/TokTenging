"""Grouped validation, metrics, and conservative ML acceptance policy."""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from hashlib import sha256
from itertools import product
from typing import Any, Iterable, Sequence
import math
import numbers

import numpy as np
import pandas as pd


@dataclass(frozen=True)
class GroupedFold:
    name: str
    fold: int
    group_columns: tuple[str, ...]
    train_indices: tuple[int, ...]
    test_indices: tuple[int, ...]


@dataclass(frozen=True)
class ConfidencePolicy:
    probability_threshold: float
    margin_threshold: float
    min_class_support: int
    unseen_merchant_threshold: float
    weak_text_threshold: float
    sensitive_6220_threshold: float
    target_precision: float = 0.95
    validation_precision: float = 0.0
    validation_coverage: float = 0.0
    validation_accepted: int = 0
    production_enabled: bool = False

    def __post_init__(self) -> None:
        threshold_names = (
            "probability_threshold",
            "margin_threshold",
            "unseen_merchant_threshold",
            "weak_text_threshold",
            "sensitive_6220_threshold",
            "target_precision",
            "validation_precision",
            "validation_coverage",
        )
        for name in threshold_names:
            original = getattr(self, name)
            if isinstance(original, bool) or not isinstance(original, numbers.Real):
                raise ValueError(f"{name} must be a real number")
            value = float(original)
            if not math.isfinite(value) or not 0.0 <= value <= 1.0:
                raise ValueError(f"{name} must be finite and between zero and one")
            object.__setattr__(self, name, value)
        if (
            isinstance(self.min_class_support, bool)
            or not isinstance(self.min_class_support, numbers.Integral)
            or self.min_class_support < 0
        ):
            raise ValueError("min_class_support must be a non-negative integer")
        object.__setattr__(self, "min_class_support", int(self.min_class_support))
        if (
            isinstance(self.validation_accepted, bool)
            or not isinstance(self.validation_accepted, numbers.Integral)
            or self.validation_accepted < 0
        ):
            raise ValueError("validation_accepted must be a non-negative integer")
        object.__setattr__(self, "validation_accepted", int(self.validation_accepted))
        if type(self.production_enabled) is not bool:
            raise ValueError("production_enabled must be a boolean")

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ConfidencePolicy":
        if not isinstance(payload, dict):
            raise ValueError("confidence policy must be an object")
        allowed = {field.name for field in cls.__dataclass_fields__.values()}
        required = {
            "probability_threshold",
            "margin_threshold",
            "min_class_support",
            "unseen_merchant_threshold",
            "weak_text_threshold",
            "sensitive_6220_threshold",
        }
        missing = sorted(required - set(payload))
        if missing:
            raise ValueError(f"confidence policy missing fields: {', '.join(missing)}")
        return cls(**{key: payload[key] for key in allowed if key in payload})


class _UnionFind:
    def __init__(self, size: int):
        self.parent = list(range(size))
        self.rank = [0] * size

    def find(self, value: int) -> int:
        parent = self.parent[value]
        if parent != value:
            self.parent[value] = self.find(parent)
        return self.parent[value]

    def union(self, first: int, second: int) -> None:
        first_root = self.find(first)
        second_root = self.find(second)
        if first_root == second_root:
            return
        if self.rank[first_root] < self.rank[second_root]:
            first_root, second_root = second_root, first_root
        self.parent[second_root] = first_root
        if self.rank[first_root] == self.rank[second_root]:
            self.rank[first_root] += 1


def _normalized_group_value(value: Any, fallback: str) -> str:
    if value is None:
        return fallback
    try:
        if pd.isna(value):
            return fallback
    except Exception:
        pass
    text = str(value).strip()
    return text or fallback


def _protected_components(
    frame: pd.DataFrame,
    group_columns: Sequence[str],
) -> dict[int, list[int]]:
    size = len(frame)
    union_find = _UnionFind(size)
    first_for_group: dict[tuple[str, str], int] = {}
    first_for_identity: dict[str, int] = {}
    for position, (_, row) in enumerate(frame.reset_index(drop=True).iterrows()):
        # Each requested dimension is independently protected.  Treating a
        # (client, session) pair as one composite would still allow the same
        # client to appear in both train and test under different sessions.
        for column in group_columns:
            key = (
                str(column),
                _normalized_group_value(row.get(column), f"__missing_{column}_{position}"),
            )
            if key in first_for_group:
                union_find.union(position, first_for_group[key])
            else:
                first_for_group[key] = position
        identity = _normalized_group_value(
            row.get("semantic_identity"),
            f"__unique_identity_{position}",
        )
        if identity in first_for_identity:
            union_find.union(position, first_for_identity[identity])
        else:
            first_for_identity[identity] = position
    components: dict[int, list[int]] = defaultdict(list)
    for position in range(size):
        components[union_find.find(position)].append(position)
    return components


def make_grouped_folds(
    frame: pd.DataFrame,
    *,
    name: str,
    group_columns: Sequence[str],
    label_column: str = "model_target",
    n_splits: int = 5,
    seed: int = 1729,
) -> list[GroupedFold]:
    """Create deterministic stratified-ish folds over protected group components."""
    if label_column not in frame.columns:
        raise ValueError(f"missing label column: {label_column}")
    missing_groups = [column for column in group_columns if column not in frame.columns]
    if missing_groups:
        raise ValueError(f"missing group columns: {', '.join(missing_groups)}")
    if n_splits < 2:
        raise ValueError("n_splits must be at least two")
    if len(frame) < n_splits:
        raise ValueError("not enough rows for requested folds")

    normalized = frame.reset_index(drop=True)
    components = _protected_components(normalized, group_columns)
    if len(components) < n_splits:
        raise ValueError("not enough independent groups for requested folds")
    labels = normalized[label_column].astype(str).tolist()
    total_class = Counter(labels)
    total_size = len(normalized)

    component_records = []
    for positions in components.values():
        counts = Counter(labels[position] for position in positions)
        material = f"{seed}|{','.join(map(str, positions))}".encode("utf-8")
        tie_break = sha256(material).hexdigest()
        component_records.append((positions, counts, tie_break))
    component_records.sort(
        key=lambda item: (-len(item[0]), -max(item[1].values()), item[2]),
    )

    fold_positions: list[list[int]] = [[] for _ in range(n_splits)]
    fold_counts: list[Counter] = [Counter() for _ in range(n_splits)]

    def score(candidate_fold: int, positions: list[int], counts: Counter) -> tuple[float, int, int]:
        class_score = 0.0
        for label, total in total_class.items():
            proportions = []
            for fold_index in range(n_splits):
                value = fold_counts[fold_index][label]
                if fold_index == candidate_fold:
                    value += counts[label]
                proportions.append(value / total)
            class_score += float(np.std(proportions))
        sizes = []
        for fold_index in range(n_splits):
            value = len(fold_positions[fold_index])
            if fold_index == candidate_fold:
                value += len(positions)
            sizes.append(value / total_size)
        size_score = float(np.std(sizes))
        return class_score + (0.35 * size_score), len(fold_positions[candidate_fold]), candidate_fold

    for component_index, (positions, counts, _tie) in enumerate(component_records):
        if component_index < n_splits:
            chosen = component_index
        else:
            chosen = min(range(n_splits), key=lambda fold_index: score(fold_index, positions, counts))
        fold_positions[chosen].extend(positions)
        fold_counts[chosen].update(counts)

    all_positions = set(range(len(normalized)))
    folds = []
    for fold_index, test in enumerate(fold_positions):
        test_indices = tuple(sorted(test))
        train_indices = tuple(sorted(all_positions - set(test_indices)))
        folds.append(GroupedFold(
            name=name,
            fold=fold_index,
            group_columns=tuple(group_columns),
            train_indices=train_indices,
            test_indices=test_indices,
        ))
    return folds


def default_validation_schemes(frame: pd.DataFrame, *, n_splits: int = 5) -> dict[str, list[GroupedFold]]:
    merchant_groups = frame.copy()
    merchant_groups["_evaluation_merchant"] = _merchant_evaluation_components(merchant_groups)
    return {
        "unseen_session": make_grouped_folds(
            frame,
            name="unseen_session",
            group_columns=("session_id",),
            n_splits=n_splits,
        ),
        "unseen_client": make_grouped_folds(
            frame,
            name="unseen_client",
            group_columns=("client_id",),
            n_splits=n_splits,
        ),
        "unseen_merchant": make_grouped_folds(
            merchant_groups,
            name="unseen_merchant",
            group_columns=("_evaluation_merchant",),
            n_splits=n_splits,
        ),
    }


def _merchant_evaluation_components(frame: pd.DataFrame) -> list[str]:
    """Link every merchant identity visible to the feature layer.

    A blank merchant family must not make a row unique when its validated
    kennitala or normalized text repeats.  Conversely, rows sharing a family
    stay linked even if only some contain a kennitala.  This prevents the
    nominal unseen-merchant split from leaking an exact identifier.
    """
    normalized = frame.reset_index(drop=True)
    union_find = _UnionFind(len(normalized))
    first_for_signal: dict[tuple[str, str], int] = {}
    for position, (_, row) in enumerate(normalized.iterrows()):
        signals = []
        for column, prefix in (
            ("merchant_family", "family"),
            ("valid_kennitala", "kt"),
            ("normalized_text", "text"),
        ):
            value = _normalized_group_value(row.get(column), "")
            if value:
                signals.append((prefix, value))
        for signal in signals:
            if signal in first_for_signal:
                union_find.union(position, first_for_signal[signal])
            else:
                first_for_signal[signal] = position
    roots: dict[int, str] = {}
    output = []
    for position in range(len(normalized)):
        root = union_find.find(position)
        roots.setdefault(root, f"merchant_component_{len(roots):06d}")
        output.append(roots[root])
    return output


def audit_fold_leakage(frame: pd.DataFrame, fold: GroupedFold) -> dict[str, Any]:
    normalized = frame.reset_index(drop=True)
    missing_group_columns = []
    for column in fold.group_columns:
        if column in normalized.columns:
            continue
        if column == "_evaluation_merchant":
            normalized[column] = _merchant_evaluation_components(normalized)
        else:
            missing_group_columns.append(column)
    train = normalized.iloc[list(fold.train_indices)]
    test = normalized.iloc[list(fold.test_indices)]
    group_overlap = {}
    for column in fold.group_columns:
        if column not in normalized.columns:
            continue
        overlap = set(train[column].astype(str)) & set(test[column].astype(str))
        group_overlap[column] = sorted(overlap)
    identity_overlap = []
    if "semantic_identity" in normalized.columns:
        identity_overlap = sorted(
            set(train["semantic_identity"].astype(str))
            & set(test["semantic_identity"].astype(str))
        )
    return {
        "ok": not missing_group_columns and not any(group_overlap.values()) and not identity_overlap,
        "missing_group_columns": missing_group_columns,
        "group_overlap": group_overlap,
        "semantic_identity_overlap": identity_overlap,
        "train_rows": len(train),
        "test_rows": len(test),
    }


def classification_metrics(y_true: Sequence[str], y_pred: Sequence[str]) -> dict[str, Any]:
    if len(y_true) != len(y_pred):
        raise ValueError("y_true and y_pred must have equal length")
    labels = sorted(set(map(str, y_true)) | set(map(str, y_pred)))
    rows = []
    total = len(y_true)
    correct = sum(str(actual) == str(predicted) for actual, predicted in zip(y_true, y_pred))
    for label in labels:
        tp = sum(str(a) == label and str(p) == label for a, p in zip(y_true, y_pred))
        fp = sum(str(a) != label and str(p) == label for a, p in zip(y_true, y_pred))
        fn = sum(str(a) == label and str(p) != label for a, p in zip(y_true, y_pred))
        support = sum(str(a) == label for a in y_true)
        precision = tp / (tp + fp) if tp + fp else 0.0
        recall = tp / (tp + fn) if tp + fn else 0.0
        f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
        rows.append({
            "account": label,
            "precision": precision,
            "recall": recall,
            "f1": f1,
            "support": support,
        })
    macro_f1 = sum(row["f1"] for row in rows) / len(rows) if rows else 0.0
    weighted_f1 = sum(row["f1"] * row["support"] for row in rows) / total if total else 0.0
    return {
        "rows": total,
        "accuracy": correct / total if total else 0.0,
        "macro_f1": macro_f1,
        "weighted_f1": weighted_f1,
        "per_class": rows,
    }


def top_k_accuracy(y_true: Sequence[str], ranked_predictions: Sequence[Sequence[str]], k: int) -> float:
    if len(y_true) != len(ranked_predictions):
        raise ValueError("truth and predictions must have equal length")
    if not y_true:
        return 0.0
    hits = sum(str(actual) in {str(value) for value in ranked[:k]} for actual, ranked in zip(y_true, ranked_predictions))
    return hits / len(y_true)


def calibration_summary(
    probabilities: Sequence[float],
    correct: Sequence[bool],
    *,
    bins: int = 10,
) -> dict[str, Any]:
    if len(probabilities) != len(correct):
        raise ValueError("probabilities and correct flags must have equal length")
    if not probabilities:
        return {"ece": 0.0, "brier_top_label": 0.0, "bins": []}
    probabilities_array = np.asarray(probabilities, dtype=float)
    correct_array = np.asarray(correct, dtype=float)
    boundaries = np.linspace(0.0, 1.0, bins + 1)
    output_bins = []
    ece = 0.0
    for index in range(bins):
        lower, upper = boundaries[index], boundaries[index + 1]
        mask = (
            (probabilities_array >= lower)
            & (probabilities_array <= upper if index == bins - 1 else probabilities_array < upper)
        )
        count = int(mask.sum())
        if not count:
            continue
        mean_confidence = float(probabilities_array[mask].mean())
        accuracy = float(correct_array[mask].mean())
        ece += (count / len(probabilities_array)) * abs(mean_confidence - accuracy)
        output_bins.append({
            "lower": float(lower),
            "upper": float(upper),
            "count": count,
            "mean_confidence": mean_confidence,
            "accuracy": accuracy,
        })
    brier = float(np.mean((probabilities_array - correct_array) ** 2))
    return {"ece": ece, "brier_top_label": brier, "bins": output_bins}


def prediction_is_accepted(
    record: Any,
    policy: ConfidencePolicy,
    *,
    require_production_enabled: bool = True,
) -> tuple[bool, str]:
    if require_production_enabled and not policy.production_enabled:
        return False, "policy_not_enabled"
    getter = record.get if hasattr(record, "get") else lambda key, default=None: getattr(record, key, default)
    try:
        probability = float(getter("probability", 0.0) or 0.0)
        margin = float(getter("margin", 0.0) or 0.0)
        support = int(getter("class_support", 0) or 0)
    except (TypeError, ValueError, OverflowError):
        return False, "invalid_prediction_values"
    predicted = str(getter("predicted", "") or "")
    if (
        not predicted
        or not math.isfinite(probability)
        or not math.isfinite(margin)
        or not 0.0 <= probability <= 1.0
        or not 0.0 <= margin <= 1.0
        or support < 0
    ):
        return False, "invalid_prediction_values"
    if bool(getter("unknown_treasury", False)):
        return False, "unknown_treasury"
    if support < policy.min_class_support:
        return False, "rare_or_unsupported_class"
    if margin < policy.margin_threshold:
        return False, "small_candidate_margin"
    threshold = policy.probability_threshold
    if not bool(getter("merchant_familiar", False)):
        threshold = max(threshold, policy.unseen_merchant_threshold)
    if bool(getter("weak_text", False)):
        threshold = max(threshold, policy.weak_text_threshold)
    if predicted == "6220":
        threshold = max(threshold, policy.sensitive_6220_threshold)
    if probability < threshold:
        return False, "probability_below_policy"
    return True, "accepted"


def select_confidence_policy(
    validation_predictions: pd.DataFrame,
    *,
    target_precision: float = 0.95,
    minimum_accepted: int | None = None,
) -> ConfidencePolicy:
    required = {
        "correct",
        "probability",
        "margin",
        "class_support",
        "merchant_familiar",
        "weak_text",
        "predicted",
    }
    missing = sorted(required - set(validation_predictions.columns))
    if missing:
        raise ValueError(f"validation predictions missing columns: {', '.join(missing)}")
    total = len(validation_predictions)
    if not total:
        return ConfidencePolicy(0.99, 0.50, 20, 0.995, 0.995, 0.999, target_precision=target_precision)
    minimum_accepted = minimum_accepted or max(25, math.ceil(total * 0.01))
    candidates = []
    for probability, margin, support in product(
        (0.55, 0.65, 0.75, 0.82, 0.88, 0.92, 0.95, 0.97, 0.99),
        (0.05, 0.10, 0.15, 0.20, 0.30, 0.40),
        (3, 5, 10, 20),
    ):
        policy = ConfidencePolicy(
            probability_threshold=probability,
            margin_threshold=margin,
            min_class_support=support,
            unseen_merchant_threshold=min(0.995, probability + 0.05),
            weak_text_threshold=min(0.995, probability + 0.08),
            sensitive_6220_threshold=min(0.999, probability + 0.07),
            target_precision=target_precision,
        )
        accepted = [
            prediction_is_accepted(row, policy, require_production_enabled=False)[0]
            for _, row in validation_predictions.iterrows()
        ]
        accepted_count = sum(accepted)
        if not accepted_count:
            continue
        accepted_frame = validation_predictions.loc[accepted]
        precision = float(accepted_frame["correct"].astype(bool).mean())
        coverage = accepted_count / total
        candidates.append((precision >= target_precision and accepted_count >= minimum_accepted, coverage, precision, accepted_count, policy))
    if not candidates:
        return ConfidencePolicy(0.99, 0.50, 20, 0.995, 0.995, 0.999, target_precision=target_precision)
    meeting = [candidate for candidate in candidates if candidate[0]]
    chosen = max(meeting or candidates, key=lambda item: (item[1] if item[0] else item[2], item[2], item[3]))
    enabled, coverage, precision, accepted_count, policy = chosen
    return ConfidencePolicy(
        probability_threshold=policy.probability_threshold,
        margin_threshold=policy.margin_threshold,
        min_class_support=policy.min_class_support,
        unseen_merchant_threshold=policy.unseen_merchant_threshold,
        weak_text_threshold=policy.weak_text_threshold,
        sensitive_6220_threshold=policy.sensitive_6220_threshold,
        target_precision=target_precision,
        validation_precision=precision,
        validation_coverage=coverage,
        validation_accepted=accepted_count,
        production_enabled=bool(enabled),
    )


__all__ = [
    "ConfidencePolicy",
    "GroupedFold",
    "audit_fold_leakage",
    "calibration_summary",
    "classification_metrics",
    "default_validation_schemes",
    "make_grouped_folds",
    "prediction_is_accepted",
    "select_confidence_policy",
    "top_k_accuracy",
]

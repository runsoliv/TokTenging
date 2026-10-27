"""Leakage-resistant feature preparation shared by ML training and inference."""

from __future__ import annotations

from decimal import Decimal, InvalidOperation
import math
import re
from typing import Any, Mapping

try:
    from .auto_coder import merchant_family, normalize_text
    from .ml_dataset import VEHICLE_EXPENSE_FAMILY, normalize_kennitala
except ImportError:  # pragma: no cover - direct script compatibility
    from auto_coder import merchant_family, normalize_text
    from ml_dataset import VEHICLE_EXPENSE_FAMILY, normalize_kennitala


FEATURE_VERSION = "tok-ml-document-v1"

MASKED_CARD_SUFFIX = re.compile(r"\*{1,3}\s*-?\s*\d{4,}", re.IGNORECASE)
LONG_REFERENCE_FRAGMENT = re.compile(r"(?<!\d)\d{8,}(?!\d)")

# These fields are post-coding artifacts or grouping metadata.  They are kept
# here as an explicit guardrail so callers cannot accidentally pass them into
# the free-form feature document.
FORBIDDEN_FEATURE_FIELDS = frozenset({
    "cnf",
    "confidence",
    "status",
    "auto_code_source",
    "_auto_code_source",
    "auto_debit_status",
    "auto_debit_confidence",
    "auto_debit_reason",
    "target_code",
    "model_target",
    "debit",
    "credit",
    "session_id",
    "source_workbook",
    "source_hash",
    "client_id",
})


def _value(row: Mapping[str, Any], name: str, default: Any = "") -> Any:
    try:
        return row.get(name, default)
    except AttributeError:
        return default


def _finite_float(value: Any) -> float | None:
    if value is None:
        return None
    try:
        number = float(value)
    except (TypeError, ValueError, InvalidOperation):
        return None
    return number if math.isfinite(number) else None


def amount_bucket(value: Any) -> str:
    number = _finite_float(value)
    if number is None:
        return "amount_missing"
    amount = abs(number)
    limits = (
        (1_000, "amount_0000_001k"),
        (3_000, "amount_001k_003k"),
        (7_500, "amount_003k_007k"),
        (15_000, "amount_007k_015k"),
        (30_000, "amount_015k_030k"),
        (75_000, "amount_030k_075k"),
        (150_000, "amount_075k_150k"),
        (300_000, "amount_150k_300k"),
        (750_000, "amount_300k_750k"),
    )
    for limit, label in limits:
        if amount <= limit:
            return label
    return "amount_750k_plus"


def amount_log_bucket(value: Any) -> str:
    number = _finite_float(value)
    if number is None:
        return "amount_log_missing"
    level = min(7, max(0, int(math.log10(abs(number) + 1))))
    return f"amount_log_{level}"


def amount_roundness(value: Any) -> str:
    number = _finite_float(value)
    if number is None or number == 0:
        return "amount_round_missing"
    try:
        amount = abs(Decimal(str(number)))
        if amount % Decimal("10000") == 0:
            return "amount_round_10000"
        if amount % Decimal("1000") == 0:
            return "amount_round_1000"
        if amount % Decimal("100") == 0:
            return "amount_round_100"
    except (InvalidOperation, ValueError):
        return "amount_round_missing"
    return "amount_not_round"


def month_token(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    if not text:
        return "month_missing"
    match = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", text)
    if match:
        month = int(match.group(2))
        return f"month_{month:02d}"
    return "month_missing"


def model_target_code(code: Any) -> str:
    text = "" if code is None else str(code).strip()
    return VEHICLE_EXPENSE_FAMILY if text in {"4635", "4640"} else text


def normalize_ml_text(value: Any) -> str:
    """Normalize changing bank/card references before text vectorization."""
    original = "" if value is None else str(value)
    original = MASKED_CARD_SUFFIX.sub(" ", original)
    original = LONG_REFERENCE_FRAGMENT.sub(" ", original)
    return normalize_text(original)


def build_feature_document(row: Mapping[str, Any]) -> str:
    """Create one deterministic document without client or answer leakage."""
    # Prefer original text so masked suffixes remain visible to the stricter ML
    # normalizer; ``normalized_text`` is only a fallback for API callers.
    text = normalize_ml_text(_value(row, "text") or _value(row, "normalized_text"))
    raw = normalize_ml_text(_value(row, "raw_explanation"))
    family = normalize_text(_value(row, "merchant_family") or merchant_family(text))
    bank = normalize_text(_value(row, "bank_type")) or "unknown"
    direction = normalize_text(_value(row, "direction")) or "unknown"
    kennitala = normalize_kennitala(_value(row, "valid_kennitala"))
    weak_text = bool(_value(row, "weak_text", False))

    parts = [
        f"bank_{bank.replace(' ', '_')}",
        f"direction_{direction.replace(' ', '_')}",
        amount_bucket(_value(row, "amount")),
        amount_log_bucket(_value(row, "amount")),
        amount_roundness(_value(row, "amount")),
        month_token(_value(row, "date")),
        "weak_text" if weak_text else "informative_text",
        "raw_present" if raw else "raw_missing",
        "kennitala_present" if kennitala else "kennitala_missing",
    ]
    if text:
        parts.append(f"text {text}")
    if family:
        parts.append(f"merchant_family {family}")
    if raw:
        parts.append(f"raw_explanation {raw}")
    if kennitala:
        parts.append(f"kennitala_value kt{kennitala}")
    return " | ".join(parts)


def feature_documents(records) -> list[str]:
    if hasattr(records, "to_dict"):
        records = records.to_dict("records")
    return [build_feature_document(record) for record in records]


__all__ = [
    "FEATURE_VERSION",
    "FORBIDDEN_FEATURE_FIELDS",
    "amount_bucket",
    "build_feature_document",
    "feature_documents",
    "model_target_code",
    "normalize_ml_text",
]

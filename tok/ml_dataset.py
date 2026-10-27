"""Deterministic, read-only training dataset construction for Tok Tenging ML.

The loader deliberately keeps workbook discovery, row validation, raw/coded
pairing, and deduplication separate from model training.  Import sessions are
accepted only through their manifests.  A configured legacy ``trainingCoded``
directory is also supported, but generated/combined files are excluded.

Nothing in this module writes to a source workbook or training archive.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from decimal import Decimal, InvalidOperation
from hashlib import sha256
from pathlib import Path
from typing import Any, Iterable
import datetime as _datetime
import json
import math
import re

import pandas as pd

try:
    from .auto_coder import (
        COUNTER_ACCOUNT_CODES,
        clean_code,
        merchant_family,
        merchant_key,
        normalize_text,
    )
except ImportError:  # pragma: no cover - direct script compatibility
    from auto_coder import (
        COUNTER_ACCOUNT_CODES,
        clean_code,
        merchant_family,
        merchant_key,
        normalize_text,
    )


DATASET_SCHEMA_VERSION = "1.0"
FEATURE_SCHEMA_VERSION = "tok-ml-features-v1"
VEHICLE_EXPENSE_FAMILY = "__VEHICLE_EXPENSE_4635_4640__"
PRODUCTION_EXCLUDED_BANK_TYPES = frozenset({
    "islandsbanki_special",
    "islandsbanki_innheimta",
    "sala_yfirlit",
})

# These roles were verified over all 10,260 explicitly signed retained import
# rows.  Bank-side validation and target exclusion are intentionally separate:
# the legacy constant wrongly includes sales accounts 1000/1010, while missing
# bank-side 7840.  Per-row sign still decides which side is the target.
BANK_SIDE_ACCOUNT_CODES = frozenset({
    "7810", "7820", "7825", "7830", "7840", "7846", "9336",
})
NON_CLASSIFICATION_ACCOUNT_CODES = frozenset(set(BANK_SIDE_ACCOUNT_CODES) | {
    "7841", "7842", "7850", "9330",
})
DEFAULT_CONTROL_ACCOUNT_CODES = NON_CLASSIFICATION_ACCOUNT_CODES

CANONICAL_COLUMN_ALIASES = {
    "DATE": {"date", "dagsetning", "dags"},
    "TEXT": {"text", "texti", "lysing", "skyring"},
    "DEBIT": {"debit", "debet"},
    "CREDIT": {"credit", "kredit"},
    "ID": {"id", "kennitala", "kt"},
    "AMOUNT": {"amount", "upphaed", "upphaedisk", "fjarmagn"},
    "Positive/Negative": {
        "positivenegative",
        "plusminus",
        "stefna",
        "merki",
    },
}

REQUIRED_CODED_COLUMNS = ("DATE", "TEXT", "DEBIT", "CREDIT", "AMOUNT")

RAW_TEXT_HEADER_TOKENS = {
    "texti",
    "skyring",
    "skyringar",
    "lysing",
    "nanariskyring",
    "mottakandi",
    "greidandi",
    "greidanda",
    "vidtakanda",
    "nafnvidtakanda",
    "nafngreidanda",
    "krafuheiti",
    "upplysingar",
    "motadili",
    "soluadili",
    "tegund",
    "textalykill",
    "faerslulykill",
    "soluadilaflokkun",
}

REFERENCE_HEADER_TOKENS = (
    "einkvaemurlykill",
    "faersluaukenni",
    "arnaukenni",
    "arn",
    "sedilnumer",
    "sedilnr",
    "tnrsedilnr",
    "krofunumer",
    "bokhaldstilvisun",
    "heimild",
    "tilvisun",
)

KENNITALA_HEADER_TOKENS = (
    "kennitala",
    "kt",
    "tilvisun",
    "mottakandi",
    "greidandi",
)

# Card statements can contain the client's/cardholder's kennitala as well as
# the merchant's.  The former is effectively client identity leakage and must
# never be used as a transaction/merchant feature.  Prefer counterparty fields
# explicitly, then consider other strict ten-digit candidates.
KENNITALA_EXCLUDED_HEADER_TOKENS = (
    "korthafa",
    "cardholder",
    "kortnumer",
    "kortanumer",
    "cardnumber",
)

KENNITALA_PREFERRED_HEADER_TOKENS = (
    "kennitalasoluadila",
    "kennitalamotadili",
    "kennitalamottakanda",
    "kennitalavidtakanda",
    "kennitalagreidanda",
    "merchantkennitala",
    "counterpartykennitala",
)

LEAKAGE_HEADER_TOKENS = {
    "cnf",
    "confidence",
    "status",
    "autocodesource",
    "autodebitstatus",
    "autodebitconfidence",
    "autodebitreason",
}

WEAK_NORMALIZED_TEXT = {
    "",
    "reikningur",
    "reikn",
    "kreditkort",
    "debitkortafaersla",
    "greidsla",
    "innborgun",
    "utborgun",
    "laun",
}

LEGACY_GENERIC_PATTERNS = (
    re.compile(r"^(reikn|reikningur)\s+\d+$"),
    re.compile(r"^(innb|innborgun|innborgad|sala|vsk|laun)$"),
)

SKIP_LEGACY_NAME_PARTS = (
    "combined",
    "compressed_txs",
    "compressed transactions",
    "temporary",
    "temp",
)


@dataclass(frozen=True)
class DatasetIssue:
    level: str
    reason: str
    source: str = ""
    row: int | None = None
    detail: str = ""


@dataclass(frozen=True)
class TrainingSource:
    path: Path
    source_type: str
    source_id: str
    session_dir: Path | None = None
    manifest: dict[str, Any] = field(default_factory=dict)
    bank_type: str = ""
    client_id: str = ""
    paired_path: Path | None = None
    raw_path: Path | None = None
    compression_map_path: Path | None = None
    source_hash: str = ""
    raw_hash: str = ""
    paired_hash: str = ""
    compression_map_hash: str = ""
    manifest_hash: str = ""


@dataclass
class PairingInfo:
    status: str
    raw_values: dict[str, Any] = field(default_factory=dict)
    raw_names: dict[str, str] = field(default_factory=dict)
    raw_index: str = ""


@dataclass
class TrainingDataset:
    frame: pd.DataFrame
    report: dict[str, Any]
    fingerprint: str
    issues: list[DatasetIssue]

    def copy_frame(self) -> pd.DataFrame:
        return self.frame.copy(deep=True)

    def model_frame(self) -> pd.DataFrame:
        """Return only rows from formats that production autocoding handles."""
        if "production_eligible" not in self.frame.columns:
            return self.copy_frame()
        return self.frame.loc[self.frame["production_eligible"].astype(bool)].copy(deep=True)

    def report_payload(self) -> dict[str, Any]:
        payload = dict(self.report)
        payload["fingerprint"] = self.fingerprint
        payload["issues"] = [asdict(issue) for issue in self.issues]
        return payload


def _missing(value: Any) -> bool:
    if value is None:
        return True
    try:
        result = pd.isna(value)
        return bool(result) if not hasattr(result, "__len__") else False
    except Exception:
        return False


def _text(value: Any) -> str:
    if _missing(value):
        return ""
    return str(value).strip()


def _header_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", normalize_text(value))


def normalize_kennitala(value: Any) -> str:
    """Return exactly ten digits, allowing only one optional source hyphen."""
    text = _text(value)
    text = re.sub(r"^(\d+)\.0+$", r"\1", text)
    if not re.fullmatch(r"\d{6}-?\d{4}", text):
        return ""
    return text.replace("-", "")


def _canonical_reference(value: Any) -> str:
    text = _text(value)
    text = re.sub(r"^(\d+)\.0+$", r"\1", text)
    if not text:
        return ""
    return re.sub(r"[^a-z0-9]", "", normalize_text(text))


def _identity_id(value: Any) -> str:
    return normalize_kennitala(value) or _canonical_reference(value)


def _coded_identity_key(row: Any) -> tuple[str, str, str, str, str]:
    return (
        _date_key(row.get("DATE")),
        _amount_key(row.get("AMOUNT")),
        _text(row.get("Positive/Negative")),
        normalize_text(row.get("TEXT")),
        _identity_id(row.get("ID")),
    )


def _amount_decimal(value: Any) -> Decimal | None:
    if _missing(value):
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        try:
            if isinstance(value, float) and math.isnan(value):
                return None
            return Decimal(str(value))
        except (InvalidOperation, ValueError):
            return None
    text = _text(value).replace("\u00a0", "").replace(" ", "")
    if not text:
        return None
    negative = text.startswith("-") or (text.startswith("(") and text.endswith(")"))
    text = text.strip("()+-")
    text = re.sub(r"(?i)(?:isk|kr\.?)$", "", text)
    text = re.sub(r"[^0-9,.]", "", text)
    if not text:
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    elif text.count(".") == 1 and len(text.rsplit(".", 1)[1]) == 3:
        text = text.replace(".", "")
    elif text.count(".") > 1:
        text = text.replace(".", "")
    try:
        amount = Decimal(text)
    except InvalidOperation:
        return None
    return -amount if negative else amount


def _amount_key(value: Any) -> str:
    amount = _amount_decimal(value)
    if amount is None:
        return ""
    return format(amount.quantize(Decimal("0.01")), "f")


def _date_key(value: Any) -> str:
    if _missing(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.date().isoformat()
    if isinstance(value, (_datetime.datetime, _datetime.date)):
        return value.date().isoformat() if isinstance(value, _datetime.datetime) else value.isoformat()
    text = _text(value)
    if not text:
        return ""
    try:
        parsed = pd.to_datetime(text, dayfirst=True, errors="raise")
        return parsed.date().isoformat()
    except Exception:
        return normalize_text(text)


def _direction(value: Any, amount: Any = None) -> str:
    text = normalize_text(value)
    if _text(value) == "-" or text in {"negative", "minus", "ut", "debit"}:
        return "outgoing"
    if _text(value) == "+" or text in {"positive", "plus", "inn", "credit"}:
        return "incoming"
    parsed = _amount_decimal(amount)
    if parsed is not None and parsed < 0:
        return "outgoing"
    return ""


def _file_sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _read_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    if not isinstance(payload, dict):
        raise ValueError("JSON root must be an object")
    return payload


def _safe_child(directory: Path, name: Any) -> Path | None:
    if not name:
        return None
    candidate = directory / Path(str(name)).name
    try:
        if candidate.resolve().parent != directory.resolve():
            return None
    except Exception:
        return None
    return candidate


def _column_map(columns: Iterable[Any]) -> tuple[dict[Any, str], list[str]]:
    by_key: dict[str, list[Any]] = defaultdict(list)
    for column in columns:
        by_key[_header_key(column)].append(column)
    rename: dict[Any, str] = {}
    ambiguous: list[str] = []
    for canonical, aliases in CANONICAL_COLUMN_ALIASES.items():
        matches: list[Any] = []
        for alias in aliases | {_header_key(canonical)}:
            matches.extend(by_key.get(alias, []))
        matches = list(dict.fromkeys(matches))
        if len(matches) == 1:
            rename[matches[0]] = canonical
        elif len(matches) > 1:
            ambiguous.append(canonical)
    return rename, ambiguous


def _canonical_frame(frame: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    rename, ambiguous = _column_map(frame.columns)
    return frame.rename(columns=rename).copy(), ambiguous


class TrainingDataLoader:
    """Build a reproducible, defensively validated local training dataset."""

    def __init__(
        self,
        *,
        training_data_dir: str | Path,
        legacy_training_dir: str | Path | None = None,
        include_legacy: bool = True,
        control_account_codes: Iterable[str] | None = None,
    ):
        self.training_data_dir = Path(training_data_dir)
        self.legacy_training_dir = Path(legacy_training_dir) if legacy_training_dir else None
        self.include_legacy = bool(include_legacy)
        self.control_account_codes = set(control_account_codes or DEFAULT_CONTROL_ACCOUNT_CODES)

    def discover_sources(self) -> tuple[list[TrainingSource], list[DatasetIssue], dict[str, int]]:
        issues: list[DatasetIssue] = []
        counts = Counter()
        candidates: list[TrainingSource] = []
        imports_root = self.training_data_dir / "imports"

        if imports_root.is_dir():
            for session_dir in sorted(
                (path for path in imports_root.iterdir() if path.is_dir()),
                key=lambda path: str(path).casefold(),
            ):
                counts["import_sessions_discovered"] += 1
                manifest_path = session_dir / "manifest.json"
                if not manifest_path.is_file():
                    issues.append(DatasetIssue("skip", "missing_manifest", str(session_dir)))
                    counts["sources_skipped"] += 1
                    continue
                try:
                    manifest = _read_json(manifest_path)
                except Exception as exc:
                    issues.append(DatasetIssue("skip", "invalid_manifest", str(session_dir), detail=str(exc)))
                    counts["sources_skipped"] += 1
                    continue

                archive_files = manifest.get("archive_files") if isinstance(manifest.get("archive_files"), dict) else {}
                coded_name = archive_files.get("coded") or manifest.get("output_coded_filename")
                coded_path = _safe_child(session_dir, coded_name)
                if coded_path is None or not coded_path.is_file():
                    issues.append(DatasetIssue("skip", "manifest_coded_workbook_missing", str(session_dir)))
                    counts["sources_skipped"] += 1
                    continue
                if coded_path.suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
                    issues.append(DatasetIssue("skip", "unsupported_coded_workbook", str(coded_path)))
                    counts["sources_skipped"] += 1
                    continue

                paired_path = _safe_child(session_dir, archive_files.get("paired_rows") or "paired_rows.csv")
                raw_path = _safe_child(session_dir, archive_files.get("raw") or "raw.xlsx")
                compression_path = _safe_child(
                    session_dir,
                    archive_files.get("compression_map") or "compression_map.json",
                )
                try:
                    source_hash = _file_sha256(coded_path)
                    raw_hash = _file_sha256(raw_path) if raw_path and raw_path.is_file() else ""
                    paired_hash = _file_sha256(paired_path) if paired_path and paired_path.is_file() else ""
                    compression_map_hash = (
                        _file_sha256(compression_path)
                        if compression_path and compression_path.is_file()
                        else ""
                    )
                    manifest_hash = _file_sha256(manifest_path)
                except OSError as exc:
                    issues.append(DatasetIssue("skip", "source_hash_failed", str(coded_path), detail=str(exc)))
                    counts["sources_skipped"] += 1
                    continue
                candidates.append(TrainingSource(
                    path=coded_path,
                    source_type="import",
                    source_id=session_dir.name,
                    session_dir=session_dir,
                    manifest=manifest,
                    bank_type=_text(manifest.get("detected_bank_type")),
                    client_id=normalize_text(manifest.get("client_hint") or manifest.get("source_raw_filename")),
                    paired_path=paired_path if paired_path and paired_path.is_file() else None,
                    raw_path=raw_path if raw_path and raw_path.is_file() else None,
                    compression_map_path=(
                        compression_path if compression_path and compression_path.is_file() else None
                    ),
                    source_hash=source_hash,
                    raw_hash=raw_hash,
                    paired_hash=paired_hash,
                    compression_map_hash=compression_map_hash,
                    manifest_hash=manifest_hash,
                ))
        else:
            issues.append(DatasetIssue("warning", "imports_directory_missing", str(imports_root)))

        if self.include_legacy and self.legacy_training_dir and self.legacy_training_dir.is_dir():
            for path in sorted(self.legacy_training_dir.rglob("*.xls*"), key=lambda value: str(value).casefold()):
                if not path.is_file() or path.name.startswith("~$"):
                    continue
                relative_parts = {_header_key(part) for part in path.relative_to(self.legacy_training_dir).parts[:-1]}
                if relative_parts & {"codedkeys", "keys", "keymappings"}:
                    continue
                stem = normalize_text(path.stem)
                if any(part in stem for part in SKIP_LEGACY_NAME_PARTS):
                    issues.append(DatasetIssue("skip", "generated_legacy_workbook", str(path)))
                    counts["sources_skipped"] += 1
                    continue
                try:
                    source_hash = _file_sha256(path)
                except OSError as exc:
                    issues.append(DatasetIssue("skip", "source_hash_failed", str(path), detail=str(exc)))
                    counts["sources_skipped"] += 1
                    continue
                candidates.append(TrainingSource(
                    path=path,
                    source_type="legacy",
                    source_id=f"legacy:{path.relative_to(self.legacy_training_dir).as_posix()}",
                    client_id=normalize_text(path.stem),
                    source_hash=source_hash,
                ))
                counts["legacy_workbooks_discovered"] += 1

        sources: list[TrainingSource] = []
        seen_hashes: dict[str, TrainingSource] = {}
        for source in sorted(candidates, key=lambda item: (item.source_type, str(item.path).casefold())):
            duplicate = seen_hashes.get(source.source_hash)
            if duplicate is not None:
                issues.append(DatasetIssue(
                    "skip",
                    "exact_duplicate_workbook",
                    str(source.path),
                    detail=f"same bytes as {duplicate.path}",
                ))
                counts["exact_duplicate_workbooks"] += 1
                counts["sources_skipped"] += 1
                continue
            seen_hashes[source.source_hash] = source
            sources.append(source)
        counts["sources_selected"] = len(sources)
        counts["import_workbooks_selected"] = sum(source.source_type == "import" for source in sources)
        counts["legacy_workbooks_selected"] = sum(source.source_type == "legacy" for source in sources)
        return sources, issues, dict(counts)

    def load(self) -> TrainingDataset:
        sources, issues, discovery = self.discover_sources()
        stats = Counter(discovery)
        pairing_counts = Counter()
        bank_counts = Counter()
        direction_counts = Counter()
        class_counts = Counter()
        rows: list[dict[str, Any]] = []

        for source in sources:
            try:
                original = pd.read_excel(source.path, dtype=object, keep_default_na=False)
            except Exception as exc:
                issues.append(DatasetIssue("skip", "unreadable_coded_workbook", str(source.path), detail=str(exc)))
                stats["workbooks_unreadable"] += 1
                continue
            frame, ambiguous = _canonical_frame(original)
            if ambiguous:
                issues.append(DatasetIssue(
                    "skip",
                    "ambiguous_coded_columns",
                    str(source.path),
                    detail=", ".join(ambiguous),
                ))
                stats["workbooks_invalid"] += 1
                continue
            missing_columns = [column for column in REQUIRED_CODED_COLUMNS if column not in frame.columns]
            if missing_columns:
                issues.append(DatasetIssue(
                    "skip",
                    "missing_coded_columns",
                    str(source.path),
                    detail=", ".join(missing_columns),
                ))
                stats["workbooks_invalid"] += 1
                continue

            stats["workbooks_readable"] += 1
            stats["coded_rows_seen"] += len(frame)
            workbook_controls = self._workbook_control_codes(frame)
            pairings = self._pairings(source, frame, issues)
            if len(pairings) != len(frame):
                pairings = [PairingInfo("pairing_internal_error") for _ in range(len(frame))]
            semantic_occurrences = Counter()
            overlap_occurrences = Counter()

            for position, (_, row) in enumerate(frame.iterrows()):
                excel_row = position + 2
                pair = pairings[position]
                pairing_counts[pair.status] += 1
                if source.source_type == "import" and self._legacy_row_is_usable(row):
                    # Preserve the historical 8,927-row audit definition as a
                    # separately reported baseline.  The new dataset then
                    # applies stricter direction/control-account validation.
                    stats["legacy_compatible_import_rows"] += 1
                direction = _direction(row.get("Positive/Negative"), row.get("AMOUNT"))
                debit = clean_code(row.get("DEBIT"))
                credit = clean_code(row.get("CREDIT"))
                if not direction:
                    if debit in workbook_controls and credit and credit not in workbook_controls:
                        direction = "incoming"
                    elif credit in workbook_controls and debit and debit not in workbook_controls:
                        direction = "outgoing"
                if direction == "outgoing":
                    target_code, counter_code = debit, credit
                elif direction == "incoming":
                    target_code, counter_code = credit, debit
                else:
                    issues.append(DatasetIssue("row_skip", "unknown_direction", str(source.path), excel_row))
                    stats["rows_skipped_unknown_direction"] += 1
                    continue
                if not target_code:
                    issues.append(DatasetIssue("row_skip", "missing_or_invalid_target", str(source.path), excel_row))
                    stats["rows_skipped_invalid_target"] += 1
                    continue
                if not counter_code:
                    issues.append(DatasetIssue("row_skip", "missing_or_invalid_counter_account", str(source.path), excel_row))
                    stats["rows_skipped_invalid_counter"] += 1
                    continue
                if counter_code not in workbook_controls:
                    issues.append(DatasetIssue(
                        "row_skip",
                        "counter_account_not_validated",
                        str(source.path),
                        excel_row,
                        detail=counter_code,
                    ))
                    stats["rows_skipped_unvalidated_counter"] += 1
                    continue
                if target_code in workbook_controls:
                    stats["rows_skipped_control_account_target"] += 1
                    continue

                amount = _amount_decimal(row.get("AMOUNT"))
                date = _date_key(row.get("DATE"))
                if amount is None or not date:
                    reason = "missing_amount" if amount is None else "missing_date"
                    issues.append(DatasetIssue("row_skip", reason, str(source.path), excel_row))
                    stats[f"rows_skipped_{reason}"] += 1
                    continue

                text = _text(row.get("TEXT"))
                raw_explanation, kennitala, reference = self._raw_features(
                    pair,
                    coded_id=row.get("ID"),
                    bank_type=source.bank_type,
                )
                normalized = normalize_text(text)
                family = merchant_family(text)
                weak_text = normalized in WEAK_NORMALIZED_TEXT or not merchant_key(text)
                model_target = (
                    VEHICLE_EXPENSE_FAMILY
                    if target_code in {"4635", "4640"}
                    else target_code
                )
                signed_amount = -abs(amount) if direction == "outgoing" else abs(amount)
                semantic_identity = self._semantic_identity(
                    source=source,
                    position=position,
                    pair=pair,
                    date=date,
                    signed_amount=signed_amount,
                    normalized_text=normalized,
                    reference=reference,
                    kennitala=kennitala,
                )
                if semantic_identity.startswith("ref:"):
                    occurrence = semantic_occurrences[semantic_identity]
                    semantic_occurrences[semantic_identity] += 1
                    semantic_identity = f"{semantic_identity}:occurrence:{occurrence}"
                overlap_identity = self._statement_overlap_identity(
                    source=source,
                    date=date,
                    signed_amount=signed_amount,
                    normalized_text=normalized,
                    kennitala=kennitala,
                )
                if overlap_identity:
                    occurrence = overlap_occurrences[overlap_identity]
                    overlap_occurrences[overlap_identity] += 1
                    overlap_identity = f"{overlap_identity}:occurrence:{occurrence}"
                rows.append({
                    "example_id": f"{source.source_id}:{position}",
                    "semantic_identity": semantic_identity,
                    "statement_overlap_identity": overlap_identity,
                    "source_type": source.source_type,
                    "session_id": source.source_id,
                    "source_workbook": str(source.path),
                    "source_hash": source.source_hash,
                    "row_number": excel_row,
                    "bank_type": source.bank_type or "legacy_unknown",
                    "client_id": source.client_id,
                    "date": date,
                    "text": text,
                    "normalized_text": normalized,
                    "merchant_family": family,
                    "raw_explanation": raw_explanation,
                    "valid_kennitala": kennitala,
                    "transaction_reference": reference,
                    "amount": float(abs(amount)),
                    "signed_amount": float(signed_amount),
                    "direction": direction,
                    "counter_code": counter_code,
                    "target_code": target_code,
                    "model_target": model_target,
                    "raw_pair_status": pair.status,
                    "weak_text": bool(weak_text),
                    "production_eligible": (
                        source.bank_type not in PRODUCTION_EXCLUDED_BANK_TYPES
                    ),
                })
                direction_counts[direction] += 1
                bank_counts[source.bank_type or "legacy_unknown"] += 1
                class_counts[target_code] += 1
                stats["rows_valid_before_dedup"] += 1

        rows, dedup_stats = self._deduplicate(rows, issues)
        stats.update(dedup_stats)
        frame = pd.DataFrame(rows)
        if not frame.empty:
            frame = frame.sort_values(
                ["source_type", "session_id", "source_workbook", "row_number"],
                kind="stable",
            ).reset_index(drop=True)
        stats["usable_rows"] = len(frame)
        stats["production_excluded_rows"] = (
            int((~frame["production_eligible"].astype(bool)).sum())
            if not frame.empty and "production_eligible" in frame.columns
            else 0
        )
        stats["production_model_rows"] = stats["usable_rows"] - stats["production_excluded_rows"]

        # Recompute distributions after defensive deduplication/conflict removal.
        if not frame.empty:
            class_counts = Counter(frame["target_code"].astype(str))
            direction_counts = Counter(frame["direction"].astype(str))
            bank_counts = Counter(frame["bank_type"].astype(str))
        fingerprint = self._dataset_fingerprint(frame, sources)
        report = {
            "dataset_schema_version": DATASET_SCHEMA_VERSION,
            "feature_schema_version": FEATURE_SCHEMA_VERSION,
            "counts": dict(sorted(stats.items())),
            "pairing": dict(sorted(pairing_counts.items())),
            "directions": dict(sorted(direction_counts.items())),
            "banks": dict(sorted(bank_counts.items())),
            "class_count": len(class_counts),
            "class_support": dict(sorted(class_counts.items())),
            "rare_classes": {
                "singletons": sum(value == 1 for value in class_counts.values()),
                "support_below_5": sum(value < 5 for value in class_counts.values()),
                "support_below_10": sum(value < 10 for value in class_counts.values()),
            },
            "sources": [
                {
                    "source_id": source.source_id,
                    "source_type": source.source_type,
                    "path": str(source.path),
                    "bank_type": source.bank_type,
                    "source_hash": source.source_hash,
                }
                for source in sources
            ],
        }
        return TrainingDataset(frame=frame, report=report, fingerprint=fingerprint, issues=issues)

    def _workbook_control_codes(self, frame: pd.DataFrame) -> set[str]:
        controls = set(self.control_account_codes)
        bank_side_counts = Counter()
        for _, row in frame.iterrows():
            direction = _direction(row.get("Positive/Negative"), row.get("AMOUNT"))
            if direction == "outgoing":
                code = clean_code(row.get("CREDIT"))
            elif direction == "incoming":
                code = clean_code(row.get("DEBIT"))
            else:
                continue
            if code:
                bank_side_counts[code] += 1
        if bank_side_counts:
            maximum = bank_side_counts.most_common(1)[0][1]
            controls.update(
                code for code, count in bank_side_counts.items()
                if count >= 2 and (count >= maximum * 0.05 or code in self.control_account_codes)
            )
        return controls

    def _pairings(
        self,
        source: TrainingSource,
        coded: pd.DataFrame,
        issues: list[DatasetIssue],
    ) -> list[PairingInfo]:
        if source.source_type != "import" or source.session_dir is None:
            return [PairingInfo("legacy_coded_only") for _ in range(len(coded))]
        if (source.session_dir / "paired_rows_sync_error.txt").is_file():
            return [PairingInfo("unsafe_sync_error") for _ in range(len(coded))]
        sync_mode = _text(source.manifest.get("paired_rows_sync_mode"))
        if sync_mode == "coded_only_rebuild":
            return [PairingInfo("coded_only_rebuild") for _ in range(len(coded))]
        if source.paired_path is None:
            return [PairingInfo("paired_rows_missing") for _ in range(len(coded))]
        try:
            paired = pd.read_csv(
                source.paired_path,
                dtype=object,
                keep_default_na=False,
                encoding="utf-8-sig",
            )
        except Exception as exc:
            issues.append(DatasetIssue("warning", "paired_rows_unreadable", str(source.paired_path), detail=str(exc)))
            return [PairingInfo("paired_rows_unreadable") for _ in range(len(coded))]

        raw_names = {
            str(key): str(value)
            for key, value in (
                source.manifest.get("columns", {}).get("raw_column_map", {})
                if isinstance(source.manifest.get("columns"), dict)
                else {}
            ).items()
        }
        raw_columns = [column for column in paired.columns if str(column).startswith("raw_")]
        if len(coded) == len(paired) and self._pair_identity_matches(coded, paired):
            output = []
            for position in range(len(coded)):
                raw_values = {str(column): paired.iloc[position].get(column, "") for column in raw_columns}
                has_raw = any(_text(value) for value in raw_values.values())
                output.append(PairingInfo(
                    "validated_pair" if has_raw else "coded_only",
                    raw_values=raw_values if has_raw else {},
                    raw_names=raw_names,
                    raw_index=_text(paired.iloc[position].get("raw_index", "")),
                ))
            return output

        if source.compression_map_path is not None:
            try:
                compression_map = _read_json(source.compression_map_path)
                entries = compression_map.get("rows", [])
                if not isinstance(entries, list) or len(entries) != len(coded):
                    raise ValueError("compression map row count does not match coded workbook")
                coverage: list[int] = []
                normalized_entries: list[tuple[dict[str, Any], list[int]]] = []
                for position, entry in enumerate(entries):
                    if not isinstance(entry, dict) or entry.get("coded_index") != position:
                        raise ValueError("compression map coded row order is invalid")
                    indices = entry.get("paired_indices")
                    if not isinstance(indices, list) or not indices:
                        raise ValueError("compression map contains an empty group")
                    try:
                        integer_indices = [int(value) for value in indices]
                    except (TypeError, ValueError) as exc:
                        raise ValueError("compression map contains a non-numeric paired index") from exc
                    if any(value < 0 or value >= len(paired) for value in integer_indices):
                        raise ValueError("compression map paired index is out of range")
                    identity = entry.get("identity") if isinstance(entry.get("identity"), dict) else {}
                    if not self._map_identity_matches(coded.iloc[position], identity):
                        raise ValueError(f"compression identity mismatch at coded row {position + 1}")
                    coverage.extend(integer_indices)
                    normalized_entries.append((entry, integer_indices))
                if sorted(coverage) != list(range(len(paired))):
                    raise ValueError("compression map does not cover every paired row exactly once")

                output: list[PairingInfo] = []
                for _entry, indices in normalized_entries:
                    if len(indices) != 1:
                        output.append(PairingInfo("ambiguous_compressed_group"))
                        continue
                    paired_row = paired.iloc[indices[0]]
                    coded_position = len(output)
                    if _coded_identity_key(coded.iloc[coded_position]) != _coded_identity_key(paired_row):
                        output.append(PairingInfo("compression_identity_mismatch"))
                        continue
                    raw_values = {str(column): paired_row.get(column, "") for column in raw_columns}
                    has_raw = any(_text(value) for value in raw_values.values())
                    output.append(PairingInfo(
                        "validated_compression_pair" if has_raw else "compression_coded_only",
                        raw_values=raw_values if has_raw else {},
                        raw_names=raw_names,
                        raw_index=_text(paired_row.get("raw_index", "")),
                    ))
                return output
            except Exception as exc:
                issues.append(DatasetIssue("warning", "compression_map_invalid", str(source.path), detail=str(exc)))

        # A stale row count/order is never paired positionally.  Recover only
        # identities that are unique in both current coded rows and the paired
        # archive; ambiguous or unmatched rows remain valid coded-only labels.
        return self._identity_aligned_pairings(
            coded,
            paired,
            raw_columns=raw_columns,
            raw_names=raw_names,
        )

    @staticmethod
    def _pair_identity_matches(coded: pd.DataFrame, paired: pd.DataFrame) -> bool:
        if len(coded) != len(paired):
            return False
        coded_keys = [_coded_identity_key(coded.iloc[position]) for position in range(len(coded))]
        paired_keys = [_coded_identity_key(paired.iloc[position]) for position in range(len(paired))]
        # Repeated identical transactions cannot be safely associated with
        # different raw-only explanations by row position alone.
        if len(set(coded_keys)) != len(coded_keys) or len(set(paired_keys)) != len(paired_keys):
            return False
        for position in range(len(coded)):
            if coded_keys[position] != paired_keys[position]:
                return False
        return bool(len(coded))

    @staticmethod
    def _identity_aligned_pairings(
        coded: pd.DataFrame,
        paired: pd.DataFrame,
        *,
        raw_columns: list[Any],
        raw_names: dict[str, str],
    ) -> list[PairingInfo]:
        coded_by_key: dict[tuple[str, ...], list[int]] = defaultdict(list)
        paired_by_key: dict[tuple[str, ...], list[int]] = defaultdict(list)
        for position in range(len(coded)):
            coded_by_key[_coded_identity_key(coded.iloc[position])].append(position)
        for position in range(len(paired)):
            paired_by_key[_coded_identity_key(paired.iloc[position])].append(position)
        output = [PairingInfo("identity_unmatched") for _ in range(len(coded))]
        for key, coded_positions in coded_by_key.items():
            paired_positions = paired_by_key.get(key, [])
            # Empty/generic identities or repeated same-value transactions are
            # deliberately not guessed by occurrence or row position.
            if not all(key[:4]) or len(coded_positions) != 1 or len(paired_positions) != 1:
                status = "identity_ambiguous" if paired_positions else "identity_unmatched"
                for coded_position in coded_positions:
                    output[coded_position] = PairingInfo(status)
                continue
            coded_position = coded_positions[0]
            paired_row = paired.iloc[paired_positions[0]]
            raw_values = {str(column): paired_row.get(column, "") for column in raw_columns}
            has_raw = any(_text(value) for value in raw_values.values())
            output[coded_position] = PairingInfo(
                "validated_identity_pair" if has_raw else "identity_coded_only",
                raw_values=raw_values if has_raw else {},
                raw_names=raw_names,
                raw_index=_text(paired_row.get("raw_index", "")),
            )
        return output

    @staticmethod
    def _map_identity_matches(coded_row: pd.Series, identity: dict[str, Any]) -> bool:
        compared = 0
        for column, normalizer in (
            ("DATE", _date_key),
            ("AMOUNT", _amount_key),
            ("Positive/Negative", lambda value: _text(value)),
            ("TEXT", normalize_text),
            ("ID", _identity_id),
        ):
            if column not in identity:
                continue
            compared += 1
            if normalizer(coded_row.get(column)) != normalizer(identity.get(column)):
                return False
        return compared > 0

    @staticmethod
    def _raw_features(
        pair: PairingInfo,
        *,
        coded_id: Any,
        bank_type: str = "",
    ) -> tuple[str, str, str]:
        kennitala = normalize_kennitala(coded_id)
        explanation_parts: list[str] = []
        references: list[tuple[int, str]] = []
        kennitala_candidates: list[tuple[int, str, str]] = []
        for raw_column, value in pair.raw_values.items():
            value_text = _text(value)
            if not value_text:
                continue
            original_name = pair.raw_names.get(raw_column, raw_column)
            header = _header_key(original_name)
            if header in LEAKAGE_HEADER_TOKENS:
                continue
            if any(token in header for token in RAW_TEXT_HEADER_TOKENS):
                normalized_value = value_text.casefold()
                if normalized_value not in {part.casefold() for part in explanation_parts}:
                    explanation_parts.append(value_text)
            for priority, token in enumerate(REFERENCE_HEADER_TOKENS):
                if token in header:
                    reference = _canonical_reference(value)
                    if reference:
                        references.append((priority, f"{token}:{reference}"))
                    break
            if any(token in header for token in KENNITALA_EXCLUDED_HEADER_TOKENS):
                continue
            candidate_kennitala = normalize_kennitala(value)
            if candidate_kennitala and any(token in header for token in KENNITALA_HEADER_TOKENS):
                preferred = any(token in header for token in KENNITALA_PREFERRED_HEADER_TOKENS)
                # Landsbankinn card data is the known dual-ID format.  The
                # explicit merchant field wins regardless of workbook order;
                # the same ordering is conservative for every other format.
                priority = 0 if preferred else (2 if bank_type == "landsbankinn_kort" else 1)
                kennitala_candidates.append((priority, header, candidate_kennitala))
        if not kennitala and kennitala_candidates:
            kennitala_candidates.sort(key=lambda item: (item[0], item[1], item[2]))
            kennitala = kennitala_candidates[0][2]
        references.sort(key=lambda item: (item[0], item[1]))
        reference = references[0][1] if references else ""
        return " | ".join(explanation_parts), kennitala, reference

    @staticmethod
    def _semantic_identity(
        *,
        source: TrainingSource,
        position: int,
        pair: PairingInfo,
        date: str,
        signed_amount: Decimal,
        normalized_text: str,
        reference: str,
        kennitala: str,
    ) -> str:
        amount = format(signed_amount.quantize(Decimal("0.01")), "f")
        if reference:
            material = "|".join((
                source.bank_type,
                source.client_id,
                date,
                amount,
                normalized_text,
                kennitala,
                reference,
            ))
            return "ref:" + sha256(material.encode("utf-8")).hexdigest()
        if source.raw_hash and pair.raw_index and pair.status.startswith("validated"):
            return f"raw:{source.raw_hash}:{pair.raw_index}"
        return f"source:{source.source_hash}:{position}"

    @staticmethod
    def _statement_overlap_identity(
        *,
        source: TrainingSource,
        date: str,
        signed_amount: Decimal,
        normalized_text: str,
        kennitala: str,
    ) -> str:
        if not source.client_id or not kennitala or not normalized_text:
            return ""
        amount = format(signed_amount.quantize(Decimal("0.01")), "f")
        material = "|".join((
            source.bank_type,
            source.client_id,
            date,
            amount,
            normalized_text,
            kennitala,
        ))
        return "overlap:" + sha256(material.encode("utf-8")).hexdigest()

    @staticmethod
    def _deduplicate(
        rows: list[dict[str, Any]],
        issues: list[DatasetIssue],
    ) -> tuple[list[dict[str, Any]], Counter]:
        groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for row in rows:
            groups[str(row["semantic_identity"])].append(row)
        primary_kept: list[dict[str, Any]] = []
        stats = Counter()
        for identity in sorted(groups):
            group = groups[identity]
            if len(group) == 1:
                primary_kept.append(group[0])
                continue
            targets = {str(row["target_code"]) for row in group}
            if len(targets) > 1:
                stats["duplicate_conflict_rows_removed"] += len(group)
                issues.append(DatasetIssue(
                    "row_skip",
                    "duplicate_label_conflict",
                    group[0]["source_workbook"],
                    int(group[0]["row_number"]),
                    detail=f"identity {identity[:28]} has targets {sorted(targets)}",
                ))
                continue
            primary_kept.append(group[0])
            stats["duplicate_rows_removed"] += len(group) - 1

        # References can legitimately distinguish repeated payments, but some
        # bank exports regenerate references when an overlapping statement is
        # downloaded again.  Deduplicate weaker identities only between pairs
        # of sessions that demonstrably overlap in bulk.  A single coincidental
        # same-day payment can never trigger this rule.
        session_sizes = Counter(str(row["session_id"]) for row in primary_kept)
        overlap_groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for row in primary_kept:
            overlap_identity = str(row.get("statement_overlap_identity", ""))
            if overlap_identity:
                overlap_groups[overlap_identity].append(row)
        pair_shared = Counter()
        for group in overlap_groups.values():
            sessions = sorted({str(row["session_id"]) for row in group})
            for first_index in range(len(sessions)):
                for second_index in range(first_index + 1, len(sessions)):
                    pair_shared[(sessions[first_index], sessions[second_index])] += 1
        overlapping_session_pairs = {
            pair
            for pair, shared in pair_shared.items()
            if shared >= 5
            and shared / max(1, min(session_sizes[pair[0]], session_sizes[pair[1]])) >= 0.25
        }

        remove_ids: set[str] = set()
        for identity, group in overlap_groups.items():
            if len(group) < 2:
                continue
            group_sessions = {str(row["session_id"]) for row in group}
            group_targets = {str(row["target_code"]) for row in group}
            if len(group_sessions) > 1 and len(group_targets) > 1:
                for row in group:
                    remove_ids.add(str(row["example_id"]))
                stats["duplicate_conflict_rows_removed"] += len(group)
                issues.append(DatasetIssue(
                    "row_skip",
                    "duplicate_label_conflict",
                    group[0]["source_workbook"],
                    int(group[0]["row_number"]),
                    detail=f"overlap identity {identity[:28]} has targets {sorted(group_targets)}",
                ))
                continue
            eligible_rows = []
            for row in group:
                session = str(row["session_id"])
                if any(
                    session in pair
                    and any(str(other["session_id"]) == (pair[1] if session == pair[0] else pair[0]) for other in group)
                    for pair in overlapping_session_pairs
                ):
                    eligible_rows.append(row)
            if len(eligible_rows) < 2:
                continue
            preferred = max(
                eligible_rows,
                key=lambda row: (
                    str(row.get("raw_pair_status", "")).startswith("validated"),
                    bool(row.get("raw_explanation")),
                    str(row["session_id"]),
                ),
            )
            for row in eligible_rows:
                if row is not preferred:
                    remove_ids.add(str(row["example_id"]))
                    stats["duplicate_rows_removed"] += 1

        kept = [row for row in primary_kept if str(row["example_id"]) not in remove_ids]
        return kept, stats

    def _legacy_row_is_usable(self, row: pd.Series) -> bool:
        debit = clean_code(row.get("DEBIT"))
        credit = clean_code(row.get("CREDIT"))
        sign = _text(row.get("Positive/Negative"))
        if not debit or debit == "7810":
            return False
        if sign != "-" and not (not sign and credit in COUNTER_ACCOUNT_CODES):
            return False
        key = merchant_key(row.get("TEXT"))
        return bool(key) and not any(pattern.match(key) for pattern in LEGACY_GENERIC_PATTERNS)

    @staticmethod
    def _dataset_fingerprint(frame: pd.DataFrame, sources: list[TrainingSource]) -> str:
        digest = sha256()
        digest.update(DATASET_SCHEMA_VERSION.encode("ascii"))
        for source in sorted(sources, key=lambda item: (item.source_id, item.source_hash)):
            source_material = (
                source.source_id,
                source.source_type,
                source.source_hash,
                source.raw_hash,
                source.paired_hash,
                source.compression_map_hash,
                source.manifest_hash,
                source.bank_type,
                source.client_id,
            )
            digest.update(("|".join(source_material) + "\n").encode("utf-8"))
        if not frame.empty:
            for row in frame.itertuples(index=False):
                material = "|".join((
                    str(row.semantic_identity),
                    str(row.model_target),
                    str(row.target_code),
                    str(row.bank_type),
                    str(row.direction),
                    str(row.date),
                    f"{float(row.signed_amount):.2f}",
                    str(row.normalized_text),
                    normalize_text(row.raw_explanation),
                    str(row.valid_kennitala),
                    str(row.transaction_reference),
                    str(row.raw_pair_status),
                    "1" if bool(row.weak_text) else "0",
                ))
                digest.update((material + "\n").encode("utf-8"))
        return digest.hexdigest()


def load_training_dataset(
    *,
    training_data_dir: str | Path,
    legacy_training_dir: str | Path | None = None,
    include_legacy: bool = True,
) -> TrainingDataset:
    return TrainingDataLoader(
        training_data_dir=training_data_dir,
        legacy_training_dir=legacy_training_dir,
        include_legacy=include_legacy,
    ).load()


def build_runtime_feature_record(
    coded_row: Any,
    *,
    source_row: Any = None,
    bank_type: str = "",
) -> dict[str, Any]:
    """Build the same missing-safe feature record used by the trainer.

    ``source_row`` is optional because rebuilt archives can legitimately contain
    only coded output.  Header filtering in :meth:`_raw_features` also keeps
    confidence/status artifacts out of the model input.
    """
    try:
        raw_items = list(source_row.items()) if source_row is not None else []
    except Exception:
        raw_items = []
    raw_values = {str(column): value for column, value in raw_items}
    pair = PairingInfo(
        status="runtime" if raw_values else "runtime_coded_only",
        raw_values=raw_values,
        raw_names={column: column for column in raw_values},
    )
    getter = coded_row.get if hasattr(coded_row, "get") else lambda key, default=None: default
    raw_explanation, kennitala, reference = TrainingDataLoader._raw_features(
        pair,
        coded_id=getter("ID", ""),
        bank_type=str(bank_type or ""),
    )
    text = _text(getter("TEXT", ""))
    normalized = normalize_text(text)
    amount = _amount_decimal(getter("AMOUNT", ""))
    direction = _direction(getter("Positive/Negative", ""), amount)
    return {
        "bank_type": str(bank_type or "") or "unknown",
        "date": _date_key(getter("DATE", "")),
        "text": text,
        "normalized_text": normalized,
        "merchant_family": merchant_family(text),
        "raw_explanation": raw_explanation,
        "valid_kennitala": kennitala,
        "transaction_reference": reference,
        "amount": float(abs(amount)) if amount is not None else 0.0,
        "signed_amount": (
            float(-abs(amount)) if amount is not None and direction == "outgoing"
            else float(abs(amount)) if amount is not None
            else 0.0
        ),
        "direction": direction or "unknown",
        "weak_text": bool(normalized in WEAK_NORMALIZED_TEXT or not merchant_key(text)),
        "raw_pair_status": pair.status,
    }


__all__ = [
    "DATASET_SCHEMA_VERSION",
    "BANK_SIDE_ACCOUNT_CODES",
    "DEFAULT_CONTROL_ACCOUNT_CODES",
    "FEATURE_SCHEMA_VERSION",
    "NON_CLASSIFICATION_ACCOUNT_CODES",
    "PRODUCTION_EXCLUDED_BANK_TYPES",
    "TrainingDataLoader",
    "TrainingDataset",
    "VEHICLE_EXPENSE_FAMILY",
    "build_runtime_feature_record",
    "load_training_dataset",
    "normalize_kennitala",
]

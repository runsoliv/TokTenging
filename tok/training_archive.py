"""Training-data archive helpers for raw/coded bank import pairs."""

from __future__ import annotations

import datetime as _datetime
import json
import os
from pathlib import Path
import re
import shutil
import unicodedata

import pandas as pd


SCHEMA_VERSION = "1.0"
TRAINING_RAW_INDEX_COLUMN = "_TRAINING_RAW_INDEX"

CODED_FIELDS = ["DATE", "TEXT", "ID", "AMOUNT", "DEBIT", "CREDIT", "Positive/Negative"]
AUTO_CODE_FIELDS = [
    "STATUS",
    "CONFIDENCE",
    "AUTO_DEBIT_STATUS",
    "AUTO_DEBIT_CONFIDENCE",
    "AUTO_DEBIT_REASON",
    "_AUTO_CODE_SOURCE",
]
CODE_LABEL_FIELDS = ["DEBIT", "CREDIT"]
CODED_COLUMN_ALIASES = {
    "CONFIDENCE": ("Cnf",),
}
COMPRESSED_SYNC_FIELDS = ["DEBIT", "CREDIT", "STATUS", "CONFIDENCE"]
SYNC_ERROR_FILENAME = "paired_rows_sync_error.txt"


def sync_all_training_archives(training_data_dir: str | os.PathLike) -> dict:
    """Synchronize every valid import session under a training archive."""
    imports_root = Path(training_data_dir) / "imports"
    if not imports_root.is_dir():
        return {
            "ok": False,
            "total": 0,
            "synced": 0,
            "failed": 0,
            "failures": [],
            "message": f"Mappan „imports“ fannst ekki: {imports_root}",
        }

    session_dirs = sorted({path.parent for path in imports_root.rglob("manifest.json")})
    synced = 0
    failures = []
    synced_rows = 0

    for session_dir in session_dirs:
        coded_path = _find_session_coded_workbook(session_dir)
        result = refresh_paired_rows_from_coded(
            session_dir=session_dir,
            coded_file_path=coded_path,
        )
        if result.get("ok"):
            synced += 1
            synced_rows += int(result.get("rows", 0))
        else:
            failures.append({
                "session": session_dir.name,
                "message": result.get("message", "Óþekkt villa kom upp við samstillingu."),
                "error_file": result.get("error_file", ""),
            })

    failed = len(failures)
    total = len(session_dirs)
    return {
        "ok": failed == 0,
        "total": total,
        "synced": synced,
        "failed": failed,
        "synced_rows": synced_rows,
        "failures": failures,
        "message": f"Samstillti {synced} af {total} lotum í safni; {failed} mistókust.",
    }


def refresh_paired_rows_from_coded(
    *,
    session_dir: str | os.PathLike,
    coded_file_path: str | os.PathLike | None = None,
) -> dict:
    """Refresh ML labels from the final, user-edited coded workbook."""
    session_path = Path(session_dir)
    manifest_path = session_path / "manifest.json"
    paired_path = session_path / "paired_rows.csv"
    coded_path = Path(coded_file_path) if coded_file_path else None

    def fail(message):
        return _record_sync_failure(session_path, coded_path, message)

    if not manifest_path.exists():
        return fail("Stýriskrá þjálfunargagna fannst ekki.")
    if not paired_path.exists():
        return fail("Skrá með pöruðum línum fannst ekki.")

    try:
        with open(manifest_path, "r", encoding="utf-8") as handle:
            manifest = json.load(handle)
    except Exception as exc:
        return fail(f"Ekki tókst að lesa stýriskrá þjálfunargagna: {exc}")

    coded_path = coded_path or _manifest_coded_path(session_path, manifest)
    if not coded_path.exists():
        return fail("Kóðuð vinnubók fannst ekki.")

    try:
        coded_df = pd.read_excel(coded_path, dtype=object, keep_default_na=False)
        paired_df = pd.read_csv(paired_path, dtype=object, keep_default_na=False, encoding="utf-8-sig")
    except Exception as exc:
        return fail(f"Ekki tókst að lesa þjálfunarskrár: {exc}")

    sync_mode = "direct"
    sync_warning = ""
    compressed_sync = len(coded_df) != len(paired_df)
    if compressed_sync:
        sync_result = _apply_compressed_codes(session_path, coded_df, paired_df)
        if not sync_result["ok"]:
            sync_warning = sync_result["message"]
            paired_df = _rebuild_paired_rows_from_coded(coded_df, paired_df, manifest)
            sync_mode = "coded_only_rebuild"
        else:
            paired_df = sync_result["paired_df"]
            sync_mode = "compressed_map"
    else:
        mismatch = _first_coded_identity_mismatch(coded_df, paired_df)
        if mismatch:
            sync_warning = (
                f"Færslulína {mismatch} samsvarar ekki lengur þjálfunarlínunni."
            )
            paired_df = _rebuild_paired_rows_from_coded(coded_df, paired_df, manifest)
            sync_mode = "coded_only_rebuild"
        else:
            for field in CODED_FIELDS:
                coded_column = _coded_column_name(coded_df, field)
                if coded_column:
                    paired_df[field] = [_clean_value(value) for value in coded_df[coded_column]]
            for field in AUTO_CODE_FIELDS:
                coded_column = _coded_column_name(coded_df, field)
                if coded_column:
                    paired_df[field] = [_clean_value(value) for value in coded_df[coded_column]]
            paired_df["coded_index"] = list(range(len(coded_df)))

    synced_at = _datetime.datetime.now().isoformat(timespec="seconds")
    manifest["output_coded_filename"] = coded_path.name
    manifest.setdefault("archive_files", {})["coded"] = coded_path.name
    manifest.setdefault("row_counts", {})["coded"] = int(len(coded_df))
    manifest["row_counts"]["paired"] = int(len(paired_df))
    manifest.setdefault("columns", {})["coded"] = [str(column) for column in coded_df.columns]
    manifest["columns"]["paired"] = [str(column) for column in paired_df.columns]
    manifest["paired_rows_last_synced_at"] = synced_at
    manifest["paired_rows_sync_source"] = coded_path.name
    manifest["paired_rows_sync_mode"] = sync_mode
    if sync_warning:
        manifest["paired_rows_sync_warning"] = sync_warning
    else:
        manifest.pop("paired_rows_sync_warning", None)

    paired_temp = paired_path.with_suffix(".csv.tmp")
    manifest_temp = manifest_path.with_suffix(".json.tmp")
    try:
        paired_df.to_csv(paired_temp, index=False, encoding="utf-8-sig")
        with open(manifest_temp, "w", encoding="utf-8") as handle:
            json.dump(manifest, handle, indent=2, ensure_ascii=False)
        os.replace(paired_temp, paired_path)
        os.replace(manifest_temp, manifest_path)
    except Exception as exc:
        for temp_path in (paired_temp, manifest_temp):
            try:
                temp_path.unlink(missing_ok=True)
            except Exception:
                pass
        return fail(f"Ekki tókst að uppfæra safn þjálfunargagna: {exc}")

    _clear_sync_failure(session_path)
    return {
        "ok": True,
        "message": (
            f"Þjálfunarlínur voru endurgerðar úr núverandi kóðuðu vinnubók {coded_path.name}."
            if sync_mode == "coded_only_rebuild"
            else (
                f"Þjálfunarlínur voru uppfærðar úr þjöppuðu skránni {coded_path.name}."
                if sync_mode == "compressed_map"
                else f"Þjálfunarlínur voru uppfærðar úr {coded_path.name}."
            )
        ),
        "rows": int(len(paired_df)),
        "synced_at": synced_at,
        "sync_mode": sync_mode,
        "warning": sync_warning,
    }


def _record_sync_failure(session_dir: Path, coded_path: Path | None, message: str) -> dict:
    error_path = session_dir / SYNC_ERROR_FILENAME
    timestamp = _datetime.datetime.now().isoformat(timespec="seconds")
    contents = (
        "Samstilling paraðra lína mistókst.\n"
        f"Tími: {timestamp}\n"
        f"Kóðuð skrá: {coded_path.name if coded_path else 'óþekkt'}\n"
        f"Ástæða: {message}\n"
        "paired_rows.csv var látin óbreytt.\n"
    )
    try:
        session_dir.mkdir(parents=True, exist_ok=True)
        error_path.write_text(contents, encoding="utf-8")
    except Exception:
        pass
    return {"ok": False, "message": message, "error_file": str(error_path)}


def _clear_sync_failure(session_dir: Path) -> None:
    try:
        (session_dir / SYNC_ERROR_FILENAME).unlink(missing_ok=True)
    except Exception:
        pass


def _find_session_coded_workbook(session_dir: Path) -> Path | None:
    manifest_path = session_dir / "manifest.json"
    manifest = {}
    try:
        with open(manifest_path, "r", encoding="utf-8") as handle:
            manifest = json.load(handle)
    except Exception:
        pass

    named_candidates = [
        manifest.get("archive_files", {}).get("coded"),
        manifest.get("output_coded_filename"),
    ]
    for name in named_candidates:
        if not name:
            continue
        candidate = session_dir / Path(str(name)).name
        if candidate.is_file():
            return candidate

    candidates = [
        path
        for path in session_dir.glob("*.xlsx")
        if path.name.lower() != "raw.xlsx"
        and "compressed_txs" not in path.stem.lower()
    ]
    if not candidates:
        return None
    coded_named = [path for path in candidates if path.stem.lower().endswith("_coded")]
    pool = coded_named or candidates
    return max(pool, key=lambda path: path.stat().st_mtime)


def write_compression_map(
    *,
    session_dir: str | os.PathLike,
    coded_file_path: str | os.PathLike,
    compressed_df: pd.DataFrame,
    paired_index_groups: list[list[int]],
) -> dict:
    """Persist how each compressed workbook row maps to original paired rows."""
    session_path = Path(session_dir)
    if len(compressed_df) != len(paired_index_groups):
        return {"ok": False, "message": "Línufjöldi í þjöppunarkorti samsvarar ekki úttakslínum."}

    rows = []
    for position, (_, row) in enumerate(compressed_df.iterrows()):
        rows.append({
            "coded_index": position,
            "paired_indices": [int(index) for index in paired_index_groups[position]],
            "identity": {
                field: _clean_value(row.get(field, ""))
                for field in ("DATE", "TEXT", "ID", "AMOUNT", "Positive/Negative")
                if field in compressed_df.columns
            },
        })

    payload = {
        "schema_version": SCHEMA_VERSION,
        "created_at": _datetime.datetime.now().isoformat(timespec="seconds"),
        "coded_filename": Path(coded_file_path).name,
        "source_row_count": sum(len(group) for group in paired_index_groups),
        "compressed_row_count": int(len(compressed_df)),
        "rows": rows,
    }
    map_path = session_path / "compression_map.json"
    temp_path = map_path.with_suffix(".json.tmp")
    try:
        with open(temp_path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, ensure_ascii=False)
        os.replace(temp_path, map_path)
    except Exception as exc:
        try:
            temp_path.unlink(missing_ok=True)
        except Exception:
            pass
        return {"ok": False, "message": f"Ekki tókst að vista þjöppunarkort: {exc}"}

    manifest_path = session_path / "manifest.json"
    if not manifest_path.exists():
        return {
            "ok": True,
            "path": str(map_path),
            "warning": "Þjöppunarkort var vistað en manifest.json fannst ekki.",
        }
    try:
        with open(manifest_path, "r", encoding="utf-8") as handle:
            manifest = json.load(handle)
        manifest.setdefault("archive_files", {})["compression_map"] = map_path.name
        manifest["compression"] = {
            "created_at": payload["created_at"],
            "source_row_count": payload["source_row_count"],
            "compressed_row_count": payload["compressed_row_count"],
        }
        manifest_temp = manifest_path.with_suffix(".json.tmp")
        with open(manifest_temp, "w", encoding="utf-8") as handle:
            json.dump(manifest, handle, indent=2, ensure_ascii=False)
        os.replace(manifest_temp, manifest_path)
    except Exception as exc:
        return {
            "ok": True,
            "path": str(map_path),
            "warning": f"Þjöppunarkort var vistað en lýsigögn í manifest.json voru ekki uppfærð: {exc}",
        }

    return {"ok": True, "path": str(map_path)}


def default_training_data_dir(base_dir: str | os.PathLike | None = None) -> str:
    base = Path(base_dir) if base_dir else Path(__file__).resolve().parent
    return str(base / "trainingData")


def create_import_session(
    *,
    raw_file_path: str,
    bank_type: str,
    client_hint: str = "",
    training_data_dir: str | os.PathLike | None = None,
    timestamp: _datetime.datetime | None = None,
) -> str:
    root = Path(training_data_dir or default_training_data_dir())
    imports_root = root / "imports"
    imports_root.mkdir(parents=True, exist_ok=True)
    session_dir = _unique_session_dir(
        imports_root,
        timestamp or _datetime.datetime.now(),
        client_hint or Path(raw_file_path).stem,
        bank_type,
    )
    session_dir.mkdir(parents=True, exist_ok=False)
    return str(session_dir)


def coded_output_path(session_dir: str | os.PathLike) -> str:
    return str(Path(session_dir) / "coded.xlsx")


def archive_bank_import(
    *,
    raw_file_path: str,
    raw_df: pd.DataFrame,
    coded_df: pd.DataFrame,
    output_file_path: str,
    bank_type: str,
    bank_name: str = "",
    client_hint: str = "",
    training_data_dir: str | os.PathLike | None = None,
    session_dir: str | os.PathLike | None = None,
    app_version: str = "",
) -> dict:
    """Create a per-import ML archive folder for a processed bank file."""
    timestamp = _datetime.datetime.now()
    if session_dir:
        session_dir = Path(session_dir)
        session_dir.mkdir(parents=True, exist_ok=True)
    else:
        session_dir = Path(create_import_session(
            raw_file_path=raw_file_path,
            bank_type=bank_type or bank_name,
            client_hint=client_hint or Path(raw_file_path).stem,
            training_data_dir=training_data_dir,
            timestamp=timestamp,
        ))

    raw_archive_name = _archive_raw_file(raw_file_path, session_dir)
    coded_archive_name = Path(output_file_path).name or "coded.xlsx"
    coded_target = session_dir / coded_archive_name
    if not _same_file(output_file_path, coded_target):
        shutil.copy2(output_file_path, coded_target)

    paired_df, raw_column_map = build_paired_rows(
        raw_df=raw_df,
        coded_df=coded_df,
        bank_type=bank_type,
        bank_name=bank_name,
        client_hint=client_hint or Path(raw_file_path).stem,
    )
    paired_archive_name = "paired_rows.csv"
    paired_df.to_csv(session_dir / paired_archive_name, index=False, encoding="utf-8-sig")

    manifest = {
        "schema_version": SCHEMA_VERSION,
        "import_timestamp": timestamp.isoformat(timespec="seconds"),
        "source_raw_filename": Path(raw_file_path).name,
        "output_coded_filename": Path(output_file_path).name,
        "archive_files": {
            "raw": raw_archive_name,
            "coded": coded_archive_name,
            "paired_rows": paired_archive_name,
            "manifest": "manifest.json",
        },
        "detected_bank_type": bank_type,
        "detected_bank_name": bank_name,
        "client_hint": client_hint or Path(raw_file_path).stem,
        "app_version": app_version or "",
        "row_counts": {
            "raw": int(len(raw_df)) if raw_df is not None else 0,
            "coded": int(len(coded_df)) if coded_df is not None else 0,
            "paired": int(len(paired_df)),
        },
        "columns": {
            "coded": [str(column) for column in coded_df.columns] if coded_df is not None else [],
            "raw": [str(column) for column in raw_df.columns] if raw_df is not None else [],
            "paired": [str(column) for column in paired_df.columns],
            "raw_column_map": raw_column_map,
        },
    }
    with open(session_dir / "manifest.json", "w", encoding="utf-8") as handle:
        json.dump(manifest, handle, indent=2, ensure_ascii=False)

    return {
        "session_dir": str(session_dir),
        "manifest_path": str(session_dir / "manifest.json"),
        "paired_rows_path": str(session_dir / paired_archive_name),
        "raw_path": str(session_dir / raw_archive_name),
        "coded_path": str(session_dir / coded_archive_name),
    }


def build_paired_rows(
    *,
    raw_df: pd.DataFrame,
    coded_df: pd.DataFrame,
    bank_type: str,
    bank_name: str = "",
    client_hint: str = "",
) -> tuple[pd.DataFrame, dict[str, str]]:
    raw_df = raw_df if raw_df is not None else pd.DataFrame()
    coded_df = coded_df if coded_df is not None else pd.DataFrame()
    raw_columns, raw_column_map = _prefixed_raw_columns(raw_df.columns)
    rows = []

    for coded_position, (_, coded_row) in enumerate(coded_df.iterrows()):
        raw_index = _raw_index_for_coded_row(coded_row, coded_position)
        raw_row = _raw_row_by_index(raw_df, raw_index)
        record = {
            "raw_index": _clean_value(raw_index),
            "coded_index": coded_position,
            "bank_type": bank_type or "",
            "bank_name": bank_name or "",
            "source_file_hint": client_hint or "",
            "client_hint": client_hint or "",
        }
        for field in CODED_FIELDS:
            record[field] = _clean_value(coded_row.get(field, ""))
        for field in AUTO_CODE_FIELDS:
            if field in coded_df.columns:
                record[field] = _clean_value(coded_row.get(field, ""))
        for original_column, output_column in raw_columns:
            record[output_column] = _clean_value(raw_row.get(original_column, "") if raw_row is not None else "")
        rows.append(record)

    return pd.DataFrame(rows), raw_column_map


def _archive_raw_file(raw_file_path: str, session_dir: Path) -> str:
    source = Path(raw_file_path)
    suffix = source.suffix.lower() if source.suffix else ".xlsx"
    raw_archive_name = f"raw{suffix}"
    target = session_dir / raw_archive_name
    if not _same_file(source, target):
        shutil.copy2(source, target)
    return raw_archive_name


def _rebuild_paired_rows_from_coded(coded_df: pd.DataFrame, paired_df: pd.DataFrame, manifest: dict) -> pd.DataFrame:
    """Create current training rows from the coded workbook when raw pairing is stale.

    This is a fallback for user-edited workbooks whose row count/order no longer
    matches the archived raw pairing. The raw_* columns are intentionally left
    blank because we cannot safely reconstruct those relationships, but the
    coded labels remain usable for auto-coding.
    """
    base_columns = list(paired_df.columns) if paired_df is not None and not paired_df.empty else []
    required_columns = [
        "raw_index",
        "coded_index",
        "bank_type",
        "bank_name",
        "source_file_hint",
        "client_hint",
        *CODED_FIELDS,
        *AUTO_CODE_FIELDS,
    ]
    columns = []
    for column in [*base_columns, *required_columns]:
        if column not in columns:
            columns.append(column)

    first_existing = paired_df.iloc[0] if paired_df is not None and not paired_df.empty else {}

    def existing_or_manifest(field: str, manifest_key: str = ""):
        try:
            value = first_existing.get(field, "")
        except Exception:
            value = ""
        if not _is_missing(value) and str(value).strip():
            return _clean_value(value)
        if manifest_key:
            return _clean_value(manifest.get(manifest_key, ""))
        return ""

    bank_type = existing_or_manifest("bank_type", "detected_bank_type")
    bank_name = existing_or_manifest("bank_name", "detected_bank_name")
    client_hint = existing_or_manifest("client_hint", "client_hint")
    source_file_hint = existing_or_manifest("source_file_hint", "source_raw_filename") or client_hint

    rows = []
    for coded_position, (_, coded_row) in enumerate(coded_df.iterrows()):
        record = {column: "" for column in columns}
        record["raw_index"] = ""
        record["coded_index"] = coded_position
        record["bank_type"] = bank_type
        record["bank_name"] = bank_name
        record["source_file_hint"] = source_file_hint
        record["client_hint"] = client_hint
        for field in CODED_FIELDS:
            coded_column = _coded_column_name(coded_df, field)
            if coded_column:
                record[field] = _clean_value(coded_row.get(coded_column, ""))
        for field in AUTO_CODE_FIELDS:
            coded_column = _coded_column_name(coded_df, field)
            if coded_column:
                record[field] = _clean_value(coded_row.get(coded_column, ""))
        rows.append(record)

    return pd.DataFrame(rows, columns=columns)


def _manifest_coded_path(session_dir: Path, manifest: dict) -> Path:
    archive_name = manifest.get("archive_files", {}).get("coded")
    return session_dir / (archive_name or "coded.xlsx")


def _first_coded_identity_mismatch(coded_df: pd.DataFrame, paired_df: pd.DataFrame) -> int | None:
    # The coded workbook is intentionally user-editable after export. TEXT, ID,
    # DEBIT/CREDIT, status, and confidence may be changed by the user before
    # compression or before leaving the success screen. Use only stable
    # transaction identity fields here so those edits can be synced back.
    identity_fields = [
        field
        for field in ("DATE", "AMOUNT", "Positive/Negative")
        if field in coded_df.columns and field in paired_df.columns
    ]
    for position in range(len(coded_df)):
        for field in identity_fields:
            coded_value = _comparison_value(coded_df.iloc[position][field])
            paired_value = _comparison_value(paired_df.iloc[position][field])
            if coded_value != paired_value:
                return position + 1
    return None


def _apply_compressed_codes(session_dir: Path, coded_df: pd.DataFrame, paired_df: pd.DataFrame) -> dict:
    map_path = session_dir / "compression_map.json"
    if not map_path.exists():
        return {
            "ok": False,
            "message": (
                f"Línufjöldi breyttist ({len(paired_df)} paraðar, {len(coded_df)} kóðaðar) "
                "og ekkert þjöppunarkort fannst. Pöruðum línum var ekki breytt."
            ),
        }

    try:
        with open(map_path, "r", encoding="utf-8") as handle:
            compression_map = json.load(handle)
    except Exception as exc:
        return {"ok": False, "message": f"Ekki tókst að lesa þjöppunarkort: {exc}"}

    mapped_rows = compression_map.get("rows", [])
    if len(mapped_rows) != len(coded_df):
        return {
            "ok": False,
            "message": "Þjöppuð vinnubók samsvarar ekki lengur þjöppunarkortinu.",
        }

    covered_indices = []
    for position, entry in enumerate(mapped_rows):
        if entry.get("coded_index") != position:
            return {"ok": False, "message": "Röð lína í þjöppunarkorti er ógild."}
        paired_indices = entry.get("paired_indices", [])
        if not paired_indices:
            return {"ok": False, "message": f"Þjöppuð lína {position + 1} hefur engar upprunalínur."}
        covered_indices.extend(paired_indices)

        identity = entry.get("identity", {})
        for field in ("DATE", "AMOUNT", "Positive/Negative"):
            if field not in identity:
                continue
            expected = identity[field]
            if field not in coded_df.columns:
                return {"ok": False, "message": f"Þjöppuðu vinnubókina vantar dálkinn {field}."}
            if _comparison_value(coded_df.iloc[position][field]) != _comparison_value(expected):
                return {
                    "ok": False,
                    "message": (
                        f"Þjöppuð færslulína {position + 1} samsvarar ekki lengur þjöppunarkortinu. "
                        "Pöruðum línum var ekki breytt."
                    ),
                }

        for paired_index in paired_indices:
            try:
                paired_position = int(paired_index)
            except Exception:
                return {"ok": False, "message": "Þjöppunarkort inniheldur ógilda upprunalínu."}
            if not 0 <= paired_position < len(paired_df):
                return {"ok": False, "message": "Þjöppunarkort vísar út fyrir pöruðu gögnin."}
            for field in COMPRESSED_SYNC_FIELDS:
                coded_column = _coded_column_name(coded_df, field)
                if coded_column:
                    paired_df.at[paired_df.index[paired_position], field] = _clean_value(
                        coded_df.iloc[position][coded_column]
                    )

    normalized_coverage = sorted(int(index) for index in covered_indices)
    if normalized_coverage != list(range(len(paired_df))):
        return {
            "ok": False,
            "message": "Þjöppunarkort nær ekki nákvæmlega einu sinni yfir hverja upprunalega færslu.",
        }

    return {"ok": True, "paired_df": paired_df}


def _comparison_value(value) -> str:
    if _is_missing(value):
        return ""
    if isinstance(value, (_datetime.datetime, _datetime.date)):
        return value.isoformat()
    text = str(value).strip()
    try:
        number = float(text.replace(",", "."))
        if number.is_integer():
            return str(int(number))
    except Exception:
        pass
    return text


def _coded_column_name(coded_df: pd.DataFrame, field: str) -> str | None:
    if field in coded_df.columns:
        return field
    for alias in CODED_COLUMN_ALIASES.get(field, ()):
        if alias in coded_df.columns:
            return alias
    return None


def _same_file(first, second) -> bool:
    try:
        return Path(first).resolve() == Path(second).resolve()
    except Exception:
        return False


def _unique_session_dir(imports_root: Path, timestamp: _datetime.datetime, client_hint: str, bank_type: str) -> Path:
    stamp = timestamp.strftime("%Y-%m-%d_%H%M%S")
    client_slug = _slug(client_hint, "client")
    bank_slug = _slug(bank_type, "bank")
    base = imports_root / f"{stamp}_{client_slug}_{bank_slug}"
    candidate = base
    suffix = 2
    while candidate.exists():
        candidate = imports_root / f"{base.name}_{suffix:02d}"
        suffix += 1
    return candidate


def _prefixed_raw_columns(columns) -> tuple[list[tuple[object, str]], dict[str, str]]:
    pairs = []
    mapping = {}
    used = set()
    for position, column in enumerate(columns):
        name = f"raw_{_slug(column, f'column_{position + 1}')}"
        base = name
        suffix = 2
        while name in used:
            name = f"{base}_{suffix}"
            suffix += 1
        used.add(name)
        pairs.append((column, name))
        mapping[name] = str(column)
    return pairs, mapping


def _raw_index_for_coded_row(coded_row, fallback_position: int):
    if TRAINING_RAW_INDEX_COLUMN in getattr(coded_row, "index", []):
        value = coded_row.get(TRAINING_RAW_INDEX_COLUMN)
        if not _is_missing(value):
            return _normalize_index_value(value)
    return fallback_position


def _raw_row_by_index(raw_df: pd.DataFrame, raw_index):
    if raw_df is None or raw_df.empty:
        return None
    if raw_index in raw_df.index:
        return raw_df.loc[raw_index]
    try:
        pos = int(raw_index)
        if 0 <= pos < len(raw_df):
            return raw_df.iloc[pos]
    except Exception:
        pass
    return None


def _normalize_index_value(value):
    try:
        if float(value).is_integer():
            return int(value)
    except Exception:
        pass
    return value


def _clean_value(value):
    if _is_missing(value):
        return ""
    if isinstance(value, (_datetime.datetime, _datetime.date)):
        return value.isoformat()
    try:
        if hasattr(value, "item"):
            return value.item()
    except Exception:
        pass
    return value


def _is_missing(value) -> bool:
    try:
        return bool(pd.isna(value))
    except Exception:
        return value is None


def _slug(value, fallback: str) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_").lower()
    text = re.sub(r"_+", "_", text)
    return (text or fallback)[:72]

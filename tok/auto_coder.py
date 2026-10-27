from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import json
import math
import os
import re
import sys
import unicodedata
import uuid

import pandas as pd


CODE_PATTERN = re.compile(r"^\d{3,5}$")
CARD_MASK_PATTERN = re.compile(r"\*\*-?\d+|\d{8,}|\d{4,}\*+\d+")

COUNTER_ACCOUNT_CODES = {
    "1000", "1007", "1010",
    "7810", "7820", "7825", "7830", "7841", "7842", "7846", "7850",
    "9330", "9336",
}

TEXT_SIGNAL_COLUMNS = [
    "Söluaðilaflokkun",
    "Soluaðilaflokkun",
    "Soluflokkur",
    "Tegund",
    "Textalykill",
    "Upplýsingar",
    "Upplysingar",
    "Skýring greiðslu",
    "Skyring greidslu",
    "Skýring",
    "Skyring",
    "Texti",
    "Aðrar upplýsingar",
    "Adrar upplysingar",
    "Upplýsingar um færslu",
    "Upplysingar um faerslu",
    "Færslulykill",
    "Faerslulykill",
    "Söluaðili",
    "Soluaðili",
    "Nafn viðtakanda eða greiðanda",
    "Nafn vidtakanda eda greidanda",
]

PAIR_REFERENCE_COLUMN_KEYS = {
    "tnrsedilnr",
    "sedilnr",
    "sedilnumer",
    "krofunumer",
}

LEGAL_SUFFIXES = {
    "ehf", "hf", "slf", "ohf", "ses", "sf", "bs", "llc", "ltd", "inc",
    "co", "com", "ab", "as", "gmbh",
}

LOCATION_WORDS = {
    "reykjavik", "rvk", "kopavogur", "hafnarfjordur", "akureyri",
    "selfoss", "selfossi", "keflavik", "gardabaer", "garoabaer",
    "mosfellsbaer", "reykjanesbaer", "sudurnesjum", "sudurnes",
    "bildshofda", "glaes", "smara", "lagmuli", "louholar", "skeifan",
    "gnodarvogi", "artunshofdi", "holtagordum", "sundagordum",
    "heidrun", "fitjar", "kaplakriki", "bustadav", "flugvallarv",
    "sudurhellu", "kleppsvegi", "alfheimum", "nordlingaholt",
    "gylfaflot", "sudurfelli", "haholt", "blidubakka",
}

GENERIC_WORDS = {
    "greidsla", "greitt", "millifaert", "innborgun", "utborgun",
    "debitkortafaersla", "uttekt", "med", "debetkorti", "visa",
    "kort", "kreditkort", "reikningur", "reikn", "nr", "kostnadur",
}

PROCESSOR_PREFIXES = {"str", "paypal", "sumup"}

GENERIC_TRAINING_PATTERNS = [
    re.compile(r"^(reikn|reikningur)\s+\d+$"),
    re.compile(r"^(innb|innborgun|innborgad|sala|vsk|laun)$"),
]

MEMORY_FILENAME = "auto_code_memory.json"
MEMORY_VERSION = 6
TRAINING_RAW_INDEX_COLUMN = "_TRAINING_RAW_INDEX"

COMPANY_WORDS = {
    "verslun", "solur", "soluturn", "kiosk", "bakari", "bakarameistarinn",
    "restaurant", "roasters", "kaffi", "grill", "pizza", "dominos",
    "byggingar", "verktaekni", "verkfaeri", "verkstedi", "verkstaedi",
    "smurstod", "bilapunkturinn", "banki", "bankinn", "lifeyrissjodur",
    "tryggingar", "bokhald", "logmenn", "legal", "workspace", "hotel",
}

PERSON_NAME_SUFFIXES = ("son", "dottir", "dottur")

RESTAURANT_SUPERMARKET_KEYWORDS = {
    "bonus", "kronan", "netto", "hagkaup", "costco", "costco wholesale",
}

RESTAURANT_SUPERMARKET_EXCLUDE_KEYWORDS = {
    "fuel", "bensin", "bensinstod", "olis", "orkan", "atlantsolia", "n1",
}

GAS_STATION_KEYWORDS = {
    "n1", "olis", "orkan", "atlantsolia", "bensin", "bensinstod",
}

VEHICLE_EXPENSE_ACCOUNT_CODES = {"4635", "4640"}
RAW_DETERMINISTIC_CONFIDENCE_SOURCES = {
    "hardcoded_bank_signal",
    "hardcoded_collection_cost",
}


@dataclass
class AutoCodeResult:
    code: str = ""
    confidence: int = 0
    status: str = "review"
    reason: str = "No confident match"
    source: str = ""
    matched_text: str = ""


def _base_dir() -> Path:
    module_dir = Path(__file__).resolve().parent
    candidates = [
        module_dir,
        Path(sys.executable).resolve().parent / "tok",
        Path(sys.executable).resolve().parent.parent / "tok",
        Path.cwd() / "tok",
    ]
    for candidate in candidates:
        if (candidate / "trainingCoded").exists() or (candidate / "coded keys").exists():
            return candidate
    return module_dir


def _translate_icelandic(text: str) -> str:
    table = str.maketrans({
        "ð": "d", "Ð": "d",
        "þ": "th", "Þ": "th",
        "æ": "ae", "Æ": "ae",
        "ö": "o", "Ö": "o",
        "á": "a", "Á": "a",
        "é": "e", "É": "e",
        "í": "i", "Í": "i",
        "ó": "o", "Ó": "o",
        "ú": "u", "Ú": "u",
        "ý": "y", "Ý": "y",
    })
    return text.translate(table)


def normalize_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip().lower()
    if not text:
        return ""
    text = _translate_icelandic(text)
    text = "".join(
        ch for ch in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(ch)
    )
    text = text.replace("&", " og ")
    text = text.replace("*", " ")
    text = CARD_MASK_PATTERN.sub(" ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def tokenize(value, keep_generic: bool = False, drop_processors: bool = False) -> list[str]:
    text = normalize_text(value)
    if not text:
        return []
    tokens = []
    for token in text.split():
        if token.isdigit():
            continue
        if len(token) <= 1 and not token.isalpha():
            continue
        if token in LEGAL_SUFFIXES:
            continue
        if token in LOCATION_WORDS:
            continue
        if not keep_generic and token in GENERIC_WORDS:
            continue
        tokens.append(token)
    if drop_processors and len(tokens) > 1 and tokens[0] in PROCESSOR_PREFIXES:
        tokens = tokens[1:]
    return tokens


def merchant_key(value) -> str:
    return " ".join(tokenize(value))


def merchant_family(value) -> str:
    tokens = tokenize(value, drop_processors=True)
    if not tokens:
        return ""
    first = tokens[0]
    special_two = {
        ("google", "workspace"),
        ("google", "one"),
        ("atlantsolia", "kaplakriki"),
        ("atlantsolia", "bustadav"),
        ("orkan", "sudurfelli"),
        ("orkan", "gylfaflot"),
        ("costco", "wholesale"),
    }
    if len(tokens) >= 2 and (tokens[0], tokens[1]) in special_two:
        return " ".join(tokens[:2])
    chain_words = {
        "bonus", "kronan", "netto", "hagkaup", "costco", "byko",
        "bauhaus", "husasmidjan", "parka", "n1", "olis", "orkan",
        "atlantsolia", "nova", "siminn", "landsbankinn", "arion",
        "islandsbanki", "dominos", "kfc", "subway", "teya", "straumur",
    }
    if first in chain_words:
        return first
    return " ".join(tokens[:2]) if len(tokens) >= 2 else first


def looks_like_person_name(value) -> bool:
    original = "" if value is None else str(value).strip()
    text = normalize_text(value)
    if not text:
        return False
    raw_tokens = [token for token in text.split() if token and not token.isdigit()]
    if len(raw_tokens) < 2 or len(raw_tokens) > 4:
        return False
    if any(token in LEGAL_SUFFIXES or token in COMPANY_WORDS for token in raw_tokens):
        return False
    if any(token in LOCATION_WORDS for token in raw_tokens):
        return False
    tokens = tokenize(value, keep_generic=True, drop_processors=True)
    if len(tokens) < 2 or len(tokens) > 4:
        return False
    if any(token in COMPANY_WORDS or token in GENERIC_WORDS for token in tokens):
        return False
    if any(token in RESTAURANT_SUPERMARKET_KEYWORDS or token in GAS_STATION_KEYWORDS for token in tokens):
        return False
    if any(token.endswith(PERSON_NAME_SUFFIXES) for token in tokens[1:]):
        return True
    if original and original.upper() == original and any(ch.isalpha() for ch in original):
        return False
    return all(token.isalpha() and len(token) >= 3 for token in tokens)


def looks_like_person_id(value) -> bool:
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except Exception:
        pass
    digits = re.sub(r"\D", "", str(value))
    if len(digits) == 9:
        digits = "0" + digits
    if len(digits) != 10:
        return False
    try:
        day = int(digits[:2])
        month = int(digits[2:4])
    except Exception:
        return False
    return 1 <= day <= 31 and 1 <= month <= 12


def client_key_from_path(path) -> str:
    if not path:
        return ""
    stem = Path(str(path)).stem
    stem = re.sub(r"_innlestur_\d{8}_\d{6}$", "", stem, flags=re.I)
    text = normalize_text(stem)
    tokens = []
    for token in text.split():
        if token.isdigit():
            continue
        if re.fullmatch(r"\d+[a-z]?", token):
            continue
        if token in {"innlestur", "coded", "raw", "xlsx", "xls", "bank", "kort"}:
            continue
        tokens.append(token)
    return " ".join(tokens[:3])


def clean_code(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    try:
        dec = Decimal(text.replace(",", "."))
        if dec == dec.to_integral_value():
            text = str(int(dec))
    except (InvalidOperation, ValueError):
        text = re.sub(r"\.0$", "", text)
    return text if CODE_PATTERN.match(text) else ""


def _existing_confidence_value(row):
    for column in ("CONFIDENCE", "Cnf"):
        try:
            value = row.get(column, "")
        except Exception:
            value = ""
        if not is_missing_value(value):
            return value
    return ""


def _existing_debit_is_protected(row, overwrite_existing: bool) -> bool:
    """Protect manual/nonblank work before any deterministic pre-pass."""
    try:
        existing = clean_code(row.get("DEBIT"))
        marker = str(_existing_confidence_value(row) or "").strip().upper()
        source = str(row.get("_AUTO_CODE_SOURCE", "") or "").strip().lower()
    except Exception:
        return False
    if not existing:
        return False
    if marker in {"AD", "MN"} or source == "manual_review" or source.startswith("hardcoded_"):
        return True
    return not overwrite_existing


def _apply_vehicle_expense_code_choice(
    result: AutoCodeResult,
    selected_code,
) -> AutoCodeResult:
    """Normalize only automatically predicted 4635/4640 vehicle expenses."""
    target_code = clean_code(selected_code)
    predicted_code = clean_code(result.code)
    if (
        target_code not in VEHICLE_EXPENSE_ACCOUNT_CODES
        or predicted_code not in VEHICLE_EXPENSE_ACCOUNT_CODES
        or predicted_code == target_code
    ):
        return result

    result.code = target_code
    note = f"vehicle expense account selected: {target_code}"
    if note not in result.reason:
        result.reason = f"{result.reason}; {note}" if result.reason else note
    return result


def _output_confidence(result: AutoCodeResult):
    """Use AD only when a deterministic rule was proven by raw bank text."""
    if result.source in RAW_DETERMINISTIC_CONFIDENCE_SOURCES:
        return "AD"
    return result.confidence if result.confidence else ""


def _parse_amount(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, (int, float)):
        if math.isnan(value) if isinstance(value, float) else False:
            return None
        return Decimal(str(value))
    text = str(value).strip().replace(" ", "")
    if not text:
        return None
    text = text.replace("\u00a0", "")
    sign = Decimal("-1") if text.startswith("(") and text.endswith(")") else Decimal("1")
    if text.startswith("-"):
        sign = Decimal("-1")
        text = text[1:]
    elif text.startswith("+"):
        text = text[1:]
    if text.startswith("(") and text.endswith(")"):
        text = text[1:-1]
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
        return Decimal(text) * sign
    except InvalidOperation:
        return None


def is_missing_value(value) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    if isinstance(value, str):
        return not value.strip()
    return False


def _parse_day(value) -> int | None:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, pd.Timestamp):
        return int(value.day)
    if isinstance(value, datetime.datetime):
        return int(value.day)
    if isinstance(value, datetime.date):
        return int(value.day)
    text = str(value).strip()
    if not text:
        return None
    match = re.match(r"^\s*(\d{1,2})\s*[.\-/]", text)
    if match:
        day = int(match.group(1))
        return day if 1 <= day <= 31 else None
    match = re.match(r"^\s*\d{4}\s*[.\-/]\s*\d{1,2}\s*[.\-/]\s*(\d{1,2})", text)
    if match:
        day = int(match.group(1))
        return day if 1 <= day <= 31 else None
    return None


def _parse_month(value) -> int | None:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, pd.Timestamp):
        return int(value.month)
    if isinstance(value, datetime.datetime):
        return int(value.month)
    if isinstance(value, datetime.date):
        return int(value.month)
    text = str(value).strip()
    if not text:
        return None
    match = re.match(r"^\s*\d{1,2}\s*[.\-/]\s*(\d{1,2})\s*[.\-/]", text)
    if match:
        month = int(match.group(1))
        return month if 1 <= month <= 12 else None
    match = re.match(r"^\s*\d{4}\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*\d{1,2}", text)
    if match:
        month = int(match.group(1))
        return month if 1 <= month <= 12 else None
    return None


def _day_bucket(value) -> str:
    day = _parse_day(value)
    if day is None:
        return ""
    if day <= 7:
        return "d01_07"
    if day <= 14:
        return "d08_14"
    if day <= 21:
        return "d15_21"
    return "d22_31"


def _day_exact(value) -> str:
    day = _parse_day(value)
    return f"d{day:02d}" if day is not None else ""


def _day_window(value) -> str:
    day = _parse_day(value)
    if day is None:
        return ""
    if 1 <= day <= 4:
        return "dw01_04"
    if 5 <= day <= 7:
        return "dw05_07"
    if 8 <= day <= 14:
        return "dw08_14"
    if 15 <= day <= 17:
        return "dw15_17"
    if 18 <= day <= 24:
        return "dw18_24"
    return "dw25_31"


def _month_exact(value) -> str:
    month = _parse_month(value)
    return f"m{month:02d}" if month is not None else ""


def _month_parity(value) -> str:
    month = _parse_month(value)
    if month is None:
        return ""
    return "month_even" if month % 2 == 0 else "month_odd"


def _amount_bucket(value) -> str:
    amount = _parse_amount(value)
    if amount is None:
        return ""
    amount = abs(amount)
    buckets = [
        (Decimal("1000"), "a0000_001k"),
        (Decimal("3000"), "a001k_003k"),
        (Decimal("7500"), "a003k_007k"),
        (Decimal("15000"), "a007k_015k"),
        (Decimal("30000"), "a015k_030k"),
        (Decimal("75000"), "a030k_075k"),
        (Decimal("150000"), "a075k_150k"),
        (Decimal("300000"), "a150k_300k"),
        (Decimal("750000"), "a300k_750k"),
    ]
    for limit, label in buckets:
        if amount <= limit:
            return label
    return "a750k_plus"


def _amount_round_bucket(value) -> str:
    amount = _parse_amount(value)
    if amount is None:
        return ""
    amount = abs(amount)
    if amount == 0:
        return ""
    try:
        if amount % Decimal("10000") == 0:
            return "round_10k"
        if amount % Decimal("1000") == 0:
            return "round_1k"
        if amount % Decimal("100") == 0:
            return "round_100"
    except Exception:
        return ""
    return "not_round"


def _repeat_bucket(count) -> str:
    try:
        count = int(count or 0)
    except Exception:
        count = 0
    if count <= 1:
        return "r01"
    if count <= 3:
        return "r02_03"
    if count <= 7:
        return "r04_07"
    if count <= 15:
        return "r08_15"
    return "r16_plus"


def _context_features(text, amount=None, date=None, repeat_count=None, id_value=None) -> list[str]:
    key = merchant_key(text)
    family = merchant_family(text)
    amount_bucket = _amount_bucket(amount)
    round_bucket = _amount_round_bucket(amount)
    day_bucket = _day_bucket(date)
    day_exact = _day_exact(date)
    day_window = _day_window(date)
    month_exact = _month_exact(date)
    month_parity = _month_parity(date)
    repeat_bucket = _repeat_bucket(repeat_count)
    person_like = looks_like_person_name(text) or looks_like_person_id(id_value)
    features = []

    if key:
        if amount_bucket:
            features.append(f"key:{key}|amount:{amount_bucket}")
        if round_bucket:
            features.append(f"key:{key}|round:{round_bucket}")
        if day_bucket:
            features.append(f"key:{key}|day_bucket:{day_bucket}")
        if day_exact:
            features.append(f"key:{key}|day:{day_exact}")
        if day_window:
            features.append(f"key:{key}|day_window:{day_window}")
        if month_exact:
            features.append(f"key:{key}|month:{month_exact}")
        if month_parity:
            features.append(f"key:{key}|month_parity:{month_parity}")
        if repeat_bucket:
            features.append(f"key:{key}|repeat:{repeat_bucket}")
        if amount_bucket and day_bucket:
            features.append(f"key:{key}|amount:{amount_bucket}|day_bucket:{day_bucket}")
        if day_window and month_parity:
            features.append(f"key:{key}|day_window:{day_window}|month_parity:{month_parity}")
        if day_exact and month_parity:
            features.append(f"key:{key}|day:{day_exact}|month_parity:{month_parity}")
        if amount_bucket and day_window and month_parity:
            features.append(f"key:{key}|amount:{amount_bucket}|day_window:{day_window}|month_parity:{month_parity}")
        if amount_bucket and repeat_bucket:
            features.append(f"key:{key}|amount:{amount_bucket}|repeat:{repeat_bucket}")

    if family and family != key:
        if amount_bucket:
            features.append(f"family:{family}|amount:{amount_bucket}")
        if round_bucket:
            features.append(f"family:{family}|round:{round_bucket}")
        if day_bucket:
            features.append(f"family:{family}|day_bucket:{day_bucket}")
        if day_window:
            features.append(f"family:{family}|day_window:{day_window}")
        if month_parity:
            features.append(f"family:{family}|month_parity:{month_parity}")
        if repeat_bucket:
            features.append(f"family:{family}|repeat:{repeat_bucket}")
        if amount_bucket and day_bucket:
            features.append(f"family:{family}|amount:{amount_bucket}|day_bucket:{day_bucket}")
        if day_window and month_parity:
            features.append(f"family:{family}|day_window:{day_window}|month_parity:{month_parity}")

    if person_like:
        features.append("type:person")
        if amount_bucket:
            features.append(f"type:person|amount:{amount_bucket}")
        if round_bucket:
            features.append(f"type:person|round:{round_bucket}")
        if day_bucket:
            features.append(f"type:person|day_bucket:{day_bucket}")
        if day_window:
            features.append(f"type:person|day_window:{day_window}")
        if month_parity:
            features.append(f"type:person|month_parity:{month_parity}")
        if repeat_bucket:
            features.append(f"type:person|repeat:{repeat_bucket}")
        if amount_bucket and repeat_bucket:
            features.append(f"type:person|amount:{amount_bucket}|repeat:{repeat_bucket}")
        if day_bucket and repeat_bucket:
            features.append(f"type:person|day_bucket:{day_bucket}|repeat:{repeat_bucket}")

    return list(dict.fromkeys(features))


def _is_generic_training_text(text: str) -> bool:
    key = merchant_key(text)
    if not key:
        return True
    return any(pattern.match(key) for pattern in GENERIC_TRAINING_PATTERNS)


def _best(counter: Counter) -> tuple[str, int, int, float]:
    if not counter:
        return "", 0, 0, 0.0
    code, count = counter.most_common(1)[0]
    total = sum(counter.values())
    ratio = count / total if total else 0.0
    return code, count, total, ratio


def _confidence_from_history(count: int, total: int, ratio: float, exact: bool) -> int:
    if total <= 0:
        return 0
    base = 57 if exact else 45
    count_boost = min(20, int(math.log1p(count) * 8))
    ratio_boost = int(ratio * (28 if exact else 24))
    ambiguity_penalty = 0 if ratio >= 0.9 else int((0.9 - ratio) * 25)
    return max(0, min(99, base + count_boost + ratio_boost - ambiguity_penalty))


def _hardcoded_bank_signal_rule(text: str, signal_text: str) -> AutoCodeResult | None:
    """Return mandatory account codes for narrowly scoped bank explanations."""
    combined = " ".join((normalize_text(text), normalize_text(signal_text))).strip()
    if "rikissjodsinnheimt" not in combined:
        return None

    rules = (
        (("gjold utan stadgreidslu", "gjald utan stadgreidslu"), "9610", "treasury non-withholding levy rule"),
        (("kilometragjald", "kilometrjagjald"), "4610", "treasury kilometre fee rule"),
        (("bifreidagjold", "bifreidagjald"), "4610", "treasury vehicle fee rule"),
        (("gr frestur adflutningsgjalda", "greidslufrestur adflutningsgjalda"), "9540", "treasury deferred import fee rule"),
        (("tryggingagjald",), "9615", "treasury insurance levy rule"),
        (("stadgreidsluskattur",), "9615", "treasury withholding tax rule"),
        (("kostnadur",), "6220", "treasury collection cost rule"),
    )
    for phrases, code, reason in rules:
        if any(phrase in combined for phrase in phrases):
            return AutoCodeResult(
                code=code,
                confidence=100,
                status="coded",
                reason=reason,
                source="hardcoded_bank_signal",
            )
    return None


def _cap_unrecognized_treasury_confidence(
    result: AutoCodeResult,
    text: str,
    signal_text: str,
) -> AutoCodeResult:
    """Keep unfamiliar treasury explanations from appearing fully certain."""
    combined = " ".join((normalize_text(text), normalize_text(signal_text))).strip()
    if "rikissjodsinnheimt" not in combined:
        return result
    if _hardcoded_bank_signal_rule(text, signal_text) is not None:
        return result
    try:
        confidence = int(result.confidence or 0)
    except (TypeError, ValueError):
        confidence = 0
    if confidence > 70:
        result.confidence = 70
        note = "unrecognized treasury explanation confidence cap"
        if note not in result.reason:
            result.reason = f"{result.reason}; {note}" if result.reason else note
    return result


def _is_explicit_collection_cost_signal(text: str) -> bool:
    """Recognize unambiguous collection-cost wording used by bank exports."""
    normalized = normalize_text(text)
    if "kostnadur" not in normalized:
        return False
    return "innheimt" in normalized or "krof" in normalized


def _collection_cost_context_rule(context_flags=None) -> AutoCodeResult | None:
    flags = set(context_flags or ())
    if "explicit_collection_cost" in flags:
        return AutoCodeResult(
            code="6220",
            confidence=100,
            status="coded",
            reason="explicit raw bank collection-cost rule",
            source="hardcoded_collection_cost",
        )
    return None


def _category_rule(
    text: str,
    signal_text: str,
    amount=None,
    industry_context: str = "",
    context_flags=None,
) -> AutoCodeResult | None:
    combined = " ".join([normalize_text(text), normalize_text(signal_text)]).strip()
    if not combined:
        return None
    flags = set(context_flags or ())
    amount_value = _parse_amount(amount)
    abs_amount = abs(amount_value) if amount_value is not None else None

    collection_cost_rule = _collection_cost_context_rule(flags)
    if collection_cost_rule:
        return collection_cost_rule

    if "simgreidsla" in combined:
        return AutoCodeResult(
            code="2200",
            confidence=100,
            status="coded",
            reason="raw bank texti simgreidsla rule",
            source="bank_signal_rule",
        )

    if abs_amount is not None and abs_amount < Decimal("5000"):
        if any(keyword in combined for keyword in GAS_STATION_KEYWORDS):
            return AutoCodeResult(
                code="4082",
                confidence=100,
                status="coded",
                reason="small gas-station amount food rule",
                source="amount_rule",
            )

    if _is_explicit_collection_cost_signal(combined):
        return AutoCodeResult(
            code="6220",
            confidence=96,
            status="coded",
            reason="raw bank collection-cost rule",
            source="bank_signal_rule",
        )

    if industry_context == "restaurant":
        if any(keyword in combined for keyword in RESTAURANT_SUPERMARKET_KEYWORDS):
            if not any(keyword in combined for keyword in RESTAURANT_SUPERMARKET_EXCLUDE_KEYWORDS):
                return AutoCodeResult(
                    code="2107",
                    confidence=100,
                    status="coded",
                    reason="restaurant mode supermarket/raw material rule",
                    source="restaurant_context",
                )
    elif any(keyword in combined for keyword in RESTAURANT_SUPERMARKET_KEYWORDS):
        if not any(keyword in combined for keyword in RESTAURANT_SUPERMARKET_EXCLUDE_KEYWORDS):
            return AutoCodeResult(
                code="4061",
                confidence=100,
                status="coded",
                reason="general mode supermarket/food rule",
                source="general_context",
            )

    rules = [
        (("thjonustugjald", "faerslugjold", "bankagjald"), "4400", 82, "bank fee keyword"),
        (("utvextir", "vaxtagjold", "vextir af"), "6200", 78, "interest expense keyword"),
        (("simi og net", "fjarskipta", "nova", "siminn", "vodafone", "hringdu"), "4030", 72, "phone/internet keyword"),
        (("google workspace", "google one", "godaddy", "domain", "hosting"), "4055", 70, "software/web service keyword"),
        (("parka", "bilastaedi", "bifrei", "parking"), "4640", 70, "vehicle/parking keyword"),
        (("bensin", "bensinstod", "olis", "orkan", "atlantsolia", "n1"), "4600", 68, "fuel keyword"),
        (("vidgerdarverkstaedi", "smurstod", "bilasmid", "dekk"), "4620", 68, "vehicle repair keyword"),
        (("skyndibitastadir", "veiting", "bakari", "kaff", "bonus", "kronan", "netto", "dominos", "kfc"), "4061", 66, "food/coffee keyword"),
        (("jarnvoru", "bygging", "byko", "bauhaus", "husasmidjan", "verkfaeri"), "2100", 65, "hardware/materials keyword"),
        (("trygging", "sjova", "vis tryggingar"), "4230", 62, "insurance keyword"),
        (("bokhald", "gaius"), "4200", 62, "bookkeeping keyword"),
        (("prentun", "ritfong", "pappir"), "4031", 62, "office/printing keyword"),
        (("auglysing", "facebook", "meta", "ads"), "4050", 62, "advertising keyword"),
    ]
    for keywords, code, confidence, reason in rules:
        if any(keyword in combined for keyword in keywords):
            return AutoCodeResult(code=code, confidence=confidence, status="coded", reason=reason, source="keyword")
    return None


def _source_signal_text(source_row) -> str:
    if source_row is None:
        return ""
    parts = []
    seen_columns = set()
    seen_values = set()

    def append_value(value):
        if is_missing_value(value):
            return
        value_text = str(value).strip()
        if not value_text:
            return
        value_key = value_text.casefold()
        if value_key in seen_values:
            return
        seen_values.add(value_key)
        parts.append(value_text)

    for col in TEXT_SIGNAL_COLUMNS:
        try:
            value = source_row.get(col)
        except Exception:
            continue
        if value is None:
            continue
        seen_columns.add(col)
        append_value(value)

    return " | ".join(parts)


def _source_rule_text(source_row) -> str:
    """Collect all textual raw cells for deterministic cross-bank rules."""
    if source_row is None:
        return ""
    signal_text = _source_signal_text(source_row)
    parts = [signal_text] if signal_text else []
    seen_columns = set(TEXT_SIGNAL_COLUMNS)
    seen_values = {signal_text.casefold()} if signal_text else set()

    # Bank exports do not use one stable name for merchant and explanation
    # fields. Deterministic rules can safely inspect every textual cell while
    # the learned coder continues to receive only the curated signal columns.
    try:
        source_items = source_row.items()
    except Exception:
        source_items = ()
    for col, value in source_items:
        if col in seen_columns or not isinstance(value, str):
            continue
        value_text = value.strip()
        if not value_text:
            continue
        value_key = value_text.casefold()
        if value_key in seen_values:
            continue
        seen_values.add(value_key)
        parts.append(value_text)
    return " | ".join(parts)


def _source_collection_cost_kind(source_row) -> str:
    """Return ``explicit`` or ``bare`` for known cross-bank cost signals."""
    if source_row is None:
        return ""
    if _is_explicit_collection_cost_signal(_source_rule_text(source_row)):
        return "explicit"
    try:
        source_items = source_row.items()
    except Exception:
        source_items = ()
    for _column, value in source_items:
        if not isinstance(value, str):
            continue
        if normalize_text(value).strip() == "kostnadur":
            return "bare"
    return ""


def _normalized_column_key(value) -> str:
    return re.sub(r"[^a-z0-9]", "", normalize_text(value))


def _canonical_pair_reference(value) -> str:
    if is_missing_value(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"^(\d+)\.0+$", r"\1", text)
    return re.sub(r"[^a-z0-9]", "", normalize_text(text))


def _source_pair_reference(source_row) -> str:
    """Read a claim/reference key without confusing a kennitala for one."""
    if source_row is None:
        return ""
    try:
        source_items = list(source_row.items())
    except Exception:
        return ""

    for column, value in source_items:
        if _normalized_column_key(column) in PAIR_REFERENCE_COLUMN_KEYS:
            reference = _canonical_pair_reference(value)
            if reference:
                return reference

    # New seven-column Islandsbanki statements use Tilvísun as the claim key.
    # Legacy exports use that column for the payee kennitala, so accept it only
    # when it differs from the workbook's dedicated kennitala value.
    kennitala_values = {
        re.sub(r"\D", "", re.sub(r"^(\d+)\.0+$", r"\1", str(value).strip()))
        for column, value in source_items
        if "kennitala" in _normalized_column_key(column) and not is_missing_value(value)
    }
    for column, value in source_items:
        if _normalized_column_key(column) != "tilvisun":
            continue
        reference = _canonical_pair_reference(value)
        reference_digits = re.sub(r"\D", "", reference)
        if reference and kennitala_values and reference_digits not in kennitala_values:
            return reference
    return ""


def _source_row_for_output(source_df, output_row, output_index):
    if source_df is None:
        return None

    candidates = []
    try:
        raw_index = output_row.get(TRAINING_RAW_INDEX_COLUMN)
    except Exception:
        raw_index = None
    if not is_missing_value(raw_index):
        candidates.append(raw_index)
        try:
            raw_index_number = float(raw_index)
            if raw_index_number.is_integer():
                candidates.append(int(raw_index_number))
        except Exception:
            pass
    candidates.append(output_index)

    checked = set()
    for candidate in candidates:
        try:
            candidate_key = (type(candidate), candidate)
            if candidate_key in checked:
                continue
            checked.add(candidate_key)
        except Exception:
            pass
        try:
            return source_df.loc[candidate]
        except Exception:
            continue

    try:
        return source_df.iloc[output_index]
    except Exception:
        return None


def _is_outgoing_row(row, amount=None) -> bool:
    sign = str(row.get("Positive/Negative", "")).strip()
    if sign == "-":
        return True
    amount_value = _parse_amount(row.get("AMOUNT") if amount is None else amount)
    return sign == "" and amount_value is not None and amount_value < 0


def _row_context_flags(
    df: pd.DataFrame,
    source_df: pd.DataFrame | None = None,
) -> dict[object, set[str]]:
    flags: dict[object, set[str]] = defaultdict(set)
    grouped: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)

    for idx, row in df.iterrows():
        amount = _parse_amount(row.get("AMOUNT"))
        if amount is None or not _is_outgoing_row(row, amount):
            continue
        source_row = _source_row_for_output(source_df, row, idx)
        cost_kind = _source_collection_cost_kind(source_row)
        if cost_kind == "explicit":
            # The wording itself is decisive and must outrank merchant memory.
            flags[idx].add("explicit_collection_cost")
        id_text = re.sub(r"^(\d+)\.0+$", r"\1", str(row.get("ID", "") or "").strip())
        id_digits = re.sub(r"\D", "", id_text)
        identity = f"id:{id_digits}" if id_digits else f"text:{merchant_key(row.get('TEXT', ''))}"
        if identity in {"id:", "text:"}:
            continue
        date_key = str(row.get("DATE", "")).strip()
        if not date_key:
            continue
        grouped[(date_key, identity)].append({
            "idx": idx,
            "reference": _source_pair_reference(source_row),
            "cost_kind": cost_kind,
        })

    # Some new bank exports provide only the bare value "Kostnaður". Treat it
    # as collection cost only when another outgoing row shares its claim key.
    # Partitioning by reference keeps 3/4/6-row dates from contaminating one
    # another and deliberately avoids inferring 6220 from amount alone.
    for rows in grouped.values():
        by_reference: dict[str, list[dict[str, object]]] = defaultdict(list)
        for record in rows:
            reference = str(record["reference"] or "")
            if reference:
                by_reference[reference].append(record)
        for reference_rows in by_reference.values():
            if len(reference_rows) < 2:
                continue
            for record in reference_rows:
                if record["cost_kind"] == "bare":
                    flags[record["idx"]].add("explicit_collection_cost")
    return flags


class TransactionAutoCoder:
    def __init__(
        self,
        training_dir=None,
        key_dir=None,
        memory_path=None,
        training_data_dir=None,
        force_rebuild: bool = False,
    ):
        base = _base_dir()
        self.training_dir = Path(training_dir) if training_dir else base / "trainingCoded"
        self.training_data_dir = (
            Path(training_data_dir)
            if training_data_dir
            else self._resolve_training_data_dir(base)
        )
        self.key_dir = self._resolve_key_dir(key_dir, base)
        self.memory_path = Path(memory_path) if memory_path else self.training_dir / MEMORY_FILENAME
        self.global_exact: dict[str, Counter] = defaultdict(Counter)
        self.global_family: dict[str, Counter] = defaultdict(Counter)
        self.client_exact: dict[tuple[str, str], Counter] = defaultdict(Counter)
        self.client_family: dict[tuple[str, str], Counter] = defaultdict(Counter)
        self.context_counts: dict[str, Counter] = defaultdict(Counter)
        self.client_context_counts: dict[tuple[str, str], Counter] = defaultdict(Counter)
        self.token_counts: dict[str, Counter] = defaultdict(Counter)
        self.bigram_counts: dict[str, Counter] = defaultdict(Counter)
        self.code_counts: Counter = Counter()
        self.account_descriptions: dict[str, str] = {}
        self.training_rows = 0
        self.training_files = 0
        self.source_fingerprint: dict[str, dict[str, int]] = {}
        self.loaded_from_memory = False
        self.memory_saved = False
        self.load(force_rebuild=force_rebuild)

    def _reset_model(self):
        self.global_exact = defaultdict(Counter)
        self.global_family = defaultdict(Counter)
        self.client_exact = defaultdict(Counter)
        self.client_family = defaultdict(Counter)
        self.context_counts = defaultdict(Counter)
        self.client_context_counts = defaultdict(Counter)
        self.token_counts = defaultdict(Counter)
        self.bigram_counts = defaultdict(Counter)
        self.code_counts = Counter()
        self.account_descriptions = {}
        self.training_rows = 0
        self.training_files = 0
        self.source_fingerprint = {}
        self.loaded_from_memory = False
        self.memory_saved = False

    def _resolve_training_data_dir(self, base: Path) -> Path:
        candidates = [
            self.training_dir.parent / "trainingData",
            base / "trainingData",
        ]
        for candidate in candidates:
            if candidate.exists():
                return candidate
        return candidates[0]

    def _resolve_key_dir(self, key_dir, base: Path) -> Path:
        candidates = []
        if key_dir:
            candidates.append(Path(key_dir))
        candidates.extend([
            self.training_dir / "coded keys",
            self.training_dir / "keys",
            self.training_dir / "codedKeys",
            self.training_dir / "key mappings",
            base / "coded keys",
        ])
        for candidate in candidates:
            if candidate.exists():
                return candidate
        return candidates[0] if candidates else base / "coded keys"

    def load(self, force_rebuild: bool = False):
        # Normal auto-coding uses the saved model. Only the explicit rebuild
        # action re-reads training workbooks, so a newly exported/unreviewed
        # file cannot silently become training data during an ordinary run.
        if not force_rebuild and self._load_memory():
            self.loaded_from_memory = True
            return

        # A rebuild always starts empty and never merges the previous JSON.
        self._reset_model()
        self.account_descriptions.update(self._load_account_descriptions())
        training_sources = self._training_table_sources()
        if not training_sources:
            self.save_memory()
            return
        for path in training_sources:
            if path.name.startswith("~$"):
                continue
            source_key, source_meta = self._source_fingerprint_for_path(path)
            try:
                df = self._read_training_table(path)
            except Exception:
                continue
            self.training_files += 1
            client_key = client_key_from_path(path)
            records = []
            for _, row in df.iterrows():
                target = self._target_from_training_row(row)
                if not target:
                    continue
                text = row.get("TEXT", "")
                if _is_generic_training_text(text):
                    continue
                key = merchant_key(text)
                family = merchant_family(text)
                if not key:
                    continue
                records.append((row, target, text, key, family))

            key_counts = Counter(key for _, _, _, key, _ in records)
            family_counts = Counter(family for _, _, _, _, family in records if family)
            for row, target, text, key, family in records:
                self.training_rows += 1
                self.code_counts[target] += 1
                self.global_exact[key][target] += 1
                if family:
                    self.global_family[family][target] += 1
                if client_key:
                    self.client_exact[(client_key, key)][target] += 1
                    if family:
                        self.client_family[(client_key, family)][target] += 1
                tokens = tokenize(text, drop_processors=True)
                for token in set(tokens):
                    self.token_counts[token][target] += 1
                for bigram in set(" ".join(tokens[i:i + 2]) for i in range(len(tokens) - 1)):
                    self.bigram_counts[bigram][target] += 1
                repeat_count = key_counts.get(key) or family_counts.get(family) or 1
                for feature in _context_features(
                    text,
                    amount=row.get("AMOUNT"),
                    date=row.get("DATE"),
                    repeat_count=repeat_count,
                    id_value=row.get("ID"),
                ):
                    self.context_counts[feature][target] += 1
                    if client_key:
                        self.client_context_counts[(client_key, feature)][target] += 1
            self.source_fingerprint[source_key] = source_meta
        self.source_fingerprint.update(self._source_fingerprint(include_excel=False))
        self.save_memory()

    def _training_sources(self) -> list[Path]:
        paths = []
        paths.extend(self._training_table_sources())
        if self.key_dir.exists():
            paths.extend(
                p for p in self.key_dir.rglob("*")
                if p.is_file() and not p.name.startswith("~$")
            )
        return paths

    def _training_table_sources(self) -> list[Path]:
        sources = []
        sources.extend(self._training_excel_sources())
        sources.extend(self._training_archive_sources())
        return sorted(dict.fromkeys(sources), key=lambda path: str(path).lower())

    def _training_excel_sources(self) -> list[Path]:
        if not self.training_dir.exists():
            return []
        return [
            p for p in self.training_dir.rglob("*.xls*")
            if p.is_file() and not p.name.startswith("~$")
        ]

    def _training_archive_sources(self) -> list[Path]:
        imports_root = self.training_data_dir / "imports"
        if not imports_root.is_dir():
            return []
        session_dirs = set()
        session_dirs.update(path.parent for path in imports_root.rglob("manifest.json"))
        session_dirs.update(path for path in imports_root.iterdir() if path.is_dir())

        sources = []
        for session_dir in sorted(session_dirs, key=lambda path: str(path).lower()):
            coded_path = self._session_coded_workbook(session_dir)
            if coded_path is not None:
                sources.append(coded_path)
        return sources

    def _session_coded_workbook(self, session_dir: Path) -> Path | None:
        manifest_path = session_dir / "manifest.json"
        try:
            with open(manifest_path, "r", encoding="utf-8") as handle:
                manifest = json.load(handle)
        except Exception:
            manifest = {}

        names = [
            manifest.get("archive_files", {}).get("coded") if isinstance(manifest, dict) else None,
            manifest.get("output_coded_filename") if isinstance(manifest, dict) else None,
        ]
        for name in names:
            if not name:
                continue
            candidate = session_dir / Path(str(name)).name
            if candidate.is_file() and candidate.suffix.lower() in {".xlsx", ".xls", ".xlsm"}:
                return candidate

        # A manifest explicitly names the coded workbook. If that workbook was
        # deleted, the session must disappear from the next clean rebuild; do
        # not fall back to paired_rows.csv, raw.xlsx, or a combined workbook.
        if any(names):
            return None

        candidates = [
            path
            for path in session_dir.glob("*.xls*")
            if path.is_file()
            and not path.name.startswith("~$")
            and path.name.lower() != "raw.xlsx"
            and "compressed_txs" not in path.stem.lower()
        ]
        coded_named = [path for path in candidates if path.stem.lower().endswith("_coded")]
        if not coded_named:
            return None
        return max(coded_named, key=lambda path: path.stat().st_mtime)

    def _read_training_table(self, path: Path) -> pd.DataFrame:
        suffix = path.suffix.lower()
        if suffix == ".csv":
            return pd.read_csv(path, dtype=object, keep_default_na=False, encoding="utf-8-sig")
        return pd.read_excel(path, dtype=object, keep_default_na=False)

    def _source_fingerprint_for_path(self, path: Path) -> tuple[str, dict[str, int]]:
        try:
            stat = path.stat()
        except OSError:
            return str(path), {"mtime_ns": 0, "size": 0}
        try:
            rel = str(path.resolve().relative_to(_base_dir().resolve()))
        except Exception:
            rel = str(path.resolve())
        metadata = {
            "mtime_ns": int(stat.st_mtime_ns),
            "size": int(stat.st_size),
        }
        sync_error_path = path.parent / "paired_rows_sync_error.txt"
        if path.parent.parent.name == "imports" and sync_error_path.exists():
            try:
                error_stat = sync_error_path.stat()
                metadata["sync_error_mtime_ns"] = int(error_stat.st_mtime_ns)
                metadata["sync_error_size"] = int(error_stat.st_size)
            except OSError:
                pass
        return rel, metadata

    def _source_fingerprint(self, include_excel: bool = True) -> dict[str, dict[str, int]]:
        fingerprint = {}
        sources = self._training_sources() if include_excel else []
        if not include_excel and self.key_dir.exists():
            sources = [
                p for p in self.key_dir.rglob("*")
                if p.is_file() and not p.name.startswith("~$")
            ]
        for path in sources:
            key, meta = self._source_fingerprint_for_path(path)
            fingerprint[key] = meta
        return fingerprint

    @staticmethod
    def _portable_path(path: Path) -> str:
        try:
            return str(path.resolve().relative_to(_base_dir().resolve()))
        except Exception:
            return str(path)

    @staticmethod
    def _counter_to_dict(counter: Counter) -> dict[str, int]:
        return {str(key): int(value) for key, value in counter.items()}

    @staticmethod
    def _dict_to_counter(values) -> Counter:
        return Counter({str(key): int(value) for key, value in (values or {}).items()})

    @classmethod
    def _counter_map_to_payload(cls, mapping) -> dict[str, dict[str, int]]:
        return {str(key): cls._counter_to_dict(counter) for key, counter in mapping.items()}

    @classmethod
    def _payload_to_counter_map(cls, values) -> defaultdict:
        output = defaultdict(Counter)
        for key, counter in (values or {}).items():
            output[str(key)] = cls._dict_to_counter(counter)
        return output

    @classmethod
    def _tuple_counter_map_to_payload(cls, mapping) -> list[dict[str, object]]:
        payload = []
        for (client, key), counter in mapping.items():
            payload.append({
                "client": str(client),
                "key": str(key),
                "counts": cls._counter_to_dict(counter),
            })
        return payload

    @classmethod
    def _payload_to_tuple_counter_map(cls, values) -> defaultdict:
        output = defaultdict(Counter)
        for item in values or []:
            if not isinstance(item, dict):
                continue
            client = str(item.get("client", ""))
            key = str(item.get("key", ""))
            if not client or not key:
                continue
            output[(client, key)] = cls._dict_to_counter(item.get("counts", {}))
        return output

    def _to_payload(self) -> dict[str, object]:
        return {
            "version": MEMORY_VERSION,
            "training_dir": self._portable_path(self.training_dir),
            "training_data_dir": self._portable_path(self.training_data_dir),
            "key_dir": self._portable_path(self.key_dir),
            "source_fingerprint": dict(self.source_fingerprint) or self._source_fingerprint(),
            "training_rows": int(self.training_rows),
            "training_files": int(self.training_files),
            "global_exact": self._counter_map_to_payload(self.global_exact),
            "global_family": self._counter_map_to_payload(self.global_family),
            "client_exact": self._tuple_counter_map_to_payload(self.client_exact),
            "client_family": self._tuple_counter_map_to_payload(self.client_family),
            "context_counts": self._counter_map_to_payload(self.context_counts),
            "client_context_counts": self._tuple_counter_map_to_payload(self.client_context_counts),
            "token_counts": self._counter_map_to_payload(self.token_counts),
            "bigram_counts": self._counter_map_to_payload(self.bigram_counts),
            "code_counts": self._counter_to_dict(self.code_counts),
            "account_descriptions": dict(self.account_descriptions),
        }

    def _load_payload(self, payload: dict) -> bool:
        if not isinstance(payload, dict) or payload.get("version") != MEMORY_VERSION:
            return False
        self.training_rows = int(payload.get("training_rows", 0) or 0)
        self.training_files = int(payload.get("training_files", 0) or 0)
        self.global_exact = self._payload_to_counter_map(payload.get("global_exact"))
        self.global_family = self._payload_to_counter_map(payload.get("global_family"))
        self.client_exact = self._payload_to_tuple_counter_map(payload.get("client_exact"))
        self.client_family = self._payload_to_tuple_counter_map(payload.get("client_family"))
        self.context_counts = self._payload_to_counter_map(payload.get("context_counts"))
        self.client_context_counts = self._payload_to_tuple_counter_map(payload.get("client_context_counts"))
        self.token_counts = self._payload_to_counter_map(payload.get("token_counts"))
        self.bigram_counts = self._payload_to_counter_map(payload.get("bigram_counts"))
        self.code_counts = self._dict_to_counter(payload.get("code_counts"))
        self.account_descriptions = {
            str(key): str(value)
            for key, value in (payload.get("account_descriptions") or {}).items()
        }
        self.source_fingerprint = {
            str(key): {
                str(meta_key): int(meta_value or 0)
                for meta_key, meta_value in (value or {}).items()
                if str(meta_key) in {
                    "mtime_ns",
                    "size",
                    "sync_error_mtime_ns",
                    "sync_error_size",
                }
            }
            for key, value in (payload.get("source_fingerprint") or {}).items()
            if isinstance(value, dict)
        }
        return True

    def _load_memory(self) -> bool:
        if not self.memory_path.exists():
            return False
        try:
            with open(self.memory_path, "r", encoding="utf-8") as handle:
                payload = json.load(handle)
        except Exception:
            return False
        return self._load_payload(payload)

    def save_memory(self) -> bool:
        temp_path = self.memory_path.with_name(
            f".{self.memory_path.name}.{os.getpid()}.{uuid.uuid4().hex}.tmp"
        )
        try:
            self.memory_path.parent.mkdir(parents=True, exist_ok=True)
            with open(temp_path, "w", encoding="utf-8") as handle:
                json.dump(
                    self._to_payload(),
                    handle,
                    ensure_ascii=False,
                    separators=(",", ":"),
                )
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temp_path, self.memory_path)
            self.memory_saved = True
            return True
        except Exception:
            return False
        finally:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass

    def _target_from_training_row(self, row) -> str:
        debit = clean_code(row.get("DEBIT"))
        credit = clean_code(row.get("CREDIT"))
        if not debit or debit == "7810":
            return ""
        sign = str(row.get("Positive/Negative", "")).strip()
        if sign == "-":
            return debit
        if sign == "+":
            return ""
        if credit in COUNTER_ACCOUNT_CODES and debit not in {"7810"}:
            return debit
        return ""

    def _load_account_descriptions(self) -> dict[str, str]:
        descriptions: dict[str, str] = {}
        if not self.key_dir.exists():
            return descriptions
        for path in sorted(self.key_dir.rglob("*")):
            if path.name.startswith("~$") or not path.is_file():
                continue
            suffix = path.suffix.lower()
            if suffix in {".xlsx", ".xls", ".xlsm"}:
                descriptions.update(self._load_account_descriptions_from_excel(path))
            elif suffix == ".csv":
                descriptions.update(self._load_account_descriptions_from_csv(path))
            elif suffix == ".pdf":
                descriptions.update(self._load_account_descriptions_from_pdf(path))
        return descriptions

    def _load_account_descriptions_from_excel(self, path: Path) -> dict[str, str]:
        try:
            df = pd.read_excel(path)
        except Exception:
            return {}
        code_col = None
        desc_col = None
        for col in df.columns:
            norm = normalize_text(col)
            if code_col is None and any(word in norm for word in ("lykill", "code", "konto", "reikningur")):
                code_col = col
            if desc_col is None and any(word in norm for word in ("lysing", "description", "heiti", "name")):
                desc_col = col
        if code_col is None:
            code_col = df.columns[0] if len(df.columns) else None
        if desc_col is None and len(df.columns) > 1:
            desc_col = df.columns[1]
        if code_col is None or desc_col is None:
            return {}
        descriptions = {}
        for _, row in df.iterrows():
            code = clean_code(row.get(code_col))
            desc = str(row.get(desc_col, "")).strip()
            if code and desc and desc.lower() != "nan":
                descriptions[code] = desc
        return descriptions

    def _load_account_descriptions_from_csv(self, path: Path) -> dict[str, str]:
        try:
            df = pd.read_csv(path)
        except Exception:
            return {}
        tmp = path.with_suffix(".xlsx")
        # Reuse the same column-detection logic without writing any files.
        code_col = df.columns[0] if len(df.columns) else None
        desc_col = df.columns[1] if len(df.columns) > 1 else None
        descriptions = {}
        if code_col is None or desc_col is None:
            return descriptions
        for _, row in df.iterrows():
            code = clean_code(row.get(code_col))
            desc = str(row.get(desc_col, "")).strip()
            if code and desc and desc.lower() != "nan":
                descriptions[code] = desc
        return descriptions

    def _load_account_descriptions_from_pdf(self, path: Path) -> dict[str, str]:
        try:
            from pypdf import PdfReader
        except Exception:
            return {}
        try:
            text = "\n".join(page.extract_text() or "" for page in PdfReader(str(path)).pages)
        except Exception:
            return {}
        descriptions = {}
        for line in text.splitlines():
            match = re.match(r"^\s*(\d{3,5})\s+(.+?)\s+[RE]\s+", line)
            if match:
                descriptions[match.group(1)] = match.group(2).strip()
        return descriptions

    def ready(self) -> bool:
        return self.training_rows > 0

    def summary(self) -> dict[str, int]:
        return {
            "training_files": self.training_files,
            "training_rows": self.training_rows,
            "merchant_keys": len(self.global_exact),
            "families": len(self.global_family),
            "context_features": len(self.context_counts),
            "account_descriptions": len(self.account_descriptions),
            "loaded_from_memory": int(self.loaded_from_memory),
            "memory_saved": int(self.memory_saved),
        }

    def predict(
        self,
        text,
        signal_text: str = "",
        amount=None,
        date=None,
        id_value=None,
        repeat_count=None,
        client_hint: str = "",
        industry_context: str = "",
        context_flags=None,
    ) -> AutoCodeResult:
        key = merchant_key(text)
        family = merchant_family(text)
        client = client_key_from_path(client_hint) if client_hint else ""
        person_like = looks_like_person_name(text) or looks_like_person_id(id_value)

        mandatory_rule = _hardcoded_bank_signal_rule(str(text or ""), signal_text)
        if mandatory_rule:
            mandatory_rule.matched_text = key
            if mandatory_rule.code in self.account_descriptions:
                mandatory_rule.reason = (
                    f"{mandatory_rule.reason}; "
                    f"{mandatory_rule.code} {self.account_descriptions[mandatory_rule.code]}"
                )
            return mandatory_rule

        candidates: list[AutoCodeResult] = []

        consistent_history = self._consistent_exact_history(key, client)
        if consistent_history:
            consistent_history.matched_text = key
            if consistent_history.code in self.account_descriptions:
                consistent_history.reason = f"{consistent_history.reason}; {consistent_history.code} {self.account_descriptions[consistent_history.code]}"
            return consistent_history

        exact_history_is_mixed = self._exact_history_is_mixed(key, client)

        if client and key:
            candidates.append(self._history_result(
                self.client_exact.get((client, key)),
                f"client merchant history: {key}",
                "client_exact",
                exact=True,
                minimum_total=1,
            ))
        if key:
            candidates.append(self._history_result(
                self.global_exact.get(key),
                f"merchant history: {key}",
                "global_exact",
                exact=True,
                minimum_total=1,
            ))
        if client and family:
            candidates.append(self._history_result(
                self.client_family.get((client, family)),
                f"client family history: {family}",
                "client_family",
                exact=False,
                minimum_total=2,
            ))
        if family:
            candidates.append(self._history_result(
                self.global_family.get(family),
                f"merchant family history: {family}",
                "global_family",
                exact=False,
                minimum_total=3,
            ))

        rule = _category_rule(
            str(text or ""),
            signal_text,
            amount=amount,
            industry_context=industry_context,
            context_flags=context_flags,
        )
        if rule:
            candidates.append(rule)

        context_result = self._context_result(
            str(text or ""),
            amount=amount,
            date=date,
            id_value=id_value,
            repeat_count=repeat_count,
            client_hint=client_hint,
        )
        if context_result:
            candidates.append(context_result)

        if not person_like:
            token_result = self._token_result(str(text or ""), signal_text)
            if token_result:
                candidates.append(token_result)

        candidates = [c for c in candidates if c and c.code]
        if exact_history_is_mixed and context_result and context_result.code and context_result.confidence >= 55:
            context_result.matched_text = key
            if context_result.code in self.account_descriptions:
                context_result.reason = f"{context_result.reason}; {context_result.code} {self.account_descriptions[context_result.code]}"
            return context_result

        if not candidates:
            if person_like:
                return AutoCodeResult(
                    code="9410",
                    confidence=42,
                    status="review",
                    reason="person-like payee default wage review",
                    source="person_default",
                    matched_text=key,
                )
            return AutoCodeResult(
                code="4510",
                confidence=20,
                status="review",
                reason="fallback generic outgoing; no confident match",
                source="fallback",
                matched_text=key,
            )

        candidates.sort(key=lambda c: (c.confidence, c.source == "client_exact", c.source == "global_exact"), reverse=True)
        best = candidates[0]
        best.matched_text = key
        if best.code in self.account_descriptions:
            best.reason = f"{best.reason}; {best.code} {self.account_descriptions[best.code]}"
        return best

    def _consistent_exact_history(self, key: str, client: str = "") -> AutoCodeResult | None:
        checks = []
        if client and key:
            checks.append((self.client_exact.get((client, key)), "client merchant consistent history", "client_exact_consistent"))
        if key:
            checks.append((self.global_exact.get(key), "merchant consistent history", "global_exact_consistent"))
        for counter, reason, source in checks:
            if not counter:
                continue
            code, count, total, ratio = _best(counter)
            if total < 2:
                continue
            if ratio >= 0.98 or (total >= 5 and ratio >= 0.94):
                confidence = min(99, _confidence_from_history(count, total, ratio, exact=True) + 3)
                return AutoCodeResult(
                    code=code,
                    confidence=confidence,
                    status="coded",
                    reason=f"{reason}: {key} ({count}/{total})",
                    source=source,
                )
        return None

    def _exact_history_is_mixed(self, key: str, client: str = "") -> bool:
        counters = []
        if client and key:
            counters.append(self.client_exact.get((client, key)))
        if key:
            counters.append(self.global_exact.get(key))
        for counter in counters:
            if not counter:
                continue
            _code, _count, total, ratio = _best(counter)
            if total >= 2 and len(counter) >= 2 and ratio < 0.98:
                return True
        return False

    def _history_result(self, counter, reason: str, source: str, exact: bool, minimum_total: int) -> AutoCodeResult:
        if not counter:
            return AutoCodeResult()
        code, count, total, ratio = _best(counter)
        if total < minimum_total:
            return AutoCodeResult()
        if total >= 2 and ratio < (0.58 if exact else 0.66):
            return AutoCodeResult(
                code=code,
                confidence=max(35, int(ratio * 70)),
                status="review",
                reason=f"ambiguous {reason} ({count}/{total})",
                source=source,
            )
        confidence = _confidence_from_history(count, total, ratio, exact)
        if total == 1:
            confidence = min(confidence, 72 if exact else 58)
        status = "coded" if confidence >= 55 else "review"
        return AutoCodeResult(
            code=code,
            confidence=confidence,
            status=status,
            reason=f"{reason} ({count}/{total})",
            source=source,
        )

    @staticmethod
    def _context_feature_weight(feature: str) -> float:
        if feature.startswith("key:"):
            weight = 1.35
        elif feature.startswith("family:"):
            weight = 1.05
        elif feature.startswith("type:person"):
            weight = 0.50
        else:
            weight = 0.75
        if "|amount:" in feature:
            weight += 0.12
        if "|day:" in feature or "|day_bucket:" in feature or "|day_window:" in feature:
            weight += 0.10
        if "|month:" in feature or "|month_parity:" in feature:
            weight += 0.12
        if "|repeat:" in feature:
            weight += 0.14
        if feature.count("|") >= 2:
            weight += 0.15
        return weight

    @staticmethod
    def _context_thresholds(feature: str, client_specific: bool) -> tuple[int, float]:
        if feature == "type:person":
            return (8, 0.70)
        if feature.startswith("type:person"):
            return ((2, 0.58) if client_specific else (4, 0.62))
        if "|month_parity:" in feature and ("|day:" in feature or "|day_window:" in feature):
            return ((2, 0.62) if client_specific else (4, 0.68))
        if ("|day:" in feature or "|day_bucket:" in feature or "|day_window:" in feature) and "|amount:" not in feature and "|repeat:" not in feature:
            return ((2, 0.65) if client_specific else (4, 0.70))
        if feature.startswith("key:"):
            return ((1, 0.55) if client_specific else (2, 0.58))
        if feature.startswith("family:"):
            return ((2, 0.58) if client_specific else (3, 0.62))
        return ((2, 0.60) if client_specific else (4, 0.65))

    def _context_result(
        self,
        text: str,
        amount=None,
        date=None,
        id_value=None,
        repeat_count=None,
        client_hint: str = "",
    ) -> AutoCodeResult | None:
        features = _context_features(text, amount=amount, date=date, repeat_count=repeat_count, id_value=id_value)
        if not features:
            return None

        client = client_key_from_path(client_hint) if client_hint else ""
        scores = Counter()
        evidence = []
        strongest_feature = ""

        def add_counter(feature: str, counter, client_specific: bool):
            nonlocal strongest_feature
            if not counter:
                return
            code, count, total, ratio = _best(counter)
            min_total, min_ratio = self._context_thresholds(feature, client_specific)
            if total < min_total or ratio < min_ratio:
                return
            source_weight = 1.25 if client_specific else 1.0
            feature_weight = self._context_feature_weight(feature) * source_weight
            score_weight = feature_weight * ratio * math.log1p(total)
            for candidate_code, candidate_count in counter.items():
                scores[candidate_code] += score_weight * (candidate_count / total)
            label = "client context" if client_specific else "context"
            evidence.append(f"{label}: {feature} -> {code} ({count}/{total})")
            if not strongest_feature or feature_weight > self._context_feature_weight(strongest_feature):
                strongest_feature = feature

        for feature in features:
            if client:
                add_counter(feature, self.client_context_counts.get((client, feature)), True)
            add_counter(feature, self.context_counts.get(feature), False)

        if not scores:
            return None

        ranked = scores.most_common(2)
        code, top = ranked[0]
        second = ranked[1][1] if len(ranked) > 1 else 0
        margin = (top - second) / top if top else 0
        if top < 1.05 or margin < 0.12:
            return None

        confidence_cap = 94 if any(
            ("|month_parity:" in item and ("|day:" in item or "|day_window:" in item))
            for item in evidence
        ) else 88
        confidence = int(min(confidence_cap, 44 + top * 5.5 + margin * 16 + min(6, len(evidence))))
        if strongest_feature.startswith("type:person") and not any(
            item.startswith("client context: key:") or item.startswith("context: key:")
            or item.startswith("client context: family:") or item.startswith("context: family:")
            for item in evidence
        ):
            confidence = min(confidence, 68)
        return AutoCodeResult(
            code=code,
            confidence=max(35, confidence),
            status="coded" if confidence >= 55 else "review",
            reason="; ".join(evidence[:4]),
            source="context_model",
        )

    def _token_result(self, text: str, signal_text: str) -> AutoCodeResult | None:
        tokens = tokenize(text, drop_processors=True)
        signal_tokens = tokenize(signal_text, keep_generic=True, drop_processors=True)
        if not tokens and not signal_tokens:
            return None
        scores = Counter()
        evidence = []
        all_tokens = tokens + signal_tokens
        for token in set(all_tokens):
            counter = self.token_counts.get(token)
            if not counter:
                continue
            code, count, total, ratio = _best(counter)
            if total < 2 or ratio < 0.55:
                continue
            weight = ratio * math.log1p(total)
            if token in signal_tokens:
                weight *= 0.65
            for c, n in counter.items():
                scores[c] += weight * (n / total)
            evidence.append(f"{token}:{code}")
        for i in range(len(tokens) - 1):
            bigram = " ".join(tokens[i:i + 2])
            counter = self.bigram_counts.get(bigram)
            if not counter:
                continue
            code, count, total, ratio = _best(counter)
            if ratio < 0.55:
                continue
            weight = ratio * math.log1p(total) * 1.4
            for c, n in counter.items():
                scores[c] += weight * (n / total)
            evidence.append(f"{bigram}:{code}")
        if not scores:
            return None
        ranked = scores.most_common(2)
        code, top = ranked[0]
        second = ranked[1][1] if len(ranked) > 1 else 0
        margin = (top - second) / top if top else 0
        if top < 1.15 or margin < 0.18:
            return None
        confidence = int(min(74, 48 + top * 7 + margin * 16))
        return AutoCodeResult(
            code=code,
            confidence=confidence,
            status="coded" if confidence >= 55 else "review",
            reason="token model: " + ", ".join(evidence[:4]),
            source="token_model",
        )


_CACHED_CODER: TransactionAutoCoder | None = None
_CACHED_CODER_KEY = None


def get_auto_coder(
    force_reload: bool = False,
    training_dir=None,
    key_dir=None,
    memory_path=None,
    training_data_dir=None,
    force_rebuild: bool = False,
) -> TransactionAutoCoder:
    global _CACHED_CODER, _CACHED_CODER_KEY
    cache_key = (
        str(Path(training_dir).resolve()) if training_dir else "",
        str(Path(key_dir).resolve()) if key_dir else "",
        str(Path(memory_path).resolve()) if memory_path else "",
        str(Path(training_data_dir).resolve()) if training_data_dir else "",
    )
    if force_reload or force_rebuild or _CACHED_CODER is None or _CACHED_CODER_KEY != cache_key:
        _CACHED_CODER = TransactionAutoCoder(
            training_dir=training_dir,
            key_dir=key_dir,
            memory_path=memory_path,
            training_data_dir=training_data_dir,
            force_rebuild=force_rebuild,
        )
        _CACHED_CODER_KEY = cache_key
    return _CACHED_CODER


def apply_auto_debit_codes(
    output_df: pd.DataFrame,
    source_df: pd.DataFrame | None = None,
    input_file_path: str = "",
    enabled: bool = True,
    fill_threshold: int = 35,
    overwrite_existing: bool = False,
    training_dir=None,
    key_dir=None,
    memory_path=None,
    training_data_dir=None,
    industry_context: str = "",
    vehicle_expense_code: str = "",
) -> pd.DataFrame:
    df = output_df.copy()
    if not enabled:
        df["STATUS"] = "review needed"
        df["CONFIDENCE"] = ""
        df["_AUTO_CODE_SOURCE"] = "disabled"
        return df

    context_flags_by_index = _row_context_flags(df, source_df=source_df)
    mandatory_results = {}
    for idx, row in df.iterrows():
        amount = _parse_amount(row.get("AMOUNT"))
        if not _is_outgoing_row(row, amount):
            continue
        if _existing_debit_is_protected(row, overwrite_existing):
            continue
        source_row = _source_row_for_output(source_df, row, idx)
        mandatory = _hardcoded_bank_signal_rule(
            str(row.get("TEXT", "") or ""),
            _source_rule_text(source_row),
        )
        if mandatory is None:
            mandatory = _collection_cost_context_rule(context_flags_by_index.get(idx, set()))
        if mandatory:
            df.at[idx, "DEBIT"] = mandatory.code
            mandatory_results[idx] = mandatory

    coder = get_auto_coder(
        training_dir=training_dir,
        key_dir=key_dir,
        memory_path=memory_path,
        training_data_dir=training_data_dir,
    )
    if not coder.ready():
        statuses = []
        confidences = []
        sources = []
        for idx, row in df.iterrows():
            if idx in mandatory_results:
                mandatory = mandatory_results[idx]
                statuses.append("coded")
                confidences.append(_output_confidence(mandatory))
                sources.append(mandatory.source)
            elif _existing_debit_is_protected(row, overwrite_existing):
                marker = _existing_confidence_value(row)
                source = str(row.get("_AUTO_CODE_SOURCE", "") or "").strip()
                statuses.append(str(row.get("STATUS", "") or "coded"))
                confidences.append(marker)
                sources.append(source or ("manual_review" if str(marker).upper() == "MN" else "existing"))
            elif _is_outgoing_row(row):
                statuses.append("review needed")
                confidences.append("")
                sources.append("unavailable")
            else:
                statuses.append("")
                confidences.append("")
                sources.append("")
        df["STATUS"] = statuses
        df["CONFIDENCE"] = confidences
        df["_AUTO_CODE_SOURCE"] = sources
        return df

    repeat_counts = Counter()
    for _, row in df.iterrows():
        amount = _parse_amount(row.get("AMOUNT"))
        if not _is_outgoing_row(row, amount):
            continue
        key = merchant_key(row.get("TEXT", ""))
        if key:
            repeat_counts[key] += 1

    statuses = []
    confidences = []
    sources = []
    for idx, row in df.iterrows():
        sign = str(row.get("Positive/Negative", "")).strip()
        amount = _parse_amount(row.get("AMOUNT"))
        if sign != "-":
            if sign == "" and amount is not None and amount < 0:
                sign = "-"
            else:
                statuses.append("")
                confidences.append("")
                sources.append("")
                continue

        if idx in mandatory_results:
            mandatory = mandatory_results[idx]
            statuses.append("coded")
            confidences.append(_output_confidence(mandatory))
            sources.append(mandatory.source)
            continue

        existing = clean_code(row.get("DEBIT"))
        if existing and _existing_debit_is_protected(row, overwrite_existing):
            marker = _existing_confidence_value(row)
            source = str(row.get("_AUTO_CODE_SOURCE", "") or "").strip()
            statuses.append(str(row.get("STATUS", "") or "coded"))
            confidences.append(marker)
            sources.append(source or ("manual_review" if str(marker).upper() == "MN" else "existing"))
            continue

        source_row = _source_row_for_output(source_df, row, idx)
        signal_text = _source_signal_text(source_row)
        text_value = row.get("TEXT", "")
        result = coder.predict(
            text_value,
            signal_text=signal_text,
            amount=row.get("AMOUNT"),
            date=row.get("DATE"),
            id_value=row.get("ID"),
            repeat_count=repeat_counts.get(merchant_key(text_value), 1),
            client_hint=input_file_path,
            industry_context=industry_context,
            context_flags=context_flags_by_index.get(idx, set()),
        )
        result = _cap_unrecognized_treasury_confidence(
            result,
            str(text_value or ""),
            _source_rule_text(source_row),
        )
        result = _apply_vehicle_expense_code_choice(result, vehicle_expense_code)
        should_fill = result.code and (result.confidence >= fill_threshold or result.source == "fallback")
        if should_fill:
            df.at[idx, "DEBIT"] = result.code
            if result.source == "fallback":
                statuses.append("review needed")
            else:
                statuses.append("coded" if result.confidence >= 55 else "review needed")
        elif result.code:
            statuses.append("review needed")
        else:
            statuses.append("review needed")
        confidences.append(_output_confidence(result))
        sources.append(result.source or "")

    df["STATUS"] = statuses
    df["CONFIDENCE"] = confidences
    df["_AUTO_CODE_SOURCE"] = sources
    return df

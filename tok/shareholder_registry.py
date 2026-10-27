"""Modern, read-only shareholder-registry UI for Tok Tenging.

The workbook parsing and entity reconciliation live in the companion data-model
module.  This module deliberately contains presentation and interaction logic
only: asynchronous loading, safe search, source inspection, and credential
reveal/copy controls.
"""

from __future__ import annotations

import importlib
import os
from pathlib import Path
import queue
import re
import threading
import tkinter as tk
from tkinter import filedialog, ttk
import unicodedata

try:
    from . import ui_components
    from .ui_components import (
        create_panel,
        create_styled_button,
        create_styled_entry,
        create_styled_label,
        set_button_accent,
        _rounded_rect,
    )
except ImportError:  # Direct-script mode used by TokTenging.py.
    import ui_components
    from ui_components import (
        create_panel,
        create_styled_button,
        create_styled_entry,
        create_styled_label,
        set_button_accent,
        _rounded_rect,
    )


globals().update(ui_components.get_theme("light"))

REGISTRY_SHEET_NAMES = (
    "Hlutafjármiðar 2025",
    "Hluthafaskrá",
    "Áreiðanleikakönnun allir",
)
DEFAULT_WORKBOOK_FILENAME = "data codex.xlsx"
DEFAULT_WORKBOOK_RELATIVE_PATH = os.path.join(
    "tok", "trainingData", "data folder", DEFAULT_WORKBOOK_FILENAME
)
LEGACY_WORKBOOK_RELATIVE_PATH = os.path.join(
    "tok", "data folder", DEFAULT_WORKBOOK_FILENAME
)
RAW_PAGE_SIZE = 250
SEARCH_DEBOUNCE_MS = 160
SECRET_REMASK_MS = 30_000

MISSING_VALUE_LABEL = "Ekki skráð"
ALL_RECORDS_LABEL = "Allar færslur"
ALL_SHEETS_LABEL = "Öll þrjú blöðin"

_SECTION_LABELS = {
    "2025 shareholder returns": "Hlutafjármiðar 2025",
    "Búferlar breakdown": "B — Nánari sundurliðun",
    "Búferlar owner detail": "Nánari sundurliðun eigenda — Búferlar",
    "Búferlar breakdown and unlabelled calculations": "Nánari sundurliðun og ómerktir útreikningar — Búferlar",
    "A — Companies in due diligence": "A — Hlutafélög í áreiðanleikakönnun",
    "B — Búferlar breakdown": "B — Nánari sundurliðun — Búferlar",
    "C — Companies not due-diligence checked": "C — Hlutafélög ekki áreiðanleikakönnuð",
    "D — SLF entities not due-diligence checked": "D — SLF félög — ekki áreiðanleikakönnuð",
    "E — Companies handled by Anna Bragina": "E — Félög sem Anna Bragina færir og ber ábyrgð á",
    "E — Companies handled by Fribba, Íris and Alli": "E — Félög sem Fribba, Íris og Alli færa og bera ábyrgð á",
    "F — Individuals with operations": "F — Einstaklingar með rekstur",
    "G — Individuals not due-diligence checked": "G — Einstaklingar með rekstur — ekki áreiðanleikakannað",
    "Worked for third parties — not due-diligence checked": "Unnið fyrir þriðja aðila — ekki áreiðanleikakannað",
    "SLF entities": "SLF félög",
    "Summary counts": "Samantekt",
}

_FIELD_LABELS = {
    "Unlabelled": "Án fyrirsagnar",
    "Breakdown": "SUNDURLIÐUN",
    "Owner": "Eigandi",
    "Owner ID / country": "Kennitala eiganda / land",
    "Share": "Hlutur",
    "Percentage": "Hlutfall",
    "Access key": "Veflykill",
}

_MATCH_REASON_LABELS = {
    "Exact ID": "Nákvæm kennitala",
    "Exact company name": "Nákvæmt nafn",
    "Company name starts with search": "Leitartexti í upphafi nafns",
    "Company name": "Nafn",
    "Alternate name": "Annað nafn",
    "Related public record": "Tengd opinber færsla",
    "kennitala": "Kennitala",
    "id": "Kennitala",
    "name": "Nafn",
    "name prefix": "Leitartexti í upphafi nafns",
    "alias": "Annað nafn",
    "email": "Netfang",
    "phone": "Sími",
    "shareholder": "Hluthafi",
}

_WARNING_LABELS = {
    "Columns I–K on this row are a separate unlabelled side calculation.":
        "Í dálkum I–K í þessari línu er sérstakur útreikningur án fyrirsagnar.",
    "The source header is damaged and some appended rows use shifted columns.":
        "Haus frumgagnanna er skemmdur og dálkar í sumum viðbættum línum hafa hliðrast.",
    "Unrelated cells I–K remain available in the global raw-source explorer.":
        "Ótengdir reitir I–K eru áfram aðgengilegir í yfirliti frumgagna.",
    "Owner identifier was stored in a share column.":
        "Kennitala eiganda var skráð í dálk fyrir hlut.",
    "Owner name was stored in a share column.":
        "Nafn eiganda var skráð í dálk fyrir hlut.",
    "Owner/share columns appear shifted; values were paired by type.":
        "Dálkar eiganda og hlutar virðast hafa hliðrast; gildi voru pöruð eftir tegund.",
    "The cached share formula contains an Excel error.":
        "Vistaða formúlan fyrir hlut inniheldur Excel-villu.",
    "No valid kennitala is recorded; this entity remains linked to its source row.":
        "Engin gild kennitala er skráð; færslan er áfram tengd við upprunalínuna.",
    "The source sheets contain more than one name spelling for this identity.":
        "Nafn sama aðila er ritað á fleiri en einn hátt í upprunablöðunum.",
    "Share-capital values conflict between source observations.":
        "Gildi hlutafjár stangast á milli upprunafærslna.",
    "Access credentials conflict between source observations; all values remain source-labelled.":
        "Veflyklar stangast á milli upprunafærslna; hvert gildi er áfram merkt uppruna sínum.",
}

root = None
frame = None
registry_workbook_path = None


def _not_configured(*_args, **_kwargs):
    raise RuntimeError("shareholder_registry.configure() must be called before opening the registry")


set_page_title = _not_configured
initialize_main_menu = _not_configured
resource_path = lambda relative: relative
open_file_location = lambda _path: None
load_registry_data = None
_training_data_dir = lambda: str(Path(__file__).resolve().parent / "trainingData")


def _default_window_work_area(window):
    try:
        left = int(window.winfo_vrootx())
        top = int(window.winfo_vrooty())
        return left, top, left + int(window.winfo_vrootwidth()), top + int(window.winfo_vrootheight())
    except Exception:
        return 0, 0, int(window.winfo_screenwidth()), int(window.winfo_screenheight())


window_work_area = _default_window_work_area

_active_page = None


def configure(**dependencies):
    """Inject application-shell dependencies without creating import cycles."""
    globals().update(dependencies)


def _geometry_spec(width, height, x, y):
    x_part = f"+{int(x)}" if int(x) >= 0 else str(int(x))
    y_part = f"+{int(y)}" if int(y) >= 0 else str(int(y))
    return f"{int(width)}x{int(height)}{x_part}{y_part}"


def _capture_window_layout(window):
    """Capture the window layout before the registry temporarily enlarges it."""
    if window is None:
        return None
    try:
        window.update_idletasks()
        return {
            "state": str(window.state()),
            "geometry": _geometry_spec(
                max(1, int(window.winfo_width())),
                max(1, int(window.winfo_height())),
                int(window.winfo_x()),
                int(window.winfo_y()),
            ),
            "minsize": tuple(int(value) for value in window.minsize()),
        }
    except (tk.TclError, TypeError, ValueError):
        return None


def _restore_window_layout(window, layout):
    """Restore a layout captured by :func:`_capture_window_layout`."""
    if window is None or not layout:
        return
    try:
        minimum = layout.get("minsize")
        if minimum and len(minimum) == 2:
            window.minsize(*minimum)
        previous_state = str(layout.get("state") or "normal")
        if previous_state == "zoomed":
            window.state("zoomed")
            return
        if window.state() == "zoomed":
            window.state("normal")
        geometry = layout.get("geometry")
        if geometry:
            window.geometry(geometry)
        if previous_state not in {"", "normal", "fullscreen"}:
            window.state(previous_state)
    except (tk.TclError, TypeError, ValueError):
        pass


def set_theme(theme):
    """Receive the active shared theme from the application shell."""
    if isinstance(theme, str):
        theme = ui_components.get_theme(theme)
    globals().update(theme)


def _get(obj, *names, default=None):
    """Read the first available attribute/key from dataclasses or mappings."""
    if obj is None:
        return default
    for name in names:
        if isinstance(obj, dict) and name in obj:
            value = obj[name]
        elif hasattr(obj, name):
            value = getattr(obj, name)
        else:
            continue
        if value is not None:
            return value
    return default


def _items(value):
    if value is None:
        return []
    if isinstance(value, dict):
        return list(value.values())
    if isinstance(value, (str, bytes)):
        return [value]
    try:
        return list(value)
    except TypeError:
        return [value]


def _text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _normalize(value):
    text = _text(value).casefold()
    text = text.translate(str.maketrans({"ð": "d", "þ": "th", "æ": "ae", "ö": "o"}))
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    return " ".join(re.findall(r"[a-z0-9]+", text))


def _digits(value):
    return "".join(re.findall(r"\d", _text(value)))


def _localize_section(value):
    text = _text(value)
    return _SECTION_LABELS.get(text, text)


def _localize_field_label(value):
    text = _text(value)
    if text in _FIELD_LABELS:
        return _FIELD_LABELS[text]
    match = re.fullmatch(r"Unlabelled\s*\(([^)]+)\)", text, flags=re.IGNORECASE)
    if match:
        return f"Án fyrirsagnar ({match.group(1)})"
    match = re.fullmatch(r"Owner\s+(\d+)", text, flags=re.IGNORECASE)
    if match:
        return f"Eigandi {match.group(1)}"
    match = re.fullmatch(r"Shareholder\s+(\d+)", text, flags=re.IGNORECASE)
    if match:
        return f"Hluthafi {match.group(1)}"
    return text


def _localize_warning(value):
    text = _text(value)
    if text in _WARNING_LABELS:
        return _WARNING_LABELS[text]
    match = re.fullmatch(r"Owner identifier is a country/note rather than a kennitala: (.+)\.", text)
    if match:
        return f"Auðkenni eiganda er land eða athugasemd en ekki kennitala: {match.group(1)}."
    match = re.fullmatch(r"(\d+) source rows share this identity in (.+); none were overwritten\.", text)
    if match:
        return f"{match.group(1)} línur í frumgögnum hafa sama auðkenni á blaðinu {match.group(2)}; engu var skrifað yfir."
    return text


def _localize_relationship_role(value):
    text = _text(value).replace("_", " ")
    normalized = _normalize(text)
    if normalized == "shareholder":
        return "Hluthafi"
    match = re.fullmatch(r"shareholder\s+(\d+)", text, flags=re.IGNORECASE)
    if match:
        return f"Hluthafi {match.group(1)}"
    if normalized == "owns shares in":
        return "Á hlut í"
    if normalized == "owner":
        return "Eigandi"
    return text


def _entity_kind_key(entity):
    return _normalize(_get(entity, "entity_type", "kind", "record_type", "type", default="company"))


def _bind_responsive_wrap(label, container, horizontal_padding=28, minimum=140):
    """Keep wrapped copy inside the surface that owns it as windows resize."""
    def sync(event):
        try:
            label.configure(wraplength=max(minimum, int(event.width) - horizontal_padding))
        except tk.TclError:
            pass

    container.bind("<Configure>", sync, add="+")
    return label


def _format_kennitala(value):
    digits = _digits(value)
    if len(digits) == 10:
        return f"{digits[:6]}-{digits[6:]}"
    return _text(value) or MISSING_VALUE_LABEL


def _format_amount(value):
    if value in (None, ""):
        return MISSING_VALUE_LABEL
    if isinstance(value, bool):
        return _text(value)
    try:
        number = float(value)
    except (TypeError, ValueError):
        return _text(value)
    if number.is_integer():
        rendered = f"{int(number):,}".replace(",", ".")
    else:
        rendered = f"{number:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{rendered} kr."


def _format_percent(value):
    if value in (None, ""):
        return ""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return _text(value)
    if abs(number) <= 1:
        number *= 100
    return (f"{number:.2f}".rstrip("0").rstrip(".") + "%").replace(".", ",")


def _format_count(value):
    try:
        return f"{int(value):,}".replace(",", ".")
    except (TypeError, ValueError):
        return _text(value)


def _source_sheet_count(value):
    count = int(value)
    return "1 upprunablað" if count == 1 else f"{_format_count(count)} upprunablöð"


def _entity_name(entity):
    return _text(_get(entity, "canonical_name", "display_name", "name", "company_name", default="Ónefnd færsla"))


def _entity_kennitala(entity):
    value = _get(entity, "normalized_kennitala", "kennitala", "company_id", "id_number")
    if not value:
        entity_id = _text(_get(entity, "entity_id", "key"))
        match = re.search(r"(?<!\d)(\d{10})(?!\d)", entity_id)
        value = match.group(1) if match else ""
    return _digits(value) or _text(value)


def _entity_key(entity):
    key = _get(entity, "entity_id", "key")
    if key not in (None, ""):
        return _text(key)
    return _entity_kennitala(entity) or f"name:{_normalize(_entity_name(entity))}:{id(entity)}"


def _entity_kind(entity):
    value = _text(_get(entity, "entity_type", "kind", "record_type", "type", default="company"))
    normalized = _normalize(value)
    if normalized in {"company", "fyrirtaeki", "hlutafelag"}:
        return "Hlutafélag"
    if normalized == "slf":
        return "SLF félag"
    if normalized in {"partnership", "sameignarfelag"}:
        return "Sameignarfélag"
    if normalized in {"individual", "person", "einstaklingur"}:
        return "Einstaklingur"
    return value.replace("_", " ") or "Hlutafélag"


def _entity_aliases(entity):
    values = _get(entity, "aliases", "name_aliases", "alternate_names", default=[])
    return [_text(value) for value in _items(values) if _text(value) and _text(value) != _entity_name(entity)]


def _entity_statuses(entity):
    values = _get(entity, "statuses", "categories", "classifications", "tags", "status", default=[])
    output = [_text(value) for value in _items(values) if _text(value)]
    if not output:
        output = list(dict.fromkeys(_text(_get(row, "section")) for row in _entity_source_records(entity) if _text(_get(row, "section"))))
    return output


def _entity_source_records(entity):
    records = _items(_get(entity, "source_rows", "source_records", "source_observations", "observations", "records", default=[]))
    details = _items(_get(entity, "detail_rows", default=[]))
    seen = {id(record) for record in records}
    records.extend(record for record in details if id(record) not in seen)
    return records


def _entity_credentials(entity):
    explicit = _items(_get(entity, "credentials", "access_credentials", "veflyklar", default=[]))
    if explicit:
        return explicit
    return [cell for cell in _record_cells(entity) if _cell_sensitive(cell)]


def _entity_ownerships(entity):
    return _items(_get(entity, "ownerships", "shareholders", "owners", "ownership", "related_parties", default=[]))


def _entity_reverse_ownerships(entity):
    return _items(_get(entity, "reverse_ownerships", "owned_companies", "investments", "owns_shares_in", default=[]))


def _model_entities(model):
    value = _get(model, "entities", "companies", "records", default=[])
    if callable(value):
        value = value()
    return _items(value)


def _model_raw_cells(model):
    value = _get(model, "raw_cells", "all_cells", "source_cells", default=[])
    if callable(value):
        value = value()
    if value:
        return _items(value)
    iterator = _get(model, "iter_raw_cells", "iter_cells")
    if callable(iterator):
        return list(iterator())
    row_iterator = _get(model, "iter_source_rows")
    if callable(row_iterator):
        output = []
        for row in row_iterator():
            output.extend(_cells_from_record(row))
        return output
    return []


def _source_name(value):
    return _text(_get(value, "sheet_name", "sheet", "source", "source_sheet", default="Óþekkt upprunablað"))


def _cell_coordinate(cell):
    coordinate = _text(_get(cell, "coordinate", "cell", "excel_cell"))
    if coordinate:
        return coordinate
    row = _get(cell, "row", "excel_row")
    column = _get(cell, "column", "column_letter")
    return f"{column or ''}{row or ''}" or "—"


def _cell_row(cell):
    value = _get(cell, "row", "excel_row")
    if value not in (None, ""):
        try:
            return int(value)
        except (TypeError, ValueError):
            pass
    match = re.search(r"\d+", _cell_coordinate(cell))
    return int(match.group(0)) if match else None


def _cell_label(cell):
    return _text(_get(cell, "label", "header", "original_header", "field_name", "semantic_field", default="Án fyrirsagnar"))


def _cell_value(cell):
    return _get(cell, "value", "cached_value", "original_value", "raw_value", "display_value")


def _cell_original_value(cell):
    return _get(cell, "original_value", "raw_value", "value")


def _cell_formula(cell):
    return _text(_get(cell, "formula", "formula_text"))


def _cell_formula_display(cell, reveal_sensitive=False):
    if _cell_sensitive(cell) and not reveal_sensitive and _cell_formula(cell):
        return "••••••••"
    return _cell_formula(cell)


def _cell_hyperlink(cell):
    return _text(_get(cell, "hyperlink", "hyperlink_target", "link", "url"))


def _cell_hyperlink_display(cell, reveal_sensitive=False):
    if _cell_sensitive(cell) and not reveal_sensitive and _cell_hyperlink(cell):
        return "••••••••"
    return _cell_hyperlink(cell)


def _cell_display_value(cell, reveal_sensitive=False):
    display = _get(cell, "display_value")
    if callable(display):
        try:
            return _text(display(reveal_sensitive=reveal_sensitive))
        except TypeError:
            try:
                return _text(display(reveal_sensitive))
            except TypeError:
                pass
    if _cell_sensitive(cell) and not reveal_sensitive:
        return "••••••••"
    return _text(_cell_value(cell))


def _cell_original_display(cell, reveal_sensitive=False):
    if _cell_sensitive(cell) and not reveal_sensitive:
        return "••••••••"
    return _text(_cell_original_value(cell))


def _cell_sensitive(cell):
    explicit = _get(cell, "sensitive", "is_sensitive", "secret")
    if explicit is not None:
        return bool(explicit)
    label = _normalize(_cell_label(cell))
    semantic = _normalize(_get(cell, "semantic_field", "field_name", default=""))
    value = _normalize(_cell_value(cell))
    if value in {"veflykill", "lykilord", "password"}:
        return False
    return any(word in f"{label} {semantic}" for word in ("veflykill", "lykilord", "password", "credential")) and _cell_row(cell) != 3


def _cells_from_record(record):
    cells = _get(record, "cells", "source_cells", "fields", default=[])
    if isinstance(cells, dict):
        sheet = _source_name(record)
        row = _get(record, "row", "excel_row")
        return [
            {
                "sheet": sheet,
                "row": row,
                "coordinate": key,
                "label": key,
                "raw_value": value,
                "sensitive": "veflykill" in _normalize(key) or "credential" in _normalize(key),
            }
            for key, value in cells.items()
        ]
    return _items(cells)


def _record_cells(entity):
    output = []
    for record in _entity_source_records(entity):
        output.extend(_cells_from_record(record))
    direct = _get(entity, "source_cells", "raw_cells", default=[])
    output.extend(_items(direct))
    return output


def _credential_parts(credential):
    if isinstance(credential, str):
        return "Veflykill", credential, ""
    label = _localize_field_label(_get(credential, "label", "field_name", "kind", default="Veflykill"))
    value = _text(_get(credential, "value", "raw_value", "secret", "credential"))
    source = _source_name(credential)
    coordinate = _cell_coordinate(credential)
    if coordinate != "—":
        source = f"{source} · {coordinate}" if source else coordinate
    return label or "Veflykill", value, source


def _resolve_default_workbook():
    configured = _text(registry_workbook_path)
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    candidates = []
    try:
        archive_root = _text(_training_data_dir())
        if archive_root:
            candidates.append(
                str(Path(archive_root).expanduser() / "data folder" / DEFAULT_WORKBOOK_FILENAME)
            )
    except Exception:
        pass
    try:
        candidates.append(resource_path(DEFAULT_WORKBOOK_RELATIVE_PATH))
    except Exception:
        pass
    candidates.append(
        str(Path(__file__).resolve().parent / "trainingData" / "data folder" / DEFAULT_WORKBOOK_FILENAME)
    )
    # Keep old installations working until their workbook is moved beside the
    # imports folder.
    try:
        candidates.append(resource_path(LEGACY_WORKBOOK_RELATIVE_PATH))
    except Exception:
        pass
    candidates.append(
        str(Path(__file__).resolve().parent / "data folder" / DEFAULT_WORKBOOK_FILENAME)
    )
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return os.path.abspath(candidate)
    return os.path.abspath(candidates[0] if candidates else DEFAULT_WORKBOOK_RELATIVE_PATH)


def _find_data_module():
    package = __package__
    names = ("shareholder_data", "shareholder_registry_data", "shareholder_model")
    for name in names:
        try:
            return importlib.import_module(f".{name}", package) if package else importlib.import_module(name)
        except ImportError:
            continue
    return None


def _load_model(path):
    loader = load_registry_data if callable(load_registry_data) else None
    module = None
    if loader is None:
        module = _find_data_module()
        if module is not None:
            for name in ("load_registry_data", "load_registry", "load_workbook", "build_registry"):
                candidate = getattr(module, name, None)
                if callable(candidate):
                    loader = candidate
                    break
    if loader is None:
        raise RuntimeError("Gagnalesari hluthafaskrárinnar er ekki tiltækur.")
    try:
        return loader(path)
    except TypeError as positional_error:
        try:
            return loader(workbook_path=path)
        except TypeError:
            raise positional_error


def _public_search_terms(entity):
    """Return intentionally non-secret fallback search fields only."""
    terms = [_entity_name(entity), _entity_kennitala(entity), _entity_kind(entity)]
    terms.extend(_entity_aliases(entity))
    terms.extend(_entity_statuses(entity))
    for field_name in ("emails", "email", "phones", "phone", "managers", "manager_names", "contacts", "search_terms"):
        values = _get(entity, field_name, default=[])
        if field_name == "search_terms" and not _get(entity, "search_terms_are_public", default=False):
            continue
        for value in _items(values):
            if isinstance(value, (str, int, float)):
                terms.append(_text(value))
            else:
                terms.extend(
                    _text(_get(value, name, default=""))
                    for name in ("name", "display_name", "kennitala", "email", "phone")
                )
    for relationship in _entity_ownerships(entity):
        party = _get(relationship, "party", "shareholder", "owner")
        terms.extend(
            [
                _text(_get(relationship, "party_name", "shareholder_name", "owner_name", "name")),
                _text(_get(relationship, "party_id", "shareholder_id", "owner_id", "kennitala")),
                _entity_name(party) if party is not None else "",
                _entity_kennitala(party) if party is not None else "",
            ]
        )
    return [term for term in terms if term]


def _fallback_search(entities, query):
    query_normalized = _normalize(query)
    query_digits = _digits(query)
    if not query_normalized:
        return [{"entity": entity, "score": 0, "match_reason": ""} for entity in sorted(entities, key=lambda item: _normalize(_entity_name(item)))]
    results = []
    tokens = query_normalized.split()
    for entity in entities:
        name = _normalize(_entity_name(entity))
        kennitala = _digits(_entity_kennitala(entity))
        aliases = [_normalize(alias) for alias in _entity_aliases(entity)]
        public_terms = [_normalize(term) for term in _public_search_terms(entity)]
        score = 0
        reason = ""
        if query_digits and query_digits == kennitala:
            score, reason = 1000, "Exact ID"
        elif query_normalized == name:
            score, reason = 900, "Exact company name"
        elif name.startswith(query_normalized):
            score, reason = 800, "Company name starts with search"
        elif tokens and all(token in name for token in tokens):
            score, reason = 700, "Company name"
        elif query_normalized in name:
            score, reason = 650, "Company name"
        elif any(query_normalized == alias for alias in aliases):
            score, reason = 600, "Alternate name"
        elif any(query_normalized in term for term in public_terms):
            score, reason = 400, "Related public record"
        if score:
            results.append({"entity": entity, "score": score, "match_reason": reason})
    return sorted(results, key=lambda result: (-result["score"], _normalize(_entity_name(result["entity"]))))


def _result_entity(result):
    return _get(result, "entity", "record", "company", default=result)


def _result_reason(result):
    reason = _text(_get(result, "match_reason", "reason", "matched_on"))
    if not reason:
        return ""
    return _MATCH_REASON_LABELS.get(reason, reason.replace("_", " "))


class _ResultCard(tk.Canvas):
    def __init__(self, parent, entity, reason, command, selected=False):
        self.entity = entity
        self.reason = reason
        self.command = command
        self.selected = selected
        self.hovered = False
        super().__init__(
            parent,
            height=68,
            bg=BG_CARD,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
            takefocus=True,
        )
        self.bind("<Configure>", self._draw)
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)
        self.bind("<Button-1>", self._activate)
        self.bind("<Return>", self._activate)
        self.bind("<space>", self._activate)
        self._draw()

    def set_selected(self, selected):
        self.selected = bool(selected)
        self._draw()

    def _draw(self, _event=None):
        self.delete("all")
        width = max(40, self.winfo_width())
        fill = ACCENT_SOFT if self.selected else (BG_INPUT if self.hovered else BG_CARD)
        outline = ACCENT if self.selected else BORDER
        _rounded_rect(self, 2, 2, width - 2, 66, 12, fill=fill, outline=outline, width=2 if self.selected else 1)
        self.create_text(15, 20, text=_entity_name(self.entity), fill=TEXT_PRIMARY, font=("Segoe UI Semibold", 10), anchor="w", width=max(80, width - 30))
        subtitle = _format_kennitala(_entity_kennitala(self.entity))
        if self.reason:
            subtitle = f"{subtitle}  ·  {self.reason}"
        self.create_text(15, 47, text=subtitle, fill=TEXT_SECONDARY, font=("Segoe UI", 8), anchor="w", width=max(80, width - 30))

    def _enter(self, _event):
        self.hovered = True
        self._draw()

    def _leave(self, _event):
        self.hovered = False
        self._draw()

    def _activate(self, _event=None):
        self.focus_set()
        self.command(self.entity)
        return "break"


class _VerticalScrollFrame(tk.Frame):
    def __init__(self, parent, background=None):
        background = background or BG_CARD
        super().__init__(parent, bg=background)
        # A Tk canvas otherwise advertises a roughly 400 px requested width.
        # Keeping its request deliberately tiny lets the surrounding grid, not
        # the canvas default, decide how the results/detail split should resize.
        self.canvas = tk.Canvas(self, width=1, height=1, bg=background, bd=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(
            self,
            orient=tk.VERTICAL,
            command=self.canvas.yview,
            style="Registry.Vertical.TScrollbar",
        )
        self.inner = tk.Frame(self.canvas, bg=background)
        self.window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.inner.bind("<Configure>", self._sync_scrollregion)
        self.canvas.bind("<Configure>", self._sync_width)
        self.canvas.bind("<MouseWheel>", self._wheel)
        self.inner.bind("<MouseWheel>", self._wheel)

    def _sync_scrollregion(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _sync_width(self, event):
        self.canvas.itemconfigure(self.window, width=max(1, event.width))

    def _wheel(self, event):
        delta = int(-event.delta / 120)
        if delta == 0 and event.delta:
            delta = -1 if event.delta > 0 else 1
        if delta:
            self.canvas.yview_scroll(delta, "units")
        return "break"

    def bind_mousewheel_tree(self):
        """Make wheel scrolling work over labels/cards, not only blank canvas."""
        def bind_descendants(widget):
            for child in widget.winfo_children():
                if isinstance(child, (ttk.Treeview, tk.Text, tk.Listbox)):
                    continue
                child.bind("<MouseWheel>", self._wheel, add="+")
                bind_descendants(child)

        bind_descendants(self.inner)

    def reset(self):
        self.canvas.yview_moveto(0)


class _CredentialRow(tk.Frame):
    def __init__(self, parent, label, value, source, on_copy, schedule_remask):
        super().__init__(parent, bg=BG_INPUT, padx=12, pady=10, highlightthickness=1, highlightbackground=BORDER)
        self.secret = value
        self.revealed = False
        self.schedule_remask = schedule_remask
        self.value_var = tk.StringVar(value="••••••••" if value else MISSING_VALUE_LABEL)
        left = tk.Frame(self, bg=BG_INPUT)
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(left, text=label, bg=BG_INPUT, fg=TEXT_PRIMARY, font=("Segoe UI Semibold", 9)).pack(anchor="w")
        tk.Label(left, textvariable=self.value_var, bg=BG_INPUT, fg=TEXT_SECONDARY, font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 0))
        if source:
            tk.Label(left, text=source, bg=BG_INPUT, fg=TEXT_SECONDARY, font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))
        actions = tk.Frame(self, bg=BG_INPUT)
        actions.pack(side=tk.RIGHT, padx=(10, 0))
        self.reveal_button = create_styled_button(actions, "Sýna", self.toggle, width=7, accent=False, height=30)
        self.reveal_button.pack(side=tk.LEFT, padx=(0, 6))
        copy_button = create_styled_button(actions, "Afrita", lambda: on_copy(value), width=7, accent=False, height=30)
        copy_button.pack(side=tk.LEFT)

    def toggle(self):
        if not self.secret:
            return
        self.revealed = not self.revealed
        self.value_var.set(self.secret if self.revealed else "••••••••")
        self.reveal_button.set_text("Fela" if self.revealed else "Sýna")
        if self.revealed:
            self.schedule_remask()

    def remask(self):
        self.revealed = False
        if self.secret:
            self.value_var.set("••••••••")
        self.reveal_button.set_text("Sýna")


class _RegistryPage(tk.Frame):
    def __init__(self, parent, workbook_path=None):
        super().__init__(parent, bg=BG_DARK)
        self.workbook_path = os.path.abspath(workbook_path or _resolve_default_workbook())
        self.model = None
        self.entities = []
        self.results = []
        self.selected_entity = None
        self.selected_key = None
        self.current_tab = "overview"
        self.alive = True
        self.load_queue = queue.Queue()
        self.load_token = 0
        self.poll_after_id = None
        self.search_after_id = None
        self.remask_after_id = None
        self.credential_rows = []
        self.raw_dialogs = []
        self.root_bindings = []
        self.result_cards = []
        self._previous_window_layout = _capture_window_layout(root)
        self._resize_for_registry()
        self.pack(fill=tk.BOTH, expand=True)
        self._build_header()
        self.content = tk.Frame(self, bg=BG_DARK)
        self.content.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        self._bind_shortcuts()
        if os.path.isfile(self.workbook_path):
            self.start_load(self.workbook_path)
        else:
            self._show_missing_file()

    def _resize_for_registry(self):
        if root is None:
            return
        root.update_idletasks()
        work_left, work_top, work_right, work_bottom = window_work_area(root)
        work_width = max(320, work_right - work_left)
        work_height = max(320, work_bottom - work_top)
        current_state = root.state()
        current_x = root.winfo_x()
        current_y = root.winfo_y()
        current_width = max(root.winfo_width(), 1)
        current_height = max(root.winfo_height(), 1)
        available_width = max(320, work_width - 32)
        available_height = max(320, work_height - 32)
        width = min(available_width, max(1120, current_width))
        height = min(available_height, max(790, current_height))
        center_x = current_x + current_width // 2
        center_y = current_y + current_height // 2
        x = max(work_left + 16, min(work_right - width - 16, center_x - width // 2))
        y = max(work_top + 8, min(work_bottom - height - 16, center_y - height // 2))
        root.minsize(min(880, available_width), min(600, available_height))
        # Respect a user's maximized workspace.  Geometry changes while zoomed
        # can unexpectedly change the restore size on Windows.
        if current_state != "zoomed":
            root.geometry(_geometry_spec(width, height, x, y))

    def _build_header(self):
        header = tk.Frame(self, bg=BG_DARK)
        header.pack(fill=tk.X)
        back_button = create_styled_button(header, "← Til baka", self.go_back, width=10, accent=False, height=38)
        back_button.pack(side=tk.LEFT)

        source_column = tk.Frame(header, bg=BG_DARK)
        source_column.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(14, 12))
        self.source_label = tk.Label(
            source_column,
            text=os.path.basename(self.workbook_path),
            bg=BG_DARK,
            fg=TEXT_PRIMARY,
            font=("Segoe UI Semibold", 10),
            anchor="w",
        )
        self.source_label.pack(fill=tk.X)
        self.status_label = tk.Label(
            source_column,
            text="",
            bg=BG_DARK,
            fg=TEXT_SECONDARY,
            font=("Segoe UI", 8),
            anchor="w",
        )
        self.status_label.pack(fill=tk.X, pady=(2, 0))

        self.raw_button = create_styled_button(header, "Frumgögn", self.open_raw_explorer, width=12, accent=False, height=38)
        self.raw_button.pack(side=tk.RIGHT, padx=(7, 0))
        reload_button = create_styled_button(header, "Endurhlaða", self.reload, width=10, accent=False, height=38)
        reload_button.pack(side=tk.RIGHT, padx=(7, 0))
        browse_button = create_styled_button(header, "Velja skrá", self.browse, width=11, accent=False, height=38)
        browse_button.pack(side=tk.RIGHT)

    def _bind_shortcuts(self):
        if root is None:
            return
        for sequence, callback in (
            ("<Control-f>", self._focus_search),
            ("<Escape>", self._escape),
        ):
            binding_id = root.bind(sequence, callback, add="+")
            self.root_bindings.append((sequence, binding_id))

    def _focus_search(self, _event=None):
        if hasattr(self, "search_entry") and self.search_entry.winfo_exists():
            self.search_entry.focus_set()
            self.search_entry.selection_range(0, tk.END)
            return "break"
        return None

    def _escape(self, _event=None):
        if any(dialog.winfo_exists() for dialog in self.raw_dialogs):
            dialog = next(dialog for dialog in reversed(self.raw_dialogs) if dialog.winfo_exists())
            dialog.destroy()
            return "break"
        self.go_back()
        return "break"

    def _clear_content(self):
        for widget in self.content.winfo_children():
            widget.destroy()

    def _show_center_state(self, title, detail, color=None, actions=()):
        self._clear_content()
        shell = tk.Frame(self.content, bg=BG_DARK)
        shell.place(relx=0.5, rely=0.44, anchor="center", relwidth=0.74)
        panel, inner = create_panel(shell, padx=30, pady=28)
        panel.pack(fill=tk.X)
        tk.Label(inner, text=title, bg=BG_CARD, fg=color or TEXT_PRIMARY, font=("Segoe UI Semibold", 16)).pack()
        detail_label = tk.Label(
            inner,
            text=detail,
            bg=BG_CARD,
            fg=TEXT_SECONDARY,
            font=("Segoe UI", 9),
            justify="center",
            wraplength=620,
        )
        detail_label.pack(fill=tk.X, pady=(8, 18))
        _bind_responsive_wrap(detail_label, inner, horizontal_padding=8, minimum=180)
        if actions:
            row = tk.Frame(inner, bg=BG_CARD)
            row.pack()
            for index, (label, command, accent) in enumerate(actions):
                button = create_styled_button(row, label, command, width=14, accent=accent, height=38)
                button.pack(side=tk.LEFT, padx=(0 if index == 0 else 6, 6 if index < len(actions) - 1 else 0))

    def _show_missing_file(self):
        self.status_label.configure(text="Vinnubók fannst ekki", fg=TEXT_WARNING)
        self._show_center_state(
            "Veldu vinnubók hluthafaskrár",
            f"Væntanleg skrá fannst ekki:\n{self.workbook_path}\n\nVeldu Excel-vinnubókina og reyndu aftur.",
            color=TEXT_WARNING,
            actions=(("Velja skrá", self.browse, True), ("Reyna aftur", self.reload, False), ("Til baka", self.go_back, False)),
        )

    def browse(self):
        initial = os.path.dirname(self.workbook_path)
        if not os.path.isdir(initial):
            initial = os.path.dirname(__file__)
        chosen = filedialog.askopenfilename(
            parent=root,
            title="Veldu vinnubók hluthafaskrár",
            initialdir=initial,
            filetypes=(("Excel-vinnubækur", "*.xlsx"), ("Allar skrár", "*.*")),
        )
        if chosen:
            self.start_load(chosen)

    def reload(self):
        if os.path.isfile(self.workbook_path):
            self.start_load(self.workbook_path)
        else:
            self._show_missing_file()

    def start_load(self, path):
        path = os.path.abspath(path)
        self.workbook_path = path
        self.source_label.configure(text=os.path.basename(path))
        self.status_label.configure(text="Hleð inn þremur völdum upprunablöðum…", fg=TEXT_SECONDARY)
        self._show_center_state(
            "Hleð hluthafaskrá",
            "Les og tengi saman færslur um fyrirtæki, hluthafa, samskiptaupplýsingar og reiti úr frumgögnum. Viðmótið helst virkt á meðan.",
        )
        self.load_token += 1
        token = self.load_token

        def worker():
            try:
                model = _load_model(path)
                self.load_queue.put((token, model, None))
            except Exception as error:  # Delivered to the UI without logging workbook content.
                self.load_queue.put((token, None, error))

        threading.Thread(target=worker, daemon=True, name="shareholder-registry-loader").start()
        if self.poll_after_id is None and root is not None:
            self.poll_after_id = root.after(60, self._poll_load)

    def _poll_load(self):
        self.poll_after_id = None
        if not self.alive or not self.winfo_exists():
            return
        current_completion = None
        try:
            while True:
                completion = self.load_queue.get_nowait()
                if completion[0] == self.load_token:
                    current_completion = completion
        except queue.Empty:
            pass
        if current_completion is None:
            self.poll_after_id = root.after(60, self._poll_load)
            return
        _token, model, error = current_completion
        if error is not None:
            self._load_failed(error)
            return
        self.model = _get(model, "registry", "data", default=model)
        self.entities = _model_entities(self.model)
        self._load_succeeded()

    def _load_failed(self, error):
        if isinstance(error, FileNotFoundError):
            self._show_missing_file()
            return
        self.status_label.configure(text="Ekki tókst að hlaða vinnubók", fg=TEXT_ERROR)
        detail = _text(error) or error.__class__.__name__
        missing_prefix = "The selected workbook is missing required registry sheet(s): "
        if detail.startswith(missing_prefix):
            detail = "Valda vinnubók vantar nauðsynleg blöð hluthafaskrárinnar: " + detail[len(missing_prefix):]
        if len(detail) > 500:
            detail = detail[:497] + "…"
        self._show_center_state(
            "Ekki tókst að hlaða vinnubókinni",
            detail,
            color=TEXT_ERROR,
            actions=(("Reyna aftur", self.reload, True), ("Velja aðra skrá", self.browse, False), ("Til baka", self.go_back, False)),
        )

    def _load_succeeded(self):
        sheet_names = _items(_get(self.model, "sheet_names", "selected_sheet_names", default=REGISTRY_SHEET_NAMES))
        source_count = len(sheet_names) or len(REGISTRY_SHEET_NAMES)
        self.status_label.configure(
            text=f"{_format_count(len(self.entities))} tengdar færslur · {_source_sheet_count(source_count)}",
            fg=TEXT_SUCCESS,
        )
        self._build_workspace()

    def _build_workspace(self):
        self._clear_content()
        self._configure_ttk_styles()
        search_panel, search_inner = create_panel(self.content, padx=14, pady=12)
        search_panel.pack(fill=tk.X, pady=(0, 10))
        search_shell = tk.Frame(search_inner, bg=BORDER, padx=1, pady=1)
        search_shell.pack(side=tk.LEFT, fill=tk.X, expand=True)
        search_surface = tk.Frame(search_shell, bg=BG_INPUT)
        search_surface.pack(fill=tk.BOTH, expand=True)
        self.search_var = tk.StringVar()
        self.search_entry = create_styled_entry(search_surface, width=48)
        self.search_entry.configure(textvariable=self.search_var, font=("Segoe UI", 11), bg=BG_INPUT)
        self.search_entry.pack(fill=tk.X, ipady=9, padx=12)
        self.search_hint = tk.Label(
            search_surface,
            text="Leita eftir nafni, kennitölu, hluthafa, netfangi eða síma",
            bg=BG_INPUT,
            fg=TEXT_SECONDARY,
            font=("Segoe UI", 10),
            anchor="w",
            cursor="xterm",
        )
        self.search_hint.place(x=12, rely=0.5, anchor="w")
        self.search_hint.bind("<Button-1>", lambda _event: self.search_entry.focus_set())
        self.search_entry.bind("<FocusIn>", lambda _event: search_shell.configure(bg=ACCENT), add="+")
        self.search_entry.bind("<FocusOut>", lambda _event: search_shell.configure(bg=BORDER), add="+")
        self.search_var.trace_add("write", lambda *_args: self._sync_search_hint())
        self.search_entry.bind("<FocusIn>", lambda _event: self._sync_search_hint(), add="+")
        self.search_entry.bind("<FocusOut>", lambda _event: self._sync_search_hint(), add="+")
        self.search_entry.bind("<KeyRelease>", self._schedule_search)
        self.search_entry.bind("<Down>", lambda _event: self._move_result_selection(1))
        self.search_entry.bind("<Up>", lambda _event: self._move_result_selection(-1))
        self.search_entry.bind("<Return>", self._select_active_result)

        self.type_filter = tk.StringVar(value=ALL_RECORDS_LABEL)
        filter_values = (ALL_RECORDS_LABEL, "Hlutafélög", "SLF félög", "Einstaklingar")
        filter_box = ttk.Combobox(
            search_inner,
            textvariable=self.type_filter,
            values=filter_values,
            state="readonly",
            style="Registry.TCombobox",
            width=14,
        )
        filter_box.pack(side=tk.LEFT, padx=(10, 0), ipady=5)
        filter_box.bind("<<ComboboxSelected>>", lambda _event: self._apply_search())
        clear_button = create_styled_button(search_inner, "Hreinsa", self._clear_search, width=8, accent=False, height=38)
        clear_button.pack(side=tk.LEFT, padx=(8, 0))

        body = tk.Frame(self.content, bg=BG_DARK)
        body.pack(fill=tk.BOTH, expand=True)
        body.columnconfigure(0, minsize=300, weight=0)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        body.bind("<Configure>", self._sync_body_columns)

        left = tk.Frame(body, bg=BG_CARD, highlightbackground=BORDER, highlightthickness=1)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        list_header = tk.Frame(left, bg=BG_CARD)
        list_header.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 7))
        self.result_count_label = tk.Label(list_header, text="Niðurstöður", bg=BG_CARD, fg=TEXT_PRIMARY, font=("Segoe UI Semibold", 10))
        self.result_count_label.pack(side=tk.LEFT)
        self.match_hint_label = tk.Label(list_header, text="", bg=BG_CARD, fg=TEXT_SECONDARY, font=("Segoe UI", 8))
        self.match_hint_label.pack(side=tk.RIGHT)
        self.result_scroll = _VerticalScrollFrame(left, background=BG_CARD)
        self.result_scroll.grid(row=1, column=0, sticky="nsew", padx=7, pady=(0, 7))

        self.detail_panel = tk.Frame(body, bg=BG_CARD, highlightbackground=BORDER, highlightthickness=1)
        self.detail_panel.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self.detail_panel.columnconfigure(0, weight=1)
        self.detail_panel.rowconfigure(2, weight=1)

        self._apply_search()

    def _sync_body_columns(self, event):
        """Give detail records priority while keeping search results usable."""
        width = max(1, int(event.width))
        if width < 860:
            result_width = 270
        else:
            result_width = min(320, max(292, int(width * 0.29)))
        if getattr(self, "_result_column_width", None) == result_width:
            return
        self._result_column_width = result_width
        event.widget.columnconfigure(0, minsize=result_width)

    def _sync_search_hint(self):
        if not hasattr(self, "search_hint") or not self.search_hint.winfo_exists():
            return
        if self.search_var.get() or self.focus_get() == self.search_entry:
            self.search_hint.place_forget()
        else:
            self.search_hint.place(x=12, rely=0.5, anchor="w")

    def _configure_ttk_styles(self):
        style = ttk.Style(root)
        # Bypass the theme-native Treeview field frame; on Windows/clam it can
        # render as a bright white focus rectangle even in a dark theme.
        style.layout("Registry.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        style.configure(
            "Registry.Treeview",
            background=BG_CARD,
            fieldbackground=BG_CARD,
            foreground=TEXT_PRIMARY,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            relief="flat",
            borderwidth=0,
            rowheight=32,
            font=("Segoe UI", 9),
        )
        style.map(
            "Registry.Treeview",
            background=[("selected", ACCENT_SOFT)],
            foreground=[("selected", TEXT_PRIMARY)],
        )
        style.configure(
            "Registry.Treeview.Heading",
            background=BUTTON_MUTED,
            foreground=TEXT_PRIMARY,
            relief="flat",
            borderwidth=0,
            padding=(10, 8),
            font=("Segoe UI Semibold", 9),
        )
        style.map("Registry.Treeview.Heading", background=[("active", BUTTON_MUTED_HOVER)])
        style.configure(
            "Registry.TCombobox",
            fieldbackground=BG_INPUT,
            background=BUTTON_MUTED,
            foreground=TEXT_PRIMARY,
            selectbackground=BG_INPUT,
            selectforeground=TEXT_PRIMARY,
            arrowcolor=TEXT_SECONDARY,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            relief="flat",
            borderwidth=1,
            padding=(9, 6),
            arrowsize=12,
        )
        style.map(
            "Registry.TCombobox",
            fieldbackground=[("readonly", BG_INPUT), ("focus", BG_INPUT)],
            foreground=[("readonly", TEXT_PRIMARY), ("focus", TEXT_PRIMARY)],
            selectbackground=[("readonly", BG_INPUT)],
            selectforeground=[("readonly", TEXT_PRIMARY)],
            background=[("readonly", BUTTON_MUTED), ("active", BUTTON_MUTED_HOVER)],
            bordercolor=[("focus", ACCENT), ("readonly", BORDER)],
            lightcolor=[("focus", ACCENT), ("readonly", BORDER)],
            darkcolor=[("focus", ACCENT), ("readonly", BORDER)],
            arrowcolor=[("active", TEXT_PRIMARY), ("readonly", TEXT_SECONDARY)],
        )

        # Arrowless, narrow scrollbars avoid the bright native Windows chrome
        # while keeping the full scrollbar keyboard/mouse behaviour from ttk.
        style.layout(
            "Registry.Vertical.TScrollbar",
            [("Vertical.Scrollbar.trough", {"sticky": "ns", "children": [
                ("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"}),
            ]})],
        )
        style.layout(
            "Registry.Horizontal.TScrollbar",
            [("Horizontal.Scrollbar.trough", {"sticky": "ew", "children": [
                ("Horizontal.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"}),
            ]})],
        )
        for scrollbar_style in ("Registry.Vertical.TScrollbar", "Registry.Horizontal.TScrollbar"):
            style.configure(
                scrollbar_style,
                background=BUTTON_MUTED,
                troughcolor=BG_CARD,
                bordercolor=BG_CARD,
                lightcolor=BUTTON_MUTED,
                darkcolor=BUTTON_MUTED,
                relief="flat",
                borderwidth=0,
                width=10,
            )
            style.map(
                scrollbar_style,
                background=[("pressed", ACCENT), ("active", BUTTON_MUTED_HOVER)],
                lightcolor=[("pressed", ACCENT), ("active", BUTTON_MUTED_HOVER)],
                darkcolor=[("pressed", ACCENT), ("active", BUTTON_MUTED_HOVER)],
            )

        if root is not None:
            dropdown_font = ("Segoe UI", 9)
            root.option_add("*TCombobox*Listbox.background", BG_CARD)
            root.option_add("*TCombobox*Listbox.foreground", TEXT_PRIMARY)
            root.option_add("*TCombobox*Listbox.selectBackground", ACCENT_SOFT)
            root.option_add("*TCombobox*Listbox.selectForeground", TEXT_PRIMARY)
            root.option_add("*TCombobox*Listbox.font", dropdown_font)

    def _schedule_search(self, _event=None):
        if self.search_after_id is not None:
            root.after_cancel(self.search_after_id)
        self.search_after_id = root.after(SEARCH_DEBOUNCE_MS, self._apply_search)

    def _clear_search(self):
        self.search_var.set("")
        self._apply_search()
        self.search_entry.focus_set()

    def _search_model(self, query):
        if not query.strip():
            return _fallback_search(self.entities, "")
        search = _get(self.model, "search", "search_entities")
        if callable(search):
            attempts = (
                lambda: search(query, limit=500),
                lambda: search(query=query, limit=500),
                lambda: search(query),
            )
            for attempt in attempts:
                try:
                    return _items(attempt())
                except TypeError:
                    continue
        return _fallback_search(self.entities, query)

    def _matches_type_filter(self, entity):
        selected = self.type_filter.get() if hasattr(self, "type_filter") else ALL_RECORDS_LABEL
        kind = _entity_kind_key(entity)
        if selected == "Hlutafélög":
            return not any(word in kind for word in ("person", "individual", "einstaklingur", "partnership", "sameignarfelag", "slf"))
        if selected == "SLF félög":
            return any(word in kind for word in ("partnership", "sameignarfelag", "slf"))
        if selected == "Einstaklingar":
            return any(word in kind for word in ("person", "individual", "einstaklingur"))
        return True

    def _apply_search(self):
        self.search_after_id = None
        if not self.model or not hasattr(self, "search_var"):
            return
        query = self.search_var.get().strip()
        self.results = [result for result in self._search_model(query) if self._matches_type_filter(_result_entity(result))]
        self._render_results(query)

    def _render_results(self, query):
        for child in self.result_scroll.inner.winfo_children():
            child.destroy()
        self.result_cards = []
        self.result_count_label.configure(text=f"Niðurstöður ({_format_count(len(self.results))})")
        self.match_hint_label.configure(text="eftir samsvörun" if query else "A–Ö")
        if not self.results:
            tk.Label(
                self.result_scroll.inner,
                text=f'Engar færslur passa við „{query}“.' if query else "Engar færslur fundust í hluthafaskránni.",
                bg=BG_CARD,
                fg=TEXT_SECONDARY,
                font=("Segoe UI", 9),
                wraplength=280,
                justify="center",
            ).pack(fill=tk.X, padx=18, pady=36)
            self._show_empty_detail("Engin niðurstaða valin", "Prófaðu nafn fyrirtækis, kennitölu, hluthafa, netfang eða símanúmer.")
            return
        available_keys = {_entity_key(_result_entity(result)) for result in self.results}
        if self.selected_key not in available_keys:
            self.selected_entity = _result_entity(self.results[0])
            self.selected_key = _entity_key(self.selected_entity)
        for result in self.results[:500]:
            entity = _result_entity(result)
            card = _ResultCard(
                self.result_scroll.inner,
                entity,
                _result_reason(result),
                self.select_entity,
                selected=_entity_key(entity) == self.selected_key,
            )
            card.pack(fill=tk.X, padx=5, pady=3)
            card.bind("<MouseWheel>", self.result_scroll._wheel, add="+")
            card.bind("<Down>", lambda _event: self._move_result_selection(1), add="+")
            card.bind("<Up>", lambda _event: self._move_result_selection(-1), add="+")
            self.result_cards.append(card)
        self.result_scroll.reset()
        self._render_detail()

    def _move_result_selection(self, delta):
        if not self.results:
            return "break"
        keys = [_entity_key(_result_entity(result)) for result in self.results]
        try:
            index = keys.index(self.selected_key)
        except ValueError:
            index = 0
        index = max(0, min(len(self.results) - 1, index + delta))
        self.select_entity(_result_entity(self.results[index]))
        if index < len(self.result_cards):
            card = self.result_cards[index]
            card.focus_set()
            self._ensure_result_visible(card)
        return "break"

    def _ensure_result_visible(self, card):
        self.result_scroll.inner.update_idletasks()
        content_height = max(1, self.result_scroll.inner.winfo_reqheight())
        viewport_height = max(1, self.result_scroll.canvas.winfo_height())
        top = self.result_scroll.canvas.canvasy(0)
        bottom = top + viewport_height
        card_top = card.winfo_y()
        card_bottom = card_top + card.winfo_height()
        if card_top < top:
            self.result_scroll.canvas.yview_moveto(max(0.0, card_top / content_height))
        elif card_bottom > bottom:
            target = max(0.0, (card_bottom - viewport_height) / content_height)
            self.result_scroll.canvas.yview_moveto(min(1.0, target))

    def _select_active_result(self, _event=None):
        if self.results:
            self.select_entity(_result_entity(self.results[0] if self.selected_key is None else next((item for item in self.results if _entity_key(_result_entity(item)) == self.selected_key), self.results[0])))
        return "break"

    def select_entity(self, entity):
        self._remask_credentials()
        self.selected_entity = entity
        self.selected_key = _entity_key(entity)
        for card in self.result_cards:
            card.set_selected(_entity_key(card.entity) == self.selected_key)
        self._render_detail()

    def _show_empty_detail(self, title, detail):
        for child in self.detail_panel.winfo_children():
            child.destroy()
        holder = tk.Frame(self.detail_panel, bg=BG_CARD)
        holder.place(relx=0.5, rely=0.44, relwidth=0.82, anchor="center")
        tk.Label(holder, text=title, bg=BG_CARD, fg=TEXT_PRIMARY, font=("Segoe UI Semibold", 14)).pack()
        detail_label = tk.Label(
            holder,
            text=detail,
            bg=BG_CARD,
            fg=TEXT_SECONDARY,
            font=("Segoe UI", 9),
            wraplength=440,
            justify="center",
        )
        detail_label.pack(fill=tk.X, pady=(6, 0))
        _bind_responsive_wrap(detail_label, holder, horizontal_padding=8)

    def _render_detail(self):
        if not self.selected_entity:
            self._show_empty_detail("Veldu færslu", "Veldu niðurstöðu til að sjá tengdar upplýsingar úr vinnubókinni.")
            return
        self._remask_credentials()
        self.revealed_source_cells = set()
        for child in self.detail_panel.winfo_children():
            child.destroy()
        entity = self.selected_entity

        header = tk.Frame(self.detail_panel, bg=BG_CARD)
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(15, 8))
        header.columnconfigure(0, weight=1)
        name = tk.Label(
            header,
            text=_entity_name(entity),
            bg=BG_CARD,
            fg=TEXT_PRIMARY,
            font=("Segoe UI Semibold", 16),
            anchor="w",
        )
        name.grid(row=0, column=0, sticky="ew")
        subtitle_parts = [_format_kennitala(_entity_kennitala(entity)), _entity_kind(entity)]
        source_names = list(dict.fromkeys(_source_name(record) for record in _entity_source_records(entity) if _source_name(record)))
        if source_names:
            subtitle_parts.append(_source_sheet_count(len(source_names)))
        tk.Label(
            header,
            text="  ·  ".join(part for part in subtitle_parts if part),
            bg=BG_CARD,
            fg=TEXT_SECONDARY,
            font=("Segoe UI", 9),
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", pady=(4, 0))
        if _entity_kennitala(entity):
            copy_id = create_styled_button(
                header,
                "Afrita kennitölu",
                lambda: self._copy_to_clipboard(_entity_kennitala(entity), "Kennitala afrituð"),
                width=9,
                accent=False,
                height=32,
            )
            copy_id.grid(row=0, column=1, rowspan=2, padx=(12, 0))

        tabs = tk.Frame(self.detail_panel, bg=BG_CARD)
        tabs.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 8))
        self.tab_buttons = {}
        for index, (key, label) in enumerate(
            (("overview", "Yfirlit"), ("ownership", "Eignarhald"), ("sources", "Upprunafærslur"))
        ):
            button = create_styled_button(
                tabs,
                label,
                lambda selected=key: self._select_detail_tab(selected),
                width=13,
                accent=key == self.current_tab,
                height=34,
            )
            button.pack(side=tk.LEFT, padx=(0, 7) if index < 2 else (0, 0))
            self.tab_buttons[key] = button

        self.detail_body = tk.Frame(self.detail_panel, bg=BG_CARD)
        self.detail_body.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.detail_body.columnconfigure(0, weight=1)
        self.detail_body.rowconfigure(0, weight=1)
        self._render_current_tab()

    def _select_detail_tab(self, key):
        if key == self.current_tab:
            return
        self._remask_credentials()
        self.current_tab = key
        for tab_key, button in self.tab_buttons.items():
            set_button_accent(button, tab_key == key)
        self._render_current_tab()

    def _render_current_tab(self):
        for child in self.detail_body.winfo_children():
            child.destroy()
        if self.current_tab == "ownership":
            self._render_ownership_tab()
        elif self.current_tab == "sources":
            self._render_sources_tab()
        else:
            self._render_overview_tab()

    def _section(self, parent, title, subtitle=""):
        shell = tk.Frame(parent, bg=BG_INPUT, highlightbackground=BORDER, highlightthickness=1, padx=14, pady=12)
        shell.pack(fill=tk.X, padx=5, pady=(5, 8))
        tk.Label(shell, text=title, bg=BG_INPUT, fg=TEXT_PRIMARY, font=("Segoe UI Semibold", 11), anchor="w").pack(fill=tk.X)
        if subtitle:
            subtitle_label = tk.Label(
                shell,
                text=subtitle,
                bg=BG_INPUT,
                fg=TEXT_SECONDARY,
                font=("Segoe UI", 8),
                anchor="w",
                justify="left",
                wraplength=620,
            )
            subtitle_label.pack(fill=tk.X, pady=(2, 8))
            _bind_responsive_wrap(subtitle_label, shell)
        content = tk.Frame(shell, bg=BG_INPUT)
        content.pack(fill=tk.X)
        return content

    def _add_field(self, parent, label, value, source="", row=None):
        row = row if row is not None else len(parent.grid_slaves()) // 3
        tk.Label(parent, text=label, bg=BG_INPUT, fg=TEXT_SECONDARY, font=("Segoe UI", 8), anchor="w").grid(
            row=row, column=0, sticky="nw", padx=(0, 14), pady=5
        )
        value_text = _text(value) or MISSING_VALUE_LABEL
        value_label = tk.Label(
            parent,
            text=value_text,
            bg=BG_INPUT,
            fg=TEXT_PRIMARY if value_text != MISSING_VALUE_LABEL else TEXT_SECONDARY,
            font=("Segoe UI Semibold", 9),
            anchor="w",
            justify="left",
            wraplength=420,
        )
        value_label.grid(row=row, column=1, sticky="ew", pady=5)
        _bind_responsive_wrap(value_label, parent, horizontal_padding=250, minimum=120)
        if source:
            tk.Label(parent, text=source, bg=BG_INPUT, fg=TEXT_SECONDARY, font=("Segoe UI", 8), anchor="e").grid(
                row=row, column=2, sticky="ne", padx=(10, 0), pady=5
            )

    def _observed_values(self, entity, header_terms):
        values = []
        seen = set()
        normalized_terms = tuple(_normalize(term) for term in header_terms)
        for cell in _record_cells(entity):
            if _cell_sensitive(cell):
                continue
            label = _normalize(_cell_label(cell))
            if not label or not any(term and term in label for term in normalized_terms):
                continue
            value = _cell_display_value(cell)
            if not value:
                continue
            key = (value, _source_name(cell), _cell_coordinate(cell))
            if key in seen:
                continue
            seen.add(key)
            values.append((value, f"{_source_name(cell)} · {_cell_coordinate(cell)}"))
        return values

    def _render_overview_tab(self):
        scroll = _VerticalScrollFrame(self.detail_body, background=BG_CARD)
        scroll.grid(row=0, column=0, sticky="nsew")
        entity = self.selected_entity

        identity = self._section(scroll.inner, "Grunnupplýsingar")
        identity.columnconfigure(1, weight=1)
        fields = [
            ("Nafn", _entity_name(entity), ""),
            ("Kennitala", _format_kennitala(_entity_kennitala(entity)), ""),
            ("Tegund færslu", _entity_kind(entity), ""),
        ]
        aliases = _entity_aliases(entity)
        if aliases:
            fields.append(("Önnur nöfn", ", ".join(aliases), ""))
        for index, field in enumerate(fields):
            self._add_field(identity, *field, row=index)

        statuses = _entity_statuses(entity)
        if statuses:
            category = self._section(scroll.inner, "Flokkun")
            pill_row = tk.Frame(category, bg=BG_INPUT)
            pill_row.pack(fill=tk.X, pady=(8, 0))
            for status in statuses:
                status_label = tk.Label(
                    pill_row,
                    text=_localize_section(status),
                    bg=ACCENT_SOFT,
                    fg=TEXT_PRIMARY,
                    font=("Segoe UI Semibold", 8),
                    padx=9,
                    pady=5,
                    anchor="w",
                    justify="left",
                    wraplength=560,
                )
                status_label.pack(fill=tk.X, pady=(0, 6))
                _bind_responsive_wrap(status_label, pill_row, horizontal_padding=22)

        financial = []
        capital_direct = _get(entity, "capital", "share_capital", "current_capital")
        if capital_direct not in (None, ""):
            financial.append(("Hlutafé", _format_amount(capital_direct), "Tengd færsla"))
        for value, source in self._observed_values(entity, ("Hlutafé", "Stofnfé")):
            financial.append(("Hlutafé", _format_amount(value), source))
        for value, source in self._observed_values(entity, ("Arður",)):
            financial.append(("Arður", _format_amount(value), source))
        for value, source in self._observed_values(entity, ("Fjármtsk", "Fjármagnstekjuskattur")):
            financial.append(("Fjármtsk", _format_amount(value), source))
        deduped_financial = []
        seen_financial = set()
        for field in financial:
            marker = tuple(field)
            if marker not in seen_financial:
                seen_financial.add(marker)
                deduped_financial.append(field)
        if deduped_financial:
            finance_section = self._section(
                scroll.inner,
                "Hlutafé, arður og fjármagnstekjuskattur",
                "Gildi eru sýnd eftir upprunablaði þegar færslur í vinnubókinni stangast á.",
            )
            finance_section.columnconfigure(1, weight=1)
            for index, field in enumerate(deduped_financial):
                self._add_field(finance_section, *field, row=index)

        emails = [_text(value) for value in _items(_get(entity, "emails", "email", default=[])) if _text(value)]
        phones = [_text(value) for value in _items(_get(entity, "phones", "phone", default=[])) if _text(value)]
        managers = self._observed_values(entity, ("Framkv", "Manager"))
        if emails or phones or managers:
            contact = self._section(scroll.inner, "Samskiptaupplýsingar")
            contact.columnconfigure(1, weight=1)
            index = 0
            for email in emails:
                self._add_field(contact, "Netfang", email, "", row=index)
                index += 1
            for phone in phones:
                self._add_field(contact, "Sími", phone, "", row=index)
                index += 1
            for value, source in managers:
                self._add_field(contact, "Framkv.stj", _format_kennitala(value), source, row=index)
                index += 1

        credentials = [credential for credential in _entity_credentials(entity) if _credential_parts(credential)[1]]
        if credentials:
            credential_section = self._section(
                scroll.inner,
                "Veflyklar",
                "Viðkvæm gildi eru undanskilin leit og hulin þar til þú velur sérstaklega að sýna eða afrita þau.",
            )
            for credential in credentials:
                label, value, source = _credential_parts(credential)
                row = _CredentialRow(
                    credential_section,
                    label,
                    value,
                    source,
                    self._copy_secret,
                    self._schedule_remask,
                )
                row.pack(fill=tk.X, pady=(8, 0))
                self.credential_rows.append(row)

        warnings = [_text(value) for value in _items(_get(entity, "warnings", "quality_flags", default=[])) if _text(value)]
        for record in _entity_source_records(entity):
            warning = _text(_get(record, "layout_warning", "warning"))
            if warning:
                warnings.append(warning)
        warnings = list(dict.fromkeys(warnings))
        if warnings:
            quality = self._section(
                scroll.inner,
                "Athugasemdir við gögn",
                "Upprunaleg gildi vinnubókarinnar eru aðgengileg á flipanum „Upprunafærslur“.",
            )
            for warning in warnings:
                warning_label = tk.Label(
                    quality,
                    text=f"• {_localize_warning(warning)}",
                    bg=BG_INPUT,
                    fg=TEXT_WARNING,
                    font=("Segoe UI", 9),
                    anchor="w",
                    justify="left",
                    wraplength=620,
                )
                warning_label.pack(fill=tk.X, pady=2)
                _bind_responsive_wrap(warning_label, quality, horizontal_padding=4)

        records = _entity_source_records(entity)
        coverage = self._section(scroll.inner, "Gagnaumfang")
        coverage.columnconfigure(1, weight=1)
        sheets = list(dict.fromkeys(_source_name(record) for record in records if _source_name(record)))
        cell_count = sum(len(_cells_from_record(record)) for record in records)
        self._add_field(coverage, "Upprunablöð", ", ".join(sheets) or "Ekki tengt", "", row=0)
        self._add_field(coverage, "Línur í frumgögnum", len(records), "", row=1)
        self._add_field(coverage, "Reitir í frumgögnum", cell_count, "", row=2)
        scroll.bind_mousewheel_tree()

    def _relationship_values(self, relationship):
        party = _get(relationship, "party", "shareholder", "owner")
        role = _localize_relationship_role(_get(relationship, "role", "relationship_type", default="Hluthafi"))
        name = _text(_get(relationship, "party_name", "shareholder_name", "owner_name", "name"))
        identifier = _text(_get(relationship, "party_id", "shareholder_id", "owner_id", "kennitala"))
        if party is not None:
            name = name or _entity_name(party)
            identifier = identifier or _entity_kennitala(party)
        share = _get(relationship, "share", "amount", "holding")
        percentage = _get(relationship, "percentage", "ownership_percentage", "ratio")
        source = _text(_get(relationship, "source_sheet", "sheet", "source"))
        row = _get(relationship, "excel_row", "row")
        if row not in (None, ""):
            source = f"{source} · lína {row}" if source else f"Lína {row}"
        warning = _text(_get(relationship, "warning", "quality_warning"))
        return (
            role,
            name or "Ógreindur aðili",
            _format_kennitala(identifier) if identifier else "—",
            _format_amount(share) if share not in (None, "") else MISSING_VALUE_LABEL,
            _format_percent(percentage),
            source,
            _localize_warning(warning),
        )

    def _relationship_table(self, parent, title, relationships):
        section = self._section(parent, title)
        table_shell = tk.Frame(section, bg=BG_INPUT)
        table_shell.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        table_shell.columnconfigure(0, weight=1)
        table_shell.rowconfigure(0, weight=1)
        columns = ("role", "party", "id", "share", "percent", "source")
        tree = ttk.Treeview(table_shell, columns=columns, show="headings", style="Registry.Treeview", height=min(12, max(3, len(relationships))))
        headings = (("role", "Hlutverk", 92), ("party", "Aðili", 165), ("id", "Kennitala", 104), ("share", "Hlutur", 105), ("percent", "%", 52), ("source", "Uppruni", 150))
        for column, label, width in headings:
            tree.heading(column, text=label)
            tree.column(column, width=width, minwidth=45, stretch=column in {"party", "source"})
        scrollbar = ttk.Scrollbar(
            table_shell,
            orient=tk.VERTICAL,
            command=tree.yview,
            style="Registry.Vertical.TScrollbar",
        )
        horizontal = ttk.Scrollbar(
            table_shell,
            orient=tk.HORIZONTAL,
            command=tree.xview,
            style="Registry.Horizontal.TScrollbar",
        )
        tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        warnings = []
        for index, relationship in enumerate(relationships):
            values = self._relationship_values(relationship)
            tree.insert("", tk.END, values=values[:6], tags=("even" if index % 2 == 0 else "odd",))
            if values[6]:
                warnings.append(values[6])
        tree.tag_configure("even", background=BG_CARD, foreground=TEXT_PRIMARY)
        tree.tag_configure("odd", background=BG_INPUT, foreground=TEXT_PRIMARY)
        return warnings

    def _render_ownership_tab(self):
        scroll = _VerticalScrollFrame(self.detail_body, background=BG_CARD)
        scroll.grid(row=0, column=0, sticky="nsew")
        relationships = _entity_ownerships(self.selected_entity)
        reverse = _entity_reverse_ownerships(self.selected_entity)
        if not relationships and not reverse:
            empty = self._section(scroll.inner, "Engar greinanlegar eignarhaldsfærslur")
            empty_label = tk.Label(
                empty,
                text="Ekki tókst að greina eignarhaldstengsl með nægilegri vissu. Skoðaðu flipann „Upprunafærslur“ til að sjá upprunalegu reitina.",
                bg=BG_INPUT,
                fg=TEXT_SECONDARY,
                font=("Segoe UI", 9),
                wraplength=600,
                justify="left",
            )
            empty_label.pack(fill=tk.X, pady=(8, 0))
            _bind_responsive_wrap(empty_label, empty, horizontal_padding=4)
            scroll.bind_mousewheel_tree()
            return
        warnings = []
        if relationships:
            warnings.extend(self._relationship_table(scroll.inner, "Hluthafar og tengdir aðilar", relationships))
        if reverse:
            warnings.extend(self._relationship_table(scroll.inner, "Félög sem aðili á hlut í", reverse))
        warnings = list(dict.fromkeys(warning for warning in warnings if warning))
        if warnings:
            section = self._section(scroll.inner, "Athugasemdir við eignarhald")
            for warning in warnings:
                warning_label = tk.Label(section, text=f"• {warning}", bg=BG_INPUT, fg=TEXT_WARNING, font=("Segoe UI", 9), anchor="w", justify="left", wraplength=620)
                warning_label.pack(fill=tk.X, pady=2)
                _bind_responsive_wrap(warning_label, section, horizontal_padding=4)
        scroll.bind_mousewheel_tree()

    def _render_sources_tab(self):
        cells = _record_cells(self.selected_entity)
        wrapper = tk.Frame(self.detail_body, bg=BG_CARD)
        wrapper.grid(row=0, column=0, sticky="nsew")
        wrapper.columnconfigure(0, weight=1)
        wrapper.rowconfigure(1, weight=1)
        toolbar = tk.Frame(wrapper, bg=BG_CARD)
        toolbar.grid(row=0, column=0, sticky="ew", padx=4, pady=(0, 7))
        tk.Label(
            toolbar,
            text=f"Tengdir reitir úr frumgögnum: {_format_count(len(cells))}",
            bg=BG_CARD,
            fg=TEXT_SECONDARY,
            font=("Segoe UI", 8),
        ).pack(side=tk.LEFT)
        create_styled_button(toolbar, "Sýna / fela", self._toggle_selected_source_cell, width=11, accent=False, height=30).pack(side=tk.RIGHT, padx=(6, 0))
        create_styled_button(toolbar, "Afrita gildi", self._copy_selected_source_cell, width=11, accent=False, height=30).pack(side=tk.RIGHT)

        columns = ("sheet", "cell", "field", "value", "raw", "formula", "link")
        tree = ttk.Treeview(wrapper, columns=columns, show="headings", style="Registry.Treeview")
        specs = (("sheet", "Upprunablað", 150), ("cell", "Reitur", 58), ("field", "Heiti", 120), ("value", "Gildi", 180), ("raw", "Upprunalegt", 160), ("formula", "Formúla", 150), ("link", "Tengill", 145))
        for column, label, width in specs:
            tree.heading(column, text=label)
            tree.column(column, width=width, minwidth=45, stretch=column in {"field", "value", "formula"})
        vertical = ttk.Scrollbar(
            wrapper,
            orient=tk.VERTICAL,
            command=tree.yview,
            style="Registry.Vertical.TScrollbar",
        )
        horizontal = ttk.Scrollbar(
            wrapper,
            orient=tk.HORIZONTAL,
            command=tree.xview,
            style="Registry.Horizontal.TScrollbar",
        )
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=1, column=0, sticky="nsew")
        vertical.grid(row=1, column=1, sticky="ns")
        horizontal.grid(row=2, column=0, sticky="ew")
        self.source_tree = tree
        self.source_cell_by_iid = {}
        for index, cell in enumerate(cells):
            iid = f"cell-{index}"
            tree.insert(
                "",
                tk.END,
                iid=iid,
                values=(
                    _source_name(cell),
                    _cell_coordinate(cell),
                    _localize_field_label(_cell_label(cell)),
                    _cell_display_value(cell, reveal_sensitive=False),
                    _cell_original_display(cell, reveal_sensitive=False),
                    _cell_formula_display(cell, reveal_sensitive=False) or "—",
                    _cell_hyperlink_display(cell, reveal_sensitive=False) or "—",
                ),
                tags=("sensitive",) if _cell_sensitive(cell) else (),
            )
            self.source_cell_by_iid[iid] = cell
        tree.tag_configure("sensitive", foreground=TEXT_WARNING)
        if cells:
            tree.selection_set("cell-0")
        else:
            tk.Label(wrapper, text="Engir reitir úr frumgögnum tengjast þessari færslu.", bg=BG_CARD, fg=TEXT_SECONDARY, font=("Segoe UI", 9)).place(relx=0.5, rely=0.5, anchor="center")

    def _selected_source_cell(self):
        if not hasattr(self, "source_tree"):
            return None, None
        selection = self.source_tree.selection()
        if not selection:
            return None, None
        iid = selection[0]
        return iid, self.source_cell_by_iid.get(iid)

    def _toggle_selected_source_cell(self):
        iid, cell = self._selected_source_cell()
        if cell is None or not _cell_sensitive(cell):
            return
        if iid in self.revealed_source_cells:
            self.revealed_source_cells.discard(iid)
        else:
            self.revealed_source_cells.add(iid)
            self._schedule_remask()
        values = list(self.source_tree.item(iid, "values"))
        values[3] = _cell_display_value(cell, reveal_sensitive=iid in self.revealed_source_cells)
        values[4] = _cell_original_display(cell, reveal_sensitive=iid in self.revealed_source_cells)
        values[5] = _cell_formula_display(cell, reveal_sensitive=iid in self.revealed_source_cells) or "—"
        values[6] = _cell_hyperlink_display(cell, reveal_sensitive=iid in self.revealed_source_cells) or "—"
        self.source_tree.item(iid, values=values)

    def _copy_selected_source_cell(self):
        _iid, cell = self._selected_source_cell()
        if cell is not None:
            self._copy_to_clipboard(_cell_display_value(cell, reveal_sensitive=True), "Gildi úr frumgögnum afritað")

    def _copy_secret(self, value):
        self._copy_to_clipboard(value, "Veflykill afritaður")

    def _copy_to_clipboard(self, value, confirmation="Afritað"):
        value = _text(value)
        if not value or root is None:
            return
        try:
            root.clipboard_clear()
            root.clipboard_append(value)
            root.update_idletasks()
            self.status_label.configure(text=confirmation, fg=TEXT_SUCCESS)
            root.after(1800, self._restore_status)
        except tk.TclError:
            self.status_label.configure(text="Ekki tókst að opna klippiborðið", fg=TEXT_ERROR)

    def _restore_status(self):
        if self.alive and self.winfo_exists() and self.model is not None:
            stats = _get(self.model, "stats")
            entity_count = _get(stats, "entity_count", default=len(self.entities))
            self.status_label.configure(text=f"{_format_count(entity_count)} tengdar færslur · 3 upprunablöð", fg=TEXT_SUCCESS)

    def _schedule_remask(self):
        if root is None:
            return
        if self.remask_after_id is not None:
            root.after_cancel(self.remask_after_id)
        self.remask_after_id = root.after(SECRET_REMASK_MS, self._remask_credentials)

    def _remask_credentials(self):
        if self.remask_after_id is not None and root is not None:
            try:
                root.after_cancel(self.remask_after_id)
            except tk.TclError:
                pass
        self.remask_after_id = None
        for row in self.credential_rows:
            if row.winfo_exists():
                row.remask()
        self.credential_rows = [row for row in self.credential_rows if row.winfo_exists()]
        if hasattr(self, "revealed_source_cells"):
            self.revealed_source_cells.clear()
        if hasattr(self, "source_tree") and self.source_tree.winfo_exists():
            for iid, cell in self.source_cell_by_iid.items():
                if self.source_tree.exists(iid):
                    values = list(self.source_tree.item(iid, "values"))
                    values[3] = _cell_display_value(cell, reveal_sensitive=False)
                    values[4] = _cell_original_display(cell, reveal_sensitive=False)
                    values[5] = _cell_formula_display(cell, reveal_sensitive=False) or "—"
                    values[6] = _cell_hyperlink_display(cell, reveal_sensitive=False) or "—"
                    self.source_tree.item(iid, values=values)

    def open_raw_explorer(self):
        if self.model is None:
            self.status_label.configure(text="Hladdu vinnubók áður en frumgögn eru opnuð", fg=TEXT_WARNING)
            return
        for dialog in reversed(self.raw_dialogs):
            if dialog.winfo_exists():
                dialog.deiconify()
                dialog.lift()
                dialog.focus_force()
                return
        dialog = _RawSourceExplorer(self, self.model, self.workbook_path)
        self.raw_dialogs.append(dialog)

    def dispose(self):
        global _active_page
        if not self.alive:
            return
        self.alive = False
        self.load_token += 1
        self._remask_credentials()
        for after_id in (self.poll_after_id, self.search_after_id, self.remask_after_id):
            if after_id is not None and root is not None:
                try:
                    root.after_cancel(after_id)
                except tk.TclError:
                    pass
        self.poll_after_id = self.search_after_id = self.remask_after_id = None
        for sequence, binding_id in self.root_bindings:
            if binding_id and root is not None:
                try:
                    root.unbind(sequence, binding_id)
                except tk.TclError:
                    pass
        self.root_bindings.clear()
        for dialog in list(self.raw_dialogs):
            if dialog.winfo_exists():
                dialog.close()
        self.raw_dialogs.clear()
        if _active_page is self:
            _active_page = None

    def go_back(self):
        previous_layout = self._previous_window_layout
        self.dispose()
        initialize_main_menu()
        if root is not None and root.winfo_exists():
            # The registry is a deliberately large workspace. Once it closes,
            # restore the exact menu layout that was active before it opened.
            _restore_window_layout(root, previous_layout)


class _RawSourceExplorer(tk.Toplevel):
    def __init__(self, page, model, workbook_path):
        super().__init__(root)
        self.page = page
        self.model = model
        self.workbook_path = workbook_path
        self.all_cells = _model_raw_cells(model)
        self.filtered_cells = list(self.all_cells)
        self.page_index = 0
        self.revealed = set()
        self.remask_after_id = None
        self.filter_after_id = None
        self.cell_by_iid = {}
        self.title("Frumgögn")
        self.configure(bg=BG_DARK)
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.transient(root)
        self._size_window()
        self._build()
        self._apply_filters()
        self.bind("<Escape>", lambda _event: self.close())
        self.bind("<Control-f>", self._focus_filter)

    def _size_window(self):
        self.update_idletasks()
        work_left, work_top, work_right, work_bottom = window_work_area(root if root is not None else self)
        work_width = max(320, work_right - work_left)
        work_height = max(320, work_bottom - work_top)
        available_width = max(320, work_width - 32)
        available_height = max(320, work_height - 32)
        width = min(1100, available_width)
        height = min(720, available_height)
        if root is not None:
            x = root.winfo_x() + (root.winfo_width() - width) // 2
            y = root.winfo_y() + (root.winfo_height() - height) // 2
        else:
            x = work_left + (work_width - width) // 2
            y = work_top + (work_height - height) // 2
        x = max(work_left + 16, min(x, work_right - width - 16))
        y = max(work_top + 8, min(y, work_bottom - height - 16))
        self.geometry(_geometry_spec(width, height, x, y))
        self.minsize(min(760, width), min(500, height))

    def _build(self):
        outer = tk.Frame(self, bg=BG_DARK)
        outer.pack(fill=tk.BOTH, expand=True, padx=18, pady=16)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1, minsize=100)

        top = tk.Frame(outer, bg=BG_DARK)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        title_column = tk.Frame(top, bg=BG_DARK)
        title_column.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_column, text="Frumgögn", bg=BG_DARK, fg=TEXT_PRIMARY, font=("Segoe UI Semibold", 16), anchor="w").pack(fill=tk.X)
        source_detail = tk.Label(
            title_column,
            text=f"Allir reitir frumgagna sem innihalda gildi, þar á meðal formúlur og gildi án fyrirsagnar · {os.path.basename(self.workbook_path)}",
            bg=BG_DARK,
            fg=TEXT_SECONDARY,
            font=("Segoe UI", 8),
            anchor="w",
            justify="left",
            wraplength=760,
        )
        source_detail.pack(fill=tk.X, pady=(3, 0))
        _bind_responsive_wrap(source_detail, title_column, horizontal_padding=12)
        create_styled_button(top, "Loka", self.close, width=8, accent=False, height=36).pack(side=tk.RIGHT)

        filter_panel, filter_inner = create_panel(outer, padx=12, pady=10)
        filter_panel.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        tk.Label(filter_inner, text="Upprunablað", bg=BG_CARD, fg=TEXT_SECONDARY, font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(0, 6))
        sheet_names = list(dict.fromkeys(_source_name(cell) for cell in self.all_cells if _source_name(cell)))
        ordered = [name for name in REGISTRY_SHEET_NAMES if name in sheet_names]
        ordered.extend(name for name in sheet_names if name not in ordered)
        self.sheet_var = tk.StringVar(value=ALL_SHEETS_LABEL)
        sheet_box = ttk.Combobox(
            filter_inner,
            textvariable=self.sheet_var,
            values=(ALL_SHEETS_LABEL, *ordered),
            state="readonly",
            style="Registry.TCombobox",
            width=27,
        )
        sheet_box.pack(side=tk.LEFT, padx=(0, 10), ipady=4)
        sheet_box.bind("<<ComboboxSelected>>", lambda _event: self._apply_filters())

        search_shell = tk.Frame(filter_inner, bg=BORDER, padx=1, pady=1)
        search_shell.pack(side=tk.LEFT, fill=tk.X, expand=True)
        search_surface = tk.Frame(search_shell, bg=BG_INPUT)
        search_surface.pack(fill=tk.BOTH, expand=True)
        self.filter_var = tk.StringVar()
        self.filter_entry = create_styled_entry(search_surface, width=36)
        self.filter_entry.configure(textvariable=self.filter_var, bg=BG_INPUT)
        self.filter_entry.pack(fill=tk.X, padx=10, ipady=7)
        self.filter_entry.bind("<FocusIn>", lambda _event: search_shell.configure(bg=ACCENT), add="+")
        self.filter_entry.bind("<FocusOut>", lambda _event: search_shell.configure(bg=BORDER), add="+")
        self.filter_entry.bind("<KeyRelease>", self._schedule_filter)
        create_styled_button(filter_inner, "Hreinsa", self._clear_filter, width=8, accent=False, height=34).pack(side=tk.LEFT, padx=(8, 0))

        table_shell = tk.Frame(outer, bg=BG_CARD, highlightbackground=BORDER, highlightthickness=1)
        table_shell.grid(row=2, column=0, sticky="nsew")
        table_shell.columnconfigure(0, weight=1)
        table_shell.rowconfigure(0, weight=1)
        columns = ("source", "section", "cell", "field", "value", "raw", "formula", "link")
        self.tree = ttk.Treeview(table_shell, columns=columns, show="headings", style="Registry.Treeview")
        specs = (
            ("source", "Upprunablað", 170),
            ("section", "Flokkur", 145),
            ("cell", "Reitur", 60),
            ("field", "Heiti", 125),
            ("value", "Gildi", 210),
            ("raw", "Upprunalegt", 170),
            ("formula", "Formúla", 165),
            ("link", "Tengill", 150),
        )
        for column, label, width in specs:
            self.tree.heading(column, text=label)
            self.tree.column(column, width=width, minwidth=45, stretch=column in {"section", "field", "value", "formula"})
        vertical = ttk.Scrollbar(
            table_shell,
            orient=tk.VERTICAL,
            command=self.tree.yview,
            style="Registry.Vertical.TScrollbar",
        )
        horizontal = ttk.Scrollbar(
            table_shell,
            orient=tk.HORIZONTAL,
            command=self.tree.xview,
            style="Registry.Horizontal.TScrollbar",
        )
        self.tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        self.tree.tag_configure("sensitive", foreground=TEXT_WARNING)
        self.tree.bind("<Double-1>", lambda _event: self.toggle_selected())

        footer = tk.Frame(outer, bg=BG_DARK)
        footer.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        footer.columnconfigure(0, weight=1)
        self.count_label = tk.Label(footer, text="", bg=BG_DARK, fg=TEXT_SECONDARY, font=("Segoe UI", 8))
        self.count_label.grid(row=0, column=0, sticky="w")
        self.notice_label = tk.Label(footer, text="", bg=BG_DARK, fg=TEXT_SUCCESS, font=("Segoe UI", 8))
        self.notice_label.grid(row=0, column=1, sticky="e", padx=(12, 0))

        footer_actions = tk.Frame(footer, bg=BG_DARK)
        footer_actions.grid(row=1, column=0, columnspan=2, sticky="e", pady=(7, 0))
        self.previous_button = create_styled_button(footer_actions, "Fyrri", self.previous_page, width=8, accent=False, height=32)
        self.previous_button.pack(side=tk.LEFT)
        self.page_label = tk.Label(footer_actions, text="", bg=BG_DARK, fg=TEXT_PRIMARY, font=("Segoe UI Semibold", 8), width=13)
        self.page_label.pack(side=tk.LEFT, padx=4)
        self.next_button = create_styled_button(footer_actions, "Næsta", self.next_page, width=7, accent=False, height=32)
        self.next_button.pack(side=tk.LEFT)
        create_styled_button(footer_actions, "Afrita gildi", self.copy_selected, width=11, accent=False, height=32).pack(side=tk.LEFT, padx=(7, 0))
        create_styled_button(footer_actions, "Sýna / fela", self.toggle_selected, width=11, accent=False, height=32).pack(side=tk.LEFT, padx=(7, 0))

    def _focus_filter(self, _event=None):
        self.filter_entry.focus_set()
        self.filter_entry.selection_range(0, tk.END)
        return "break"

    def _schedule_filter(self, _event=None):
        if self.filter_after_id is not None:
            self.after_cancel(self.filter_after_id)
        self.filter_after_id = self.after(SEARCH_DEBOUNCE_MS, self._apply_filters)

    def _clear_filter(self):
        self.filter_var.set("")
        self._apply_filters()
        self.filter_entry.focus_set()

    def _safe_filter_haystack(self, cell):
        pieces = [
            _source_name(cell),
            _text(_get(cell, "section")),
            _localize_section(_get(cell, "section")),
            _cell_coordinate(cell),
            _cell_label(cell),
            _localize_field_label(_cell_label(cell)),
        ]
        if not _cell_sensitive(cell):
            pieces.append(_cell_display_value(cell, reveal_sensitive=False))
            pieces.append(_cell_original_display(cell, reveal_sensitive=False))
            pieces.append(_cell_formula(cell))
            pieces.append(_cell_hyperlink(cell))
        return _normalize(" ".join(piece for piece in pieces if piece))

    def _apply_filters(self):
        self.filter_after_id = None
        selected_sheet = self.sheet_var.get()
        query = _normalize(self.filter_var.get())
        self.filtered_cells = []
        for cell in self.all_cells:
            if selected_sheet != ALL_SHEETS_LABEL and _source_name(cell) != selected_sheet:
                continue
            if query and query not in self._safe_filter_haystack(cell):
                continue
            self.filtered_cells.append(cell)
        self.page_index = 0
        self.revealed.clear()
        self._render_page()

    def _render_page(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.cell_by_iid = {}
        total = len(self.filtered_cells)
        page_count = max(1, (total + RAW_PAGE_SIZE - 1) // RAW_PAGE_SIZE)
        self.page_index = max(0, min(self.page_index, page_count - 1))
        start = self.page_index * RAW_PAGE_SIZE
        end = min(total, start + RAW_PAGE_SIZE)
        for local_index, cell in enumerate(self.filtered_cells[start:end]):
            absolute_index = start + local_index
            iid = f"raw-{absolute_index}"
            section = _localize_section(_get(cell, "section")) or "—"
            self.tree.insert(
                "",
                tk.END,
                iid=iid,
                values=(
                    _source_name(cell),
                    section,
                    _cell_coordinate(cell),
                    _localize_field_label(_cell_label(cell)),
                    _cell_display_value(cell, reveal_sensitive=iid in self.revealed),
                    _cell_original_display(cell, reveal_sensitive=iid in self.revealed),
                    _cell_formula_display(cell, reveal_sensitive=iid in self.revealed) or "—",
                    _cell_hyperlink_display(cell, reveal_sensitive=iid in self.revealed) or "—",
                ),
                tags=("sensitive",) if _cell_sensitive(cell) else (),
            )
            self.cell_by_iid[iid] = cell
        self.page_label.configure(text=f"Síða {self.page_index + 1} af {page_count}")
        if hasattr(self.previous_button, "set_enabled"):
            self.previous_button.set_enabled(self.page_index > 0)
        if hasattr(self.next_button, "set_enabled"):
            self.next_button.set_enabled(end < total)
        if total:
            self.count_label.configure(
                text=f"Sýni {_format_count(start + 1)}–{_format_count(end)} af {_format_count(total)} samsvarandi reitum · {_format_count(len(self.all_cells))} alls"
            )
        else:
            self.count_label.configure(text=f"Engir samsvarandi reitir · {_format_count(len(self.all_cells))} alls")
        if self.tree.get_children():
            first = self.tree.get_children()[0]
            self.tree.selection_set(first)
            self.tree.focus(first)

    def previous_page(self):
        if self.page_index > 0:
            self.page_index -= 1
            self.revealed.clear()
            self._render_page()

    def next_page(self):
        if (self.page_index + 1) * RAW_PAGE_SIZE < len(self.filtered_cells):
            self.page_index += 1
            self.revealed.clear()
            self._render_page()

    def _selected(self):
        selection = self.tree.selection()
        if not selection:
            return None, None
        iid = selection[0]
        return iid, self.cell_by_iid.get(iid)

    def toggle_selected(self):
        iid, cell = self._selected()
        if cell is None or not _cell_sensitive(cell):
            return
        if iid in self.revealed:
            self.revealed.discard(iid)
        else:
            self.revealed.add(iid)
            self._schedule_remask()
        values = list(self.tree.item(iid, "values"))
        values[4] = _cell_display_value(cell, reveal_sensitive=iid in self.revealed)
        values[5] = _cell_original_display(cell, reveal_sensitive=iid in self.revealed)
        values[6] = _cell_formula_display(cell, reveal_sensitive=iid in self.revealed) or "—"
        values[7] = _cell_hyperlink_display(cell, reveal_sensitive=iid in self.revealed) or "—"
        self.tree.item(iid, values=values)

    def copy_selected(self):
        _iid, cell = self._selected()
        if cell is None:
            return
        value = _cell_display_value(cell, reveal_sensitive=True)
        if not value:
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(value)
            self.update_idletasks()
            self.notice_label.configure(text="Gildi afritað")
            self.after(1600, lambda: self.notice_label.configure(text="") if self.winfo_exists() else None)
        except tk.TclError:
            self.notice_label.configure(text="Klippiborð ekki tiltækt", fg=TEXT_ERROR)

    def _schedule_remask(self):
        if self.remask_after_id is not None:
            self.after_cancel(self.remask_after_id)
        self.remask_after_id = self.after(SECRET_REMASK_MS, self._remask)

    def _remask(self):
        self.remask_after_id = None
        self.revealed.clear()
        self._render_page()

    def close(self):
        for after_id in (self.remask_after_id, self.filter_after_id):
            if after_id is not None:
                try:
                    self.after_cancel(after_id)
                except tk.TclError:
                    pass
        self.remask_after_id = self.filter_after_id = None
        if self in self.page.raw_dialogs:
            self.page.raw_dialogs.remove(self)
        self.destroy()


def initialize_shareholder_registry(workbook_path=None):
    """Open the searchable registry inside the application's shared content frame."""
    global _active_page
    if _active_page is not None:
        _active_page.dispose()
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Hluthafaskrá")
    _active_page = _RegistryPage(frame, workbook_path=workbook_path)
    return _active_page


# Helpful alias for integrations that use the Icelandic service name.
initialize_hluthafaskra = initialize_shareholder_registry

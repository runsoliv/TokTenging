"""Bank formatter workflow, review dialogs, compression, and success views."""

import datetime
from decimal import Decimal, ROUND_HALF_UP
import json
import os
import re
import tkinter as tk
from tkinter import filedialog, ttk
import uuid

import pandas as pd

try:
    try:
        from .auto_coder import apply_auto_debit_codes
    except ImportError:
        from auto_coder import apply_auto_debit_codes
    AUTO_CODER_AVAILABLE = True
except Exception:
    apply_auto_debit_codes = None
    AUTO_CODER_AVAILABLE = False

try:
    try:
        from .ml_runtime import apply_ml_hybrid_codes
    except ImportError:
        from ml_runtime import apply_ml_hybrid_codes
    ML_RUNTIME_AVAILABLE = True
except Exception:
    apply_ml_hybrid_codes = None
    ML_RUNTIME_AVAILABLE = False

try:
    from . import ui_components
    from .ui_components import (
        RoundedPanel,
        attach_drop_target,
        create_drop_box,
        create_panel,
        create_segmented_setting,
        create_styled_button,
        create_styled_entry,
        create_styled_label,
        finalize_fixed_action_dialog_grid,
        set_button_accent,
        _rounded_rect,
    )
    from .bank_detection import BANK_CONFIGS, detect_bank_type
    from .training_archive import (
        TRAINING_RAW_INDEX_COLUMN,
        archive_bank_import,
        create_import_session,
        refresh_paired_rows_from_coded,
        write_compression_map,
    )
    from .bank_utils import (
        autofit_excel_columns,
        format_date_as_text,
        _bank_amount_abs,
        _bank_amount_sign,
        _format_kennitala,
        _parse_bank_amount_decimal,
        _parse_icelandic_date,
        _parse_innheimta_amount,
        _round_half_up_decimal,
        _strip_accents,
    )
except ImportError:
    import ui_components
    from ui_components import (
        RoundedPanel,
        attach_drop_target,
        create_drop_box,
        create_panel,
        create_segmented_setting,
        create_styled_button,
        create_styled_entry,
        create_styled_label,
        finalize_fixed_action_dialog_grid,
        set_button_accent,
        _rounded_rect,
    )
    from bank_detection import BANK_CONFIGS, detect_bank_type
    from training_archive import (
        TRAINING_RAW_INDEX_COLUMN,
        archive_bank_import,
        create_import_session,
        refresh_paired_rows_from_coded,
        write_compression_map,
    )
    from bank_utils import (
        autofit_excel_columns,
        format_date_as_text,
        _bank_amount_abs,
        _bank_amount_sign,
        _format_kennitala,
        _parse_bank_amount_decimal,
        _parse_icelandic_date,
        _parse_innheimta_amount,
        _round_half_up_decimal,
        _strip_accents,
    )


globals().update(ui_components.get_theme("light"))

RECENT_VISIBLE_LIMIT = 3
root = None
frame = None


class CompressionMapMismatchError(RuntimeError):
    """Raised when an existing compression map cannot safely describe a workbook."""


def _not_configured(*_args, **_kwargs):
    raise RuntimeError("bank_formatter.configure() must be called before using the bank formatter")


set_page_title = _not_configured
initialize_main_menu = _not_configured
open_tok_input_with_file = _not_configured
open_file_location = _not_configured
play_success_sound = _not_configured
play_error_sound = _not_configured
ensure_window_fits = _not_configured
fit_dialog_to_content = _not_configured
schedule_layout_refresh = lambda: None
get_recent_files = lambda _key: []
remember_recent_file = lambda _key, _path: None
_short_file_name = lambda path, limit=22: os.path.basename(str(path))[:limit]
_status_with_file_detail = lambda message, _path: message
_auto_code_training_dir = _not_configured
_auto_code_key_dir = _not_configured
_auto_code_memory_path = _not_configured
_training_data_dir = _not_configured
_auto_code_mode = lambda: "legacy"
_ml_artifact_dir = _not_configured


def configure(**dependencies):
    globals().update(dependencies)


def set_theme(theme):
    globals().update(theme)


def initialize_bank_formatter():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Sniðmótun bankaskráa")
    
    display_bank_formatter_controls()

def set_bank_formatter_status(message, color=TEXT_SECONDARY, show_open=False):
    if "bank_status_label" in globals() and bank_status_label and bank_status_label.winfo_exists():
        bank_status_label.config(text=message, fg=color)
    if "bank_status_open_btn" not in globals() or not bank_status_open_btn.winfo_exists():
        return
    if not show_open and "bank_input_file_entry" in globals() and bank_input_file_entry and bank_input_file_entry.winfo_exists():
        selected_path = bank_input_file_entry.get().strip()
        show_open = bool(selected_path and os.path.exists(selected_path))
    if show_open:
        bank_status_open_btn.pack(side=tk.RIGHT, padx=(8, 8), pady=4)
    else:
        bank_status_open_btn.pack_forget()

def refresh_bank_detection_status(path):
    if not path:
        set_bank_formatter_status("Veldu Excel-skrá frá banka áður en þú endurhleður.", TEXT_WARNING)
        return None
    if "input_drop" in globals() and input_drop and input_drop.winfo_exists():
        input_drop.config(text=os.path.basename(path), fg=TEXT_PRIMARY)
    bank_type, _ = detect_bank_type(path)
    if bank_type:
        suffix = " (sjálfvirkri kóðun sleppt)" if BANK_CONFIGS[bank_type].get('skip_auto_coding') else ""
        set_bank_formatter_status(f"Banki greindur: {BANK_CONFIGS[bank_type]['name']}{suffix}", TEXT_SUCCESS)
    else:
        set_bank_formatter_status("Óþekkt bankasnið", TEXT_ERROR, show_open=True)
    if "bank_status_label" in globals() and bank_status_label and bank_status_label.winfo_exists():
        base_message = bank_status_label.cget("text").splitlines()[0]
        bank_status_label.config(text=_status_with_file_detail(base_message, path))
    schedule_layout_refresh()
    return bank_type

def refresh_bank_recent_buttons():
    if "bank_recent_files_frame" not in globals() or not bank_recent_files_frame.winfo_exists():
        return
    for widget in bank_recent_files_frame.winfo_children():
        widget.destroy()
    recent = get_recent_files("recent_bank_files")[:RECENT_VISIBLE_LIMIT]
    if not recent:
        return
    label = create_styled_label(bank_recent_files_frame, "Nýlegar skrár", size=9, color=TEXT_SECONDARY, bg=BG_CARD)
    label.pack(anchor='w', pady=(0, 4))
    row = tk.Frame(bank_recent_files_frame, bg=BG_CARD)
    row.pack(fill=tk.X)
    for index, path in enumerate(recent):
        btn = create_styled_button(
            row,
            _short_file_name(path, 19),
            lambda p=path: select_bank_input_file(p),
            width=13,
            accent=False,
            height=34,
        )
        btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6) if index < len(recent) - 1 else (0, 0))
    schedule_layout_refresh()

def select_bank_input_file(file_path, remember=True):
    if not file_path:
        return None
    path = os.path.normpath(os.path.abspath(file_path))
    if "bank_input_file_entry" in globals() and bank_input_file_entry and bank_input_file_entry.winfo_exists():
        bank_input_file_entry.delete(0, tk.END)
        bank_input_file_entry.insert(0, path)
    if "input_drop" in globals() and input_drop and input_drop.winfo_exists():
        input_drop.config(text=os.path.basename(path), fg=TEXT_PRIMARY)
    bank_type = refresh_bank_detection_status(path)
    if remember and os.path.exists(path):
        remember_recent_file("recent_bank_files", path)
        refresh_bank_recent_buttons()
    schedule_layout_refresh()
    return bank_type

def display_bank_formatter_controls():
    global bank_input_file_entry, bank_status_label, bank_output_name_entry, bank_restaurant_mode_var
    global bank_vehicle_expense_code_var
    global bank_counter_mode_var, bank_counter_code_entry
    global bank_status_open_btn, input_drop, bank_recent_files_frame

    controls_card, controls_frame = create_panel(frame, padx=22, pady=18)
    controls_card.pack(fill=tk.X, padx=28, pady=(0, 12))

    # Input file row
    # Using two drop areas stacked: top for input, bottom for output
    drop_frame = tk.Frame(controls_frame, bg=BG_CARD)
    drop_frame.pack(fill=tk.X)
    
    # Input drop
    input_label = create_styled_label(drop_frame, "Inntaksskrá", size=10, color=TEXT_PRIMARY, bold=True)
    input_label.pack(anchor='w', pady=(0, 6))
    input_drop = create_drop_box(drop_frame, "Dragðu Excel-bankaskrá hingað eða smelltu til að velja", height=5)
    input_drop.pack(fill=tk.X, pady=(0, 8))
    
    bank_input_file_entry = create_styled_entry(drop_frame, width=35)
    bank_input_file_entry.pack_forget()

    bank_recent_files_frame = tk.Frame(drop_frame, bg=BG_CARD)
    bank_recent_files_frame.pack(fill=tk.X, pady=(0, 8))
    
    def on_drop_input(path):
        if path:
            select_bank_input_file(path)
    if not attach_drop_target(input_drop, on_drop_input):
        input_drop.config(text="Smelltu til að velja. Settu upp tkinterdnd2 til að draga og sleppa.", fg=TEXT_WARNING)
    input_drop.bind("<Button-1>", lambda _event: browse_bank_input_file())

    # Required name used to identify the archive session.
    output_name_label = create_styled_label(drop_frame, "Skráarheiti", size=10, color=TEXT_PRIMARY, bold=True)
    output_name_label.pack(anchor='w', pady=(0, 5))
    bank_output_name_entry = create_styled_entry(drop_frame, width=35)
    bank_output_name_entry.pack(fill=tk.X, ipady=5, pady=(0, 8))

    options_frame = tk.Frame(drop_frame, bg=BG_CARD)
    options_frame.pack(fill=tk.X, pady=(2, 12))
    options_label = create_styled_label(options_frame, "Valkostir kóðunar", size=9, color=TEXT_SECONDARY, bg=BG_CARD)
    options_label.pack(anchor='w', pady=(0, 8))

    bank_restaurant_mode_var = tk.BooleanVar(value=False)
    restaurant_mode_setting = create_segmented_setting(
        options_frame,
        "Tegund rekstrar",
        bank_restaurant_mode_var,
        [("Almennt", False), ("Veitingarstaður", True)],
        hint="Vörukaup breytast í 2107",
    )
    restaurant_mode_setting.pack(fill=tk.X, pady=(0, 10))

    bank_vehicle_expense_code_var = tk.StringVar(value="4640")
    vehicle_expense_setting = create_segmented_setting(
        options_frame,
        "Bifreiðalykill",
        bank_vehicle_expense_code_var,
        [("4635", "4635"), ("4640", "4640")],
        hint="Sjálfkóðaðar færslur með 4635 eða 4640 nota valinn lykil.",
    )
    vehicle_expense_setting.pack(fill=tk.X, pady=(0, 10))

    counter_block = tk.Frame(options_frame, bg=BG_CARD)
    counter_block.pack(fill=tk.X)

    counter_setting_row = tk.Frame(counter_block, bg=BG_CARD)
    counter_setting_row.pack(fill=tk.X)
    counter_text_col = tk.Frame(counter_setting_row, bg=BG_CARD)
    counter_text_col.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
    create_styled_label(
        counter_text_col,
        "Nota 7810 eða annan lykil",
        size=9,
        color=TEXT_PRIMARY,
        bold=True,
        bg=BG_CARD,
    ).pack(anchor='w')
    create_styled_label(
        counter_text_col,
        "Veldu hvaða lykill er settur á bankahlið",
        size=8,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    ).pack(anchor='w', pady=(1, 0))

    bank_counter_mode_var = tk.StringVar(value="7810")
    counter_mode_row = tk.Frame(counter_setting_row, bg=BG_CARD)
    counter_mode_row.pack(side=tk.RIGHT)
    counter_mode_buttons = []

    def sync_bank_controls_layout():
        try:
            if not controls_card.winfo_exists():
                return
            controls_frame.update_idletasks()
            controls_card._sync_height()
            ensure_window_fits()
        except Exception:
            pass

    def refresh_counter_mode():
        current = bank_counter_mode_var.get()
        for button, value in counter_mode_buttons:
            set_button_accent(button, current == value)
        custom_enabled = current == "custom"
        state = "normal" if custom_enabled else "disabled"
        try:
            if custom_enabled:
                custom_code_row.pack(fill=tk.X, pady=(8, 0))
            else:
                custom_code_row.pack_forget()
            bank_counter_code_entry.configure(
                state=state,
                bg=BG_INPUT,
                fg=TEXT_PRIMARY,
                disabledbackground=BG_INPUT,
                disabledforeground=TEXT_PRIMARY,
            )
            custom_code_shell.fill = BG_INPUT
            custom_code_shell.outline = ACCENT
            custom_code_shell.inner.configure(bg=BG_INPUT)
            custom_code_shell._sync_height()
            custom_code_shell._draw()
            sync_bank_controls_layout()
            root.after_idle(sync_bank_controls_layout)
            root.after(60, sync_bank_controls_layout)
            schedule_layout_refresh()
        except Exception:
            pass

    def select_counter_mode(value):
        bank_counter_mode_var.set(value)
        refresh_counter_mode()

    for index, (text, value) in enumerate((("7810", "7810"), ("Annað", "custom"))):
        button = create_styled_button(counter_mode_row, text, lambda v=value: select_counter_mode(v), width=9, accent=False, height=34)
        button.pack(side=tk.LEFT, padx=(0, 6) if index == 0 else (0, 0))
        counter_mode_buttons.append((button, value))

    custom_code_row = tk.Frame(counter_block, bg=BG_CARD)
    create_styled_label(custom_code_row, "Annar lykill (auður reitur skilur mótlykil eftir auðan)", size=8, color=TEXT_SECONDARY, bg=BG_CARD).pack(side=tk.LEFT)
    custom_code_shell = RoundedPanel(custom_code_row, fill=BG_INPUT, outline=ACCENT, radius=8, padx=9, pady=4)
    custom_code_shell.canvas.configure(width=104)
    custom_code_shell.pack(side=tk.RIGHT)
    bank_counter_code_entry = tk.Entry(
        custom_code_shell.inner,
        width=8,
        bg=BG_INPUT,
        fg=TEXT_PRIMARY,
        disabledbackground=BG_INPUT,
        disabledforeground=TEXT_PRIMARY,
        insertbackground=TEXT_PRIMARY,
        font=('Segoe UI', 10),
        relief='flat',
        bd=0,
        justify='center',
    )
    bank_counter_code_entry.pack(fill=tk.X)
    bank_counter_code_entry.insert(0, "")
    bank_counter_mode_var.trace_add('write', lambda *_args: refresh_counter_mode())
    refresh_counter_mode()

    status_shell = RoundedPanel(controls_frame, fill=BUTTON_MUTED, outline=BORDER, radius=10, padx=0, pady=0)
    status_row = status_shell.inner
    status_shell.pack(fill=tk.X, pady=(0, 12))
    bank_status_label = tk.Label(
        status_row,
        text="Veldu inntaksskrá til að greina bankann sjálfkrafa",
        bg=BUTTON_MUTED,
        fg=TEXT_SECONDARY,
        font=('Segoe UI', 10),
        padx=14,
        pady=7,
        anchor='w',
        justify='left',
    )
    bank_status_label.configure(wraplength=360)
    bank_status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
    bank_status_open_btn = create_styled_button(
        status_row,
        "Opna skrá",
        lambda: os.startfile(bank_input_file_entry.get().strip()) if bank_input_file_entry.get().strip() else None,
        width=10,
        accent=False,
        height=42,
    )

    def refresh_bank_file():
        path = bank_input_file_entry.get().strip()
        if not path:
            set_bank_formatter_status("Veldu Excel-skrá frá banka áður en þú endurhleður.", TEXT_WARNING)
            return
        select_bank_input_file(path)

    # Buttons
    action_row = tk.Frame(controls_frame, bg=BG_CARD)
    action_row.pack(fill=tk.X, pady=(0, 8))
    run_btn = create_styled_button(action_row, "Keyra", lambda: run_bank_formatter_script(autofill_7810=True), width=38, height=42)
    run_btn.pack(fill=tk.X, expand=True)
    
    nav_row = tk.Frame(controls_frame, bg=BG_CARD)
    nav_row.pack(fill=tk.X)
    refresh_btn = create_styled_button(nav_row, "Endurhlaða skrá", refresh_bank_file, width=18, accent=False, height=42)
    refresh_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    back_btn = create_styled_button(nav_row, "← Til baka í valmynd", initialize_main_menu, width=18, accent=False)
    if hasattr(back_btn, "configure"):
        back_btn.configure(height=42)
    back_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))
    refresh_bank_recent_buttons()
    ensure_window_fits()

def browse_bank_input_file():
    filename = filedialog.askopenfilename(
        initialdir=os.getcwd(), 
        title="Veldu inntaksskrá",
        filetypes=(("Excel-skrár", "*.xlsx *.xls *.xlsm"), ("Allar skrár", "*.*"))
    )
    if filename:
        select_bank_input_file(filename)

def _bank_client_hint(input_file_path="", output_name="", output_file_path=""):
    for value in (output_name, input_file_path, output_file_path):
        if not value:
            continue
        name = os.path.splitext(os.path.basename(str(value).strip()))[0]
        if name:
            return name
    return "bank_import"


def _coded_output_filename(file_name):
    name = os.path.basename(str(file_name).strip().strip('"'))
    stem = os.path.splitext(name)[0].strip().rstrip(". ")
    stem = re.sub(r'[<>:"/\\|?*]', "_", stem).strip().rstrip(". ")
    if not stem:
        return ""
    if not stem.lower().endswith("_coded"):
        stem += "_coded"
    return f"{stem}.xlsx"


def _create_training_output_path(input_file_path, bank_type, client_hint, file_name):
    session_dir = create_import_session(
        raw_file_path=input_file_path,
        bank_type=bank_type,
        client_hint=client_hint,
        training_data_dir=_training_data_dir(),
    )
    return os.path.join(session_dir, _coded_output_filename(file_name))


def _with_training_raw_index(df):
    output = df.copy()
    output[TRAINING_RAW_INDEX_COLUMN] = output.index
    return output


def _archive_training_import(input_file_path, raw_df, coded_df, output_file_path, bank_type, bank_name, client_hint):
    try:
        if not input_file_path or not output_file_path:
            return ""
        if not os.path.exists(input_file_path) or not os.path.exists(output_file_path):
            return ""
        coded_df = _validated_id_copy(coded_df)
        result = archive_bank_import(
            raw_file_path=input_file_path,
            raw_df=raw_df,
            coded_df=coded_df,
            output_file_path=output_file_path,
            bank_type=bank_type,
            bank_name=bank_name,
            client_hint=client_hint,
            training_data_dir=_training_data_dir(),
            session_dir=os.path.dirname(output_file_path),
        )
        return result.get("session_dir", "")
    except Exception as exc:
        print(f"Training archive failed: {exc}")
        return ""


def _bank_restaurant_mode_enabled():
    if "bank_restaurant_mode_var" not in globals():
        return False
    try:
        return bool(bank_restaurant_mode_var.get())
    except Exception:
        return False


def _bank_vehicle_expense_code():
    if "bank_vehicle_expense_code_var" not in globals():
        return "4640"
    try:
        code = str(bank_vehicle_expense_code_var.get()).strip()
    except Exception:
        return "4640"
    return code if code in {"4635", "4640"} else "4640"


def _normalize_bank_counter_code(value):
    text = "" if value is None else str(value).strip()
    if not text:
        return ""
    text = text.replace(" ", "")
    try:
        dec = Decimal(text.replace(",", "."))
        if dec == dec.to_integral_value():
            text = str(int(dec))
    except Exception:
        text = re.sub(r"\.0$", "", text)
    if not re.fullmatch(r"\d{3,5}", text):
        return ""
    return int(text) if not text.startswith("0") else text


def _bank_counter_mode():
    if "bank_counter_mode_var" not in globals():
        return "7810"
    try:
        raw_mode = str(bank_counter_mode_var.get()).strip()
        mode = raw_mode.lower()
    except Exception:
        return "7810"
    if mode == "custom":
        return "custom"
    if mode == "7810":
        return "7810"
    return "7810"


def _bank_counter_fill(autofill_counter=True):
    if not autofill_counter:
        return False, ""
    mode = _bank_counter_mode()
    if mode == "custom":
        if "bank_counter_code_entry" not in globals():
            return False, ""
        try:
            code = _normalize_bank_counter_code(bank_counter_code_entry.get())
        except Exception:
            code = ""
        return (True, code) if code else (False, "")
    return True, 7810


def _apply_bank_auto_coding(
    output_df,
    source_df,
    input_file_path,
    enabled=True,
    skip=False,
    note="",
    industry_context="",
    vehicle_expense_code="",
    bank_type="",
):
    if skip:
        output_df = output_df.copy()
        output_df["STATUS"] = ""
        output_df["CONFIDENCE"] = ""
        output_df["_AUTO_CODE_SOURCE"] = "skipped"
        return output_df
    if not enabled:
        output_df = output_df.copy()
        output_df["STATUS"] = "review needed"
        output_df["CONFIDENCE"] = ""
        output_df["_AUTO_CODE_SOURCE"] = "disabled"
        return output_df
    if not AUTO_CODER_AVAILABLE or apply_auto_debit_codes is None:
        output_df = output_df.copy()
        output_df["STATUS"] = "review needed"
        output_df["CONFIDENCE"] = ""
        output_df["_AUTO_CODE_SOURCE"] = "unavailable"
        return output_df
    legacy_result = apply_auto_debit_codes(
        output_df,
        source_df=source_df,
        input_file_path=input_file_path,
        enabled=True,
        training_dir=_auto_code_training_dir(),
        key_dir=_auto_code_key_dir(),
        memory_path=_auto_code_memory_path(),
        training_data_dir=_training_data_dir(),
        industry_context=industry_context,
        vehicle_expense_code=vehicle_expense_code,
    )
    try:
        mode = str(_auto_code_mode() or "legacy").strip().lower()
    except Exception:
        mode = "legacy"
    if mode != "ml_hybrid" or not ML_RUNTIME_AVAILABLE or apply_ml_hybrid_codes is None:
        return legacy_result
    try:
        return apply_ml_hybrid_codes(
            legacy_result,
            source_df=source_df,
            input_file_path=input_file_path,
            bank_type=bank_type,
            vehicle_expense_code=vehicle_expense_code,
            artifact_dir=_ml_artifact_dir(),
        )
    except Exception:
        # Runtime inference is optional.  Its own adapter is fail-closed, and
        # this final boundary guarantees the historical path remains usable.
        return legacy_result


def _auto_review_cluster_key(text):
    try:
        if pd.isna(text):
            return "missing text"
    except Exception:
        pass
    normalized = _strip_accents(str(text or "")).lower()
    normalized = re.sub(r"[^a-z0-9]+", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return "missing text" if normalized in {"", "nan", "none", "nat", "na"} else normalized


def _auto_review_confidence(value):
    try:
        if pd.isna(value):
            return 0
    except Exception:
        pass
    try:
        return int(float(str(value).strip()))
    except Exception:
        return 0


def _should_review_auto_code_row(row):
    sign = str(row.get("Positive/Negative", "")).strip()
    if sign != "-":
        return False
    source = str(row.get("_AUTO_CODE_SOURCE", "")).strip().lower()
    if "_AUTO_CODE_SOURCE" in row.index:
        return source in {"fallback", "person_default"}
    status = str(row.get("STATUS", "")).strip().lower()
    confidence = _auto_review_confidence(row.get("CONFIDENCE"))
    if status in {"coded", ""} and confidence > 45:
        return False
    return status in {"review needed", "review", "fallback_review", "fallback review"} or confidence <= 45


def _low_confidence_auto_code_clusters(df):
    clusters = {}
    for idx, row in df.iterrows():
        if not _should_review_auto_code_row(row):
            continue
        key = _auto_review_cluster_key(row.get("TEXT", ""))
        if not key:
            continue
        clusters.setdefault(key, []).append(idx)
    return sorted(
        clusters.items(),
        key=lambda item: (
            min(_auto_review_confidence(df.loc[idx].get("CONFIDENCE")) for idx in item[1]),
            -len(item[1]),
            item[0],
        ),
    )


def _format_auto_review_value(value, limit=41):
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat", "<na>"}:
        return ""
    limit = max(4, int(limit or 41))
    return text[:limit - 3] + "..." if len(text) > limit else text


def _auto_review_display_name(value):
    text = _format_auto_review_value(value, limit=94).strip()
    if not text or text == "missing text":
        return "Texta vantar"
    missing_split = re.fullmatch(r"missing text split (.+)", text)
    if missing_split:
        return f"Texta vantar · hluti {missing_split.group(1)}"
    return text


def _format_auto_review_amount(value):
    amount = _parse_bank_amount_decimal(value)
    if amount is None:
        return _format_auto_review_value(value, limit=18)
    sign = "-" if amount < 0 else ""
    amount = abs(amount)
    if amount == amount.to_integral_value():
        return f"{sign}{int(amount):,}kr"
    text = f"{amount:,.2f}".rstrip("0").rstrip(".")
    return f"{sign}{text}kr"


def _attach_themed_tree_scrollbar(parent, tree, width=12):
    scrollbar = tk.Canvas(parent, width=width, bg=BG_INPUT, highlightthickness=0, bd=0)
    thumb = scrollbar.create_rectangle(3, 2, width - 3, 30, fill=ACCENT, outline="")
    scroll_state = {"first": 0.0, "last": 1.0}

    def redraw(first=None, last=None):
        if first is not None and last is not None:
            try:
                scroll_state["first"] = float(first)
                scroll_state["last"] = float(last)
            except Exception:
                pass
        h = max(1, scrollbar.winfo_height())
        first_pos = max(0.0, min(1.0, scroll_state["first"]))
        last_pos = max(first_pos, min(1.0, scroll_state["last"]))
        top = int(first_pos * h)
        bottom = int(last_pos * h)
        min_thumb = min(34, h)
        if bottom - top < min_thumb:
            center = (top + bottom) // 2
            top = max(2, center - min_thumb // 2)
            bottom = min(h - 2, top + min_thumb)
            top = max(2, bottom - min_thumb)
        scrollbar.coords(thumb, 3, top, width - 3, max(top + 1, bottom))

    def move_to_event(event):
        h = max(1, scrollbar.winfo_height())
        thumb_height = max(1, scrollbar.coords(thumb)[3] - scrollbar.coords(thumb)[1])
        movable = max(1, h - thumb_height)
        fraction = max(0.0, min(1.0, (event.y - (thumb_height / 2)) / movable))
        tree.yview_moveto(fraction)
        return "break"

    scrollbar.bind("<Configure>", lambda _event: redraw())
    scrollbar.bind("<Button-1>", move_to_event)
    scrollbar.bind("<B1-Motion>", move_to_event)
    scrollbar.bind("<Enter>", lambda _event: scrollbar.itemconfigure(thumb, fill=ACCENT_HOVER))
    scrollbar.bind("<Leave>", lambda _event: scrollbar.itemconfigure(thumb, fill=ACCENT))
    tree.configure(yscrollcommand=redraw)
    tree.after_idle(lambda: redraw(*tree.yview()))
    return scrollbar


def _show_auto_code_cluster_review(cluster_df, cluster_name, current_index, total_clusters, can_go_back=False):
    result = {"action": "skip", "code": "", "row_indices": list(cluster_df.index), "split_batches": []}

    dialog = tk.Toplevel(root)
    dialog.title("Yfirfara sjálfvirka kóðun")
    dialog.configure(bg=BG_DARK)
    dialog.transient(root)
    dialog.grab_set()
    dialog.resizable(True, True)
    dialog.minsize(560, 420)

    outer = tk.Frame(dialog, bg=BG_DARK)
    outer.pack(fill=tk.BOTH, expand=True, padx=14, pady=14)
    outer.columnconfigure(0, weight=1)
    outer.rowconfigure(1, weight=1)

    active_indices = list(cluster_df.index)
    selected_positions = set()
    split_batches = []
    item_to_position = {}
    item_base_tags = {}

    confidences = [_auto_review_confidence(row.get("CONFIDENCE")) for _, row in cluster_df.iterrows()]
    min_confidence = min(confidences) if confidences else 0
    debit_values = []
    for _, row in cluster_df.iterrows():
        debit_value = _format_auto_review_value(row.get("DEBIT"), limit=12).strip()
        if debit_value and debit_value not in debit_values:
            debit_values.append(debit_value)
    current_debit = debit_values[0] if len(debit_values) == 1 else ("Blandað" if debit_values else "-")

    header_panel, header = create_panel(outer, padx=18, pady=14)
    header_panel.grid(row=0, column=0, sticky="ew")

    eyebrow_row = tk.Frame(header, bg=BG_CARD)
    eyebrow_row.pack(fill=tk.X)
    create_styled_label(
        eyebrow_row,
        "YFIRFERÐ SJÁLFVIRKRAR KÓÐUNAR",
        size=8,
        color=TEXT_SECONDARY,
        bold=True,
        bg=BG_CARD,
    ).pack(side=tk.LEFT)
    progress_label = create_styled_label(
        eyebrow_row,
        f"{current_index} af {total_clusters}",
        size=8,
        color=TEXT_ON_ACCENT,
        bold=True,
        bg=ACCENT,
    )
    progress_label.pack(side=tk.RIGHT, ipadx=10, ipady=4)

    create_styled_label(
        header,
        "Yfirfara færsluflokk",
        size=14,
        bold=True,
        bg=BG_CARD,
    ).pack(anchor="w", pady=(6, 2))
    cluster_label = create_styled_label(
        header,
        _auto_review_display_name(cluster_name),
        size=10,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    )
    cluster_label.configure(wraplength=500, justify="left")
    cluster_label.pack(anchor="w")

    stats_row = tk.Frame(header, bg=BG_CARD)
    stats_row.pack(fill=tk.X, pady=(10, 0))

    def add_stat(parent, label, value, pad):
        tile = tk.Frame(parent, bg=BG_INPUT, highlightbackground=BORDER, highlightthickness=1)
        tile.pack(side=tk.LEFT, padx=pad)
        value_label = create_styled_label(tile, f"{label}: {value}", size=8, color=TEXT_PRIMARY, bold=True, bg=BG_INPUT)
        value_label.pack(
            anchor="w",
            padx=10,
            pady=5,
        )
        return value_label

    rows_stat = add_stat(stats_row, "Færslur", str(len(cluster_df)), (0, 8))
    confidence_stat = add_stat(stats_row, "Lægsta vissustig", f"{min_confidence}", (0, 8))
    debit_stat = add_stat(stats_row, "Núverandi debetlykill", current_debit, (0, 0))

    table_panel, table_inner = create_panel(outer, padx=0, pady=0)
    table_panel.grid(row=1, column=0, sticky="nsew", pady=(10, 0))

    table_title_row = tk.Frame(table_inner, bg=BG_CARD)
    table_title_row.pack(fill=tk.X, padx=14, pady=(10, 6))
    create_styled_label(
        table_title_row,
        "Færslur með lágt vissustig",
        size=10,
        color=TEXT_PRIMARY,
        bold=True,
        bg=BG_CARD,
    ).pack(side=tk.LEFT)
    selected_count_label = create_styled_label(table_title_row, "", size=8, color=TEXT_SECONDARY, bg=BG_CARD)
    selected_count_label.pack(side=tk.RIGHT)

    table_body = tk.Frame(table_inner, bg=BG_CARD)
    table_body.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 12))

    tree_style = ttk.Style(dialog)
    try:
        tree_style.configure(
            "AutoReview.Treeview",
            background=BG_CARD,
            fieldbackground=BG_CARD,
            foreground=TEXT_PRIMARY,
            borderwidth=0,
            rowheight=27,
            font=("Segoe UI", 9),
        )
        tree_style.configure(
            "AutoReview.Treeview.Heading",
            background=BG_INPUT,
            foreground=TEXT_SECONDARY,
            relief="flat",
            font=("Segoe UI Semibold", 8),
        )
        tree_style.configure(
            "AutoReview.Vertical.TScrollbar",
            background=BUTTON_MUTED,
            troughcolor=BG_INPUT,
            bordercolor=BORDER,
            arrowcolor=TEXT_SECONDARY,
            darkcolor=BUTTON_MUTED,
            lightcolor=BUTTON_MUTED,
            relief="flat",
            width=12,
        )
        tree_style.map("AutoReview.Treeview", background=[("selected", ACCENT)], foreground=[("selected", TEXT_ON_ACCENT)])
        tree_style.map(
            "AutoReview.Vertical.TScrollbar",
            background=[("active", ACCENT), ("pressed", ACCENT_HOVER)],
            arrowcolor=[("active", TEXT_ON_ACCENT), ("pressed", TEXT_ON_ACCENT)],
        )
    except Exception:
        pass

    visible_rows = min(5, max(3, len(cluster_df)))
    columns = ("date", "amount", "debit", "confidence", "text")
    tree = ttk.Treeview(
        table_body,
        columns=columns,
        show="headings",
        selectmode="none",
        height=visible_rows,
        style="AutoReview.Treeview",
    )
    tree.heading("date", text="DATE", anchor=tk.W)
    tree.heading("amount", text="AMOUNT", anchor=tk.E)
    tree.heading("debit", text="DEBIT", anchor=tk.E)
    tree.heading("confidence", text="VISSA", anchor=tk.E)
    tree.heading("text", text="TEXT", anchor=tk.W)
    tree.column("date", width=90, minwidth=82, stretch=False, anchor=tk.W)
    tree.column("amount", width=92, minwidth=86, stretch=False, anchor=tk.E)
    tree.column("debit", width=64, minwidth=58, stretch=False, anchor=tk.E)
    tree.column("confidence", width=58, minwidth=52, stretch=False, anchor=tk.E)
    tree.column("text", width=190, minwidth=140, stretch=True, anchor=tk.W)
    tree.tag_configure("even", background=BG_CARD, foreground=TEXT_PRIMARY)
    tree.tag_configure("odd", background=BG_INPUT, foreground=TEXT_PRIMARY)
    tree.tag_configure("marked", background=ACCENT, foreground=TEXT_ON_ACCENT)

    yscrollbar = _attach_themed_tree_scrollbar(table_body, tree)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    yscrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def refresh_selected_count():
        selected_count = len(selected_positions)
        selected_text = "1 valin" if selected_count == 1 else f"{selected_count} valdar"
        selected_count_label.config(text=selected_text if selected_positions else "")

    def refresh_stats():
        rows_text = str(len(active_indices))
        if split_batches:
            split_count = sum(len(batch) for batch in split_batches)
            split_text = "1 færð" if split_count == 1 else f"{split_count} færðar"
            rows_text += f" ({split_text})"
        current_rows = [cluster_df.loc[idx] for idx in active_indices if idx in cluster_df.index]
        current_conf = [_auto_review_confidence(row.get("CONFIDENCE")) for row in current_rows]
        current_debits = []
        for row in current_rows:
            debit_value = _format_auto_review_value(row.get("DEBIT"), limit=12).strip()
            if debit_value and debit_value not in current_debits:
                current_debits.append(debit_value)
        rows_stat.config(text=f"Færslur: {rows_text}")
        confidence_stat.config(text=f"Lægsta vissustig: {min(current_conf) if current_conf else 0}")
        debit_stat.config(text=f"Núverandi debetlykill: {current_debits[0] if len(current_debits) == 1 else ('Blandað' if current_debits else '-')}")

    def row_values(row):
        return (
            _format_auto_review_value(row.get("DATE"), limit=18),
            _format_auto_review_amount(row.get("AMOUNT")),
            _format_auto_review_value(row.get("DEBIT"), limit=12),
            _format_auto_review_value(row.get("CONFIDENCE"), limit=8),
            _format_auto_review_value(row.get("TEXT"), limit=90),
        )

    def refresh_table():
        selected_positions.clear()
        item_to_position.clear()
        item_base_tags.clear()
        for item in tree.get_children():
            tree.delete(item)
        for position, idx in enumerate(active_indices):
            if idx not in cluster_df.index:
                continue
            base_tag = "odd" if position % 2 else "even"
            item = tree.insert("", tk.END, values=row_values(cluster_df.loc[idx]), tags=(base_tag,))
            item_to_position[item] = position
            item_base_tags[item] = base_tag
        refresh_selected_count()
        refresh_stats()

    def toggle_tree_row(event):
        region = tree.identify("region", event.x, event.y)
        if region not in {"cell", "tree"}:
            return None
        item = tree.identify_row(event.y)
        if not item:
            return "break"
        position = item_to_position.get(item)
        if position is None:
            return "break"
        if position in selected_positions:
            selected_positions.remove(position)
            tree.item(item, tags=(item_base_tags.get(item, "even"),))
        else:
            selected_positions.add(position)
            tree.item(item, tags=("marked",))
        refresh_selected_count()
        status_label.config(text="", fg=TEXT_WARNING)
        return "break"

    tree.bind("<Button-1>", toggle_tree_row)

    action_panel, action_inner = create_panel(outer, padx=14, pady=12)
    action_panel.grid(row=2, column=0, sticky="ew", pady=(10, 0))

    form_row = tk.Frame(action_inner, bg=BG_CARD)
    form_row.pack(fill=tk.X)
    create_styled_label(form_row, "Debetlykill", size=9, color=TEXT_SECONDARY, bg=BG_CARD).pack(side=tk.LEFT, padx=(0, 10))
    entry_shell = tk.Frame(form_row, bg=BORDER, padx=1, pady=1)
    entry_shell.pack(side=tk.LEFT)
    code_entry = create_styled_entry(entry_shell, width=12)
    code_entry.pack(ipady=5, padx=0, pady=0)
    status_label = create_styled_label(form_row, "", size=8, color=TEXT_WARNING, bg=BG_CARD)
    status_label.pack(side=tk.LEFT, padx=(12, 0))

    button_row = tk.Frame(action_inner, bg=BG_CARD)
    button_row.pack(fill=tk.X, pady=(10, 0))
    secondary_button_row = tk.Frame(action_inner, bg=BG_CARD)
    secondary_button_row.pack(fill=tk.X, pady=(6, 0))

    def store_current_split_result():
        result["row_indices"] = list(active_indices)
        result["split_batches"] = [list(batch) for batch in split_batches if batch]

    def submit_code():
        raw = code_entry.get().strip()
        code = _normalize_bank_counter_code(raw)
        if not code:
            status_label.config(text="Sláðu inn 3–5 stafa lykil.")
            return
        result["action"] = "code"
        result["code"] = code
        store_current_split_result()
        dialog.destroy()

    def skip_cluster():
        result["action"] = "skip"
        store_current_split_result()
        dialog.destroy()

    def skip_all():
        result["action"] = "skip_all"
        dialog.destroy()

    def go_back():
        if not can_go_back:
            return
        result["action"] = "back"
        dialog.destroy()

    def exclude_selected():
        positions = sorted(selected_positions)
        if not positions:
            status_label.config(text="Veldu færslur sem eiga að fara í annan flokk.")
            return
        if len(positions) >= len(active_indices):
            status_label.config(text="Ekki er hægt að færa allar færslurnar úr þessum flokki.")
            return
        excluded_indices = [active_indices[pos] for pos in positions if 0 <= pos < len(active_indices)]
        for pos in reversed(positions):
            if 0 <= pos < len(active_indices):
                active_indices.pop(pos)
        split_batches.append(excluded_indices)
        moved_count = len(excluded_indices)
        moved_text = "1 færsla færð" if moved_count == 1 else f"{moved_count} færslur færðar"
        status_label.config(text=f"{moved_text} í annan flokk.", fg=TEXT_SUCCESS)
        refresh_table()

    refresh_table()

    button_row.grid_columnconfigure(0, weight=1, uniform="review_actions")
    button_row.grid_columnconfigure(1, weight=1, uniform="review_actions")
    code_btn = create_styled_button(button_row, "Kóða", submit_code, width=10, height=34)
    code_btn.grid(row=0, column=0, sticky="ew", padx=(0, 5))
    exclude_btn = create_styled_button(button_row, "Færa valdar", exclude_selected, width=16, height=34, accent=False)
    exclude_btn.grid(row=0, column=1, sticky="ew", padx=(5, 0))

    secondary_column_count = 3 if can_go_back else 2
    for column in range(secondary_column_count):
        secondary_button_row.grid_columnconfigure(column, weight=1, uniform="review_secondary_actions")
    if can_go_back:
        back_btn = create_styled_button(secondary_button_row, "Til baka", go_back, width=10, height=34, accent=False)
        back_btn.grid(row=0, column=0, sticky="ew", padx=(0, 5))
    skip_btn = create_styled_button(secondary_button_row, "Sleppa", skip_cluster, width=10, height=34, accent=False)
    skip_column = 1 if can_go_back else 0
    skip_btn.grid(
        row=0,
        column=skip_column,
        sticky="ew",
        padx=((5, 5) if can_go_back else (0, 5)),
    )
    skip_all_btn = create_styled_button(secondary_button_row, "Sleppa öllum", skip_all, width=12, height=34, accent=False)
    skip_all_btn.grid(row=0, column=skip_column + 1, sticky="ew", padx=(5, 0))

    code_entry.bind("<Return>", lambda _event: submit_code())
    dialog.bind("<Return>", lambda _event: submit_code())
    dialog.bind("<Escape>", lambda _event: skip_cluster())
    dialog.bind("<Alt-Left>", lambda _event: go_back())
    dialog.protocol("WM_DELETE_WINDOW", skip_cluster)
    finalize_fixed_action_dialog_grid(outer, header_panel, table_panel, action_panel)
    fit_dialog_to_content(dialog, min_width=560, min_height=420, preferred_width=600)
    code_entry.focus_set()
    root.wait_window(dialog)
    return result


def review_low_confidence_auto_code_clusters(df, enabled=True):
    if not enabled:
        return df
    if "root" not in globals() or root is None:
        return df
    clusters = _low_confidence_auto_code_clusters(df)
    if not clusters:
        return df

    reviewed = df.copy()
    original_rows = {}
    cluster_index = 0
    while cluster_index < len(clusters):
        total = len(clusters)
        cluster_key, row_indices = clusters[cluster_index]
        available_indices = [idx for idx in row_indices if idx in reviewed.index]
        if not available_indices:
            cluster_index += 1
            continue
        if cluster_index not in original_rows:
            original_rows[cluster_index] = reviewed.loc[available_indices].copy()
        cluster_df = reviewed.loc[available_indices].copy()
        display_candidates = []
        if "TEXT" in cluster_df:
            for value in cluster_df["TEXT"].tolist():
                display_value = _format_auto_review_value(value, limit=94).strip()
                if display_value:
                    display_candidates.append(display_value)
        display_name = (
            str(pd.Series(display_candidates).mode().iloc[0])
            if display_candidates and not pd.Series(display_candidates).mode().empty
            else _auto_review_display_name(cluster_key)
        )
        choice = _show_auto_code_cluster_review(
            cluster_df,
            display_name,
            cluster_index + 1,
            total,
            can_go_back=cluster_index > 0,
        )
        if choice.get("action") == "skip_all":
            break
        if choice.get("action") == "back":
            cluster_index = max(0, cluster_index - 1)
            previous_key, previous_rows = clusters[cluster_index]
            previous_indices = [idx for idx in previous_rows if idx in reviewed.index]
            if cluster_index in original_rows and previous_indices:
                reviewed.loc[previous_indices, original_rows[cluster_index].columns] = original_rows[cluster_index]
            continue
        current_indices = [idx for idx in choice.get("row_indices", available_indices) if idx in reviewed.index]
        if not current_indices:
            current_indices = list(available_indices)
        split_batches = []
        for batch in choice.get("split_batches", []):
            clean_batch = [idx for idx in batch if idx in reviewed.index and idx not in current_indices]
            if clean_batch:
                split_batches.append(clean_batch)
        if split_batches or current_indices != available_indices:
            clusters[cluster_index] = (cluster_key, current_indices)
            if cluster_index in original_rows:
                kept_original_rows = [idx for idx in current_indices if idx in original_rows[cluster_index].index]
                if kept_original_rows:
                    original_rows[cluster_index] = original_rows[cluster_index].loc[kept_original_rows].copy()
        for split_number, split_indices in enumerate(split_batches, start=1):
            clusters.append((f"{cluster_key} split {cluster_index + 1}.{split_number}", split_indices))
        if choice.get("action") != "code":
            cluster_index += 1
            continue
        code = choice.get("code")
        reviewed.loc[current_indices, "DEBIT"] = code
        reviewed.loc[current_indices, "STATUS"] = "coded"
        reviewed.loc[current_indices, "CONFIDENCE"] = "MN"
        reviewed.loc[current_indices, "_AUTO_CODE_SOURCE"] = "manual_review"
        cluster_index += 1
    return reviewed


def _strip_auto_code_internal_columns(df):
    internal_prefixes = ("_AUTO_CODE_", "_TRAINING_")
    return df.drop(
        columns=[column for column in df.columns if str(column).startswith(internal_prefixes)],
        errors="ignore",
    )


def _prepare_bank_excel_output(df):
    output = _strip_auto_code_internal_columns(_validated_id_copy(df))
    if "CONFIDENCE" in output.columns and "Cnf" not in output.columns:
        output.rename(columns={"CONFIDENCE": "Cnf"}, inplace=True)
    preferred = ["DATE", "TEXT", "DEBIT", "Cnf", "ID", "AMOUNT", "CREDIT", "Positive/Negative", "STATUS"]
    ordered = [column for column in preferred if column in output.columns]
    ordered.extend(column for column in output.columns if column not in ordered)
    return output[ordered]


def _validated_id_copy(df):
    """Return a copy whose ID column contains kennitölur only."""
    output = df.copy()
    if "ID" in output.columns:
        output["ID"] = output["ID"].apply(_format_kennitala)
    return output


def _normalize_tx_text_key(value):
    return _auto_review_cluster_key(value)


def _normalize_tx_code(value):
    code = _normalize_bank_counter_code(value)
    return str(code) if code else ""


def _tx_sign(row):
    sign = str(row.get("Positive/Negative", "")).strip()
    if sign in {"+", "-"}:
        return sign
    amount = _parse_bank_amount_decimal(row.get("AMOUNT"))
    if amount is None:
        return ""
    return "+" if amount >= 0 else "-"


def _tx_code_pair(row):
    sign = _tx_sign(row)
    debit = _normalize_tx_code(row.get("DEBIT"))
    credit = _normalize_tx_code(row.get("CREDIT"))
    if sign not in {"+", "-"}:
        return "", "", ""
    return sign, debit, credit


def _format_compressed_amount(value):
    amount = _parse_bank_amount_decimal(value)
    if amount is None:
        return value
    if amount == amount.to_integral_value():
        return int(amount)
    return float(amount)


def _sum_tx_amounts(df):
    total = Decimal("0")
    for value in df["AMOUNT"]:
        amount = _parse_bank_amount_decimal(value)
        if amount is None:
            return None
        total += abs(amount)
    return total


def _compressed_reference_path(output_file_path):
    folder = os.path.dirname(os.path.abspath(output_file_path)) if output_file_path else _training_data_dir()
    stem, ext = os.path.splitext(os.path.basename(output_file_path))
    ext = ext or ".xlsx"
    default_path = os.path.join(folder, f"{stem}_compressed_txs{ext}")
    map_path = os.path.join(folder, "compression_map.json")
    try:
        with open(map_path, "r", encoding="utf-8") as handle:
            map_payload = json.load(handle)
        remembered_name = str(map_payload.get("reference_filename", "")).strip()
        mapped_coded_name = str(map_payload.get("coded_filename", "")).strip()
        same_workbook = (
            not mapped_coded_name
            or os.path.normcase(mapped_coded_name)
            == os.path.normcase(os.path.basename(output_file_path))
        )
        if same_workbook and remembered_name and os.path.basename(remembered_name) == remembered_name:
            remembered_path = os.path.join(folder, remembered_name)
            if (
                os.path.normcase(os.path.abspath(remembered_path))
                != os.path.normcase(os.path.abspath(output_file_path))
                and os.path.exists(remembered_path)
            ):
                return remembered_path
    except Exception:
        pass
    return default_path


_COMPRESSION_STORE_LABELS = {
    "atlantsolia": "ATLANTSOLÍA",
    "bakarameistarinn": "BAKARAMEISTARINN",
    "bauhaus": "BAUHAUS",
    "bilanaust": "BÍLANAUST",
    "bonus": "BÓNUS",
    "byko": "BYKO",
    "clippers": "CLIPPERS",
    "dominos": "DOMINO'S",
    "elko": "ELKO",
    "frumherji": "FRUMHERJI",
    "hagkaup": "HAGKAUP",
    "husasmidjan": "HÚSASMIÐJAN",
    "ikea": "IKEA",
    "iskraft": "ÍSKRAFT",
    "kfc": "KFC",
    "kjorbudin": "KJÖRBÚÐIN",
    "kronan": "KRÓNAN",
    "murbudin": "MÚRBÚÐIN",
    "n1": "N1",
    "netto": "NETTÓ",
    "nova": "NOVA",
    "ob": "OB",
    "olis": "OLÍS",
    "orkan": "ORKAN",
    "parka": "PARKA",
    "siminn": "SÍMINN",
    "subway": "SUBWAY",
    "vinbudin": "VÍNBÚÐIN",
}
_COMPRESSION_MERCHANT_ALIASES = {
    "facebk": ("facebook", "FACEBOOK"),
    "facebook": ("facebook", "FACEBOOK"),
}
_COMPRESSION_PROCESSOR_PREFIX = re.compile(
    r"^\s*(?:str|paypal|sumup)(?:\s*\*+\s*|\s+)",
    re.IGNORECASE,
)
_COMPRESSION_CARD_SUFFIX = re.compile(
    r"\s*,?\s*\*{2,}\s*-?\s*\d+\s*$",
    re.IGNORECASE,
)
_COMPRESSION_TERMINAL_SUFFIX = re.compile(
    r"\s*-\s*(?:sölunóta|solunota|greiðsluseðill|greidslusedill)\s*$",
    re.IGNORECASE,
)
_AUTO_COMPRESSION_BLOCKED_ACCOUNT_CODES = {"6220"}
_COMPRESSION_REFERENCE_PRINT_COLUMNS = ["DATE", "TEXT", "DEBIT", "ID", "AMOUNT"]
_COMPRESSED_TRANSACTION_STATUS = "Þjöppuð færsla"
_COMPRESSION_GENERIC_REFERENCE_STEMS = {
    "faersla",
    "greidsla",
    "kort",
    "reikningur",
    "transaction",
    "visa",
}


def _clean_compression_merchant_text(value):
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value or "").strip()
    if text.lower() in {"", "nan", "none", "nat", "<na>", "missing text"}:
        return ""
    text = _COMPRESSION_PROCESSOR_PREFIX.sub("", text)
    text = _COMPRESSION_CARD_SUFFIX.sub("", text)
    text = _COMPRESSION_TERMINAL_SUFFIX.sub("", text)
    return re.sub(r"\s+", " ", text).strip(" ,;-")


def _is_opaque_compression_reference(token):
    token = str(token or "").strip().lower()
    return (
        len(token) >= 6
        and token.isalnum()
        and any(character.isalpha() for character in token)
        and any(character.isdigit() for character in token)
    )


def _compression_reference_stem(cleaned, tokens):
    """Remove one strong machine-reference suffix without fuzzy name matching."""
    if len(tokens) < 2 or not _is_opaque_compression_reference(tokens[-1]):
        return "", ""
    normalized_stem = " ".join(tokens[:-1]).strip()
    if not normalized_stem or normalized_stem in _COMPRESSION_GENERIC_REFERENCE_STEMS:
        return "", ""
    display_stem = cleaned.rsplit(maxsplit=1)[0].strip(" ,;-")
    if not display_stem:
        return "", ""
    return normalized_stem, display_stem


def _compression_store_identity(value):
    """Return a conservative merchant key and output label for compression."""
    raw_text = "" if value is None else str(value).strip()
    processor_wrapped = bool(_COMPRESSION_PROCESSOR_PREFIX.match(raw_text))
    cleaned = _clean_compression_merchant_text(value)
    normalized = _normalize_tx_text_key(cleaned)
    if not normalized or normalized == "missing text":
        return "", ""
    tokens = normalized.split()
    first = tokens[0]

    if first == "costco":
        token_set = set(tokens)
        if token_set & {"fuel", "bensin", "bensinstod"}:
            return "store:costco:fuel", "COSTCO FUEL"
        if token_set & {"online", "netverslun", "vefverslun"}:
            return "store:costco:online", "COSTCO ONLINE"
        if token_set & {"warehouse", "wholesale", "verslun"}:
            return "store:costco:warehouse", "COSTCO WAREHOUSE"
        return "store:costco", "COSTCO"

    merchant_alias = _COMPRESSION_MERCHANT_ALIASES.get(first)
    if merchant_alias:
        canonical_key, merchant_label = merchant_alias
        return f"store:{canonical_key}", merchant_label

    store_label = _COMPRESSION_STORE_LABELS.get(first)
    if store_label:
        source_label = cleaned.split(maxsplit=1)[0].strip(" ,;-").upper()
        return f"store:{first}", source_label or store_label

    if not processor_wrapped:
        reference_stem, reference_label = _compression_reference_stem(cleaned, tokens)
        if reference_stem:
            return f"exact:{reference_stem}", reference_label.upper()

    # Unknown merchants use an exact cleaned key. This deliberately avoids
    # merging unrelated people or processor payments merely because their
    # first word is the same.
    return f"exact:{normalized}", cleaned


def _tx_counter_account(row, sign=None):
    sign = sign or _tx_sign(row)
    if sign == "-":
        return _normalize_tx_code(row.get("CREDIT"))
    if sign == "+":
        return _normalize_tx_code(row.get("DEBIT"))
    return ""


def _tx_category_account(row, sign=None):
    """Return the expense/income account rather than the bank counter-account."""
    sign = sign or _tx_sign(row)
    if sign == "-":
        return _normalize_tx_code(row.get("DEBIT"))
    if sign == "+":
        return _normalize_tx_code(row.get("CREDIT"))
    return ""


def _find_compressible_transaction_clusters(df, minimum_count=5, excluded_indices=None):
    required = {"DATE", "TEXT", "DEBIT", "AMOUNT", "CREDIT"}
    if not required.issubset(set(df.columns)):
        return []

    excluded_indices = set(excluded_indices or ())
    blocked_merchant_keys = set()
    for _, row in df.iterrows():
        merchant_key, _ = _compression_store_identity(row.get("TEXT"))
        if not merchant_key:
            continue
        sign = _tx_sign(row)
        category_account = _tx_category_account(row, sign)
        if category_account in _AUTO_COMPRESSION_BLOCKED_ACCOUNT_CODES:
            blocked_merchant_keys.add(merchant_key)

    grouped = {}
    for position, (idx, row) in enumerate(df.iterrows()):
        if idx in excluded_indices:
            continue
        merchant_key, display_name = _compression_store_identity(row.get("TEXT"))
        if not merchant_key or merchant_key in blocked_merchant_keys:
            continue
        sign, debit_code, credit_code = _tx_code_pair(row)
        if not sign:
            continue
        counter_account = _tx_counter_account(row, sign)
        if not counter_account:
            continue
        amount_ok = _parse_bank_amount_decimal(row.get("AMOUNT")) is not None
        grouped.setdefault((merchant_key, sign, counter_account), []).append({
            "idx": idx,
            "position": position,
            "display_name": display_name,
            "debit_code": debit_code,
            "credit_code": credit_code,
            "amount_ok": amount_ok,
        })

    clusters = []
    for (merchant_key, sign, counter_account), rows_info in grouped.items():
        if len(rows_info) < minimum_count:
            continue
        if any(not item["amount_ok"] for item in rows_info):
            continue
        row_indices = [item["idx"] for item in rows_info]
        cluster_df = df.loc[row_indices].copy()
        total = _sum_tx_amounts(cluster_df)
        if total is None:
            continue
        display_name = rows_info[0]["display_name"]
        debit_code = rows_info[0]["debit_code"]
        credit_code = rows_info[0]["credit_code"]
        code_pairs = {
            (item["debit_code"], item["credit_code"])
            for item in rows_info
        }
        clusters.append({
            "key": (merchant_key, sign, counter_account),
            "display_name": display_name,
            "summary_text": display_name,
            "sign": sign,
            "counter_account": counter_account,
            "debit_code": debit_code,
            "credit_code": credit_code,
            "mixed_codes": len(code_pairs) > 1,
            "row_indices": row_indices,
            "count": len(row_indices),
            "total": total,
            "first_position": rows_info[0]["position"],
        })
    return sorted(clusters, key=lambda item: item["first_position"])


def _show_transaction_compression_review(cluster_df, cluster, current_index, total_clusters):
    result = {"action": "skip"}

    dialog = tk.Toplevel(root)
    dialog.title("Sjálfvirk þjöppun færslna")
    dialog.configure(bg=BG_DARK)
    dialog.transient(root)
    dialog.grab_set()
    dialog.resizable(True, True)
    dialog.minsize(560, 500)

    outer = tk.Frame(dialog, bg=BG_DARK)
    outer.pack(fill=tk.BOTH, expand=True, padx=14, pady=14)
    outer.columnconfigure(0, weight=1)
    outer.rowconfigure(1, weight=1)

    header_panel, header = create_panel(outer, padx=18, pady=14)
    header_panel.grid(row=0, column=0, sticky="ew")
    eyebrow_row = tk.Frame(header, bg=BG_CARD)
    eyebrow_row.pack(fill=tk.X)
    create_styled_label(
        eyebrow_row,
        "SJÁLFVIRK ÞJÖPPUN FÆRSLNA",
        size=8,
        color=TEXT_SECONDARY,
        bold=True,
        bg=BG_CARD,
    ).pack(side=tk.LEFT)
    progress_label = create_styled_label(
        eyebrow_row,
        f"{current_index} af {total_clusters}",
        size=8,
        color=TEXT_ON_ACCENT,
        bold=True,
        bg=ACCENT,
    )
    progress_label.pack(side=tk.RIGHT, ipadx=10, ipady=4)
    title = create_styled_label(
        header,
        "Tillaga að þjöppun",
        size=14,
        bold=True,
        bg=BG_CARD,
    )
    title.pack(anchor="w", pady=(6, 2))
    active_indices = [idx for idx in cluster["row_indices"] if idx in cluster_df.index]
    original_count = len(active_indices)

    def active_total():
        total = _sum_tx_amounts(cluster_df.loc[active_indices]) if active_indices else Decimal("0")
        return total if total is not None else Decimal("0")

    def active_code_pairs():
        pairs = []
        for idx in active_indices:
            _sign, debit_code, credit_code = _tx_code_pair(cluster_df.loc[idx])
            pairs.append((debit_code, credit_code))
        return pairs

    def active_code_label():
        pairs = active_code_pairs()
        if not pairs:
            return "-"
        first_debit, first_credit = pairs[0]
        if len(set(pairs)) == 1:
            return f"debet {first_debit or '-'} / kredit {first_credit or '-'}"
        return "Blandað"

    name_label = create_styled_label(
        header,
        _auto_review_display_name(cluster.get("display_name")),
        size=9,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    )
    name_label.configure(wraplength=560, justify="left")
    name_label.pack(anchor="w")
    mixed_code_note = create_styled_label(
        header,
        "",
        size=8,
        color=TEXT_WARNING,
        bg=BG_CARD,
    )
    mixed_code_note.configure(wraplength=560, justify="left")
    mixed_code_note.pack(anchor="w", pady=(4, 0))

    stats_row = tk.Frame(header, bg=BG_CARD)
    stats_row.pack(fill=tk.X, pady=(10, 0))

    def add_stat(parent, label, value, pad):
        tile = tk.Frame(parent, bg=BG_INPUT, highlightbackground=BORDER, highlightthickness=1)
        tile.pack(side=tk.LEFT, padx=pad)
        value_label = create_styled_label(
            tile,
            f"{label}: {value}",
            size=8,
            color=TEXT_PRIMARY,
            bold=True,
            bg=BG_INPUT,
        )
        value_label.pack(anchor="w", padx=10, pady=5)
        return value_label

    rows_stat = add_stat(stats_row, "Færslur", str(len(active_indices)), (0, 8))
    sign_stat = add_stat(stats_row, "Merki", cluster.get("sign") or "-", (0, 8))
    code_stat = add_stat(stats_row, "Lyklar", active_code_label(), (0, 8))
    total_stat = add_stat(stats_row, "Samtals", _format_auto_review_amount(active_total()), (0, 0))

    table_panel, table_inner = create_panel(outer, padx=0, pady=0)
    table_panel.grid(row=1, column=0, sticky="nsew", pady=(10, 0))

    table_title_row = tk.Frame(table_inner, bg=BG_CARD)
    table_title_row.pack(fill=tk.X, padx=14, pady=(10, 6))
    create_styled_label(
        table_title_row,
        "Færslur",
        size=10,
        color=TEXT_PRIMARY,
        bold=True,
        bg=BG_CARD,
    ).pack(side=tk.LEFT)
    selected_count_label = create_styled_label(table_title_row, "", size=8, color=TEXT_SECONDARY, bg=BG_CARD)
    selected_count_label.pack(side=tk.RIGHT)

    table_body = tk.Frame(table_inner, bg=BG_CARD)
    table_body.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 12))

    tree_style = ttk.Style(dialog)
    try:
        tree_style.configure(
            "CompressReview.Treeview",
            background=BG_CARD,
            fieldbackground=BG_CARD,
            foreground=TEXT_PRIMARY,
            borderwidth=0,
            rowheight=27,
            font=("Segoe UI", 9),
        )
        tree_style.configure(
            "CompressReview.Treeview.Heading",
            background=BG_INPUT,
            foreground=TEXT_SECONDARY,
            relief="flat",
            font=("Segoe UI Semibold", 8),
        )
        tree_style.map("CompressReview.Treeview", background=[("selected", ACCENT)], foreground=[("selected", TEXT_ON_ACCENT)])
    except Exception:
        pass

    visible_rows = min(6, max(4, len(active_indices)))
    columns = ("date", "amount", "debit", "credit", "sign", "text")
    tree = ttk.Treeview(
        table_body,
        columns=columns,
        show="headings",
        selectmode="none",
        height=visible_rows,
        style="CompressReview.Treeview",
    )
    tree.heading("date", text="DATE", anchor=tk.W)
    tree.heading("amount", text="AMOUNT", anchor=tk.E)
    tree.heading("debit", text="DEBIT", anchor=tk.E)
    tree.heading("credit", text="CREDIT", anchor=tk.E)
    tree.heading("sign", text="MERKI", anchor=tk.CENTER)
    tree.heading("text", text="TEXT", anchor=tk.W)
    tree.column("date", width=90, minwidth=82, stretch=False, anchor=tk.W)
    tree.column("amount", width=92, minwidth=86, stretch=False, anchor=tk.E)
    tree.column("debit", width=64, minwidth=58, stretch=False, anchor=tk.E)
    tree.column("credit", width=64, minwidth=58, stretch=False, anchor=tk.E)
    tree.column("sign", width=48, minwidth=42, stretch=False, anchor=tk.CENTER)
    tree.column("text", width=170, minwidth=120, stretch=True, anchor=tk.W)
    tree.tag_configure("even", background=BG_CARD, foreground=TEXT_PRIMARY)
    tree.tag_configure("odd", background=BG_INPUT, foreground=TEXT_PRIMARY)
    tree.tag_configure("marked", background=ACCENT, foreground=TEXT_ON_ACCENT)

    scrollbar = _attach_themed_tree_scrollbar(table_body, tree)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    selected_positions = set()
    item_to_position = {}
    item_base_tags = {}

    action_panel, action_inner = create_panel(outer, padx=14, pady=12)
    action_panel.grid(row=2, column=0, sticky="ew", pady=(10, 0))

    status_label = create_styled_label(action_inner, "", size=8, color=TEXT_SECONDARY, bg=BG_CARD)
    status_label.pack(anchor="w")

    def refresh_summary():
        excluded_count = original_count - len(active_indices)
        rows_text = f"{len(active_indices)}"
        if excluded_count:
            excluded_text = "1 undanskilin" if excluded_count == 1 else f"{excluded_count} undanskildar"
            rows_text += f" ({excluded_text})"
        rows_stat.config(text=f"Færslur: {rows_text}")
        sign_stat.config(text=f"Merki: {cluster.get('sign') or '-'}")
        pairs = active_code_pairs()
        code_stat.config(text=f"Lyklar: {active_code_label()}")
        if len(set(pairs)) > 1:
            first_debit, first_credit = pairs[0]
            mixed_code_note.config(
                text=(
                    "Færslurnar eru með mismunandi lykla. Við þjöppun verða "
                    "lyklar efstu færslunnar sem eftir er notaðir: "
                    f"debet {first_debit or '-'} / kredit {first_credit or '-'}."
                )
            )
        else:
            mixed_code_note.config(text="")
        total_stat.config(text=f"Samtals: {_format_auto_review_amount(active_total())}")

    def refresh_selected_count():
        selected_count = len(selected_positions)
        selected_text = "1 valin" if selected_count == 1 else f"{selected_count} valdar"
        selected_count_label.config(text=selected_text if selected_positions else "")

    def row_values(row):
        return (
            _format_auto_review_value(row.get("DATE"), limit=18),
            _format_auto_review_amount(row.get("AMOUNT")),
            _format_auto_review_value(row.get("DEBIT"), limit=12),
            _format_auto_review_value(row.get("CREDIT"), limit=12),
            _format_auto_review_value(row.get("Positive/Negative"), limit=4),
            _format_auto_review_value(row.get("TEXT"), limit=90),
        )

    def refresh_table():
        selected_positions.clear()
        item_to_position.clear()
        item_base_tags.clear()
        for item in tree.get_children():
            tree.delete(item)
        for position, idx in enumerate(active_indices):
            base_tag = "odd" if position % 2 else "even"
            item = tree.insert("", tk.END, values=row_values(cluster_df.loc[idx]), tags=(base_tag,))
            item_to_position[item] = position
            item_base_tags[item] = base_tag
        refresh_selected_count()
        refresh_summary()

    def toggle_tree_row(event):
        region = tree.identify("region", event.x, event.y)
        if region not in {"cell", "tree"}:
            return None
        item = tree.identify_row(event.y)
        if not item:
            return "break"
        position = item_to_position.get(item)
        if position is None:
            return "break"
        if position in selected_positions:
            selected_positions.remove(position)
            tree.item(item, tags=(item_base_tags.get(item, "even"),))
        else:
            selected_positions.add(position)
            tree.item(item, tags=("marked",))
        refresh_selected_count()
        status_label.config(text="", fg=TEXT_SECONDARY)
        return "break"

    tree.bind("<Button-1>", toggle_tree_row)

    def exclude_selected():
        positions = sorted(selected_positions, reverse=True)
        if not positions:
            status_label.config(text="Veldu eina eða fleiri færslur til að undanskilja.", fg=TEXT_WARNING)
            return
        if len(positions) >= len(active_indices):
            status_label.config(text="Ekki er hægt að undanskilja allar færslurnar í flokknum.", fg=TEXT_WARNING)
            return
        for pos in positions:
            if 0 <= pos < len(active_indices):
                active_indices.pop(pos)
        excluded_count = len(positions)
        excluded_text = "1 færsla undanskilin" if excluded_count == 1 else f"{excluded_count} færslur undanskildar"
        status_label.config(text=f"{excluded_text} frá þessari þjöppun.", fg=TEXT_SUCCESS)
        refresh_table()

    refresh_table()

    button_row = tk.Frame(action_inner, bg=BG_CARD)
    button_row.pack(fill=tk.X, pady=(10, 0))

    def compress_cluster():
        if len(active_indices) < 5:
            status_label.config(
                text="Að minnsta kosti fimm færslur þurfa að vera eftir fyrir sjálfvirka þjöppun.",
                fg=TEXT_WARNING,
            )
            return
        result["action"] = "compress"
        result["row_indices"] = list(active_indices)
        result["count"] = len(active_indices)
        result["total"] = active_total()
        result["summary_text"] = cluster.get("summary_text", cluster.get("display_name", ""))
        first_sign, first_debit, first_credit = _tx_code_pair(cluster_df.loc[active_indices[0]])
        result["sign"] = first_sign
        result["debit_code"] = first_debit
        result["credit_code"] = first_credit
        dialog.destroy()

    def skip_cluster():
        result["action"] = "skip"
        dialog.destroy()

    def skip_all():
        result["action"] = "skip_all"
        dialog.destroy()

    compress_btn = create_styled_button(button_row, "Þjappa", compress_cluster, width=12, height=34)
    compress_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    exclude_btn = create_styled_button(button_row, "Undanskilja valdar", exclude_selected, width=16, height=34, accent=False)
    exclude_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    skip_btn = create_styled_button(button_row, "Sleppa", skip_cluster, width=10, height=34, accent=False)
    skip_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    skip_all_btn = create_styled_button(button_row, "Sleppa öllum", skip_all, width=12, height=34, accent=False)
    skip_all_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

    dialog.bind("<Return>", lambda _event: compress_cluster())
    dialog.bind("<Escape>", lambda _event: skip_cluster())
    dialog.bind("<Delete>", lambda _event: exclude_selected())
    dialog.protocol("WM_DELETE_WINDOW", skip_cluster)
    finalize_fixed_action_dialog_grid(outer, header_panel, table_panel, action_panel)
    fit_dialog_to_content(dialog, min_width=560, min_height=500, preferred_width=620)
    dialog.focus_set()
    root.wait_window(dialog)
    return result


def _drag_positions_between(previous_position, current_position):
    """Return every row position crossed by one drag movement, in order."""
    if current_position is None:
        return ()
    if previous_position is None:
        return (current_position,)
    step = 1 if current_position >= previous_position else -1
    return range(previous_position, current_position + step, step)


def _top_selected_transaction_text(df, row_indices, selected_positions):
    """Return the full TEXT value from the topmost valid selected row."""
    valid_positions = sorted(
        position
        for position in selected_positions
        if 0 <= position < len(row_indices)
    )
    if not valid_positions:
        return ""
    value = df.loc[row_indices[valid_positions[0]]].get("TEXT", "")
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    return "" if text.lower() in {"nan", "none", "nat", "<na>"} else text


def _show_manual_transaction_compression_dialog(df):
    result = {"action": "cancel"}

    required = {"DATE", "TEXT", "DEBIT", "AMOUNT", "CREDIT"}
    if not required.issubset(set(df.columns)):
        return {
            "action": "error",
            "message": "Handvirk þjöppun krefst dálkanna DATE, TEXT, DEBIT, AMOUNT og CREDIT.",
        }

    dialog = tk.Toplevel(root)
    dialog.title("Handvirk þjöppun færslna")
    dialog.configure(bg=BG_DARK)
    dialog.transient(root)
    dialog.grab_set()
    dialog.resizable(True, True)
    dialog.minsize(760, 560)

    outer = tk.Frame(dialog, bg=BG_DARK)
    outer.pack(fill=tk.BOTH, expand=True, padx=14, pady=14)
    outer.columnconfigure(0, weight=1)
    outer.rowconfigure(1, weight=1)

    header_panel, header = create_panel(outer, padx=18, pady=14)
    header_panel.grid(row=0, column=0, sticky="ew")
    eyebrow_row = tk.Frame(header, bg=BG_CARD)
    eyebrow_row.pack(fill=tk.X)
    create_styled_label(
        eyebrow_row,
        "HANDVIRK ÞJÖPPUN FÆRSLNA",
        size=8,
        color=TEXT_SECONDARY,
        bold=True,
        bg=BG_CARD,
    ).pack(side=tk.LEFT)
    row_count_label = create_styled_label(
        eyebrow_row,
        f"Færslur: {len(df)}",
        size=8,
        color=TEXT_ON_ACCENT,
        bold=True,
        bg=ACCENT,
    )
    row_count_label.pack(side=tk.RIGHT, ipadx=10, ipady=4)
    create_styled_label(
        header,
        "Veldu færslur til að þjappa",
        size=14,
        bold=True,
        bg=BG_CARD,
    ).pack(anchor="w", pady=(6, 2))
    help_label = create_styled_label(
        header,
        "Færslur birtast í sömu röð og í vinnubókinni. Veldu tvær eða fleiri færslur með sama formerki og sláðu síðan inn texta fyrir þjöppuðu færsluna. Debet- og kreditlyklar fyrstu völdu færslunnar verða notaðir.",
        size=9,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    )
    help_label.configure(wraplength=690, justify="left")
    help_label.pack(anchor="w")

    stats_row = tk.Frame(header, bg=BG_CARD)
    stats_row.pack(fill=tk.X, pady=(10, 0))

    def add_stat(parent, label, value, pad):
        tile = tk.Frame(parent, bg=BG_INPUT, highlightbackground=BORDER, highlightthickness=1)
        tile.pack(side=tk.LEFT, padx=pad)
        value_label = create_styled_label(
            tile,
            f"{label}: {value}",
            size=8,
            color=TEXT_PRIMARY,
            bold=True,
            bg=BG_INPUT,
        )
        value_label.pack(anchor="w", padx=10, pady=5)
        return value_label

    selected_stat = add_stat(stats_row, "Valdar", "0", (0, 8))
    total_stat = add_stat(stats_row, "Samtals", "0kr", (0, 8))
    code_stat = add_stat(stats_row, "Lyklar", "-", (0, 0))

    table_panel, table_inner = create_panel(outer, padx=0, pady=0)
    table_panel.grid(row=1, column=0, sticky="nsew", pady=(10, 0))

    table_title_row = tk.Frame(table_inner, bg=BG_CARD)
    table_title_row.pack(fill=tk.X, padx=14, pady=(10, 6))
    create_styled_label(
        table_title_row,
        "Færslur í vinnubók",
        size=10,
        color=TEXT_PRIMARY,
        bold=True,
        bg=BG_CARD,
    ).pack(side=tk.LEFT)
    selected_count_label = create_styled_label(table_title_row, "", size=8, color=TEXT_SECONDARY, bg=BG_CARD)
    selected_count_label.pack(side=tk.RIGHT)

    table_body = tk.Frame(table_inner, bg=BG_CARD)
    table_body.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 12))

    tree_style = ttk.Style(dialog)
    try:
        tree_style.configure(
            "ManualCompress.Treeview",
            background=BG_CARD,
            fieldbackground=BG_CARD,
            foreground=TEXT_PRIMARY,
            borderwidth=0,
            rowheight=27,
            font=("Segoe UI", 9),
        )
        tree_style.configure(
            "ManualCompress.Treeview.Heading",
            background=BG_INPUT,
            foreground=TEXT_SECONDARY,
            relief="flat",
            font=("Segoe UI Semibold", 8),
        )
        tree_style.map("ManualCompress.Treeview", background=[("selected", ACCENT)], foreground=[("selected", TEXT_ON_ACCENT)])
    except Exception:
        pass

    columns = ("row", "date", "amount", "debit", "credit", "sign", "text")
    visible_rows = min(14, max(7, len(df)))
    tree = ttk.Treeview(
        table_body,
        columns=columns,
        show="headings",
        selectmode="none",
        height=visible_rows,
        style="ManualCompress.Treeview",
    )
    tree.heading("row", text="RÖÐ", anchor=tk.E)
    tree.heading("date", text="DATE", anchor=tk.W)
    tree.heading("amount", text="AMOUNT", anchor=tk.E)
    tree.heading("debit", text="DEBIT", anchor=tk.E)
    tree.heading("credit", text="CREDIT", anchor=tk.E)
    tree.heading("sign", text="MERKI", anchor=tk.CENTER)
    tree.heading("text", text="TEXT", anchor=tk.W)
    tree.column("row", width=52, minwidth=46, stretch=False, anchor=tk.E)
    tree.column("date", width=90, minwidth=82, stretch=False, anchor=tk.W)
    tree.column("amount", width=92, minwidth=86, stretch=False, anchor=tk.E)
    tree.column("debit", width=64, minwidth=58, stretch=False, anchor=tk.E)
    tree.column("credit", width=64, minwidth=58, stretch=False, anchor=tk.E)
    tree.column("sign", width=48, minwidth=42, stretch=False, anchor=tk.CENTER)
    tree.column("text", width=260, minwidth=160, stretch=True, anchor=tk.W)
    tree.tag_configure("even", background=BG_CARD, foreground=TEXT_PRIMARY)
    tree.tag_configure("odd", background=BG_INPUT, foreground=TEXT_PRIMARY)
    tree.tag_configure("marked", background=ACCENT, foreground=TEXT_ON_ACCENT)

    scrollbar = _attach_themed_tree_scrollbar(table_body, tree)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    row_indices = list(df.index)
    selected_positions = set()
    item_to_position = {}
    position_to_item = {}
    item_base_tags = {}
    drag_select_state = {
        "active": False,
        "mark": True,
        "seen": set(),
        "last_position": None,
        "pointer_y": None,
        "scroll_job": None,
    }
    drag_edge_size = 38
    drag_scroll_interval_ms = 55

    action_panel, action_inner = create_panel(outer, padx=14, pady=12)
    action_panel.grid(row=2, column=0, sticky="ew", pady=(10, 0))

    name_row = tk.Frame(action_inner, bg=BG_CARD)
    name_row.pack(fill=tk.X)
    create_styled_label(name_row, "Texti þjappaðrar færslu", size=9, color=TEXT_SECONDARY, bg=BG_CARD).pack(side=tk.LEFT, padx=(0, 10))
    text_var = tk.StringVar()
    text_entry = create_styled_entry(name_row, width=34)
    text_entry.configure(textvariable=text_var)
    text_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
    text_autofill_state = {"programmatic": False, "user_edited": False}

    def track_manual_text_edit(*_args):
        if not text_autofill_state["programmatic"]:
            text_autofill_state["user_edited"] = True

    text_var.trace_add("write", track_manual_text_edit)
    status_label = create_styled_label(action_inner, "", size=8, color=TEXT_SECONDARY, bg=BG_CARD)
    status_label.pack(anchor="w", pady=(8, 0))

    def selected_indices():
        return [row_indices[pos] for pos in sorted(selected_positions) if 0 <= pos < len(row_indices)]

    def selected_rows_df():
        indices = selected_indices()
        return df.loc[indices].copy() if indices else pd.DataFrame(columns=df.columns)

    def selected_total(selected_df=None):
        if selected_df is None:
            selected_df = selected_rows_df()
        if selected_df.empty:
            return Decimal("0")
        return _sum_tx_amounts(selected_df)

    def selected_code_label(selected_df=None):
        if selected_df is None:
            selected_df = selected_rows_df()
        if selected_df.empty:
            return "-"
        pairs = []
        for _, row in selected_df.iterrows():
            sign, debit_code, credit_code = _tx_code_pair(row)
            pairs.append((sign, debit_code, credit_code))
        unique_pairs = {pair for pair in pairs}
        if len(unique_pairs) == 1:
            sign, debit_code, credit_code = next(iter(unique_pairs))
            return f"{sign or '-'} / debet {debit_code or '-'} / kredit {credit_code or '-'}"
        return "Blandað; lyklar fyrstu völdu færslunnar verða notaðir"

    def refresh_compression_text(force_reset=False):
        if text_autofill_state["user_edited"] and not force_reset:
            return
        suggested_text = _top_selected_transaction_text(
            df,
            row_indices,
            selected_positions,
        )
        if text_var.get() != suggested_text:
            text_autofill_state["programmatic"] = True
            try:
                text_var.set(suggested_text)
            finally:
                text_autofill_state["programmatic"] = False
        text_autofill_state["user_edited"] = False

    def refresh_selected_summary():
        count = len(selected_positions)
        selected_count_label.config(text=("1 valin" if count == 1 else f"{count} valdar") if count else "")
        selected_stat.config(text=f"Valdar: {count}")
        selected_df = selected_rows_df()
        total = selected_total(selected_df)
        total_stat.config(text=f"Samtals: {_format_auto_review_amount(total) if total is not None else 'ógild upphæð'}")
        code_stat.config(text=f"Lyklar: {selected_code_label(selected_df)}")
        refresh_compression_text()

    def row_values(position, row):
        sign, _debit_code, _credit_code = _tx_code_pair(row)
        return (
            str(position + 2),
            _format_auto_review_value(row.get("DATE"), limit=18),
            _format_auto_review_amount(row.get("AMOUNT")),
            _format_auto_review_value(row.get("DEBIT"), limit=12),
            _format_auto_review_value(row.get("CREDIT"), limit=12),
            sign or _format_auto_review_value(row.get("Positive/Negative"), limit=4),
            _format_auto_review_value(row.get("TEXT"), limit=110),
        )

    def refresh_table():
        item_to_position.clear()
        position_to_item.clear()
        item_base_tags.clear()
        for item in tree.get_children():
            tree.delete(item)
        for position, idx in enumerate(row_indices):
            base_tag = "odd" if position % 2 else "even"
            tags = ("marked",) if position in selected_positions else (base_tag,)
            item = tree.insert("", tk.END, values=row_values(position, df.loc[idx]), tags=tags)
            item_to_position[item] = position
            position_to_item[position] = item
            item_base_tags[item] = base_tag
        refresh_selected_summary()

    def row_item_at_event(event):
        region = tree.identify("region", event.x, event.y)
        if region not in {"cell", "tree"}:
            return None, None
        item = tree.identify_row(event.y)
        if not item:
            return None, None
        position = item_to_position.get(item)
        if position is None:
            return None, None
        return item, position

    def apply_drag_selection(position):
        if position in drag_select_state["seen"]:
            return False
        item = position_to_item.get(position)
        if item is None:
            return False
        drag_select_state["seen"].add(position)
        if drag_select_state["mark"]:
            selected_positions.add(position)
            tree.item(item, tags=("marked",))
        else:
            selected_positions.discard(position)
            tree.item(item, tags=(item_base_tags.get(item, "even"),))
        return True

    def apply_drag_through(position):
        if position is None:
            return
        previous = drag_select_state["last_position"]
        changed = False
        for crossed_position in _drag_positions_between(previous, position):
            changed = apply_drag_selection(crossed_position) or changed
        drag_select_state["last_position"] = position
        if changed:
            status_label.config(text="", fg=TEXT_SECONDARY)
            refresh_selected_summary()

    def edge_scroll_direction(pointer_y):
        height = max(1, tree.winfo_height())
        if pointer_y <= drag_edge_size:
            return -1
        if pointer_y >= height - drag_edge_size:
            return 1
        return 0

    def row_position_near_y(pointer_y, direction=0):
        height = max(1, tree.winfo_height())
        y = max(0, min(int(pointer_y), height - 1))
        item = tree.identify_row(y)
        if item:
            return item_to_position.get(item)

        # The top probe can land in the heading and the bottom probe can land
        # below the last fully visible row. Search inward only as far as one
        # row plus the heading instead of jumping the viewport by a page.
        probe_limit = min(height, drag_edge_size + 28)
        if direction < 0:
            probes = range(y + 1, min(height, y + probe_limit))
        elif direction > 0:
            probes = range(y - 1, max(-1, y - probe_limit), -1)
        else:
            return None
        for probe_y in probes:
            item = tree.identify_row(probe_y)
            if item:
                return item_to_position.get(item)
        return None

    def cancel_drag_autoscroll():
        job = drag_select_state.get("scroll_job")
        drag_select_state["scroll_job"] = None
        if job is not None:
            try:
                tree.after_cancel(job)
            except (tk.TclError, RuntimeError):
                pass

    def run_drag_autoscroll():
        drag_select_state["scroll_job"] = None
        if not drag_select_state["active"]:
            return
        try:
            pointer_y = tree.winfo_pointery() - tree.winfo_rooty()
            drag_select_state["pointer_y"] = pointer_y
            direction = edge_scroll_direction(pointer_y)
            if not direction:
                return

            before = tree.yview()
            tree.yview_scroll(direction, "units")
            after = tree.yview()
            position = row_position_near_y(pointer_y, direction)
            if position is not None:
                apply_drag_through(position)

            if after != before and drag_select_state["active"]:
                drag_select_state["scroll_job"] = tree.after(
                    drag_scroll_interval_ms,
                    run_drag_autoscroll,
                )
        except (tk.TclError, RuntimeError):
            drag_select_state["scroll_job"] = None

    def update_drag_autoscroll():
        pointer_y = drag_select_state.get("pointer_y")
        if (
            drag_select_state["active"]
            and pointer_y is not None
            and edge_scroll_direction(pointer_y)
        ):
            if drag_select_state.get("scroll_job") is None:
                drag_select_state["scroll_job"] = tree.after(
                    drag_scroll_interval_ms,
                    run_drag_autoscroll,
                )
        else:
            cancel_drag_autoscroll()

    def start_drag_selection(event):
        item, position = row_item_at_event(event)
        if position is None:
            return None
        cancel_drag_autoscroll()
        drag_select_state["active"] = True
        drag_select_state["mark"] = position not in selected_positions
        drag_select_state["seen"] = set()
        drag_select_state["last_position"] = None
        drag_select_state["pointer_y"] = event.y
        apply_drag_through(position)
        return "break"

    def drag_select_rows(event):
        if not drag_select_state["active"]:
            return None
        drag_select_state["pointer_y"] = event.y
        item, position = row_item_at_event(event)
        if position is None:
            direction = edge_scroll_direction(event.y)
            position = row_position_near_y(event.y, direction)
        if position is not None:
            apply_drag_through(position)
        update_drag_autoscroll()
        return "break"

    def end_drag_selection(_event=None):
        was_active = drag_select_state["active"]
        cancel_drag_autoscroll()
        drag_select_state["active"] = False
        drag_select_state["seen"] = set()
        drag_select_state["last_position"] = None
        drag_select_state["pointer_y"] = None
        return "break" if was_active else None

    def toggle_tree_row(event):
        item, position = row_item_at_event(event)
        if position is None:
            return None
        if position in selected_positions:
            selected_positions.remove(position)
            tree.item(item, tags=(item_base_tags.get(item, "even"),))
        else:
            selected_positions.add(position)
            tree.item(item, tags=("marked",))
        status_label.config(text="", fg=TEXT_SECONDARY)
        refresh_selected_summary()
        return "break"

    def clear_selection():
        selected_positions.clear()
        refresh_compression_text(force_reset=True)
        refresh_table()
        status_label.config(text="Valið var hreinsað.", fg=TEXT_SECONDARY)

    def validate_manual_selection():
        indices = selected_indices()
        if len(indices) < 2:
            return None, "Veldu að minnsta kosti tvær færslur."
        new_text = text_entry.get().strip()
        if not new_text:
            return None, "Sláðu inn texta fyrir þjöppuðu færsluna."
        selected_df = df.loc[indices].copy()
        total = _sum_tx_amounts(selected_df)
        if total is None:
            return None, "Allar valdar færslur verða að vera með gildar upphæðir."
        pairs = []
        for _, row in selected_df.iterrows():
            sign, debit_code, credit_code = _tx_code_pair(row)
            if sign not in {"+", "-"}:
                return None, "Valdar færslur verða að vera með gilt formerki."
            pairs.append((sign, debit_code, credit_code))
        signs = {sign for sign, _debit_code, _credit_code in pairs}
        if len(signs) != 1:
            return None, "Valdar færslur verða að vera með sama formerki."
        sign = next(iter(signs))
        _first_sign, debit_code, credit_code = pairs[0]
        return {
            "action": "compress",
            "row_indices": indices,
            "count": len(indices),
            "total": total,
            "summary_text": new_text,
            "display_name": new_text,
            "sign": sign,
            "debit_code": debit_code,
            "credit_code": credit_code,
            "first_position": int(min(sorted(selected_positions))) if selected_positions else 0,
        }, ""

    def compress_selected():
        selection, message = validate_manual_selection()
        if not selection:
            status_label.config(text=message, fg=TEXT_WARNING)
            return
        result.update(selection)
        end_drag_selection()
        dialog.destroy()

    def cancel_dialog():
        result["action"] = "cancel"
        end_drag_selection()
        dialog.destroy()

    tree.bind("<ButtonPress-1>", start_drag_selection)
    tree.bind("<B1-Motion>", drag_select_rows)
    tree.bind("<ButtonRelease-1>", end_drag_selection)
    refresh_table()

    button_row = tk.Frame(action_inner, bg=BG_CARD)
    button_row.pack(fill=tk.X, pady=(10, 0))
    compress_btn = create_styled_button(button_row, "Þjappa völdum", compress_selected, width=18, height=34)
    compress_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    clear_btn = create_styled_button(button_row, "Hreinsa", clear_selection, width=10, height=34, accent=False)
    clear_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    cancel_btn = create_styled_button(button_row, "Hætta við", cancel_dialog, width=10, height=34, accent=False)
    cancel_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

    dialog.bind("<Return>", lambda _event: compress_selected())
    dialog.bind("<Escape>", lambda _event: cancel_dialog())
    dialog.protocol("WM_DELETE_WINDOW", cancel_dialog)
    finalize_fixed_action_dialog_grid(outer, header_panel, table_panel, action_panel)
    fit_dialog_to_content(dialog, min_width=760, min_height=560, preferred_width=820)
    text_entry.focus_set()
    root.wait_window(dialog)
    return result


def _compression_map_path(output_file_path):
    return os.path.join(os.path.dirname(os.path.abspath(output_file_path)), "compression_map.json")


def _update_compression_map_file_names(output_file_path, reference_path=None):
    map_path = _compression_map_path(output_file_path)
    if not os.path.exists(map_path):
        return False
    try:
        with open(map_path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
        payload["coded_filename"] = os.path.basename(output_file_path)
        if reference_path:
            payload["reference_filename"] = os.path.basename(reference_path)
        temp_path = f"{map_path}.{uuid.uuid4().hex}.tmp"
        with open(temp_path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, ensure_ascii=False)
        os.replace(temp_path, map_path)
        return True
    except Exception:
        try:
            if "temp_path" in locals():
                _remove_if_exists(temp_path)
        except Exception:
            pass
        return False


def _remember_compressed_reference_path(output_file_path, reference_path):
    return _update_compression_map_file_names(output_file_path, reference_path)


def _remember_compressed_output_path(output_file_path):
    return _update_compression_map_file_names(output_file_path)


def _compression_identity_date(value):
    text = str(value or "").strip()
    iso_match = re.match(r"^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$", text)
    if iso_match:
        year, month, day = iso_match.groups()
        return f"{day}.{month}.{year}"
    return _parse_icelandic_date(value)


def _source_groups_for_current_output(output_file_path, df):
    map_path = _compression_map_path(output_file_path)
    if not os.path.exists(map_path):
        return {}
    try:
        with open(map_path, "r", encoding="utf-8") as handle:
            compression_map = json.load(handle)
    except Exception as exc:
        raise CompressionMapMismatchError(
            "Þjöppunarkortið er skemmt eða ólæsilegt. Kóðuðu skránni var ekki breytt."
        ) from exc
    mapped_coded_name = str(compression_map.get("coded_filename", "")).strip()
    if (
        mapped_coded_name
        and os.path.normcase(mapped_coded_name)
        != os.path.normcase(os.path.basename(output_file_path))
    ):
        return {}
    mapped_rows = compression_map.get("rows", [])
    if len(mapped_rows) != len(df):
        raise CompressionMapMismatchError(
            "Þjöppunarkortið passar ekki við línufjölda kóðuðu skrárinnar. "
            "Kóðuðu skránni var ekki breytt."
        )

    # A compression map is stored beside the workbook. Confirm that it still
    # describes this exact workbook before using it; otherwise a different
    # same-length file in the folder could inherit unrelated source rows. TEXT
    # and ID are intentionally editable, so only stable row identity is used.
    for position, (_, row) in enumerate(df.iterrows()):
        entry = mapped_rows[position] if position < len(mapped_rows) else {}
        identity = entry.get("identity", {})
        if not isinstance(identity, dict) or not identity:
            continue
        for field in ("DATE", "AMOUNT", "Positive/Negative"):
            if field not in identity or field not in df.columns:
                continue
            expected = identity.get(field, "")
            actual = row.get(field, "")
            if field == "AMOUNT":
                expected_amount = _parse_bank_amount_decimal(expected)
                actual_amount = _parse_bank_amount_decimal(actual)
                if expected_amount != actual_amount:
                    raise CompressionMapMismatchError(
                        "Upphæðir í þjöppunarkortinu passa ekki við kóðuðu skrána. "
                        "Kóðuðu skránni var ekki breytt."
                    )
            elif field == "DATE":
                expected_date = _compression_identity_date(expected)
                actual_date = _compression_identity_date(actual)
                if expected_date != actual_date:
                    raise CompressionMapMismatchError(
                        "Dagsetningar í þjöppunarkortinu passa ekki við kóðuðu skrána. "
                        "Kóðuðu skránni var ekki breytt."
                    )
            else:
                if str(expected).strip() != str(actual).strip():
                    raise CompressionMapMismatchError(
                        "Formerki í þjöppunarkortinu passa ekki við kóðuðu skrána. "
                        "Kóðuðu skránni var ekki breytt."
                    )

    groups = {}
    for position, idx in enumerate(df.index):
        entry = mapped_rows[position] if position < len(mapped_rows) else {}
        paired_indices = entry.get("paired_indices", [])
        clean_indices = []
        for paired_index in paired_indices:
            try:
                clean_indices.append(int(paired_index))
            except Exception:
                pass
        if clean_indices:
            groups[idx] = clean_indices
    return groups


def _compression_reference_print_view(df):
    if df is None:
        return pd.DataFrame(columns=_COMPRESSION_REFERENCE_PRINT_COLUMNS)
    return df.reindex(columns=_COMPRESSION_REFERENCE_PRINT_COLUMNS, fill_value="").copy()


def _append_reference_rows(reference_path, reference_df):
    reference_print = _compression_reference_print_view(reference_df)
    if not reference_path or not os.path.exists(reference_path):
        return reference_print
    try:
        existing = pd.read_excel(reference_path, dtype=object, keep_default_na=False)
    except Exception:
        return reference_print
    existing_print = _compression_reference_print_view(existing)
    return pd.concat([existing_print, reference_print], ignore_index=True)


def _row_source_group(row_source_groups, row_index):
    source_group = row_source_groups.get(row_index)
    if source_group:
        return [int(index) for index in source_group]
    return [int(row_index)]


def _apply_transaction_compressions(df, selected_clusters, row_source_groups=None):
    row_source_groups = row_source_groups or {}
    selected_by_index = {}
    cluster_by_first = {}
    for cluster in selected_clusters:
        row_indices = list(cluster["row_indices"])
        if not row_indices:
            continue
        first_idx = row_indices[0]
        cluster_by_first[first_idx] = cluster
        for idx in row_indices:
            selected_by_index[idx] = first_idx

    output_rows = []
    paired_index_groups = []
    reference_rows = []
    reference_columns = list(_COMPRESSION_REFERENCE_PRINT_COLUMNS)
    blank_reference_row = {column: "" for column in reference_columns}
    for idx, row in df.iterrows():
        first_idx = selected_by_index.get(idx)
        if first_idx is None:
            output_rows.append(row.to_dict())
            paired_index_groups.append(_row_source_group(row_source_groups, idx))
            continue
        if idx != first_idx:
            continue

        cluster = cluster_by_first[first_idx]
        row_indices = list(cluster["row_indices"])
        cluster_df = df.loc[row_indices].copy()
        summary = cluster_df.iloc[0].copy()
        summary_text = str(cluster.get("summary_text", "")).strip()
        if summary_text and "TEXT" in summary.index:
            summary["TEXT"] = summary_text
        summary["AMOUNT"] = _format_compressed_amount(cluster["total"])
        if "STATUS" in summary.index:
            summary["STATUS"] = _COMPRESSED_TRANSACTION_STATUS
        if "ID" in summary.index:
            id_values = {
                _format_auto_review_value(value, limit=200).strip()
                for value in cluster_df["ID"].tolist()
                if _format_auto_review_value(value, limit=200).strip()
            }
            if len(id_values) > 1:
                summary["ID"] = ""
        output_rows.append(summary.to_dict())
        paired_group = []
        for row_index in row_indices:
            paired_group.extend(_row_source_group(row_source_groups, row_index))
        paired_index_groups.append(paired_group)

        for _, detail_row in cluster_df.iterrows():
            ref_row = {
                column: detail_row.get(column, "")
                for column in reference_columns
            }
            reference_rows.append(ref_row)

        reference_rows.append(dict(blank_reference_row))
        sum_row = {
            column: summary.get(column, "")
            for column in reference_columns
        }
        reference_rows.append(sum_row)
        reference_rows.append(dict(blank_reference_row))

    compressed_df = pd.DataFrame(output_rows, columns=df.columns)
    reference_df = pd.DataFrame(reference_rows, columns=reference_columns)
    return compressed_df, reference_df, paired_index_groups


def _compression_file_signature(file_path):
    try:
        stat = os.stat(file_path)
        return stat.st_mtime_ns, stat.st_size
    except OSError:
        return None


def _read_transaction_compression_workbook(output_file_path):
    """Read a stable, current snapshot of a coded workbook from disk."""
    if not output_file_path or not os.path.exists(output_file_path):
        return None, None, "Úttaksskrá fannst ekki."
    last_error = ""
    for _attempt in range(2):
        before = _compression_file_signature(output_file_path)
        try:
            df = pd.read_excel(output_file_path, dtype=object, keep_default_na=False)
        except Exception as exc:
            last_error = f"Ekki tókst að endurhlaða úttaksskránni: {str(exc)[:70]}"
            continue
        after = _compression_file_signature(output_file_path)
        if before is not None and before == after:
            return df, after, ""
        last_error = "Úttaksskráin breyttist á meðan verið var að endurhlaða henni. Lokaðu Excel og reyndu aftur."
    return None, None, last_error or "Ekki tókst að endurhlaða úttaksskránni."


def _remove_if_exists(path):
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


def _save_transaction_compressions(output_file_path, df, selected, expected_signature=None):
    if expected_signature is not None and _compression_file_signature(output_file_path) != expected_signature:
        return {
            "ok": False,
            "message": "Kóðaða skráin breyttist eftir að hún var hlaðin inn. Lokaðu Excel og veldu Þjappa aftur.",
        }

    reference_path = _compressed_reference_path(output_file_path)
    try:
        row_source_groups = _source_groups_for_current_output(output_file_path, df)
    except CompressionMapMismatchError as exc:
        return {"ok": False, "message": str(exc)}
    compressed_df, reference_df, paired_index_groups = _apply_transaction_compressions(
        df,
        selected,
        row_source_groups=row_source_groups,
    )
    reference_output_df = _append_reference_rows(reference_path, reference_df)

    token = uuid.uuid4().hex
    output_folder = os.path.dirname(os.path.abspath(output_file_path))
    reference_folder = os.path.dirname(os.path.abspath(reference_path))
    output_temp = os.path.join(output_folder, f".{os.path.basename(output_file_path)}.{token}.tmp.xlsx")
    reference_temp = os.path.join(reference_folder, f".{os.path.basename(reference_path)}.{token}.tmp.xlsx")
    output_backup = os.path.join(output_folder, f".{os.path.basename(output_file_path)}.{token}.backup")
    reference_backup = os.path.join(reference_folder, f".{os.path.basename(reference_path)}.{token}.backup")

    try:
        _prepare_bank_excel_output(compressed_df).to_excel(output_temp, index=False)
        _prepare_bank_excel_output(reference_output_df).to_excel(reference_temp, index=False)
        autofit_excel_columns(output_temp)
        autofit_excel_columns(reference_temp)
    except Exception as exc:
        _remove_if_exists(output_temp)
        _remove_if_exists(reference_temp)
        return {"ok": False, "message": f"Ekki tókst að vista þjöppunina: {str(exc)[:75]}"}

    if expected_signature is not None and _compression_file_signature(output_file_path) != expected_signature:
        _remove_if_exists(output_temp)
        _remove_if_exists(reference_temp)
        return {
            "ok": False,
            "message": "Kóðaða skráin breyttist á meðan þjöppunarglugginn var opinn. Veldu Þjappa aftur til að endurhlaða henni.",
        }

    output_moved = False
    output_replaced = False
    reference_moved = False
    reference_replaced = False
    try:
        # Moving the original first safely detects a workbook that is still open
        # in Excel. Nothing is overwritten until both replacement files are ready.
        os.replace(output_file_path, output_backup)
        output_moved = True
        os.replace(output_temp, output_file_path)
        output_replaced = True

        if os.path.exists(reference_path):
            os.replace(reference_path, reference_backup)
            reference_moved = True
        os.replace(reference_temp, reference_path)
        reference_replaced = True
    except Exception as exc:
        if reference_replaced:
            _remove_if_exists(reference_path)
        if reference_moved and os.path.exists(reference_backup):
            try:
                os.replace(reference_backup, reference_path)
            except Exception:
                pass
        if output_replaced:
            _remove_if_exists(output_file_path)
        if output_moved and os.path.exists(output_backup):
            try:
                os.replace(output_backup, output_file_path)
            except Exception:
                pass
        _remove_if_exists(output_temp)
        _remove_if_exists(reference_temp)
        return {
            "ok": False,
            "message": f"Ekki tókst að uppfæra kóðuðu skrána. Lokaðu henni í Excel og reyndu aftur: {str(exc)[:55]}",
        }

    _remove_if_exists(output_backup)
    _remove_if_exists(reference_backup)

    map_result = write_compression_map(
        session_dir=os.path.dirname(os.path.abspath(output_file_path)),
        coded_file_path=output_file_path,
        compressed_df=compressed_df,
        paired_index_groups=paired_index_groups,
    )
    if map_result.get("ok"):
        _remember_compressed_reference_path(output_file_path, reference_path)

    deleted_count = sum(cluster["count"] for cluster in selected)
    map_warning = ""
    if not map_result.get("ok"):
        map_warning = f" Þjöppun tókst, en ekki tókst að vista þjálfunarkortið: {map_result.get('message', 'óþekkt villa')}"
    output_row_text = "eina færslu" if len(selected) == 1 else f"{len(selected)} færslur"
    return {
        "ok": True,
        "message": (
            f"{deleted_count} færslum var þjappað í {output_row_text}. "
            f"Tilvísunarskrá: {os.path.basename(reference_path)}"
            f"{map_warning}"
        ),
        "reference_path": reference_path,
        "clusters": len(selected),
        "deleted_count": deleted_count,
    }


def compress_transactions_interactive(output_file_path):
    df, source_signature, read_error = _read_transaction_compression_workbook(output_file_path)
    if read_error:
        return {"ok": False, "message": read_error}

    try:
        source_groups = _source_groups_for_current_output(output_file_path, df)
    except CompressionMapMismatchError as exc:
        return {"ok": False, "message": str(exc)}
    already_compressed = {
        idx
        for idx, paired_indices in source_groups.items()
        if len(paired_indices) > 1
    }
    clusters = _find_compressible_transaction_clusters(
        df,
        minimum_count=5,
        excluded_indices=already_compressed,
    )
    if not clusters:
        return {
            "ok": False,
            "message": "Engir flokkar með að minnsta kosti 5 óþjöppuðum færslum frá sama mótaðila fundust.",
        }

    selected = []
    total = len(clusters)
    for index, cluster in enumerate(clusters, start=1):
        cluster_df = df.loc[cluster["row_indices"]].copy()
        choice = _show_transaction_compression_review(cluster_df, cluster, index, total)
        if choice.get("action") == "skip_all":
            break
        if choice.get("action") == "compress":
            selected_cluster = dict(cluster)
            if choice.get("row_indices"):
                selected_cluster["row_indices"] = list(choice["row_indices"])
            if choice.get("count"):
                selected_cluster["count"] = int(choice["count"])
            if choice.get("total") is not None:
                selected_cluster["total"] = choice["total"]
            for field in ("summary_text", "sign", "debit_code", "credit_code"):
                if choice.get(field) is not None:
                    selected_cluster[field] = choice[field]
            selected.append(selected_cluster)

    if not selected:
        return {"ok": False, "message": "Engum færsluflokkum var þjappað."}

    return _save_transaction_compressions(
        output_file_path,
        df,
        selected,
        expected_signature=source_signature,
    )


def manual_compress_transactions_interactive(output_file_path):
    df, source_signature, read_error = _read_transaction_compression_workbook(output_file_path)
    if read_error:
        return {"ok": False, "message": read_error}

    try:
        _source_groups_for_current_output(output_file_path, df)
    except CompressionMapMismatchError as exc:
        return {"ok": False, "message": str(exc)}

    choice = _show_manual_transaction_compression_dialog(df)
    if choice.get("action") == "error":
        return {"ok": False, "message": choice.get("message", "Handvirk þjöppun mistókst.")}
    if choice.get("action") != "compress":
        return {"ok": False, "message": "Engin handvirk þjöppun var valin."}

    selected_cluster = {
        "row_indices": list(choice.get("row_indices", [])),
        "count": int(choice.get("count", 0) or 0),
        "total": choice.get("total"),
        "summary_text": choice.get("summary_text", ""),
        "display_name": choice.get("display_name", ""),
        "sign": choice.get("sign", ""),
        "debit_code": choice.get("debit_code", ""),
        "credit_code": choice.get("credit_code", ""),
        "first_position": int(choice.get("first_position", 0) or 0),
    }
    if len(selected_cluster["row_indices"]) < 2:
        return {"ok": False, "message": "Veldu að minnsta kosti tvær færslur."}

    return _save_transaction_compressions(
        output_file_path,
        df,
        [selected_cluster],
        expected_signature=source_signature,
    )


def _find_bank_column(df, candidates):
    normalized_candidates = {
        _strip_accents(str(candidate)).lower().replace(" ", "")
        for candidate in candidates
    }
    for column in df.columns:
        normalized = _strip_accents(str(column)).lower().replace(" ", "")
        if normalized in normalized_candidates:
            return column
    return None


def _arion_amount_source(df, fallback_amount_col):
    currency_col = _find_bank_column(df, ["Mynt"])
    isk_amount_col = _find_bank_column(df, ["Upphæð í ISK", "Upphæð ISK", "Upphæð(ISK)"])
    if currency_col is None or isk_amount_col is None:
        return df[fallback_amount_col]

    fallback_amounts = df[fallback_amount_col]
    isk_amounts = df[isk_amount_col]
    use_isk = df[currency_col].fillna("").astype(str).str.strip().str.upper().ne("ISK")
    use_isk = use_isk & isk_amounts.notna() & isk_amounts.astype(str).str.strip().ne("")
    return fallback_amounts.where(~use_isk, isk_amounts)


def run_bank_formatter_script(autofill_7810=True):
    input_file_path = bank_input_file_entry.get()
    output_name = bank_output_name_entry.get().strip() if "bank_output_name_entry" in globals() else ""
    auto_code_debits = True
    restaurant_mode = _bank_restaurant_mode_enabled()
    industry_context = "restaurant" if restaurant_mode else ""
    vehicle_expense_code = _bank_vehicle_expense_code()
    fill_counter, counter_code = _bank_counter_fill(autofill_7810)

    if not input_file_path:
        set_bank_formatter_status("Veldu inntaksskrá", TEXT_WARNING)
        return
    if not output_name:
        set_bank_formatter_status("Sláðu inn skráarheiti", TEXT_WARNING)
        try:
            bank_output_name_entry.focus_set()
        except Exception:
            pass
        return
    if not _coded_output_filename(output_name):
        set_bank_formatter_status("Sláðu inn gilt skráarheiti", TEXT_WARNING)
        try:
            bank_output_name_entry.focus_set()
        except Exception:
            pass
        return
    
    bank_type, df = detect_bank_type(input_file_path)
    
    if bank_type is None:
        set_bank_formatter_status("Ekki tókst að greina bankasnið", TEXT_ERROR, show_open=True)
        play_error_sound()
        return

    config = BANK_CONFIGS[bank_type]
    client_hint = _bank_client_hint(input_file_path, output_name)
    output_file_path = _create_training_output_path(input_file_path, bank_type, client_hint, output_name)
    
    try:
        # Check if this bank needs custom processing
        if config.get('custom_processor'):
            if bank_type == 'islandsbanki_innheimta':
                process_islandsbanki_innheimta(
                    df,
                    output_file_path,
                    config['name'],
                    autofill_7810=fill_counter,
                    input_file_path=input_file_path,
                    bank_type=bank_type,
                    client_hint=client_hint,
                )
            elif bank_type == 'islandsbanki_kort':
                process_islandsbanki_kort(
                    df,
                    output_file_path,
                    config['name'],
                    autofill_7810=fill_counter,
                    input_file_path=input_file_path,
                    auto_code_debits=auto_code_debits,
                    industry_context=industry_context,
                    vehicle_expense_code=vehicle_expense_code,
                    counter_code=counter_code,
                    bank_type=bank_type,
                    client_hint=client_hint,
                )
            elif bank_type == 'islandsbanki_special':
                process_islandsbanki_special(df, output_file_path, config['name'], autofill_7810=fill_counter, counter_code=counter_code, input_file_path=input_file_path, bank_type=bank_type, client_hint=client_hint)
            elif bank_type == 'sala_yfirlit':
                process_sala_yfirlit(df, output_file_path, config['name'], autofill_7810=fill_counter, input_file_path=input_file_path, bank_type=bank_type, client_hint=client_hint)
            elif bank_type == 'landsbankinn_kort':
                process_landsbankinn_kort(
                    df,
                    output_file_path,
                    config['name'],
                    autofill_7810=fill_counter,
                    input_file_path=input_file_path,
                    auto_code_debits=auto_code_debits,
                    industry_context=industry_context,
                    vehicle_expense_code=vehicle_expense_code,
                    counter_code=counter_code,
                    bank_type=bank_type,
                    client_hint=client_hint,
                )
            elif bank_type == 'arion_kort':
                process_arion_kort(
                    df,
                    output_file_path,
                    config['name'],
                    autofill_7810=fill_counter,
                    input_file_path=input_file_path,
                    auto_code_debits=auto_code_debits,
                    industry_context=industry_context,
                    vehicle_expense_code=vehicle_expense_code,
                    counter_code=counter_code,
                    bank_type=bank_type,
                    client_hint=client_hint,
                )
        else:
            # Standard processing for regular banks
            extracted_columns = _with_training_raw_index(df[config['columns']])
            amount_col = config['columns'][3]
            amount_source = (
                _arion_amount_source(df, amount_col)
                if bank_type in {'arion', 'arion_reikningsyfirlit'}
                else extracted_columns[amount_col]
            )
            extracted_columns['Positive/Negative'] = amount_source.apply(_bank_amount_sign)
            extracted_columns[amount_col] = amount_source.apply(_bank_amount_abs)
            
            # Rename to standard format
            extracted_columns.rename(columns=config['rename'], inplace=True)
            
            # Insert DEBIT and CREDIT columns
            extracted_columns.insert(2, 'DEBIT', '')
            extracted_columns.insert(5, 'CREDIT', '')

            if fill_counter:
                extracted_columns.loc[extracted_columns['Positive/Negative'] == '+', 'DEBIT'] = counter_code
                extracted_columns.loc[extracted_columns['Positive/Negative'] == '-', 'CREDIT'] = counter_code

            extracted_columns = _apply_bank_auto_coding(
                extracted_columns,
                df,
                input_file_path,
                enabled=auto_code_debits,
                skip=config.get('skip_auto_coding', False),
                note=config.get('auto_coding_note', ''),
                industry_context=industry_context,
                vehicle_expense_code=vehicle_expense_code,
                bank_type=bank_type,
            )
            extracted_columns = review_low_confidence_auto_code_clusters(
                extracted_columns,
                enabled=auto_code_debits and not config.get('skip_auto_coding', False),
            )

            # Sort by Positive/Negative and TEXT
            extracted_columns = extracted_columns.sort_values(by=['Positive/Negative', 'TEXT'])
            
            # Convert DATE to text to preserve format on copy-paste
            extracted_columns['DATE'] = extracted_columns['DATE'].apply(format_date_as_text)
            
            # Save to Excel
            _prepare_bank_excel_output(extracted_columns).to_excel(output_file_path, index=False)
            
            # Auto-fit column widths
            autofit_excel_columns(output_file_path)

            archive_path = _archive_training_import(
                input_file_path,
                df,
                extracted_columns,
                output_file_path,
                bank_type,
                config['name'],
                client_hint,
            )

            display_bank_formatter_success(
                config['name'],
                output_file_path,
                training_archive_path=archive_path,
                raw_file_path=input_file_path,
            )
        
    except Exception as e:
        set_bank_formatter_status(f"✗ Villa: {str(e)[:40]}", TEXT_ERROR)
        play_error_sound()


def process_islandsbanki_innheimta(df, output_file_path, bank_name, autofill_7810=True, input_file_path="", bank_type="islandsbanki_innheimta", client_hint=""):
    # Do not run merchant/purpose auto-coding on this format; it already creates fixed calculation rows.
    columns_to_select = ['Kennitala', 'Greiðsludagur', 'Fjármagnstekjuskattur', 'Dráttarvextir', 'Rst. Upphæð', 'Upphæð']
    df_selected = _with_training_raw_index(df[columns_to_select])
    df_selected = df_selected[
        df_selected['Kennitala'].apply(lambda value: bool(_format_kennitala(value)))
        & df_selected['Greiðsludagur'].apply(lambda value: bool(format_date_as_text(value)))
    ]

    entries = []
    for raw_idx, row in df_selected.iterrows():
        date_val = format_date_as_text(row['Greiðsludagur'])
        id_val = _format_kennitala(row['Kennitala'])
        text_val = 'kostnaður'
        fjarmagn = _parse_innheimta_amount(row['Fjármagnstekjuskattur'])
        drattar = _parse_innheimta_amount(row['Dráttarvextir'])
        rst_amount = _parse_innheimta_amount(row['Rst. Upphæð'], force_thousands=True)
        upphaed = _parse_innheimta_amount(row['Upphæð'], force_thousands=True)

        if fjarmagn != 0:
            entries.append({
                'DATE': date_val,
                'TEXT': text_val,
                'DEBIT': 7660,
                'AMOUNT': fjarmagn,
                'ID': id_val,
                'CREDIT': 7620,
                TRAINING_RAW_INDEX_COLUMN: raw_idx
            })
        if drattar != 0:
            entries.append({
                'DATE': date_val,
                'TEXT': text_val,
                'DEBIT': 7620,
                'AMOUNT': drattar,
                'ID': id_val,
                'CREDIT': 6100,
                TRAINING_RAW_INDEX_COLUMN: raw_idx
            })
        # Rst. Upphæð + Fjármagnstekjuskattur - Dráttarvextir - Upphæð (J + F - G - E).
        calc_amount = rst_amount + fjarmagn - drattar - upphaed
        if calc_amount != 0:
            entries.append({
                'DATE': date_val,
                'TEXT': text_val,
                'DEBIT': 7620,
                'AMOUNT': calc_amount,
                'ID': id_val,
                'CREDIT': 4400,
                TRAINING_RAW_INDEX_COLUMN: raw_idx
            })

    df_output = pd.DataFrame(entries, columns=['DATE', 'TEXT', 'DEBIT', 'AMOUNT', 'ID', 'CREDIT', TRAINING_RAW_INDEX_COLUMN])
    _prepare_bank_excel_output(df_output).to_excel(output_file_path, index=False)
    autofit_excel_columns(output_file_path)
    archive_path = _archive_training_import(input_file_path, df, df_output, output_file_path, bank_type, bank_name, client_hint)
    display_bank_formatter_success(
        bank_name,
        output_file_path,
        training_archive_path=archive_path,
        raw_file_path=input_file_path,
    )


def process_sala_yfirlit(df, output_file_path, bank_name, autofill_7810=True, input_file_path="", bank_type="sala_yfirlit", client_hint=""):
    # Do not run merchant/purpose auto-coding on this format; Reikningur nr is an internal incremental invoice number.
    # Drop total rows if present (where Kennitala is NaN and/or Nafn starts with 'Alls')
    df_clean = df.copy()
    df_clean = df_clean[~df_clean['Nafn'].astype(str).str.lower().str.startswith('alls')]
    cols = ['Nafn', 'Kennitala', 'Upphæð með vsk', 'Reikningur nr', 'Dagsetning']
    df_selected = _with_training_raw_index(df_clean[cols])

    # ID: strip dash and spaces
    df_selected['ID'] = df_selected['Kennitala'].astype(str).str.replace('-', '').str.replace(' ', '')

    # TEXT: "Reikningur X" (strip trailing .0)
    def _fmt_reikningur(x):
        try:
            as_float = float(x)
            if as_float.is_integer():
                return f"Reikningur {int(as_float)}"
        except Exception:
            pass
        return f"Reikningur {str(x).rstrip('0').rstrip('.') if isinstance(x, str) else x}"
    df_selected['TEXT'] = df_selected['Reikningur nr'].apply(_fmt_reikningur)

    # AMOUNT: round half up from Upphæð með vsk
    def _round_amount(v):
        dec = _round_half_up_decimal(v)
        if dec is None:
            return v
        return int(dec.quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    df_selected['AMOUNT'] = df_selected['Upphæð með vsk'].apply(_round_amount)

    # DATE: format as text
    df_selected['DATE'] = df_selected['Dagsetning'].apply(_parse_icelandic_date)

    # Positive/Negative and absolute amount
    df_selected['Positive/Negative'] = df_selected['AMOUNT'].apply(lambda x: '+' if _round_half_up_decimal(x) is None or _round_half_up_decimal(x) >= 0 else '-')
    df_selected['AMOUNT'] = df_selected['AMOUNT'].apply(lambda x: abs(int(_round_half_up_decimal(x).quantize(Decimal("1"), rounding=ROUND_HALF_UP))) if _round_half_up_decimal(x) is not None else x)

    # Insert DEBIT and CREDIT
    df_selected.insert(2, 'DEBIT', '')
    df_selected.insert(5, 'CREDIT', '')

    if autofill_7810:
        df_selected['DEBIT'] = 7620
        df_selected['CREDIT'] = 1000

    df_selected = df_selected.sort_values(by=['Positive/Negative', 'TEXT'])

    df_final = df_selected[['DATE', 'TEXT', 'DEBIT', 'AMOUNT', 'ID', 'CREDIT', TRAINING_RAW_INDEX_COLUMN]].copy()
    _prepare_bank_excel_output(df_final).to_excel(output_file_path, index=False)
    autofit_excel_columns(output_file_path)
    archive_path = _archive_training_import(input_file_path, df, df_final, output_file_path, bank_type, bank_name, client_hint)
    display_bank_formatter_success(
        bank_name,
        output_file_path,
        training_archive_path=archive_path,
        raw_file_path=input_file_path,
    )

def process_islandsbanki_special(df, output_file_path, bank_name, autofill_7810=True, counter_code=7810, input_file_path="", bank_type="islandsbanki_special", client_hint=""):
    # Do not run merchant/purpose auto-coding on this format; it is a special calculation/subtraction import.
    config = BANK_CONFIGS['islandsbanki_special']
    df_selected = _with_training_raw_index(df[config['columns']])
    df_selected.rename(columns=config['rename'], inplace=True)
    df_selected = df_selected.dropna(how='all')
    df_selected = df_selected[df_selected['DATE'].notna() & df_selected['AMOUNT'].notna()]
    df_selected = df_selected[df_selected['AMOUNT'].apply(_parse_bank_amount_decimal).notna()]

    df_selected['DATE'] = df_selected['DATE'].apply(format_date_as_text)
    df_selected['TEXT'] = df_selected['TEXT'].fillna('').astype(str).str.strip()
    df_selected['Positive/Negative'] = df_selected['AMOUNT'].apply(_bank_amount_sign)
    df_selected['AMOUNT'] = df_selected['AMOUNT'].apply(_bank_amount_abs)
    df_selected['DEBIT'] = ''
    df_selected['ID'] = ''
    df_selected['CREDIT'] = ''

    if autofill_7810:
        df_selected.loc[df_selected['Positive/Negative'] == '+', 'DEBIT'] = counter_code
        df_selected.loc[df_selected['Positive/Negative'] == '-', 'CREDIT'] = counter_code

    df_selected = df_selected.sort_values(by=['Positive/Negative', 'TEXT'])
    df_output = df_selected[['DATE', 'TEXT', 'DEBIT', 'ID', 'AMOUNT', 'CREDIT', 'Positive/Negative', TRAINING_RAW_INDEX_COLUMN]].copy()
    _prepare_bank_excel_output(df_output).to_excel(output_file_path, index=False)
    autofit_excel_columns(output_file_path)
    archive_path = _archive_training_import(input_file_path, df, df_output, output_file_path, bank_type, bank_name, client_hint)
    display_bank_formatter_success(
        bank_name,
        output_file_path,
        training_archive_path=archive_path,
        raw_file_path=input_file_path,
    )

def process_islandsbanki_kort(
    df,
    output_file_path,
    bank_name,
    autofill_7810=True,
    input_file_path="",
    auto_code_debits=True,
    industry_context="",
    counter_code=7810,
    bank_type="islandsbanki_kort",
    client_hint="",
    vehicle_expense_code="",
):
    config = BANK_CONFIGS['islandsbanki_kort']
    df_selected = _with_training_raw_index(df[config['columns']])
    df_selected.rename(columns=config['rename'], inplace=True)
    df_selected = df_selected.dropna(how='all')
    df_selected = df_selected[df_selected['DATE'].notna() & df_selected['AMOUNT'].notna()]
    df_selected = df_selected[df_selected['AMOUNT'].apply(_parse_bank_amount_decimal).notna()]

    df_selected['DATE'] = df_selected['DATE'].apply(format_date_as_text)
    df_selected['TEXT'] = df_selected['TEXT'].fillna('').astype(str).str.strip()
    amount_values = df_selected['AMOUNT'].apply(_parse_bank_amount_decimal)
    # Credit-card charges are exported as positive amounts, but they are outgoing purchases.
    df_selected['Positive/Negative'] = amount_values.apply(lambda amount: '+' if amount < 0 else '-')
    df_selected['AMOUNT'] = amount_values.apply(lambda amount: _bank_amount_abs(amount))
    df_selected['DEBIT'] = ''
    df_selected['ID'] = ''
    df_selected['CREDIT'] = ''

    if autofill_7810:
        df_selected.loc[df_selected['Positive/Negative'] == '+', 'DEBIT'] = counter_code
        df_selected.loc[df_selected['Positive/Negative'] == '-', 'CREDIT'] = counter_code

    df_selected = df_selected.sort_values(by=['Positive/Negative', 'TEXT'])
    df_output = df_selected[['DATE', 'TEXT', 'DEBIT', 'ID', 'AMOUNT', 'CREDIT', 'Positive/Negative', TRAINING_RAW_INDEX_COLUMN]].copy()
    df_output = _apply_bank_auto_coding(
        df_output,
        df,
        input_file_path,
        enabled=auto_code_debits,
        industry_context=industry_context,
        vehicle_expense_code=vehicle_expense_code,
        bank_type=bank_type,
    )
    df_output = review_low_confidence_auto_code_clusters(df_output, enabled=auto_code_debits)
    _prepare_bank_excel_output(df_output).to_excel(output_file_path, index=False)
    autofit_excel_columns(output_file_path)
    archive_path = _archive_training_import(input_file_path, df, df_output, output_file_path, bank_type, bank_name, client_hint)
    display_bank_formatter_success(
        bank_name,
        output_file_path,
        training_archive_path=archive_path,
        raw_file_path=input_file_path,
    )

def process_landsbankinn_kort(
    df,
    output_file_path,
    bank_name,
    autofill_7810=True,
    input_file_path="",
    auto_code_debits=True,
    industry_context="",
    counter_code=7810,
    bank_type="landsbankinn_kort",
    client_hint="",
    vehicle_expense_code="",
):
    config = BANK_CONFIGS['landsbankinn_kort']
    df_selected = _with_training_raw_index(df[config['columns']])
    df_selected.rename(columns=config['rename'], inplace=True)
    df_selected = df_selected.dropna(how='all')
    df_selected = df_selected[df_selected['DATE'].notna() & df_selected['AMOUNT'].notna()]

    df_selected['DATE'] = df_selected['DATE'].apply(format_date_as_text)
    df_selected['TEXT'] = df_selected['TEXT'].fillna('').astype(str).str.strip()
    df_selected['Positive/Negative'] = df_selected['AMOUNT'].apply(_bank_amount_sign)
    df_selected['AMOUNT'] = df_selected['AMOUNT'].apply(_bank_amount_abs)
    df_selected['DEBIT'] = ''
    df_selected['ID'] = ''
    df_selected['CREDIT'] = ''

    if autofill_7810:
        df_selected.loc[df_selected['Positive/Negative'] == '+', 'DEBIT'] = counter_code
        df_selected.loc[df_selected['Positive/Negative'] == '-', 'CREDIT'] = counter_code

    df_selected = df_selected.sort_values(by=['Positive/Negative', 'TEXT'])
    df_output = df_selected[['DATE', 'TEXT', 'DEBIT', 'ID', 'AMOUNT', 'CREDIT', 'Positive/Negative', TRAINING_RAW_INDEX_COLUMN]].copy()
    df_output = _apply_bank_auto_coding(
        df_output,
        df,
        input_file_path,
        enabled=auto_code_debits,
        industry_context=industry_context,
        vehicle_expense_code=vehicle_expense_code,
        bank_type=bank_type,
    )
    df_output = review_low_confidence_auto_code_clusters(df_output, enabled=auto_code_debits)
    _prepare_bank_excel_output(df_output).to_excel(output_file_path, index=False)
    autofit_excel_columns(output_file_path)
    archive_path = _archive_training_import(input_file_path, df, df_output, output_file_path, bank_type, bank_name, client_hint)
    display_bank_formatter_success(
        bank_name,
        output_file_path,
        training_archive_path=archive_path,
        raw_file_path=input_file_path,
    )

def process_arion_kort(
    df,
    output_file_path,
    bank_name,
    autofill_7810=True,
    input_file_path="",
    auto_code_debits=True,
    industry_context="",
    counter_code=7810,
    bank_type="arion_kort",
    client_hint="",
    vehicle_expense_code="",
):
    config = BANK_CONFIGS['arion_kort']
    df_selected = _with_training_raw_index(df[config['columns']])
    df_selected.rename(columns=config['rename'], inplace=True)
    df_selected = df_selected.dropna(how='all')
    df_selected = df_selected[df_selected['DATE'].notna() & df_selected['AMOUNT'].notna()]

    df_selected['DATE'] = df_selected['DATE'].apply(_parse_icelandic_date)
    df_selected['TEXT'] = df_selected['TEXT'].fillna('').astype(str).str.strip()
    df_selected['Positive/Negative'] = df_selected['AMOUNT'].apply(_bank_amount_sign)
    df_selected['AMOUNT'] = df_selected['AMOUNT'].apply(_bank_amount_abs)
    df_selected['DEBIT'] = ''
    df_selected['ID'] = ''
    df_selected['CREDIT'] = ''

    if autofill_7810:
        df_selected.loc[df_selected['Positive/Negative'] == '+', 'DEBIT'] = counter_code
        df_selected.loc[df_selected['Positive/Negative'] == '-', 'CREDIT'] = counter_code

    df_selected = df_selected.sort_values(by=['Positive/Negative', 'TEXT'])
    df_output = df_selected[['DATE', 'TEXT', 'DEBIT', 'ID', 'AMOUNT', 'CREDIT', 'Positive/Negative', TRAINING_RAW_INDEX_COLUMN]].copy()
    df_output = _apply_bank_auto_coding(
        df_output,
        df,
        input_file_path,
        enabled=auto_code_debits,
        industry_context=industry_context,
        vehicle_expense_code=vehicle_expense_code,
        bank_type=bank_type,
    )
    df_output = review_low_confidence_auto_code_clusters(df_output, enabled=auto_code_debits)
    _prepare_bank_excel_output(df_output).to_excel(output_file_path, index=False)
    autofit_excel_columns(output_file_path)
    archive_path = _archive_training_import(input_file_path, df, df_output, output_file_path, bank_type, bank_name, client_hint)
    display_bank_formatter_success(
        bank_name,
        output_file_path,
        training_archive_path=archive_path,
        raw_file_path=input_file_path,
    )

def display_bank_formatter_success(bank_name, output_file_path, training_archive_path="", raw_file_path=""):
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Tilbúið", TEXT_SUCCESS)
    play_success_sound()

    path_state = {"path": output_file_path}
    raw_file_path = os.path.normpath(os.path.abspath(raw_file_path)) if raw_file_path else ""
    edit_state = {"active": False}
    compression_reference_state = {"path": ""}
    compression_reference_edit_state = {"active": False}
    archive_exit_state = {"active": True}

    def current_output_path():
        return path_state["path"]

    def sync_training_archive():
        if not training_archive_path:
            return True
        result = refresh_paired_rows_from_coded(
            session_dir=training_archive_path,
            coded_file_path=current_output_path(),
        )
        if result.get("ok"):
            show_rename_status(result.get("message", "Þjálfunarfærslur voru uppfærðar."), TEXT_SUCCESS)
            return True
        show_rename_status(result.get("message", "Ekki tókst að uppfæra þjálfunarfærslur."), TEXT_ERROR)
        play_error_sound()
        return False

    def leave_success_page(command):
        if not archive_exit_state["active"]:
            return
        sync_training_archive()
        archive_exit_state["active"] = False
        root.protocol("WM_DELETE_WINDOW", root.destroy)
        command()

    def close_app_from_success():
        if not archive_exit_state["active"]:
            root.destroy()
            return
        sync_training_archive()
        archive_exit_state["active"] = False
        root.destroy()

    def open_tok_input_after_sync():
        leave_success_page(lambda: open_tok_input_with_file(current_output_path()))

    def open_raw_file():
        if not raw_file_path or not os.path.exists(raw_file_path):
            show_rename_status("Upprunalega inntaksskráin fannst ekki.", TEXT_WARNING)
            play_error_sound()
            return
        try:
            os.startfile(raw_file_path)
        except Exception as exc:
            show_rename_status(f"Ekki tókst að opna upprunalegu inntaksskrána: {str(exc)[:55]}", TEXT_ERROR)
            play_error_sound()

    class MiniPillButton(tk.Canvas):
        def __init__(self, parent, text, command, width=64, accent=False):
            super().__init__(
                parent,
                width=width,
                height=30,
                bg=BUTTON_MUTED,
                highlightthickness=0,
                bd=0,
                cursor='hand2',
            )
            self.text = text
            self.command = command
            self.accent = accent
            self._hover = False
            self.bind("<Configure>", lambda _event: self._draw())
            self.bind("<Enter>", self._on_enter)
            self.bind("<Leave>", self._on_leave)
            self.bind("<ButtonRelease-1>", self._on_click)
            self._draw()

        def _draw(self):
            self.delete("all")
            w = max(4, self.winfo_width())
            h = max(4, self.winfo_height())
            if self.accent:
                fill = ACCENT_HOVER if self._hover else ACCENT
                text_fill = TEXT_ON_ACCENT
                outline = fill
            else:
                fill = BUTTON_MUTED_HOVER if self._hover else BUTTON_MUTED
                text_fill = TEXT_PRIMARY
                outline = BORDER
            _rounded_rect(self, 1, 1, w - 1, h - 1, 10, fill=fill, outline=outline, width=1)
            self.create_text(w / 2, h / 2, text=self.text, fill=text_fill,
                             font=('Segoe UI Semibold', 9))

        def _on_enter(self, _event):
            self._hover = True
            self._draw()

        def _on_leave(self, _event):
            self._hover = False
            self._draw()

        def _on_click(self, _event):
            self.command()

        def set_text(self, text):
            self.text = text
            self._draw()

        def set_accent(self, accent):
            self.accent = accent
            self._draw()

    def _fit_status_canvas(canvas, status_item):
        """Grow a filename card when a wrapped status needs extra lines."""
        def resize():
            try:
                width = max(120, canvas.winfo_width() - 30)
                canvas.itemconfigure(status_item, width=width)
                bbox = canvas.bbox(status_item)
                target_height = max(92, (bbox[3] + 12) if bbox else 92)
                current_height = int(float(canvas.cget("height")))
                if current_height != target_height:
                    canvas.configure(height=target_height)
                schedule_layout_refresh()
                ensure_window_fits()
            except Exception:
                pass

        try:
            canvas.after_idle(resize)
        except Exception:
            pass

    def show_rename_status(message, color=TEXT_SECONDARY):
        try:
            filename_canvas.itemconfigure(status_text, text=message, fill=color)
            _fit_status_canvas(filename_canvas, status_text)
        except Exception:
            pass

    def fit_file_name_for_display(name, width_px):
        max_chars = max(18, int(width_px / 7))
        if len(name) <= max_chars:
            return name
        suffix_len = min(14, max_chars // 2)
        prefix_len = max(4, max_chars - suffix_len - 3)
        return f"{name[:prefix_len]}...{name[-suffix_len:]}"

    def rename_output_file():
        old_path = current_output_path()
        if not old_path:
            return

        new_name = output_name_var.get().strip().strip('"')
        if not new_name:
            show_rename_status("Sláðu inn skráarheiti.", TEXT_WARNING)
            return

        new_name = os.path.basename(new_name)
        if not new_name.lower().endswith(".xlsx"):
            new_name += ".xlsx"
            output_name_var.set(new_name)

        if re.search(r'[<>:"/\\|?*]', new_name):
            show_rename_status("Skráarheitið inniheldur ógilda stafi.", TEXT_ERROR)
            return

        new_path = os.path.join(os.path.dirname(old_path), new_name)
        old_abs = os.path.normcase(os.path.abspath(old_path))
        new_abs = os.path.normcase(os.path.abspath(new_path))
        if old_abs == new_abs:
            show_rename_status("Skráin heitir þetta nú þegar.", TEXT_SECONDARY)
            return
        if os.path.exists(new_path):
            show_rename_status("Skrá með þessu heiti er þegar til.", TEXT_ERROR)
            return

        try:
            os.rename(old_path, new_path)
            path_state["path"] = new_path
            _remember_compressed_output_path(new_path)
            output_name_var.set(os.path.basename(new_path))
            refresh_filename_display()
            end_filename_edit()
            show_rename_status("Skrá endurnefnd.", TEXT_SUCCESS)
        except Exception as exc:
            show_rename_status(f"Ekki tókst að endurnefna skrána: {str(exc)[:45]}", TEXT_ERROR)

    def start_filename_edit():
        edit_state["active"] = True
        output_name_entry.delete(0, tk.END)
        output_name_entry.insert(0, output_name_var.get())
        filename_canvas.itemconfigure(filename_text, state="hidden")
        filename_canvas.itemconfigure(entry_window, state="normal")
        filename_action_btn.set_text("Vista")
        filename_action_btn.set_accent(True)
        show_rename_status("", TEXT_SECONDARY)
        output_name_entry.focus_set()
        output_name_entry.select_range(0, tk.END)

    def end_filename_edit():
        edit_state["active"] = False
        filename_canvas.itemconfigure(entry_window, state="hidden")
        filename_canvas.itemconfigure(filename_text, state="normal")
        filename_action_btn.set_text("Breyta")
        filename_action_btn.set_accent(False)

    def toggle_filename_edit():
        if edit_state["active"]:
            rename_output_file()
        else:
            start_filename_edit()

    def compress_output_transactions():
        if not sync_training_archive():
            return
        show_rename_status("Leita að færsluflokkum sem hægt er að þjappa...", TEXT_SECONDARY)
        result = compress_transactions_interactive(current_output_path())
        if result.get("ok"):
            reference_path = result.get("reference_path", "")
            if reference_path:
                show_compressed_reference_box(reference_path)
            show_rename_status(result.get("message", "Færslum var þjappað."), TEXT_SUCCESS)
            play_success_sound()
        else:
            show_rename_status(result.get("message", "Engum færslum var þjappað."), TEXT_WARNING)

    def manual_compress_output_transactions():
        if not sync_training_archive():
            return
        show_rename_status("Opna handvirka þjöppun færslna...", TEXT_SECONDARY)
        result = manual_compress_transactions_interactive(current_output_path())
        if result.get("ok"):
            reference_path = result.get("reference_path", "")
            if reference_path:
                show_compressed_reference_box(reference_path)
            show_rename_status(result.get("message", "Færslum var þjappað."), TEXT_SUCCESS)
            play_success_sound()
        else:
            show_rename_status(result.get("message", "Engum færslum var þjappað."), TEXT_WARNING)

    def open_compressed_reference_file():
        reference_path = compression_reference_state.get("path", "")
        if reference_path and os.path.exists(reference_path):
            os.startfile(reference_path)
        else:
            show_compressed_reference_status("Tilvísunarskrá þjappaðra færslna fannst ekki.", TEXT_WARNING)

    def open_compressed_reference_location():
        reference_path = compression_reference_state.get("path", "")
        if reference_path:
            open_file_location(reference_path)
        else:
            show_compressed_reference_status("Tilvísunarskrá þjappaðra færslna fannst ekki.", TEXT_WARNING)

    def show_compressed_reference_status(message, color=TEXT_SECONDARY):
        try:
            compressed_canvas.itemconfigure(compressed_status_text, text=message, fill=color)
            _fit_status_canvas(compressed_canvas, compressed_status_text)
        except Exception:
            show_rename_status(message, color)

    def refresh_compressed_filename_display():
        try:
            w = max(4, compressed_canvas.winfo_width())
            available_width = max(120, w - 105)
            compressed_canvas.itemconfigure(
                compressed_filename_text,
                text=fit_file_name_for_display(compressed_name_var.get(), available_width),
                width=0,
            )
        except Exception:
            pass

    def end_compressed_filename_edit():
        compression_reference_edit_state["active"] = False
        compressed_canvas.itemconfigure(compressed_entry_window, state="hidden")
        compressed_canvas.itemconfigure(compressed_filename_text, state="normal")
        compressed_action_btn.set_text("Breyta")
        compressed_action_btn.set_accent(False)

    def rename_compressed_reference_file():
        old_path = compression_reference_state.get("path", "")
        if not old_path:
            show_compressed_reference_status("Engin þjöppuð skrá er tilbúin enn.", TEXT_WARNING)
            return

        new_name = compressed_name_var.get().strip().strip('"')
        if not new_name:
            show_compressed_reference_status("Sláðu inn skráarheiti.", TEXT_WARNING)
            return
        new_name = os.path.basename(new_name)
        if not new_name.lower().endswith(".xlsx"):
            new_name += ".xlsx"
            compressed_name_var.set(new_name)
        if re.search(r'[<>:"/\\|?*]', new_name):
            show_compressed_reference_status("Skráarheitið inniheldur ógilda stafi.", TEXT_ERROR)
            return

        new_path = os.path.join(os.path.dirname(old_path), new_name)
        old_abs = os.path.normcase(os.path.abspath(old_path))
        new_abs = os.path.normcase(os.path.abspath(new_path))
        if old_abs == new_abs:
            show_compressed_reference_status("Skráin heitir þetta nú þegar.", TEXT_SECONDARY)
            return
        if os.path.exists(new_path):
            show_compressed_reference_status("Skrá með þessu heiti er þegar til.", TEXT_ERROR)
            return
        try:
            os.rename(old_path, new_path)
            compression_reference_state["path"] = new_path
            _remember_compressed_reference_path(current_output_path(), new_path)
            compressed_name_var.set(os.path.basename(new_path))
            refresh_compressed_filename_display()
            end_compressed_filename_edit()
            show_compressed_reference_status("Skrá endurnefnd.", TEXT_SUCCESS)
        except Exception as exc:
            show_compressed_reference_status(f"Ekki tókst að endurnefna skrána: {str(exc)[:45]}", TEXT_ERROR)

    def start_compressed_filename_edit():
        compression_reference_edit_state["active"] = True
        compressed_name_entry.delete(0, tk.END)
        compressed_name_entry.insert(0, compressed_name_var.get())
        compressed_canvas.itemconfigure(compressed_filename_text, state="hidden")
        compressed_canvas.itemconfigure(compressed_entry_window, state="normal")
        compressed_action_btn.set_text("Vista")
        compressed_action_btn.set_accent(True)
        show_compressed_reference_status("", TEXT_SECONDARY)
        compressed_name_entry.focus_set()
        compressed_name_entry.select_range(0, tk.END)

    def toggle_compressed_filename_edit():
        if compression_reference_edit_state["active"]:
            rename_compressed_reference_file()
        else:
            start_compressed_filename_edit()

    def show_compressed_reference_box(reference_path):
        compression_reference_state["path"] = reference_path
        compressed_name_var.set(os.path.basename(reference_path))
        refresh_compressed_filename_display()
        show_compressed_reference_status("Tilvísunarskráin er tilbúin.", TEXT_SUCCESS)
        try:
            if not compressed_canvas.winfo_manager():
                compressed_canvas.pack(fill=tk.X, pady=(10, 0))
            if not compressed_file_row.winfo_manager():
                compressed_file_row.pack(fill=tk.X, pady=(0, 8), after=manual_compress_btn)
            schedule_layout_refresh()
            ensure_window_fits()
        except Exception:
            pass

    def draw_processed_box(event=None):
        processed_canvas.delete("all")
        w = max(4, processed_canvas.winfo_width())
        _rounded_rect(processed_canvas, 1, 1, w - 1, 72, 12,
                      fill=BUTTON_MUTED, outline=BORDER, width=1)
        processed_canvas.create_text(15, 19, text="Unnið sem", anchor="w",
                                     fill=TEXT_SECONDARY, font=('Segoe UI', 9))
        processed_canvas.create_text(15, 47, text=bank_name, anchor="w",
                                     fill=TEXT_PRIMARY, font=('Segoe UI', 11, 'bold'))
    
    card_shell, card = create_panel(frame, padx=22, pady=18)
    card_shell.pack(fill=tk.X, padx=42, pady=(8, 16))

    processed_canvas = tk.Canvas(card, height=74, bg=BG_CARD, highlightthickness=0, bd=0)
    processed_canvas.pack(fill=tk.X, pady=(0, 10))
    processed_canvas.bind("<Configure>", draw_processed_box)
    processed_canvas.after_idle(draw_processed_box)

    output_name = os.path.basename(output_file_path) if output_file_path else ""
    if output_name:
        output_name_var = tk.StringVar(value=output_name)
        filename_canvas = tk.Canvas(card, height=92, bg=BG_CARD, highlightthickness=0, bd=0)
        filename_canvas.pack(fill=tk.X)
        output_name_entry = tk.Entry(
            filename_canvas,
            textvariable=output_name_var,
            bg=BUTTON_MUTED,
            fg=TEXT_PRIMARY,
            insertbackground=TEXT_PRIMARY,
            relief='flat',
            bd=0,
            highlightthickness=0,
            font=('Segoe UI', 10),
        )
        output_name_entry.bind("<Return>", lambda _event: rename_output_file())
        output_name_entry.bind("<Escape>", lambda _event: end_filename_edit())

        filename_action_btn = MiniPillButton(filename_canvas, "Breyta", toggle_filename_edit, width=62)
        filename_label = filename_canvas.create_text(
            15, 19, text="Úttaksskrá", anchor="w",
            fill=TEXT_SECONDARY, font=('Segoe UI', 9)
        )
        filename_text = filename_canvas.create_text(
            15, 48, text=output_name, anchor="w",
            fill=TEXT_PRIMARY, font=('Segoe UI', 10)
        )
        status_text = filename_canvas.create_text(
            15, 74, text="", anchor="w",
            fill=TEXT_SECONDARY, font=('Segoe UI', 8)
        )
        entry_window = filename_canvas.create_window(
            15, 48, anchor="w", window=output_name_entry, state="hidden"
        )
        button_window = filename_canvas.create_window(
            0, 46, anchor="e", window=filename_action_btn
        )

        def draw_filename_box(event=None):
            w = max(4, filename_canvas.winfo_width())
            h = max(92, filename_canvas.winfo_height())
            filename_canvas.delete("filename_bg")
            _rounded_rect(filename_canvas, 1, 1, w - 1, h - 2, 12,
                          fill=BUTTON_MUTED, outline=BORDER, width=1, tags="filename_bg")
            filename_canvas.tag_lower("filename_bg")
            filename_canvas.coords(filename_label, 15, 19)
            filename_canvas.coords(filename_text, 15, 48)
            filename_canvas.coords(status_text, 15, 74)
            filename_canvas.itemconfigure(status_text, width=max(120, w - 30))
            filename_canvas.coords(entry_window, 15, 48)
            filename_canvas.itemconfigure(entry_window, width=max(120, w - 105), height=27)
            filename_canvas.coords(button_window, w - 14, 46)
            refresh_filename_display()

        def refresh_filename_display():
            w = max(4, filename_canvas.winfo_width())
            available_width = max(120, w - 105)
            filename_canvas.itemconfigure(
                filename_text,
                text=fit_file_name_for_display(output_name_var.get(), available_width),
                width=0,
            )

        filename_canvas.bind("<Configure>", draw_filename_box)
        filename_canvas.after_idle(draw_filename_box)

        compressed_name_var = tk.StringVar(value="")
        compressed_canvas = tk.Canvas(card, height=92, bg=BG_CARD, highlightthickness=0, bd=0)
        compressed_name_entry = tk.Entry(
            compressed_canvas,
            textvariable=compressed_name_var,
            bg=BUTTON_MUTED,
            fg=TEXT_PRIMARY,
            insertbackground=TEXT_PRIMARY,
            relief='flat',
            bd=0,
            highlightthickness=0,
            font=('Segoe UI', 10),
        )
        compressed_name_entry.bind("<Return>", lambda _event: rename_compressed_reference_file())
        compressed_name_entry.bind("<Escape>", lambda _event: end_compressed_filename_edit())

        compressed_action_btn = MiniPillButton(compressed_canvas, "Breyta", toggle_compressed_filename_edit, width=62)
        compressed_label = compressed_canvas.create_text(
            15, 19, text="Skrá með þjöppuðum færslum", anchor="w",
            fill=TEXT_SECONDARY, font=('Segoe UI', 9)
        )
        compressed_filename_text = compressed_canvas.create_text(
            15, 48, text="", anchor="w",
            fill=TEXT_PRIMARY, font=('Segoe UI', 10)
        )
        compressed_status_text = compressed_canvas.create_text(
            15, 74, text="", anchor="w",
            fill=TEXT_SECONDARY, font=('Segoe UI', 8)
        )
        compressed_entry_window = compressed_canvas.create_window(
            15, 48, anchor="w", window=compressed_name_entry, state="hidden"
        )
        compressed_button_window = compressed_canvas.create_window(
            0, 46, anchor="e", window=compressed_action_btn
        )

        def draw_compressed_filename_box(event=None):
            w = max(4, compressed_canvas.winfo_width())
            h = max(92, compressed_canvas.winfo_height())
            compressed_canvas.delete("compressed_bg")
            _rounded_rect(compressed_canvas, 1, 1, w - 1, h - 2, 12,
                          fill=BUTTON_MUTED, outline=BORDER, width=1, tags="compressed_bg")
            compressed_canvas.tag_lower("compressed_bg")
            compressed_canvas.coords(compressed_label, 15, 19)
            compressed_canvas.coords(compressed_filename_text, 15, 48)
            compressed_canvas.coords(compressed_status_text, 15, 74)
            compressed_canvas.itemconfigure(compressed_status_text, width=max(120, w - 30))
            compressed_canvas.coords(compressed_entry_window, 15, 48)
            compressed_canvas.itemconfigure(compressed_entry_window, width=max(120, w - 105), height=27)
            compressed_canvas.coords(compressed_button_window, w - 14, 46)
            refresh_compressed_filename_display()

        compressed_canvas.bind("<Configure>", draw_compressed_filename_box)
        compressed_canvas.after_idle(draw_compressed_filename_box)

        action_frame = tk.Frame(frame, bg=BG_DARK)
        action_frame.pack(fill=tk.X, padx=42, pady=(4, 10))

        tok_input_btn = create_styled_button(action_frame, "Keyra inn í tok", open_tok_input_after_sync, width=20)
        tok_input_btn.pack(fill=tk.X, pady=(0, 8))

        format_another_btn = create_styled_button(
            action_frame,
            "Sníða aðra skrá",
            lambda: leave_success_page(initialize_bank_formatter),
            width=20,
            accent=False,
        )
        format_another_btn.pack(fill=tk.X, pady=(0, 8))

        auto_compress_btn = create_styled_button(
            action_frame,
            "Þjappa færslum sjálfvirkt",
            compress_output_transactions,
            width=20,
            accent=False,
        )
        auto_compress_btn.pack(fill=tk.X, pady=(0, 8))

        manual_compress_btn = create_styled_button(action_frame, "Þjappa færslum handvirkt", manual_compress_output_transactions, width=20, accent=False)
        manual_compress_btn.pack(fill=tk.X, pady=(0, 8))

        compressed_file_row = tk.Frame(action_frame, bg=BG_DARK)
        compressed_file_row.pack(fill=tk.X, pady=(0, 8))
        open_compressed_btn = create_styled_button(compressed_file_row, "Opna þjappaðar færslur", open_compressed_reference_file, width=18, accent=False)
        open_compressed_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        open_compressed_location_btn = create_styled_button(compressed_file_row, "Opna staðsetningu", open_compressed_reference_location, width=18, accent=False)
        open_compressed_location_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))
        compressed_file_row.pack_forget()

        file_row = tk.Frame(action_frame, bg=BG_DARK)
        file_row.pack(fill=tk.X, pady=(0, 8))
        open_btn = create_styled_button(file_row, "Opna skrá", lambda: os.startfile(current_output_path()), width=18, accent=False)
        open_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        open_location_btn = create_styled_button(file_row, "Opna staðsetningu", lambda: open_file_location(current_output_path()), width=18, accent=False)
        open_location_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

        tok_input_btn = create_styled_button(frame, "Keyra inn í tok", open_tok_input_after_sync, width=20, accent=False)

    nav_parent = action_frame if output_name else frame
    nav_row = tk.Frame(nav_parent, bg=BG_DARK)
    nav_row.pack(fill=tk.X, padx=0 if output_name else 42, pady=(0, 10))
    open_raw_btn = create_styled_button(
        nav_row,
        "Opna upprunaskrá",
        open_raw_file,
        width=18,
        accent=False,
    )
    open_raw_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    back_nav_btn = create_styled_button(
        nav_row,
        "Til baka í valmynd",
        lambda: leave_success_page(initialize_main_menu),
        width=18,
        accent=False,
    )
    back_nav_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

    back_button = create_styled_button(
        frame,
        "← Til baka í valmynd",
        lambda: leave_success_page(initialize_main_menu),
        width=20,
    )
    root.protocol("WM_DELETE_WINDOW", close_app_from_success)
    ensure_window_fits()


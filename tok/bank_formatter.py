"""Bank formatter workflow, review dialogs, compression, and success views."""

import datetime
from decimal import Decimal, ROUND_HALF_UP
import os
import re
import tkinter as tk
from tkinter import filedialog, ttk

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
_bank_output_dir = _not_configured
_compressed_output_dir = _not_configured
_auto_code_training_dir = _not_configured
_auto_code_key_dir = _not_configured
_auto_code_memory_path = _not_configured


def configure(**dependencies):
    globals().update(dependencies)


def set_theme(theme):
    globals().update(theme)


def initialize_bank_formatter():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Bank Formatter")

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
        set_bank_formatter_status("Select a bank Excel file before refreshing.", TEXT_WARNING)
        return None
    if "input_drop" in globals() and input_drop and input_drop.winfo_exists():
        input_drop.config(text=os.path.basename(path), fg=TEXT_PRIMARY)
    bank_type, _ = detect_bank_type(path)
    if bank_type:
        suffix = " (auto-code skipped)" if BANK_CONFIGS[bank_type].get('skip_auto_coding') else ""
        set_bank_formatter_status(f"Detected: {BANK_CONFIGS[bank_type]['name']}{suffix}", TEXT_SUCCESS)
    else:
        set_bank_formatter_status("Unknown bank format", TEXT_ERROR, show_open=True)
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
    label = create_styled_label(bank_recent_files_frame, "Recent files", size=9, color=TEXT_SECONDARY, bg=BG_CARD)
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

def clear_bank_input_selection():
    if "bank_input_file_entry" in globals() and bank_input_file_entry and bank_input_file_entry.winfo_exists():
        bank_input_file_entry.delete(0, tk.END)
    if "input_drop" in globals() and input_drop and input_drop.winfo_exists():
        input_drop.config(text="Drop bank Excel file here or click to browse", fg=TEXT_SECONDARY)
    set_bank_formatter_status("Select input file to auto-detect bank", TEXT_SECONDARY, show_open=False)
    schedule_layout_refresh()

def display_bank_formatter_controls():
    global bank_input_file_entry, bank_status_label, bank_output_name_entry, bank_auto_code_var, bank_restaurant_mode_var
    global bank_counter_mode_var, bank_counter_code_entry
    global bank_status_open_btn, input_drop, bank_recent_files_frame

    controls_card, controls_frame = create_panel(frame, padx=22, pady=18)
    controls_card.pack(fill=tk.X, padx=28, pady=(0, 12))

    # Input file row
    # Using two drop areas stacked: top for input, bottom for output
    drop_frame = tk.Frame(controls_frame, bg=BG_CARD)
    drop_frame.pack(fill=tk.X)

    # Input drop
    input_label = create_styled_label(drop_frame, "Input file", size=10, color=TEXT_PRIMARY, bold=True)
    input_label.pack(anchor='w', pady=(0, 6))
    input_drop = create_drop_box(drop_frame, "Drop bank Excel file here or click to browse", height=5)
    input_drop.pack(fill=tk.X, pady=(0, 8))

    bank_input_file_entry = create_styled_entry(drop_frame, width=35)
    bank_input_file_entry.pack_forget()

    bank_recent_files_frame = tk.Frame(drop_frame, bg=BG_CARD)
    bank_recent_files_frame.pack(fill=tk.X, pady=(0, 8))

    def on_drop_input(path):
        if path:
            select_bank_input_file(path)
    if not attach_drop_target(input_drop, on_drop_input):
        input_drop.config(text="Click to browse. Install tkinterdnd2 for drag-and-drop.", fg=TEXT_WARNING)
    input_drop.bind("<Button-1>", lambda _event: browse_bank_input_file())

    # Output name (optional)
    output_name_label = create_styled_label(drop_frame, "Output file name (optional)", size=10, color=TEXT_PRIMARY, bold=True)
    output_name_label.pack(anchor='w', pady=(0, 5))
    bank_output_name_entry = create_styled_entry(drop_frame, width=35)
    bank_output_name_entry.pack(fill=tk.X, ipady=5, pady=(0, 8))

    output_info = create_styled_label(drop_frame, f"Output folder: {_bank_output_dir()}", size=9, color=TEXT_SECONDARY)
    output_info.configure(wraplength=480, justify='center')
    output_info.pack(pady=(0, 9))

    options_frame = tk.Frame(drop_frame, bg=BG_CARD)
    options_frame.pack(fill=tk.X, pady=(2, 12))
    options_label = create_styled_label(options_frame, "Coding options", size=9, color=TEXT_SECONDARY, bg=BG_CARD)
    options_label.pack(anchor='w', pady=(0, 8))

    bank_auto_code_var = tk.BooleanVar(value=True)
    auto_code_setting = create_segmented_setting(
        options_frame,
        "Kóða færslur",
        bank_auto_code_var,
        [("Auto", True), ("Manual", False)],
    )
    auto_code_setting.pack(fill=tk.X, pady=(0, 10))

    bank_restaurant_mode_var = tk.BooleanVar(value=False)
    restaurant_mode_setting = create_segmented_setting(
        options_frame,
        "Business típa",
        bank_restaurant_mode_var,
        [("Almennt", False), ("Veitingarstaður", True)],
        hint="Vörukaup breytast í 2107",
    )
    restaurant_mode_setting.pack(fill=tk.X, pady=(0, 10))

    counter_block = tk.Frame(options_frame, bg=BG_CARD)
    counter_block.pack(fill=tk.X)
    create_styled_label(counter_block, "Nota 7810 eða eithhvað annað", size=9, color=TEXT_PRIMARY, bold=True, bg=BG_CARD).pack(anchor='w')
    create_styled_label(counter_block, "Choose what code fills on bank side", size=8, color=TEXT_SECONDARY, bg=BG_CARD).pack(anchor='w', pady=(1, 6))

    bank_counter_mode_var = tk.StringVar(value="7810")
    counter_mode_row = tk.Frame(counter_block, bg=BG_CARD)
    counter_mode_row.pack(fill=tk.X, pady=(0, 8))
    counter_mode_buttons = []

    def refresh_counter_mode():
        current = bank_counter_mode_var.get()
        for button, value in counter_mode_buttons:
            set_button_accent(button, current == value)
        custom_enabled = current == "custom"
        state = "normal" if custom_enabled else "disabled"
        try:
            if custom_enabled:
                custom_code_row.pack(fill=tk.X)
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
            schedule_layout_refresh()
        except Exception:
            pass

    def select_counter_mode(value):
        bank_counter_mode_var.set(value)
        refresh_counter_mode()

    for index, (text, value) in enumerate((("7810", "7810"), ("Custom", "custom"))):
        button = create_styled_button(counter_mode_row, text, lambda v=value: select_counter_mode(v), width=12, accent=False, height=34)
        button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6) if index == 0 else (0, 0))
        counter_mode_buttons.append((button, value))

    custom_code_row = tk.Frame(counter_block, bg=BG_CARD)
    create_styled_label(custom_code_row, "Custom code (blank leaves counter empty)", size=8, color=TEXT_SECONDARY, bg=BG_CARD).pack(side=tk.LEFT)
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
        text="Select input file to auto-detect bank",
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
        "Open File",
        lambda: os.startfile(bank_input_file_entry.get().strip()) if bank_input_file_entry.get().strip() else None,
        width=10,
        accent=False,
        height=42,
    )

    def refresh_bank_file():
        path = bank_input_file_entry.get().strip()
        if not path:
            set_bank_formatter_status("Select a bank Excel file before refreshing.", TEXT_WARNING)
            return
        select_bank_input_file(path)

    # Buttons
    action_row = tk.Frame(controls_frame, bg=BG_CARD)
    action_row.pack(fill=tk.X, pady=(0, 8))
    run_btn = create_styled_button(action_row, "Run", lambda: run_bank_formatter_script(autofill_7810=True), width=38, height=42)
    run_btn.pack(fill=tk.X, expand=True)

    nav_row = tk.Frame(controls_frame, bg=BG_CARD)
    nav_row.pack(fill=tk.X)
    refresh_btn = create_styled_button(nav_row, "Refresh File", refresh_bank_file, width=18, accent=False, height=42)
    refresh_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    back_btn = create_styled_button(nav_row, "← Back to Menu", initialize_main_menu, width=18, accent=False)
    if hasattr(back_btn, "configure"):
        back_btn.configure(height=42)
    back_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))
    refresh_bank_recent_buttons()
    ensure_window_fits()

def browse_bank_input_file():
    filename = filedialog.askopenfilename(
        initialdir=os.getcwd(),
        title="Select Input File",
        filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*"))
    )
    if filename:
        select_bank_input_file(filename)

def _bank_output_path(output_name, input_file_path=""):
    output_dir = _bank_output_dir()
    output_name = output_name.strip()
    if output_name:
        if not output_name.lower().endswith(".xlsx"):
            output_name += ".xlsx"
        filename = output_name
    else:
        base_name = os.path.splitext(os.path.basename(input_file_path))[0] if input_file_path else "innlestur"
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{base_name}_innlestur_{timestamp}.xlsx"
    return os.path.join(output_dir, filename)

def _bank_auto_code_enabled():
    if "bank_auto_code_var" not in globals():
        return True
    try:
        return bool(bank_auto_code_var.get())
    except Exception:
        return True


def _bank_restaurant_mode_enabled():
    if "bank_restaurant_mode_var" not in globals():
        return False
    try:
        return bool(bank_restaurant_mode_var.get())
    except Exception:
        return False


def _bank_industry_context():
    return "restaurant" if _bank_restaurant_mode_enabled() else ""


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


def _apply_bank_auto_coding(output_df, source_df, input_file_path, enabled=True, skip=False, note="", industry_context=""):
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
    return apply_auto_debit_codes(
        output_df,
        source_df=source_df,
        input_file_path=input_file_path,
        enabled=True,
        training_dir=_auto_code_training_dir(),
        key_dir=_auto_code_key_dir(),
        memory_path=_auto_code_memory_path(),
        industry_context=industry_context,
    )


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
    return text or "Missing text"


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
    dialog.title("Review Auto-Code Cluster")
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
    current_debit = debit_values[0] if len(debit_values) == 1 else ("Mixed" if debit_values else "-")

    header_panel, header = create_panel(outer, padx=18, pady=14)
    header_panel.grid(row=0, column=0, sticky="ew")

    eyebrow_row = tk.Frame(header, bg=BG_CARD)
    eyebrow_row.pack(fill=tk.X)
    create_styled_label(
        eyebrow_row,
        "AUTO-CODE REVIEW",
        size=8,
        color=TEXT_SECONDARY,
        bold=True,
        bg=BG_CARD,
    ).pack(side=tk.LEFT)
    progress_label = create_styled_label(
        eyebrow_row,
        f"{current_index} of {total_clusters}",
        size=8,
        color=TEXT_ON_ACCENT,
        bold=True,
        bg=ACCENT,
    )
    progress_label.pack(side=tk.RIGHT, ipadx=10, ipady=4)

    create_styled_label(
        header,
        "Review cluster",
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

    rows_stat = add_stat(stats_row, "Rows", str(len(cluster_df)), (0, 8))
    confidence_stat = add_stat(stats_row, "Lowest confidence", f"{min_confidence}", (0, 8))
    debit_stat = add_stat(stats_row, "Current debit", current_debit, (0, 0))

    table_panel, table_inner = create_panel(outer, padx=0, pady=0)
    table_panel.grid(row=1, column=0, sticky="nsew", pady=(10, 0))

    table_title_row = tk.Frame(table_inner, bg=BG_CARD)
    table_title_row.pack(fill=tk.X, padx=14, pady=(10, 6))
    create_styled_label(
        table_title_row,
        "Low-confidence rows",
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
    tree.heading("confidence", text="CONF", anchor=tk.E)
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
        selected_count_label.config(text=f"{len(selected_positions)} selected" if selected_positions else "")

    def refresh_stats():
        rows_text = str(len(active_indices))
        if split_batches:
            rows_text += f" ({sum(len(batch) for batch in split_batches)} split)"
        current_rows = [cluster_df.loc[idx] for idx in active_indices if idx in cluster_df.index]
        current_conf = [_auto_review_confidence(row.get("CONFIDENCE")) for row in current_rows]
        current_debits = []
        for row in current_rows:
            debit_value = _format_auto_review_value(row.get("DEBIT"), limit=12).strip()
            if debit_value and debit_value not in current_debits:
                current_debits.append(debit_value)
        rows_stat.config(text=f"Rows: {rows_text}")
        confidence_stat.config(text=f"Lowest confidence: {min(current_conf) if current_conf else 0}")
        debit_stat.config(text=f"Current debit: {current_debits[0] if len(current_debits) == 1 else ('Mixed' if current_debits else '-')}")

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
    create_styled_label(form_row, "Debit code", size=9, color=TEXT_SECONDARY, bg=BG_CARD).pack(side=tk.LEFT, padx=(0, 10))
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
            status_label.config(text="Enter a 3-5 digit code.")
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
            status_label.config(text="Select txs to split into a later batch.")
            return
        if len(positions) >= len(active_indices):
            status_label.config(text="Cannot split every tx out of this batch.")
            return
        excluded_indices = [active_indices[pos] for pos in positions if 0 <= pos < len(active_indices)]
        for pos in reversed(positions):
            if 0 <= pos < len(active_indices):
                active_indices.pop(pos)
        split_batches.append(excluded_indices)
        status_label.config(text=f"Split {len(excluded_indices)} tx(s) to a later batch.", fg=TEXT_SUCCESS)
        refresh_table()

    refresh_table()

    code_btn = create_styled_button(button_row, "Code", submit_code, width=10, height=34)
    code_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    exclude_btn = create_styled_button(button_row, "Exclude Selected", exclude_selected, width=16, height=34, accent=False)
    exclude_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
    if can_go_back:
        back_btn = create_styled_button(secondary_button_row, "Back", go_back, width=10, height=34, accent=False)
        back_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    skip_btn = create_styled_button(secondary_button_row, "Skip", skip_cluster, width=10, height=34, accent=False)
    skip_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0 if not can_go_back else 5, 5))
    skip_all_btn = create_styled_button(secondary_button_row, "Skip All", skip_all, width=12, height=34, accent=False)
    skip_all_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

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
        reviewed.loc[current_indices, "CONFIDENCE"] = 100
        reviewed.loc[current_indices, "_AUTO_CODE_SOURCE"] = "manual_review"
        cluster_index += 1
    return reviewed


def _strip_auto_code_internal_columns(df):
    return df.drop(columns=[column for column in df.columns if str(column).startswith("_AUTO_CODE_")], errors="ignore")


def _prepare_bank_excel_output(df):
    output = _strip_auto_code_internal_columns(df.copy())
    if "CONFIDENCE" in output.columns and "Cnf" not in output.columns:
        output.rename(columns={"CONFIDENCE": "Cnf"}, inplace=True)
    preferred = ["DATE", "TEXT", "DEBIT", "Cnf", "ID", "AMOUNT", "CREDIT", "Positive/Negative", "STATUS"]
    ordered = [column for column in preferred if column in output.columns]
    ordered.extend(column for column in output.columns if column not in ordered)
    return output[ordered]


def _normalize_tx_text_key(value):
    return _auto_review_cluster_key(value)


def _normalize_tx_id_key(value):
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    if not text:
        return ""
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\s+", "", text)


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
    folder = _compressed_output_dir()
    stem, ext = os.path.splitext(os.path.basename(output_file_path))
    ext = ext or ".xlsx"
    base = os.path.join(folder, f"{stem}_compressed_txs{ext}")
    if not os.path.exists(base):
        return base
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(folder, f"{stem}_compressed_txs_{timestamp}{ext}")


def _find_compressible_transaction_clusters(df, minimum_count=10):
    required = {"DATE", "TEXT", "DEBIT", "AMOUNT", "CREDIT"}
    if not required.issubset(set(df.columns)):
        return []

    grouped = {}
    for idx, row in df.iterrows():
        text_key = _normalize_tx_text_key(row.get("TEXT"))
        if not text_key:
            continue
        sign, debit_code, credit_code = _tx_code_pair(row)
        if not sign:
            continue
        amount_ok = _parse_bank_amount_decimal(row.get("AMOUNT")) is not None
        grouped.setdefault((text_key, sign), []).append({
            "idx": idx,
            "code_key": (debit_code, credit_code),
            "debit_code": debit_code,
            "credit_code": credit_code,
            "id_key": _normalize_tx_id_key(row.get("ID")),
            "amount_ok": amount_ok,
        })

    clusters = []
    for (text_key, sign), rows_info in grouped.items():
        if len(rows_info) < minimum_count:
            continue
        if any(not item["amount_ok"] for item in rows_info):
            continue
        code_keys = {item["code_key"] for item in rows_info}
        if len(code_keys) != 1:
            continue
        debit_code, credit_code = next(iter(code_keys))
        if not debit_code and not credit_code:
            continue
        nonblank_ids = {item["id_key"] for item in rows_info if item["id_key"]}
        if len(nonblank_ids) > 1:
            continue
        row_indices = [item["idx"] for item in rows_info]
        cluster_df = df.loc[row_indices].copy()
        total = _sum_tx_amounts(cluster_df)
        if total is None:
            continue
        first_row = cluster_df.iloc[0]
        display_name = _auto_review_display_name(first_row.get("TEXT", text_key))
        try:
            first_position = int(row_indices[0])
        except Exception:
            first_position = len(clusters)
        clusters.append({
            "key": (text_key, sign),
            "display_name": display_name,
            "sign": sign,
            "debit_code": debit_code,
            "credit_code": credit_code,
            "row_indices": row_indices,
            "count": len(row_indices),
            "total": total,
            "first_position": first_position,
        })
    return sorted(clusters, key=lambda item: item["first_position"])


def _show_transaction_compression_review(cluster_df, cluster, current_index, total_clusters):
    result = {"action": "skip"}

    dialog = tk.Toplevel(root)
    dialog.title("Compress Transactions")
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
        "TX COMPRESSION",
        size=8,
        color=TEXT_SECONDARY,
        bold=True,
        bg=BG_CARD,
    ).pack(side=tk.LEFT)
    progress_label = create_styled_label(
        eyebrow_row,
        f"{current_index} of {total_clusters}",
        size=8,
        color=TEXT_ON_ACCENT,
        bold=True,
        bg=ACCENT,
    )
    progress_label.pack(side=tk.RIGHT, ipadx=10, ipady=4)
    title = create_styled_label(
        header,
        "Compress transactions",
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

    code_label = f"debit {cluster.get('debit_code') or '-'} / credit {cluster.get('credit_code') or '-'}"
    name_label = create_styled_label(
        header,
        _auto_review_display_name(cluster.get("display_name")),
        size=9,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    )
    name_label.configure(wraplength=560, justify="left")
    name_label.pack(anchor="w")

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

    rows_stat = add_stat(stats_row, "Rows", str(len(active_indices)), (0, 8))
    sign_stat = add_stat(stats_row, "Sign", cluster.get("sign") or "-", (0, 8))
    code_stat = add_stat(stats_row, "Codes", code_label, (0, 8))
    total_stat = add_stat(stats_row, "Sum", _format_auto_review_amount(active_total()), (0, 0))

    table_panel, table_inner = create_panel(outer, padx=0, pady=0)
    table_panel.grid(row=1, column=0, sticky="nsew", pady=(10, 0))

    table_title_row = tk.Frame(table_inner, bg=BG_CARD)
    table_title_row.pack(fill=tk.X, padx=14, pady=(10, 6))
    create_styled_label(
        table_title_row,
        "Transactions",
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
    tree.heading("sign", text="SIGN", anchor=tk.CENTER)
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
            rows_text += f" ({excluded_count} excluded)"
        rows_stat.config(text=f"Rows: {rows_text}")
        sign_stat.config(text=f"Sign: {cluster.get('sign') or '-'}")
        code_stat.config(text=f"Codes: {code_label}")
        total_stat.config(text=f"Sum: {_format_auto_review_amount(active_total())}")

    def refresh_selected_count():
        selected_count_label.config(text=f"{len(selected_positions)} selected" if selected_positions else "")

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
            status_label.config(text="Select one or more txs to exclude.", fg=TEXT_WARNING)
            return
        if len(positions) >= len(active_indices):
            status_label.config(text="Cannot exclude every tx in the cluster.", fg=TEXT_WARNING)
            return
        for pos in positions:
            if 0 <= pos < len(active_indices):
                active_indices.pop(pos)
        status_label.config(text=f"Excluded {len(positions)} tx(s) from this compression.", fg=TEXT_SUCCESS)
        refresh_table()

    refresh_table()

    button_row = tk.Frame(action_inner, bg=BG_CARD)
    button_row.pack(fill=tk.X, pady=(10, 0))

    def compress_cluster():
        if len(active_indices) < 2:
            status_label.config(text="Need at least 2 txs left to compress.", fg=TEXT_WARNING)
            return
        result["action"] = "compress"
        result["row_indices"] = list(active_indices)
        result["count"] = len(active_indices)
        result["total"] = active_total()
        dialog.destroy()

    def skip_cluster():
        result["action"] = "skip"
        dialog.destroy()

    def skip_all():
        result["action"] = "skip_all"
        dialog.destroy()

    compress_btn = create_styled_button(button_row, "Compress", compress_cluster, width=12, height=34)
    compress_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    exclude_btn = create_styled_button(button_row, "Exclude Selected", exclude_selected, width=16, height=34, accent=False)
    exclude_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    skip_btn = create_styled_button(button_row, "Skip", skip_cluster, width=10, height=34, accent=False)
    skip_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    skip_all_btn = create_styled_button(button_row, "Skip All", skip_all, width=12, height=34, accent=False)
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


def _apply_transaction_compressions(df, selected_clusters):
    selected_by_index = {}
    cluster_by_first = {}
    for group_number, cluster in enumerate(selected_clusters, start=1):
        row_indices = list(cluster["row_indices"])
        if not row_indices:
            continue
        first_idx = row_indices[0]
        cluster_by_first[first_idx] = (group_number, cluster)
        for idx in row_indices:
            selected_by_index[idx] = first_idx

    output_rows = []
    reference_rows = []
    reference_columns = list(df.columns) + [
        "COMPRESSED_GROUP",
        "REFERENCE_TYPE",
        "COMPRESSED_COUNT",
        "COMPRESSED_SUM_AMOUNT",
    ]
    blank_reference_row = {column: "" for column in reference_columns}
    for idx, row in df.iterrows():
        first_idx = selected_by_index.get(idx)
        if first_idx is None:
            output_rows.append(row.to_dict())
            continue
        if idx != first_idx:
            continue

        group_number, cluster = cluster_by_first[first_idx]
        row_indices = list(cluster["row_indices"])
        cluster_df = df.loc[row_indices].copy()
        summary = cluster_df.iloc[0].copy()
        summary["AMOUNT"] = _format_compressed_amount(cluster["total"])
        output_rows.append(summary.to_dict())

        for _, detail_row in cluster_df.iterrows():
            ref_row = detail_row.to_dict()
            ref_row["COMPRESSED_GROUP"] = group_number
            ref_row["REFERENCE_TYPE"] = "deleted_tx"
            ref_row["COMPRESSED_COUNT"] = cluster["count"]
            ref_row["COMPRESSED_SUM_AMOUNT"] = _format_compressed_amount(cluster["total"])
            reference_rows.append(ref_row)

        reference_rows.append(dict(blank_reference_row))
        sum_row = summary.to_dict()
        sum_row["COMPRESSED_GROUP"] = group_number
        sum_row["REFERENCE_TYPE"] = "compressed_sum"
        sum_row["COMPRESSED_COUNT"] = cluster["count"]
        sum_row["COMPRESSED_SUM_AMOUNT"] = _format_compressed_amount(cluster["total"])
        reference_rows.append(sum_row)
        reference_rows.append(dict(blank_reference_row))

    compressed_df = pd.DataFrame(output_rows, columns=df.columns)
    reference_df = pd.DataFrame(reference_rows, columns=reference_columns)
    return compressed_df, reference_df


def compress_transactions_interactive(output_file_path):
    if not output_file_path or not os.path.exists(output_file_path):
        return {"ok": False, "message": "Output file not found."}
    try:
        df = pd.read_excel(output_file_path)
    except Exception as exc:
        return {"ok": False, "message": f"Could not read output file: {str(exc)[:70]}"}

    clusters = _find_compressible_transaction_clusters(df, minimum_count=10)
    if not clusters:
        return {"ok": False, "message": "No 10+ same-code transaction clusters found."}

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
            selected.append(selected_cluster)

    if not selected:
        return {"ok": False, "message": "No transaction clusters compressed."}

    compressed_df, reference_df = _apply_transaction_compressions(df, selected)
    reference_path = _compressed_reference_path(output_file_path)
    try:
        _prepare_bank_excel_output(compressed_df).to_excel(output_file_path, index=False)
        _prepare_bank_excel_output(reference_df).to_excel(reference_path, index=False)
        autofit_excel_columns(output_file_path)
        autofit_excel_columns(reference_path)
    except Exception as exc:
        return {"ok": False, "message": f"Compression save failed: {str(exc)[:75]}"}

    deleted_count = sum(cluster["count"] for cluster in selected)
    return {
        "ok": True,
        "message": (
            f"Compressed {deleted_count} txs into {len(selected)} row(s). "
            f"Reference: {os.path.basename(reference_path)}"
        ),
        "reference_path": reference_path,
        "clusters": len(selected),
        "deleted_count": deleted_count,
    }


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
    output_name = bank_output_name_entry.get() if "bank_output_name_entry" in globals() else ""
    output_file_path = _bank_output_path(output_name, input_file_path=input_file_path)
    auto_code_debits = _bank_auto_code_enabled()
    restaurant_mode = _bank_restaurant_mode_enabled()
    industry_context = "restaurant" if restaurant_mode else ""
    fill_counter, counter_code = _bank_counter_fill(autofill_7810)

    if not input_file_path:
        set_bank_formatter_status("Please select an input file", TEXT_WARNING)
        return

    bank_type, df = detect_bank_type(input_file_path)

    if bank_type is None:
        set_bank_formatter_status("Could not detect bank format", TEXT_ERROR, show_open=True)
        play_error_sound()
        return

    config = BANK_CONFIGS[bank_type]

    try:
        # Check if this bank needs custom processing
        if config.get('custom_processor'):
            if bank_type == 'islandsbanki_innheimta':
                process_islandsbanki_innheimta(df, output_file_path, config['name'], autofill_7810=fill_counter)
            elif bank_type == 'islandsbanki_kort':
                process_islandsbanki_kort(df, output_file_path, config['name'], autofill_7810=fill_counter, input_file_path=input_file_path, auto_code_debits=auto_code_debits, industry_context=industry_context, counter_code=counter_code)
            elif bank_type == 'islandsbanki_special':
                process_islandsbanki_special(df, output_file_path, config['name'], autofill_7810=fill_counter, counter_code=counter_code)
            elif bank_type == 'sala_yfirlit':
                process_sala_yfirlit(df, output_file_path, config['name'], autofill_7810=fill_counter)
            elif bank_type == 'landsbankinn_kort':
                process_landsbankinn_kort(df, output_file_path, config['name'], autofill_7810=fill_counter, input_file_path=input_file_path, auto_code_debits=auto_code_debits, industry_context=industry_context, counter_code=counter_code)
            elif bank_type == 'arion_kort':
                process_arion_kort(df, output_file_path, config['name'], autofill_7810=fill_counter, input_file_path=input_file_path, auto_code_debits=auto_code_debits, industry_context=industry_context, counter_code=counter_code)
        else:
            # Standard processing for regular banks
            extracted_columns = df[config['columns']].copy()
            amount_col = config['columns'][3]
            amount_source = _arion_amount_source(df, amount_col) if bank_type == 'arion' else extracted_columns[amount_col]
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
            )
            extracted_columns = review_low_confidence_auto_code_clusters(
                extracted_columns,
                enabled=auto_code_debits and not config.get('skip_auto_coding', False),
            )
            extracted_columns = _strip_auto_code_internal_columns(extracted_columns)

            # Sort by Positive/Negative and TEXT
            extracted_columns = extracted_columns.sort_values(by=['Positive/Negative', 'TEXT'])

            # Convert DATE to text to preserve format on copy-paste
            extracted_columns['DATE'] = extracted_columns['DATE'].apply(format_date_as_text)

            # Save to Excel
            _prepare_bank_excel_output(extracted_columns).to_excel(output_file_path, index=False)

            # Auto-fit column widths
            autofit_excel_columns(output_file_path)

            display_bank_formatter_success(config['name'], output_file_path)

    except Exception as e:
        set_bank_formatter_status(f"✗ Error: {str(e)[:40]}", TEXT_ERROR)
        play_error_sound()


def process_islandsbanki_innheimta(df, output_file_path, bank_name, autofill_7810=True):
    # Do not run merchant/purpose auto-coding on this format; it already creates fixed calculation rows.
    columns_to_select = ['Kennitala', 'Greiðsludagur', 'Fjármagnstekjuskattur', 'Dráttarvextir', 'Rst. Upphæð', 'Upphæð']
    df_selected = df[columns_to_select].copy()

    entries = []
    for _, row in df_selected.iterrows():
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
                'CREDIT': 7620
            })
        if drattar != 0:
            entries.append({
                'DATE': date_val,
                'TEXT': text_val,
                'DEBIT': 7620,
                'AMOUNT': drattar,
                'ID': id_val,
                'CREDIT': 6100
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
                'CREDIT': 4400
            })

    df_output = pd.DataFrame(entries, columns=['DATE', 'TEXT', 'DEBIT', 'AMOUNT', 'ID', 'CREDIT'])
    _prepare_bank_excel_output(df_output).to_excel(output_file_path, index=False)
    autofit_excel_columns(output_file_path)
    display_bank_formatter_success(bank_name, output_file_path)


def process_sala_yfirlit(df, output_file_path, bank_name, autofill_7810=True):
    # Do not run merchant/purpose auto-coding on this format; Reikningur nr is an internal incremental invoice number.
    # Drop total rows if present (where Kennitala is NaN and/or Nafn starts with 'Alls')
    df_clean = df.copy()
    df_clean = df_clean[~df_clean['Nafn'].astype(str).str.lower().str.startswith('alls')]
    cols = ['Nafn', 'Kennitala', 'Upphæð með vsk', 'Reikningur nr', 'Dagsetning']
    df_selected = df_clean[cols].copy()

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

    df_final = df_selected[['DATE', 'TEXT', 'DEBIT', 'AMOUNT', 'ID', 'CREDIT']].copy()
    _prepare_bank_excel_output(df_final).to_excel(output_file_path, index=False)
    autofit_excel_columns(output_file_path)
    display_bank_formatter_success(bank_name, output_file_path)

def process_islandsbanki_special(df, output_file_path, bank_name, autofill_7810=True, counter_code=7810):
    # Do not run merchant/purpose auto-coding on this format; it is a special calculation/subtraction import.
    config = BANK_CONFIGS['islandsbanki_special']
    df_selected = df[config['columns']].copy()
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
    df_output = df_selected[['DATE', 'TEXT', 'DEBIT', 'ID', 'AMOUNT', 'CREDIT', 'Positive/Negative']].copy()
    _prepare_bank_excel_output(df_output).to_excel(output_file_path, index=False)
    autofit_excel_columns(output_file_path)
    display_bank_formatter_success(bank_name, output_file_path)

def process_islandsbanki_kort(df, output_file_path, bank_name, autofill_7810=True, input_file_path="", auto_code_debits=True, industry_context="", counter_code=7810):
    config = BANK_CONFIGS['islandsbanki_kort']
    df_selected = df[config['columns']].copy()
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
    df_output = df_selected[['DATE', 'TEXT', 'DEBIT', 'ID', 'AMOUNT', 'CREDIT', 'Positive/Negative']].copy()
    df_output = _apply_bank_auto_coding(df_output, df, input_file_path, enabled=auto_code_debits, industry_context=industry_context)
    df_output = review_low_confidence_auto_code_clusters(df_output, enabled=auto_code_debits)
    df_output = _strip_auto_code_internal_columns(df_output)
    _prepare_bank_excel_output(df_output).to_excel(output_file_path, index=False)
    autofit_excel_columns(output_file_path)
    display_bank_formatter_success(bank_name, output_file_path)

def process_landsbankinn_kort(df, output_file_path, bank_name, autofill_7810=True, input_file_path="", auto_code_debits=True, industry_context="", counter_code=7810):
    config = BANK_CONFIGS['landsbankinn_kort']
    df_selected = df[config['columns']].copy()
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
    df_output = df_selected[['DATE', 'TEXT', 'DEBIT', 'ID', 'AMOUNT', 'CREDIT', 'Positive/Negative']].copy()
    df_output = _apply_bank_auto_coding(df_output, df, input_file_path, enabled=auto_code_debits, industry_context=industry_context)
    df_output = review_low_confidence_auto_code_clusters(df_output, enabled=auto_code_debits)
    df_output = _strip_auto_code_internal_columns(df_output)
    _prepare_bank_excel_output(df_output).to_excel(output_file_path, index=False)
    autofit_excel_columns(output_file_path)
    display_bank_formatter_success(bank_name, output_file_path)

def process_arion_kort(df, output_file_path, bank_name, autofill_7810=True, input_file_path="", auto_code_debits=True, industry_context="", counter_code=7810):
    config = BANK_CONFIGS['arion_kort']
    df_selected = df[config['columns']].copy()
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
    df_output = df_selected[['DATE', 'TEXT', 'DEBIT', 'ID', 'AMOUNT', 'CREDIT', 'Positive/Negative']].copy()
    df_output = _apply_bank_auto_coding(df_output, df, input_file_path, enabled=auto_code_debits, industry_context=industry_context)
    df_output = review_low_confidence_auto_code_clusters(df_output, enabled=auto_code_debits)
    df_output = _strip_auto_code_internal_columns(df_output)
    _prepare_bank_excel_output(df_output).to_excel(output_file_path, index=False)
    autofit_excel_columns(output_file_path)
    display_bank_formatter_success(bank_name, output_file_path)

def display_bank_formatter_success(bank_name, output_file_path):
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Success", TEXT_SUCCESS)
    play_success_sound()

    path_state = {"path": output_file_path}
    edit_state = {"active": False}
    compression_reference_state = {"path": ""}
    compression_reference_edit_state = {"active": False}

    def current_output_path():
        return path_state["path"]

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

    def show_rename_status(message, color=TEXT_SECONDARY):
        try:
            filename_canvas.itemconfigure(status_text, text=message, fill=color)
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
            show_rename_status("Enter a file name.", TEXT_WARNING)
            return

        new_name = os.path.basename(new_name)
        if not new_name.lower().endswith(".xlsx"):
            new_name += ".xlsx"
            output_name_var.set(new_name)

        if re.search(r'[<>:"/\\|?*]', new_name):
            show_rename_status("That file name has invalid characters.", TEXT_ERROR)
            return

        new_path = os.path.join(os.path.dirname(old_path), new_name)
        old_abs = os.path.normcase(os.path.abspath(old_path))
        new_abs = os.path.normcase(os.path.abspath(new_path))
        if old_abs == new_abs:
            show_rename_status("File name is already up to date.", TEXT_SECONDARY)
            return
        if os.path.exists(new_path):
            show_rename_status("A file with that name already exists.", TEXT_ERROR)
            return

        try:
            os.rename(old_path, new_path)
            path_state["path"] = new_path
            output_name_var.set(os.path.basename(new_path))
            refresh_filename_display()
            end_filename_edit()
            show_rename_status("Renamed.", TEXT_SUCCESS)
        except Exception as exc:
            show_rename_status(f"Rename failed: {str(exc)[:45]}", TEXT_ERROR)

    def start_filename_edit():
        edit_state["active"] = True
        output_name_entry.delete(0, tk.END)
        output_name_entry.insert(0, output_name_var.get())
        filename_canvas.itemconfigure(filename_text, state="hidden")
        filename_canvas.itemconfigure(entry_window, state="normal")
        filename_action_btn.set_text("Save")
        filename_action_btn.set_accent(True)
        show_rename_status("", TEXT_SECONDARY)
        output_name_entry.focus_set()
        output_name_entry.select_range(0, tk.END)

    def end_filename_edit():
        edit_state["active"] = False
        filename_canvas.itemconfigure(entry_window, state="hidden")
        filename_canvas.itemconfigure(filename_text, state="normal")
        filename_action_btn.set_text("Edit")
        filename_action_btn.set_accent(False)

    def toggle_filename_edit():
        if edit_state["active"]:
            rename_output_file()
        else:
            start_filename_edit()

    def compress_output_transactions():
        show_rename_status("Checking compressible transaction clusters...", TEXT_SECONDARY)
        result = compress_transactions_interactive(current_output_path())
        if result.get("ok"):
            reference_path = result.get("reference_path", "")
            if reference_path:
                show_compressed_reference_box(reference_path)
            show_rename_status(result.get("message", "Compressed transactions."), TEXT_SUCCESS)
            play_success_sound()
        else:
            show_rename_status(result.get("message", "No transactions compressed."), TEXT_WARNING)

    def open_compressed_reference_file():
        reference_path = compression_reference_state.get("path", "")
        if reference_path and os.path.exists(reference_path):
            os.startfile(reference_path)
        else:
            show_compressed_reference_status("Compressed TX reference file not found.", TEXT_WARNING)

    def open_compressed_reference_location():
        reference_path = compression_reference_state.get("path", "")
        if reference_path:
            open_file_location(reference_path)
        else:
            show_compressed_reference_status("Compressed TX reference file not found.", TEXT_WARNING)

    def show_compressed_reference_status(message, color=TEXT_SECONDARY):
        try:
            compressed_canvas.itemconfigure(compressed_status_text, text=message, fill=color)
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
        compressed_action_btn.set_text("Edit")
        compressed_action_btn.set_accent(False)

    def rename_compressed_reference_file():
        old_path = compression_reference_state.get("path", "")
        if not old_path:
            show_compressed_reference_status("No compressed file yet.", TEXT_WARNING)
            return

        new_name = compressed_name_var.get().strip().strip('"')
        if not new_name:
            show_compressed_reference_status("Enter a file name.", TEXT_WARNING)
            return
        new_name = os.path.basename(new_name)
        if not new_name.lower().endswith(".xlsx"):
            new_name += ".xlsx"
            compressed_name_var.set(new_name)
        if re.search(r'[<>:"/\\|?*]', new_name):
            show_compressed_reference_status("That file name has invalid characters.", TEXT_ERROR)
            return

        new_path = os.path.join(os.path.dirname(old_path), new_name)
        old_abs = os.path.normcase(os.path.abspath(old_path))
        new_abs = os.path.normcase(os.path.abspath(new_path))
        if old_abs == new_abs:
            show_compressed_reference_status("File name is already up to date.", TEXT_SECONDARY)
            return
        if os.path.exists(new_path):
            show_compressed_reference_status("A file with that name already exists.", TEXT_ERROR)
            return
        try:
            os.rename(old_path, new_path)
            compression_reference_state["path"] = new_path
            compressed_name_var.set(os.path.basename(new_path))
            refresh_compressed_filename_display()
            end_compressed_filename_edit()
            show_compressed_reference_status("Renamed.", TEXT_SUCCESS)
        except Exception as exc:
            show_compressed_reference_status(f"Rename failed: {str(exc)[:45]}", TEXT_ERROR)

    def start_compressed_filename_edit():
        compression_reference_edit_state["active"] = True
        compressed_name_entry.delete(0, tk.END)
        compressed_name_entry.insert(0, compressed_name_var.get())
        compressed_canvas.itemconfigure(compressed_filename_text, state="hidden")
        compressed_canvas.itemconfigure(compressed_entry_window, state="normal")
        compressed_action_btn.set_text("Save")
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
        show_compressed_reference_status("Reference file ready.", TEXT_SUCCESS)
        try:
            if not compressed_canvas.winfo_manager():
                compressed_canvas.pack(fill=tk.X, pady=(10, 0))
            if not compressed_file_row.winfo_manager():
                compressed_file_row.pack(fill=tk.X, pady=(0, 8), after=compress_btn)
            schedule_layout_refresh()
            ensure_window_fits()
        except Exception:
            pass

    def draw_processed_box(event=None):
        processed_canvas.delete("all")
        w = max(4, processed_canvas.winfo_width())
        _rounded_rect(processed_canvas, 1, 1, w - 1, 72, 12,
                      fill=BUTTON_MUTED, outline=BORDER, width=1)
        processed_canvas.create_text(15, 19, text="Processed as", anchor="w",
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

        filename_action_btn = MiniPillButton(filename_canvas, "Edit", toggle_filename_edit, width=62)
        filename_label = filename_canvas.create_text(
            15, 19, text="Output file", anchor="w",
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
            filename_canvas.delete("filename_bg")
            _rounded_rect(filename_canvas, 1, 1, w - 1, 90, 12,
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

        compressed_action_btn = MiniPillButton(compressed_canvas, "Edit", toggle_compressed_filename_edit, width=62)
        compressed_label = compressed_canvas.create_text(
            15, 19, text="Compressed TXs file", anchor="w",
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
            compressed_canvas.delete("compressed_bg")
            _rounded_rect(compressed_canvas, 1, 1, w - 1, 90, 12,
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

        tok_input_btn = create_styled_button(action_frame, "Keyra inn í tok", lambda: open_tok_input_with_file(current_output_path()), width=20)
        tok_input_btn.pack(fill=tk.X, pady=(0, 8))

        compress_btn = create_styled_button(action_frame, "Compress TXs", compress_output_transactions, width=20, accent=False)
        compress_btn.pack(fill=tk.X, pady=(0, 8))

        compressed_file_row = tk.Frame(action_frame, bg=BG_DARK)
        compressed_file_row.pack(fill=tk.X, pady=(0, 8))
        open_compressed_btn = create_styled_button(compressed_file_row, "Open Compressed TXs", open_compressed_reference_file, width=18, accent=False)
        open_compressed_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        open_compressed_location_btn = create_styled_button(compressed_file_row, "Open Location", open_compressed_reference_location, width=18, accent=False)
        open_compressed_location_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))
        compressed_file_row.pack_forget()

        file_row = tk.Frame(action_frame, bg=BG_DARK)
        file_row.pack(fill=tk.X, pady=(0, 8))
        open_btn = create_styled_button(file_row, "Open File", lambda: os.startfile(current_output_path()), width=18, accent=False)
        open_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        open_location_btn = create_styled_button(file_row, "Open Location", lambda: open_file_location(current_output_path()), width=18, accent=False)
        open_location_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

        tok_input_btn = create_styled_button(frame, "Keyra inn í tok", lambda: open_tok_input_with_file(current_output_path()), width=20, accent=False)

    nav_parent = action_frame if output_name else frame
    nav_row = tk.Frame(nav_parent, bg=BG_DARK)
    nav_row.pack(fill=tk.X, padx=0 if output_name else 42, pady=(0, 10))
    format_another_btn = create_styled_button(nav_row, "Format Another", initialize_bank_formatter, width=18, accent=False)
    format_another_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    back_nav_btn = create_styled_button(nav_row, "Back to Menu", initialize_main_menu, width=18, accent=False)
    back_nav_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

    back_button = create_styled_button(frame, "← Back to Menu", initialize_main_menu, width=20)
    ensure_window_fits()

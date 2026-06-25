import tkinter as tk
from tkinter import ttk, filedialog, PhotoImage
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKDND_AVAILABLE = True
    TKDND_ERROR = ""
except Exception as exc:
    DND_FILES = None
    TkinterDnD = None
    TKDND_AVAILABLE = False
    TKDND_ERROR = str(exc)
import pandas as pd
import pyautogui
import time
import logging
import pyperclip
import gc
import datetime
import os
import sys
import json
import subprocess
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import winsound
from threading import Timer

try:
    try:
        from .auto_coder import apply_auto_debit_codes, get_auto_coder
    except ImportError:
        from auto_coder import apply_auto_debit_codes, get_auto_coder
    AUTO_CODER_AVAILABLE = True
    AUTO_CODER_ERROR = ""
except Exception as exc:
    apply_auto_debit_codes = None
    get_auto_coder = None
    AUTO_CODER_AVAILABLE = False
    AUTO_CODER_ERROR = str(exc)

action_delay = 0.1
focus_delay = 0.02
paste_delay = 0.0
post_paste_delay = 0.02
start_delay = 3
running = True
watchdog_timer = None
after_id = None
settings = {}
SETTINGS_PATH = os.path.join(os.path.expanduser("~"), "AppData", "Local", "TokTenging", "settings.json")
DEBUG_LOG_PATH = os.path.join(os.path.dirname(SETTINGS_PATH), "tok_input_debug.log")
pending_tok_df = None
pending_tok_file_path = None
current_tok_file_path = None
current_tok_run_start_index = 0
current_run_speed_label = "Saved"
automation_stop_reason = ""
AUTOMATION_WATCHDOG_SECONDS = 15
MOUSE_FAILSAFE_MARGIN = 8
RECENT_FILE_LIMIT = 5
RECENT_VISIBLE_LIMIT = 3
SPEED_PRESETS = [
    ("Fast", {"action": 0.048, "focus": 0.080, "paste": 0.018, "post_paste": 0.090, "start": 3}),
    ("Balanced", {"action": 0.060, "focus": 0.100, "paste": 0.022, "post_paste": 0.112, "start": 3}),
    ("Slow", {"action": 0.081, "focus": 0.272, "paste": 0.030, "post_paste": 0.152, "start": 3}),
]
SPEED_PRESET_MAP = {label: values for label, values in SPEED_PRESETS}
SPEED_PRESET_MAP["Test Run"] = {"action": 3.0, "focus": 1.0, "paste": 3.0, "post_paste": 1.0, "start": 3}

# PyAutoGUI sleeps 0.1s after every call by default. The app already controls
# pacing with the delay settings below, so keep PyAutoGUI itself immediate.
pyautogui.PAUSE = 0
pyautogui.FAILSAFE = True

# Sentinel used by the action queue to mean "use the current global action_delay".
USE_ACTION_DELAY = object()

# =============================================================================
# APP MODULES
# =============================================================================
try:
    from . import bank_formatter, ui_components
    from .bank_formatter import initialize_bank_formatter
    from .bank_utils import format_date_as_text, _format_date_input_line
    from .ui_components import (
        RoundedPanel,
        attach_auto_resize_text,
        attach_drop_target,
        create_drop_box,
        create_panel,
        create_segmented_setting,
        create_styled_button,
        create_styled_entry,
        create_styled_label,
        create_styled_text_area,
        finalize_fixed_action_dialog_grid,
        get_text_area_value,
        set_button_accent,
        set_text_area_value,
        _rounded_rect,
    )
except ImportError:
    import bank_formatter
    import ui_components
    from bank_formatter import initialize_bank_formatter
    from bank_utils import format_date_as_text, _format_date_input_line
    from ui_components import (
        RoundedPanel,
        attach_auto_resize_text,
        attach_drop_target,
        create_drop_box,
        create_panel,
        create_segmented_setting,
        create_styled_button,
        create_styled_entry,
        create_styled_label,
        create_styled_text_area,
        finalize_fixed_action_dialog_grid,
        get_text_area_value,
        set_button_accent,
        set_text_area_value,
        _rounded_rect,
    )

ui_components.configure_drag_and_drop(TKDND_AVAILABLE, DND_FILES)
ui_components.set_layout_refresh_callback(lambda: schedule_layout_refresh())
globals().update(ui_components.get_theme("light"))

# Settings layout helpers
SETTINGS_LABEL_WIDTH = 24


def apply_theme(theme_name):
    theme = ui_components.set_theme(theme_name)
    globals().update(theme)
    bank_formatter.set_theme(theme)
    if "root" in globals():
        root.configure(bg=BG_DARK)
    if "header_frame" in globals():
        header_frame.configure(bg=BG_DARK)
    if "logo_label" in globals():
        logo_label.configure(bg=BG_DARK)
    if "title_frame" in globals():
        title_frame.configure(bg=BG_DARK)
    if "page_title_label" in globals():
        page_title_label.configure(bg=BG_DARK, fg=TEXT_PRIMARY)
    if "frame" in globals():
        frame.configure(bg=BG_DARK)
    if "style" in globals():
        style.configure('Horizontal.TProgressbar',
                        background=ACCENT,
                        troughcolor=BG_INPUT,
                        bordercolor=BG_DARK,
                        lightcolor=ACCENT,
                        darkcolor=ACCENT)

def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
    return os.path.join(base_path, relative_path)

def _default_output_dir():
    return os.path.join(os.path.expanduser("~"), "Desktop", "innlestur")


def _ensure_folder(path, fallback):
    requested = (path or fallback or _default_output_dir()).strip()
    for candidate in (requested, fallback, _default_output_dir(), os.path.expanduser("~")):
        if not candidate:
            continue
        try:
            os.makedirs(candidate, exist_ok=True)
            if os.path.isdir(candidate):
                return candidate
        except Exception:
            continue
    return os.getcwd()


def _ensure_bank_output_dir(path=None):
    return _ensure_folder(path or settings.get("bank_output_dir"), _default_output_dir())


def _bank_output_dir():
    output_dir = _ensure_bank_output_dir()
    if settings.get("bank_output_dir") != output_dir:
        settings["bank_output_dir"] = output_dir
        save_settings()
    return output_dir


def _ensure_compressed_output_dir(path=None, bank_dir=None):
    fallback = bank_dir or settings.get("bank_output_dir") or _default_output_dir()
    return _ensure_folder(path or settings.get("compressed_output_dir"), fallback)


def _compressed_output_dir():
    output_dir = _ensure_compressed_output_dir()
    if settings.get("compressed_output_dir") != output_dir:
        settings["compressed_output_dir"] = output_dir
        save_settings()
    return output_dir


def _default_auto_code_training_dir():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "trainingCoded")


def _default_auto_code_key_dir():
    training_keys = os.path.join(_default_auto_code_training_dir(), "coded keys")
    if os.path.isdir(training_keys):
        return training_keys
    training_keys_alt = os.path.join(_default_auto_code_training_dir(), "keys")
    if os.path.isdir(training_keys_alt):
        return training_keys_alt
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "coded keys")


def _auto_code_training_dir():
    return settings.get("auto_code_training_dir") or _default_auto_code_training_dir()


def _auto_code_key_dir():
    training_dir = _auto_code_training_dir()
    for folder_name in ("coded keys", "keys", "codedKeys", "key mappings"):
        candidate = os.path.join(training_dir, folder_name)
        if os.path.isdir(candidate):
            return candidate
    saved_key_dir = settings.get("auto_code_key_dir")
    if saved_key_dir and os.path.isdir(saved_key_dir):
        return saved_key_dir
    return _default_auto_code_key_dir()


def _auto_code_memory_path():
    return os.path.join(_auto_code_training_dir(), "auto_code_memory.json")

def _clean_file_list(paths):
    cleaned = []
    seen = set()
    if not isinstance(paths, list):
        return cleaned
    for path in paths:
        if not isinstance(path, str) or not path.strip():
            continue
        try:
            normalized = os.path.normpath(os.path.abspath(path.strip()))
        except Exception:
            continue
        key = os.path.normcase(normalized)
        if key in seen or not os.path.exists(normalized):
            continue
        seen.add(key)
        cleaned.append(normalized)
        if len(cleaned) >= RECENT_FILE_LIMIT:
            break
    return cleaned

def get_recent_files(setting_key):
    recent = _clean_file_list(settings.get(setting_key, []))
    if settings.get(setting_key) != recent:
        settings[setting_key] = recent
        save_settings()
    return recent

def remember_recent_file(setting_key, path):
    if not path or not os.path.exists(path):
        return
    normalized = os.path.normpath(os.path.abspath(path))
    existing = [p for p in get_recent_files(setting_key) if os.path.normcase(p) != os.path.normcase(normalized)]
    settings[setting_key] = [normalized] + existing[:RECENT_FILE_LIMIT - 1]
    save_settings()

def _short_file_name(path, limit=22):
    name = os.path.basename(path) if path else ""
    if len(name) <= limit:
        return name
    keep_end = max(8, limit // 2)
    keep_start = max(6, limit - keep_end - 3)
    return f"{name[:keep_start]}...{name[-keep_end:]}"

def _format_file_size(bytes_count):
    try:
        size = float(bytes_count)
    except Exception:
        return ""
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
    return ""

def _file_loaded_detail(path):
    if not path:
        return ""
    if not os.path.exists(path):
        return "File not found."
    try:
        modified = datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%d.%m.%Y %H:%M")
        size = _format_file_size(os.path.getsize(path))
        return f"Loaded saved version: {modified}" + (f" · {size}" if size else "")
    except Exception:
        return ""

def _status_with_file_detail(message, path):
    detail = _file_loaded_detail(path)
    return f"{message}\n{detail}" if detail else message

def _short_log_value(value, limit=90):
    text = "" if value is None else str(value)
    text = text.replace("\r", "\\r").replace("\n", "\\n")
    if len(text) <= limit:
        return text
    return f"{text[:limit - 3]}..."

def write_tok_debug_log(message):
    try:
        os.makedirs(os.path.dirname(DEBUG_LOG_PATH), exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        with open(DEBUG_LOG_PATH, "a", encoding="utf-8") as handle:
            handle.write(f"{stamp} {message}\n")
    except Exception:
        pass

def open_tok_debug_log():
    try:
        if not os.path.exists(DEBUG_LOG_PATH):
            write_tok_debug_log("Debug log created.")
        os.startfile(DEBUG_LOG_PATH)
    except Exception:
        pass

def load_settings():
    defaults = {
        "action_delay": 0.1,
        "focus_delay": 0.05,
        "paste_delay": 0.0,
        "post_paste_delay": 0.05,
        "start_delay": 3,
        "bank_output_dir": _default_output_dir(),
        "compressed_output_dir": "",
        "auto_code_training_dir": _default_auto_code_training_dir(),
        "auto_code_key_dir": _default_auto_code_key_dir(),
        "bank_restaurant_mode": False,
        "theme": "light",
        "recent_tok_files": [],
        "recent_bank_files": []
    }
    try:
        if os.path.exists(SETTINGS_PATH):
            with open(SETTINGS_PATH, "r", encoding="utf-8") as handle:
                data = json.load(handle)
            defaults.update(data if isinstance(data, dict) else {})
    except Exception:
        pass
    try:
        defaults["action_delay"] = float(defaults["action_delay"])
    except Exception:
        defaults["action_delay"] = 0.1
    try:
        defaults["focus_delay"] = float(defaults.get("focus_delay", 0.02))
    except Exception:
        defaults["focus_delay"] = 0.02
    try:
        defaults["paste_delay"] = float(defaults.get("paste_delay", 0.0))
    except Exception:
        defaults["paste_delay"] = 0.0
    try:
        defaults["post_paste_delay"] = float(defaults.get("post_paste_delay", 0.02))
    except Exception:
        defaults["post_paste_delay"] = 0.02
    try:
        defaults["start_delay"] = int(defaults["start_delay"])
    except Exception:
        defaults["start_delay"] = 3
    if not defaults.get("bank_output_dir"):
        defaults["bank_output_dir"] = _default_output_dir()
    defaults["bank_output_dir"] = _ensure_bank_output_dir(defaults.get("bank_output_dir"))
    if not defaults.get("compressed_output_dir"):
        defaults["compressed_output_dir"] = defaults["bank_output_dir"]
    defaults["compressed_output_dir"] = _ensure_compressed_output_dir(
        defaults.get("compressed_output_dir"),
        bank_dir=defaults.get("bank_output_dir"),
    )
    if not defaults.get("auto_code_training_dir"):
        defaults["auto_code_training_dir"] = _default_auto_code_training_dir()
    if not defaults.get("auto_code_key_dir"):
        defaults["auto_code_key_dir"] = _default_auto_code_key_dir()
    defaults["bank_restaurant_mode"] = bool(defaults.get("bank_restaurant_mode", False))
    if defaults.get("theme") not in THEMES:
        defaults["theme"] = "light"
    defaults["recent_tok_files"] = _clean_file_list(defaults.get("recent_tok_files", []))
    defaults["recent_bank_files"] = _clean_file_list(defaults.get("recent_bank_files", []))
    return defaults

def save_settings():
    try:
        os.makedirs(os.path.dirname(SETTINGS_PATH), exist_ok=True)
        with open(SETTINGS_PATH, "w", encoding="utf-8") as handle:
            json.dump(settings, handle, indent=2)
    except Exception:
        pass

def _after_ms(seconds):
    # Tkinter expects milliseconds. Always schedule at least 1ms.
    try:
        return max(1, int(float(seconds) * 1000))
    except Exception:
        return 1

def _cancel_scheduled_after():
    global after_id
    if after_id is not None:
        try:
            root.after_cancel(after_id)
        except Exception:
            pass
        after_id = None

def _cancel_watchdog():
    global watchdog_timer
    if watchdog_timer:
        try:
            watchdog_timer.cancel()
        except Exception:
            pass
    watchdog_timer = None

def _mark_automation_stopped(reason=""):
    global running, actions, automation_stop_reason
    running = False
    actions = []
    automation_stop_reason = reason

def _mouse_in_main_top_left():
    try:
        x, y = pyautogui.position()
        return 0 <= x <= MOUSE_FAILSAFE_MARGIN and 0 <= y <= MOUSE_FAILSAFE_MARGIN
    except Exception:
        return False

def _stop_for_mouse_failsafe():
    _mark_automation_stopped("Stopped because the mouse was moved to the top-left of the main screen.")
    _cancel_scheduled_after()
    _cancel_watchdog()

def _arm_watchdog():
    global watchdog_timer
    _cancel_watchdog()
    watchdog_timer = Timer(AUTOMATION_WATCHDOG_SECONDS, on_watchdog_timeout)
    watchdog_timer.daemon = True
    watchdog_timer.start()

def _sync_settings_from_ui():
    global action_delay, start_delay, focus_delay, paste_delay, post_paste_delay
    if "settings_theme_var" in globals():
        theme = settings_theme_var.get()
        if theme in THEMES:
            settings["theme"] = theme
    if "settings_action_entry" in globals() and settings_action_entry and settings_action_entry.winfo_exists():
        try:
            action_delay = float(settings_action_entry.get())
        except Exception:
            pass
        settings["action_delay"] = action_delay
    if "settings_focus_entry" in globals() and settings_focus_entry and settings_focus_entry.winfo_exists():
        try:
            focus_delay = float(settings_focus_entry.get())
        except Exception:
            pass
        settings["focus_delay"] = focus_delay
    if "settings_paste_entry" in globals() and settings_paste_entry and settings_paste_entry.winfo_exists():
        try:
            paste_delay = float(settings_paste_entry.get())
        except Exception:
            pass
        settings["paste_delay"] = paste_delay
    if "settings_post_paste_entry" in globals() and settings_post_paste_entry and settings_post_paste_entry.winfo_exists():
        try:
            post_paste_delay = float(settings_post_paste_entry.get())
        except Exception:
            pass
        settings["post_paste_delay"] = post_paste_delay
    if "settings_start_entry" in globals() and settings_start_entry and settings_start_entry.winfo_exists():
        try:
            start_delay = int(float(settings_start_entry.get()))
        except Exception:
            pass
        settings["start_delay"] = start_delay
    if "settings_output_dir_entry" in globals() and settings_output_dir_entry and settings_output_dir_entry.winfo_exists():
        output_dir = settings_output_dir_entry.get().strip()
        if output_dir:
            settings["bank_output_dir"] = output_dir
    if "settings_compressed_output_dir_entry" in globals() and settings_compressed_output_dir_entry and settings_compressed_output_dir_entry.winfo_exists():
        compressed_output_dir = settings_compressed_output_dir_entry.get().strip()
        if compressed_output_dir:
            settings["compressed_output_dir"] = compressed_output_dir
    if "settings_auto_code_training_entry" in globals() and settings_auto_code_training_entry and settings_auto_code_training_entry.winfo_exists():
        training_dir = settings_auto_code_training_entry.get().strip()
        if training_dir:
            settings["auto_code_training_dir"] = training_dir
    save_settings()

def _set_runtime_delays(values):
    global action_delay, focus_delay, paste_delay, post_paste_delay
    action_delay = float(values.get("action", action_delay))
    focus_delay = float(values.get("focus", focus_delay))
    paste_delay = float(values.get("paste", paste_delay))
    post_paste_delay = float(values.get("post_paste", post_paste_delay))

def apply_tok_run_speed_selection():
    global action_delay, focus_delay, paste_delay, post_paste_delay, current_run_speed_label
    selected = "Saved"
    if "tok_run_speed_var" in globals() and tok_run_speed_var:
        selected = tok_run_speed_var.get() or "Saved"
    current_run_speed_label = selected
    if selected in SPEED_PRESET_MAP:
        _set_runtime_delays(SPEED_PRESET_MAP[selected])
    else:
        action_delay = float(settings.get("action_delay", action_delay))
        focus_delay = float(settings.get("focus_delay", focus_delay))
        paste_delay = float(settings.get("paste_delay", paste_delay))
        post_paste_delay = float(settings.get("post_paste_delay", post_paste_delay))

def ensure_window_fits():
    root.update_idletasks()
    req_w = root.winfo_reqwidth()
    req_h = root.winfo_reqheight()
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    target_w = min(max(req_w, 620), screen_w - 40)
    target_h = min(max(req_h, 560), screen_h - 60)
    move_window_onscreen(target_w, target_h)

def fit_window_to_content():
    root.update_idletasks()
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    target_w = min(max(root.winfo_reqwidth(), 620), screen_w - 40)
    target_h = min(max(root.winfo_reqheight(), 560), screen_h - 60)
    move_window_onscreen(target_w, target_h)

def move_window_onscreen(width=None, height=None):
    root.update_idletasks()
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    width = int(width if width is not None else root.winfo_width())
    height = int(height if height is not None else root.winfo_height())
    width = min(width, screen_w - 40)
    height = min(height, screen_h - 60)

    x = root.winfo_x()
    y = root.winfo_y()
    if y < 0 or y + height > screen_h - 40:
        y = max(0, min(y, screen_h - height - 40))

    root.geometry(f"{width}x{height}+{int(x)}+{int(y)}")

def fit_dialog_to_content(dialog, min_width=560, min_height=420, preferred_width=None):
    dialog.update_idletasks()
    screen_w = dialog.winfo_screenwidth()
    screen_h = dialog.winfo_screenheight()
    target_w = min(max(dialog.winfo_reqwidth(), min_width, preferred_width or 0), screen_w - 80)
    target_h = min(max(dialog.winfo_reqheight(), min_height), screen_h - 120)
    dialog.minsize(min_width, min_height)

    try:
        parent_x = root.winfo_rootx()
        parent_y = root.winfo_rooty()
        parent_w = root.winfo_width()
        parent_h = root.winfo_height()
        x = parent_x + (parent_w - target_w) // 2
        y = parent_y + (parent_h - target_h) // 2
    except Exception:
        x = (screen_w - target_w) // 2
        y = (screen_h - target_h) // 2

    x = max(20, min(int(x), screen_w - int(target_w) - 20))
    y = max(20, min(int(y), screen_h - int(target_h) - 60))
    dialog.geometry(f"{int(target_w)}x{int(target_h)}+{x}+{y}")

def finalize_page_layout():
    root.after_idle(fit_window_to_content)

def schedule_layout_refresh():
    if "root" in globals():
        root.after_idle(fit_window_to_content)
        try:
            root.after(80, fit_window_to_content)
            root.after(180, fit_window_to_content)
        except Exception:
            pass

def set_page_title(title, color=None):
    if "page_title_label" in globals():
        page_title_label.configure(text=title, fg=color or TEXT_PRIMARY)

def _stop_automation_only():
    """Stop any currently running auto-input as fast as possible (non-blocking)."""
    _mark_automation_stopped()
    _cancel_scheduled_after()
    _cancel_watchdog()

def stop_script():
    _stop_automation_only()
    logging.info("Stopping...")
    display_stopped_screen()

def _legacy_display_stopped_screen(reason=""):
    """Legacy stopped screen kept only as a fallback reference."""
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Script Stopped", TEXT_ERROR)
    
    # Calculate runtime
    end_time = datetime.datetime.now()
    time_elapsed = end_time - start_time if 'start_time' in globals() and start_time else datetime.timedelta(0)
    
    # Get current row info
    current_row_num = row_index + 1 if 'row_index' in globals() else 0
    if 'rows' in globals() and rows:
        total_rows_count = len(rows)
    elif pending_tok_df is not None:
        total_rows_count = len(pending_tok_df)
    else:
        total_rows_count = 0
    
    if reason:
        reason_label = create_styled_label(frame, reason, size=9, color=TEXT_WARNING)
        reason_label.pack(pady=(0, 8))
    
    # Info card
    card = tk.Frame(frame, bg=BG_CARD, padx=20, pady=15)
    card.pack(fill=tk.X, padx=20, pady=10)
    
    # Row info
    row_info_label = tk.Label(card, text=f"Stopped at Row: {current_row_num} / {total_rows_count}",
                              bg=BG_CARD, fg=TEXT_PRIMARY, font=('Segoe UI', 11, 'bold'))
    row_info_label.pack(pady=5)
    
    # Show current row details if available (truncate long text)
    if 'rows' in globals() and rows and row_index < len(rows):
        current_row = rows[row_index]
        date_val = str(current_row.get('DATE', 'N/A'))
        text_val = str(current_row.get('TEXT', 'N/A'))[:30] + ('...' if len(str(current_row.get('TEXT', ''))) > 30 else '')
        amount_val = str(current_row.get('AMOUNT', 'N/A'))
        
        details_frame = tk.Frame(card, bg=BG_CARD)
        details_frame.pack(pady=10)
        
        for label, value in [("DATE", date_val), ("TEXT", text_val), ("AMOUNT", amount_val)]:
            row_frame = tk.Frame(details_frame, bg=BG_CARD)
            row_frame.pack(fill=tk.X, pady=2)
            tk.Label(row_frame, text=f"{label}:", bg=BG_CARD, fg=TEXT_SECONDARY,
                    font=('Segoe UI', 9), width=8, anchor='e').pack(side=tk.LEFT)
            tk.Label(row_frame, text=value, bg=BG_CARD, fg=TEXT_PRIMARY,
                    font=('Segoe UI', 9)).pack(side=tk.LEFT, padx=5)
    
    # Stats card
    stats_card = tk.Frame(frame, bg=BG_CARD, padx=20, pady=15)
    stats_card.pack(fill=tk.X, padx=20, pady=10)
    
    stats_frame = tk.Frame(stats_card, bg=BG_CARD)
    stats_frame.pack()
    
    # Runtime
    tk.Label(stats_frame, text="⏱ Runtime:", bg=BG_CARD, fg=TEXT_SECONDARY,
             font=('Segoe UI', 10)).pack(side=tk.LEFT)
    tk.Label(stats_frame, text=format_timedelta(time_elapsed), bg=BG_CARD, fg=TEXT_PRIMARY,
             font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=(5, 20))
    
    # Rows processed
    tk.Label(stats_frame, text="📊 Rows:", bg=BG_CARD, fg=TEXT_SECONDARY,
             font=('Segoe UI', 10)).pack(side=tk.LEFT)
    tk.Label(stats_frame, text=str(row_index if 'row_index' in globals() else 0), bg=BG_CARD, fg=TEXT_PRIMARY,
             font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=5)
    
    if current_tok_file_path:
        tok_input_btn = create_styled_button(
            frame,
            "Keyra aftur inn",
            lambda path=current_tok_file_path: open_tok_input_with_file(path),
            width=20,
            accent=False
        )
        tok_input_btn.pack(pady=(5, 8))

        open_btn = create_styled_button(frame, "Open File", lambda: os.startfile(current_tok_file_path), width=20, accent=False)
        open_btn.pack(pady=(0, 10))

    action_row = tk.Frame(frame, bg=BG_DARK)
    action_row.pack(fill=tk.X, padx=28, pady=(0, 10))
    log_button = create_styled_button(action_row, "Log", open_tok_debug_log, width=16, accent=False)
    log_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
    back_button = create_styled_button(action_row, "Back to Menu", initialize_main_menu, width=16)
    back_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
    back_button.pack(pady=20)
    finalize_page_layout()

def display_stopped_screen(reason=""):
    """Show a recovery-focused summary when the script is stopped."""
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Script Stopped", TEXT_ERROR)

    end_time = datetime.datetime.now()
    time_elapsed = end_time - start_time if 'start_time' in globals() and start_time else datetime.timedelta(0)
    if 'rows' in globals() and rows:
        total_rows_count = len(rows)
    elif pending_tok_df is not None:
        total_rows_count = len(pending_tok_df)
    else:
        total_rows_count = 0

    current_index = row_index if 'row_index' in globals() else 0
    current_index = max(0, min(current_index, max(total_rows_count - 1, 0))) if total_rows_count else 0
    stopped_row = current_index + 1 if total_rows_count else 0
    rows_this_run = max(0, current_index - current_tok_run_start_index)
    write_tok_debug_log(
        f"RUN_STOP stopped_row={stopped_row} rows_this_run={rows_this_run} "
        f"reason='{_short_log_value(reason)}'"
    )

    if reason:
        reason_label = create_styled_label(frame, reason, size=9, color=TEXT_WARNING)
        reason_label.configure(wraplength=470, justify='center')
        reason_label.pack(pady=(0, 8))

    summary_shell, summary = create_panel(frame, padx=22, pady=18)
    summary_shell.pack(fill=tk.X, padx=28, pady=(4, 12))

    def add_info_row(parent, label, value, value_color=None, bold=False):
        row = tk.Frame(parent, bg=BG_CARD)
        row.pack(fill=tk.X, pady=3)
        tk.Label(row, text=label, bg=BG_CARD, fg=TEXT_SECONDARY,
                 font=('Segoe UI', 9), width=18, anchor='w').pack(side=tk.LEFT)
        tk.Label(row, text=value, bg=BG_CARD, fg=value_color or TEXT_PRIMARY,
                 font=('Segoe UI', 10, 'bold' if bold else 'normal'),
                 anchor='w').pack(side=tk.LEFT, fill=tk.X, expand=True)

    stopped_row_frame = tk.Frame(summary, bg=BG_CARD)
    stopped_row_frame.pack(fill=tk.X, pady=3)
    tk.Label(stopped_row_frame, text="Stopped at", bg=BG_CARD, fg=TEXT_SECONDARY,
             font=('Segoe UI', 9), width=18, anchor='w').pack(side=tk.LEFT)
    tk.Label(stopped_row_frame, text=f"Row {stopped_row} / {total_rows_count}", bg=BG_CARD, fg=TEXT_ERROR,
             font=('Segoe UI', 10, 'bold'), anchor='w').pack(side=tk.LEFT, fill=tk.X, expand=True)
    log_icon = create_styled_button(stopped_row_frame, "Log", open_tok_debug_log, width=5, accent=False, height=30)
    log_icon.pack(side=tk.RIGHT)
    add_info_row(summary, "Rows this run", str(rows_this_run))
    add_info_row(summary, "Runtime", format_timedelta(time_elapsed))

    if current_tok_file_path:
        file_row = tk.Frame(frame, bg=BG_DARK)
        file_row.pack(fill=tk.X, padx=28, pady=(0, 8))
        tok_input_btn = create_styled_button(file_row, "Keyra aftur inn", lambda path=current_tok_file_path: open_tok_input_with_file(path), width=16, accent=False)
        tok_input_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        open_btn = create_styled_button(file_row, "Open File", lambda: os.startfile(current_tok_file_path), width=16, accent=False)
        open_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))

    back_button = create_styled_button(frame, "Back to Menu", initialize_main_menu, width=20)
    back_button.pack(pady=(8, 14))
    finalize_page_layout()

def format_number(number):
    number_str = str(number).split('.')[0]
    return number_str.replace(',', '')

def copy_to_clipboard_verified(text, field="", row_label="?"):
    expected = "" if text is None else str(text)
    last_value = ""
    for attempt in range(8):
        pyperclip.copy(expected)
        # Give Windows clipboard a tiny chance to settle on older machines.
        time.sleep(0.015 if attempt < 3 else 0.035)
        try:
            last_value = pyperclip.paste()
        except Exception:
            last_value = ""
        if last_value == expected:
            if attempt:
                write_tok_debug_log(f"CLIPBOARD_VERIFIED_AFTER_RETRY row={row_label} field={field} attempts={attempt + 1}")
            return
    write_tok_debug_log(
        f"CLIPBOARD_VERIFY_FAILED row={row_label} field={field} "
        f"expected='{_short_log_value(expected)}' actual='{_short_log_value(last_value)}'"
    )
    raise RuntimeError("Clipboard did not update to the expected value.")

def enter_data(row, row_number=None):
    date_str = row['DATE'].strftime('%Y-%m-%d') if isinstance(row['DATE'], pd.Timestamp) else row['DATE']
    actions = []
    field_delay = 0.03
    row_label = row_number if row_number is not None else "?"

    def _step(func, delay=USE_ACTION_DELAY, debug=""):
        return (func, delay, debug)

    def _press_enter(debug_label="ENTER"):
        pyautogui.press('enter')

    def _paste_text_steps(field, text):
        # Split into multiple steps to keep the UI responsive and allow Stop mid-row.
        text = "" if text is None else str(text)
        return [
            _step(lambda t=text, f=field: copy_to_clipboard_verified(t, f, row_label), paste_delay, f"row={row_label} field={field} copy value='{_short_log_value(text)}'"),
            _step(lambda: pyautogui.hotkey('ctrl', 'v'), post_paste_delay, f"row={row_label} field={field} paste"),
        ]

    actions.extend(_paste_text_steps("DATE", date_str))
    actions.append(_step(lambda: _press_enter("after DATE"), focus_delay, f"row={row_label} after DATE enter"))
    actions.append(_step(lambda: None, field_delay, f"row={row_label} after DATE wait"))

    actions.extend(_paste_text_steps("TEXT", str(row['TEXT'])))
    actions.append(_step(lambda: _press_enter("after TEXT 1"), focus_delay, f"row={row_label} after TEXT enter 1"))
    actions.append(_step(lambda: _press_enter("after TEXT 2"), focus_delay, f"row={row_label} after TEXT enter 2"))
    actions.append(_step(lambda: None, field_delay, f"row={row_label} after TEXT wait"))

    actions.extend(_paste_text_steps("DEBIT", format_number(row['DEBIT'])))
    actions.append(_step(lambda: _press_enter("after DEBIT"), focus_delay, f"row={row_label} after DEBIT enter"))
    actions.append(_step(lambda: None, field_delay, f"row={row_label} after DEBIT wait"))

    if pd.notna(row['ID']):
        actions.extend(_paste_text_steps("ID_DEBIT_SIDE", format_number(row['ID'])))
    actions.append(_step(lambda: _press_enter("after ID debit side"), focus_delay, f"row={row_label} after debit-side ID enter"))
    actions.append(_step(lambda: None, field_delay, f"row={row_label} after debit-side ID wait"))

    actions.extend(_paste_text_steps("AMOUNT", format_number(row['AMOUNT'])))
    actions.append(_step(lambda: _press_enter("after AMOUNT 1"), focus_delay, f"row={row_label} after AMOUNT enter 1"))
    actions.append(_step(lambda: _press_enter("after AMOUNT 2"), focus_delay, f"row={row_label} after AMOUNT enter 2"))
    actions.append(_step(lambda: _press_enter("after AMOUNT 3"), focus_delay, f"row={row_label} after AMOUNT enter 3"))
    actions.append(_step(lambda: _press_enter("after AMOUNT 4"), focus_delay, f"row={row_label} after AMOUNT enter 4"))
    actions.append(_step(lambda: None, field_delay, f"row={row_label} after AMOUNT wait"))
    actions.extend(_paste_text_steps("CREDIT", format_number(row['CREDIT'])))
    actions.append(_step(lambda: _press_enter("after CREDIT"), focus_delay, f"row={row_label} after CREDIT enter"))
    actions.append(_step(lambda: None, field_delay, f"row={row_label} after CREDIT wait"))

    if pd.notna(row['ID']):
        actions.extend(_paste_text_steps("ID_CREDIT_SIDE", format_number(row['ID'])))
    actions.append(_step(lambda: _press_enter("after credit-side ID 1"), focus_delay, f"row={row_label} after credit-side ID enter 1"))
    actions.append(_step(lambda: _press_enter("after credit-side ID 2"), focus_delay, f"row={row_label} after credit-side ID enter 2"))

    return actions

def init_progress_bar(total_rows):
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Processing")

    global progress_bar, progress_label, stop_button_processing

    progress_label = create_styled_label(frame, "", size=10, color=TEXT_SECONDARY)
    progress_label.pack(pady=10)

    # Progress bar with custom style
    progress_bar = ttk.Progressbar(frame, orient='horizontal', length=350, mode='determinate',
                                   style='Horizontal.TProgressbar')
    progress_bar.pack(pady=15)
    progress_bar['maximum'] = total_rows
    progress_bar['value'] = 0
    
    stop_button_processing = create_styled_button(frame, "⏹ Stop", stop_script, width=15, accent=False)
    stop_button_processing.pack(pady=20)
    finalize_page_layout()

def update_progress(row_count):
    progress_bar['value'] = row_count
    progress_label.config(text=f"Processing... {row_count}/{progress_bar['maximum']} rows")
    root.update_idletasks()

def display_results(time_elapsed, rows_processed):
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Completed", TEXT_SUCCESS)
    
    card_shell, card = create_panel(frame, padx=24, pady=18)
    card_shell.pack(fill=tk.X, padx=28, pady=(4, 14))
    
    tk.Label(card, text=f"⏱ Time: {format_timedelta(time_elapsed)}", bg=BG_CARD, fg=TEXT_PRIMARY,
             font=('Segoe UI', 11)).pack(pady=5)
    tk.Label(card, text=f"📊 Rows Processed: {rows_processed}", bg=BG_CARD, fg=TEXT_PRIMARY,
             font=('Segoe UI', 11)).pack(pady=5)
    
    play_success_sound()

    if current_tok_file_path:
        file_row = tk.Frame(frame, bg=BG_DARK)
        file_row.pack(fill=tk.X, padx=28, pady=(0, 8))
        rerun_button = create_styled_button(
            file_row,
            "Keyra aftur inn",
            lambda path=current_tok_file_path: open_tok_input_with_file(path),
            width=16,
            accent=False,
        )
        rerun_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        open_button = create_styled_button(
            file_row,
            "Open File",
            lambda path=current_tok_file_path: os.startfile(path),
            width=16,
            accent=False,
        )
        open_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))

    action_row = tk.Frame(frame, bg=BG_DARK)
    action_row.pack(fill=tk.X, padx=28, pady=(0, 10))
    log_button = create_styled_button(action_row, "Log", open_tok_debug_log, width=16, accent=False)
    log_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
    back_button = create_styled_button(action_row, "Back to Menu", initialize_main_menu, width=16)
    back_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
    finalize_page_layout()

def format_timedelta(td):
    minutes, seconds = divmod(td.total_seconds(), 60)
    hours, minutes = divmod(minutes, 60)
    return f"{int(hours):02}:{int(minutes):02}:{seconds:.2f}"

TOK_REQUIRED_COLUMNS = ['DATE', 'TEXT', 'DEBIT', 'ID', 'AMOUNT', 'CREDIT']

def _format_missing_rows(missing_rows):
    if len(missing_rows) > 10:
        return ", ".join(str(r) for r in missing_rows[:10]) + f" (+{len(missing_rows) - 10} more)"
    return ", ".join(str(r) for r in missing_rows)

def _parse_tok_decimal(value):
    if pd.isna(value):
        return None
    if isinstance(value, str):
        cleaned = value.replace(" ", "").replace(",", ".")
    else:
        cleaned = str(value)
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None

def _is_missing_tok_value(val):
    if pd.isna(val):
        return True
    if isinstance(val, str) and not val.strip():
        return True
    return False

def _is_tok_stop_row(row):
    date_val = row.get('DATE')
    if pd.isna(date_val):
        return False
    return str(date_val).strip().lower() == 'xx'

def _tok_row_has_input_data(row):
    if _is_tok_stop_row(row):
        for col in TOK_REQUIRED_COLUMNS:
            if col == 'DATE':
                continue
            val = row.get(col)
            if pd.isna(val):
                continue
            if isinstance(val, str) and not val.strip():
                continue
            return True
        return False
    for col in TOK_REQUIRED_COLUMNS:
        val = row.get(col)
        if pd.isna(val):
            continue
        if isinstance(val, str) and not val.strip():
            continue
        return True
    return False

def _extract_tok_date_year(date_val):
    if pd.isna(date_val):
        return None
    text = str(date_val).strip()
    match = re.search(r'(\d{4})\s*$', text)
    return match.group(1) if match else None

def _validate_tok_single_date_year(check_df):
    rows_by_year = {}
    invalid_rows = []
    for idx, row in check_df.iterrows():
        if not _tok_row_has_input_data(row):
            continue
        year = _extract_tok_date_year(row.get('DATE'))
        excel_row = idx + 2  # Excel row number (header is row 1)
        if not year:
            invalid_rows.append(excel_row)
            continue
        rows_by_year.setdefault(year, []).append(excel_row)

    if invalid_rows:
        return (
            False,
            f"DATE must include a year on row(s): {_format_missing_rows(invalid_rows)}"
        )
    if len(rows_by_year) > 1:
        details = "; ".join(
            f"{year}: row(s) {_format_missing_rows(row_numbers)}"
            for year, row_numbers in sorted(rows_by_year.items())
        )
        return False, f"Mixed DATE years found. Use one year per file. {details}"
    return True, ""

def prepare_tok_input_file(file_path):
    if not file_path:
        return {'ok': False, 'message': 'Select an Excel file first.', 'df': None, 'missing_rows': []}

    try:
        df = pd.read_excel(file_path)
    except Exception as exc:
        return {'ok': False, 'message': f"Failed to read file: {str(exc)[:80]}", 'df': None, 'missing_rows': []}

    df = df.dropna(how='all')
    if df.empty:
        return {'ok': False, 'message': 'No rows to process.', 'df': None, 'missing_rows': []}

    missing_cols = [col for col in TOK_REQUIRED_COLUMNS if col not in df.columns]
    if missing_cols:
        return {'ok': False, 'message': f"Missing column(s): {', '.join(missing_cols)}", 'df': None, 'missing_rows': []}

    df = df.copy()
    has_decimals = False
    for val in df['AMOUNT']:
        dec = _parse_tok_decimal(val)
        if dec is not None and dec != dec.to_integral_value():
            has_decimals = True
            break
    if has_decimals:
        def _round_half_up(val):
            dec = _parse_tok_decimal(val)
            if dec is None:
                return val
            return int(dec.quantize(Decimal("1"), rounding=ROUND_HALF_UP))
        df['AMOUNT'] = df['AMOUNT'].apply(_round_half_up)

    df['DATE'] = df['DATE'].apply(format_date_as_text)

    check_df = df[df.apply(_tok_row_has_input_data, axis=1)].copy()
    if check_df.empty:
        return {'ok': False, 'message': 'No rows to process.', 'df': None, 'missing_rows': []}

    dates_ok, date_message = _validate_tok_single_date_year(check_df)
    if not dates_ok:
        return {'ok': False, 'message': date_message, 'df': None, 'missing_rows': []}

    missing_rows = []
    for idx, row in check_df.iterrows():
        if _tok_row_has_input_data(row) and (_is_missing_tok_value(row.get('DEBIT')) or _is_missing_tok_value(row.get('CREDIT'))):
            missing_rows.append(idx + 2)  # Excel row number (header is row 1)

    if missing_rows:
        return {
            'ok': True,
            'message': f"Compatible, but missing DEBIT/CREDIT on row(s): {_format_missing_rows(missing_rows)}",
            'df': check_df,
            'missing_rows': missing_rows,
            'warning': True
        }

    ready_rows = len(check_df)
    return {'ok': True, 'message': f"Compatible. {ready_rows} row(s) ready.", 'df': check_df, 'missing_rows': [], 'warning': False}

def update_tok_input_status(file_path):
    global pending_tok_df, pending_tok_file_path
    result = prepare_tok_input_file(file_path)
    if result['ok']:
        pending_tok_df = result['df']
        pending_tok_file_path = file_path
        color = TEXT_WARNING if result.get('warning') else TEXT_SUCCESS
    else:
        pending_tok_df = None
        pending_tok_file_path = None
        color = TEXT_ERROR

    if "tok_status_label" in globals() and tok_status_label and tok_status_label.winfo_exists():
        tok_status_label.config(text=_status_with_file_detail(result['message'], file_path), fg=color)
    if "tok_open_file_button" in globals() and tok_open_file_button and tok_open_file_button.winfo_exists():
        if file_path and os.path.exists(file_path):
            tok_open_file_button.pack(side=tk.RIGHT, padx=(8, 8), pady=4)
        else:
            tok_open_file_button.pack_forget()
    schedule_layout_refresh()
    return result

def refresh_tok_recent_buttons():
    if "tok_recent_files_frame" not in globals() or not tok_recent_files_frame.winfo_exists():
        return
    for widget in tok_recent_files_frame.winfo_children():
        widget.destroy()
    recent = get_recent_files("recent_tok_files")[:RECENT_VISIBLE_LIMIT]
    if not recent:
        return
    label = create_styled_label(tok_recent_files_frame, "Recent files", size=9, color=TEXT_SECONDARY, bg=BG_CARD)
    label.pack(anchor='w', pady=(0, 4))
    row = tk.Frame(tok_recent_files_frame, bg=BG_CARD)
    row.pack(fill=tk.X)
    for index, path in enumerate(recent):
        btn = create_styled_button(
            row,
            _short_file_name(path, 19),
            lambda p=path: select_tok_input_file(p),
            width=13,
            accent=False,
            height=34,
        )
        btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6) if index < len(recent) - 1 else (0, 0))
    schedule_layout_refresh()

def select_tok_input_file(file_path, remember=True, reset_start=True):
    if not file_path:
        return None
    path = os.path.normpath(os.path.abspath(file_path))
    if "file_path_entry" in globals() and file_path_entry and file_path_entry.winfo_exists():
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, path)
    if "tok_drop_area" in globals() and tok_drop_area and tok_drop_area.winfo_exists():
        tok_drop_area.config(text=os.path.basename(path), fg=TEXT_PRIMARY)
    result = update_tok_input_status(path)
    if reset_start:
        reset_tok_start_row_controls()
    if remember and os.path.exists(path):
        remember_recent_file("recent_tok_files", path)
        refresh_tok_recent_buttons()
    schedule_layout_refresh()
    return result

def clear_tok_input_selection():
    global pending_tok_df, pending_tok_file_path
    pending_tok_df = None
    pending_tok_file_path = None
    if "file_path_entry" in globals() and file_path_entry and file_path_entry.winfo_exists():
        file_path_entry.delete(0, tk.END)
    if "tok_drop_area" in globals() and tok_drop_area and tok_drop_area.winfo_exists():
        tok_drop_area.config(text="Drop Tok input file here or click to browse", fg=TEXT_SECONDARY)
    if "tok_status_label" in globals() and tok_status_label and tok_status_label.winfo_exists():
        tok_status_label.config(text="Select a Tok input Excel file to check compatibility.", fg=TEXT_SECONDARY)
    if "tok_open_file_button" in globals() and tok_open_file_button and tok_open_file_button.winfo_exists():
        tok_open_file_button.pack_forget()
    reset_tok_start_row_controls()
    schedule_layout_refresh()


def reset_tok_start_row_controls():
    if "tok_start_mode_var" in globals() and tok_start_mode_var:
        try:
            tok_start_mode_var.set("first")
        except Exception:
            pass
    if "tok_start_row_entry" in globals() and tok_start_row_entry and tok_start_row_entry.winfo_exists():
        tok_start_row_entry.delete(0, tk.END)
        tok_start_row_entry.insert(0, "1")
    if "refresh_tok_start_row_controls" in globals():
        try:
            refresh_tok_start_row_controls()
        except Exception:
            pass


def _tok_selected_start_excel_row():
    if "tok_start_mode_var" not in globals() or tok_start_mode_var.get() != "custom":
        return 1
    raw = ""
    if "tok_start_row_entry" in globals() and tok_start_row_entry and tok_start_row_entry.winfo_exists():
        raw = tok_start_row_entry.get().strip()
    if not raw:
        return 1
    try:
        row_number = int(float(raw))
    except Exception:
        if "tok_status_label" in globals() and tok_status_label and tok_status_label.winfo_exists():
            tok_status_label.config(text="Start row must be a number.", fg=TEXT_WARNING)
        return None
    if row_number < 1:
        if "tok_status_label" in globals() and tok_status_label and tok_status_label.winfo_exists():
            tok_status_label.config(text="Start row must be 1 or higher.", fg=TEXT_WARNING)
        return None
    return row_number


def _tok_start_index_from_excel_row(df, excel_row):
    if df is None or len(df) == 0:
        return None
    try:
        excel_row = int(excel_row)
    except Exception:
        return None
    if excel_row <= 2:
        return 0
    available_rows = []
    for position, index_value in enumerate(df.index):
        try:
            row_number = int(index_value) + 2
        except Exception:
            row_number = position + 2
        available_rows.append((position, row_number))
        if row_number >= excel_row:
            return position
    last_row = available_rows[-1][1] if available_rows else 1
    if "tok_status_label" in globals() and tok_status_label and tok_status_label.winfo_exists():
        tok_status_label.config(text=f"Start row {excel_row} is after the last usable row ({last_row}).", fg=TEXT_WARNING)
    return None

def display_input_controls():
    global file_path_entry, browse_button, run_button, refresh_button, tok_drop_area, tok_status_label, tok_open_file_button
    global input_controls_frame, tok_recent_files_frame, tok_run_speed_var, tok_speed_buttons, tok_test_run_button
    global tok_start_mode_var, tok_start_row_entry, tok_start_buttons, tok_start_custom_frame, refresh_tok_start_row_controls

    input_controls_frame, inner_controls = create_panel(frame, padx=22, pady=18)
    input_controls_frame.pack(fill=tk.X, padx=28, pady=(0, 12))

    file_frame = tk.Frame(inner_controls, bg=BG_CARD)
    file_frame.pack(fill=tk.X, pady=(2, 8))
    
    drop_area = create_drop_box(file_frame, "Drop Tok input file here or click to browse", height=5)
    drop_area.pack(fill=tk.X)
    tok_drop_area = drop_area
    
    file_path_entry = create_styled_entry(file_frame, width=35)
    file_path_entry.pack_forget()

    tok_recent_files_frame = tk.Frame(inner_controls, bg=BG_CARD)
    tok_recent_files_frame.pack(fill=tk.X, pady=(0, 8))

    status_shell = RoundedPanel(inner_controls, fill=BUTTON_MUTED, outline=BORDER, radius=10, padx=0, pady=0)
    status_row = status_shell.inner
    status_shell.pack(fill=tk.X, pady=(0, 10))
    tok_status_label = tk.Label(
        status_row,
        text="Select a Tok input Excel file to check compatibility.",
        bg=BUTTON_MUTED,
        fg=TEXT_SECONDARY,
        font=('Segoe UI', 9),
        padx=14,
        pady=7,
        anchor='w',
        justify='left',
    )
    tok_status_label.configure(wraplength=360, justify='left')
    tok_status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
    tok_open_file_button = create_styled_button(
        status_row,
        "Open File",
        lambda: os.startfile(file_path_entry.get().strip()) if file_path_entry.get().strip() else None,
        width=10,
        accent=False,
        height=42,
    )

    options_label = create_styled_label(inner_controls, "Run options", size=10, color=TEXT_PRIMARY, bold=True, bg=BG_CARD)
    options_label.pack(anchor='w', pady=(0, 6))

    speed_label = create_styled_label(inner_controls, "Speed for this run", size=9, color=TEXT_SECONDARY, bg=BG_CARD)
    speed_label.pack(anchor='w', pady=(0, 6))
    speed_row = tk.Frame(inner_controls, bg=BG_CARD)
    speed_row.pack(fill=tk.X, pady=(0, 10))
    tok_run_speed_var = tk.StringVar(value="Saved")
    tok_speed_buttons = {}

    def refresh_speed_buttons():
        selected = tok_run_speed_var.get()
        for label, button in tok_speed_buttons.items():
            set_button_accent(button, label == selected)
        if "tok_test_run_button" in globals() and tok_test_run_button and tok_test_run_button.winfo_exists():
            set_button_accent(tok_test_run_button, selected == "Test Run")

    def set_run_speed(label):
        tok_run_speed_var.set(label)
        refresh_speed_buttons()

    speed_labels = ["Saved"] + [label for label, _values in SPEED_PRESETS]
    for index, label in enumerate(speed_labels):
        btn = create_styled_button(speed_row, label, lambda value=label: set_run_speed(value), width=9, accent=False, height=40)
        tok_speed_buttons[label] = btn
        btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6) if index < len(speed_labels) - 1 else (0, 0))
    refresh_speed_buttons()

    start_label = create_styled_label(inner_controls, "Start row", size=9, color=TEXT_SECONDARY, bg=BG_CARD)
    start_label.pack(anchor='w', pady=(0, 6))
    start_row = tk.Frame(inner_controls, bg=BG_CARD)
    start_row.pack(fill=tk.X, pady=(0, 10))
    tok_start_mode_var = tk.StringVar(value="first")
    tok_start_buttons = {}

    start_button_row = tk.Frame(start_row, bg=BG_CARD)
    start_button_row.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    tok_start_buttons["first"] = create_styled_button(
        start_button_row,
        "Start at 1",
        lambda: set_tok_start_mode("first"),
        width=12,
        accent=False,
        height=38,
    )
    tok_start_buttons["first"].pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
    tok_start_buttons["custom"] = create_styled_button(
        start_button_row,
        "Custom",
        lambda: set_tok_start_mode("custom"),
        width=12,
        accent=False,
        height=38,
    )
    tok_start_buttons["custom"].pack(side=tk.LEFT, fill=tk.X, expand=True)

    tok_start_custom_frame = tk.Frame(start_row, bg=BG_CARD)
    tok_start_row_entry = create_styled_entry(tok_start_custom_frame, width=8)
    tok_start_row_entry.insert(0, "1")
    tok_start_row_entry.pack(side=tk.LEFT, ipady=4)

    def refresh_tok_start_row_controls():
        selected = tok_start_mode_var.get()
        for value, button in tok_start_buttons.items():
            set_button_accent(button, value == selected)
        if selected == "custom":
            if not tok_start_custom_frame.winfo_manager():
                tok_start_custom_frame.pack(side=tk.RIGHT)
        else:
            tok_start_custom_frame.pack_forget()

    def set_tok_start_mode(value):
        tok_start_mode_var.set(value)
        refresh_tok_start_row_controls()
        if value == "custom":
            tok_start_row_entry.focus_set()
            tok_start_row_entry.select_range(0, tk.END)

    tok_start_row_entry.bind("<Return>", lambda _event: run_script_from_gui())
    refresh_tok_start_row_controls()
    
    def browse_tok():
        filename = filedialog.askopenfilename(
            initialdir=os.getcwd(),
            title="Select a File",
            filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*"))
        )
        if filename:
            select_tok_input_file(filename)

    def on_drop_tok(path):
        if path:
            select_tok_input_file(path)

    if not attach_drop_target(drop_area, on_drop_tok):
        drop_area.config(text="Click to browse. Install tkinterdnd2 for drag-and-drop.", fg=TEXT_WARNING)
    
    drop_area.bind("<Button-1>", lambda _event: browse_tok())

    def refresh_tok_file():
        path = file_path_entry.get().strip()
        if not path:
            tok_status_label.config(text="Select a Tok input Excel file before refreshing.", fg=TEXT_WARNING)
            if tok_open_file_button.winfo_manager():
                tok_open_file_button.pack_forget()
            return
        select_tok_input_file(path, reset_start=False)

    btn_frame = tk.Frame(inner_controls, bg=BG_CARD)
    btn_frame.pack(fill=tk.X, pady=(2, 0))

    file_action_row = tk.Frame(btn_frame, bg=BG_CARD)
    file_action_row.pack(fill=tk.X, pady=(0, 7))
    refresh_button = create_styled_button(file_action_row, "Refresh File", refresh_tok_file, width=14, accent=False, height=42)
    refresh_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    tok_test_run_button = create_styled_button(file_action_row, "Test Run", lambda: set_run_speed("Test Run"), width=14, accent=False, height=42)
    tok_test_run_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

    run_button = create_styled_button(btn_frame, "Run", run_script_from_gui, width=20, height=42)
    run_button.pack(fill=tk.X)

    refresh_tok_recent_buttons()

def run_script_with_df(row_data, start_index=0):
    global running, actions, rows, row_index, start_time, after_id, automation_stop_reason, pending_tok_df
    running = True
    automation_stop_reason = ""
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

    rows = row_data if isinstance(row_data, list) else row_data.to_dict(orient='records')
    pending_tok_df = None
    gc.collect()

    total_rows = len(rows)
    if total_rows == 0:
        display_tok_input_error("No rows to process.")
        return

    start_index = max(0, min(int(start_index), total_rows - 1))
    start_time = datetime.datetime.now()
    actions = enter_data(rows[start_index], start_index + 1)
    row_index = start_index
    write_tok_debug_log(f"ROW_START row={row_index + 1}")
    update_progress(row_index)
    _cancel_scheduled_after()
    _cancel_watchdog()
    # Countdown already gave the user time to focus the target app.
    after_id = root.after(100, process_next_action)

def display_tok_input_error(message):
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Incompatible File", TEXT_ERROR)
    msg_label = create_styled_label(frame, message, size=10, color=TEXT_SECONDARY)
    msg_label.pack(pady=(0, 15))
    back_button = create_styled_button(frame, "← Back to Menu", initialize_main_menu, width=20)
    back_button.pack(pady=10)
    ensure_window_fits()

def display_tok_missing_fields_prompt(missing_rows, start_index=0):
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Missing Debit/Credit", TEXT_WARNING)

    shown = _format_missing_rows(missing_rows)
    msg_label = create_styled_label(frame, f"Missing DEBIT/CREDIT on row(s): {shown}", size=10, color=TEXT_SECONDARY)
    msg_label.pack(pady=(0, 15))

    def continue_anyway():
        if pending_tok_df is not None and pending_tok_file_path:
            start_tok_run(pending_tok_df, pending_tok_file_path, start_index=start_index)

    def open_file():
        if pending_tok_file_path:
            os.startfile(pending_tok_file_path)

    def refresh_file():
        if not pending_tok_file_path:
            msg_label.config(text="No selected file to refresh.", fg=TEXT_ERROR)
            return
        result = update_tok_input_status(pending_tok_file_path)
        if not result['ok']:
            display_tok_input_error(result['message'])
            return
        if result.get('missing_rows'):
            display_tok_missing_fields_prompt(result['missing_rows'], start_index=start_index)
            return
        open_tok_input_with_file(pending_tok_file_path)

    btn_frame = tk.Frame(frame, bg=BG_DARK)
    btn_frame.pack(pady=5)
    continue_btn = create_styled_button(btn_frame, "Continue Anyway", continue_anyway, width=20)
    continue_btn.pack(pady=5)
    refresh_btn = create_styled_button(btn_frame, "Refresh File", refresh_file, width=20, accent=False)
    refresh_btn.pack(pady=5)
    open_btn = create_styled_button(btn_frame, "Open File", open_file, width=20, accent=False)
    open_btn.pack(pady=5)

    back_button = create_styled_button(frame, "← Back to Menu", initialize_main_menu, width=20)
    back_button.pack(pady=10)
    ensure_window_fits()

def start_tok_run(df, file_path, start_index=0):
    global pending_tok_df, pending_tok_file_path, current_tok_file_path, running, automation_stop_reason, rows, row_index, current_tok_run_start_index
    running = True
    automation_stop_reason = ""
    pending_tok_df = df
    pending_tok_file_path = file_path
    current_tok_file_path = file_path
    total_rows = len(df)
    if total_rows == 0:
        display_tok_input_error("No rows to process.")
        return

    run_rows = df.to_dict(orient='records')
    pending_tok_df = None
    del df
    gc.collect()

    start_index = max(0, min(int(start_index), total_rows - 1))
    current_tok_run_start_index = start_index
    rows = run_rows
    row_index = start_index
    write_tok_debug_log(
        f"RUN_START file='{file_path}' total_rows={total_rows} start_row={start_index + 1} "
        f"speed={current_run_speed_label} action={action_delay:.3f} focus={focus_delay:.3f} "
        f"paste={paste_delay:.3f} post_paste={post_paste_delay:.3f}"
    )
    init_progress_bar(total_rows)
    countdown_label = create_styled_label(frame, f"Starting in {int(start_delay)} seconds...", size=11, color=TEXT_PRIMARY)
    countdown_label.pack(pady=10)
    progress_bar['value'] = start_index
    progress_label.config(text=f"Rows processed: {start_index} / {total_rows}")

    def update_countdown(seconds_left):
        global after_id
        if not running and automation_stop_reason:
            display_stopped_screen(automation_stop_reason)
            return
        countdown_label.config(text=f"Starting in {seconds_left} seconds...")
        progress_bar['value'] = start_index
        progress_label.config(text=f"Rows processed: {start_index} / {total_rows}")
        if seconds_left > 0:
            after_id = root.after(1000, update_countdown, seconds_left - 1)
        else:
            countdown_label.pack_forget()
            after_id = root.after(100, lambda: run_script_with_df(run_rows, start_index=start_index))

    update_countdown(int(start_delay))

def process_next_action():
    global row_index, actions, rows, watchdog_timer, after_id
    after_id = None
    if not running:
        if automation_stop_reason:
            display_stopped_screen(automation_stop_reason)
        return

    _cancel_watchdog()

    if actions:
        action = actions.pop(0)
        if len(action) == 3:
            func, delay, debug_text = action
        else:
            func, delay = action
            debug_text = ""
        if not running:
            return
        if _mouse_in_main_top_left():
            _stop_for_mouse_failsafe()
            display_stopped_screen(automation_stop_reason)
            return
        _arm_watchdog()
        try:
            if debug_text:
                write_tok_debug_log(f"ACTION {debug_text}")
            func()
        except pyautogui.FailSafeException:
            logging.info("PyAutoGUI fail-safe triggered.")
            _stop_for_mouse_failsafe()
            display_stopped_screen(automation_stop_reason)
            return
        except Exception:
            logging.exception("Action failed; stopping automation.")
            _mark_automation_stopped("Automation stopped after a paste/key action failed.")
            _cancel_watchdog()
            display_stopped_screen(automation_stop_reason)
            play_error_sound()
            return
        if not running:
            display_stopped_screen(automation_stop_reason)
            return
        _arm_watchdog()
        next_delay = action_delay if delay is USE_ACTION_DELAY else delay
        after_id = root.after(_after_ms(next_delay), process_next_action)
    else:
        write_tok_debug_log(f"ROW_DONE row={row_index + 1}")
        row_index += 1
        if row_index >= len(rows):
            complete_script()
        else:
            actions = enter_data(rows[row_index], row_index + 1)
            write_tok_debug_log(f"ROW_START row={row_index + 1}")
            update_progress(row_index)
            _arm_watchdog()
            after_id = root.after(_after_ms(action_delay), process_next_action)

def complete_script():
    global start_time, watchdog_timer, running, actions
    running = False
    actions = []
    _cancel_watchdog()
    end_time = datetime.datetime.now()
    time_elapsed = end_time - start_time
    write_tok_debug_log(f"RUN_COMPLETE rows_processed={row_index} runtime={time_elapsed.total_seconds():.2f}s")
    display_results(time_elapsed, row_index)
    gc.collect()
    logging.info("Completed.")
    logging.info(f"Total time elapsed: {time_elapsed.total_seconds():.2f} seconds")
    logging.info(f"Number of rows processed: {row_index}")

def run_script_from_gui():
    file_path = file_path_entry.get()
    global current_tok_file_path
    current_tok_file_path = file_path
    _sync_settings_from_ui()
    apply_tok_run_speed_selection()
    result = update_tok_input_status(file_path)
    if not result['ok']:
        play_error_sound()
        return

    df = result['df']
    start_excel_row = _tok_selected_start_excel_row()
    if start_excel_row is None:
        play_error_sound()
        return
    start_index = _tok_start_index_from_excel_row(df, start_excel_row)
    if start_index is None:
        play_error_sound()
        return

    missing_rows = [row for row in result['missing_rows'] if row >= max(2, start_excel_row)]
    if missing_rows:
        global pending_tok_df, pending_tok_file_path
        pending_tok_df = df
        pending_tok_file_path = file_path
        display_tok_missing_fields_prompt(missing_rows, start_index=start_index)
        return

    start_tok_run(df, file_path, start_index=start_index)

def increase_start_delay():
    current_value = int(settings_start_entry.get())
    settings_start_entry.delete(0, tk.END)
    settings_start_entry.insert(0, str(current_value + 1))
    _sync_settings_from_ui()

def decrease_start_delay():
    current_value = int(settings_start_entry.get())
    current_value = max(0, current_value - 1)
    settings_start_entry.delete(0, tk.END)
    settings_start_entry.insert(0, str(current_value))
    _sync_settings_from_ui()

def increase_action_delay():
    global action_delay
    current_value = float(settings_action_entry.get())
    new_value = current_value + 0.001
    settings_action_entry.delete(0, tk.END)
    settings_action_entry.insert(0, f"{new_value:.3f}")
    action_delay = new_value
    _sync_settings_from_ui()

def decrease_action_delay():
    global action_delay
    current_value = float(settings_action_entry.get())
    new_value = max(0, current_value - 0.001)
    settings_action_entry.delete(0, tk.END)
    settings_action_entry.insert(0, f"{new_value:.3f}")
    action_delay = new_value
    _sync_settings_from_ui()

def increase_focus_delay():
    global focus_delay
    current_value = float(settings_focus_entry.get())
    new_value = current_value + 0.001
    settings_focus_entry.delete(0, tk.END)
    settings_focus_entry.insert(0, f"{new_value:.3f}")
    focus_delay = new_value
    _sync_settings_from_ui()

def decrease_focus_delay():
    global focus_delay
    current_value = float(settings_focus_entry.get())
    new_value = max(0, current_value - 0.001)
    settings_focus_entry.delete(0, tk.END)
    settings_focus_entry.insert(0, f"{new_value:.3f}")
    focus_delay = new_value
    _sync_settings_from_ui()

def increase_paste_delay():
    global paste_delay
    current_value = float(settings_paste_entry.get())
    new_value = current_value + 0.001
    settings_paste_entry.delete(0, tk.END)
    settings_paste_entry.insert(0, f"{new_value:.3f}")
    paste_delay = new_value
    _sync_settings_from_ui()

def decrease_paste_delay():
    global paste_delay
    current_value = float(settings_paste_entry.get())
    new_value = max(0, current_value - 0.001)
    settings_paste_entry.delete(0, tk.END)
    settings_paste_entry.insert(0, f"{new_value:.3f}")
    paste_delay = new_value
    _sync_settings_from_ui()

def increase_post_paste_delay():
    global post_paste_delay
    current_value = float(settings_post_paste_entry.get())
    new_value = current_value + 0.001
    settings_post_paste_entry.delete(0, tk.END)
    settings_post_paste_entry.insert(0, f"{new_value:.3f}")
    post_paste_delay = new_value
    _sync_settings_from_ui()

def decrease_post_paste_delay():
    global post_paste_delay
    current_value = float(settings_post_paste_entry.get())
    new_value = max(0, current_value - 0.001)
    settings_post_paste_entry.delete(0, tk.END)
    settings_post_paste_entry.insert(0, f"{new_value:.3f}")
    post_paste_delay = new_value
    _sync_settings_from_ui()

def initialize_main_menu():
    _sync_settings_from_ui()
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Veldu aðgerð")

    menu_card, menu_inner = create_panel(frame, padx=28, pady=26)
    menu_card.pack(fill=tk.X, padx=28, pady=(16, 10))

    primary_row = tk.Frame(menu_inner, bg=BG_CARD)
    primary_row.pack(fill=tk.X, pady=(0, 14))
    tok_input_button = create_styled_button(primary_row, "Tok Input", initialize_tok_input, width=18)
    tok_input_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    bank_formatter_button = create_styled_button(primary_row, "Bank Formatter", initialize_bank_formatter, width=18)
    bank_formatter_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

    tools_grid = tk.Frame(menu_inner, bg=BG_CARD)
    tools_grid.pack(fill=tk.X)
    tool_buttons = [
        ("Round Numbers", initialize_round_numbers),
        ("Format Dates", initialize_format_dates),
        ("Format ID Numbers", initialize_format_ids),
        ("Settings", initialize_settings),
    ]
    for idx, (label, command) in enumerate(tool_buttons):
        btn = create_styled_button(tools_grid, label, command, width=17, accent=False)
        btn.grid(row=idx // 2, column=idx % 2, sticky='ew', padx=(0, 8) if idx % 2 == 0 else (8, 0), pady=7)
    tools_grid.columnconfigure(0, weight=1)
    tools_grid.columnconfigure(1, weight=1)
    ensure_window_fits()

def initialize_tok_input():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Tok Input")
    
    display_input_controls()
    
    nav_row = tk.Frame(frame, bg=BG_DARK)
    nav_row.pack(fill=tk.X, padx=28, pady=(0, 10))
    back_button = create_styled_button(nav_row, "← Back to Menu", initialize_main_menu, width=18, accent=False)
    back_button.pack(side=tk.RIGHT, fill=tk.X, expand=False)
    ensure_window_fits()

def open_tok_input_with_file(file_path):
    initialize_tok_input()
    if not file_path:
        return
    select_tok_input_file(file_path)

def open_file_location(file_path):
    if not file_path:
        return
    file_path = os.path.abspath(file_path)
    if os.name == "nt":
        try:
            if os.path.exists(file_path):
                subprocess.run(["explorer", "/select,", file_path], check=False)
            else:
                folder = os.path.dirname(file_path)
                while folder and not os.path.isdir(folder):
                    parent = os.path.dirname(folder)
                    if parent == folder:
                        break
                    folder = parent
                os.startfile(folder if os.path.isdir(folder) else _bank_output_dir())
            return
        except Exception:
            pass
    try:
        folder = os.path.dirname(file_path)
        os.startfile(folder if os.path.isdir(folder) else _bank_output_dir())
    except Exception:
        pass

def initialize_settings():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Settings")

    def _sync_settings_scroll_region(_event=None):
        schedule_layout_refresh()

    settings_card, settings_frame = create_panel(frame, padx=24, pady=22)
    settings_card.pack(fill=tk.X, padx=28, pady=(0, 12))

    help_label = create_styled_label(
        settings_frame,
        "If Tok skips fields or values are missing, use Slow. If everything works, use Fast.",
        size=9,
        color=TEXT_SECONDARY
    )
    help_label.configure(wraplength=420, justify='left')
    help_label.pack(anchor='w', pady=(0, 8))

    global settings_action_entry
    global settings_focus_entry
    global settings_paste_entry
    global settings_post_paste_entry
    global settings_start_entry
    global settings_theme_var
    global settings_output_dir_entry
    global settings_compressed_output_dir_entry
    global settings_auto_code_training_entry, settings_auto_code_status_label

    manual_section_open = tk.BooleanVar(value=False)
    manual_grid = tk.Frame(settings_frame, bg=BG_CARD)

    def _make_entry(value, decimals=True):
        entry = create_styled_entry(manual_grid, width=7)
        entry.configure(justify='center')
        entry.insert(0, f"{value:.3f}" if decimals else str(value))
        return entry

    settings_action_entry = _make_entry(settings.get('action_delay', 0.1))
    settings_focus_entry = _make_entry(settings.get('focus_delay', 0.02))
    settings_paste_entry = _make_entry(settings.get('paste_delay', 0.0))
    settings_post_paste_entry = _make_entry(settings.get('post_paste_delay', 0.02))
    settings_start_entry = _make_entry(settings.get("start_delay", 3), decimals=False)

    def _set_entry(entry, value, decimals=True):
        entry.delete(0, tk.END)
        entry.insert(0, f"{value:.3f}" if decimals else str(value))

    def _apply_preset(values):
        _set_entry(settings_action_entry, values["action"])
        _set_entry(settings_focus_entry, values["focus"])
        _set_entry(settings_paste_entry, values["paste"])
        _set_entry(settings_post_paste_entry, values["post_paste"])
        _set_entry(settings_start_entry, values["start"], decimals=False)
        _sync_settings_from_ui()

    settings_theme_var = tk.StringVar(value=settings.get("theme", "light"))
    theme_label = create_styled_label(settings_frame, "Theme", size=10, bold=True)
    theme_label.pack(anchor='w', pady=(4, 6))

    theme_frame = tk.Frame(settings_frame, bg=BG_CARD)
    theme_frame.pack(fill=tk.X, pady=(0, 14))
    theme_buttons = {}

    def _refresh_theme_buttons():
        current_theme = settings_theme_var.get()
        for theme_key, button in theme_buttons.items():
            set_button_accent(button, theme_key == current_theme)

    def _set_theme(theme_key):
        settings_theme_var.set(theme_key)
        settings["theme"] = theme_key
        save_settings()
        apply_theme(theme_key)
        initialize_settings()

    for label, theme_key in (("Light", "light"), ("Dark", "dark")):
        theme_btn = create_styled_button(
            theme_frame,
            label,
            lambda key=theme_key: _set_theme(key),
            width=10,
            accent=False
        )
        theme_buttons[theme_key] = theme_btn
        theme_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8) if theme_key == "light" else (8, 0))
    _refresh_theme_buttons()

    preset_label = create_styled_label(settings_frame, "Speed preset", size=10, bold=True)
    preset_label.pack(anchor='w', pady=(4, 6))

    preset_frame = tk.Frame(settings_frame, bg=BG_CARD)
    preset_frame.pack(fill=tk.X, pady=(0, 14))
    presets = SPEED_PRESETS
    preset_buttons = {}

    def _entry_float(entry, fallback=0):
        try:
            return float(entry.get())
        except Exception:
            return fallback

    def _refresh_preset_buttons():
        current = {
            "action": _entry_float(settings_action_entry),
            "focus": _entry_float(settings_focus_entry),
            "paste": _entry_float(settings_paste_entry),
            "post_paste": _entry_float(settings_post_paste_entry),
            "start": _entry_float(settings_start_entry),
        }
        for label, values in presets:
            is_active = all(abs(current[key] - float(values[key])) < 0.0005 for key in values)
            set_button_accent(preset_buttons[label], is_active)

    def _preset_command(values):
        _apply_preset(values)
        _refresh_preset_buttons()

    for label, values in presets:
        preset_btn = create_styled_button(
            preset_frame,
            label,
            lambda vals=values: _preset_command(vals),
            width=10,
            accent=False
        )
        preset_buttons[label] = preset_btn
        preset_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8) if label != "Slow" else (0, 0))

    manual_toggle_frame = tk.Frame(settings_frame, bg=BG_CARD, cursor='hand2')
    manual_toggle_frame.pack(fill=tk.X, pady=(0, 16))
    tk.Frame(manual_toggle_frame, bg=BORDER, height=1).pack(fill=tk.X, pady=(0, 8))
    manual_toggle_inner = tk.Frame(manual_toggle_frame, bg=BG_CARD, cursor='hand2')
    manual_toggle_inner.pack(fill=tk.X)
    manual_toggle_label = create_styled_label(
        manual_toggle_inner,
        "+ Show manual tuning",
        size=9,
        color=ACCENT,
        bold=True,
        bg=BG_CARD
    )
    manual_toggle_label.pack(side=tk.LEFT)
    manual_toggle_hint = create_styled_label(
        manual_toggle_inner,
        "Advanced",
        size=8,
        color=TEXT_SECONDARY,
        bg=BG_CARD
    )
    manual_toggle_hint.pack(side=tk.RIGHT)

    def _small_button(parent, text, command):
        btn = tk.Button(parent, text=text, command=command,
                        bg=BUTTON_MUTED, fg=TEXT_PRIMARY, font=('Segoe UI', 11, 'bold'),
                        activebackground=BUTTON_MUTED_HOVER, activeforeground=TEXT_PRIMARY,
                        relief='flat', width=2, cursor='hand2',
                        bd=0, highlightthickness=0)
        btn.bind('<Enter>', lambda _e, b=btn: b.configure(bg=BUTTON_MUTED_HOVER))
        btn.bind('<Leave>', lambda _e, b=btn: b.configure(bg=BUTTON_MUTED))
        return btn

    def _with_preset_refresh(command):
        def _wrapped():
            command()
            _refresh_preset_buttons()
        return _wrapped

    def _timing_row(row, title, hint, entry, decrease_cmd, increase_cmd):
        label_frame = tk.Frame(manual_grid, bg=BG_CARD)
        label_frame.grid(row=row, column=0, sticky='w', pady=5)
        title_label = create_styled_label(label_frame, title, size=9, bold=True, bg=BG_CARD)
        title_label.pack(anchor='w')
        hint_label = create_styled_label(label_frame, hint, size=8, color=TEXT_SECONDARY, bg=BG_CARD)
        hint_label.configure(wraplength=260, justify='left')
        hint_label.pack(anchor='w')

        entry.grid(row=row, column=1, padx=(8, 8), pady=5)
        _small_button(manual_grid, "-", _with_preset_refresh(decrease_cmd)).grid(row=row, column=2, padx=(0, 5), pady=5)
        _small_button(manual_grid, "+", _with_preset_refresh(increase_cmd)).grid(row=row, column=3, padx=(0, 0), pady=5)

    manual_grid.columnconfigure(0, weight=1)
    _timing_row(0, "Between steps", "Overall speed. Increase if Tok jumps to the wrong field.", settings_action_entry, decrease_action_delay, increase_action_delay)
    _timing_row(1, "After Enter", "Wait after moving fields. Increase if the cursor cannot keep up.", settings_focus_entry, decrease_focus_delay, increase_focus_delay)
    _timing_row(2, "Before paste", "Wait after copying to clipboard. Usually keep this near zero.", settings_paste_entry, decrease_paste_delay, increase_paste_delay)
    _timing_row(3, "After paste", "Most important for lag. Increase if values are missing or cut off.", settings_post_paste_entry, decrease_post_paste_delay, increase_post_paste_delay)
    _timing_row(4, "Start countdown", "Time to click into Tok before typing starts.", settings_start_entry, decrease_start_delay, increase_start_delay)
    _refresh_preset_buttons()

    def _toggle_manual_tuning(_event=None):
        if manual_section_open.get():
            manual_grid.pack_forget()
            manual_section_open.set(False)
            manual_toggle_label.config(text="+ Show manual tuning")
            manual_toggle_hint.config(text="Advanced")
        else:
            manual_grid.pack(fill=tk.X, pady=(0, 8), before=output_frame)
            manual_section_open.set(True)
            manual_toggle_label.config(text="- Hide manual tuning")
            manual_toggle_hint.config(text="Delay controls")
        _sync_settings_scroll_region()
        schedule_layout_refresh()
        ensure_window_fits()

    for widget in (manual_toggle_frame, manual_toggle_inner, manual_toggle_label, manual_toggle_hint):
        widget.bind("<Button-1>", _toggle_manual_tuning)

    # Bank output folder
    output_frame = tk.Frame(settings_frame, bg=BG_CARD)
    output_frame.pack(fill=tk.X, pady=(12, 4))
    output_label = create_styled_label(output_frame, "Bank output folder:", size=9, color=TEXT_SECONDARY)
    output_label.pack(anchor='w')

    settings_output_dir_entry = create_styled_entry(output_frame, width=42)
    settings_output_dir_entry.insert(0, _bank_output_dir())
    settings_output_dir_entry.pack(side=tk.LEFT, padx=(0, 8), pady=5, ipady=5, fill=tk.X, expand=True)

    def browse_output_dir():
        chosen = filedialog.askdirectory(initialdir=os.path.expanduser("~"), title="Select Output Folder")
        if chosen:
            settings_output_dir_entry.delete(0, tk.END)
            settings_output_dir_entry.insert(0, chosen)
            _sync_settings_from_ui()
            schedule_layout_refresh()

    browse_btn = create_styled_button(output_frame, "Browse Folder", browse_output_dir, width=15, accent=False)
    browse_btn.pack(side=tk.LEFT)

    # Compressed transaction reference output folder
    compressed_output_frame = tk.Frame(settings_frame, bg=BG_CARD)
    compressed_output_frame.pack(fill=tk.X, pady=(10, 4))
    compressed_output_label = create_styled_label(compressed_output_frame, "Compressed TX output folder:", size=9, color=TEXT_SECONDARY)
    compressed_output_label.pack(anchor='w')

    settings_compressed_output_dir_entry = create_styled_entry(compressed_output_frame, width=42)
    settings_compressed_output_dir_entry.insert(0, _compressed_output_dir())
    settings_compressed_output_dir_entry.pack(side=tk.LEFT, padx=(0, 8), pady=5, ipady=5, fill=tk.X, expand=True)

    def browse_compressed_output_dir():
        current_dir = settings_compressed_output_dir_entry.get().strip() or _bank_output_dir()
        chosen = filedialog.askdirectory(initialdir=current_dir, title="Select Compressed TX Output Folder")
        if chosen:
            settings_compressed_output_dir_entry.delete(0, tk.END)
            settings_compressed_output_dir_entry.insert(0, chosen)
            _sync_settings_from_ui()
            schedule_layout_refresh()

    compressed_browse_btn = create_styled_button(compressed_output_frame, "Browse Folder", browse_compressed_output_dir, width=15, accent=False)
    compressed_browse_btn.pack(side=tk.LEFT)

    # Auto-code training folder
    training_frame = tk.Frame(settings_frame, bg=BG_CARD)
    training_frame.pack(fill=tk.X, pady=(10, 4))
    training_label = create_styled_label(training_frame, "Auto-code training folder:", size=9, color=TEXT_SECONDARY)
    training_label.pack(anchor='w')

    settings_auto_code_training_entry = create_styled_entry(training_frame, width=42)
    settings_auto_code_training_entry.insert(0, settings.get("auto_code_training_dir", _default_auto_code_training_dir()))
    settings_auto_code_training_entry.pack(side=tk.LEFT, padx=(0, 8), pady=5, ipady=5, fill=tk.X, expand=True)

    def browse_training_dir():
        current_dir = settings_auto_code_training_entry.get().strip() or os.path.expanduser("~")
        chosen = filedialog.askdirectory(initialdir=current_dir, title="Select Auto-Code Training Folder")
        if chosen:
            settings_auto_code_training_entry.delete(0, tk.END)
            settings_auto_code_training_entry.insert(0, chosen)
            _sync_settings_from_ui()
            schedule_layout_refresh()

    training_browse_btn = create_styled_button(training_frame, "Browse Folder", browse_training_dir, width=15, accent=False)
    training_browse_btn.pack(side=tk.LEFT)

    memory_frame = tk.Frame(settings_frame, bg=BG_CARD)
    memory_frame.pack(fill=tk.X, pady=(2, 10))
    settings_auto_code_status_label = create_styled_label(
        memory_frame,
        f"Memory cache: {os.path.join(settings.get('auto_code_training_dir', _default_auto_code_training_dir()), 'auto_code_memory.json')}",
        size=8,
        color=TEXT_SECONDARY,
        bg=BG_CARD
    )
    settings_auto_code_status_label.configure(wraplength=420, justify='left')
    settings_auto_code_status_label.pack(anchor='w', pady=(0, 6))

    def rebuild_auto_code_memory():
        _sync_settings_from_ui()
        training_dir = _auto_code_training_dir()
        memory_path = _auto_code_memory_path()
        if not AUTO_CODER_AVAILABLE or get_auto_coder is None:
            settings_auto_code_status_label.config(text=f"Auto-coder unavailable: {AUTO_CODER_ERROR[:80]}", fg=TEXT_ERROR)
            return
        try:
            coder = get_auto_coder(
                force_reload=True,
                force_rebuild=True,
                training_dir=training_dir,
                key_dir=_auto_code_key_dir(),
                memory_path=memory_path,
            )
            summary = coder.summary()
            if coder.ready():
                settings_auto_code_status_label.config(
                    text=(
                        f"Rebuilt memory: {summary['training_rows']} rows, "
                        f"{summary['merchant_keys']} merchants. Saved: {memory_path}"
                    ),
                    fg=TEXT_SUCCESS
                )
            else:
                settings_auto_code_status_label.config(
                    text=f"No usable coded rows found in: {training_dir}",
                    fg=TEXT_WARNING
                )
        except Exception as exc:
            settings_auto_code_status_label.config(text=f"Rebuild failed: {str(exc)[:90]}", fg=TEXT_ERROR)

    rebuild_memory_btn = create_styled_button(memory_frame, "Rebuild Auto-Code Memory", rebuild_auto_code_memory, width=22, accent=False)
    rebuild_memory_btn.pack(anchor='w')

    def _save_settings_and_refresh():
        _sync_settings_from_ui()
        _refresh_preset_buttons()
        _sync_settings_scroll_region()
        schedule_layout_refresh()

    bottom_row = tk.Frame(frame, bg=BG_DARK)
    bottom_row.pack(fill=tk.X, padx=28, pady=(2, 10))
    save_btn = create_styled_button(bottom_row, "Save Settings", _save_settings_and_refresh, width=16, accent=True)
    save_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

    back_button = create_styled_button(bottom_row, "← Back", initialize_main_menu, width=14, accent=False)
    back_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))
    ensure_window_fits()

def initialize_round_numbers():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Round Numbers")

    container = tk.Frame(frame, bg=BG_DARK)
    container.pack(fill=tk.X, padx=28)

    input_label = create_styled_label(container, "Paste numbers (one per line):", size=9, color=TEXT_SECONDARY)
    input_label.pack(anchor='w')

    input_shell, input_text = create_styled_text_area(container, "Paste numbers here, one per line", height=6)
    input_shell.pack(fill=tk.X, pady=(5, 12))
    input_resize = attach_auto_resize_text(input_text, min_lines=6, max_lines=8)

    output_label = create_styled_label(container, "Rounded output:", size=9, color=TEXT_SECONDARY)
    output_label.pack(anchor='w')

    output_shell, output_text = create_styled_text_area(container, "Rounded numbers will appear here", height=6)
    output_shell.pack(fill=tk.X, pady=(5, 12))
    output_resize = attach_auto_resize_text(output_text, min_lines=5, max_lines=7)

    status_label = create_styled_label(container, "", size=9, color=TEXT_SECONDARY)
    status_label.pack(pady=(0, 8))

    def round_numbers():
        raw = get_text_area_value(input_text)
        lines = [line.strip() for line in raw.splitlines() if line.strip()]
        rounded = []
        bad = 0
        for line in lines:
            normalized = line.replace(" ", "").replace(",", ".")
            try:
                value = Decimal(normalized)
                rounded_value = value.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
                rounded.append(str(int(rounded_value)))
            except (InvalidOperation, ValueError):
                bad += 1
        set_text_area_value(output_text, "\n".join(rounded))
        output_resize()
        schedule_layout_refresh()
        if bad:
            status_label.config(text=f"Skipped {bad} invalid line(s).", fg=TEXT_WARNING)
        else:
            status_label.config(text="Rounded successfully.", fg=TEXT_SUCCESS)

    def copy_output():
        text = get_text_area_value(output_text).strip()
        if text:
            pyperclip.copy(text)
            status_label.config(text="Copied to clipboard (one per row).", fg=TEXT_SUCCESS)
        else:
            status_label.config(text="No output to copy.", fg=TEXT_WARNING)

    btn_row = tk.Frame(container, bg=BG_DARK)
    btn_row.pack(pady=5)

    round_btn = create_styled_button(btn_row, "Round", round_numbers, width=12)
    round_btn.pack(side=tk.LEFT, padx=5)

    copy_btn = create_styled_button(btn_row, "Copy Output", copy_output, width=12, accent=False)
    copy_btn.pack(side=tk.LEFT, padx=5)

    back_button = create_styled_button(frame, "← Back to Menu", initialize_main_menu, width=20)
    back_button.pack(pady=10)
    ensure_window_fits()

def initialize_format_dates():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Format Dates")

    container = tk.Frame(frame, bg=BG_DARK)
    container.pack(fill=tk.X, padx=28)

    input_label = create_styled_label(container, "Paste dates (one per line):", size=9, color=TEXT_SECONDARY)
    input_label.pack(anchor='w')

    input_shell, input_text = create_styled_text_area(container, "Paste dates like 31.12.2025, 31/12/25, or 31 desember 2025", height=6)
    input_shell.pack(fill=tk.X, pady=(5, 12))
    input_resize = attach_auto_resize_text(input_text, min_lines=6, max_lines=8)

    output_label = create_styled_label(container, "Formatted output (DD.MM.YYYY):", size=9, color=TEXT_SECONDARY)
    output_label.pack(anchor='w')

    output_shell, output_text = create_styled_text_area(container, "Formatted dates will appear here", height=6)
    output_shell.pack(fill=tk.X, pady=(5, 12))
    output_resize = attach_auto_resize_text(output_text, min_lines=5, max_lines=7)

    status_label = create_styled_label(container, "", size=9, color=TEXT_SECONDARY)
    status_label.pack(pady=(0, 8))

    def format_dates():
        raw = get_text_area_value(input_text)
        lines = [line.strip() for line in raw.splitlines() if line.strip()]
        formatted = []
        bad = 0
        for line in lines:
            parsed = _format_date_input_line(line)
            if parsed:
                formatted.append(parsed)
            else:
                bad += 1

        set_text_area_value(output_text, "\n".join(formatted))
        output_resize()
        schedule_layout_refresh()
        if bad:
            status_label.config(text=f"Skipped {bad} invalid line(s).", fg=TEXT_WARNING)
        else:
            status_label.config(text="Formatted successfully.", fg=TEXT_SUCCESS)

    def copy_output():
        text = get_text_area_value(output_text).strip()
        if text:
            pyperclip.copy(text)
            status_label.config(text="Copied to clipboard (one per row).", fg=TEXT_SUCCESS)
        else:
            status_label.config(text="No output to copy.", fg=TEXT_WARNING)

    btn_row = tk.Frame(container, bg=BG_DARK)
    btn_row.pack(pady=5)

    format_btn = create_styled_button(btn_row, "Format", format_dates, width=12)
    format_btn.pack(side=tk.LEFT, padx=5)

    copy_btn = create_styled_button(btn_row, "Copy Output", copy_output, width=12, accent=False)
    copy_btn.pack(side=tk.LEFT, padx=5)

    back_button = create_styled_button(frame, "← Back to Menu", initialize_main_menu, width=20)
    back_button.pack(pady=10)
    ensure_window_fits()

def initialize_format_ids():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Format ID Numbers")

    container = tk.Frame(frame, bg=BG_DARK)
    container.pack(fill=tk.X, padx=28)

    input_label = create_styled_label(container, "Paste IDs (one per line):", size=9, color=TEXT_SECONDARY)
    input_label.pack(anchor='w')

    input_shell, input_text = create_styled_text_area(container, "Paste kennitala values here, one per line", height=6)
    input_shell.pack(fill=tk.X, pady=(5, 12))
    input_resize = attach_auto_resize_text(input_text, min_lines=6, max_lines=8)

    toggle_frame = tk.Frame(container, bg=BG_DARK)
    toggle_frame.pack(pady=(0, 8))

    mode_var = tk.StringVar(value="remove")
    mode_buttons = {}

    def set_id_mode(mode):
        mode_var.set(mode)
        for value, button in mode_buttons.items():
            set_button_accent(button, value == mode)

    mode_buttons["remove"] = create_styled_button(
        toggle_frame,
        "Remove dash",
        lambda: set_id_mode("remove"),
        width=14,
        accent=True,
    )
    mode_buttons["remove"].pack(side=tk.LEFT, padx=(0, 8))

    mode_buttons["add"] = create_styled_button(
        toggle_frame,
        "Add dash",
        lambda: set_id_mode("add"),
        width=14,
        accent=False,
    )
    mode_buttons["add"].pack(side=tk.LEFT, padx=(8, 0))

    output_label = create_styled_label(container, "Formatted output:", size=9, color=TEXT_SECONDARY)
    output_label.pack(anchor='w')

    output_shell, output_text = create_styled_text_area(container, "Formatted ID numbers will appear here", height=6)
    output_shell.pack(fill=tk.X, pady=(5, 12))
    output_resize = attach_auto_resize_text(output_text, min_lines=5, max_lines=7)

    status_label = create_styled_label(container, "", size=9, color=TEXT_SECONDARY)
    status_label.pack(pady=(0, 8))

    def format_ids():
        raw = get_text_area_value(input_text)
        lines = [line.strip() for line in raw.splitlines() if line.strip()]
        formatted = []
        bad = 0
        for line in lines:
            digits = "".join(ch for ch in line if ch.isdigit())
            if len(digits) < 10:
                bad += 1
                continue
            base = digits[:10]
            if mode_var.get() == "add":
                formatted.append(f"{base[:6]}-{base[6:]}")
            else:
                formatted.append(base)

        set_text_area_value(output_text, "\n".join(formatted))
        output_resize()
        schedule_layout_refresh()
        if bad:
            status_label.config(text=f"Skipped {bad} invalid line(s).", fg=TEXT_WARNING)
        else:
            status_label.config(text="Formatted successfully.", fg=TEXT_SUCCESS)

    def copy_output():
        text = get_text_area_value(output_text).strip()
        if text:
            pyperclip.copy(text)
            status_label.config(text="Copied to clipboard (one per row).", fg=TEXT_SUCCESS)
        else:
            status_label.config(text="No output to copy.", fg=TEXT_WARNING)

    btn_row = tk.Frame(container, bg=BG_DARK)
    btn_row.pack(pady=5)

    format_btn = create_styled_button(btn_row, "Format", format_ids, width=12)
    format_btn.pack(side=tk.LEFT, padx=5)

    copy_btn = create_styled_button(btn_row, "Copy Output", copy_output, width=12, accent=False)
    copy_btn.pack(side=tk.LEFT, padx=5)

    back_button = create_styled_button(frame, "← Back to Menu", initialize_main_menu, width=20)
    back_button.pack(pady=10)
    ensure_window_fits()

# Bank formatter workflow lives in bank_formatter.py.

def play_success_sound():
    sound_file_path = resource_path(os.path.join("tok", "success.wav"))
    winsound.PlaySound(sound_file_path, winsound.SND_FILENAME | winsound.SND_ASYNC)

def play_error_sound():
    error_sound_file_path = resource_path(os.path.join("tok", "error.wav"))
    winsound.PlaySound(error_sound_file_path, winsound.SND_FILENAME | winsound.SND_ASYNC)

def on_watchdog_timeout():
    logging.error("Watchdog timer exceeded %s seconds. Stopping automation.", AUTOMATION_WATCHDOG_SECONDS)
    _mark_automation_stopped(
        f"Automation paused for over {AUTOMATION_WATCHDOG_SECONDS} seconds. "
        "Check Tok, then resume from the current or next row."
    )
    try:
        play_error_sound()
    except Exception:
        pass

def configure_bank_formatter_bridge():
    bank_formatter.configure(
        root=root,
        frame=frame,
        set_page_title=set_page_title,
        initialize_main_menu=initialize_main_menu,
        open_tok_input_with_file=open_tok_input_with_file,
        open_file_location=open_file_location,
        play_success_sound=play_success_sound,
        play_error_sound=play_error_sound,
        ensure_window_fits=ensure_window_fits,
        fit_dialog_to_content=fit_dialog_to_content,
        schedule_layout_refresh=schedule_layout_refresh,
        get_recent_files=get_recent_files,
        remember_recent_file=remember_recent_file,
        _short_file_name=_short_file_name,
        _status_with_file_detail=_status_with_file_detail,
        _bank_output_dir=_bank_output_dir,
        _compressed_output_dir=_compressed_output_dir,
        _auto_code_training_dir=_auto_code_training_dir,
        _auto_code_key_dir=_auto_code_key_dir,
        _auto_code_memory_path=_auto_code_memory_path,
        RECENT_VISIBLE_LIMIT=RECENT_VISIBLE_LIMIT,
    )

# =============================================================================
# MAIN APPLICATION SETUP
# =============================================================================

def main():
    global root, settings, action_delay, focus_delay, paste_delay, post_paste_delay, start_delay
    global style, header_frame, logo_image, logo_image_small, logo_label
    global title_frame, page_title_label, frame

    if TKDND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    settings = load_settings()
    apply_theme(settings.get("theme", "light"))
    action_delay = settings.get("action_delay", 0.1)
    focus_delay = settings.get("focus_delay", 0.02)
    paste_delay = settings.get("paste_delay", 0.0)
    post_paste_delay = settings.get("post_paste_delay", 0.02)
    start_delay = settings.get("start_delay", 3)
    root.title("Tok Tenging")
    root.geometry("680x700")
    root.configure(bg=BG_DARK)
    root.resizable(True, True)
    root.minsize(620, 560)

    # Configure ttk styles (progressbar only)
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('Horizontal.TProgressbar',
                    background=ACCENT,
                    troughcolor=BG_INPUT,
                    bordercolor=BG_DARK,
                    lightcolor=ACCENT,
                    darkcolor=ACCENT)

    # Header with smaller logo and always-available window fit control
    header_frame = tk.Frame(root, bg=BG_DARK)
    header_frame.pack(fill=tk.X, padx=22, pady=(10, 0))
    header_frame.columnconfigure(0, weight=1)
    header_frame.columnconfigure(1, weight=0)
    header_frame.columnconfigure(2, weight=1)

    logo_path = resource_path(os.path.join("tok", "temp2.png"))
    logo_image = PhotoImage(file=logo_path)
    logo_image_small = logo_image.subsample(2, 2)
    logo_label = tk.Label(header_frame, image=logo_image_small, bg=BG_DARK)
    logo_label.image = logo_image_small
    logo_label.grid(row=0, column=1)
    root.iconphoto(True, logo_image)

    title_frame = tk.Frame(root, bg=BG_DARK, height=44)
    title_frame.pack(fill=tk.X, padx=30, pady=(18, 0))
    title_frame.pack_propagate(False)
    page_title_label = create_styled_label(title_frame, "", size=16, bold=True, bg=BG_DARK)
    page_title_label.pack(expand=True)

    # Main content frame - simple, no scroll needed with proper sizing
    frame = tk.Frame(root, bg=BG_DARK)
    frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=(8, 20))

    configure_bank_formatter_bridge()
    initialize_main_menu()
    root.mainloop()


if __name__ == "__main__":
    main()

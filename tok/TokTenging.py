import tkinter as tk
from tkinter import ttk, filedialog, messagebox, PhotoImage
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKDND_AVAILABLE = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    TKDND_AVAILABLE = False
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
import ctypes
import queue
from ctypes import wintypes
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import winsound
from threading import Event, Thread, Timer

try:
    try:
        from .auto_coder import get_auto_coder
    except ImportError:
        from auto_coder import get_auto_coder
    AUTO_CODER_AVAILABLE = True
    AUTO_CODER_ERROR = ""
except Exception as exc:
    get_auto_coder = None
    AUTO_CODER_AVAILABLE = False
    AUTO_CODER_ERROR = str(exc)

try:
    try:
        from .ml_runtime import ml_runtime_status
    except ImportError:
        from ml_runtime import ml_runtime_status
    ML_STATUS_AVAILABLE = True
except Exception:
    ml_runtime_status = None
    ML_STATUS_AVAILABLE = False

settings = {}
runtime_timing = {}
automation_state = {
    "active": False,
    "watchdog": None,
    "after": None,
    "stop_reason": "",
    "next_due": None,
    "target_hwnd": 0,
    "target_pid": 0,
    "target_title": "",
    "focus_mismatch_since": None,
    "target_unresponsive_since": None,
    "current_action": "",
    "lag_events": 0,
}
SETTINGS_PATH = os.path.join(os.path.expanduser("~"), "AppData", "Local", "TokTenging", "settings.json")
DEBUG_LOG_PATH = os.path.join(os.path.dirname(SETTINGS_PATH), "tok_input_debug.log")
pending_tok_df = None
pending_tok_file_path = None
current_tok_file_path = None
current_tok_run_start_index = 0
current_tok_rows_completed = 0
tok_excel_row_numbers = []
current_run_speed_label = "Saved"
AUTOMATION_WATCHDOG_SECONDS = 15
MOUSE_FAILSAFE_MARGIN = 8
SCHEDULER_LAG_THRESHOLD_SECONDS = 0.18
FOCUS_MISMATCH_GRACE_SECONDS = 0.6
TARGET_UNRESPONSIVE_TIMEOUT_SECONDS = 4.0
RECENT_FILE_LIMIT = 5
RECENT_VISIBLE_LIMIT = 3
SPEED_PRESETS = [
    ("Fast", {"action": 0.048, "focus": 0.080, "paste": 0.018, "post_paste": 0.090, "start": 3}),
    ("Balanced", {"action": 0.060, "focus": 0.100, "paste": 0.022, "post_paste": 0.112, "start": 3}),
    ("Slow", {"action": 0.081, "focus": 0.272, "paste": 0.030, "post_paste": 0.152, "start": 3}),
]
SPEED_PRESET_MAP = {label: values for label, values in SPEED_PRESETS}
SPEED_PRESET_MAP["Test Run"] = {"action": 3.0, "focus": 1.0, "paste": 3.0, "post_paste": 1.0, "start": 3}
SPEED_DISPLAY_LABELS = {
    "Saved": "Vistað",
    "Fast": "Hratt",
    "Balanced": "Miðlungs",
    "Slow": "Hægt",
    "Test Run": "Prufukeyrsla",
}
AUTO_CODE_MODES = frozenset({"legacy", "ml_hybrid"})

# PyAutoGUI sleeps 0.1s after every call by default. The app already controls
# pacing with the delay settings below, so keep PyAutoGUI itself immediate.
pyautogui.PAUSE = 0
pyautogui.FAILSAFE = True

# Sentinel used by the action queue to mean "use the current action delay".
USE_ACTION_DELAY = object()

# =============================================================================
# APP MODULES
# =============================================================================
try:
    from . import bank_formatter, coded_file_tools, email_downloader, folder_printer, shareholder_registry, ui_components
    from .bank_formatter import initialize_bank_formatter
    from .bank_utils import format_date_as_text, _format_date_input_line
    from .shareholder_registry import initialize_shareholder_registry
    from .training_archive import sync_all_training_archives
    from .ui_components import (
        RoundedPanel,
        attach_auto_resize_text,
        attach_drop_target,
        create_drop_box,
        create_panel,
        create_styled_button,
        create_styled_entry,
        create_styled_label,
        create_styled_text_area,
        get_text_area_value,
        set_button_accent,
        set_text_area_value,
    )
except ImportError:
    import bank_formatter
    import coded_file_tools
    import email_downloader
    import folder_printer
    import shareholder_registry
    import ui_components
    from bank_formatter import initialize_bank_formatter
    from bank_utils import format_date_as_text, _format_date_input_line
    from shareholder_registry import initialize_shareholder_registry
    from training_archive import sync_all_training_archives
    from ui_components import (
        RoundedPanel,
        attach_auto_resize_text,
        attach_drop_target,
        create_drop_box,
        create_panel,
        create_styled_button,
        create_styled_entry,
        create_styled_label,
        create_styled_text_area,
        get_text_area_value,
        set_button_accent,
        set_text_area_value,
    )

ui_components.configure_drag_and_drop(TKDND_AVAILABLE, DND_FILES)
ui_components.set_layout_refresh_callback(lambda: schedule_layout_refresh())
THEMES = ui_components.THEMES
globals().update(ui_components.get_theme("light"))

# Settings layout helpers
SETTINGS_LABEL_WIDTH = 24


def apply_theme(theme_name):
    theme = ui_components.set_theme(theme_name)
    globals().update(theme)
    bank_formatter.set_theme(theme)
    coded_file_tools.set_theme(theme)
    shareholder_registry.set_theme(theme)
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

def _ensure_folder(path, fallback):
    requested = (path or fallback or os.path.expanduser("~")).strip()
    for candidate in (requested, fallback, os.path.expanduser("~")):
        if not candidate:
            continue
        try:
            os.makedirs(candidate, exist_ok=True)
            if os.path.isdir(candidate):
                return candidate
        except Exception:
            continue
    return os.getcwd()


def _default_auto_code_training_dir():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "trainingCoded")


def _default_training_data_dir():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "trainingData")


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


def _training_data_dir():
    output_dir = _ensure_folder(settings.get("training_data_dir"), _default_training_data_dir())
    if settings.get("training_data_dir") != output_dir:
        settings["training_data_dir"] = output_dir
        save_settings()
    return output_dir


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


def _normalized_auto_code_mode(value):
    mode = str(value or "").strip().lower()
    return mode if mode in AUTO_CODE_MODES else "legacy"


def _auto_code_mode():
    return _normalized_auto_code_mode(settings.get("auto_code_mode", "legacy"))


def _ml_artifact_dir():
    return os.path.join(_training_data_dir(), "ml_artifacts")

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
        return "Skráin fannst ekki."
    try:
        modified = datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%d.%m.%Y %H:%M")
        size = _format_file_size(os.path.getsize(path))
        return f"Vistaða útgáfan var lesin: {modified}" + (f" · {size}" if size else "")
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
        "auto_code_training_dir": _default_auto_code_training_dir(),
        "training_data_dir": _default_training_data_dir(),
        "auto_code_key_dir": _default_auto_code_key_dir(),
        "auto_code_mode": "legacy",
        "bank_restaurant_mode": False,
        "theme": "light",
        "recent_tok_files": [],
        "recent_bank_files": [],
        "yahoo_email": "",
        "yahoo_app_password": "",
        "email_start_date": "",
        "email_end_date": "",
        "email_print_folder": "",
    }
    try:
        if os.path.exists(SETTINGS_PATH):
            with open(SETTINGS_PATH, "r", encoding="utf-8") as handle:
                data = json.load(handle)
            defaults.update(data if isinstance(data, dict) else {})
    except Exception:
        pass
    defaults.pop("bank_output_dir", None)
    defaults.pop("compressed_output_dir", None)
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
    if not defaults.get("auto_code_training_dir"):
        defaults["auto_code_training_dir"] = _default_auto_code_training_dir()
    if not defaults.get("training_data_dir"):
        defaults["training_data_dir"] = _default_training_data_dir()
    defaults["training_data_dir"] = _ensure_folder(defaults.get("training_data_dir"), _default_training_data_dir())
    if not defaults.get("auto_code_key_dir"):
        defaults["auto_code_key_dir"] = _default_auto_code_key_dir()
    defaults["auto_code_mode"] = _normalized_auto_code_mode(defaults.get("auto_code_mode"))
    defaults["bank_restaurant_mode"] = bool(defaults.get("bank_restaurant_mode", False))
    if defaults.get("theme") not in THEMES:
        defaults["theme"] = "light"
    defaults["recent_tok_files"] = _clean_file_list(defaults.get("recent_tok_files", []))
    defaults["recent_bank_files"] = _clean_file_list(defaults.get("recent_bank_files", []))
    defaults.pop("recent_coded_files", None)
    return defaults

def save_settings():
    try:
        os.makedirs(os.path.dirname(SETTINGS_PATH), exist_ok=True)
        with open(SETTINGS_PATH, "w", encoding="utf-8") as handle:
            json.dump(settings, handle, indent=2)
    except Exception:
        pass

def _timing_value(setting_key, fallback, integer=False):
    value = runtime_timing.get(setting_key, settings.get(setting_key, fallback))
    try:
        value = float(value)
    except Exception:
        value = fallback
    return int(value) if integer else value

def _set_timing_value(setting_key, value, fallback, integer=False, persist=True):
    try:
        value = float(value)
    except Exception:
        value = fallback
    if integer:
        value = int(value)
    runtime_timing[setting_key] = value
    if persist:
        settings[setting_key] = value
    return value

def _sync_runtime_timing_from_settings():
    _set_timing_value("action_delay", settings.get("action_delay", 0.1), 0.1)
    _set_timing_value("focus_delay", settings.get("focus_delay", 0.05), 0.05)
    _set_timing_value("paste_delay", settings.get("paste_delay", 0.0), 0.0)
    _set_timing_value("post_paste_delay", settings.get("post_paste_delay", 0.05), 0.05)
    _set_timing_value("start_delay", settings.get("start_delay", 3), 3, integer=True)

def _after_ms(seconds):
    # Tkinter expects milliseconds. Always schedule at least 1ms.
    try:
        return max(1, int(float(seconds) * 1000))
    except Exception:
        return 1

def _cancel_scheduled_after():
    current_after_id = automation_state.get("after")
    if current_after_id is not None:
        try:
            root.after_cancel(current_after_id)
        except Exception:
            pass
        automation_state["after"] = None
    automation_state["next_due"] = None


def _schedule_next_action(delay_seconds):
    delay_seconds = max(0.001, float(delay_seconds))
    automation_state["next_due"] = time.monotonic() + delay_seconds
    automation_state["after"] = root.after(_after_ms(delay_seconds), process_next_action)


def _scheduler_recovery_delay():
    due = automation_state.get("next_due")
    automation_state["next_due"] = None
    if due is None:
        return 0
    lag = max(0.0, time.monotonic() - float(due))
    if lag < SCHEDULER_LAG_THRESHOLD_SECONDS:
        return 0
    automation_state["lag_events"] = int(automation_state.get("lag_events", 0)) + 1
    recovery = min(2.0, max(0.5, lag * 1.5))
    write_tok_debug_log(
        f"SCHEDULER_LAG lag_ms={int(lag * 1000)} recovery_ms={int(recovery * 1000)} "
        f"action='{_short_log_value(automation_state.get('current_action', ''))}'"
    )
    return recovery


def _foreground_window_info():
    if os.name != "nt":
        return 0, 0, ""
    try:
        user32 = ctypes.windll.user32
        hwnd = int(user32.GetForegroundWindow())
        if not hwnd:
            return 0, 0, ""
        process_id = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
        length = int(user32.GetWindowTextLengthW(hwnd))
        buffer = ctypes.create_unicode_buffer(max(1, length + 1))
        user32.GetWindowTextW(hwnd, buffer, len(buffer))
        return hwnd, int(process_id.value), buffer.value
    except Exception:
        return 0, 0, ""


def _capture_target_window():
    hwnd, process_id, title = _foreground_window_info()
    if os.name == "nt" and (not hwnd or process_id == os.getpid()):
        return False
    automation_state["target_hwnd"] = hwnd
    automation_state["target_pid"] = process_id
    automation_state["target_title"] = title
    automation_state["focus_mismatch_since"] = None
    automation_state["target_unresponsive_since"] = None
    write_tok_debug_log(
        f"TARGET_CAPTURED hwnd={hwnd} pid={process_id} title='{_short_log_value(title)}'"
    )
    return True


def _target_window_responsive(hwnd):
    if os.name != "nt" or not hwnd:
        return True
    try:
        result = wintypes.WPARAM()
        ctypes.set_last_error(0)
        sent = ctypes.windll.user32.SendMessageTimeoutW(
            int(hwnd),
            0,
            0,
            0,
            0x0002,
            75,
            ctypes.byref(result),
        )
        if sent:
            return True
        # Access-denied means Windows prevented the probe, not that Tok is hung.
        return ctypes.get_last_error() == 5
    except Exception:
        return True


def _target_focus_state():
    target_hwnd = int(automation_state.get("target_hwnd") or 0)
    if os.name != "nt" or not target_hwnd:
        return "ready"
    current_hwnd, current_pid, current_title = _foreground_window_info()
    if current_hwnd == target_hwnd:
        automation_state["focus_mismatch_since"] = None
        if _target_window_responsive(target_hwnd):
            automation_state["target_unresponsive_since"] = None
            return "ready"
        now = time.monotonic()
        unresponsive_since = automation_state.get("target_unresponsive_since")
        if unresponsive_since is None:
            automation_state["target_unresponsive_since"] = now
            write_tok_debug_log(
                f"TARGET_UNRESPONSIVE hwnd={target_hwnd} "
                f"action='{_short_log_value(automation_state.get('current_action', ''))}'"
            )
            return "waiting"
        if now - float(unresponsive_since) < TARGET_UNRESPONSIVE_TIMEOUT_SECONDS:
            return "waiting"
        return "unresponsive"

    now = time.monotonic()
    mismatch_since = automation_state.get("focus_mismatch_since")
    if mismatch_since is None:
        automation_state["focus_mismatch_since"] = now
        write_tok_debug_log(
            f"TARGET_FOCUS_CHANGED expected_hwnd={target_hwnd} actual_hwnd={current_hwnd} "
            f"actual_pid={current_pid} actual_title='{_short_log_value(current_title)}'"
        )
        return "waiting"
    if now - float(mismatch_since) < FOCUS_MISMATCH_GRACE_SECONDS:
        return "waiting"
    return "lost"


def _clear_target_window():
    automation_state["target_hwnd"] = 0
    automation_state["target_pid"] = 0
    automation_state["target_title"] = ""
    automation_state["focus_mismatch_since"] = None
    automation_state["target_unresponsive_since"] = None
    automation_state["current_action"] = ""

def _cancel_watchdog():
    current_watchdog = automation_state.get("watchdog")
    if current_watchdog:
        try:
            current_watchdog.cancel()
        except Exception:
            pass
    automation_state["watchdog"] = None

def _mark_automation_stopped(reason=""):
    global actions
    automation_state["active"] = False
    actions = []
    automation_state["stop_reason"] = reason
    automation_state["next_due"] = None

def _automation_is_running():
    return bool(automation_state.get("active"))

def _automation_stop_reason():
    return automation_state.get("stop_reason", "")

def _mouse_in_main_top_left():
    try:
        x, y = pyautogui.position()
        return 0 <= x <= MOUSE_FAILSAFE_MARGIN and 0 <= y <= MOUSE_FAILSAFE_MARGIN
    except Exception:
        return False

def _stop_for_mouse_failsafe():
    _mark_automation_stopped("Keyrsla stöðvaðist þar sem músin var færð efst í vinstra horn aðalskjásins.")
    _cancel_scheduled_after()
    _cancel_watchdog()

def _arm_watchdog():
    _cancel_watchdog()
    automation_state["watchdog"] = Timer(AUTOMATION_WATCHDOG_SECONDS, on_watchdog_timeout)
    automation_state["watchdog"].daemon = True
    automation_state["watchdog"].start()

def _sync_settings_from_ui():
    if "settings_theme_var" in globals():
        theme = settings_theme_var.get()
        if theme in THEMES:
            settings["theme"] = theme
    if "settings_action_entry" in globals() and settings_action_entry and settings_action_entry.winfo_exists():
        _set_timing_value("action_delay", settings_action_entry.get(), _timing_value("action_delay", 0.1))
    if "settings_focus_entry" in globals() and settings_focus_entry and settings_focus_entry.winfo_exists():
        _set_timing_value("focus_delay", settings_focus_entry.get(), _timing_value("focus_delay", 0.05))
    if "settings_paste_entry" in globals() and settings_paste_entry and settings_paste_entry.winfo_exists():
        _set_timing_value("paste_delay", settings_paste_entry.get(), _timing_value("paste_delay", 0.0))
    if "settings_post_paste_entry" in globals() and settings_post_paste_entry and settings_post_paste_entry.winfo_exists():
        _set_timing_value("post_paste_delay", settings_post_paste_entry.get(), _timing_value("post_paste_delay", 0.05))
    if "settings_start_entry" in globals() and settings_start_entry and settings_start_entry.winfo_exists():
        _set_timing_value("start_delay", settings_start_entry.get(), _timing_value("start_delay", 3, integer=True), integer=True)
    if "settings_auto_code_training_entry" in globals() and settings_auto_code_training_entry and settings_auto_code_training_entry.winfo_exists():
        training_dir = settings_auto_code_training_entry.get().strip()
        if training_dir:
            settings["auto_code_training_dir"] = training_dir
    if "settings_training_data_entry" in globals() and settings_training_data_entry and settings_training_data_entry.winfo_exists():
        training_data_dir = settings_training_data_entry.get().strip()
        if training_data_dir:
            settings["training_data_dir"] = training_data_dir
    if "settings_auto_code_mode_var" in globals():
        settings["auto_code_mode"] = _normalized_auto_code_mode(settings_auto_code_mode_var.get())
    if "settings_yahoo_email_entry" in globals() and settings_yahoo_email_entry and settings_yahoo_email_entry.winfo_exists():
        settings["yahoo_email"] = settings_yahoo_email_entry.get().strip()
    if "settings_yahoo_password_entry" in globals() and settings_yahoo_password_entry and settings_yahoo_password_entry.winfo_exists():
        settings["yahoo_app_password"] = settings_yahoo_password_entry.get().replace(" ", "")
    save_settings()

def _set_runtime_delays(values):
    _set_timing_value("action_delay", values.get("action"), _timing_value("action_delay", 0.1), persist=False)
    _set_timing_value("focus_delay", values.get("focus"), _timing_value("focus_delay", 0.05), persist=False)
    _set_timing_value("paste_delay", values.get("paste"), _timing_value("paste_delay", 0.0), persist=False)
    _set_timing_value("post_paste_delay", values.get("post_paste"), _timing_value("post_paste_delay", 0.05), persist=False)
    _set_timing_value("start_delay", values.get("start"), _timing_value("start_delay", 3, integer=True), integer=True, persist=False)

def apply_tok_run_speed_selection():
    global current_run_speed_label
    selected = "Saved"
    if "tok_run_speed_var" in globals() and tok_run_speed_var:
        selected = tok_run_speed_var.get() or "Saved"
    current_run_speed_label = selected
    if selected in SPEED_PRESET_MAP:
        _set_runtime_delays(SPEED_PRESET_MAP[selected])
    else:
        _sync_runtime_timing_from_settings()


def _window_work_area(window):
    """Return the work area for the monitor containing a Tk window."""
    if os.name == "nt":
        try:
            class RECT(ctypes.Structure):
                _fields_ = (
                    ("left", wintypes.LONG),
                    ("top", wintypes.LONG),
                    ("right", wintypes.LONG),
                    ("bottom", wintypes.LONG),
                )

            class MONITORINFO(ctypes.Structure):
                _fields_ = (
                    ("cbSize", wintypes.DWORD),
                    ("rcMonitor", RECT),
                    ("rcWork", RECT),
                    ("dwFlags", wintypes.DWORD),
                )

            monitor_from_window = ctypes.windll.user32.MonitorFromWindow
            monitor_from_window.argtypes = (wintypes.HWND, wintypes.DWORD)
            monitor_from_window.restype = ctypes.c_void_p
            monitor = monitor_from_window(int(window.winfo_id()), 2)  # nearest monitor
            info = MONITORINFO()
            info.cbSize = ctypes.sizeof(MONITORINFO)
            get_monitor_info = ctypes.windll.user32.GetMonitorInfoW
            get_monitor_info.argtypes = (ctypes.c_void_p, ctypes.POINTER(MONITORINFO))
            get_monitor_info.restype = wintypes.BOOL
            if monitor and get_monitor_info(monitor, ctypes.byref(info)):
                area = info.rcWork
                return int(area.left), int(area.top), int(area.right), int(area.bottom)
        except Exception:
            pass
    try:
        left = int(window.winfo_vrootx())
        top = int(window.winfo_vrooty())
        return left, top, left + int(window.winfo_vrootwidth()), top + int(window.winfo_vrootheight())
    except Exception:
        return 0, 0, int(window.winfo_screenwidth()), int(window.winfo_screenheight())

def ensure_window_fits():
    root.update_idletasks()
    if not root.winfo_ismapped():
        return
    req_w = root.winfo_reqwidth()
    req_h = root.winfo_reqheight()
    target_w = max(root.winfo_width(), req_w, 620)
    target_h = max(root.winfo_height(), req_h, 560)
    move_window_onscreen(target_w, target_h)

def fit_window_to_content():
    root.update_idletasks()
    if not root.winfo_ismapped():
        return
    # Page transitions may need more room, but should not undo a size chosen by
    # the user.  This makes resizing stable across every service, not only the
    # shareholder registry.
    target_w = max(root.winfo_width(), root.winfo_reqwidth(), 620)
    target_h = max(root.winfo_height(), root.winfo_reqheight(), 560)
    move_window_onscreen(target_w, target_h)

def move_window_onscreen(width=None, height=None):
    root.update_idletasks()
    try:
        if root.state() in {"zoomed", "fullscreen"}:
            return
    except tk.TclError:
        return
    left, top, right, bottom = _window_work_area(root)
    width = int(width if width is not None else root.winfo_width())
    height = int(height if height is not None else root.winfo_height())
    width = min(width, max(320, right - left - 32))
    height = min(height, max(320, bottom - top - 32))

    x = root.winfo_x()
    y = root.winfo_y()
    x = max(left + 16, min(x, right - width - 16))
    y = max(top + 8, min(y, bottom - height - 16))

    x_part = f"+{int(x)}" if int(x) >= 0 else str(int(x))
    y_part = f"+{int(y)}" if int(y) >= 0 else str(int(y))
    root.geometry(f"{width}x{height}{x_part}{y_part}")

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

def display_stopped_screen(reason=""):
    """Show a recovery-focused summary when the script is stopped."""
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Keyrsla stöðvuð", TEXT_ERROR)

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
    stopped_row = _tok_excel_row_for_position(current_index) if total_rows_count else 0
    rows_this_run = current_tok_rows_completed
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
    tk.Label(stopped_row_frame, text="Stöðvað við", bg=BG_CARD, fg=TEXT_SECONDARY,
             font=('Segoe UI', 9), width=18, anchor='w').pack(side=tk.LEFT)
    tk.Label(stopped_row_frame, text=f"Excel-línu {stopped_row}", bg=BG_CARD, fg=TEXT_ERROR,
             font=('Segoe UI', 10, 'bold'), anchor='w').pack(side=tk.LEFT, fill=tk.X, expand=True)
    log_icon = create_styled_button(stopped_row_frame, "Keyrsluskrá", open_tok_debug_log, width=10, accent=False, height=30)
    log_icon.pack(side=tk.RIGHT)
    add_info_row(summary, "Keyrslutími", format_timedelta(time_elapsed))

    if current_tok_file_path:
        file_row = tk.Frame(frame, bg=BG_DARK)
        file_row.pack(fill=tk.X, padx=28, pady=(0, 8))
        tok_input_btn = create_styled_button(file_row, "Keyra aftur inn", lambda path=current_tok_file_path: open_tok_input_with_file(path), width=16, accent=False)
        tok_input_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        open_btn = create_styled_button(file_row, "Opna skrá", lambda: os.startfile(current_tok_file_path), width=16, accent=False)
        open_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))

    back_button = create_styled_button(frame, "Til baka í valmynd", initialize_main_menu, width=20)
    back_button.pack(pady=(8, 14))
    finalize_page_layout()

def format_number(number):
    if pd.isna(number):
        return ""
    number_str = str(number).split('.')[0]
    if number_str.strip().lower() in {"nan", "none"}:
        return ""
    return number_str.replace(',', '')


def _tok_excel_row_number(index_value, position):
    try:
        return int(index_value) + 2
    except Exception:
        return int(position) + 2


def _tok_excel_row_for_position(position):
    try:
        if 0 <= int(position) < len(tok_excel_row_numbers):
            return int(tok_excel_row_numbers[int(position)])
    except Exception:
        pass
    return int(position) + 2


def _tok_last_excel_row():
    if tok_excel_row_numbers:
        return int(tok_excel_row_numbers[-1])
    if "rows" in globals() and rows:
        return len(rows) + 1
    return 0

def _send_unicode_text(text):
    """Type Unicode text with Win32 SendInput when the clipboard is unavailable."""
    if os.name != "nt":
        pyautogui.write(str(text))
        return

    ULONG_PTR = wintypes.WPARAM

    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", wintypes.WORD),
            ("wScan", wintypes.WORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", ULONG_PTR),
        ]

    class INPUT_UNION(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT)]

    class INPUT(ctypes.Structure):
        _anonymous_ = ("union",)
        _fields_ = [("type", wintypes.DWORD), ("union", INPUT_UNION)]

    key_events = []
    encoded = str(text).encode("utf-16-le")
    for index in range(0, len(encoded), 2):
        code_unit = int.from_bytes(encoded[index:index + 2], "little")
        key_events.append(INPUT(
            type=1,
            ki=KEYBDINPUT(0, code_unit, 0x0004, 0, 0),
        ))
        key_events.append(INPUT(
            type=1,
            ki=KEYBDINPUT(0, code_unit, 0x0004 | 0x0002, 0, 0),
        ))
    if not key_events:
        return
    event_array = (INPUT * len(key_events))(*key_events)
    sent = ctypes.windll.user32.SendInput(len(event_array), event_array, ctypes.sizeof(INPUT))
    if sent != len(event_array):
        raise ctypes.WinError()


def copy_to_clipboard_verified(text, field="", row_label="?"):
    expected = "" if text is None else str(text)
    last_value = ""
    last_error = ""
    for attempt in range(8):
        try:
            pyperclip.copy(expected)
        except Exception as exc:
            last_error = str(exc)
            time.sleep(0.025 if attempt < 3 else 0.06)
            continue
        # Give Windows clipboard a tiny chance to settle on older machines.
        time.sleep(0.015 if attempt < 3 else 0.035)
        try:
            last_value = pyperclip.paste()
        except Exception as exc:
            last_error = str(exc)
            last_value = ""
        if last_value == expected:
            if attempt:
                write_tok_debug_log(f"CLIPBOARD_VERIFIED_AFTER_RETRY row={row_label} field={field} attempts={attempt + 1}")
            return True
    write_tok_debug_log(
        f"CLIPBOARD_VERIFY_FAILED row={row_label} field={field} "
        f"expected='{_short_log_value(expected)}' actual='{_short_log_value(last_value)}' "
        f"error='{_short_log_value(last_error)}'"
    )
    return False

def enter_data(row, row_number=None):
    date_str = row['DATE'].strftime('%Y-%m-%d') if isinstance(row['DATE'], pd.Timestamp) else row['DATE']
    actions = []
    field_delay = 0.03
    focus_wait = _timing_value("focus_delay", 0.05)
    paste_wait = _timing_value("paste_delay", 0.0)
    post_paste_wait = _timing_value("post_paste_delay", 0.05)
    row_label = row_number if row_number is not None else "?"

    def _step(func, delay=USE_ACTION_DELAY, debug=""):
        return (func, delay, debug)

    def _press_enter(debug_label="ENTER"):
        pyautogui.press('enter')

    def _paste_text_steps(field, text):
        # Split into multiple steps to keep the UI responsive and allow Stop mid-row.
        text = "" if text is None else str(text)
        paste_state = {"clipboard": False}

        def prepare_text():
            paste_state["clipboard"] = copy_to_clipboard_verified(text, field, row_label)

        def enter_text():
            if paste_state["clipboard"]:
                pyautogui.hotkey("ctrl", "v")
                return
            write_tok_debug_log(
                f"UNICODE_INPUT_FALLBACK row={row_label} field={field} "
                f"value='{_short_log_value(text)}'"
            )
            _send_unicode_text(text)

        return [
            _step(prepare_text, paste_wait, f"row={row_label} field={field} copy value='{_short_log_value(text)}'"),
            _step(enter_text, post_paste_wait, f"row={row_label} field={field} paste"),
        ]

    actions.extend(_paste_text_steps("DATE", date_str))
    actions.append(_step(lambda: _press_enter("after DATE"), focus_wait, f"row={row_label} after DATE enter"))
    actions.append(_step(lambda: None, field_delay, f"row={row_label} after DATE wait"))

    actions.extend(_paste_text_steps("TEXT", str(row['TEXT'])))
    actions.append(_step(lambda: _press_enter("after TEXT 1"), focus_wait, f"row={row_label} after TEXT enter 1"))
    actions.append(_step(lambda: _press_enter("after TEXT 2"), focus_wait, f"row={row_label} after TEXT enter 2"))
    actions.append(_step(lambda: None, field_delay, f"row={row_label} after TEXT wait"))

    actions.extend(_paste_text_steps("DEBIT", format_number(row['DEBIT'])))
    actions.append(_step(lambda: _press_enter("after DEBIT"), focus_wait, f"row={row_label} after DEBIT enter"))
    actions.append(_step(lambda: None, field_delay, f"row={row_label} after DEBIT wait"))

    if pd.notna(row['ID']):
        actions.extend(_paste_text_steps("ID_DEBIT_SIDE", format_number(row['ID'])))
    actions.append(_step(lambda: _press_enter("after ID debit side"), focus_wait, f"row={row_label} after debit-side ID enter"))
    actions.append(_step(lambda: None, field_delay, f"row={row_label} after debit-side ID wait"))

    actions.extend(_paste_text_steps("AMOUNT", format_number(row['AMOUNT'])))
    actions.append(_step(lambda: _press_enter("after AMOUNT 1"), focus_wait, f"row={row_label} after AMOUNT enter 1"))
    actions.append(_step(lambda: _press_enter("after AMOUNT 2"), focus_wait, f"row={row_label} after AMOUNT enter 2"))
    actions.append(_step(lambda: _press_enter("after AMOUNT 3"), focus_wait, f"row={row_label} after AMOUNT enter 3"))
    actions.append(_step(lambda: _press_enter("after AMOUNT 4"), focus_wait, f"row={row_label} after AMOUNT enter 4"))
    actions.append(_step(lambda: None, field_delay, f"row={row_label} after AMOUNT wait"))
    actions.extend(_paste_text_steps("CREDIT", format_number(row['CREDIT'])))
    actions.append(_step(lambda: _press_enter("after CREDIT"), focus_wait, f"row={row_label} after CREDIT enter"))
    actions.append(_step(lambda: None, field_delay, f"row={row_label} after CREDIT wait"))

    if pd.notna(row['ID']):
        actions.extend(_paste_text_steps("ID_CREDIT_SIDE", format_number(row['ID'])))
    actions.append(_step(lambda: _press_enter("after credit-side ID 1"), focus_wait, f"row={row_label} after credit-side ID enter 1"))
    actions.append(_step(lambda: _press_enter("after credit-side ID 2"), focus_wait, f"row={row_label} after credit-side ID enter 2"))

    return actions

def init_progress_bar(total_rows):
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Vinnsla í gangi")

    global progress_bar, progress_label, stop_button_processing

    progress_label = create_styled_label(frame, "", size=10, color=TEXT_SECONDARY)
    progress_label.pack(pady=10)

    # Progress bar with custom style
    progress_bar = ttk.Progressbar(frame, orient='horizontal', length=350, mode='determinate',
                                   style='Horizontal.TProgressbar')
    progress_bar.pack(pady=15)
    progress_bar['maximum'] = total_rows
    progress_bar['value'] = 0
    
    stop_button_processing = create_styled_button(frame, "⏹ Stöðva", stop_script, width=15, accent=False)
    stop_button_processing.pack(pady=20)
    finalize_page_layout()

def update_progress(completed_count, active_index=None):
    maximum = int(float(progress_bar['maximum']))
    completed_count = max(0, min(int(completed_count), maximum))
    progress_bar['value'] = completed_count
    if active_index is None:
        text = ""
    else:
        text = (
            f"Vinn Excel-línu {_tok_excel_row_for_position(active_index)} "
            f"/ {_tok_last_excel_row()}"
        )
    progress_label.config(text=text)
    root.update_idletasks()

def display_results(time_elapsed, rows_processed):
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Lokið", TEXT_SUCCESS)
    
    card_shell, card = create_panel(frame, padx=24, pady=18)
    card_shell.pack(fill=tk.X, padx=28, pady=(4, 14))
    
    tk.Label(card, text=f"⏱ Tími: {format_timedelta(time_elapsed)}", bg=BG_CARD, fg=TEXT_PRIMARY,
             font=('Segoe UI', 11)).pack(pady=5)
    tk.Label(card, text=f"📊 Unnar línur: {rows_processed}", bg=BG_CARD, fg=TEXT_PRIMARY,
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
            "Opna skrá",
            lambda path=current_tok_file_path: os.startfile(path),
            width=16,
            accent=False,
        )
        open_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))

    action_row = tk.Frame(frame, bg=BG_DARK)
    action_row.pack(fill=tk.X, padx=28, pady=(0, 10))
    log_button = create_styled_button(action_row, "Keyrsluskrá", open_tok_debug_log, width=16, accent=False)
    log_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
    back_button = create_styled_button(action_row, "Til baka í valmynd", initialize_main_menu, width=16)
    back_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
    finalize_page_layout()

def format_timedelta(td):
    minutes, seconds = divmod(td.total_seconds(), 60)
    hours, minutes = divmod(minutes, 60)
    return f"{int(hours):02}:{int(minutes):02}:{seconds:.2f}"


def _countdown_status(seconds):
    seconds = int(seconds)
    unit = "sekúndu" if seconds == 1 else "sekúndur"
    return f"Keyrsla hefst eftir {seconds} {unit}..."

TOK_REQUIRED_COLUMNS = ['DATE', 'TEXT', 'DEBIT', 'ID', 'AMOUNT', 'CREDIT']

def _read_tok_id_as_text(value):
    """Keep IDs as text so Excel values such as 0123456789 retain their zero."""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)

def _format_missing_rows(missing_rows):
    if len(missing_rows) > 10:
        return ", ".join(str(r) for r in missing_rows[:10]) + f" (+{len(missing_rows) - 10} til viðbótar)"
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
            f"Ártal vantar í DATE í línum: {_format_missing_rows(invalid_rows)}"
        )
    if len(rows_by_year) > 1:
        details = "; ".join(
            f"{year}: línur {_format_missing_rows(row_numbers)}"
            for year, row_numbers in sorted(rows_by_year.items())
        )
        return False, f"Mismunandi ártöl fundust í DATE. Notaðu eitt ártal í hverri skrá. {details}"
    return True, ""

def prepare_tok_input_file(file_path):
    if not file_path:
        return {'ok': False, 'message': 'Veldu Excel-skrá fyrst.', 'df': None, 'missing_rows': []}

    try:
        df = pd.read_excel(file_path, converters={'ID': _read_tok_id_as_text})
    except Exception as exc:
        return {'ok': False, 'message': f"Ekki tókst að lesa skrána: {str(exc)[:80]}", 'df': None, 'missing_rows': []}

    df = df.dropna(how='all')
    if df.empty:
        return {'ok': False, 'message': 'Engar línur eru til vinnslu.', 'df': None, 'missing_rows': []}

    missing_cols = [col for col in TOK_REQUIRED_COLUMNS if col not in df.columns]
    if missing_cols:
        return {'ok': False, 'message': f"Dálka vantar: {', '.join(missing_cols)}", 'df': None, 'missing_rows': []}

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
        return {'ok': False, 'message': 'Engar línur eru til vinnslu.', 'df': None, 'missing_rows': []}

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
            'message': f"Skráin er samhæf, en DEBIT/CREDIT vantar í línum: {_format_missing_rows(missing_rows)}",
            'df': check_df,
            'missing_rows': missing_rows,
            'warning': True
        }

    ready_rows = len(check_df)
    return {'ok': True, 'message': f"Skráin er samhæf. Línur tilbúnar: {ready_rows}.", 'df': check_df, 'missing_rows': [], 'warning': False}

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
            tok_open_file_button.grid(row=0, column=1, sticky="e", padx=(8, 8), pady=6)
        else:
            tok_open_file_button.grid_remove()
    if "refresh_tok_status_layout" in globals():
        try:
            root.after_idle(refresh_tok_status_layout)
            root.after(60, refresh_tok_status_layout)
        except Exception:
            pass
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
    label = create_styled_label(tok_recent_files_frame, "Nýlegar skrár", size=9, color=TEXT_SECONDARY, bg=BG_CARD)
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

def reset_tok_start_row_controls():
    if "tok_start_mode_var" in globals() and tok_start_mode_var:
        try:
            tok_start_mode_var.set("first")
        except Exception:
            pass
    if "tok_start_row_entry" in globals() and tok_start_row_entry and tok_start_row_entry.winfo_exists():
        tok_start_row_entry.delete(0, tk.END)
        tok_start_row_entry.insert(0, "2")
    if "refresh_tok_start_row_controls" in globals():
        try:
            refresh_tok_start_row_controls()
        except Exception:
            pass


def _tok_selected_start_excel_row():
    if "tok_start_mode_var" not in globals() or tok_start_mode_var.get() != "custom":
        return 2
    raw = ""
    if "tok_start_row_entry" in globals() and tok_start_row_entry and tok_start_row_entry.winfo_exists():
        raw = tok_start_row_entry.get().strip()
    if not raw:
        return 2
    try:
        row_number = int(float(raw))
    except Exception:
        if "tok_status_label" in globals() and tok_status_label and tok_status_label.winfo_exists():
            tok_status_label.config(text="Upphafslína verður að vera tala.", fg=TEXT_WARNING)
        return None
    if row_number < 2:
        if "tok_status_label" in globals() and tok_status_label and tok_status_label.winfo_exists():
            tok_status_label.config(text="Fyrsta gagnalínan í Excel er lína 2.", fg=TEXT_WARNING)
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
        row_number = _tok_excel_row_number(index_value, position)
        available_rows.append((position, row_number))
        if row_number >= excel_row:
            return position
    last_row = available_rows[-1][1] if available_rows else 1
    if "tok_status_label" in globals() and tok_status_label and tok_status_label.winfo_exists():
        tok_status_label.config(text=f"Upphafslína {excel_row} er aftan við síðustu nothæfu línu ({last_row}).", fg=TEXT_WARNING)
    return None

def display_input_controls():
    global file_path_entry, browse_button, run_button, refresh_button, tok_drop_area, tok_status_label, tok_open_file_button
    global input_controls_frame, tok_recent_files_frame, tok_run_speed_var, tok_speed_buttons, tok_test_run_button
    global tok_start_mode_var, tok_start_row_entry, tok_start_buttons, tok_start_custom_frame, refresh_tok_start_row_controls
    global refresh_tok_status_layout

    input_controls_frame, inner_controls = create_panel(frame, padx=22, pady=18)
    input_controls_frame.pack(fill=tk.X, padx=28, pady=(0, 12))

    file_frame = tk.Frame(inner_controls, bg=BG_CARD)
    file_frame.pack(fill=tk.X, pady=(2, 8))
    
    drop_area = create_drop_box(file_frame, "Dragðu Tok-innsláttarskrá hingað eða smelltu til að velja", height=5)
    drop_area.pack(fill=tk.X)
    tok_drop_area = drop_area
    
    file_path_entry = create_styled_entry(file_frame, width=35)
    file_path_entry.pack_forget()

    tok_recent_files_frame = tk.Frame(inner_controls, bg=BG_CARD)
    tok_recent_files_frame.pack(fill=tk.X, pady=(0, 8))

    status_shell = RoundedPanel(inner_controls, fill=BUTTON_MUTED, outline=BORDER, radius=10, padx=0, pady=0)
    status_row = status_shell.inner
    status_shell.pack(fill=tk.X, pady=(0, 10))
    status_row.grid_columnconfigure(0, weight=1)
    tok_status_label = tk.Label(
        status_row,
        text="Veldu Excel-skrá fyrir Tok-innslátt til að kanna samhæfni.",
        bg=BUTTON_MUTED,
        fg=TEXT_SECONDARY,
        font=('Segoe UI', 9),
        padx=14,
        pady=7,
        anchor='w',
        justify='left',
    )
    tok_status_label.configure(wraplength=320, justify='left')
    tok_status_label.grid(row=0, column=0, sticky="nsew")
    tok_open_file_button = create_styled_button(
        status_row,
        "Opna skrá",
        lambda: os.startfile(file_path_entry.get().strip()) if file_path_entry.get().strip() else None,
        width=10,
        accent=False,
        height=42,
    )

    def sync_tok_input_controls_layout():
        if not input_controls_frame.winfo_exists():
            return
        inner_controls.update_idletasks()
        input_controls_frame._sync_height()
        ensure_window_fits()

    def refresh_tok_status_layout():
        if not status_row.winfo_exists():
            return
        status_row.update_idletasks()
        row_width = max(1, status_row.winfo_width())
        button_width = 0
        if tok_open_file_button.winfo_manager():
            button_width = tok_open_file_button.winfo_reqwidth() + 16
        available_width = max(150, row_width - button_width - 28)
        tok_status_label.configure(wraplength=available_width)
        status_row.update_idletasks()
        status_shell._sync_height()
        sync_tok_input_controls_layout()

    status_row.bind("<Configure>", lambda _event: root.after_idle(refresh_tok_status_layout))

    options_label = create_styled_label(inner_controls, "Keyrsluvalkostir", size=10, color=TEXT_PRIMARY, bold=True, bg=BG_CARD)
    options_label.pack(anchor='w', pady=(0, 6))

    speed_label = create_styled_label(inner_controls, "Hraði þessarar keyrslu", size=9, color=TEXT_SECONDARY, bg=BG_CARD)
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
        btn = create_styled_button(
            speed_row,
            SPEED_DISPLAY_LABELS.get(label, label),
            lambda value=label: set_run_speed(value),
            width=9,
            accent=False,
            height=40,
        )
        tok_speed_buttons[label] = btn
        btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6) if index < len(speed_labels) - 1 else (0, 0))
    refresh_speed_buttons()

    start_label = create_styled_label(inner_controls, "Upphafslína", size=9, color=TEXT_SECONDARY, bg=BG_CARD)
    start_label.pack(anchor='w', pady=(0, 6))
    start_row = tk.Frame(inner_controls, bg=BG_CARD)
    start_row.pack(fill=tk.X, pady=(0, 10))
    tok_start_mode_var = tk.StringVar(value="first")
    tok_start_buttons = {}

    start_button_row = tk.Frame(start_row, bg=BG_CARD)
    start_button_row.pack(fill=tk.X)
    start_button_row.grid_columnconfigure(0, weight=1, uniform="tok_option_actions")
    start_button_row.grid_columnconfigure(1, weight=1, uniform="tok_option_actions")
    tok_start_buttons["first"] = create_styled_button(
        start_button_row,
        "Byrja á fyrstu",
        lambda: set_tok_start_mode("first"),
        width=12,
        accent=False,
        height=42,
    )
    tok_start_buttons["first"].grid(row=0, column=0, sticky="ew", padx=(0, 8))
    tok_start_buttons["custom"] = create_styled_button(
        start_button_row,
        "Annað",
        lambda: set_tok_start_mode("custom"),
        width=12,
        accent=False,
        height=42,
    )
    tok_start_buttons["custom"].grid(row=0, column=1, sticky="ew", padx=(8, 0))

    tok_start_custom_frame = tk.Frame(start_row, bg=BG_CARD)
    tok_start_row_entry = create_styled_entry(tok_start_custom_frame, width=8)
    tok_start_row_entry.insert(0, "2")
    tok_start_row_entry.pack(side=tk.LEFT, ipady=4)

    def refresh_tok_start_row_controls():
        selected = tok_start_mode_var.get()
        for value, button in tok_start_buttons.items():
            set_button_accent(button, value == selected)
        if selected == "custom":
            if not tok_start_custom_frame.winfo_manager():
                tok_start_custom_frame.pack(fill=tk.X, pady=(8, 0))
        else:
            tok_start_custom_frame.pack_forget()
        sync_tok_input_controls_layout()
        root.after_idle(sync_tok_input_controls_layout)
        root.after(60, sync_tok_input_controls_layout)

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
            title="Veldu skrá",
            filetypes=(("Excel-skrár", "*.xlsx *.xls *.xlsm"), ("Allar skrár", "*.*"))
        )
        if filename:
            select_tok_input_file(filename)

    def on_drop_tok(path):
        if path:
            select_tok_input_file(path)

    if not attach_drop_target(drop_area, on_drop_tok):
        drop_area.config(text="Smelltu til að velja. Settu upp tkinterdnd2 til að draga og sleppa skrám.", fg=TEXT_WARNING)
    
    drop_area.bind("<Button-1>", lambda _event: browse_tok())

    def refresh_tok_file():
        path = file_path_entry.get().strip()
        if not path:
            tok_status_label.config(text="Veldu Excel-skrá fyrir Tok-innslátt áður en hún er endurlesin.", fg=TEXT_WARNING)
            if tok_open_file_button.winfo_manager():
                tok_open_file_button.grid_remove()
            root.after_idle(refresh_tok_status_layout)
            return
        select_tok_input_file(path, reset_start=False)

    btn_frame = tk.Frame(inner_controls, bg=BG_CARD)
    btn_frame.pack(fill=tk.X, pady=(2, 0))

    file_action_row = tk.Frame(btn_frame, bg=BG_CARD)
    file_action_row.pack(fill=tk.X, pady=(0, 7))
    file_action_row.grid_columnconfigure(0, weight=1, uniform="tok_option_actions")
    file_action_row.grid_columnconfigure(1, weight=1, uniform="tok_option_actions")
    refresh_button = create_styled_button(file_action_row, "Endurlesa skrá", refresh_tok_file, width=14, accent=False, height=42)
    refresh_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
    tok_test_run_button = create_styled_button(file_action_row, "Prufukeyrsla", lambda: set_run_speed("Test Run"), width=14, accent=False, height=42)
    tok_test_run_button.grid(row=0, column=1, sticky="ew", padx=(8, 0))

    run_button = create_styled_button(btn_frame, "Keyra", run_script_from_gui, width=20, height=42)
    run_button.pack(fill=tk.X)

    refresh_tok_recent_buttons()

def run_script_with_df(row_data, start_index=0, excel_row_numbers=None):
    global actions, rows, row_index, start_time, pending_tok_df, tok_excel_row_numbers
    automation_state["active"] = True
    automation_state["stop_reason"] = ""
    automation_state["lag_events"] = 0
    automation_state["current_action"] = ""
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

    rows = row_data if isinstance(row_data, list) else row_data.to_dict(orient='records')
    if excel_row_numbers is not None:
        tok_excel_row_numbers = list(excel_row_numbers)
    elif not isinstance(row_data, list):
        tok_excel_row_numbers = [
            _tok_excel_row_number(index_value, position)
            for position, index_value in enumerate(row_data.index)
        ]
    pending_tok_df = None
    gc.collect()

    total_rows = len(rows)
    if total_rows == 0:
        display_tok_input_error("Engar línur eru til vinnslu.")
        return

    start_index = max(0, min(int(start_index), total_rows - 1))
    start_time = datetime.datetime.now()
    if not _capture_target_window():
        _mark_automation_stopped(
            "Tok var ekki í forgrunni þegar niðurtalningu lauk. Settu Tok í forgrunn og keyrðu aftur."
        )
        display_stopped_screen(_automation_stop_reason())
        play_error_sound()
        return
    actions = enter_data(rows[start_index], _tok_excel_row_for_position(start_index))
    row_index = start_index
    write_tok_debug_log(f"ROW_START excel_row={_tok_excel_row_for_position(row_index)}")
    update_progress(current_tok_rows_completed, active_index=row_index)
    _cancel_scheduled_after()
    _cancel_watchdog()
    # Countdown already gave the user time to focus the target app.
    _schedule_next_action(0.1)

def display_tok_input_error(message):
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Ósamhæf skrá", TEXT_ERROR)
    msg_label = create_styled_label(frame, message, size=10, color=TEXT_SECONDARY)
    msg_label.pack(pady=(0, 15))
    back_button = create_styled_button(frame, "← Til baka í valmynd", initialize_main_menu, width=20)
    back_button.pack(pady=10)
    ensure_window_fits()

def display_tok_missing_fields_prompt(missing_rows, start_index=0):
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("DEBIT/CREDIT vantar", TEXT_WARNING)

    shown = _format_missing_rows(missing_rows)
    msg_label = create_styled_label(frame, f"DEBIT/CREDIT vantar í línum: {shown}", size=10, color=TEXT_SECONDARY)
    msg_label.pack(pady=(0, 15))

    def continue_anyway():
        if pending_tok_df is not None and pending_tok_file_path:
            start_tok_run(pending_tok_df, pending_tok_file_path, start_index=start_index)

    def open_file():
        if pending_tok_file_path:
            os.startfile(pending_tok_file_path)

    def refresh_file():
        if not pending_tok_file_path:
            msg_label.config(text="Engin skrá hefur verið valin til endurlesturs.", fg=TEXT_ERROR)
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
    continue_btn = create_styled_button(btn_frame, "Halda samt áfram", continue_anyway, width=20)
    continue_btn.pack(pady=5)
    refresh_btn = create_styled_button(btn_frame, "Endurlesa skrá", refresh_file, width=20, accent=False)
    refresh_btn.pack(pady=5)
    open_btn = create_styled_button(btn_frame, "Opna skrá", open_file, width=20, accent=False)
    open_btn.pack(pady=5)

    back_button = create_styled_button(frame, "← Til baka í valmynd", initialize_main_menu, width=20)
    back_button.pack(pady=10)
    ensure_window_fits()

def start_tok_run(df, file_path, start_index=0):
    global pending_tok_df, pending_tok_file_path, current_tok_file_path, rows, row_index
    global current_tok_run_start_index, current_tok_rows_completed, tok_excel_row_numbers
    _cancel_scheduled_after()
    _cancel_watchdog()
    _clear_target_window()
    automation_state["active"] = True
    automation_state["stop_reason"] = ""
    pending_tok_df = df
    pending_tok_file_path = file_path
    current_tok_file_path = file_path
    total_rows = len(df)
    if total_rows == 0:
        display_tok_input_error("Engar línur eru til vinnslu.")
        return

    run_rows = df.to_dict(orient='records')
    run_excel_rows = [
        _tok_excel_row_number(index_value, position)
        for position, index_value in enumerate(df.index)
    ]
    pending_tok_df = None
    del df
    gc.collect()

    start_index = max(0, min(int(start_index), total_rows - 1))
    current_tok_run_start_index = start_index
    current_tok_rows_completed = 0
    tok_excel_row_numbers = run_excel_rows
    rows = run_rows
    row_index = start_index
    action_wait = _timing_value("action_delay", 0.1)
    focus_wait = _timing_value("focus_delay", 0.05)
    paste_wait = _timing_value("paste_delay", 0.0)
    post_paste_wait = _timing_value("post_paste_delay", 0.05)
    countdown_seconds = _timing_value("start_delay", 3, integer=True)
    write_tok_debug_log(
        f"RUN_START file='{file_path}' total_rows={total_rows} "
        f"start_excel_row={_tok_excel_row_for_position(start_index)} "
        f"speed={current_run_speed_label} action={action_wait:.3f} focus={focus_wait:.3f} "
        f"paste={paste_wait:.3f} post_paste={post_paste_wait:.3f}"
    )
    rows_to_process = total_rows - start_index
    init_progress_bar(rows_to_process)
    countdown_label = create_styled_label(frame, _countdown_status(countdown_seconds), size=11, color=TEXT_PRIMARY)
    countdown_label.pack(pady=10)
    update_progress(0)

    def update_countdown(seconds_left):
        stop_reason = _automation_stop_reason()
        if not _automation_is_running() and stop_reason:
            display_stopped_screen(stop_reason)
            return
        countdown_label.config(text=_countdown_status(seconds_left))
        update_progress(0)
        if seconds_left > 0:
            automation_state["after"] = root.after(1000, update_countdown, seconds_left - 1)
        else:
            countdown_label.pack_forget()
            automation_state["after"] = root.after(
                100,
                lambda: run_script_with_df(
                    run_rows,
                    start_index=start_index,
                    excel_row_numbers=run_excel_rows,
                ),
            )

    update_countdown(countdown_seconds)

def process_next_action():
    global row_index, actions, rows, current_tok_rows_completed
    automation_state["after"] = None
    if not _automation_is_running():
        stop_reason = _automation_stop_reason()
        if stop_reason:
            display_stopped_screen(stop_reason)
        return

    _cancel_watchdog()

    recovery_delay = _scheduler_recovery_delay()
    if recovery_delay:
        _schedule_next_action(recovery_delay)
        return

    focus_state = _target_focus_state()
    if focus_state == "waiting":
        _schedule_next_action(0.1)
        return
    if focus_state in {"lost", "unresponsive"}:
        target_title = automation_state.get("target_title") or "Tok"
        if focus_state == "unresponsive":
            reason = f"Innsláttur stöðvaðist þar sem {target_title} svaraði ekki. Haltu áfram frá Excel-línunni sem birtist."
        else:
            reason = f"Innsláttur stöðvaðist þar sem {target_title} fór úr forgrunni. Haltu áfram frá Excel-línunni sem birtist."
        write_tok_debug_log(
            f"TARGET_GUARD_STOP state={focus_state} excel_row={_tok_excel_row_for_position(row_index)} "
            f"action='{_short_log_value(automation_state.get('current_action', ''))}'"
        )
        _mark_automation_stopped(reason)
        display_stopped_screen(reason)
        play_error_sound()
        return

    if actions:
        action = actions.pop(0)
        if len(action) == 3:
            func, delay, debug_text = action
        else:
            func, delay = action
            debug_text = ""
        if not _automation_is_running():
            return
        if _mouse_in_main_top_left():
            _stop_for_mouse_failsafe()
            display_stopped_screen(_automation_stop_reason())
            return
        _arm_watchdog()
        try:
            automation_state["current_action"] = debug_text
            if debug_text:
                write_tok_debug_log(f"ACTION {debug_text}")
            func()
        except pyautogui.FailSafeException:
            logging.info("PyAutoGUI fail-safe triggered.")
            _stop_for_mouse_failsafe()
            display_stopped_screen(_automation_stop_reason())
            return
        except Exception:
            logging.exception("Action failed; stopping automation.")
            _mark_automation_stopped("Keyrsla stöðvaðist eftir að innsláttaraðgerð mistókst.")
            _cancel_watchdog()
            display_stopped_screen(_automation_stop_reason())
            play_error_sound()
            return
        if not _automation_is_running():
            display_stopped_screen(_automation_stop_reason())
            return
        _arm_watchdog()
        next_delay = _timing_value("action_delay", 0.1) if delay is USE_ACTION_DELAY else delay
        _schedule_next_action(next_delay)
    else:
        write_tok_debug_log(f"ROW_DONE excel_row={_tok_excel_row_for_position(row_index)}")
        current_tok_rows_completed += 1
        row_index += 1
        if row_index >= len(rows):
            complete_script()
        else:
            actions = enter_data(rows[row_index], _tok_excel_row_for_position(row_index))
            write_tok_debug_log(f"ROW_START excel_row={_tok_excel_row_for_position(row_index)}")
            update_progress(current_tok_rows_completed, active_index=row_index)
            _arm_watchdog()
            _schedule_next_action(_timing_value("action_delay", 0.1))

def complete_script():
    global start_time, actions
    automation_state["active"] = False
    actions = []
    _cancel_watchdog()
    end_time = datetime.datetime.now()
    time_elapsed = end_time - start_time
    write_tok_debug_log(
        f"RUN_COMPLETE rows_processed={current_tok_rows_completed} "
        f"lag_events={automation_state.get('lag_events', 0)} "
        f"runtime={time_elapsed.total_seconds():.2f}s"
    )
    _clear_target_window()
    display_results(time_elapsed, current_tok_rows_completed)
    gc.collect()
    logging.info("Completed.")
    logging.info(f"Total time elapsed: {time_elapsed.total_seconds():.2f} seconds")
    logging.info(f"Number of rows processed: {current_tok_rows_completed}")

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
    current_value = float(settings_action_entry.get())
    new_value = current_value + 0.001
    settings_action_entry.delete(0, tk.END)
    settings_action_entry.insert(0, f"{new_value:.3f}")
    _sync_settings_from_ui()

def decrease_action_delay():
    current_value = float(settings_action_entry.get())
    new_value = max(0, current_value - 0.001)
    settings_action_entry.delete(0, tk.END)
    settings_action_entry.insert(0, f"{new_value:.3f}")
    _sync_settings_from_ui()

def increase_focus_delay():
    current_value = float(settings_focus_entry.get())
    new_value = current_value + 0.001
    settings_focus_entry.delete(0, tk.END)
    settings_focus_entry.insert(0, f"{new_value:.3f}")
    _sync_settings_from_ui()

def decrease_focus_delay():
    current_value = float(settings_focus_entry.get())
    new_value = max(0, current_value - 0.001)
    settings_focus_entry.delete(0, tk.END)
    settings_focus_entry.insert(0, f"{new_value:.3f}")
    _sync_settings_from_ui()

def increase_paste_delay():
    current_value = float(settings_paste_entry.get())
    new_value = current_value + 0.001
    settings_paste_entry.delete(0, tk.END)
    settings_paste_entry.insert(0, f"{new_value:.3f}")
    _sync_settings_from_ui()

def decrease_paste_delay():
    current_value = float(settings_paste_entry.get())
    new_value = max(0, current_value - 0.001)
    settings_paste_entry.delete(0, tk.END)
    settings_paste_entry.insert(0, f"{new_value:.3f}")
    _sync_settings_from_ui()

def increase_post_paste_delay():
    current_value = float(settings_post_paste_entry.get())
    new_value = current_value + 0.001
    settings_post_paste_entry.delete(0, tk.END)
    settings_post_paste_entry.insert(0, f"{new_value:.3f}")
    _sync_settings_from_ui()

def decrease_post_paste_delay():
    current_value = float(settings_post_paste_entry.get())
    new_value = max(0, current_value - 0.001)
    settings_post_paste_entry.delete(0, tk.END)
    settings_post_paste_entry.insert(0, f"{new_value:.3f}")
    _sync_settings_from_ui()

def initialize_main_menu():
    _sync_settings_from_ui()
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Veldu aðgerð")

    menu_host = tk.Frame(frame, bg=BG_DARK)
    menu_host.pack(fill=tk.X, pady=(12, 8))
    menu_host.columnconfigure(0, weight=1)
    menu_host.columnconfigure(2, weight=1)
    menu_card, menu_inner = create_panel(menu_host, padx=28, pady=20)
    menu_card.grid(row=0, column=1, sticky="ew")

    def resize_menu_column(event):
        # Keep the menu comfortably dense on wide/accounting workspaces while
        # still using nearly all available width on a small laptop display.
        available = max(1, int(event.width) - 56)
        menu_host.columnconfigure(1, minsize=min(780, available))

    menu_host.bind("<Configure>", resize_menu_column)

    registry_button = create_styled_button(
        menu_inner,
        "Hluthafaskrá",
        initialize_shareholder_registry,
        width=38,
    )
    registry_button.pack(fill=tk.X, pady=(0, 12))

    primary_row = tk.Frame(menu_inner, bg=BG_CARD)
    primary_row.pack(fill=tk.X, pady=(0, 10))
    tok_input_button = create_styled_button(primary_row, "Tok-innsláttur", initialize_tok_input, width=18)
    tok_input_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    bank_formatter_button = create_styled_button(primary_row, "Sníða bankafærslur", initialize_bank_formatter, width=18)
    bank_formatter_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

    tools_grid = tk.Frame(menu_inner, bg=BG_CARD)
    tools_grid.pack(fill=tk.X)
    tool_buttons = [
        ("Námunda tölur", initialize_round_numbers),
        ("Sníða dagsetningar", initialize_format_dates),
        ("Sníða kennitölur", initialize_format_ids),
        ("Þjappa færslum", coded_file_tools.initialize_compress_txs),
        ("Sækja tölvupóst", initialize_email_downloader),
        ("Prenta tölvupóst", initialize_folder_printer),
        ("Sameina skrár", coded_file_tools.initialize_combine_files),
        ("Stillingar", initialize_settings),
    ]
    for idx, (label, command) in enumerate(tool_buttons):
        btn = create_styled_button(tools_grid, label, command, width=17, accent=False)
        btn.grid(row=idx // 2, column=idx % 2, sticky='ew', padx=(0, 8) if idx % 2 == 0 else (8, 0), pady=5)
    tools_grid.columnconfigure(0, weight=1)
    tools_grid.columnconfigure(1, weight=1)

    def fit_main_menu_if_visible():
        """Fit after the rounded panel has reported its final requested size."""
        try:
            if not menu_card.winfo_exists() or not menu_card.winfo_manager():
                return
            # On first launch an idle callback can run before the window
            # manager has applied the requested 680x700 geometry; winfo then
            # reports 1x1 and an eager fit would collapse the window to its
            # minimum.  The timed passes below run once it is actually mapped.
            if not root.winfo_ismapped():
                return
            if hasattr(menu_card, "_sync_height"):
                menu_card._sync_height()
            root.update_idletasks()
            if root.state() in {"zoomed", "fullscreen"}:
                return
            target_w = max(root.winfo_width(), root.winfo_reqwidth(), 620)
            target_h = max(root.winfo_height(), root.winfo_reqheight(), 560)
            move_window_onscreen(target_w, target_h)
        except tk.TclError:
            return

    root.after_idle(fit_main_menu_if_visible)
    root.after(80, fit_main_menu_if_visible)
    root.after(220, fit_main_menu_if_visible)


def initialize_email_downloader():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Sækja viðhengi úr tölvupósti")

    page = tk.Frame(frame, bg=BG_DARK)
    page.pack(fill=tk.BOTH, expand=True, padx=12)

    controls_card, controls = create_panel(page, padx=20, pady=16)
    controls_card.pack(fill=tk.X, pady=(0, 10))
    controls.columnconfigure(0, weight=1)
    controls.columnconfigure(1, weight=1)

    today = datetime.date.today()
    default_start = settings.get("email_start_date") or today.replace(day=1).isoformat()
    default_end = settings.get("email_end_date") or today.isoformat()

    start_group = tk.Frame(controls, bg=BG_CARD)
    start_group.grid(row=0, column=0, sticky="ew", padx=(0, 8))
    create_styled_label(start_group, "Frá og með (ÁÁÁÁ-MM-DD)", size=9, color=TEXT_SECONDARY, bg=BG_CARD).pack(anchor="w")
    start_entry = create_styled_entry(start_group, width=18)
    start_entry.insert(0, default_start)
    start_entry.pack(fill=tk.X, pady=(5, 0), ipady=7)

    end_group = tk.Frame(controls, bg=BG_CARD)
    end_group.grid(row=0, column=1, sticky="ew", padx=(8, 0))
    create_styled_label(end_group, "Til og með (ÁÁÁÁ-MM-DD)", size=9, color=TEXT_SECONDARY, bg=BG_CARD).pack(anchor="w")
    end_entry = create_styled_entry(end_group, width=18)
    end_entry.insert(0, default_end)
    end_entry.pack(fill=tk.X, pady=(5, 0), ipady=7)

    destination = os.path.join(_training_data_dir(), "email downloads")
    destination_label = create_styled_label(
        controls,
        f"Vistun: {destination}",
        size=8,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    )
    destination_label.configure(anchor="w", justify="left", wraplength=520)
    destination_label.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))

    console_shell, console = create_styled_text_area(page, "Úttak keyrslu birtist hér", height=16)
    console_shell.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
    console.configure(font=("Consolas", 9), wrap="word")

    actions_row = tk.Frame(page, bg=BG_DARK)
    actions_row.pack(fill=tk.X)
    run_state = {"active": False, "cancel_event": None}
    output_queue = queue.Queue()

    def append_console(message):
        if not console.winfo_exists():
            return
        if getattr(console, "_placeholder_visible", False):
            console._clear_placeholder()
        console.insert(tk.END, f"{message}\n")
        console.see(tk.END)

    def finish_run(result, error):
        run_state["active"] = False
        run_state["cancel_event"] = None
        run_button.set_enabled(True)
        run_button.set_text("Keyra")
        back_button.set_enabled(True)
        cancel_button.set_enabled(False)
        cancel_button.set_text("Hætta við")
        if error:
            append_console("")
            append_console(f"Villa: {error}")
            play_error_sound()
        elif result and result.get("cancelled"):
            append_console("Niðurhali var stöðvað. Viðhengi sem þegar voru vistuð haldast í möppunni.")
            if result.get("output_dir"):
                open_button.command = lambda path=result["output_dir"]: open_output_folder(path)
                open_button.set_enabled(True)
        else:
            play_success_sound()
            if result and result.get("output_dir"):
                open_button.command = lambda path=result["output_dir"]: open_output_folder(path)
                open_button.set_enabled(True)

    def poll_worker(worker, holder):
        if not page.winfo_exists():
            return
        def drain_output():
            drained = False
            while True:
                try:
                    append_console(output_queue.get_nowait())
                    drained = True
                except queue.Empty:
                    return drained

        drain_output()
        if worker.is_alive():
            root.after(80, lambda: poll_worker(worker, holder))
            return
        # Drain once more after the worker has stopped to close the small race
        # where its final summary arrives between the first drain and is_alive().
        drain_output()
        finish_run(holder.get("result"), holder.get("error"))

    def open_output_folder(path):
        try:
            os.startfile(path)
        except Exception:
            try:
                open_file_location(path)
            except Exception:
                pass

    def cancel_download():
        cancel_event = run_state.get("cancel_event")
        if not run_state["active"] or cancel_event is None or cancel_event.is_set():
            return
        cancel_event.set()
        cancel_button.set_enabled(False)
        cancel_button.set_text("Stöðva...")
        append_console("Beiðni um stöðvun móttekin...")

    def run_download():
        if run_state["active"]:
            return
        start_value = start_entry.get().strip()
        end_value = end_entry.get().strip()
        try:
            email_downloader.parse_date_range(start_value, end_value)
        except ValueError as exc:
            append_console(f"Villa: {exc}")
            play_error_sound()
            return
        email_address = settings.get("yahoo_email", "").strip()
        app_password = settings.get("yahoo_app_password", "").replace(" ", "")
        if not email_address or not app_password:
            append_console("Villa: Skráðu Yahoo-netfang og app-lykilorð undir Stillingar.")
            play_error_sound()
            return

        settings["email_start_date"] = start_value
        settings["email_end_date"] = end_value
        save_settings()
        if getattr(console, "_placeholder_visible", False):
            console._clear_placeholder()
        console.delete("1.0", tk.END)
        run_state["active"] = True
        cancel_event = Event()
        run_state["cancel_event"] = cancel_event
        run_button.set_enabled(False)
        run_button.set_text("Sæki viðhengi...")
        back_button.set_enabled(False)
        open_button.set_enabled(False)
        cancel_button.set_text("Hætta við")
        cancel_button.set_enabled(True)
        holder = {}

        def worker_target():
            try:
                holder["result"] = email_downloader.download_inbox_attachments(
                    email_address=email_address,
                    app_password=app_password,
                    start_value=start_value,
                    end_value=end_value,
                    output_root=os.path.join(_training_data_dir(), "email downloads"),
                    emit=output_queue.put,
                    is_cancelled=cancel_event.is_set,
                )
            except Exception as exc:
                holder["error"] = str(exc)

        worker = Thread(target=worker_target, daemon=True)
        worker.start()
        poll_worker(worker, holder)

    run_button = create_styled_button(actions_row, "Keyra", run_download, width=11, accent=True)
    run_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
    cancel_button = create_styled_button(actions_row, "Hætta við", cancel_download, width=11, accent=False)
    cancel_button.set_enabled(False)
    cancel_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
    open_button = create_styled_button(actions_row, "Opna niðurstöður", lambda: None, width=16, accent=False)
    open_button.set_enabled(False)
    open_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
    back_button = create_styled_button(actions_row, "← Til baka", initialize_main_menu, width=11, accent=False)
    back_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
    ensure_window_fits()


def initialize_folder_printer():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Prenta skrár úr möppu")

    page = tk.Frame(frame, bg=BG_DARK)
    page.pack(fill=tk.BOTH, expand=True, padx=12)

    controls_card, controls = create_panel(page, padx=20, pady=16)
    controls_card.pack(fill=tk.X, pady=(0, 10))
    controls.columnconfigure(0, weight=1)

    create_styled_label(
        controls,
        "Mappa sem á að prenta",
        size=9,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    default_folder = settings.get("email_print_folder") or os.path.join(
        _training_data_dir(), "email downloads"
    )
    folder_entry = create_styled_entry(controls, width=42)
    folder_entry.insert(0, default_folder)
    folder_entry.grid(row=1, column=0, sticky="ew", pady=(5, 0), padx=(0, 8), ipady=7)

    def browse_folder():
        current = folder_entry.get().strip()
        initial = current if os.path.isdir(current) else os.path.join(_training_data_dir(), "email downloads")
        chosen = filedialog.askdirectory(initialdir=initial, title="Veldu möppu sem á að prenta")
        if chosen:
            folder_entry.delete(0, tk.END)
            folder_entry.insert(0, chosen)

    browse_button = create_styled_button(controls, "Velja möppu", browse_folder, width=14, accent=False, height=36)
    browse_button.grid(row=1, column=1, sticky="ew", pady=(5, 0))

    printer_hint = create_styled_label(
        controls,
        "Allar skrár beint í möppunni verða sendar á sjálfgefinn Windows-prentara. Undirmöppur eru ekki prentaðar.",
        size=8,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    )
    printer_hint.configure(anchor="w", justify="left", wraplength=520)
    printer_hint.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))

    console_shell, console = create_styled_text_area(page, "Úttak prentunar birtist hér", height=16)
    console_shell.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
    console.configure(font=("Consolas", 9), wrap="word")

    actions_row = tk.Frame(page, bg=BG_DARK)
    actions_row.pack(fill=tk.X)
    run_state = {"active": False}
    output_queue = queue.Queue()

    def append_console(message):
        if not console.winfo_exists():
            return
        if getattr(console, "_placeholder_visible", False):
            console._clear_placeholder()
        console.insert(tk.END, f"{message}\n")
        console.see(tk.END)

    def drain_output():
        while True:
            try:
                append_console(output_queue.get_nowait())
            except queue.Empty:
                return

    def finish_run(result, error):
        run_state["active"] = False
        run_state["cancel_event"] = None
        run_button.set_enabled(True)
        run_button.set_text("Keyra")
        back_button.set_enabled(True)
        browse_button.set_enabled(True)
        cancel_button.set_enabled(False)
        if error:
            append_console("")
            append_console(f"Villa: {error}")
            play_error_sound()
        elif result and result.get("cancelled"):
            append_console("Prentun var stöðvuð. Prentverk sem þegar voru send halda áfram í prentröðinni.")
        else:
            play_success_sound()

    def poll_worker(worker, holder):
        if not page.winfo_exists():
            return
        drain_output()
        if worker.is_alive():
            root.after(80, lambda: poll_worker(worker, holder))
            return
        drain_output()
        finish_run(holder.get("result"), holder.get("error"))

    def cancel_printing():
        cancel_event = run_state.get("cancel_event")
        if not run_state["active"] or cancel_event is None or cancel_event.is_set():
            return
        cancel_event.set()
        cancel_button.set_enabled(False)
        cancel_button.set_text("Stöðva...")
        append_console("Beiðni um stöðvun móttekin; engar fleiri skrár verða sendar eftir núverandi skrá.")

    def run_printing():
        if run_state["active"]:
            return
        selected_folder = folder_entry.get().strip()
        try:
            files = folder_printer.printable_files(selected_folder)
        except ValueError as exc:
            append_console(f"Villa: {exc}")
            play_error_sound()
            return
        if not files:
            append_console("Villa: Engar skrár fundust í valinni möppu.")
            play_error_sound()
            return
        if not messagebox.askyesno(
            "Staðfesta prentun",
            f"Senda {len(files)} skrár á sjálfgefinn Windows-prentara?",
            parent=root,
        ):
            return

        settings["email_print_folder"] = selected_folder
        save_settings()
        if getattr(console, "_placeholder_visible", False):
            console._clear_placeholder()
        console.delete("1.0", tk.END)
        run_state["active"] = True
        cancel_event = Event()
        run_state["cancel_event"] = cancel_event
        run_button.set_enabled(False)
        run_button.set_text("Prentun í gangi...")
        back_button.set_enabled(False)
        browse_button.set_enabled(False)
        cancel_button.set_text("Hætta við")
        cancel_button.set_enabled(True)
        holder = {}

        def worker_target():
            try:
                holder["result"] = folder_printer.print_folder_files(
                    selected_folder,
                    emit=output_queue.put,
                    is_cancelled=cancel_event.is_set,
                )
            except Exception as exc:
                holder["error"] = str(exc)

        worker = Thread(target=worker_target, daemon=True)
        worker.start()
        poll_worker(worker, holder)

    run_button = create_styled_button(actions_row, "Keyra", run_printing, width=14, accent=True)
    run_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
    cancel_button = create_styled_button(actions_row, "Hætta við", cancel_printing, width=14, accent=False)
    cancel_button.set_enabled(False)
    cancel_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
    back_button = create_styled_button(actions_row, "← Til baka", initialize_main_menu, width=18, accent=False)
    back_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
    ensure_window_fits()

def initialize_tok_input():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Tok-innsláttur")
    
    display_input_controls()
    
    nav_row = tk.Frame(frame, bg=BG_DARK)
    nav_row.pack(fill=tk.X, padx=28, pady=(0, 10))
    back_button = create_styled_button(nav_row, "← Til baka í valmynd", initialize_main_menu, width=18, accent=False)
    back_button.pack()
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
                os.startfile(folder if os.path.isdir(folder) else _training_data_dir())
            return
        except Exception:
            pass
    try:
        folder = os.path.dirname(file_path)
        os.startfile(folder if os.path.isdir(folder) else _training_data_dir())
    except Exception:
        pass

def initialize_settings():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Stillingar")

    # Reserve navigation space before the settings card is laid out. This keeps
    # both actions visible even on a short display or after resizing downward.
    bottom_row = tk.Frame(frame, bg=BG_DARK)
    bottom_row.pack(side=tk.BOTTOM, fill=tk.X, padx=28, pady=(2, 10))

    # Keep the primary actions fixed while the growing settings form scrolls.
    # A canvas viewport prevents the form's requested height from forcing the
    # main window beyond the monitor work area.
    settings_scroll = tk.Frame(frame, bg=BG_DARK)
    settings_scroll.pack(fill=tk.BOTH, expand=True)
    settings_canvas = tk.Canvas(
        settings_scroll,
        bg=BG_DARK,
        bd=0,
        highlightthickness=0,
        height=360,
        yscrollincrement=32,
    )
    settings_scrollbar = ttk.Scrollbar(
        settings_scroll,
        orient=tk.VERTICAL,
        command=settings_canvas.yview,
    )
    settings_canvas.configure(yscrollcommand=settings_scrollbar.set)
    settings_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(8, 0))
    settings_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    settings_content = tk.Frame(settings_canvas, bg=BG_DARK)
    settings_window = settings_canvas.create_window(
        (0, 0),
        window=settings_content,
        anchor="nw",
    )

    def _sync_settings_scroll_region(_event=None):
        if not settings_canvas.winfo_exists():
            return
        settings_canvas.configure(scrollregion=settings_canvas.bbox("all"))

    def _resize_settings_content(event):
        settings_canvas.itemconfigure(settings_window, width=max(1, event.width))
        _sync_settings_scroll_region()

    settings_canvas.bind("<Configure>", _resize_settings_content)
    settings_content.bind("<Configure>", _sync_settings_scroll_region)

    settings_card, settings_frame = create_panel(settings_content, padx=24, pady=22)
    settings_card.pack(fill=tk.X, padx=28, pady=(0, 12))

    help_label = create_styled_label(
        settings_frame,
        "Veldu Hægt ef Tok sleppir reitum eða gildi vantar. Veldu Hratt ef allt virkar rétt.",
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
    global settings_auto_code_training_entry, settings_auto_code_status_label
    global settings_training_data_entry
    global settings_auto_code_mode_var

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
    theme_label = create_styled_label(settings_frame, "Þema", size=10, bold=True)
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

    for label, theme_key in (("Ljóst", "light"), ("Dökkt", "dark")):
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

    preset_label = create_styled_label(settings_frame, "Hraðastilling", size=10, bold=True)
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
            SPEED_DISPLAY_LABELS.get(label, label),
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
        "+ Sýna ítarlegar stillingar",
        size=9,
        color=ACCENT,
        bold=True,
        bg=BG_CARD
    )
    manual_toggle_label.pack(side=tk.LEFT)
    manual_toggle_hint = create_styled_label(
        manual_toggle_inner,
        "Ítarlegt",
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
    _timing_row(0, "Milli aðgerða", "Heildarhraði. Hækkaðu gildið ef Tok fer í rangan reit.", settings_action_entry, decrease_action_delay, increase_action_delay)
    _timing_row(1, "Eftir Enter", "Bið eftir færslu milli reita. Hækkaðu gildið ef bendillinn nær ekki að fylgja.", settings_focus_entry, decrease_focus_delay, increase_focus_delay)
    _timing_row(2, "Fyrir límingu", "Bið eftir afritun á klippiborð. Yfirleitt skal hafa gildið nálægt núlli.", settings_paste_entry, decrease_paste_delay, increase_paste_delay)
    _timing_row(3, "Eftir límingu", "Mikilvægasta biðin þegar Tok er hægt. Hækkaðu gildið ef gildi vantar eða þau skerðast.", settings_post_paste_entry, decrease_post_paste_delay, increase_post_paste_delay)
    _timing_row(4, "Upphafsniðurtalning", "Tími til að velja Tok áður en innsláttur hefst.", settings_start_entry, decrease_start_delay, increase_start_delay)
    _refresh_preset_buttons()

    def _toggle_manual_tuning(_event=None):
        if manual_section_open.get():
            manual_grid.pack_forget()
            manual_section_open.set(False)
            manual_toggle_label.config(text="+ Sýna ítarlegar stillingar")
            manual_toggle_hint.config(text="Ítarlegt")
        else:
            manual_grid.pack(fill=tk.X, pady=(0, 8), before=training_frame)
            manual_section_open.set(True)
            manual_toggle_label.config(text="- Fela ítarlegar stillingar")
            manual_toggle_hint.config(text="Biðtímastillingar")
        _sync_settings_scroll_region()
        schedule_layout_refresh()
        ensure_window_fits()

    for widget in (manual_toggle_frame, manual_toggle_inner, manual_toggle_label, manual_toggle_hint):
        widget.bind("<Button-1>", _toggle_manual_tuning)

    settings_auto_code_mode_var = tk.StringVar(value=_auto_code_mode())
    auto_code_mode_setting = ui_components.create_segmented_setting(
        settings_frame,
        "Aðferð sjálfvirkrar kóðunar",
        settings_auto_code_mode_var,
        [("Hefðbundin", "legacy"), ("ML-aðstoð", "ml_hybrid")],
        hint="ML-aðstoð breytir aðeins óvissum færslum og notar hefðbundna kóðun ef gilt líkan vantar.",
        stacked=True,
    )
    auto_code_mode_setting.pack(fill=tk.X, pady=(12, 4))

    if ML_STATUS_AVAILABLE and ml_runtime_status is not None:
        ml_health = ml_runtime_status(_ml_artifact_dir())
    else:
        ml_health = {"available": False}
    if ml_health.get("available"):
        ml_health_text = (
            f"ML-líkan virkt: {ml_health.get('model_version', '')} · "
            f"{int(ml_health.get('training_rows', 0))} þjálfunarlínur."
        )
        ml_health_color = TEXT_SUCCESS
    else:
        ml_health_text = "ML-líkan er ekki virkt. ML-aðstoð notar því hefðbundna kóðun."
        ml_health_color = TEXT_WARNING
    ml_health_label = create_styled_label(
        settings_frame,
        ml_health_text,
        size=8,
        color=ml_health_color,
        bg=BG_CARD,
    )
    ml_health_label.configure(wraplength=420, justify="left")
    ml_health_label.pack(anchor="w", pady=(0, 4))

    # Auto-code training folder
    training_frame = tk.Frame(settings_frame, bg=BG_CARD)
    training_frame.pack(fill=tk.X, pady=(12, 4))
    training_label = create_styled_label(training_frame, "Þjálfunarmappa sjálfvirkrar kóðunar:", size=9, color=TEXT_SECONDARY)
    training_label.pack(anchor='w')

    settings_auto_code_training_entry = create_styled_entry(training_frame, width=42)
    settings_auto_code_training_entry.insert(0, settings.get("auto_code_training_dir", _default_auto_code_training_dir()))
    settings_auto_code_training_entry.pack(side=tk.LEFT, padx=(0, 8), pady=5, ipady=5, fill=tk.X, expand=True)

    def browse_training_dir():
        current_dir = settings_auto_code_training_entry.get().strip() or os.path.expanduser("~")
        chosen = filedialog.askdirectory(initialdir=current_dir, title="Veldu þjálfunarmöppu sjálfvirkrar kóðunar")
        if chosen:
            settings_auto_code_training_entry.delete(0, tk.END)
            settings_auto_code_training_entry.insert(0, chosen)
            _sync_settings_from_ui()
            schedule_layout_refresh()

    training_browse_btn = create_styled_button(training_frame, "Velja möppu", browse_training_dir, width=15, accent=False)
    training_browse_btn.pack(side=tk.LEFT)

    # Raw/coded training archive folder
    training_data_frame = tk.Frame(settings_frame, bg=BG_CARD)
    training_data_frame.pack(fill=tk.X, pady=(10, 4))
    training_data_label = create_styled_label(training_data_frame, "Safnmappa þjálfunargagna:", size=9, color=TEXT_SECONDARY)
    training_data_label.pack(anchor='w')

    settings_training_data_entry = create_styled_entry(training_data_frame, width=42)
    settings_training_data_entry.insert(0, settings.get("training_data_dir", _default_training_data_dir()))
    settings_training_data_entry.pack(side=tk.LEFT, padx=(0, 8), pady=5, ipady=5, fill=tk.X, expand=True)

    def browse_training_data_dir():
        current_dir = settings_training_data_entry.get().strip() or os.path.expanduser("~")
        chosen = filedialog.askdirectory(initialdir=current_dir, title="Veldu safnmöppu þjálfunargagna")
        if chosen:
            settings_training_data_entry.delete(0, tk.END)
            settings_training_data_entry.insert(0, chosen)
            _sync_settings_from_ui()
            schedule_layout_refresh()

    training_data_browse_btn = create_styled_button(training_data_frame, "Velja möppu", browse_training_data_dir, width=15, accent=False)
    training_data_browse_btn.pack(side=tk.LEFT)

    # Keep credentials in a compact dialog so the settings page remains short.
    yahoo_frame = tk.Frame(settings_frame, bg=BG_CARD)
    yahoo_frame.pack(fill=tk.X, pady=(10, 6))
    yahoo_copy = tk.Frame(yahoo_frame, bg=BG_CARD)
    yahoo_copy.pack(side=tk.LEFT, fill=tk.X, expand=True)
    create_styled_label(yahoo_copy, "Yahoo-tölvupóstur", size=9, bold=True, bg=BG_CARD).pack(anchor="w")
    yahoo_status = create_styled_label(
        yahoo_copy,
        "Innskráning vistuð" if settings.get("yahoo_email") and settings.get("yahoo_app_password") else "Innskráning ekki skráð",
        size=8,
        color=TEXT_SUCCESS if settings.get("yahoo_email") and settings.get("yahoo_app_password") else TEXT_SECONDARY,
        bg=BG_CARD,
    )
    yahoo_status.pack(anchor="w", pady=(2, 0))

    def open_yahoo_credentials():
        dialog = tk.Toplevel(root)
        dialog.title("Yahoo-innskráning")
        dialog.configure(bg=BG_DARK)
        dialog.transient(root)
        dialog.grab_set()

        card, content = create_panel(dialog, padx=22, pady=20)
        card.pack(fill=tk.BOTH, expand=True, padx=18, pady=18)
        create_styled_label(content, "Yahoo-innskráning", size=13, bold=True, bg=BG_CARD).pack(anchor="w")
        hint = create_styled_label(
            content,
            "Sláðu þetta aðeins inn einu sinni. Notaðu Yahoo app-lykilorð, ekki venjulega lykilorðið.",
            size=8,
            color=TEXT_SECONDARY,
            bg=BG_CARD,
        )
        hint.configure(wraplength=430, justify="left")
        hint.pack(anchor="w", pady=(3, 12))

        create_styled_label(content, "Netfang", size=9, color=TEXT_SECONDARY, bg=BG_CARD).pack(anchor="w")
        email_entry = create_styled_entry(content, width=42)
        email_entry.insert(0, settings.get("yahoo_email", ""))
        email_entry.pack(fill=tk.X, pady=(4, 10), ipady=7)

        create_styled_label(content, "App-lykilorð", size=9, color=TEXT_SECONDARY, bg=BG_CARD).pack(anchor="w")
        password_row = tk.Frame(content, bg=BG_CARD)
        password_row.pack(fill=tk.X, pady=(4, 14))
        password_entry = create_styled_entry(password_row, width=34)
        password_entry.configure(show="•")
        password_entry.insert(0, settings.get("yahoo_app_password", ""))
        password_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=7)
        visible = tk.BooleanVar(value=False)

        def toggle_password():
            visible.set(not visible.get())
            password_entry.configure(show="" if visible.get() else "•")
            reveal_button.set_text("Fela" if visible.get() else "Sýna")

        reveal_button = create_styled_button(password_row, "Sýna", toggle_password, width=7, accent=False, height=36)
        reveal_button.pack(side=tk.LEFT, padx=(8, 0))

        dialog_status = create_styled_label(content, "", size=8, color=TEXT_ERROR, bg=BG_CARD)
        dialog_status.pack(fill=tk.X, pady=(0, 6))

        def save_credentials():
            email_value = email_entry.get().strip()
            password_value = password_entry.get().replace(" ", "")
            if not email_value or not password_value:
                dialog_status.config(text="Netfang og app-lykilorð verða bæði að vera skráð.", fg=TEXT_ERROR)
                return
            settings["yahoo_email"] = email_value
            settings["yahoo_app_password"] = password_value
            save_settings()
            yahoo_status.config(text="Innskráning vistuð", fg=TEXT_SUCCESS)
            dialog.destroy()

        action_row = tk.Frame(content, bg=BG_CARD)
        action_row.pack(fill=tk.X)
        create_styled_button(action_row, "Vista", save_credentials, width=13, accent=True).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        create_styled_button(action_row, "Hætta við", dialog.destroy, width=13, accent=False).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
        dialog.bind("<Return>", lambda _event: save_credentials())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        fit_dialog_to_content(dialog, min_width=500, min_height=390, preferred_width=540)
        email_entry.focus_set()

    create_styled_button(
        yahoo_frame,
        "Yahoo-innskráning",
        open_yahoo_credentials,
        width=19,
        accent=False,
        height=38,
    ).pack(side=tk.RIGHT, padx=(12, 0))

    archive_sync_frame = tk.Frame(settings_frame, bg=BG_CARD)
    archive_sync_frame.pack(fill=tk.X, pady=(4, 8))
    archive_sync_status = create_styled_label(
        archive_sync_frame,
        "",
        size=8,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    )
    archive_sync_status.configure(wraplength=420, justify="left")
    archive_sync_status.pack(fill=tk.X, pady=(0, 6))
    archive_sync_state = {"active": False}

    def finish_archive_sync(result):
        archive_sync_state["active"] = False
        if not archive_sync_frame.winfo_exists():
            return
        archive_sync_btn.set_text("Samstilla pöruð þjálfunargögn")
        archive_sync_btn.command = run_archive_sync
        total = int(result.get("total", 0))
        synced = int(result.get("synced", 0))
        failed = int(result.get("failed", 0))
        if failed:
            message = f"Samstillingu lauk með villum. Samstilltar möppur: {synced} af {total}; mistókst: {failed}."
            message += " Möppur sem mistókust innihalda paired_rows_sync_error.txt."
            first_failure = (result.get("failures") or [{}])[0]
            if first_failure.get("session"):
                message += f" Fyrsta villa: {first_failure['session']}."
            color = TEXT_WARNING
        elif total == 0:
            message = "Engin pöruð þjálfunargögn fundust til samstillingar."
            color = TEXT_WARNING
        else:
            message = f"Samstillingu lokið. Samstilltar möppur: {synced} af {total}."
            color = TEXT_SUCCESS
        archive_sync_status.config(text=message, fg=color)
        settings_frame.update_idletasks()
        settings_card._sync_height()
        ensure_window_fits()

    def poll_archive_sync(worker, result_holder):
        if worker.is_alive():
            root.after(100, lambda: poll_archive_sync(worker, result_holder))
            return
        finish_archive_sync(result_holder.get("result", {
            "ok": False,
            "total": 0,
            "synced": 0,
            "failed": 1,
            "failures": [],
            "message": "Samstilling þjálfunargagna mistókst óvænt.",
        }))

    def run_archive_sync():
        if archive_sync_state["active"]:
            return
        _sync_settings_from_ui()
        archive_root = _training_data_dir()
        archive_sync_state["active"] = True
        archive_sync_status.config(text="Samstilli þjálfunargögn...", fg=TEXT_SECONDARY)
        archive_sync_btn.set_text("Samstilling í gangi...")
        archive_sync_btn.command = lambda: None
        result_holder = {}

        def worker_target():
            try:
                result_holder["result"] = sync_all_training_archives(archive_root)
            except Exception as exc:
                result_holder["result"] = {
                    "ok": False,
                    "total": 0,
                    "synced": 0,
                    "failed": 1,
                    "failures": [],
                    "message": f"Samstilling þjálfunargagna mistókst: {str(exc)[:100]}",
                }

        worker = Thread(target=worker_target, daemon=True)
        worker.start()
        poll_archive_sync(worker, result_holder)

    archive_sync_btn = create_styled_button(
        archive_sync_frame,
        "Samstilla pöruð þjálfunargögn",
        run_archive_sync,
        width=24,
        accent=False,
    )
    archive_sync_btn.pack(anchor="w")

    memory_frame = tk.Frame(settings_frame, bg=BG_CARD)
    memory_frame.pack(fill=tk.X, pady=(2, 10))
    settings_auto_code_status_label = create_styled_label(
        memory_frame,
        f"Minnisskrá: {os.path.join(settings.get('auto_code_training_dir', _default_auto_code_training_dir()), 'auto_code_memory.json')}",
        size=8,
        color=TEXT_SECONDARY,
        bg=BG_CARD
    )
    settings_auto_code_status_label.configure(wraplength=420, justify='left')
    settings_auto_code_status_label.pack(anchor='w', pady=(0, 6))

    def rebuild_auto_code_memory():
        _sync_settings_from_ui()
        training_dir = _auto_code_training_dir()
        training_data_dir = _training_data_dir()
        memory_path = _auto_code_memory_path()
        if not AUTO_CODER_AVAILABLE or get_auto_coder is None:
            settings_auto_code_status_label.config(text=f"Sjálfvirk kóðun er ekki tiltæk: {AUTO_CODER_ERROR[:80]}", fg=TEXT_ERROR)
            return
        try:
            coder = get_auto_coder(
                force_reload=True,
                force_rebuild=True,
                training_dir=training_dir,
                key_dir=_auto_code_key_dir(),
                memory_path=memory_path,
                training_data_dir=training_data_dir,
            )
            summary = coder.summary()
            if coder.ready():
                settings_auto_code_status_label.config(
                    text=(
                        f"Minnisskrá endurbyggð: {summary['training_rows']} línur, "
                        f"{summary['merchant_keys']} söluaðilar. Vistað í: {memory_path}"
                    ),
                    fg=TEXT_SUCCESS
                )
            else:
                settings_auto_code_status_label.config(
                    text=f"Engar nothæfar kóðaðar línur fundust í: {training_dir} eða {training_data_dir}",
                    fg=TEXT_WARNING
                )
        except Exception as exc:
            settings_auto_code_status_label.config(text=f"Endurbygging mistókst: {str(exc)[:90]}", fg=TEXT_ERROR)

    rebuild_memory_btn = create_styled_button(memory_frame, "Endurbyggja minni sjálfvirkrar kóðunar", rebuild_auto_code_memory, width=31, accent=False)
    rebuild_memory_btn.pack(anchor='w')

    def _save_settings_and_refresh():
        _sync_settings_from_ui()
        _refresh_preset_buttons()
        _sync_settings_scroll_region()
        schedule_layout_refresh()

    save_btn = create_styled_button(bottom_row, "Vista stillingar", _save_settings_and_refresh, width=16, accent=True)
    save_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

    back_button = create_styled_button(bottom_row, "← Til baka", initialize_main_menu, width=14, accent=False)
    back_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

    def _scroll_settings(event):
        if not settings_canvas.winfo_exists():
            return None
        if getattr(event, "num", None) == 4:
            units = -1
        elif getattr(event, "num", None) == 5:
            units = 1
        else:
            delta = int(getattr(event, "delta", 0) or 0)
            if not delta:
                return None
            step = max(1, abs(delta) // 120)
            units = -step if delta > 0 else step
        settings_canvas.yview_scroll(units, "units")
        return "break"

    def _bind_settings_wheel(widget):
        widget.bind("<MouseWheel>", _scroll_settings, add="+")
        widget.bind("<Button-4>", _scroll_settings, add="+")
        widget.bind("<Button-5>", _scroll_settings, add="+")
        for child in widget.winfo_children():
            _bind_settings_wheel(child)

    _bind_settings_wheel(settings_scroll)
    settings_canvas.yview_moveto(0.0)
    settings_canvas.after_idle(_sync_settings_scroll_region)
    ensure_window_fits()

def initialize_round_numbers():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Námunda tölur")

    container = tk.Frame(frame, bg=BG_DARK)
    container.pack(fill=tk.X, padx=28)

    input_label = create_styled_label(container, "Límdu inn tölur (eina í hverja línu):", size=9, color=TEXT_SECONDARY)
    input_label.pack(anchor='w')

    input_shell, input_text = create_styled_text_area(container, "Límdu tölurnar hér, eina í hverja línu", height=6)
    input_shell.pack(fill=tk.X, pady=(5, 12))
    input_resize = attach_auto_resize_text(input_text, min_lines=6, max_lines=8)

    output_label = create_styled_label(container, "Námunduð niðurstaða:", size=9, color=TEXT_SECONDARY)
    output_label.pack(anchor='w')

    output_shell, output_text = create_styled_text_area(container, "Námundaðar tölur birtast hér", height=6)
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
            status_label.config(text=f"Fjöldi ógildra lína sem var sleppt: {bad}.", fg=TEXT_WARNING)
        else:
            status_label.config(text="Tölurnar voru námundaðar.", fg=TEXT_SUCCESS)

    def copy_output():
        text = get_text_area_value(output_text).strip()
        if text:
            pyperclip.copy(text)
            status_label.config(text="Afritað á klippiborð (ein færsla í hverri línu).", fg=TEXT_SUCCESS)
        else:
            status_label.config(text="Engin niðurstaða er til að afrita.", fg=TEXT_WARNING)

    btn_row = tk.Frame(container, bg=BG_DARK)
    btn_row.pack(pady=5)

    round_btn = create_styled_button(btn_row, "Námunda", round_numbers, width=12)
    round_btn.pack(side=tk.LEFT, padx=5)

    copy_btn = create_styled_button(btn_row, "Afrita niðurstöðu", copy_output, width=16, accent=False)
    copy_btn.pack(side=tk.LEFT, padx=5)

    back_button = create_styled_button(frame, "← Til baka í valmynd", initialize_main_menu, width=20)
    back_button.pack(pady=10)
    ensure_window_fits()

def initialize_format_dates():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Sníða dagsetningar")

    container = tk.Frame(frame, bg=BG_DARK)
    container.pack(fill=tk.X, padx=28)

    input_label = create_styled_label(container, "Límdu inn dagsetningar (eina í hverja línu):", size=9, color=TEXT_SECONDARY)
    input_label.pack(anchor='w')

    input_shell, input_text = create_styled_text_area(container, "Límdu dagsetningar á borð við 31.12.2025, 31/12/25 eða 31 desember 2025", height=6)
    input_shell.pack(fill=tk.X, pady=(5, 12))
    input_resize = attach_auto_resize_text(input_text, min_lines=6, max_lines=8)

    output_label = create_styled_label(container, "Sniðin niðurstaða (DD.MM.YYYY):", size=9, color=TEXT_SECONDARY)
    output_label.pack(anchor='w')

    output_shell, output_text = create_styled_text_area(container, "Sniðnar dagsetningar birtast hér", height=6)
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
            status_label.config(text=f"Fjöldi ógildra lína sem var sleppt: {bad}.", fg=TEXT_WARNING)
        else:
            status_label.config(text="Dagsetningarnar voru sniðnar.", fg=TEXT_SUCCESS)

    def copy_output():
        text = get_text_area_value(output_text).strip()
        if text:
            pyperclip.copy(text)
            status_label.config(text="Afritað á klippiborð (ein færsla í hverri línu).", fg=TEXT_SUCCESS)
        else:
            status_label.config(text="Engin niðurstaða er til að afrita.", fg=TEXT_WARNING)

    btn_row = tk.Frame(container, bg=BG_DARK)
    btn_row.pack(pady=5)

    format_btn = create_styled_button(btn_row, "Sníða", format_dates, width=12)
    format_btn.pack(side=tk.LEFT, padx=5)

    copy_btn = create_styled_button(btn_row, "Afrita niðurstöðu", copy_output, width=16, accent=False)
    copy_btn.pack(side=tk.LEFT, padx=5)

    back_button = create_styled_button(frame, "← Til baka í valmynd", initialize_main_menu, width=20)
    back_button.pack(pady=10)
    ensure_window_fits()

def _format_kennitala_value(value, mode):
    digits = "".join(ch for ch in str(value) if ch.isdigit())
    if not digits:
        return None
    if mode in {"remove_dash", "add_dash"}:
        if len(digits) < 10:
            return None
        base = digits[:10]
        return f"{base[:6]}-{base[6:]}" if mode == "add_dash" else base
    if mode == "remove_zero":
        return digits[:-1] if digits.endswith("0") else digits
    if mode == "add_zero":
        return f"{digits}0"
    return None


def initialize_format_ids():
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title("Sníða kennitölur")

    container = tk.Frame(frame, bg=BG_DARK)
    container.pack(fill=tk.X, padx=28)

    input_label = create_styled_label(container, "Límdu inn kennitölur (eina í hverja línu):", size=9, color=TEXT_SECONDARY)
    input_label.pack(anchor='w')

    input_shell, input_text = create_styled_text_area(container, "Límdu kennitölurnar hér, eina í hverja línu", height=6)
    input_shell.pack(fill=tk.X, pady=(5, 12))
    input_resize = attach_auto_resize_text(input_text, min_lines=6, max_lines=8)

    toggle_frame = tk.Frame(container, bg=BG_DARK)
    toggle_frame.pack(pady=(0, 8))

    mode_var = tk.StringVar(value="remove_dash")
    mode_buttons = {}

    def set_id_mode(mode):
        mode_var.set(mode)
        for value, button in mode_buttons.items():
            set_button_accent(button, value == mode)

    mode_buttons["remove_dash"] = create_styled_button(
        toggle_frame,
        "Fjarlægja bandstrik",
        lambda: set_id_mode("remove_dash"),
        width=14,
        accent=True,
    )
    mode_buttons["remove_dash"].pack(side=tk.LEFT, padx=(0, 8))

    mode_buttons["add_dash"] = create_styled_button(
        toggle_frame,
        "Bæta við bandstriki",
        lambda: set_id_mode("add_dash"),
        width=14,
        accent=False,
    )
    mode_buttons["add_dash"].pack(side=tk.LEFT, padx=(8, 0))

    zero_toggle_frame = tk.Frame(container, bg=BG_DARK)
    zero_toggle_frame.pack(pady=(0, 12))

    mode_buttons["remove_zero"] = create_styled_button(
        zero_toggle_frame,
        "Fjarlægja núll í lokin",
        lambda: set_id_mode("remove_zero"),
        width=14,
        accent=False,
    )
    mode_buttons["remove_zero"].pack(side=tk.LEFT, padx=(0, 8))

    mode_buttons["add_zero"] = create_styled_button(
        zero_toggle_frame,
        "Bæta núlli við í lokin",
        lambda: set_id_mode("add_zero"),
        width=14,
        accent=False,
    )
    mode_buttons["add_zero"].pack(side=tk.LEFT, padx=(8, 0))

    output_label = create_styled_label(container, "Sniðin niðurstaða:", size=9, color=TEXT_SECONDARY)
    output_label.pack(anchor='w')

    output_shell, output_text = create_styled_text_area(container, "Sniðnar kennitölur birtast hér", height=6)
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
            result = _format_kennitala_value(line, mode_var.get())
            if result is None:
                bad += 1
                continue
            formatted.append(result)

        set_text_area_value(output_text, "\n".join(formatted))
        output_resize()
        schedule_layout_refresh()
        if bad:
            status_label.config(text=f"Fjöldi ógildra lína sem var sleppt: {bad}.", fg=TEXT_WARNING)
        else:
            status_label.config(text="Kennitölurnar voru sniðnar.", fg=TEXT_SUCCESS)

    def copy_output():
        text = get_text_area_value(output_text).strip()
        if text:
            pyperclip.copy(text)
            status_label.config(text="Afritað á klippiborð (ein færsla í hverri línu).", fg=TEXT_SUCCESS)
        else:
            status_label.config(text="Engin niðurstaða er til að afrita.", fg=TEXT_WARNING)

    btn_row = tk.Frame(container, bg=BG_DARK)
    btn_row.pack(pady=5)

    format_btn = create_styled_button(btn_row, "Sníða", format_ids, width=12)
    format_btn.pack(side=tk.LEFT, padx=5)

    copy_btn = create_styled_button(btn_row, "Afrita niðurstöðu", copy_output, width=16, accent=False)
    copy_btn.pack(side=tk.LEFT, padx=5)

    back_button = create_styled_button(frame, "← Til baka í valmynd", initialize_main_menu, width=20)
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
        f"Keyrsla hefur verið í bið í meira en {AUTOMATION_WATCHDOG_SECONDS} sekúndur. "
        "Athugaðu Tok og haltu síðan áfram frá núverandi eða næstu línu."
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
        _auto_code_training_dir=_auto_code_training_dir,
        _auto_code_key_dir=_auto_code_key_dir,
        _auto_code_memory_path=_auto_code_memory_path,
        _training_data_dir=_training_data_dir,
        _auto_code_mode=_auto_code_mode,
        _ml_artifact_dir=_ml_artifact_dir,
        RECENT_VISIBLE_LIMIT=RECENT_VISIBLE_LIMIT,
    )

    coded_file_tools.configure(
        root=root,
        frame=frame,
        set_page_title=set_page_title,
        initialize_main_menu=initialize_main_menu,
        open_file_location=open_file_location,
        play_success_sound=play_success_sound,
        play_error_sound=play_error_sound,
        ensure_window_fits=ensure_window_fits,
        schedule_layout_refresh=schedule_layout_refresh,
        _short_file_name=_short_file_name,
        _training_data_dir=_training_data_dir,
        compress_transactions_interactive=bank_formatter.compress_transactions_interactive,
        manual_compress_transactions_interactive=bank_formatter.manual_compress_transactions_interactive,
    )

    shareholder_registry.configure(
        root=root,
        frame=frame,
        set_page_title=set_page_title,
        initialize_main_menu=initialize_main_menu,
        resource_path=resource_path,
        _training_data_dir=_training_data_dir,
        open_file_location=open_file_location,
        window_work_area=_window_work_area,
    )

# =============================================================================
# MAIN APPLICATION SETUP
# =============================================================================

def main():
    global root, settings
    global style, header_frame, logo_image, logo_image_small, logo_label
    global title_frame, page_title_label, frame

    if TKDND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    settings = load_settings()
    apply_theme(settings.get("theme", "light"))
    _sync_runtime_timing_from_settings()
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

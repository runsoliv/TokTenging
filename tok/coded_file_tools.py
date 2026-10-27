"""Standalone tools for compressing and combining coded bank workbooks."""

from __future__ import annotations

import os
from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog, ttk
import uuid

import pandas as pd

try:
    from . import ui_components
    from .bank_utils import autofit_excel_columns
    from .training_archive import refresh_paired_rows_from_coded
    from .ui_components import (
        attach_drop_target,
        attach_drop_target_many,
        create_drop_box,
        create_panel,
        create_styled_button,
        create_styled_label,
        set_button_accent,
    )
except ImportError:
    import ui_components
    from bank_utils import autofit_excel_columns
    from training_archive import refresh_paired_rows_from_coded
    from ui_components import (
        attach_drop_target,
        attach_drop_target_many,
        create_drop_box,
        create_panel,
        create_styled_button,
        create_styled_label,
        set_button_accent,
    )


globals().update(ui_components.get_theme("light"))

RECENT_CODED_FILE_LIMIT = 6
REQUIRED_CODED_COLUMNS = ("DATE", "TEXT", "DEBIT", "AMOUNT", "CREDIT")
CODED_FILE_TYPES = (("Kóðaðar Excel-skrár", "*.xlsx"), ("Allar skrár", "*.*"))
UNNAMED_EXCEL_COLUMN = re.compile(r"^Unnamed:\s*\d+(?:_level_\d+)?$", re.IGNORECASE)

root = None
frame = None


def _not_configured(*_args, **_kwargs):
    raise RuntimeError("coded_file_tools.configure() must be called before opening a tool page")


set_page_title = _not_configured
initialize_main_menu = _not_configured
open_file_location = _not_configured
ensure_window_fits = _not_configured
schedule_layout_refresh = lambda: None
_short_file_name = lambda path, limit=22: os.path.basename(str(path))[:limit]
play_success_sound = lambda: None
play_error_sound = lambda: None
compress_transactions_interactive = _not_configured
manual_compress_transactions_interactive = _not_configured
_training_data_dir = _not_configured


class CodedWorkbookError(ValueError):
    """Raised when selected workbooks cannot be safely combined."""


def configure(**dependencies):
    globals().update(dependencies)


def set_theme(theme):
    globals().update(theme)


def _configure_scrollbar_styles():
    """Use compact, theme-coloured scrollbars instead of native white chrome."""
    style = ttk.Style(root)
    style.layout(
        "Coded.Vertical.TScrollbar",
        [("Vertical.Scrollbar.trough", {"sticky": "ns", "children": [
            ("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"}),
        ]})],
    )
    style.layout(
        "Coded.Horizontal.TScrollbar",
        [("Horizontal.Scrollbar.trough", {"sticky": "ew", "children": [
            ("Horizontal.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"}),
        ]})],
    )
    for name in ("Coded.Vertical.TScrollbar", "Coded.Horizontal.TScrollbar"):
        style.configure(
            name,
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
            name,
            background=[("pressed", ACCENT), ("active", BUTTON_MUTED_HOVER)],
            lightcolor=[("pressed", ACCENT), ("active", BUTTON_MUTED_HOVER)],
            darkcolor=[("pressed", ACCENT), ("active", BUTTON_MUTED_HOVER)],
        )


def _normalized_unique_paths(input_paths):
    normalized = []
    seen = set()
    for value in input_paths or []:
        if not value:
            continue
        path = os.path.normpath(os.path.abspath(os.fspath(value)))
        key = os.path.normcase(path)
        if key in seen:
            continue
        seen.add(key)
        normalized.append(path)
    return normalized


def find_recent_coded_files(training_data_dir, limit=RECENT_CODED_FILE_LIMIT):
    """Return the newest Bank Formatter coded outputs from the imports folder."""
    if not training_data_dir:
        return []
    root = Path(training_data_dir)
    imports_root = root / "imports"
    search_root = imports_root if imports_root.is_dir() else root
    if not search_root.is_dir():
        return []

    candidates = []
    for path in search_root.rglob("*.xlsx"):
        stem = path.stem.lower()
        if "_compressed_txs" in stem:
            continue
        if not re.search(r"_coded\d*$", stem):
            continue
        try:
            modified = path.stat().st_mtime_ns
        except OSError:
            continue
        candidates.append((modified, os.path.normcase(str(path)), str(path)))

    candidates.sort(reverse=True)
    return [path for _modified, _sort_path, path in candidates[:max(0, int(limit))]]


def read_coded_workbook(file_path):
    """Read and validate a Bank Formatter coded workbook."""
    path = Path(file_path)
    if not path.is_file():
        raise CodedWorkbookError(f"Skrá fannst ekki: {path.name or path}")
    if path.suffix.lower() != ".xlsx":
        raise CodedWorkbookError(f"{path.name} er ekki kóðuð .xlsx-vinnubók.")
    try:
        data = pd.read_excel(path, dtype=object, keep_default_na=False)
    except Exception as exc:
        raise CodedWorkbookError(f"Ekki tókst að lesa {path.name}: {str(exc)[:90]}") from exc

    # Excel creates ``Unnamed: N`` headers for side notes entered beyond the
    # coded table. They are not part of the Bank Formatter schema and must not
    # prevent otherwise compatible workbooks from being combined. This only
    # normalizes the in-memory frame; the source workbook is never changed.
    unnamed_columns = [
        column
        for column in data.columns
        if UNNAMED_EXCEL_COLUMN.fullmatch(str(column).strip())
    ]
    if unnamed_columns:
        data = data.drop(columns=unnamed_columns)

    missing = [column for column in REQUIRED_CODED_COLUMNS if column not in data.columns]
    if missing:
        raise CodedWorkbookError(
            f"{path.name} vantar dálka kóðaðrar skrár: {', '.join(missing)}"
        )
    return data


def combined_output_path(first_file_path):
    """Build the combined workbook name beside the first selected file."""
    first = Path(os.path.normpath(os.path.abspath(os.fspath(first_file_path))))
    stem = first.stem
    if not stem.lower().endswith("_combined"):
        stem += "_combined"
    return str(first.with_name(f"{stem}.xlsx"))


def combine_coded_workbooks(input_paths, output_path=None):
    """Stack coded workbooks into a ``*_combined.xlsx`` file by default.

    Every workbook must have the same columns in the same order. The default
    destination is beside the first input and leaves all source files unchanged.
    """
    paths = _normalized_unique_paths(input_paths)
    if len(paths) < 2:
        raise CodedWorkbookError("Veldu að minnsta kosti tvær mismunandi kóðaðar skrár.")

    frames = []
    expected_columns = None
    row_counts = []
    for path in paths:
        data = read_coded_workbook(path)
        columns = list(data.columns)
        if expected_columns is None:
            expected_columns = columns
        elif columns != expected_columns:
            expected_text = ", ".join(str(column) for column in expected_columns)
            actual_text = ", ".join(str(column) for column in columns)
            raise CodedWorkbookError(
                f"{os.path.basename(path)} er með aðra dálka. "
                f"Búist var við [{expected_text}], en fannst [{actual_text}]."
            )
        frames.append(data)
        row_counts.append(len(data))

    destination = os.path.normpath(os.path.abspath(output_path or combined_output_path(paths[0])))
    if Path(destination).suffix.lower() != ".xlsx":
        raise CodedWorkbookError("Sameinaða úttaksskráin verður að vera .xlsx-skrá.")
    destination_dir = os.path.dirname(destination)
    if not os.path.isdir(destination_dir):
        raise CodedWorkbookError(f"Úttaksmappa fannst ekki: {destination_dir}")

    combined = pd.concat(frames, ignore_index=True)
    temporary = os.path.join(
        destination_dir,
        f".{Path(destination).stem}.combine-{uuid.uuid4().hex}.xlsx",
    )
    try:
        combined.to_excel(temporary, index=False)
        autofit_excel_columns(temporary)
        os.replace(temporary, destination)
    except Exception as exc:
        try:
            if os.path.exists(temporary):
                os.remove(temporary)
        except Exception:
            pass
        raise CodedWorkbookError(
            f"Ekki tókst að vista {os.path.basename(destination)}: {str(exc)[:90]}"
        ) from exc

    return {
        "output_path": destination,
        "file_count": len(paths),
        "row_count": len(combined),
        "row_counts": row_counts,
        "columns": list(expected_columns or []),
    }


def _clear_page(title):
    for widget in frame.winfo_children():
        widget.destroy()
    set_page_title(title)


def _recent_file_buttons(parent, selected_paths, toggle_path):
    recent = find_recent_coded_files(_training_data_dir())
    if not recent:
        create_styled_label(
            parent,
            "Engar nýlegar kóðaðar skrár.",
            size=8,
            color=TEXT_SECONDARY,
            bg=BG_CARD,
        ).pack(anchor="w")
        return

    grid = tk.Frame(parent, bg=BG_CARD)
    grid.pack(fill=tk.X)
    selected_keys = {os.path.normcase(os.path.abspath(path)) for path in selected_paths}
    for index, path in enumerate(recent):
        selected = os.path.normcase(os.path.abspath(path)) in selected_keys
        button = create_styled_button(
            grid,
            _short_file_name(path, limit=27),
            lambda value=path: toggle_path(value),
            width=18,
            accent=selected,
            height=36,
        )
        button.grid(
            row=index // 2,
            column=index % 2,
            sticky="ew",
            padx=(0, 5) if index % 2 == 0 else (5, 0),
            pady=4,
        )
        set_button_accent(button, selected)
    grid.columnconfigure(0, weight=1)
    grid.columnconfigure(1, weight=1)


def _sync_training_archive_before_compression(file_path):
    folder = Path(file_path).parent
    if not (folder / "manifest.json").exists() or not (folder / "paired_rows.csv").exists():
        return {"ok": True, "message": ""}
    return refresh_paired_rows_from_coded(
        session_dir=folder,
        coded_file_path=file_path,
    )


def initialize_compress_txs():
    """Show the standalone coded-workbook transaction compression page."""
    _clear_page("Þjappa færslum")
    selected = {"path": ""}

    card, inner = create_panel(frame, padx=22, pady=18)
    card.pack(fill=tk.X, padx=28, pady=(8, 12))

    create_styled_label(
        inner,
        "Kóðuð bankaskrá",
        size=10,
        bold=True,
        bg=BG_CARD,
    ).pack(anchor="w")
    input_drop = create_drop_box(inner, "Dragðu kóðaða Excel-skrá hingað eða smelltu til að velja", height=5)
    input_drop.pack(fill=tk.X, pady=(8, 8))

    status_label = create_styled_label(
        inner,
        "Veldu kóðaða .xlsx-skrá.",
        size=9,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    )
    status_label.configure(anchor="w", justify="left", wraplength=520)

    recent_frame = tk.Frame(inner, bg=BG_CARD)

    def set_status(message, color=TEXT_SECONDARY):
        status_label.config(text=message, fg=color)
        schedule_layout_refresh()

    def refresh_recent():
        for widget in recent_frame.winfo_children():
            widget.destroy()
        _recent_file_buttons(
            recent_frame,
            [selected["path"]] if selected["path"] else [],
            select_file,
        )

    def select_file(path):
        try:
            data = read_coded_workbook(path)
        except CodedWorkbookError as exc:
            set_status(str(exc), TEXT_ERROR)
            play_error_sound()
            return
        path = os.path.normpath(os.path.abspath(path))
        selected["path"] = path
        input_drop.config(text=os.path.basename(path), fg=TEXT_PRIMARY)
        transaction_count = len(data)
        transaction_text = "1 færsla" if transaction_count == 1 else f"{transaction_count} færslur"
        set_status(f"Tilbúið: {transaction_text}.", TEXT_SUCCESS)
        refresh_recent()

    def browse():
        filename = filedialog.askopenfilename(
            initialdir=os.path.dirname(selected["path"]) if selected["path"] else os.getcwd(),
            title="Veldu kóðaða bankaskrá",
            filetypes=CODED_FILE_TYPES,
        )
        if filename:
            select_file(filename)

    def run_compression(automatic=False):
        path = selected["path"]
        if not path:
            set_status("Veldu fyrst kóðaða skrá.", TEXT_WARNING)
            return
        # Re-read on every press so edits saved after file selection are used.
        try:
            current_data = read_coded_workbook(path)
        except CodedWorkbookError as exc:
            set_status(str(exc), TEXT_ERROR)
            play_error_sound()
            return
        input_drop.config(text=os.path.basename(path), fg=TEXT_PRIMARY)
        if automatic:
            set_status(
                f"Færslur endurhlaðnar: {len(current_data)}. Leita að verslunarflokkum...",
                TEXT_SECONDARY,
            )
        else:
            set_status(f"Færslur endurhlaðnar: {len(current_data)}. Opna handvirkt val...", TEXT_SECONDARY)
        sync_result = _sync_training_archive_before_compression(path)
        if not sync_result.get("ok"):
            set_status(sync_result.get("message", "Samstilling þjálfunarsafns mistókst."), TEXT_ERROR)
            play_error_sound()
            return
        if automatic:
            result = compress_transactions_interactive(path)
        else:
            result = manual_compress_transactions_interactive(path)
        if result.get("ok"):
            set_status(result.get("message", "Færslum var þjappað."), TEXT_SUCCESS)
            play_success_sound()
        else:
            set_status(result.get("message", "Engum færslum var þjappað."), TEXT_WARNING)

    if not attach_drop_target(input_drop, select_file):
        input_drop.config(text="Smelltu til að velja kóðaða Excel-skrá", fg=TEXT_SECONDARY)
    input_drop.bind("<Button-1>", lambda _event: browse())

    create_styled_label(
        inner,
        "Nýlegar skrár",
        size=9,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    ).pack(anchor="w", pady=(0, 2))
    recent_frame.pack(fill=tk.X, pady=(0, 8))
    refresh_recent()

    status_label.pack(fill=tk.X, pady=(0, 10))

    compression_help = create_styled_label(
        inner,
        (
            "Sjálfvirk þjöppun finnur flokka með 5 eða fleiri færslum frá sama "
            "mótaðila eða sömu verslun og sýnir hvern flokk áður en þú samþykkir eða sleppir honum."
        ),
        size=8,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    )
    compression_help.configure(anchor="w", justify="left", wraplength=520)
    compression_help.pack(fill=tk.X, pady=(0, 8))

    compression_buttons = tk.Frame(inner, bg=BG_CARD)
    compression_buttons.pack(fill=tk.X)
    automatic_button = create_styled_button(
        compression_buttons,
        "Þjappa sjálfvirkt",
        lambda: run_compression(automatic=True),
        width=20,
        height=42,
    )
    automatic_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    manual_button = create_styled_button(
        compression_buttons,
        "Þjappa handvirkt",
        lambda: run_compression(automatic=False),
        width=20,
        height=42,
        accent=False,
    )
    manual_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

    nav = tk.Frame(frame, bg=BG_DARK)
    nav.pack(fill=tk.X, padx=28, pady=(0, 10))
    back_button = create_styled_button(
        nav,
        "← Til baka í valmynd",
        initialize_main_menu,
        width=18,
        accent=False,
        height=42,
    )
    back_button.pack(fill=tk.X)
    ensure_window_fits()


def initialize_combine_files():
    """Show the multi-select coded-workbook combine page."""
    _clear_page("Sameina skrár")
    selected_paths = []
    last_output = {"path": ""}

    card, inner = create_panel(frame, padx=22, pady=18)
    card.pack(fill=tk.X, padx=28, pady=(8, 12))

    create_styled_label(
        inner,
        "Kóðaðar bankaskrár",
        size=10,
        bold=True,
        bg=BG_CARD,
    ).pack(anchor="w")
    combine_help_label = create_styled_label(
        inner,
        "Skrárnar eru settar saman í valinni röð og vistaðar við hlið fyrstu skrárinnar með viðskeytinu _combined.",
        size=8,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    )
    combine_help_label.configure(anchor="w", justify="left")
    combine_help_label.pack(fill=tk.X, pady=(2, 8))

    input_drop = create_drop_box(
        inner,
        "Dragðu kóðaðar Excel-skrár hingað eða smelltu til að velja",
        height=3,
    )
    input_drop.pack(fill=tk.X, pady=(0, 8))
    drop_state = {"available": True}

    list_shell = tk.Frame(inner, bg=BORDER, padx=1, pady=1)
    list_shell.pack(fill=tk.X, pady=(0, 8))
    list_shell.columnconfigure(0, weight=1)
    list_shell.rowconfigure(0, weight=1)
    _configure_scrollbar_styles()
    listbox = tk.Listbox(
        list_shell,
        height=4,
        bg=BG_INPUT,
        fg=TEXT_PRIMARY,
        selectbackground=ACCENT,
        selectforeground=TEXT_ON_ACCENT,
        font=("Segoe UI", 9),
        relief="flat",
        bd=0,
        highlightthickness=0,
        activestyle="none",
    )
    scrollbar = ttk.Scrollbar(
        list_shell,
        orient=tk.VERTICAL,
        command=listbox.yview,
        style="Coded.Vertical.TScrollbar",
    )
    horizontal_scrollbar = ttk.Scrollbar(
        list_shell,
        orient=tk.HORIZONTAL,
        command=listbox.xview,
        style="Coded.Horizontal.TScrollbar",
    )
    listbox.configure(yscrollcommand=scrollbar.set, xscrollcommand=horizontal_scrollbar.set)
    listbox.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=(7, 0))
    scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 7), pady=(7, 0))
    horizontal_scrollbar.grid(row=1, column=0, sticky="ew", padx=(8, 0), pady=(0, 7))

    output_label = create_styled_label(
        inner,
        "Úttak: veldu að minnsta kosti tvær skrár",
        size=8,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    )
    output_label.configure(anchor="w", justify="left")
    output_label.pack(fill=tk.X, pady=(0, 6))

    status_label = create_styled_label(
        inner,
        "Veldu eins margar kóðaðar skrár og þörf er á.",
        size=9,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    )
    status_label.configure(anchor="w", justify="left")

    def resize_combine_copy(event):
        wrap = max(220, int(event.width) - 8)
        for label in (combine_help_label, output_label, status_label):
            label.configure(wraplength=wrap)

    inner.bind("<Configure>", resize_combine_copy, add="+")

    recent_frame = tk.Frame(inner, bg=BG_CARD)

    def refresh_combine_layout():
        """Resize the rounded card and window after dynamic rows settle."""
        def apply_layout():
            try:
                if not card.winfo_exists():
                    return
                inner.update_idletasks()
                card._sync_height()
                card._draw()
                card.update_idletasks()
                ensure_window_fits()
            except Exception:
                pass

        apply_layout()
        try:
            root.after_idle(apply_layout)
            root.after(60, apply_layout)
            root.after(160, apply_layout)
        except Exception:
            pass

    def set_status(message, color=TEXT_SECONDARY):
        status_label.config(text=message, fg=color)
        schedule_layout_refresh()
        refresh_combine_layout()

    def refresh_recent():
        for widget in recent_frame.winfo_children():
            widget.destroy()
        _recent_file_buttons(recent_frame, selected_paths, toggle_recent)

    def refresh_list(selection_index=None):
        listbox.delete(0, tk.END)
        for index, path in enumerate(selected_paths, start=1):
            marker = "  (heiti úttaksskrár)" if index == 1 else ""
            listbox.insert(tk.END, f"{index}. {os.path.basename(path)}{marker}")
        if selected_paths:
            input_drop.config(
                text=(
                    "1 kóðuð skrá valin — smelltu til að bæta við"
                    if len(selected_paths) == 1
                    else f"{len(selected_paths)} kóðaðar skrár valdar — smelltu til að bæta við"
                ),
                fg=TEXT_PRIMARY,
            )
            output_label.config(
                text=f"Úttak: {os.path.basename(combined_output_path(selected_paths[0]))}",
                fg=TEXT_PRIMARY,
            )
        else:
            input_drop.config(
                text=(
                    "Dragðu kóðaðar Excel-skrár hingað eða smelltu til að velja"
                    if drop_state["available"]
                    else "Smelltu til að velja kóðaðar Excel-skrár"
                ),
                fg=TEXT_SECONDARY,
            )
            output_label.config(text="Úttak: veldu að minnsta kosti tvær skrár", fg=TEXT_SECONDARY)
        if selection_index is not None and selected_paths:
            selection_index = max(0, min(selection_index, len(selected_paths) - 1))
            listbox.selection_set(selection_index)
            listbox.see(selection_index)
        refresh_recent()
        schedule_layout_refresh()
        refresh_combine_layout()

    def add_paths(paths):
        added = 0
        known = {os.path.normcase(path) for path in selected_paths}
        invalid = []
        for raw_path in paths:
            path = os.path.normpath(os.path.abspath(raw_path))
            key = os.path.normcase(path)
            if key in known:
                continue
            try:
                read_coded_workbook(path)
            except CodedWorkbookError as exc:
                invalid.append(str(exc))
                continue
            selected_paths.append(path)
            known.add(key)
            added += 1
        if added:
            set_status(
                "1 skrá valin." if len(selected_paths) == 1 else f"{len(selected_paths)} skrár valdar.",
                TEXT_SUCCESS,
            )
        elif invalid:
            set_status(invalid[0], TEXT_ERROR)
            play_error_sound()
        refresh_list(len(selected_paths) - 1 if selected_paths else None)

    def browse_many():
        filenames = filedialog.askopenfilenames(
            initialdir=os.path.dirname(selected_paths[0]) if selected_paths else os.getcwd(),
            title="Veldu kóðaðar bankaskrár",
            filetypes=CODED_FILE_TYPES,
        )
        if filenames:
            add_paths(filenames)

    def toggle_recent(path):
        normalized = os.path.normpath(os.path.abspath(path))
        key = os.path.normcase(normalized)
        for index, selected_path in enumerate(selected_paths):
            if os.path.normcase(selected_path) == key:
                selected_paths.pop(index)
                set_status(
                    "1 skrá valin." if len(selected_paths) == 1 else f"{len(selected_paths)} skrár valdar.",
                    TEXT_SECONDARY,
                )
                refresh_list()
                return
        add_paths([normalized])

    def selected_index():
        selection = listbox.curselection()
        return int(selection[0]) if selection else None

    def remove_selected():
        index = selected_index()
        if index is None:
            set_status("Veldu skrá á listanum til að fjarlægja hana.", TEXT_WARNING)
            return
        selected_paths.pop(index)
        set_status(
            "1 skrá valin." if len(selected_paths) == 1 else f"{len(selected_paths)} skrár valdar.",
            TEXT_SECONDARY,
        )
        refresh_list(index)

    def move_selected(direction):
        index = selected_index()
        if index is None:
            set_status("Veldu skrá á listanum til að færa hana.", TEXT_WARNING)
            return
        target = index + direction
        if target < 0 or target >= len(selected_paths):
            return
        selected_paths[index], selected_paths[target] = selected_paths[target], selected_paths[index]
        refresh_list(target)

    def clear_selected():
        selected_paths.clear()
        set_status("Valið var hreinsað.", TEXT_SECONDARY)
        refresh_list()

    def combine_selected():
        if len(selected_paths) < 2:
            set_status("Veldu að minnsta kosti tvær mismunandi kóðaðar skrár.", TEXT_WARNING)
            return
        set_status("Sameina skrár...", TEXT_SECONDARY)
        try:
            result = combine_coded_workbooks(selected_paths)
        except CodedWorkbookError as exc:
            set_status(str(exc), TEXT_ERROR)
            play_error_sound()
            return
        last_output["path"] = result["output_path"]
        set_status(
            f"Skrár sameinaðar: {result['file_count']}. Færslur: {result['row_count']}. "
            f"Úttak: {os.path.basename(result['output_path'])}.",
            TEXT_SUCCESS,
        )
        play_success_sound()
        # Continue from the new combined workbook without accidentally adding
        # all of the same source rows again on a second press.
        selected_paths[:] = [result["output_path"]]
        refresh_list(0)
        open_row.pack(fill=tk.X, pady=(8, 0))
        schedule_layout_refresh()
        refresh_combine_layout()

    drop_state["available"] = attach_drop_target_many(input_drop, add_paths)
    input_drop.bind("<Button-1>", lambda _event: browse_many())

    list_actions = tk.Frame(inner, bg=BG_CARD)
    list_actions.pack(fill=tk.X, pady=(0, 10))
    for index, (label, command) in enumerate((
        ("Færa upp", lambda: move_selected(-1)),
        ("Færa niður", lambda: move_selected(1)),
        ("Fjarlægja", remove_selected),
        ("Hreinsa", clear_selected),
    )):
        button = create_styled_button(list_actions, label, command, width=9, accent=False, height=34)
        button.pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True,
            padx=(0, 4) if index == 0 else ((4, 0) if index == 3 else 4),
        )

    create_styled_label(
        inner,
        "Nýlegar skrár (smelltu til að bæta við eða fjarlægja)",
        size=9,
        color=TEXT_SECONDARY,
        bg=BG_CARD,
    ).pack(anchor="w", pady=(0, 2))
    recent_frame.pack(fill=tk.X, pady=(0, 8))
    refresh_recent()

    status_label.pack(fill=tk.X, pady=(0, 8))
    combine_button = create_styled_button(inner, "Sameina skrár", combine_selected, width=20, height=42)
    combine_button.pack(fill=tk.X)

    open_row = tk.Frame(inner, bg=BG_CARD)
    open_button = create_styled_button(
        open_row,
        "Opna skrá",
        lambda: os.startfile(last_output["path"]) if last_output["path"] else None,
        width=18,
        accent=False,
        height=38,
    )
    open_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
    location_button = create_styled_button(
        open_row,
        "Opna staðsetningu skrár",
        lambda: open_file_location(last_output["path"]) if last_output["path"] else None,
        width=18,
        accent=False,
        height=38,
    )
    location_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))

    nav = tk.Frame(frame, bg=BG_DARK)
    nav.pack(fill=tk.X, padx=28, pady=(0, 10))
    back_button = create_styled_button(
        nav,
        "← Til baka í valmynd",
        initialize_main_menu,
        width=18,
        accent=False,
        height=42,
    )
    back_button.pack(fill=tk.X)
    refresh_list()
    refresh_combine_layout()

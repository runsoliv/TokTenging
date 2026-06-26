"""Reusable Tkinter widgets and theme primitives for Tok Tenging."""

import tkinter as tk

TKDND_AVAILABLE = False
DND_FILES = None
_layout_refresh_callback = lambda: None

BG_DARK = '#d9e8eb'        # App background
BG_CARD = '#f8fbfa'        # Panel background
BG_INPUT = '#f7fbfc'       # Input/drop area background
BORDER = '#9fb8c1'         # Subtle borders
ACCENT = '#1b87ea'         # Primary action
ACCENT_HOVER = '#1575cc'   # Primary hover
ACCENT_SOFT = '#d9ecff'    # Selected/soft accent
TEXT_PRIMARY = '#102027'   # Main text
TEXT_SECONDARY = '#4f6570' # Muted text
TEXT_ON_ACCENT = '#ffffff'
TEXT_SUCCESS = '#0f766e'   # Green for success
TEXT_ERROR = '#dc2626'     # Red for errors
TEXT_WARNING = '#b45309'   # Amber for warnings
BUTTON_MUTED = '#edf1f5'
BUTTON_MUTED_HOVER = '#e1e8ef'

THEMES = {
    "light": {
        "BG_DARK": '#d9e8eb',
        "BG_CARD": '#f8fbfa',
        "BG_INPUT": '#f7fbfc',
        "BORDER": '#9fb8c1',
        "ACCENT": '#1b87ea',
        "ACCENT_HOVER": '#1575cc',
        "ACCENT_SOFT": '#d9ecff',
        "TEXT_PRIMARY": '#102027',
        "TEXT_SECONDARY": '#4f6570',
        "TEXT_ON_ACCENT": '#ffffff',
        "TEXT_SUCCESS": '#0f766e',
        "TEXT_ERROR": '#dc2626',
        "TEXT_WARNING": '#b45309',
        "BUTTON_MUTED": '#edf1f5',
        "BUTTON_MUTED_HOVER": '#e1e8ef',
    },
    "dark": {
        "BG_DARK": '#171821',
        "BG_CARD": '#2f3140',
        "BG_INPUT": '#494c5d',
        "BORDER": '#3e4255',
        "ACCENT": '#1b87ea',
        "ACCENT_HOVER": '#3097f4',
        "ACCENT_SOFT": '#233c5d',
        "TEXT_PRIMARY": '#f5f6fb',
        "TEXT_SECONDARY": '#bbc1cf',
        "TEXT_ON_ACCENT": '#ffffff',
        "TEXT_SUCCESS": '#5ee7a8',
        "TEXT_ERROR": '#fb7185',
        "TEXT_WARNING": '#fbbf24',
        "BUTTON_MUTED": '#494c5d',
        "BUTTON_MUTED_HOVER": '#5b6078',
    },
}



def get_theme(theme_name):
    return THEMES.get(theme_name, THEMES["light"])


def set_theme(theme_name):
    theme = get_theme(theme_name)
    globals().update(theme)
    return theme


def configure_drag_and_drop(available, dnd_files):
    global TKDND_AVAILABLE, DND_FILES
    TKDND_AVAILABLE = bool(available)
    DND_FILES = dnd_files


def set_layout_refresh_callback(callback):
    global _layout_refresh_callback
    _layout_refresh_callback = callback or (lambda: None)


def _parse_drop_data(data: str) -> str:
    """Extract first file path from tkdnd drop data."""
    if not data:
        return ""
    data = data.strip()
    if not data:
        return ""
    # tkdnd uses a Tcl list format: items may be wrapped in braces if they contain spaces.
    items = []
    buf = ""
    in_brace = False
    for ch in data:
        if ch == "{" and not in_brace:
            in_brace = True
            buf = ""
            continue
        if ch == "}" and in_brace:
            in_brace = False
            items.append(buf)
            buf = ""
            continue
        if ch == " " and not in_brace:
            if buf:
                items.append(buf)
                buf = ""
            continue
        buf += ch
    if buf:
        items.append(buf)
    return items[0] if items else ""

def attach_drop_target(widget, on_path):
    """
    Try to make a widget accept file drops. Returns False if tkdnd is unavailable.
    """
    if not TKDND_AVAILABLE:
        return False
    try:
        try:
            widget.drop_target_register(DND_FILES)
        except Exception:
            widget.drop_target_register('DND_Files')
        widget.dnd_bind('<<Drop>>', lambda e: on_path(_parse_drop_data(e.data)))
        return True
    except Exception:
        return False

def _parent_bg(parent):
    try:
        return parent.cget("bg")
    except Exception:
        return BG_DARK

def _rounded_rect(canvas, x1, y1, x2, y2, radius, **kwargs):
    points = [
        x1 + radius, y1, x2 - radius, y1,
        x2, y1, x2, y1 + radius,
        x2, y2 - radius, x2, y2,
        x2 - radius, y2, x1 + radius, y2,
        x1, y2, x1, y2 - radius,
        x1, y1 + radius, x1, y1,
    ]
    return canvas.create_polygon(points, smooth=True, **kwargs)

class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, command, width=20, accent=True, height=46):
        self.command = command
        self.text = text
        self.accent = accent
        self.button_height = int(height)
        self.normal_bg = ACCENT if accent else BUTTON_MUTED
        self.hover_bg = ACCENT_HOVER if accent else BUTTON_MUTED_HOVER
        self.text_color = TEXT_ON_ACCENT if accent else TEXT_PRIMARY
        self.border_color = '' if accent else BORDER
        pixel_width = max(110, int(width) * 9)
        super().__init__(
            parent,
            width=pixel_width,
            height=self.button_height,
            bg=_parent_bg(parent),
            highlightthickness=0,
            bd=0,
            cursor='hand2',
        )
        self._current_bg = self.normal_bg
        self._pressed = False
        self.bind('<Configure>', lambda _e: self._draw())
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.bind('<ButtonPress-1>', self._on_press)
        self.bind('<ButtonRelease-1>', self._on_release)
        self._draw()

    def _draw(self):
        self.delete('all')
        w = max(4, self.winfo_width())
        h = max(4, self.winfo_height())
        offset = 1 if self._pressed else 0
        _rounded_rect(self, 2, 2 + offset, w - 2, h - 2 + offset, 14,
                      fill=self._current_bg, outline=self.border_color, width=1)
        self.create_text(
            w / 2,
            h / 2 + offset,
            text=self.text,
            fill=self.text_color,
            font=('Segoe UI Semibold', 9 if self.button_height < 42 else 10),
        )

    def _on_enter(self, _event):
        self._current_bg = self.hover_bg
        self._draw()

    def _on_leave(self, _event):
        self._pressed = False
        self._current_bg = self.normal_bg
        self._draw()

    def _on_press(self, _event):
        self._pressed = True
        self._draw()

    def _on_release(self, event):
        was_pressed = self._pressed
        self._pressed = False
        self._current_bg = self.hover_bg
        self._draw()
        if was_pressed and 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height():
            self.command()

    def set_style(self, accent=True):
        self.accent = accent
        self.normal_bg = ACCENT if accent else BUTTON_MUTED
        self.hover_bg = ACCENT_HOVER if accent else BUTTON_MUTED_HOVER
        self.text_color = TEXT_ON_ACCENT if accent else TEXT_PRIMARY
        self.border_color = '' if accent else BORDER
        self._current_bg = self.normal_bg
        self._draw()

    def set_text(self, text):
        self.text = text
        self._draw()

def create_styled_button(parent, text, command, width=20, accent=True, height=46):
    """Create a rounded app button."""
    return RoundedButton(parent, text, command, width=width, accent=accent, height=height)


def set_button_accent(button, accent=True):
    if hasattr(button, "set_style"):
        button.set_style(accent)
        return
    bg = ACCENT if accent else BUTTON_MUTED
    hover_bg = ACCENT_HOVER if accent else BUTTON_MUTED_HOVER
    fg = TEXT_ON_ACCENT if accent else TEXT_PRIMARY
    button.configure(bg=bg, fg=fg, activebackground=hover_bg, highlightbackground=bg)
    button.bind('<Enter>', lambda e, b=button, h=hover_bg: b.configure(bg=h))
    button.bind('<Leave>', lambda e, b=button, normal=bg: b.configure(bg=normal))


def create_segmented_setting(parent, label, variable, choices, hint=""):
    row = tk.Frame(parent, bg=BG_CARD)

    text_col = tk.Frame(row, bg=BG_CARD)
    text_col.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
    create_styled_label(text_col, label, size=9, color=TEXT_PRIMARY, bold=True, bg=BG_CARD).pack(anchor='w')
    if hint:
        hint_label = create_styled_label(text_col, hint, size=8, color=TEXT_SECONDARY, bg=BG_CARD)
        hint_label.pack(anchor='w', pady=(1, 0))

    button_row = tk.Frame(row, bg=BG_CARD)
    button_row.pack(side=tk.RIGHT)
    buttons = []

    def refresh():
        current = variable.get()
        for button, value in buttons:
            set_button_accent(button, current == value)

    def select(value):
        variable.set(value)
        refresh()

    for index, (text, value) in enumerate(choices):
        button = create_styled_button(button_row, text, lambda v=value: select(v), width=9, accent=False, height=34)
        button.pack(side=tk.LEFT, padx=(0, 6) if index < len(choices) - 1 else (0, 0))
        buttons.append((button, value))

    refresh()
    try:
        variable.trace_add('write', lambda *_args: refresh())
    except Exception:
        pass
    return row


class RoundedPanel(tk.Frame):
    def __init__(self, parent, fill=BG_CARD, outline=BORDER, radius=10, padx=22, pady=20):
        super().__init__(parent, bg=_parent_bg(parent), bd=0, highlightthickness=0)
        self.fill = fill
        self.outline = outline
        self.radius = radius
        self.padx = padx
        self.pady = pady
        self.canvas = tk.Canvas(self, width=1, height=1, bg=_parent_bg(parent), bd=0, highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.inner = tk.Frame(self.canvas, bg=fill)
        self._window = self.canvas.create_window(padx, pady, anchor='nw', window=self.inner)
        self.canvas.bind('<Configure>', self._draw)
        self.inner.bind('<Configure>', self._sync_height)

    def _sync_height(self, _event=None):
        inner_height = max(1, self.inner.winfo_reqheight())
        desired_height = inner_height + (self.pady * 2)
        self.canvas.itemconfigure(self._window, height=inner_height)
        if int(self.canvas.cget("height")) != desired_height:
            self.canvas.configure(height=desired_height)

    def _draw(self, _event=None):
        self._sync_height()
        self.canvas.delete('panel_bg')
        w = max(4, self.canvas.winfo_width())
        h = max(4, self.canvas.winfo_height())
        _rounded_rect(self.canvas, 1, 1, w - 1, h - 1, self.radius,
                      fill=self.fill, outline=self.outline, width=1, tags='panel_bg')
        self.canvas.tag_lower('panel_bg')
        self.canvas.coords(self._window, self.padx, self.pady)
        self.canvas.itemconfigure(self._window, width=max(1, w - (self.padx * 2)))

class RoundedLabelBox(tk.Canvas):
    def __init__(self, parent, text, height=7, fill=BUTTON_MUTED, fg=TEXT_SECONDARY,
                 font=('Segoe UI Semibold', 11), radius=10):
        super().__init__(
            parent,
            height=max(56, int(height) * 23),
            bg=_parent_bg(parent),
            bd=0,
            highlightthickness=0,
            cursor='hand2',
        )
        self.text = text
        self.fill = fill
        self.fg = fg
        self.text_font = font
        self.radius = radius
        self.bind('<Configure>', self._draw)
        self._draw()

    def _draw(self, _event=None):
        self.delete('all')
        w = max(4, self.winfo_width())
        h = max(4, self.winfo_height())
        _rounded_rect(self, 1, 1, w - 1, h - 1, self.radius,
                      fill=self.fill, outline=BORDER, width=1)
        self.create_text(w / 2, h / 2, text=self.text, fill=self.fg,
                         font=self.text_font, width=max(40, w - 32),
                         justify='center')

    def configure(self, cnf=None, **kwargs):
        if cnf:
            kwargs.update(cnf)
        redraw = False
        for key, attr in (('text', 'text'), ('fg', 'fg'), ('foreground', 'fg')):
            if key in kwargs:
                setattr(self, attr, kwargs.pop(key))
                redraw = True
        for key in ('bg', 'background'):
            if key in kwargs:
                self.fill = kwargs.pop(key)
                redraw = True
        if 'font' in kwargs:
            self.text_font = kwargs.pop('font')
            redraw = True
        result = super().configure(**kwargs) if kwargs else None
        if redraw:
            self._draw()
        return result

    config = configure

def create_panel(parent, padx=22, pady=20):
    panel = RoundedPanel(parent, fill=BG_CARD, outline=BORDER, radius=10, padx=padx, pady=pady)
    return panel, panel.inner

def finalize_fixed_action_dialog_grid(container, header_panel, table_panel, action_panel):
    for panel in (header_panel, table_panel, action_panel):
        if hasattr(panel, "_sync_height"):
            panel._sync_height()
        panel.update_idletasks()
    container.update_idletasks()
    header_height = header_panel.inner.winfo_reqheight() + (header_panel.pady * 2)
    action_height = action_panel.inner.winfo_reqheight() + (action_panel.pady * 2)
    header_panel.canvas.configure(height=header_height)
    action_panel.canvas.configure(height=action_height)
    container.rowconfigure(0, minsize=header_height)
    container.rowconfigure(1, weight=1, minsize=0)
    container.rowconfigure(2, minsize=action_height)

def create_drop_box(parent, text, height=7):
    return RoundedLabelBox(parent, text, height=height, fill=BUTTON_MUTED, fg=TEXT_SECONDARY)

def create_styled_entry(parent, width=35):
    """Create a modern styled entry field."""
    entry = tk.Entry(parent, width=width,
                     bg=BUTTON_MUTED, fg=TEXT_PRIMARY,
                     insertbackground=TEXT_PRIMARY,
                     font=('Segoe UI', 10),
                     relief='flat', bd=0,
                     highlightthickness=0)
    return entry

def create_styled_label(parent, text, size=10, color=None, bold=False, bg=None):
    """Create a modern styled label."""
    weight = 'bold' if bold else 'normal'
    if color is None:
        color = TEXT_PRIMARY
    if bg is None:
        try:
            bg = parent.cget("bg")
        except Exception:
            bg = BG_DARK
    label = tk.Label(parent, text=text, bg=bg, fg=color,
                     font=('Segoe UI', size, weight))
    return label

def create_styled_text_area(parent, placeholder="", height=6):
    shell = tk.Frame(parent, bg=BORDER, padx=1, pady=1)
    text = tk.Text(
        shell,
        width=58,
        height=height,
        bg=BG_INPUT,
        fg=TEXT_PRIMARY,
        insertbackground=TEXT_PRIMARY,
        font=('Segoe UI', 10),
        relief='flat',
        bd=0,
        padx=12,
        pady=10,
        wrap='word',
    )
    text.pack(fill=tk.BOTH, expand=True)
    text._placeholder = placeholder
    text._placeholder_visible = False

    def show_placeholder():
        if not text._placeholder:
            return
        text._placeholder_visible = True
        text.delete("1.0", tk.END)
        text.insert("1.0", text._placeholder)
        text.configure(fg=TEXT_SECONDARY)

    def clear_placeholder(_event=None):
        if text._placeholder_visible:
            text._placeholder_visible = False
            text.delete("1.0", tk.END)
            text.configure(fg=TEXT_PRIMARY)

    def restore_placeholder(_event=None):
        if not text.get("1.0", tk.END).strip():
            show_placeholder()

    text._show_placeholder = show_placeholder
    text._clear_placeholder = clear_placeholder
    text.bind("<FocusIn>", clear_placeholder)
    text.bind("<FocusOut>", restore_placeholder)
    text.bind("<FocusIn>", lambda _event: shell.configure(bg=ACCENT), add="+")
    text.bind("<FocusOut>", lambda _event: shell.configure(bg=BORDER), add="+")
    show_placeholder()
    return shell, text

def get_text_area_value(text):
    if getattr(text, "_placeholder_visible", False):
        return ""
    return text.get("1.0", tk.END)

def set_text_area_value(text, value):
    text._clear_placeholder()
    text.delete("1.0", tk.END)
    if value:
        text.insert(tk.END, value)
    else:
        text._show_placeholder()

def attach_auto_resize_text(text_widget, min_lines=6, max_lines=8):
    """Auto-resize a Text widget based on line count."""
    def _resize(_event=None):
        try:
            content = text_widget.get("1.0", "end-1c")
        except Exception:
            return
        line_count = content.count("\n") + 1 if content else 1
        new_height = max(min_lines, min(max_lines, line_count))
        if int(text_widget.cget("height")) != new_height:
            text_widget.configure(height=new_height)
            _layout_refresh_callback()

    def _on_modified(_event=None):
        try:
            text_widget.edit_modified(False)
        except Exception:
            pass
        _resize()

    def _on_mousewheel(event):
        try:
            text_widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"
        except Exception:
            return None

    text_widget.bind("<<Modified>>", _on_modified)
    text_widget.bind("<KeyRelease>", _resize)
    text_widget.bind("<Control-v>", _resize)
    text_widget.bind("<Control-V>", _resize)
    text_widget.bind("<MouseWheel>", _on_mousewheel)
    _resize()
    return _resize


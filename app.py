# app.py
"""
Excel -> MySQL Converter - Windows desktop, fully offline.

Features
--------
* Colourful modern UI (gradient header, colour-coded buttons, status colours)
* Drag & drop .xlsx / .xlsm onto the drop zone (animated hover / drag states)
* Excel preview grid with striped rows and coloured scrollbars
* Sliders: Rows per INSERT, Preview rows
* Live SQL panel with syntax highlighting - updates as you type / toggle
* Convert & Preview  |  Save .sql file  |  Cancel
* Progress bar + colour-coded status bar
* Collapsible quick guide
"""

import os
import re
import queue
import threading
import time
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

from converter import (
    get_sheet_names, load_preview,
    excel_to_sql_file, excel_to_sql_stream,
    generate_live_sql_preview,
    validate_table_name, suggest_table_name,
    ConversionCancelled, MAX_BATCH,
)

# ---------------------------------------------------------------- palette
BG        = "#f1f5f9"
CARD      = "#ffffff"
BORDER    = "#e2e8f0"
TEXT      = "#0f172a"
MUTED     = "#64748b"
PRIMARY   = "#4f46e5"
PRIMARY_D = "#4338ca"
PRIMARY_L = "#e0e7ff"
ACCENT    = "#06b6d4"
SUCCESS   = "#16a34a"
SUCCESS_D = "#15803d"
DANGER    = "#dc2626"
DANGER_D  = "#b91c1c"
DARK      = "#0b1120"

GRAD_FROM = (79, 70, 229)     # indigo
GRAD_TO   = (6, 182, 212)     # cyan

STATUS_KINDS = {              # kind: (background, foreground)
    "info": ("#e0e7ff", "#3730a3"),
    "ok":   ("#dcfce7", "#166534"),
    "warn": ("#fef3c7", "#92400e"),
    "err":  ("#fee2e2", "#991b1b"),
    "busy": ("#e0f2fe", "#075985"),
}

# SQL syntax highlighting (applied line by line - fast even for big text)
SQL_TOKEN = re.compile(
    r"(?P<comment>--[^\n]*)"
    r"|(?P<string>'(?:\\.|[^'\\])*')"
    r"|(?P<ident>`(?:[^`]|``)*`)"
    r"|(?P<kw>\b(?:SET|NAMES|DROP|TABLE|IF|EXISTS|CREATE|INSERT|INTO|VALUES|"
    r"ENGINE|DEFAULT|CHARSET|COLLATE|LONGTEXT|FOREIGN_KEY_CHECKS)\b)"
    r"|(?P<null>\bNULL\b)"
    r"|(?P<number>(?<![\w`])-?\d+(?:\.\d+)?(?![\w`]))"
)
SQL_TAG_STYLE = {
    "comment": dict(foreground="#64748b", font=("Consolas", 10, "italic")),
    "string":  dict(foreground="#86efac"),
    "ident":   dict(foreground="#67e8f9"),
    "kw":      dict(foreground="#c4b5fd", font=("Consolas", 10, "bold")),
    "null":    dict(foreground="#fb7185"),
    "number":  dict(foreground="#fcd34d"),
}
HIGHLIGHT_MAX_LINES = 4000


def _blend(c1, c2, t):
    return "#%02x%02x%02x" % tuple(int(a + (b - a) * t) for a, b in zip(c1, c2))


# =========================================================================
class Excel2SQLApp:
    PREVIEW_ROWS       = 100
    PREVIEW_CHAR_LIMIT = 200_000
    DEFAULT_BATCH      = 500
    REFRESH_DELAY_MS   = 300

    def __init__(self, root):
        self.root = root
        self.root.title("Excel to MySQL Converter  -  Offline")
        self.root.geometry("1320x880")
        self.root.minsize(1100, 700)
        self.root.configure(bg=BG)

        # ---- state ----
        self.file_path    = tk.StringVar()
        self.sheet_name   = tk.StringVar()
        self.table_name   = tk.StringVar(value="my_table")
        self.batch_size   = tk.IntVar(value=self.DEFAULT_BATCH)
        self.preview_rows = tk.IntVar(value=self.PREVIEW_ROWS)
        self.opt_drop     = tk.BooleanVar(value=True)
        self.opt_create   = tk.BooleanVar(value=True)
        self.opt_batch    = tk.BooleanVar(value=True)

        self.total_rows = 0
        self._cancel    = threading.Event()
        self._busy      = False

        self._ui_queue       = queue.Queue()   # worker thread -> UI thread
        self._refresh_job    = None            # debounced SQL refresh
        self._preview_job    = None            # debounced preview reload
        self._auto_name      = ""              # last auto-suggested table name
        self._last_prog_time = 0.0
        self._drop_state     = "normal"
        self._guide_open     = False

        self._setup_styles()
        self._build_ui()
        self._poll_ui_queue()
        self._refresh_live_sql()

    # ================================================================ THREAD -> UI
    def _post(self, fn, *args):
        """Safely schedule fn(*args) on the Tk main thread."""
        self._ui_queue.put((fn, args))

    def _poll_ui_queue(self):
        try:
            while True:
                fn, args = self._ui_queue.get_nowait()
                try:
                    fn(*args)
                except Exception:
                    traceback.print_exc()
        except queue.Empty:
            pass
        self.root.after(50, self._poll_ui_queue)

    # ================================================================ STYLES
    def _setup_styles(self):
        st = ttk.Style()
        try:
            st.theme_use("clam")       # clam honours colours (vista ignores them)
        except tk.TclError:
            pass

        st.configure(".", background=BG, foreground=TEXT,
                     font=("Segoe UI", 10), focuscolor=PRIMARY)

        # --- cards ---
        st.configure("Card.TFrame", background=CARD)
        st.configure("Card.TLabel", background=CARD, foreground=TEXT)
        st.configure("Muted.TLabel", background=CARD, foreground=MUTED,
                     font=("Segoe UI", 9))
        st.configure("Card.TLabelframe", background=CARD, bordercolor=BORDER,
                     lightcolor=BORDER, darkcolor=BORDER, relief="solid",
                     borderwidth=1)
        st.configure("Card.TLabelframe.Label", background=CARD,
                     foreground=PRIMARY, font=("Segoe UI", 10, "bold"))
        st.configure("Card.TCheckbutton", background=CARD, foreground=TEXT,
                     font=("Segoe UI", 10))
        st.map("Card.TCheckbutton", background=[("active", CARD)],
               foreground=[("active", PRIMARY)])

        # --- inputs ---
        st.configure("TEntry", fieldbackground="white", bordercolor=BORDER,
                     lightcolor=BORDER, darkcolor=BORDER, padding=5)
        st.map("TEntry", bordercolor=[("focus", PRIMARY)],
               lightcolor=[("focus", PRIMARY)], darkcolor=[("focus", PRIMARY)])
        st.configure("TCombobox", fieldbackground="white", background=PRIMARY_L,
                     bordercolor=BORDER, arrowcolor=PRIMARY, padding=4)
        st.map("TCombobox", bordercolor=[("focus", PRIMARY)],
               fieldbackground=[("readonly", "white")],
               selectbackground=[("readonly", "white")],
               selectforeground=[("readonly", TEXT)])
        st.configure("TSpinbox", fieldbackground="white", bordercolor=BORDER,
                     arrowcolor=PRIMARY, padding=4)
        self.root.option_add("*TCombobox*Listbox.font", ("Segoe UI", 10))
        self.root.option_add("*TCombobox*Listbox.selectBackground", PRIMARY)
        self.root.option_add("*TCombobox*Listbox.selectForeground", "white")

        # --- buttons ---
        def button(name, bg, hover, fg="white", pad=(18, 9)):
            st.configure(name, background=bg, foreground=fg, bordercolor=bg,
                         lightcolor=bg, darkcolor=bg, focuscolor=bg,
                         padding=pad, font=("Segoe UI", 10, "bold"),
                         relief="flat", borderwidth=0)
            st.map(name,
                   background=[("disabled", "#cbd5e1"), ("active", hover)],
                   foreground=[("disabled", "#f8fafc")],
                   bordercolor=[("disabled", "#cbd5e1"), ("active", hover)],
                   lightcolor=[("disabled", "#cbd5e1"), ("active", hover)],
                   darkcolor=[("disabled", "#cbd5e1"), ("active", hover)])

        button("Primary.TButton", PRIMARY, PRIMARY_D)
        button("Success.TButton", SUCCESS, SUCCESS_D)
        button("Danger.TButton",  DANGER,  DANGER_D)
        button("Ghost.TButton",   PRIMARY_L, "#c7d2fe", fg=PRIMARY_D, pad=(12, 6))

        # --- progress bar ---
        st.configure("Accent.Horizontal.TProgressbar", troughcolor="#e2e8f0",
                     background=PRIMARY, lightcolor=ACCENT, darkcolor=PRIMARY,
                     bordercolor="#e2e8f0", thickness=16)

        # --- scrollbars (the coloured "sliders" on the grid / SQL panel) ---
        for orient in ("Vertical", "Horizontal"):
            st.configure(f"{orient}.TScrollbar", background="#a5b4fc",
                         troughcolor="#eef2ff", bordercolor="#eef2ff",
                         lightcolor="#a5b4fc", darkcolor="#a5b4fc",
                         arrowcolor="white", arrowsize=16, gripcount=0)
            st.map(f"{orient}.TScrollbar",
                   background=[("pressed", PRIMARY_D), ("active", PRIMARY)],
                   lightcolor=[("pressed", PRIMARY_D), ("active", PRIMARY)],
                   darkcolor=[("pressed", PRIMARY_D), ("active", PRIMARY)])
            st.configure(f"Dark.{orient}.TScrollbar", background="#475569",
                         troughcolor="#111a2e", bordercolor="#111a2e",
                         lightcolor="#475569", darkcolor="#475569",
                         arrowcolor="#cbd5e1", arrowsize=16, gripcount=0)
            st.map(f"Dark.{orient}.TScrollbar",
                   background=[("pressed", ACCENT), ("active", "#6366f1")],
                   lightcolor=[("pressed", ACCENT), ("active", "#6366f1")],
                   darkcolor=[("pressed", ACCENT), ("active", "#6366f1")])

        # --- treeview ---
        st.configure("Treeview", background="white", fieldbackground="white",
                     foreground=TEXT, rowheight=26, borderwidth=0,
                     font=("Segoe UI", 9))
        st.configure("Treeview.Heading", background=PRIMARY, foreground="white",
                     font=("Segoe UI", 9, "bold"), relief="flat", padding=6)
        st.map("Treeview.Heading", background=[("active", PRIMARY_D)])
        st.map("Treeview", background=[("selected", "#c7d2fe")],
               foreground=[("selected", TEXT)])

        st.configure("TPanedwindow", background=BG)

    # ================================================================ UI
    def _build_ui(self):
        # ---- gradient header banner ----
        self.banner = tk.Canvas(self.root, height=76, highlightthickness=0, bd=0)
        self.banner.pack(fill="x")
        self.banner.bind("<Configure>", self._draw_banner)

        # ---- Status bar (packed first so it is never squeezed out) ----
        self.status_label = tk.Label(self.root, text="Ready.", anchor="w",
                                     padx=16, pady=7,
                                     font=("Segoe UI", 10, "bold"))
        self.status_label.pack(side="bottom", fill="x")
        self._status("Ready.", "info")

        self.body = body = ttk.Frame(self.root, padding=(14, 12, 14, 0))
        body.pack(fill="both", expand=True)

        # ---- drop zone ----
        self.drop_zone = tk.Canvas(body, height=92, highlightthickness=0, bd=0,
                                   cursor="hand2")
        self.drop_zone.pack(fill="x")
        self.drop_zone.bind("<Configure>", lambda e: self._draw_drop())
        self.drop_zone.bind("<Button-1>", lambda e: self.browse_file())
        self.drop_zone.bind("<Enter>", lambda e: self._draw_drop("hover"))
        self.drop_zone.bind("<Leave>", lambda e: self._draw_drop("normal"))

        if DND_AVAILABLE:
            self.drop_zone.drop_target_register(DND_FILES)
            self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)
            self.drop_zone.dnd_bind("<<DropEnter>>", self._on_drop_enter)
            self.drop_zone.dnd_bind("<<DropLeave>>", self._on_drop_leave)

        self.file_info = ttk.Label(body, text="No file selected.",
                                   foreground=MUTED, padding=(4, 8, 4, 6))
        self.file_info.pack(fill="x")

        # ---- Progress + action buttons ----
        bottom = ttk.Frame(body, padding=(0, 12, 0, 6))
        bottom.pack(side="bottom", fill="x")

        self.progress = ttk.Progressbar(bottom, mode="determinate", maximum=100,
                                        style="Accent.Horizontal.TProgressbar")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 14))

        self.btn_guide = ttk.Button(bottom, text="\U0001f4d8 Guide  \u25b8",
                                    style="Ghost.TButton",
                                    command=self._toggle_guide)
        self.btn_guide.pack(side="left", padx=3)

        self.btn_convert = ttk.Button(bottom, text="\u2705 Convert & Preview",
                                      style="Primary.TButton",
                                      command=self.start_convert_preview)
        self.btn_convert.pack(side="left", padx=3)

        self.btn_save = ttk.Button(bottom, text="\U0001f4be Save .sql file",
                                   style="Success.TButton",
                                   command=self.start_save_file)
        self.btn_save.pack(side="left", padx=3)

        self.btn_cancel = ttk.Button(bottom, text="\u2716 Cancel",
                                     style="Danger.TButton",
                                     command=self.cancel_convert,
                                     state="disabled")
        self.btn_cancel.pack(side="left", padx=3)

        # ---- Split: LEFT (preview+options) / RIGHT (live SQL) ----
        paned = ttk.Panedwindow(body, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # ============ LEFT ============
        left = ttk.Frame(paned)
        paned.add(left, weight=3)

        # --- Excel preview ---
        self.prev_frame = ttk.LabelFrame(left, style="Card.TLabelframe",
                                         text="  \U0001f4c4 Excel Preview  ",
                                         padding=8)
        px = ttk.Scrollbar(self.prev_frame, orient="horizontal",
                           style="Horizontal.TScrollbar")
        px.pack(side="bottom", fill="x")
        py = ttk.Scrollbar(self.prev_frame, orient="vertical",
                           style="Vertical.TScrollbar")
        py.pack(side="right", fill="y")

        self.preview_tree = ttk.Treeview(self.prev_frame, show="headings",
                                         height=12, yscrollcommand=py.set,
                                         xscrollcommand=px.set)
        self.preview_tree.pack(side="left", fill="both", expand=True)
        self.preview_tree.tag_configure("odd", background="#f8fafc")
        self.preview_tree.tag_configure("even", background="#ffffff")
        py.configure(command=self.preview_tree.yview)
        px.configure(command=self.preview_tree.xview)

        # --- Options ---
        opt = ttk.LabelFrame(left, style="Card.TLabelframe",
                             text="  \u2699 Options  ", padding=12)
        opt.pack(side="bottom", fill="x", pady=(10, 0), padx=(0, 6))
        # preview is packed after Options so Options can never be squeezed out
        self.prev_frame.pack(fill="both", expand=True, padx=(0, 6))
        opt.columnconfigure(1, weight=1)

        # row 0: Worksheet
        ttk.Label(opt, text="Worksheet", style="Card.TLabel")\
            .grid(row=0, column=0, sticky="w", pady=4)
        self.sheet_combo = ttk.Combobox(opt, textvariable=self.sheet_name,
                                        state="readonly")
        self.sheet_combo.grid(row=0, column=1, sticky="ew", padx=(10, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_change)

        # row 1: Table name + Auto button
        ttk.Label(opt, text="Table name", style="Card.TLabel")\
            .grid(row=1, column=0, sticky="w", pady=4)
        table_row = ttk.Frame(opt, style="Card.TFrame")
        table_row.grid(row=1, column=1, sticky="ew", padx=(10, 0))
        table_row.columnconfigure(0, weight=1)

        self.table_entry = ttk.Entry(table_row, textvariable=self.table_name)
        self.table_entry.grid(row=0, column=0, sticky="ew")
        ttk.Button(table_row, text="\u2728 Auto", style="Ghost.TButton",
                   command=self._auto_suggest_table)\
            .grid(row=0, column=1, padx=(8, 0))

        # row 2: inline hint
        self.table_hint = ttk.Label(opt, text="", style="Muted.TLabel")
        self.table_hint.grid(row=2, column=0, columnspan=2,
                             sticky="w", pady=(0, 4))

        # row 3: tick options
        tick_frame = ttk.Frame(opt, style="Card.TFrame")
        tick_frame.grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 6))
        for text, var in (("DROP TABLE first", self.opt_drop),
                          ("CREATE TABLE", self.opt_create),
                          ("BATCH INSERT", self.opt_batch)):
            ttk.Checkbutton(tick_frame, text=text, variable=var,
                            style="Card.TCheckbutton",
                            command=self._schedule_refresh)\
                .pack(side="left", padx=(0, 22))

        # row 4: rows per INSERT  (slider + spinbox share one variable)
        ttk.Label(opt, text="Rows per INSERT", style="Card.TLabel")\
            .grid(row=4, column=0, sticky="w", pady=6)
        bs_row = ttk.Frame(opt, style="Card.TFrame")
        bs_row.grid(row=4, column=1, sticky="ew", padx=(10, 0))
        bs_row.columnconfigure(0, weight=1)

        self.batch_scale = self._make_slider(bs_row, self.batch_size, 1, 2000, 1)
        self.batch_scale.grid(row=0, column=0, sticky="ew")
        ttk.Spinbox(bs_row, from_=1, to=2000, increment=50,
                    textvariable=self.batch_size, width=7,
                    command=self._schedule_refresh)\
            .grid(row=0, column=1, padx=(12, 0))

        # row 5: preview rows slider
        ttk.Label(opt, text="Preview rows", style="Card.TLabel")\
            .grid(row=5, column=0, sticky="w", pady=6)
        pv_row = ttk.Frame(opt, style="Card.TFrame")
        pv_row.grid(row=5, column=1, sticky="ew", padx=(10, 0))
        pv_row.columnconfigure(0, weight=1)

        self.preview_scale = self._make_slider(pv_row, self.preview_rows,
                                               10, 500, 10)
        self.preview_scale.grid(row=0, column=0, sticky="ew")
        self.preview_lbl = ttk.Label(pv_row, text="100", width=7,
                                     style="Card.TLabel", anchor="center",
                                     font=("Segoe UI", 10, "bold"),
                                     foreground=PRIMARY)
        self.preview_lbl.grid(row=0, column=1, padx=(12, 0))

        # trace hooks (debounced: no file re-read on every keystroke / drag)
        self.table_name.trace_add("write", lambda *a: self._schedule_refresh())
        self.batch_size.trace_add("write", lambda *a: self._schedule_refresh())
        self.preview_rows.trace_add("write", lambda *a: self._on_preview_slider())

        # ============ RIGHT ============
        right = ttk.LabelFrame(paned, style="Card.TLabelframe",
                               text="  \u26a1 Live SQL Commands  ", padding=8)
        paned.add(right, weight=4)

        sql_frame = tk.Frame(right, bg=DARK)
        sql_frame.pack(fill="both", expand=True)
        sql_frame.rowconfigure(0, weight=1)
        sql_frame.columnconfigure(0, weight=1)

        self.live_sql = tk.Text(
            sql_frame, wrap="none", font=("Consolas", 10), bg=DARK,
            fg="#e2e8f0", insertbackground="#e2e8f0", relief="flat", bd=0,
            padx=12, pady=10, selectbackground="#334155", undo=False,
        )
        sy = ttk.Scrollbar(sql_frame, orient="vertical",
                           style="Dark.Vertical.TScrollbar",
                           command=self.live_sql.yview)
        sx = ttk.Scrollbar(sql_frame, orient="horizontal",
                           style="Dark.Horizontal.TScrollbar",
                           command=self.live_sql.xview)
        self.live_sql.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.live_sql.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        for tag, cfg in SQL_TAG_STYLE.items():
            self.live_sql.tag_configure(tag, **cfg)
        self.live_sql.configure(state="disabled")

        # ---- Quick Guide (collapsible; shown by the Guide button) ----
        self.guide = ttk.LabelFrame(self.root, style="Card.TLabelframe",
                                    text="  \U0001f4d8 Quick Guide  ", padding=10)
        steps = [
            "1.  Drag & drop your .xlsx file onto the drop zone (or click it).",
            "2.  Pick the worksheet from the dropdown.",
            "3.  Enter the MySQL table name (or click 'Auto' to derive one).",
            "4.  Tick DROP TABLE / CREATE TABLE / BATCH INSERT as needed.",
            "5.  Use the sliders for Rows per INSERT and how many rows to preview.",
            "6.  'Convert & Preview' shows the full SQL; 'Save .sql file' writes it to disk.",
            "7.  Import into MySQL:  mysql -u root -p your_db < yourfile.sql",
        ]
        for s in steps:
            ttk.Label(self.guide, text=s, style="Card.TLabel")\
                .pack(anchor="w", pady=1)

    def _make_slider(self, parent, variable, lo, hi, step):
        """Coloured horizontal slider (tk.Scale is fully stylable)."""
        return tk.Scale(
            parent, variable=variable, from_=lo, to=hi, resolution=step,
            orient="horizontal", showvalue=False, length=200,
            bg=PRIMARY, troughcolor=PRIMARY_L, activebackground=PRIMARY_D,
            highlightthickness=0, bd=0, relief="flat", sliderrelief="flat",
            sliderlength=26, width=16, cursor="hand2",
        )

    # ================================================================ DRAWING
    def _draw_banner(self, event=None):
        c = self.banner
        c.delete("all")
        w, h = max(c.winfo_width(), 100), max(c.winfo_height(), 40)
        step = 4
        for x in range(0, w, step):
            c.create_rectangle(x, 0, x + step, h, outline="",
                               fill=_blend(GRAD_FROM, GRAD_TO, x / w))
        c.create_text(26, h / 2 - 11, anchor="w", fill="white",
                      text="Excel  \u2192  MySQL Converter",
                      font=("Segoe UI", 19, "bold"))
        c.create_text(28, h / 2 + 17, anchor="w", fill="#e0f2fe",
                      text="Offline  \u2022  exact copy of your cells  \u2022  "
                           "fast streaming for 100k+ rows",
                      font=("Segoe UI", 10))

    def _draw_drop(self, state=None):
        if state:
            self._drop_state = state
        palette = {
            "normal": ("#eef2ff", "#818cf8", "#4338ca"),
            "hover":  ("#e0e7ff", PRIMARY,   "#3730a3"),
            "drag":   ("#dcfce7", SUCCESS,   "#166534"),
        }
        fill, outline, fg = palette[self._drop_state]
        c = self.drop_zone
        c.configure(bg=BG)
        c.delete("all")
        w, h = max(c.winfo_width(), 200), max(c.winfo_height(), 80)
        c.create_rectangle(3, 3, w - 3, h - 3, outline=outline, width=2,
                           dash=(9, 5), fill=fill)

        if self._drop_state == "drag":
            title, sub = "\u2b07  Release to load this file", ""
        elif DND_AVAILABLE:
            title = "\u2b07  Drag & drop your Excel file here"
            sub = "or click to browse  \u2022  .xlsx  .xlsm"
        else:
            title = "\U0001f4c2  Click to browse for an Excel file"
            sub = "drag & drop unavailable - run:  pip install tkinterdnd2"
        c.create_text(w / 2, h / 2 - (10 if sub else 0), text=title, fill=fg,
                      font=("Segoe UI", 14, "bold"))
        if sub:
            c.create_text(w / 2, h / 2 + 18, text=sub, fill=MUTED,
                          font=("Segoe UI", 10))

    # ================================================================ HELPERS
    def _status(self, text, kind="info"):
        bg, fg = STATUS_KINDS.get(kind, STATUS_KINDS["info"])
        self.status_label.configure(text=text, bg=bg, fg=fg)

    def _toggle_guide(self):
        if self._guide_open:
            self.guide.pack_forget()
            self.btn_guide.configure(text="\U0001f4d8 Guide  \u25b8")
        else:
            self.guide.pack(side="bottom", fill="x", padx=14, pady=(0, 8),
                            before=self.body)
            self.btn_guide.configure(text="\U0001f4d8 Guide  \u25be")
        self._guide_open = not self._guide_open

    def _safe_batch_size(self):
        """Spinbox value as int clamped to 1..MAX_BATCH; DEFAULT_BATCH if invalid."""
        try:
            raw = int(self.batch_size.get())
        except (tk.TclError, ValueError):
            return self.DEFAULT_BATCH
        if raw <= 0:
            return self.DEFAULT_BATCH
        return min(raw, MAX_BATCH)

    def _safe_preview_rows(self):
        try:
            return max(1, int(self.preview_rows.get()))
        except (tk.TclError, ValueError):
            return self.PREVIEW_ROWS

    def _validate_table_field(self):
        name = self.table_name.get().strip()
        ok, msg = validate_table_name(name)
        if ok:
            if msg:
                self.table_hint.configure(text=f"\u26a0 {msg}", foreground="#b45309")
            else:
                self.table_hint.configure(text=f"\u2714 Will create table `{name}`",
                                          foreground=SUCCESS)
        else:
            self.table_hint.configure(text=f"\u2716 {msg}", foreground=DANGER)
        return ok

    def _auto_suggest_table(self):
        path  = self.file_path.get()
        sheet = self.sheet_name.get()
        if not path:
            messagebox.showinfo("No file", "Load an Excel file first.")
            return
        self._auto_name = suggest_table_name(path, sheet)
        self.table_name.set(self._auto_name)

    def _maybe_auto_name(self):
        """Replace the table name only if the user hasn't customised it."""
        if self.table_name.get() in ("", "my_table", self._auto_name):
            self._auto_name = suggest_table_name(self.file_path.get(),
                                                 self.sheet_name.get())
            self.table_name.set(self._auto_name)

    def _set_live_text(self, text):
        t = self.live_sql
        t.configure(state="normal")
        t.delete("1.0", tk.END)
        t.insert(tk.END, text)
        self._highlight_sql(text)
        t.configure(state="disabled")

    def _highlight_sql(self, text):
        t = self.live_sql
        for tag in SQL_TAG_STYLE:
            t.tag_remove(tag, "1.0", tk.END)
        for ln, line in enumerate(text.split("\n"), start=1):
            if ln > HIGHLIGHT_MAX_LINES:
                break
            for m in SQL_TOKEN.finditer(line):
                t.tag_add(m.lastgroup, f"{ln}.{m.start()}", f"{ln}.{m.end()}")

    # ================================================================ EVENTS
    def _on_drop_enter(self, event):
        self._draw_drop("drag")
        return event.action

    def _on_drop_leave(self, event):
        self._draw_drop("normal")
        return event.action

    def _on_drop(self, event):
        self._draw_drop("normal")
        paths = self.root.tk.splitlist(event.data)
        if paths:
            path = paths[0]
            if not path.lower().endswith((".xlsx", ".xlsm")):
                messagebox.showwarning("Unsupported file",
                                       "Please drop an .xlsx or .xlsm file.")
            else:
                self._load_file(path)
        return event.action

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Choose Excel file",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if path:
            self._load_file(path)

    def _load_file(self, path):
        try:
            sheets = get_sheet_names(path)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot read file:\n{e}")
            self._status("Could not read that file.", "err")
            return

        self.file_path.set(path)
        size_kb = os.path.getsize(path) / 1024
        self.file_info.configure(
            text=f"\U0001f4c1  {os.path.basename(path)}   \u2022   {len(sheets)} sheet(s)   "
                 f"\u2022   {size_kb:,.1f} KB   \u2022   {path}",
            foreground=PRIMARY_D,
        )
        self.sheet_combo["values"] = sheets
        self.sheet_name.set(sheets[0] if sheets else "")

        self._maybe_auto_name()
        self._on_sheet_change()

    def _on_sheet_change(self, *_):
        if self._preview_job is not None:
            try:
                self.root.after_cancel(self._preview_job)
            except tk.TclError:
                pass
            self._preview_job = None

        path  = self.file_path.get()
        sheet = self.sheet_name.get()
        if not path or not sheet:
            return

        n = self._safe_preview_rows()
        try:
            headers, rows, total = load_preview(path, sheet, n)
            self.total_rows = total
        except Exception as e:
            messagebox.showerror("Preview error", str(e))
            self._status("Preview failed.", "err")
            return

        self._populate_tree(headers, rows)
        self.prev_frame.configure(text=f"  \U0001f4c4 Excel Preview (first {n} rows)  ")
        if total:
            self._status(f"Loaded ~{total:,} rows  \u2022  {len(headers)} columns  "
                         f"(preview shows first {min(total, n)}).", "ok")
        else:
            self._status(f"Loaded {len(headers)} columns  "
                         f"(preview shows up to {n} rows).", "ok")
        self._maybe_auto_name()
        self._refresh_live_sql()

    def _on_preview_slider(self):
        """Slider moved: update label now, reload preview once dragging pauses."""
        self.preview_lbl.configure(text=str(self._safe_preview_rows()))
        if self._preview_job is not None:
            self.root.after_cancel(self._preview_job)
        self._preview_job = self.root.after(self.REFRESH_DELAY_MS,
                                            self._on_sheet_change)

    def _populate_tree(self, headers, rows):
        tree = self.preview_tree
        tree.delete(*tree.get_children())
        # internal column ids -> duplicate / odd header names can't break Treeview
        ids = [f"c{i}" for i in range(len(headers))]
        tree["columns"] = ids
        for cid, h in zip(ids, headers):
            tree.heading(cid, text=h)
            tree.column(cid, width=130, stretch=False, anchor="w")
        for i, r in enumerate(rows):
            tree.insert("", "end",
                        values=["" if v is None else str(v) for v in r],
                        tags=("odd" if i % 2 else "even",))

    # ================================================================ LIVE SQL
    def _schedule_refresh(self, *_):
        """Debounce: refresh once typing/clicking pauses."""
        if self._refresh_job is not None:
            self.root.after_cancel(self._refresh_job)
        self._refresh_job = self.root.after(self.REFRESH_DELAY_MS,
                                            self._refresh_live_sql)

    def _refresh_live_sql(self, *_):
        if self._refresh_job is not None:
            try:
                self.root.after_cancel(self._refresh_job)
            except tk.TclError:
                pass
            self._refresh_job = None

        self._validate_table_field()
        try:
            sql = generate_live_sql_preview(
                self.file_path.get(),
                self.sheet_name.get(),
                self.table_name.get().strip() or "my_table",
                batch_size=self._safe_batch_size(),
                drop_table=self.opt_drop.get(),
                create_table=self.opt_create.get(),
                batch_insert=self.opt_batch.get(),
                preview_rows=self._safe_preview_rows(),
            )
        except Exception as e:
            sql = f"-- Could not build preview: {e}"
        self._set_live_text(sql)

    # ================================================================ VALIDATION
    def _validate_inputs(self):
        path  = self.file_path.get()
        sheet = self.sheet_name.get()
        table = self.table_name.get().strip()

        if not path:
            messagebox.showwarning("No file", "Load an Excel file first.")
            return None
        if not sheet:
            messagebox.showwarning("No sheet", "Pick a worksheet.")
            return None
        if not table:
            messagebox.showwarning("No table", "Enter a MySQL table name.")
            return None

        ok, msg = validate_table_name(table)
        if not ok:
            messagebox.showwarning("Invalid table name", msg)
            return None

        return path, sheet, table

    def _set_busy(self, busy):
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.btn_convert.config(state=state)
        self.btn_save.config(state=state)
        self.btn_cancel.config(state="normal" if busy else "disabled")

    def _common_opts(self):
        return dict(
            batch_size=self._safe_batch_size(),
            drop_table=self.opt_drop.get(),
            create_table=self.opt_create.get(),
            batch_insert=self.opt_batch.get(),
        )

    # ---- UI-thread callbacks used by the workers ----
    def _show_error(self, title, msg):
        messagebox.showerror(title, msg)

    def _finish(self, status=None, kind="info"):
        self._set_busy(False)
        if status:
            self._status(status, kind)

    # ================================================================ CONVERT
    def start_convert_preview(self):
        if self._busy:
            return
        vals = self._validate_inputs()
        if not vals:
            return
        path, sheet, table = vals

        self._cancel.clear()
        self._set_busy(True)
        self.progress["value"] = 0
        self._status("Converting...", "busy")

        threading.Thread(
            target=self._run_preview,
            args=(path, sheet, table, self._common_opts()),
            daemon=True,
        ).start()

    def _run_preview(self, path, sheet, table, opts):
        gen = None
        try:
            chunks, total_chars, truncated = [], 0, False
            gen = excel_to_sql_stream(path, sheet, table,
                                      progress_cb=self._progress,
                                      cancel_event=self._cancel,
                                      **opts)
            for chunk, _processed in gen:
                room = self.PREVIEW_CHAR_LIMIT - total_chars
                if room <= 0:
                    truncated = True
                    break
                if len(chunk) > room:
                    chunks.append(chunk[:room])
                    total_chars += room
                    truncated = True
                    break
                chunks.append(chunk)
                total_chars += len(chunk)

            text = "".join(chunks)
            if truncated:
                text += ("\n\n-- Preview truncated at 200 KB. "
                         "Use 'Save .sql file' for the full output.\n")

            self._post(self._set_live_text, text)
            self._post(self._set_progress_value, 100)
            self._post(self._finish,
                       f"\u2705 Preview ready - {total_chars:,} characters shown.",
                       "ok")

        except ConversionCancelled:
            self._post(self._finish, "\u26a0 Cancelled.", "warn")
        except Exception as e:
            self._post(self._show_error, "Convert error", str(e))
            self._post(self._finish, "Convert failed.", "err")
        finally:
            if gen is not None:
                gen.close()     # releases the workbook if we stopped early

    # ================================================================ SAVE
    def start_save_file(self):
        if self._busy:
            return
        vals = self._validate_inputs()
        if not vals:
            return
        path, sheet, table = vals

        out = filedialog.asksaveasfilename(
            defaultextension=".sql",
            filetypes=[("SQL file", "*.sql")],
            initialfile=f"{table}.sql",
            title="Save generated SQL to...",
        )
        if not out:
            return

        self._cancel.clear()
        self._set_busy(True)
        self.progress["value"] = 0
        self._status("Writing .sql file...", "busy")

        threading.Thread(
            target=self._run_save,
            args=(path, sheet, table, self._common_opts(), out),
            daemon=True,
        ).start()

    def _run_save(self, path, sheet, table, opts, out_path):
        try:
            n = excel_to_sql_file(
                path, sheet, table, out_path,
                progress_cb=self._progress,
                cancel_event=self._cancel,
                **opts,
            )
            self._post(self._set_progress_value, 100)
            self._post(self._finish,
                       f"\u2705 Saved {n:,} rows -> {os.path.basename(out_path)}",
                       "ok")
            self._post(messagebox.showinfo, "Saved",
                       f"SQL file written:\n{out_path}\n\nRows: {n:,}")
        except ConversionCancelled:
            self._post(self._finish, "\u26a0 Cancelled (partial file removed).",
                       "warn")
        except Exception as e:
            self._post(self._show_error, "Save error", str(e))
            self._post(self._finish, "Save failed.", "err")

    # ================================================================ CANCEL
    def cancel_convert(self):
        self._cancel.set()
        self._status("Cancelling...", "warn")

    # ================================================================ PROGRESS
    def _set_progress_value(self, pct):
        self.progress.configure(value=pct)

    def _set_progress(self, processed, total):
        if not self._busy:
            return
        if total:
            self.progress.configure(value=min(processed / total * 100, 100))
            self._status(f"Processed {processed:,} / ~{total:,} rows...", "busy")
        else:
            self._status(f"Processed {processed:,} rows...", "busy")

    def _progress(self, processed):
        """Called from the worker thread; throttled to ~10 updates/second."""
        now = time.monotonic()
        if now - self._last_prog_time < 0.1:
            return
        self._last_prog_time = now
        self._post(self._set_progress, processed, self.total_rows)


# =========================================================================
def main():
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    Excel2SQLApp(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
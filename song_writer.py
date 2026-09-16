#!/usr/bin/env python3
# ==============================================================================
#  Song Notation Tool
#  A lightweight Tkinter desktop app for sketching guitar/bass songs by section.
#  Standard library only — no pip installs required.
#  Compatible with macOS, Windows, and Linux.
# ==============================================================================
#
#  CHANGELOG
#  ---------
#  v0.1  Initial release — sections, layers (tab/chords/notes/lyrics), export TXT
#  v0.2  Light/dark theme engine (ported from QLC+ Swiss Knife), ToolTips,
#        ttk styling, combobox dropdown fix, layer toggle buttons
#  v0.3  Added artist name field, section copy dialog, inline comments,
#        version constant (APP_VERSION), versioned file naming
#  v0.4  Two-row topbar (metadata on row 2 → no more cramping),
#        Drop D toggle and Transpose dialog added,
#        file naming fixed to song_writer_v.X.Y.py
#  v0.5  Fixed white-on-white button text on macOS (added ("!disabled", fg)
#        to ttk map — the only state entry macOS Aqua honours at rest);
#        Replaced vertical multi-line tab Text widget with one Entry per string
#        laid out horizontally, pre-filled with "- - - - " dashes
#  v0.6  Fixed layer-toggle button colours on macOS (ttk named styles);
#        replaced tk.Scrollbar with custom Canvas-drawn scroll indicators;
#        fixed copy/paste measure overwrite bug
#  v0.7  Section linking — bind sections so edits propagate automatically
#  v0.8  Resizable left panel sash; wrap mode for measure grid;
#        fixed IndexError when resizing a linked section
#  v0.10  TXT/PDF export fixes; Artist before Title; no brackets on notes/chords
#  v0.11  Chromatic note list with enharmonic aliases; transpose all layers;
#         chord/note entry widgets upgraded to root-note picker + suffix
#  v0.12  Variable tab beats (8/16/32); protected dashes; smart transposition
#  v0.13  Beats selector moved into the editor toolbar (always visible);
#         hamburger ≡ menu in topbar for all file actions on narrow windows
#  v0.16  Foundation release (no UI rewrite) — new data model (model.py),
#         chart grammar (grammar.py), render-time transposition
#         (transpose.py), shared TXT/PDF render helpers (render.py).
#         The measure-grid editor and its Section class kept working
#         against the old per-layer schema; the new modules shipped
#         unwired, ready for the UI rewrite below.
#  v0.17  Song map + chart grammar editor. The UI now speaks the v0.16
#         document schema directly: a scrollable map of section rows
#         (rendered fret/symbol chart lines, click-to-focus, drag or
#         Alt+↑/↓ to reorder), one live-parsing text field as the entire
#         editing surface for chart sections, and a riff library strip
#         where editing a block updates every section that references
#         it. The old per-measure grid survives only for render:"tab"
#         sections, now reading/writing `measure` items instead of the
#         retired Section.layers dict. TXT/PDF export and a new Stage
#         View are driven by the same render.py row renderer at two
#         scales. See DESIGN_v0_17.md.
#
# ==============================================================================

import datetime  # timestamp in PDF footer
import json      # .sng project files are plain JSON
import os
import platform
import zlib      # PDF page stream compression (FlateDecode)
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

# v0.16/v0.17 — pure-data modules (no Tk dependency): schema/migration, the
# chart grammar, render-time transposition, shared render/export helpers,
# and the song-map / riff-library logic.
import model
import grammar
import render
import transpose
import songmap
from grammar import ParseError as ChartParseError

# ── macOS: suppress deprecation noise ────────────────────────────────────────
if platform.system() == "Darwin":
    os.environ['SYSTEM_VERSION_COMPAT'] = '0'
    os.environ['TK_SILENCE_DEPRECATION'] = '1'

# ==============================================================================
#  VERSION — single source of truth; bumped here propagates everywhere
# ==============================================================================
APP_VERSION = "0.17"
APP_TITLE   = f"Song Notation Tool  v{APP_VERSION}"

# ==============================================================================
#  DOMAIN CONSTANTS
# ==============================================================================

# Section-type presets shown in the New Section dialog
SECTION_TYPES = [
    "Intro", "Verse", "Pre-Chorus", "Chorus", "Refrain",
    "Bridge", "Interlude", "Solo", "Breakdown", "Outro", "Custom",
]

# Maps instrument/tuning → string list (high → low pitch).
# Used to label tab rows and render the TXT/PDF export.
INSTRUMENT_STRINGS = {
    "Guitar (6-string)":        ["e", "B", "G", "D", "A", "E"],
    "Guitar (6-string) Drop D": ["e", "B", "G", "D", "A", "D"],
    "Guitar (7-string)":        ["e", "B", "G", "D", "A", "E", "B"],
    "Bass (4-string)":          ["G", "D", "A", "E"],
    "Bass (4-string) Drop D":   ["G", "D", "A", "D"],
    "Bass (5-string)":          ["G", "D", "A", "E", "B"],
}

RENDER_MODE_LABELS = {"chart": "Chart", "tab": "Tab grid", "both": "Chart + Tab"}

# Default number of beats per tab measure. Can be overridden per measure.
TAB_BEATS_DEFAULT = 8
TAB_BEATS_OPTIONS = [8, 16, 32, 64]   # choices in the tab-grid toolbar

# Shared fonts — Courier New keeps chart/tab columns aligned on all platforms
FONT_MAIN  = ("Courier New", 11)
FONT_HEAD  = ("Courier New", 13, "bold")
FONT_TINY  = ("Courier New", 9)
FONT_MONO  = ("Courier New", 10)   # inside measure cells
FONT_CHART = ("Courier New", 11)   # song-map row rendering
FONT_STAGE = ("Courier New", 22, "bold")  # Stage View

THEMES = {
    "dark": {
        "bg":        "#1e1e2e",   # main window background
        "fg":        "#cdd6f4",   # default text
        "card_bg":   "#16213e",   # riff chips, row cards
        "header_bg": "#181825",   # top bar rows
        "input_bg":  "#0f3460",   # Entry fill
        "border":    "#45475a",   # separators, card outline
        "surface":   "#313244",   # scrollbar track
        "btn_bg":    "#313244",   # Normal button background
        "select_bg": "#89b4fa",   # selection / hover highlight / focus ring
        "select_fg": "#1e1e2e",   # text on highlight
        "lbl_gray":  "#6c7086",   # secondary / dim labels
        "accent":    "#e94560",   # app title, section header
        "green":     "#a6e3a1",
        "red":       "#f38ba8",
        "blue":      "#89b4fa",
        "mauve":     "#cba6f7",
    },
    "light": {
        "bg":        "#eff1f5",
        "fg":        "#4c4f69",
        "card_bg":   "#dce0e8",
        "header_bg": "#dce0e8",
        "input_bg":  "#e6e9ef",
        "border":    "#bcc0cc",
        "surface":   "#ccd0da",
        "btn_bg":    "#ccd0da",
        "select_bg": "#1e66f5",
        "select_fg": "#ffffff",
        "lbl_gray":  "#8c8fa1",
        "accent":    "#d20f39",
        "green":     "#40a02b",
        "red":       "#d20f39",
        "blue":      "#1e66f5",
        "mauve":     "#8839ef",
    },
}

# ==============================================================================
#  TOOLTIP  (ported from QLC+ Swiss Knife v0.4)
# ==============================================================================

class ToolTip:
    """Lightweight hover tooltip — yellow label that appears below any widget."""

    def __init__(self, widget, text):
        self.widget = widget; self.text = text; self.tip_window = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, event=None):
        if self.tip_window:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, justify="left",
                 background="#f9e2af", foreground="#1e1e2e",
                 relief="solid", borderwidth=1,
                 font=("Helvetica", 9), padx=8, pady=4,
                 wraplength=360).pack()

    def _hide(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

# ==============================================================================
#  TAB-GRID HELPERS (render:"tab"/"both" sections only — design section 4,
#  last paragraph: "The tab grid remains, unchanged from v0.15")
#  Each string row in a measure is a tk.Entry pre-filled with dash tokens.
#  Storage format is a space-separated list of tokens, e.g. "-  -  2  -".
# ==============================================================================

import re as _re

_VALID_TOKEN_CHARS = set("0123456789-/hpbx~")


def make_blank_tab_row(beats=TAB_BEATS_DEFAULT):
    """Return the default dash-filled content for one string row."""
    return "  ".join(["-"] * beats)


def parse_tab_row(text):
    """Parse a tab row string back into a list of tokens."""
    return [t for t in text.split() if t] or ["-"]


def normalise_tab_row(text, beats=TAB_BEATS_DEFAULT):
    """Re-serialise a tab row to the canonical space-separated format."""
    tokens = parse_tab_row(text)
    while len(tokens) < beats:
        tokens.append("-")
    return "  ".join(tokens)


def _validate_tab_entry(entry_widget, beats):
    """
    After a keystroke in a tab Entry, enforce that:
      1. Content splits into exactly `beats` whitespace-separated tokens.
      2. Each token is valid (digits + decorators /hpbx~-). Invalid → "-".
      3. Token count is padded to `beats` with "-", or trimmed if over.
    """
    try:
        cur    = entry_widget.get()
        cursor = entry_widget.index(tk.INSERT)
    except Exception:
        return

    raw_tokens = cur.split()
    cleaned = []
    for tok in raw_tokens:
        tok_clean = tok.strip()
        if not tok_clean:
            continue
        if _re.fullmatch(r'0\d+', tok_clean):
            cleaned.append("-")
        elif _re.fullmatch(r'\d+', tok_clean) and int(tok_clean) > transpose.MAX_FRET:
            cleaned.append("-")
        elif all(c in _VALID_TOKEN_CHARS for c in tok_clean):
            cleaned.append(tok_clean)
        else:
            cleaned.append("-")

    if len(cleaned) == beats and cleaned == raw_tokens:
        return   # nothing to fix

    while len(cleaned) < beats:
        cleaned.append("-")
    cleaned = cleaned[:beats]

    new_val = "  ".join(cleaned)
    if new_val != cur:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, new_val)
        entry_widget.icursor(min(cursor, len(new_val)))

# ==============================================================================
#  MAIN APPLICATION
# ==============================================================================

class SongNotationApp(tk.Tk):
    """
    Main application window — the v0.17 song map.

    Layout:

        ┌─ topbar row1: logo | Title | Artist | [Save][Open][Export][Theme] ─┐
        ├─ topbar row2: Key | BPM | Time | Instrument | [Transpose] [pages] ─┤
        ├─ separator ─────────────────────────────────────────────────────────┤
        │  scrollable song map — one row per section                        │
        │    Intro    │ 7 9 10 7 │ 7 9 │×2                                  │
        │    Verse 1  │ riff1 x4                                            │
        │    Bridge   │ ▼ tab · 4 bars · bass                               │
        │    Outro    │ = chorus 1  ×1 all                                  │
        ├─────────────────────────────────────────────────────────────────── ┤
        │  chorus 1 ›  riff2 x4▏                    (chart editor bar)      │
        ├─────────────────────────────────────────────────────────────────── ┤
        │  Riffs   [riff 1 · 4 bars · used ×2]  [riff 2 · 4 bars]  [+ new]  │
        └─────────────────────────────────────────────────────────────────── ┘
    """

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1180x820")
        self.minsize(900, 600)

        self.current_theme = "dark"
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        # Song-level metadata
        self.song_title  = tk.StringVar(value="Untitled Song")
        self.song_artist = tk.StringVar(value="")
        self.song_key    = tk.StringVar(value="")
        self.song_tempo  = tk.StringVar(value="")
        self.song_time   = tk.StringVar(value="4/4")
        self.default_instrument = tk.StringVar(value="Bass (4-string)")

        # ── v0.17 document model — the whole song lives in self.doc,
        # shaped exactly like model.py's schema (see DESIGN_v0_17.md §3).
        self.doc = model.new_document(time="4/4")

        # Song-map / editor-bar focus state
        self.focus_kind: str = None      # "section" | "block" | None
        self.focus_id:   str = None      # section id or block id
        self._row_frames:   dict = {}    # section_id -> row Frame
        self._row_bodies:   dict = {}    # section_id -> body Label
        self._row_gutters:  dict = {}    # section_id -> gutter Label
        self._chip_frames:  dict = {}    # block_id   -> chip Frame
        self._expanded_tab: set  = set() # section ids currently expanded
        self._tab_panels:   dict = {}    # section_id -> expanded grid Frame
        self._tab_entries:  dict = {}    # section_id -> {(m_idx,string): Entry}

        self._parse_after_id = None
        self._drag_source_idx = None
        self._rename_entry = None

        self.dirty = False
        self._next_ids = {"section": 1, "block": 1}

        self._build_ui()
        self.apply_theme()
        self._bind_shortcuts()

        # Start with one empty section so the editor bar has somewhere to go.
        self._add_section(focus=True)
        self._rebuild_map()
        self._refresh_riff_strip()
        self._refresh_page_indicator()

    # ==========================================================================
    #  ID GENERATION
    # ==========================================================================

    def _new_id(self, kind, stem):
        while True:
            n = self._next_ids[kind]
            self._next_ids[kind] += 1
            cand = f"{stem}{n}"
            existing = ({s["id"] for s in self.doc["sections"]} if kind == "section"
                        else set(self.doc["blocks"].keys()))
            if cand not in existing:
                return cand

    # ==========================================================================
    #  UI CONSTRUCTION
    # ==========================================================================

    def _build_ui(self):
        self._topbar_entries = []
        self._topbar_labels  = []

        # ── Row 1: logo | Title | Artist | action buttons ─────────────────
        self.topbar = tk.Frame(self, pady=5)
        self.topbar.pack(fill="x", padx=10)

        self.lbl_app_title = tk.Label(
            self.topbar, text="♩ SONG NOTATION TOOL",
            font=("Courier New", 14, "bold"))
        self.lbl_app_title.pack(side="left", padx=(0, 16))

        for label, var, w in [
            ("Artist:", self.song_artist, 16),
            ("Title:",  self.song_title,  20),
        ]:
            lbl = tk.Label(self.topbar, text=label, font=FONT_TINY)
            lbl.pack(side="left")
            self._topbar_labels.append(lbl)
            e = tk.Entry(self.topbar, textvariable=var, width=w,
                         relief="flat", font=FONT_MAIN)
            e.pack(side="left", padx=(2, 10))
            self._topbar_entries.append(e)

        self.btn_theme = ttk.Button(self.topbar, text="🌗 Theme",
                                     command=self.toggle_theme,
                                     style="Normal.TButton")
        self.btn_theme.pack(side="right", padx=3)
        ToolTip(self.btn_theme, "Toggle light / dark theme")

        def _show_hamburger_menu():
            m = tk.Menu(self, tearoff=0)
            t = THEMES[self.current_theme]
            m.configure(bg=t["card_bg"], fg=t["fg"],
                        activebackground=t["select_bg"],
                        activeforeground=t["select_fg"],
                        font=FONT_TINY)
            m.add_command(label="💾  Save project",    command=self._save)
            m.add_command(label="📂  Open project",    command=self._open)
            m.add_separator()
            m.add_command(label="📋  Export PDF",      command=self._export_pdf)
            m.add_command(label="📄  Export TXT",      command=self._export_txt)
            m.add_command(label="🎤  Stage View",      command=self._open_stage_view)
            m.add_separator()
            lbl = ("🌙  Switch to Light" if self.current_theme == "dark"
                   else "☀️  Switch to Dark")
            m.add_command(label=lbl,                   command=self.toggle_theme)
            bx = self.btn_menu.winfo_rootx()
            by = self.btn_menu.winfo_rooty() + self.btn_menu.winfo_height()
            m.post(bx, by)

        self.btn_menu = ttk.Button(self.topbar, text="≡",
                                    command=_show_hamburger_menu,
                                    style="Accent.TButton", width=2)
        self.btn_menu.pack(side="right", padx=(0, 3))
        ToolTip(self.btn_menu, "Menu: Save / Open / Export / Stage View / Theme")

        for text, cmd, tip in [
            ("📄 Export TXT", self._export_txt, "Export as plain-text  (Cmd/Ctrl+E)"),
            ("📋 Export PDF", self._export_pdf, "Export as PDF  (Cmd/Ctrl+Shift+P)"),
            ("📂 Open",       self._open,       "Open a .sng project file"),
            ("💾 Save",       self._save,       "Save project  (Cmd/Ctrl+S)"),
        ]:
            b = ttk.Button(self.topbar, text=text, command=cmd, style="Accent.TButton")
            b.pack(side="right", padx=3)
            ToolTip(b, tip)

        # ── Row 2: Key | BPM | Time | Instrument | Transpose | pages ──────
        self.topbar2 = tk.Frame(self, pady=4)
        self.topbar2.pack(fill="x", padx=10)

        for label, var, w, tip in [
            ("Key:",  self.song_key,   5, "Key signature (e.g. Am, G, Bb)"),
            ("BPM:",  self.song_tempo, 5, "Tempo in beats per minute"),
            ("Time:", self.song_time,  5, "Time signature (e.g. 4/4, 3/4, 6/8)"),
        ]:
            lbl = tk.Label(self.topbar2, text=label, font=FONT_TINY)
            lbl.pack(side="left")
            self._topbar_labels.append(lbl)
            e = tk.Entry(self.topbar2, textvariable=var, width=w,
                         relief="flat", font=FONT_MAIN)
            e.pack(side="left", padx=(2, 12))
            self._topbar_entries.append(e)
            ToolTip(e, tip)

        lbl_i = tk.Label(self.topbar2, text="Default instrument:", font=FONT_TINY)
        lbl_i.pack(side="left")
        self._topbar_labels.append(lbl_i)
        self.instr_cb = ttk.Combobox(
            self.topbar2, textvariable=self.default_instrument,
            values=list(INSTRUMENT_STRINGS.keys()),
            width=20, state="readonly", font=FONT_MAIN)
        self.instr_cb.pack(side="left", padx=(2, 12))
        ToolTip(self.instr_cb, "Instrument used for new sections")

        self.btn_transpose = ttk.Button(
            self.topbar2, text="Transpose ↕",
            command=self._transpose_dialog, style="Normal.TButton")
        self.btn_transpose.pack(side="left", padx=3)
        ToolTip(self.btn_transpose,
                "Shift the whole song, or just the focused section, by ±N semitones "
                "— applied at render time only (never rewrites stored notes).")

        self.lbl_pages = tk.Label(self.topbar2, text="[1 page]",
                                   font=FONT_TINY, anchor="e")
        self.lbl_pages.pack(side="right", padx=(0, 4))
        self._topbar_labels.append(self.lbl_pages)

        # ── Separator ───────────────────────────────────────────────────
        self.sep = tk.Frame(self, height=2)
        self.sep.pack(fill="x")

        # ── Bottom-up: riff strip, then editor bar, then the map fills
        # whatever is left (design section 3: map / editor bar / riffs).
        self._build_riff_strip()
        self._build_editor_bar()
        self._build_song_map()

    def _build_song_map(self):
        outer = tk.Frame(self)
        outer.pack(fill="both", expand=True, padx=8, pady=(6, 4))
        self.map_outer = outer

        self.map_canvas = tk.Canvas(outer, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=self.map_canvas.yview)
        self.map_canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.map_canvas.pack(side="left", fill="both", expand=True)

        self.map_frame = tk.Frame(self.map_canvas)
        self.map_window = self.map_canvas.create_window((0, 0), window=self.map_frame, anchor="nw")
        self.map_frame.bind(
            "<Configure>",
            lambda e: self.map_canvas.configure(scrollregion=self.map_canvas.bbox("all")))
        self.map_canvas.bind(
            "<Configure>",
            lambda e: self.map_canvas.itemconfig(self.map_window, width=e.width))

        def _wheel(event):
            delta = -1 * int(event.delta / abs(event.delta)) if event.delta else 0
            self.map_canvas.yview_scroll(delta, "units")

        for w in (self.map_canvas, self.map_frame):
            w.bind("<MouseWheel>", _wheel)
            w.bind("<Button-4>", lambda e: self.map_canvas.yview_scroll(-1, "units"))
            w.bind("<Button-5>", lambda e: self.map_canvas.yview_scroll(1, "units"))

    def _build_editor_bar(self):
        frame = tk.Frame(self)
        frame.pack(side="bottom", fill="x", padx=8, pady=(4, 2))
        self.editor_bar_frame = frame

        self.lbl_editor_hint = tk.Label(
            frame, text="Click a section below to edit its chart line",
            font=FONT_TINY, anchor="w")
        self.lbl_editor_hint.pack(fill="x")

        self.editor_entry = tk.Entry(frame, font=FONT_CHART, relief="flat")
        self.editor_entry.pack(fill="x", ipady=4)
        self.editor_entry.bind("<KeyRelease>", self._on_editor_keystroke)
        self.editor_entry.bind("<Tab>",       lambda e: self._on_editor_tab(1))
        self.editor_entry.bind("<Shift-Tab>", lambda e: self._on_editor_tab(-1))
        self.editor_entry.bind("<Return>",    self._on_editor_return)
        self.editor_entry.bind("<Escape>",    self._on_editor_escape)
        self.editor_entry.bind("<Control-r>", self._promote_selection_to_riff)
        self.editor_entry.bind("<Command-r>", self._promote_selection_to_riff)

        self.lbl_editor_error = tk.Label(frame, text="", font=FONT_TINY, anchor="w")
        self.lbl_editor_error.pack(fill="x")
        ToolTip(self.editor_entry,
                "Type a chart line: fret+note tokens, [groups]xN, riff refs, "
                "=section refs, marks. Tab commits and moves on; Escape reverts; "
                "Ctrl/Cmd+R promotes the selection to a riff.")

    def _build_riff_strip(self):
        outer = tk.Frame(self)
        outer.pack(side="bottom", fill="x", padx=8, pady=(0, 6))
        self.riff_strip_frame = outer

        hdr = tk.Frame(outer)
        hdr.pack(fill="x")
        self.lbl_riffs = tk.Label(hdr, text="RIFFS", font=FONT_TINY, anchor="w")
        self.lbl_riffs.pack(side="left")
        btn_new = ttk.Button(hdr, text="+ new", command=self._new_riff_dialog,
                              style="Normal.TButton")
        btn_new.pack(side="right")
        ToolTip(btn_new, "Add an empty riff to the library")

        chips_outer = tk.Frame(outer)
        chips_outer.pack(fill="x")
        self.riff_canvas = tk.Canvas(chips_outer, height=46, highlightthickness=0)
        hsb = ttk.Scrollbar(chips_outer, orient="horizontal", command=self.riff_canvas.xview)
        self.riff_canvas.configure(xscrollcommand=hsb.set)
        self.riff_canvas.pack(side="top", fill="x")
        hsb.pack(side="bottom", fill="x")
        self.riff_chip_frame = tk.Frame(self.riff_canvas)
        self.riff_window = self.riff_canvas.create_window(
            (0, 0), window=self.riff_chip_frame, anchor="nw")
        self.riff_chip_frame.bind(
            "<Configure>",
            lambda e: self.riff_canvas.configure(scrollregion=self.riff_canvas.bbox("all")))

    # ==========================================================================
    #  SONG MAP — rows
    # ==========================================================================

    def _rebuild_map(self):
        for child in self.map_frame.winfo_children():
            child.destroy()
        self._row_frames.clear()
        self._row_bodies.clear()
        self._row_gutters.clear()
        self._tab_panels.clear()

        for sec in self.doc["sections"]:
            self._build_section_row(self.map_frame, sec)

        add_row = tk.Frame(self.map_frame)
        add_row.pack(fill="x", pady=(6, 16), padx=4)
        b = ttk.Button(add_row, text="+ Add Section",
                        command=lambda: self._add_section(focus=True),
                        style="Normal.TButton")
        b.pack(side="left")
        ToolTip(b, "Add a new section  (Enter, while a row is focused, does the same)")

        self._apply_row_theme_all()
        self.map_frame.update_idletasks()
        self.map_canvas.configure(scrollregion=self.map_canvas.bbox("all"))

    def _build_section_row(self, parent, sec):
        sid = sec["id"]
        row = tk.Frame(parent, bd=0)
        row.pack(fill="x", padx=4, pady=1)
        row.grid_columnconfigure(1, weight=1)
        self._row_frames[sid] = row

        gutter = tk.Label(row, text=sec["name"], font=FONT_MAIN, width=14,
                           anchor="e", cursor="fleur", justify="right")
        gutter.grid(row=0, column=0, sticky="ne", padx=(2, 8), pady=4)
        self._row_gutters[sid] = gutter
        ToolTip(gutter, "Click to edit · double-click to rename · drag to reorder")

        body = tk.Label(row, text=self._render_section_preview(sec),
                         font=FONT_CHART, justify="left", anchor="w")
        body.grid(row=0, column=1, sticky="w", pady=4)
        self._row_bodies[sid] = body

        del_btn = ttk.Button(row, text="✕", width=2, style="Normal.TButton",
                              command=lambda s=sid: self._delete_section(s))
        del_btn.grid(row=0, column=2, sticky="ne", padx=(8, 2))
        ToolTip(del_btn, "Delete section")

        for w in (gutter, body):
            w.bind("<Button-1>", lambda e, s=sid: self._on_row_click(s))
        gutter.bind("<Double-Button-1>", lambda e, s=sid: self._start_rename(s))
        gutter.bind("<ButtonPress-1>",   lambda e, s=sid: self._drag_start(s, e), add="+")
        gutter.bind("<B1-Motion>",       self._drag_motion, add="+")
        gutter.bind("<ButtonRelease-1>", self._drag_release, add="+")

        if sec.get("render") in ("tab", "both"):
            expanded = sid in self._expanded_tab
            toggle = ttk.Button(
                row, text=("▲ Hide tab grid" if expanded else "▼ Edit tab grid"),
                style="Normal.TButton",
                command=lambda s=sid: self._toggle_tab_panel(s))
            toggle.grid(row=1, column=1, sticky="w", pady=(0, 4))
            if expanded:
                self._build_tab_panel(row, sec)

    def _render_section_preview(self, sec):
        render_mode = sec.get("render", "chart")
        eff = transpose.effective_transpose(
            self.doc.get("transpose", 0), sec.get("transpose", 0))
        rep = sec.get("repeat", 1)
        rep_suffix = f"   ×{rep}" if rep and rep != 1 else ""

        lines = []
        if render_mode in ("chart", "both"):
            chart = songmap.chart_items(sec)
            if chart:
                resolved = render.resolve_display_items(chart, eff)
                lines = [ln for ln in render.render_chart_row(resolved) if ln.strip()]

        if not lines:
            if render_mode == "tab":
                n = len(songmap.measure_items(sec))
                instr = sec.get("instrument", "")
                lines = [f"{n} bar{'s' if n != 1 else ''} of tab · {instr}"]
            else:
                lines = ["(type a chart line below)"]

        lines[-1] = lines[-1] + rep_suffix
        if sec.get("annotation"):
            lines.append(f'  "{sec["annotation"]}"')
        if render.has_coda(sec.get("items", [])):
            lines.append("  ◆ coda")
        return "\n".join(lines)

    # ── focus / selection ────────────────────────────────────────────────

    def _on_row_click(self, sid):
        if self._drag_source_idx is not None:
            return
        self._focus_section(sid)

    def _find_section(self, sid):
        for s in self.doc["sections"]:
            if s["id"] == sid:
                return s
        return None

    def _section_index(self, sid):
        for i, s in enumerate(self.doc["sections"]):
            if s["id"] == sid:
                return i
        return None

    def _focus_section(self, sid):
        sec = self._find_section(sid)
        if sec is None:
            return
        self.focus_kind = "section"
        self.focus_id = sid
        annotation_suffix = f' "{sec["annotation"]}"' if sec.get("annotation") else ""
        self._set_editor_text(grammar.unparse(songmap.chart_items(sec)) + annotation_suffix)
        if sec.get("render") == "tab":
            self.editor_entry.configure(state="disabled")
            self.lbl_editor_hint.configure(
                text=f"{sec['name']} › tab-only section — edit the grid below")
        else:
            self.editor_entry.configure(state="normal")
            self.lbl_editor_hint.configure(text=f"{sec['name']} ›")
        self.lbl_editor_error.configure(text="")
        self._apply_row_theme_all()

    def _focus_block(self, bid):
        block = self.doc["blocks"].get(bid)
        if block is None:
            return
        self.focus_kind = "block"
        self.focus_id = bid
        self.editor_entry.configure(state="normal")
        self._set_editor_text(grammar.unparse(block.get("items", [])))
        self.lbl_editor_hint.configure(
            text=f"Riff ‹{block['name']}› — editing here updates every section that "
                 f"references it")
        self.lbl_editor_error.configure(text="")
        self._apply_row_theme_all()

    def _set_editor_text(self, text):
        was_disabled = self.editor_entry.cget("state") == "disabled"
        if was_disabled:
            self.editor_entry.configure(state="normal")
        self.editor_entry.delete(0, tk.END)
        self.editor_entry.insert(0, text)
        if was_disabled:
            self.editor_entry.configure(state="disabled")

    # ── rename ───────────────────────────────────────────────────────────

    def _start_rename(self, sid):
        sec = self._find_section(sid)
        gutter = self._row_gutters.get(sid)
        if sec is None or gutter is None or self._rename_entry is not None:
            return
        entry = tk.Entry(gutter.master, font=FONT_MAIN, width=14, justify="right")
        entry.insert(0, sec["name"])
        entry.select_range(0, tk.END)
        entry.grid(row=0, column=0, sticky="ne", padx=(2, 8), pady=4)
        entry.focus_set()
        self._rename_entry = entry

        def commit(event=None):
            if self._rename_entry is None:
                return
            new_name = entry.get().strip() or sec["name"]
            sec["name"] = new_name
            entry.destroy()
            self._rename_entry = None
            self.dirty = True
            was_focused = (self.focus_kind == "section" and self.focus_id == sid)
            self._rebuild_map()
            if was_focused:
                self._focus_section(sid)

        def cancel(event=None):
            entry.destroy()
            self._rename_entry = None

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>", cancel)

    # ── drag reorder (gutter) ───────────────────────────────────────────

    def _drag_start(self, sid, event):
        self._drag_source_idx = self._section_index(sid)
        self._focus_section(sid)

    def _drag_motion(self, event):
        if self._drag_source_idx is None:
            return
        y = event.widget.winfo_rooty() + event.y - self.map_frame.winfo_rooty()
        target_idx = self._row_index_at_y(y)
        if target_idx is not None and target_idx != self._drag_source_idx:
            songmap.reorder_sections(self.doc, self._drag_source_idx, target_idx)
            self._drag_source_idx = target_idx
            self.dirty = True
            focus_id = self.focus_id
            self._rebuild_map()
            if focus_id:
                self._apply_row_theme_all()

    def _drag_release(self, event):
        self._drag_source_idx = None
        self._refresh_page_indicator()

    def _row_index_at_y(self, y):
        best_idx, best_dist = None, None
        for idx, sec in enumerate(self.doc["sections"]):
            frame = self._row_frames.get(sec["id"])
            if frame is None:
                continue
            mid = frame.winfo_y() + frame.winfo_height() / 2
            dist = abs(mid - y)
            if best_dist is None or dist < best_dist:
                best_dist, best_idx = dist, idx
        return best_idx

    # ── add / delete section ────────────────────────────────────────────

    def _add_section(self, focus=False, after_id=None):
        sid = self._new_id("section", "sec")
        n = len(self.doc["sections"]) + 1
        sec = model.new_section(sid, f"Section {n}",
                                 instrument=self.default_instrument.get())
        if after_id is not None:
            idx = self._section_index(after_id)
            pos = (idx + 1) if idx is not None else len(self.doc["sections"])
            self.doc["sections"].insert(pos, sec)
        else:
            self.doc["sections"].append(sec)
        self.dirty = True
        if focus:
            self._rebuild_map()
            self._focus_section(sid)
            self._refresh_page_indicator()
        return sec

    def _delete_section(self, sid):
        sec = self._find_section(sid)
        if sec is None:
            return
        refs = songmap.section_refs_to(self.doc, sid)
        if refs and not messagebox.askyesno(
                "Delete section",
                f"{', '.join(refs)} reference this section with '='. "
                f"Delete '{sec['name']}' anyway?"):
            return
        if not refs and not messagebox.askyesno(
                "Delete section", f"Delete '{sec['name']}'?"):
            return
        self.doc["sections"] = [s for s in self.doc["sections"] if s["id"] != sid]
        if self.focus_id == sid:
            self.focus_kind = self.focus_id = None
            self._set_editor_text("")
            self.lbl_editor_hint.configure(text="Click a section below to edit its chart line")
        self.dirty = True
        self._rebuild_map()
        self._refresh_page_indicator()

    # ==========================================================================
    #  CHART EDITOR BAR — live parse, commit, revert  (design section 4)
    # ==========================================================================

    def _on_editor_keystroke(self, event=None):
        if event is not None and event.keysym in ("Tab", "Return", "Escape"):
            return
        if self._parse_after_id:
            self.after_cancel(self._parse_after_id)
        self._parse_after_id = self.after(150, self._commit_editor_line)

    def _commit_editor_line(self, force=False):
        if self._parse_after_id:
            self.after_cancel(self._parse_after_id)
            self._parse_after_id = None
        if self.focus_kind is None:
            return
        text = self.editor_entry.get()
        try:
            items, annotation = grammar.parse(text)
        except ChartParseError as e:
            # A half-typed token never clears the rendered row (section 4,
            # acceptance criteria): leave the last valid data untouched and
            # just surface the error.
            self.lbl_editor_error.configure(text=f"⚠ {e}")
            return

        self.lbl_editor_error.configure(text="")

        if self.focus_kind == "section":
            sec = self._find_section(self.focus_id)
            if sec is None:
                return
            songmap.set_chart_items(sec, items)
            sec["annotation"] = annotation or ""
            body = self._row_bodies.get(self.focus_id)
            if body is not None:
                body.configure(text=self._render_section_preview(sec))
        elif self.focus_kind == "block":
            block = self.doc["blocks"].get(self.focus_id)
            if block is None:
                return
            block["items"] = items
            self._refresh_riff_strip()

        self.dirty = True
        self._refresh_page_indicator()

    def _on_editor_tab(self, direction):
        self._commit_editor_line(force=True)
        if self.focus_kind == "section":
            idx = self._section_index(self.focus_id)
            n = len(self.doc["sections"])
            if idx is not None and n:
                self._focus_section(self.doc["sections"][(idx + direction) % n]["id"])
        return "break"

    def _on_editor_return(self, event=None):
        self._commit_editor_line(force=True)
        if self.focus_kind == "section":
            new_sec = self._add_section(focus=False, after_id=self.focus_id)
            self._rebuild_map()
            self._focus_section(new_sec["id"])
            self._refresh_page_indicator()
        return "break"

    def _on_editor_escape(self, event=None):
        if self._parse_after_id:
            self.after_cancel(self._parse_after_id)
            self._parse_after_id = None
        self.lbl_editor_error.configure(text="")
        if self.focus_kind == "section":
            self._focus_section(self.focus_id)
        elif self.focus_kind == "block":
            self._focus_block(self.focus_id)
        return "break"

    def _promote_selection_to_riff(self, event=None):
        if self.focus_kind != "section":
            return "break"
        try:
            sel_start = self.editor_entry.index("sel.first")
            sel_end = self.editor_entry.index("sel.last")
        except tk.TclError:
            messagebox.showinfo("Promote to riff",
                                 "Select the part of the chart line to promote first.")
            return "break"
        line = self.editor_entry.get()
        name = simpledialog.askstring("Promote to riff", "Name for the new riff:",
                                       parent=self)
        if not name:
            return "break"
        block_id = songmap.unique_block_id(self.doc, stem=_slug(name))
        try:
            new_line, riff_items = songmap.promote_to_riff(line, sel_start, sel_end, block_id)
        except (ValueError, ChartParseError) as e:
            messagebox.showerror("Promote to riff", f"Can't promote this selection:\n{e}")
            return "break"
        sec = self._find_section(self.focus_id)
        block = model.new_block(
            block_id, name, bars=max(1, len(riff_items) // 4 or 1),
            beats_per_bar=TAB_BEATS_DEFAULT,
            instrument=sec.get("instrument", self.default_instrument.get()) if sec else
            self.default_instrument.get())
        block["items"] = riff_items
        self.doc["blocks"][block_id] = block
        self._set_editor_text(new_line)
        self._commit_editor_line(force=True)
        self._refresh_riff_strip()
        return "break"

    def _duplicate_as_reference(self, event=None):
        if self.focus_kind != "section":
            return "break"
        idx = self._section_index(self.focus_id)
        if idx is None:
            return "break"
        self._commit_editor_line(force=True)
        new_id = self._new_id("section", "sec")
        src = self.doc["sections"][idx]
        new_name = src["name"] + " (repeat)"
        songmap.duplicate_as_reference(self.doc, idx, new_id, new_name)
        self.dirty = True
        self._rebuild_map()
        self._focus_section(new_id)
        self._refresh_page_indicator()
        return "break"

    def _move_focused_section(self, delta):
        if self.focus_kind != "section":
            return "break"
        idx = self._section_index(self.focus_id)
        if idx is None:
            return "break"
        self._commit_editor_line(force=True)
        songmap.move_section(self.doc, idx, delta)
        self.dirty = True
        sid = self.focus_id
        self._rebuild_map()
        self._focus_section(sid)
        self._refresh_page_indicator()
        return "break"

    # ==========================================================================
    #  RIFF LIBRARY  (design section 5)
    # ==========================================================================

    def _refresh_riff_strip(self):
        for child in self.riff_chip_frame.winfo_children():
            child.destroy()
        self._chip_frames.clear()
        usage = songmap.block_usage(self.doc)
        for bid, block in self.doc["blocks"].items():
            used = len(usage.get(bid, []))
            chip = tk.Frame(self.riff_chip_frame, relief="solid", bd=1)
            chip.pack(side="left", padx=4, pady=4)
            label_text = f"{block['name']} · {block.get('bars', 1)} bars"
            if used:
                label_text += f" · used ×{used}"
            lbl = tk.Label(chip, text=label_text, font=FONT_TINY, justify="left")
            lbl.pack(side="left", padx=6, pady=6)
            for w in (chip, lbl):
                w.bind("<Button-1>", lambda e, b=bid: self._focus_block(b))
            del_btn = ttk.Button(chip, text="✕", width=2, style="Normal.TButton",
                                  command=lambda b=bid: self._delete_riff(b))
            del_btn.pack(side="left", padx=(0, 4))
            ToolTip(chip, f"{used} section{'s' if used != 1 else ''} reference this riff"
                          if used else "Not referenced by any section yet")
            self._chip_frames[bid] = chip
        self._apply_row_theme_all()

    def _new_riff_dialog(self):
        name = simpledialog.askstring("New riff", "Riff name:", parent=self)
        if not name:
            return
        bid = songmap.unique_block_id(self.doc, stem=_slug(name))
        block = model.new_block(bid, name, bars=4, beats_per_bar=TAB_BEATS_DEFAULT,
                                 instrument=self.default_instrument.get())
        self.doc["blocks"][bid] = block
        self.dirty = True
        self._refresh_riff_strip()
        self._focus_block(bid)

    def _delete_riff(self, bid):
        block = self.doc["blocks"].get(bid)
        if block is None:
            return
        ok, refs = songmap.can_delete_block(self.doc, bid)
        if not ok:
            messagebox.showwarning(
                "Riff in use",
                f"'{block['name']}' is referenced by: {', '.join(refs)}.\n"
                "Remove those references before deleting it.")
            return
        if not messagebox.askyesno("Delete riff", f"Delete '{block['name']}'?"):
            return
        del self.doc["blocks"][bid]
        if self.focus_kind == "block" and self.focus_id == bid:
            self.focus_kind = self.focus_id = None
            self._set_editor_text("")
            self.lbl_editor_hint.configure(text="Click a section below to edit its chart line")
        self.dirty = True
        self._refresh_riff_strip()

    # ==========================================================================
    #  TAB GRID  (render:"tab"/"both" sections only — design section 4)
    # ==========================================================================

    def _toggle_tab_panel(self, sid):
        if sid in self._expanded_tab:
            self._expanded_tab.discard(sid)
        else:
            self._expanded_tab.add(sid)
        was_focused = (self.focus_kind == "section" and self.focus_id == sid)
        self._rebuild_map()
        if was_focused:
            self._focus_section(sid)

    def _build_tab_panel(self, row, sec):
        sid = sec["id"]
        panel = tk.Frame(row, relief="groove", bd=1)
        panel.grid(row=2, column=0, columnspan=3, sticky="ew", padx=(2, 2), pady=(0, 6))
        self._tab_panels[sid] = panel

        toolbar = tk.Frame(panel)
        toolbar.pack(fill="x", padx=4, pady=(4, 2))
        tk.Label(toolbar, text=f"Instrument: {sec.get('instrument', '')}",
                 font=FONT_TINY).pack(side="left")
        ttk.Button(toolbar, text="+ bar", style="Normal.TButton",
                   command=lambda s=sid: self._add_measure(s)).pack(side="right", padx=2)
        ttk.Button(toolbar, text="- bar", style="Normal.TButton",
                   command=lambda s=sid: self._remove_measure(s)).pack(side="right", padx=2)

        grid_frame = tk.Frame(panel)
        grid_frame.pack(fill="x", padx=4, pady=(0, 6))

        measures = songmap.measure_items(sec)
        strings = INSTRUMENT_STRINGS.get(sec.get("instrument"), ["G", "D", "A", "E"])
        self._tab_entries[sid] = {}

        tk.Label(grid_frame, text="", font=FONT_TINY, width=3).grid(row=0, column=0)
        for m_idx in range(len(measures)):
            tk.Label(grid_frame, text=f"M{m_idx + 1}", font=FONT_TINY).grid(
                row=0, column=m_idx + 1, padx=2)

        for r_idx, st in enumerate(strings):
            tk.Label(grid_frame, text=st, font=FONT_MONO, width=3).grid(
                row=r_idx + 1, column=0)
            for m_idx, measure in enumerate(measures):
                beats = measure.get("beats", TAB_BEATS_DEFAULT)
                val = measure.get("strings", {}).get(st, make_blank_tab_row(beats))
                e = tk.Entry(grid_frame, font=FONT_MONO, width=max(10, beats * 3))
                e.insert(0, val)
                e.grid(row=r_idx + 1, column=m_idx + 1, padx=2, pady=1, sticky="w")
                e.bind("<KeyRelease>", lambda ev, w=e, b=beats: _validate_tab_entry(w, b))
                e.bind("<FocusOut>", lambda ev, s=sid, m=m_idx, st=st, w=e, b=beats:
                       self._commit_tab_cell(s, m, st, w, b))
                self._tab_entries[sid][(m_idx, st)] = e

        if not measures:
            tk.Label(grid_frame, text="No bars yet — click + bar to add one.",
                     font=FONT_TINY).grid(row=1, column=0, columnspan=2, sticky="w")

    def _commit_tab_cell(self, sid, m_idx, string, entry_widget, beats):
        sec = self._find_section(sid)
        if sec is None:
            return
        measures = songmap.measure_items(sec)
        if m_idx >= len(measures):
            return
        val = normalise_tab_row(entry_widget.get(), beats)
        measures[m_idx].setdefault("strings", {})[string] = val
        songmap.set_measure_items(sec, measures)
        self.dirty = True
        body = self._row_bodies.get(sid)
        if body is not None:
            body.configure(text=self._render_section_preview(sec))
        self._refresh_page_indicator()

    def _add_measure(self, sid):
        sec = self._find_section(sid)
        if sec is None:
            return
        strings = INSTRUMENT_STRINGS.get(sec.get("instrument"), ["G", "D", "A", "E"])
        measures = songmap.measure_items(sec)
        measures.append(model.make_measure(
            TAB_BEATS_DEFAULT,
            {st: make_blank_tab_row(TAB_BEATS_DEFAULT) for st in strings}))
        songmap.set_measure_items(sec, measures)
        self.dirty = True
        self._rebuild_map()

    def _remove_measure(self, sid):
        sec = self._find_section(sid)
        if sec is None:
            return
        measures = songmap.measure_items(sec)
        if not measures:
            return
        measures.pop()
        songmap.set_measure_items(sec, measures)
        self.dirty = True
        self._rebuild_map()

    # ==========================================================================
    #  PAGE INDICATOR
    # ==========================================================================

    def _refresh_page_indicator(self):
        n = render.estimate_page_count(self.doc, instrument_strings=INSTRUMENT_STRINGS)
        self.lbl_pages.configure(text=f"[{n} page{'s' if n != 1 else ''}]")

    # ==========================================================================
    #  THEME ENGINE  (ported from QLC+ Swiss Knife v0.4)
    # ==========================================================================

    def toggle_theme(self):
        self.current_theme = "light" if self.current_theme == "dark" else "dark"
        self.apply_theme()

    def apply_theme(self):
        t = THEMES[self.current_theme]
        self.configure(bg=t["bg"])

        s = self.style
        s.configure(".", background=t["bg"], foreground=t["fg"])

        s.configure("Normal.TButton", background=t["btn_bg"], foreground=t["fg"],
                    font=FONT_TINY, padding=4)
        s.map("Normal.TButton",
              background=[("active", t["select_bg"]), ("pressed", t["select_bg"])],
              foreground=[("!disabled", t["fg"]),
                           ("active", t["select_fg"]), ("pressed", t["select_fg"])])

        s.configure("Accent.TButton", background=t["accent"], foreground="#ffffff",
                    font=("Courier New", 9, "bold"), padding=4)
        s.map("Accent.TButton",
              background=[("active", t["select_bg"]), ("pressed", t["select_bg"])],
              foreground=[("!disabled", "#ffffff"),
                           ("active", t["select_fg"]), ("pressed", t["select_fg"])])

        s.configure("TCombobox",
                    fieldbackground=t["input_bg"], background=t["btn_bg"],
                    foreground=t["fg"], arrowcolor=t["fg"],
                    selectbackground=t["select_bg"], selectforeground=t["select_fg"])
        s.map("TCombobox",
              fieldbackground=[("readonly", t["input_bg"])],
              foreground=[("readonly", t["fg"]), ("!disabled", t["fg"])],
              selectbackground=[("readonly", t["select_bg"])],
              selectforeground=[("readonly", t["select_fg"])])
        self.option_add("*TCombobox*Listbox.background",       t["input_bg"])
        self.option_add("*TCombobox*Listbox.foreground",       t["fg"])
        self.option_add("*TCombobox*Listbox.selectBackground", t["select_bg"])
        self.option_add("*TCombobox*Listbox.selectForeground", t["select_fg"])

        s.configure("TScrollbar", background=t["select_bg"], troughcolor=t["surface"],
                    arrowcolor=t["fg"], borderwidth=0, relief="flat")

        self.topbar.configure(bg=t["header_bg"])
        self.topbar2.configure(bg=t["header_bg"])
        self.sep.configure(bg=t["border"])
        self.lbl_app_title.configure(bg=t["header_bg"], fg=t["accent"])
        for lbl in self._topbar_labels:
            lbl.configure(bg=t["header_bg"], fg=t["lbl_gray"])
        for e in self._topbar_entries:
            e.configure(bg=t["input_bg"], fg=t["fg"], insertbackground=t["fg"])

        for frame in (self.map_outer, self.map_canvas, self.map_frame,
                      self.editor_bar_frame, self.riff_strip_frame,
                      self.riff_canvas, self.riff_chip_frame):
            try:
                frame.configure(bg=t["bg"])
            except tk.TclError:
                pass

        self.lbl_editor_hint.configure(bg=t["bg"], fg=t["lbl_gray"])
        self.lbl_editor_error.configure(bg=t["bg"], fg=t["red"])
        self.editor_entry.configure(bg=t["input_bg"], fg=t["fg"],
                                    insertbackground=t["fg"],
                                    disabledbackground=t["surface"],
                                    disabledforeground=t["lbl_gray"])
        self.lbl_riffs.configure(bg=t["bg"], fg=t["lbl_gray"])

        self._apply_row_theme_all()

    def _apply_row_theme_all(self):
        t = THEMES[self.current_theme]
        for sid, gutter in self._row_gutters.items():
            body = self._row_bodies.get(sid)
            row = self._row_frames.get(sid)
            is_focused = (self.focus_kind == "section" and sid == self.focus_id)
            bg = t["select_bg"] if is_focused else t["bg"]
            fg = t["select_fg"] if is_focused else t["fg"]
            try:
                gutter.configure(bg=bg, fg=fg)
                if body is not None:
                    body.configure(bg=bg, fg=fg)
                if row is not None:
                    row.configure(bg=t["bg"])
            except tk.TclError:
                pass
        for bid, chip in self._chip_frames.items():
            is_focused = (self.focus_kind == "block" and bid == self.focus_id)
            bg = t["select_bg"] if is_focused else t["card_bg"]
            fg = t["select_fg"] if is_focused else t["fg"]
            try:
                chip.configure(bg=bg)
                for child in chip.winfo_children():
                    if isinstance(child, tk.Label):
                        child.configure(bg=bg, fg=fg)
            except tk.TclError:
                pass

    # ==========================================================================
    #  TRANSPOSE  — non-destructive: sets doc/section "transpose", never
    #  rewrites a stored token (design v0.16 section 5 / v0.15 bug fix #5).
    # ==========================================================================

    def _transpose_dialog(self):
        self._commit_editor_line(force=True)
        t = THEMES[self.current_theme]
        dlg = tk.Toplevel(self)
        dlg.title("Transpose")
        dlg.configure(bg=t["bg"])
        dlg.resizable(False, False)
        dlg.grab_set()

        has_focused_section = (self.focus_kind == "section")
        scope_var = tk.StringVar(value="section" if has_focused_section else "song")

        tk.Label(dlg, text="Transpose:", bg=t["bg"], fg=t["accent"],
                 font=FONT_TINY).pack(anchor="w", padx=16, pady=(12, 2))
        tk.Radiobutton(dlg, text="Whole song", variable=scope_var, value="song",
                       bg=t["bg"], fg=t["fg"], selectcolor=t["input_bg"],
                       activebackground=t["bg"], font=FONT_TINY).pack(anchor="w", padx=28)
        tk.Radiobutton(dlg, text="Focused section only", variable=scope_var,
                       value="section", bg=t["bg"], fg=t["fg"],
                       selectcolor=t["input_bg"], activebackground=t["bg"],
                       font=FONT_TINY,
                       state=("normal" if has_focused_section else "disabled")
                       ).pack(anchor="w", padx=28)

        row = tk.Frame(dlg, bg=t["bg"])
        row.pack(padx=16, pady=(10, 4))
        tk.Label(row, text="Semitones (+/-):", bg=t["bg"], fg=t["fg"],
                 font=FONT_TINY).pack(side="left")
        amount_var = tk.StringVar(value="0")
        tk.Entry(row, textvariable=amount_var, width=5, font=FONT_MAIN).pack(
            side="left", padx=6)

        def apply_transpose():
            try:
                n = int(amount_var.get())
            except ValueError:
                messagebox.showerror("Transpose", "Enter a whole number of semitones.")
                return
            if scope_var.get() == "song":
                self.doc["transpose"] = self.doc.get("transpose", 0) + n
            else:
                sec = self._find_section(self.focus_id)
                if sec is not None:
                    sec["transpose"] = sec.get("transpose", 0) + n
            self.dirty = True
            dlg.destroy()
            self._rebuild_map()

        br = tk.Frame(dlg, bg=t["bg"])
        br.pack(fill="x", padx=16, pady=(10, 14))
        ttk.Button(br, text="Cancel", command=dlg.destroy,
                   style="Normal.TButton").pack(side="right", padx=4)
        ttk.Button(br, text="Apply", command=apply_transpose,
                   style="Accent.TButton").pack(side="right")
        dlg.bind("<Return>", lambda e: apply_transpose())

    # ==========================================================================
    #  EXPORT — TXT / PDF, both built from the same section-line renderer as
    #  the song map and the Stage View (design section 7).
    # ==========================================================================

    def _build_song_lines(self, instruments=None):
        W = 100
        div = lambda c="=": c * W
        lines = []

        artist = self.song_artist.get().strip()
        title  = self.song_title.get().strip()
        lines += [div("="), f"  {title.upper()}"]
        if artist:
            lines.append(f"  {artist}")
        meta = []
        if self.song_key.get():   meta.append(f"Key: {self.song_key.get()}")
        if self.song_tempo.get(): meta.append(f"BPM: {self.song_tempo.get()}")
        if self.song_time.get():  meta.append(f"Time: {self.song_time.get()}")
        if meta:
            lines.append("  " + "   ".join(meta))
        lines += [div("="), ""]

        TOKEN_W = 3
        for sec in self.doc["sections"]:
            if instruments is not None and sec.get("instrument") not in instruments:
                continue

            rep = sec.get("repeat", 1)
            rep_str = f"  (x{rep})" if rep and rep != 1 else ""
            lines += [div("-"),
                      f"  [{sec['name']}]{rep_str}   {sec.get('instrument', '')}",
                      div("-"), ""]

            eff = transpose.effective_transpose(
                self.doc.get("transpose", 0), sec.get("transpose", 0))
            chart = songmap.chart_items(sec)
            if chart:
                resolved = render.resolve_display_items(chart, eff)
                lines += render.render_chart_row(resolved)
                lines.append("")

            measures = songmap.measure_items(sec)
            if measures:
                all_strings = INSTRUMENT_STRINGS.get(
                    sec.get("instrument"), ["e", "B", "G", "D", "A", "E"])
                strings = render.active_strings(
                    [m.get("strings", {}) for m in measures], all_strings) or all_strings
                max_beats = max((m.get("beats", TAB_BEATS_DEFAULT) for m in measures),
                                 default=TAB_BEATS_DEFAULT)
                mpl = render.measures_per_line(len(measures), max_beats, width=W)

                for chunk_start in range(0, len(measures), mpl):
                    chunk = list(range(chunk_start, min(chunk_start + mpl, len(measures))))
                    sn_pad = max(len(st) for st in strings) + 3
                    hdr = " " * sn_pad
                    for m_idx in chunk:
                        beats = measures[m_idx].get("beats", TAB_BEATS_DEFAULT)
                        cell_w = TOKEN_W * beats + 1
                        hdr += f"{'M' + str(m_idx + 1):<{cell_w}}"
                    lines.append(hdr)
                    for st in strings:
                        row = f"{st}| "
                        for m_idx in chunk:
                            beats = measures[m_idx].get("beats", TAB_BEATS_DEFAULT)
                            raw = measures[m_idx].get("strings", {}).get(st, "")
                            tokens = [tok for tok in raw.split() if tok]
                            while len(tokens) < beats:
                                tokens.append("-")
                            tokens = tokens[:beats]
                            row += "".join(f"{tok:>{TOKEN_W}}" for tok in tokens) + "|"
                        lines.append(row)
                    lines.append("")

            if sec.get("annotation"):
                lines.append(f'  "{sec["annotation"]}"')
                lines.append("")

        lines += [div("="), f"  Generated by Song Notation Tool v{APP_VERSION}", div("=")]
        return lines

    def _export_txt(self):
        self._commit_editor_line(force=True)
        t = THEMES[self.current_theme]
        dlg = tk.Toplevel(self)
        dlg.title("Export TXT")
        dlg.configure(bg=t["bg"])
        dlg.resizable(False, False)
        dlg.grab_set()

        tk.Label(dlg, text="Include instruments:", bg=t["bg"], fg=t["accent"],
                 font=FONT_TINY).pack(padx=16, pady=(12, 2), anchor="w")
        all_instrs = sorted({s.get("instrument", "") for s in self.doc["sections"]})
        instr_vars = {}
        for instr in all_instrs:
            v = tk.BooleanVar(value=True)
            instr_vars[instr] = v
            tk.Checkbutton(dlg, text=instr, variable=v, bg=t["bg"], fg=t["fg"],
                           selectcolor=t["input_bg"], activebackground=t["bg"],
                           font=FONT_TINY).pack(anchor="w", padx=28)

        def do_export():
            instrs = {i for i, v in instr_vars.items() if v.get()}
            dlg.destroy()
            artist = self.song_artist.get().strip()
            title  = self.song_title.get().strip()
            default_name = (f"{artist} - {title}" if artist else title
                            ).replace(" ", "_") + ".txt"
            path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialfile=default_name)
            if not path:
                return
            lines = self._build_song_lines(instruments=instrs)
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
            messagebox.showinfo("Exported", f"Saved to:\n{path}")

        br = tk.Frame(dlg, bg=t["bg"])
        br.pack(fill="x", padx=16, pady=(10, 14))
        ttk.Button(br, text="Cancel", command=dlg.destroy,
                   style="Normal.TButton").pack(side="right", padx=4)
        ttk.Button(br, text="Export TXT", command=do_export,
                   style="Accent.TButton").pack(side="right")
        dlg.bind("<Return>", lambda e: do_export())

    def _export_pdf(self):
        self._commit_editor_line(force=True)
        t = THEMES[self.current_theme]
        dlg = tk.Toplevel(self)
        dlg.title("Export PDF")
        dlg.configure(bg=t["bg"])
        dlg.resizable(False, False)
        dlg.grab_set()

        tk.Label(dlg, text="Page orientation:", bg=t["bg"], fg=t["accent"],
                 font=FONT_TINY).pack(padx=16, pady=(12, 2), anchor="w")
        orient_var = tk.StringVar(value="portrait")
        for val, lbl in [("portrait", "A4 Portrait  (target: one page)"),
                          ("landscape", "A4 Landscape")]:
            tk.Radiobutton(dlg, text=lbl, variable=orient_var, value=val,
                           bg=t["bg"], fg=t["fg"], selectcolor=t["input_bg"],
                           activebackground=t["bg"], font=FONT_TINY).pack(anchor="w", padx=28)

        tk.Label(dlg, text="Include instruments:", bg=t["bg"], fg=t["accent"],
                 font=FONT_TINY).pack(padx=16, pady=(10, 2), anchor="w")
        all_instrs = sorted({s.get("instrument", "") for s in self.doc["sections"]})
        instr_vars = {}
        for instr in all_instrs:
            v = tk.BooleanVar(value=True)
            instr_vars[instr] = v
            tk.Checkbutton(dlg, text=instr, variable=v, bg=t["bg"], fg=t["fg"],
                           selectcolor=t["input_bg"], activebackground=t["bg"],
                           font=FONT_TINY).pack(anchor="w", padx=28)

        def do_export():
            instrs = {i for i, v in instr_vars.items() if v.get()}
            orient = orient_var.get()
            dlg.destroy()
            artist = self.song_artist.get().strip()
            title  = self.song_title.get().strip()
            default_name = (f"{artist} - {title}" if artist else title
                            ).replace(" ", "_") + ".pdf"
            path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
                initialfile=default_name)
            if not path:
                return
            pdf_bytes = self._build_pdf(instruments=instrs, orient=orient)
            with open(path, "wb") as f:
                f.write(pdf_bytes)
            messagebox.showinfo("Exported PDF", f"Saved to:\n{path}")

        br = tk.Frame(dlg, bg=t["bg"])
        br.pack(fill="x", padx=16, pady=(10, 14))
        ttk.Button(br, text="Cancel", command=dlg.destroy,
                   style="Normal.TButton").pack(side="right", padx=4)
        ttk.Button(br, text="Export PDF", command=do_export,
                   style="Accent.TButton").pack(side="right")
        dlg.bind("<Return>", lambda e: do_export())

    def _build_pdf(self, instruments=None, orient="portrait"):
        """
        Build a one-file PDF, reusing the QLC+ Swiss Knife v0.4 byte-level
        writer (_assemble_pdf — no external libraries). Section content
        comes from the same render.render_chart_row()/resolve_display_items()
        calls the song map and Stage View use (design section 7: one
        renderer, not two code paths).
        """
        if instruments is None:
            instruments = {s.get("instrument", "") for s in self.doc["sections"]}

        W, H = (842, 595) if orient == "landscape" else (595, 842)
        MARGIN, LINE_H = 28, 12
        MONO_SZ, HEAD_SZ, TITLE_SZ, FOOTER_H = 7.5, 9, 13, 18
        TOKEN_W, CHAR_W, SN_W = 3, 4.6, 20

        artist   = self.song_artist.get().strip()
        title    = self.song_title.get().strip()
        doc_date = datetime.date.today().strftime("%Y-%m-%d")

        pages  = []
        cur_ln = []

        def _esc(s):
            s = str(s)
            for frm, to in [
                ("\u2014", "-"), ("\u2013", "-"), ("\u00d7", "x"), ("\u00d8", "x"),
                ("\u2019", "'"), ("\u2018", "'"), ("\u201c", '"'), ("\u201d", '"'),
                ("\u00e9", "e"), ("\u00e8", "e"), ("\u00e0", "a"), ("\u00f4", "o"),
                ("\u266a", ""), ("\u2665", ""), ("\u00d6", "O"), ("\u00fc", "u"),
                ("\u25b2", "^"), ("\u25bc", "v"), ("\u25c6", "*"), ("\u203a", ">"),
                ("\u00b7", "."),
            ]:
                s = s.replace(frm, to)
            s = s.encode("latin-1", errors="replace").decode("latin-1")
            return s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

        def txt(x, y, s, sz=MONO_SZ, bold=False):
            font = "/F2" if bold else "/F1"
            cur_ln.append(f"BT {font} {sz} Tf {x:.1f} {y:.1f} Td ({_esc(s)}) Tj ET")

        def color(r, g, b):
            cur_ln.append(f"{r:.3f} {g:.3f} {b:.3f} rg")

        def hline(x1, y, x2, width=0.25, gray=0.72):
            cur_ln.append(f"{gray:.2f} G {width} w {x1:.1f} {y:.1f} m {x2:.1f} {y:.1f} l S")

        def rfill(x, y, w, h, r, g, b):
            cur_ln.append(f"{r:.3f} {g:.3f} {b:.3f} rg {x:.1f} {y:.1f} {w:.1f} {h:.1f} re f")

        def finish_page(pn):
            hline(MARGIN, FOOTER_H, W - MARGIN)
            color(0.4, 0.4, 0.4)
            txt(MARGIN, 5, f"Song Notation Tool v{APP_VERSION}  -  {doc_date}", sz=6.5)
            txt(W - MARGIN - 28, 5, f"Page {pn}", sz=6.5)
            if cur_ln:
                pages.append(zlib.compress("\n".join(cur_ln).encode("latin-1")))
                cur_ln.clear()

        max_beats = max(
            (m.get("beats", TAB_BEATS_DEFAULT)
             for s in self.doc["sections"] if s.get("instrument") in instruments
             for m in songmap.measure_items(s)),
            default=TAB_BEATS_DEFAULT,
        )
        COL_W = TOKEN_W * max_beats * CHAR_W + CHAR_W

        def mpl_for(n_measures):
            usable = W - 2 * MARGIN - SN_W
            return max(1, min(n_measures, int(usable / COL_W)))

        pn = 1
        rfill(0, H - 46, W, 46, 0.10, 0.12, 0.22)
        color(1, 1, 1)
        hdr_txt = title.upper() + (f"  -  {artist}" if artist else "")
        txt(MARGIN, H - 28, hdr_txt, sz=TITLE_SZ, bold=True)
        meta_parts = []
        if self.song_key.get():   meta_parts.append(f"Key: {self.song_key.get()}")
        if self.song_tempo.get(): meta_parts.append(f"BPM: {self.song_tempo.get()}")
        if self.song_time.get():  meta_parts.append(f"Time: {self.song_time.get()}")
        if meta_parts:
            txt(MARGIN, H - 41, "   |   ".join(meta_parts), sz=7.5)
        cy = H - 54

        for sec in self.doc["sections"]:
            if sec.get("instrument") not in instruments:
                continue

            all_strings = INSTRUMENT_STRINGS.get(
                sec.get("instrument"), ["e", "B", "G", "D", "A", "E"])
            eff = transpose.effective_transpose(
                self.doc.get("transpose", 0), sec.get("transpose", 0))
            chart = songmap.chart_items(sec)
            chart_rows = (render.render_chart_row(render.resolve_display_items(chart, eff))
                          if chart else [])
            measures = songmap.measure_items(sec)
            strings = (render.active_strings(
                        [m.get("strings", {}) for m in measures], all_strings)
                       if measures else [])
            mpl = mpl_for(len(measures)) if measures else 1

            # Acceptance criterion (design §10): never split a section
            # across a page break. Estimate the section's height with the
            # same budget the live page indicator uses, and start a new
            # page before the section — not mid-way through it — if it
            # won't fit in what's left.
            needed_h = render.estimate_section_lines(sec, strings or all_strings) * LINE_H
            if cy - needed_h < FOOTER_H + LINE_H * 2:
                finish_page(pn); pn += 1; cy = H - MARGIN

            rep = sec.get("repeat", 1)
            rep_str = f" (x{rep})" if rep and rep != 1 else ""
            sec_label = f"[ {sec['name']} ]{rep_str}   {sec.get('instrument', '')}"
            rfill(MARGIN, cy - LINE_H, W - 2 * MARGIN, LINE_H + 2, 0.16, 0.24, 0.42)
            color(1, 1, 1)
            txt(MARGIN + 4, cy - LINE_H + 3, sec_label, sz=HEAD_SZ, bold=True)
            cy -= LINE_H + 6

            if chart_rows:
                color(0, 0, 0)
                for ln in chart_rows:
                    txt(MARGIN, cy, ln, sz=MONO_SZ)
                    cy -= LINE_H
                cy -= 2

            if sec.get("annotation"):
                color(0.3, 0.3, 0.3)
                txt(MARGIN, cy, f'"{sec["annotation"]}"', sz=MONO_SZ)
                cy -= LINE_H

            for bs in range(0, len(measures), mpl):
                batch = list(range(bs, min(bs + mpl, len(measures))))

                def col_x(i, batch=batch):
                    return MARGIN + SN_W + sum(
                        TOKEN_W * measures[batch[j]].get("beats", TAB_BEATS_DEFAULT)
                        * CHAR_W + CHAR_W
                        for j in range(i))

                color(0.40, 0.58, 0.82)
                for i, m_idx in enumerate(batch):
                    txt(col_x(i), cy, f"M{m_idx + 1}", sz=7)
                cy -= LINE_H

                for st in strings:
                    color(0.16, 0.32, 0.58)
                    txt(MARGIN, cy, f"{st}|", sz=MONO_SZ)
                    for i, m_idx in enumerate(batch):
                        beats = measures[m_idx].get("beats", TAB_BEATS_DEFAULT)
                        raw = measures[m_idx].get("strings", {}).get(st, "")
                        tokens = [tok for tok in raw.split() if tok]
                        while len(tokens) < beats:
                            tokens.append("-")
                        tokens = tokens[:beats]
                        row_str = "".join(f"{tok:>{TOKEN_W}}" for tok in tokens) + "|"
                        color(0, 0, 0)
                        txt(col_x(i), cy, row_str, sz=MONO_SZ)
                    cy -= LINE_H
                    if cy < FOOTER_H + LINE_H:
                        finish_page(pn); pn += 1; cy = H - MARGIN

                hline(MARGIN, cy, W - MARGIN, gray=0.82)
                cy -= 3

            cy -= 8

        finish_page(pn)
        return self._assemble_pdf(pages, W, H)

    @staticmethod
    def _assemble_pdf(pages, W, H):
        """
        Assemble raw PDF bytes from a list of zlib-compressed page streams.
        Ported directly from QLC+ Swiss Knife v0.4 — no external libraries.
        """
        raw     = "%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"
        offsets = []

        def add(s):
            nonlocal raw
            offsets.append(len(raw))
            raw += s

        def obj(n, c):
            return f"{n} 0 obj\n{c}\nendobj\n"

        def sobj(n, data):
            body = data.decode("latin-1")
            return obj(n, (f"<< /Length {len(data)} /Filter /FlateDecode >>\n"
                           f"stream\n{body}\nendstream"))

        font_res = "<< /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R >> >>"
        kids, po, so = [], [], []
        cid = 6

        for ps in pages:
            kids.append(f"{cid} 0 R")
            po.append(obj(cid,
                f"<< /Type /Page /Parent 2 0 R "
                f"/MediaBox [0 0 {W:.2f} {H:.2f}] "
                f"/Contents {cid + 1} 0 R /Resources {font_res} >>"))
            so.append(sobj(cid + 1, ps))
            cid += 2

        add(obj(1, "<< /Type /Catalog /Pages 2 0 R >>"))
        add(obj(2, f"<< /Type /Pages /Kids [{' '.join(kids)}] /Count {len(pages)} >>"))
        add(obj(3, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica "
                   "/Encoding /WinAnsiEncoding >>"))
        add(obj(4, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold "
                   "/Encoding /WinAnsiEncoding >>"))
        add(obj(5, "<< /Type /Font /Subtype /Type1 /BaseFont /Courier "
                   "/Encoding /WinAnsiEncoding >>"))
        for p, s in zip(po, so):
            add(p); add(s)

        n    = cid - 1
        xoff = len(raw)
        raw += f"xref\n0 {n + 1}\n0000000000 65535 f \n"
        for o in offsets:
            raw += f"{o:010d} 00000 n \n"
        raw += f"trailer\n<< /Size {n + 1} /Root 1 0 R >>\nstartxref\n{xoff}\n%%EOF\n"
        return raw.encode("latin-1")

    # ==========================================================================
    #  STAGE VIEW — large, high-contrast, read-only. Same renderer as TXT/PDF
    #  (design section 7.3), just a different scale.
    # ==========================================================================

    def _open_stage_view(self):
        self._commit_editor_line(force=True)
        top = tk.Toplevel(self)
        top.title(f"Stage View — {self.song_title.get()}")
        top.configure(bg="#000000")
        try:
            top.attributes("-fullscreen", True)
        except tk.TclError:
            top.geometry("1000x700")

        txt_widget = tk.Text(top, bg="#000000", fg="#ffffff", font=FONT_STAGE,
                              wrap="none", relief="flat", insertwidth=0,
                              padx=24, pady=24)
        txt_widget.pack(fill="both", expand=True)
        txt_widget.insert("1.0", "\n".join(self._build_song_lines()))
        txt_widget.configure(state="disabled")

        top.bind("<Up>",     lambda e: txt_widget.yview_scroll(-2, "units"))
        top.bind("<Down>",   lambda e: txt_widget.yview_scroll(2, "units"))
        top.bind("<Prior>",  lambda e: txt_widget.yview_scroll(-15, "units"))
        top.bind("<Next>",   lambda e: txt_widget.yview_scroll(15, "units"))
        top.bind("<Escape>", lambda e: top.destroy())
        top.focus_set()

    # ==========================================================================
    #  SAVE / OPEN
    # ==========================================================================

    def _save(self):
        self._commit_editor_line(force=True)
        self.doc["meta"] = {
            "title": self.song_title.get(), "artist": self.song_artist.get(),
            "key": self.song_key.get(), "time": self.song_time.get(),
            "bpm": self.song_tempo.get(),
        }
        self.doc["app_version"] = APP_VERSION
        artist = self.song_artist.get().strip()
        title  = self.song_title.get().strip()
        default_name = (f"{artist} - {title}" if artist else title
                        ).replace(" ", "_") + ".sng"

        path = filedialog.asksaveasfilename(
            defaultextension=".sng",
            filetypes=[("Song files", "*.sng"), ("All files", "*.*")],
            initialfile=default_name)
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.doc, f, indent=2)
        self.dirty = False
        messagebox.showinfo("Saved", f"Project saved to:\n{path}")

    def _open(self):
        path = filedialog.askopenfilename(
            filetypes=[("Song files", "*.sng"), ("All files", "*.*")])
        if not path:
            return
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        self.doc = model.migrate_document(data)
        meta = self.doc.get("meta", {})
        self.song_title.set(meta.get("title", data.get("title", "")))
        self.song_artist.set(meta.get("artist", data.get("artist", "")))
        self.song_key.set(meta.get("key", data.get("key", "")))
        self.song_tempo.set(meta.get("bpm", data.get("tempo", "")))
        self.song_time.set(meta.get("time", data.get("time", "4/4")))
        if self.doc["sections"]:
            instr = self.doc["sections"][0].get("instrument")
            if instr:
                self.default_instrument.set(instr)

        self.focus_kind = self.focus_id = None
        self._expanded_tab.clear()
        self._set_editor_text("")
        self.lbl_editor_hint.configure(text="Click a section below to edit its chart line")

        sec_nums = [int(m.group()) for sid in
                    (s["id"] for s in self.doc["sections"])
                    for m in [_re.search(r"\d+$", sid)] if m]
        self._next_ids["section"] = (max(sec_nums) + 1) if sec_nums else 1
        blk_nums = [int(m.group()) for bid in self.doc["blocks"].keys()
                    for m in [_re.search(r"\d+$", bid)] if m]
        self._next_ids["block"] = (max(blk_nums) + 1) if blk_nums else 1

        self._rebuild_map()
        self._refresh_riff_strip()
        self._refresh_page_indicator()
        self.dirty = False

    # ==========================================================================
    #  KEYBOARD SHORTCUTS & WINDOW CLOSE  (design section 8)
    # ==========================================================================

    def _bind_shortcuts(self):
        self.bind("<Command-s>", lambda e: self._save())
        self.bind("<Control-s>", lambda e: self._save())
        self.bind("<Command-e>", lambda e: self._export_txt())
        self.bind("<Control-e>", lambda e: self._export_txt())
        self.bind("<Command-n>", lambda e: self._add_section(focus=True))
        self.bind("<Control-n>", lambda e: self._add_section(focus=True))
        self.bind("<Command-p>", lambda e: self._open_stage_view())
        self.bind("<Control-p>", lambda e: self._open_stage_view())
        self.bind("<Command-P>", lambda e: self._export_pdf())
        self.bind("<Control-P>", lambda e: self._export_pdf())
        self.bind("<Command-r>", self._promote_selection_to_riff)
        self.bind("<Control-r>", self._promote_selection_to_riff)
        self.bind("<Command-d>", self._duplicate_as_reference)
        self.bind("<Control-d>", self._duplicate_as_reference)
        self.bind("<Alt-Up>",    lambda e: self._move_focused_section(-1))
        self.bind("<Alt-Down>",  lambda e: self._move_focused_section(1))
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        self._commit_editor_line(force=True)
        if self.dirty:
            if messagebox.askyesno("Quit", "You have unsaved changes. Quit anyway?"):
                self.destroy()
        else:
            self.destroy()


def _slug(name):
    """Turn a riff name into a valid block-id stem: letters/digits only,
    starting with a letter (model.validate_block_name rejects a leading
    fret-like prefix such as '5A')."""
    s = "".join(c for c in (name or "") if c.isalnum())
    if not s or not s[0].isalpha():
        s = "riff" + s
    return s or "riff"


# ==============================================================================
#  ENTRY POINT
# ==============================================================================

if __name__ == "__main__":
    app = SongNotationApp()
    app.mainloop()

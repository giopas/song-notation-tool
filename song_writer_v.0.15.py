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
#
# ==============================================================================

import copy      # deep-copy Section objects in the copy dialog
import datetime  # timestamp in PDF footer
import json      # .sng project files are plain JSON
import os
import platform
import re        # used in the Transpose dialog to shift fret numbers
import zlib      # PDF page stream compression (FlateDecode)
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ── macOS: suppress deprecation noise ────────────────────────────────────────
if platform.system() == "Darwin":
    os.environ['SYSTEM_VERSION_COMPAT'] = '0'
    os.environ['TK_SILENCE_DEPRECATION'] = '1'

# ==============================================================================
#  VERSION — single source of truth; bumped here propagates everywhere
# ==============================================================================
APP_VERSION = "0.15"
APP_TITLE   = f"Song Notation Tool  v{APP_VERSION}"

# ==============================================================================
#  DOMAIN CONSTANTS
# ==============================================================================

# Section-type presets shown in the New / Edit Section dialog
SECTION_TYPES = [
    "Intro", "Verse", "Pre-Chorus", "Chorus", "Refrain",
    "Bridge", "Interlude", "Solo", "Breakdown", "Outro", "Custom",
]

# Maps instrument/tuning → string list (high → low pitch).
# Used to label tab rows and render the TXT export.
INSTRUMENT_STRINGS = {
    "Guitar (6-string)":        ["e", "B", "G", "D", "A", "E"],
    "Guitar (6-string) Drop D": ["e", "B", "G", "D", "A", "D"],
    "Guitar (7-string)":        ["e", "B", "G", "D", "A", "E", "B"],
    "Bass (4-string)":          ["G", "D", "A", "E"],
    "Bass (4-string) Drop D":   ["G", "D", "A", "D"],
    "Bass (5-string)":          ["G", "D", "A", "E", "B"],
}

# Default number of beats per tab measure. Can be overridden per section.
TAB_BEATS_DEFAULT = 8
TAB_BEATS_OPTIONS = [8, 16, 32, 64]   # choices in toolbar and per-measure picker

# Badge colours for the four notation layers — one shade per theme
LAYER_COLORS = {
    "tab":    {"dark": "#1e4d8c", "light": "#2563eb"},
    "chords": {"dark": "#5b21b6", "light": "#7c3aed"},
    "notes":  {"dark": "#065f46", "light": "#059669"},
    "lyrics": {"dark": "#92400e", "light": "#b45309"},
}

# Shared fonts — Courier New keeps tab columns aligned on all platforms
FONT_MAIN = ("Courier New", 11)
FONT_HEAD = ("Courier New", 13, "bold")
FONT_TINY = ("Courier New", 9)
FONT_MONO = ("Courier New", 10)   # inside measure cells

# ==============================================================================
#  CHROMATIC NOTE SYSTEM
#  Used for transposing notes/chords layers and for the root-note picker.
# ==============================================================================

# Canonical chromatic scale — sharp spelling is the primary form.
# Index 0 = A, following the A-based cycle used in guitar/bass notation.
CHROMATIC = ["A", "A#", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#"]

# Flat aliases → canonical sharp equivalent
ENHARMONIC = {
    "Bb": "A#", "Db": "C#", "Eb": "D#", "Gb": "F#", "Ab": "G#",
    # Also accept lowercase "b" variants
    "Bb": "A#", "bb": "A#",
}

# Complete display list shown in the root-note picker dropdown.
# Enharmonic pairs are shown as "A# / Bb" so users can identify either name.
NOTE_DISPLAY = [
    "A", "A# / Bb", "B", "C", "C# / Db",
    "D", "D# / Eb", "E", "F", "F# / Gb", "G", "G# / Ab",
]

# Map from display string back to canonical sharp name
NOTE_DISPLAY_TO_CANONICAL = {
    "A": "A", "A# / Bb": "A#", "B": "B", "C": "C", "C# / Db": "C#",
    "D": "D", "D# / Eb": "D#", "E": "E", "F": "F", "F# / Gb": "F#",
    "G": "G", "G# / Ab": "G#",
}

# Reverse: canonical → display string
NOTE_CANONICAL_TO_DISPLAY = {v: k for k, v in NOTE_DISPLAY_TO_CANONICAL.items()}


def note_to_index(note_str):
    """
    Convert a note name to its chromatic index (0=A … 11=G#).
    Handles sharp (#), flat (b) spellings and case variations.
    Returns None if the string is not a recognised note.
    """
    n = note_str.strip()
    # Try canonical first
    if n in CHROMATIC:
        return CHROMATIC.index(n)
    # Try enharmonic map
    if n in ENHARMONIC:
        return CHROMATIC.index(ENHARMONIC[n])
    # Case-insensitive fallback
    nu = n[0].upper() + n[1:] if len(n) > 1 else n.upper()
    if nu in CHROMATIC:
        return CHROMATIC.index(nu)
    if nu in ENHARMONIC:
        return CHROMATIC.index(ENHARMONIC[nu])
    return None


def transpose_note(note_str, semitones):
    """
    Shift a note name by `semitones` steps.
    Returns the transposed note in canonical sharp form, or the original
    string unchanged if it is not a recognised note name.
    """
    idx = note_to_index(note_str)
    if idx is None:
        return note_str   # not a note — leave untouched
    new_idx = (idx + semitones) % 12
    return CHROMATIC[new_idx]


def transpose_chord(chord_str, semitones):
    """
    Transpose a chord string such as "Am", "F#7", "Bbm7", "C# dim".
    Extracts the root note (1–2 chars), transposes it, keeps the suffix.
    Returns the transposed chord, or the original if no note is found.
    """
    if not chord_str:
        return chord_str
    chord_str = chord_str.strip()
    # Try 2-char root first (e.g. "A#", "Bb", "C#")
    for length in (2, 1):
        root = chord_str[:length]
        suffix = chord_str[length:]
        if note_to_index(root) is not None:
            new_root = transpose_note(root, semitones)
            return new_root + suffix
    return chord_str


def _parse_root_suffix(value):
    """
    Parse a chord or note string into (canonical_root, suffix).
    e.g. "Am7"  → ("A",  "m7")
         "F#dim" → ("F#", "dim")
         "Bbsus4"→ ("A#", "sus4")
         "E"     → ("E",  "")
         ""      → (None, "")
         "hello" → (None, "hello")   ← unrecognised, keep as suffix
    """
    v = (value or "").strip()
    if not v:
        return None, ""
    for length in (2, 1):
        root = v[:length]
        canon = None
        if root in CHROMATIC:
            canon = root
        elif root in ENHARMONIC:
            canon = ENHARMONIC[root]
        if canon is not None:
            return canon, v[length:]
    return None, v   # no note found — treat whole string as suffix


# ==============================================================================
#  TAB ENTRY VALIDATION  (new in v0.12)
# ==============================================================================

_VALID_TOKEN_CHARS = set("0123456789-/hpbx~")

def _validate_tab_entry(entry_widget, beats):
    """
    After a keystroke in a tab Entry, enforce that:
      1. Content splits into exactly `beats` whitespace-separated tokens.
      2. Each token is valid (digits + decorators /hpbx~-). Invalid → "-".
      3. Token count is padded to `beats` with "-", or trimmed if over.

    Special care: if the user is mid-typing a multi-digit fret (e.g. has
    typed "1" and will follow with "0" to make "10"), we must NOT reformat
    until they move on. Strategy: only reformat if the token count is wrong
    OR if there are clearly invalid chars; never touch a token that is a
    valid prefix (pure digits, no space yet typed after it at the end).
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
        if all(c in _VALID_TOKEN_CHARS for c in tok_clean):
            cleaned.append(tok_clean)
        else:
            cleaned.append("-")

    # Only rewrite if count is wrong or we have clearly invalid content
    if len(cleaned) == beats and all(
            all(c in _VALID_TOKEN_CHARS for c in t) for t in cleaned):
        # Already valid — just check nothing is obviously wrong
        if cleaned == raw_tokens:
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
#  SMART TAB TRANSPOSITION  (new in v0.12)
# ==============================================================================

# Open-string MIDI note numbers per instrument (strings listed low → high).
# MIDI note 40 = E2, 45 = A2, 50 = D3, 55 = G3, 59 = B3, 64 = e4
# Bass: 28 = E1, 33 = A1, 38 = D2, 43 = G2
_INSTRUMENT_STRINGS_LOW_HIGH = {
    "Guitar (6-string)":        ["E", "A", "D", "G", "B", "e"],
    "Guitar (6-string) Drop D": ["D", "A", "D", "G", "B", "e"],
    "Guitar (7-string)":        ["B", "E", "A", "D", "G", "B", "e"],
    "Bass (4-string)":          ["E", "A", "D", "G"],
    "Bass (4-string) Drop D":   ["D", "A", "D", "G"],
    "Bass (5-string)":          ["B", "E", "A", "D", "G"],
}

_OPEN_MIDI_BY_INSTRUMENT = {
    "Guitar (6-string)":        [40, 45, 50, 55, 59, 64],
    "Guitar (6-string) Drop D": [38, 45, 50, 55, 59, 64],
    "Guitar (7-string)":        [35, 40, 45, 50, 55, 59, 64],
    "Bass (4-string)":          [28, 33, 38, 43],
    "Bass (4-string) Drop D":   [26, 33, 38, 43],
    "Bass (5-string)":          [23, 28, 33, 38, 43],
}

_MAX_FRET = 24


def _transpose_fret_token(token, semitones, string_name, instrument):
    """
    Transpose one fret token on a given string.
    Returns (new_token, new_string_name).
    If new_string_name != string_name, the token should move to that string.

    Rules:
      - Try the same string first (fret in [0, _MAX_FRET]).
      - If fret < 0: search lower strings (ascending from lowest) for one
        where fret >= 0.  If none, use open string (fret 0) on the lowest.
      - If fret > _MAX_FRET: search higher strings.
    """
    if not re.fullmatch(r'\d+', token):
        return token, string_name   # decorator — pass through

    fret = int(token)
    strings_lh = _INSTRUMENT_STRINGS_LOW_HIGH.get(instrument)
    open_midis = _OPEN_MIDI_BY_INSTRUMENT.get(instrument)

    if strings_lh is None or open_midis is None:
        return str(max(0, fret + semitones)), string_name

    try:
        str_idx = strings_lh.index(string_name)
    except ValueError:
        return str(max(0, fret + semitones)), string_name

    target_midi = open_midis[str_idx] + fret + semitones
    new_fret    = target_midi - open_midis[str_idx]

    if 0 <= new_fret <= _MAX_FRET:
        return str(new_fret), string_name

    if new_fret < 0:
        # Look for a lower string that can play this pitch
        for i in range(str_idx - 1, -1, -1):
            candidate = target_midi - open_midis[i]
            if 0 <= candidate <= _MAX_FRET:
                return str(candidate), strings_lh[i]
        # No lower string — use open string on the lowest
        return "0", strings_lh[0]

    # new_fret > _MAX_FRET — look for a higher string
    for i in range(str_idx + 1, len(strings_lh)):
        candidate = target_midi - open_midis[i]
        if 0 <= candidate <= _MAX_FRET:
            return str(candidate), strings_lh[i]
    return str(min(new_fret, _MAX_FRET)), string_name


def transpose_tab_cell(cell_dict, semitones, instrument):
    """
    Transpose one measure tab cell {string_name: row_text} by `semitones`.
    Handles cross-string reassignment: tokens that go below fret 0 are
    moved to the appropriate lower string, or stay on the lowest with fret 0.
    """
    strings_lh = _INSTRUMENT_STRINGS_LOW_HIGH.get(instrument)
    open_midis = _OPEN_MIDI_BY_INSTRUMENT.get(instrument)

    if strings_lh is None:
        # Unknown instrument — simple numeric shift
        return {
            sn: re.sub(r'\d+',
                       lambda m, _s=semitones: str(max(0, int(m.group()) + _s)),
                       row)
            for sn, row in cell_dict.items()
        }

    beats = max((len(row.split()) for row in cell_dict.values() if row),
                default=TAB_BEATS_DEFAULT)

    # Parse all string rows into token lists, padded to `beats`
    rows = {}
    for sn in strings_lh:
        raw  = cell_dict.get(sn, "")
        toks = raw.split() if raw else []
        while len(toks) < beats:
            toks.append("-")
        rows[sn] = toks[:beats]

    # Output token lists — start with all dashes
    out = {sn: ["-"] * beats for sn in strings_lh}

    for sn in strings_lh:
        for i, tok in enumerate(rows[sn]):
            if tok == "-":
                if out[sn][i] == "-":
                    out[sn][i] = "-"
            elif re.fullmatch(r'\d+', tok):
                new_tok, new_sn = _transpose_fret_token(tok, semitones, sn, instrument)
                if new_sn not in out:
                    new_sn = sn
                out[new_sn][i] = new_tok
            else:
                out[sn][i] = tok   # decorator — keep on same string

    # Reconstruct only the strings present in the original cell
    return {sn: "  ".join(out.get(sn, ["-"] * beats))
            for sn in cell_dict}


THEMES = {
    "dark": {
        "bg":        "#1e1e2e",   # main window background
        "fg":        "#cdd6f4",   # default text
        "card_bg":   "#16213e",   # measure cards, section panel
        "header_bg": "#181825",   # top bar rows
        "input_bg":  "#0f3460",   # Entry / Listbox fill
        "border":    "#45475a",   # sash, separator, card outline
        "surface":   "#313244",   # scrollbar track
        "btn_bg":    "#313244",   # Normal button background
        "select_bg": "#89b4fa",   # selection / hover highlight
        "select_fg": "#1e1e2e",   # text on highlight
        "lbl_gray":  "#6c7086",   # secondary / dim labels
        "accent":    "#e94560",   # app title, section header, M-numbers
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
#  TAB ENTRY HELPERS
#  Each string row in a measure is a tk.Entry pre-filled with dash tokens.
#  The storage format is a space-separated list of tokens, e.g.:
#      "-  -  -  2  -  -  5  -"
#  which renders horizontally as:
#      G| -  -  -  2  -  -  5  - |
# ==============================================================================

def make_blank_tab_row(beats=TAB_BEATS_DEFAULT):
    """Return the default dash-filled content for one string row."""
    return "  ".join(["-"] * beats)

def parse_tab_row(text):
    """
    Parse a tab row string back into a list of tokens.
    Tolerates extra spaces and non-standard formatting.
    """
    # Split on one-or-more spaces; filter empty strings
    return [t for t in text.split() if t] or ["-"]

def normalise_tab_row(text, beats=TAB_BEATS_DEFAULT):
    """
    Re-serialise a tab row to the canonical space-separated format.
    Keeps existing tokens, padding with dashes if shorter than `beats`.
    """
    tokens = parse_tab_row(text)
    # Pad to at least `beats` tokens
    while len(tokens) < beats:
        tokens.append("-")
    return "  ".join(tokens)

# ==============================================================================
#  DATA MODEL
# ==============================================================================

class Section:
    """
    One structural section of a song (Verse, Chorus, etc.).

    Attributes
    ----------
    name         : display name in the listbox
    section_type : canonical type (one of SECTION_TYPES)
    measures     : number of measure columns in the editor
    instrument   : key into INSTRUMENT_STRINGS
    repeat       : repeat count (shown as ×N in the listbox)
    layers       : {layer: list[str]} — one entry per measure
                   For "tab", each entry is a dict {string_name: row_text}.
                   For other layers, each entry is a plain string.
    visible      : {layer: bool} — controls which rows are shown
    """

    def __init__(self, name, section_type, measures, instrument):
        self.name         = name
        self.section_type = section_type
        self.measures     = measures
        self.instrument   = instrument
        self.repeat    = 1
        # tab_beats: default beat count for new/empty measures in this section
        self.tab_beats = TAB_BEATS_DEFAULT
        # measure_beats: per-measure beat count overrides {m_idx: beats}
        # If a measure is not in this dict, it uses tab_beats.
        self.measure_beats: dict = {}
        # link_id
        # None means the section is not linked to anything.
        self.link_id = None
        # Tab layer stores one dict per measure: {string_name: row_text}
        # Other layers store a plain string per measure.
        self.layers = {
            "tab":    [{} for _ in range(measures)],
            "chords": [""] * measures,
            "notes":  [""] * measures,
            "lyrics": [""] * measures,
        }
        self.visible = {"tab": True, "chords": True,
                        "notes": False, "lyrics": False}

    def resize(self, n):
        """Grow or trim every layer to exactly n measures."""
        for k in self.layers:
            old = self.layers[k]
            if k == "tab":
                self.layers[k] = (old + [{} for _ in range(n - len(old))])[:n]
            else:
                self.layers[k] = (old + [""] * (n - len(old)))[:n]
        self.measures = n

    def to_dict(self):
        """Serialise to a JSON-compatible dict."""
        return {
            "name": self.name, "section_type": self.section_type,
            "measures": self.measures, "instrument": self.instrument,
            "repeat": self.repeat, "link_id": self.link_id,
            "tab_beats": self.tab_beats,
            "measure_beats": {str(k): v for k, v in self.measure_beats.items()},
            "layers": self.layers, "visible": self.visible,
        }

    @staticmethod
    def from_dict(d):
        """Deserialise from a saved dict."""
        s = Section(d["name"], d["section_type"], d["measures"], d["instrument"])
        s.repeat    = d.get("repeat", 1)
        s.link_id   = d.get("link_id", None)
        s.tab_beats = d.get("tab_beats", TAB_BEATS_DEFAULT)
        # measure_beats keys are stored as strings in JSON — convert back to int
        s.measure_beats = {int(k): v for k, v in
                           d.get("measure_beats", {}).items()}
        s.visible = d.get("visible", {k: True for k in d["layers"]})
        raw = d["layers"]
        # Migrate old format (tab as newline-joined string → new dict format)
        tab_raw = raw.get("tab", [{} for _ in range(s.measures)])
        new_tab = []
        strings = INSTRUMENT_STRINGS.get(s.instrument, ["e","B","G","D","A","E"])
        for cell in tab_raw:
            if isinstance(cell, dict):
                new_tab.append(cell)           # already new format
            elif isinstance(cell, str) and cell:
                # Old format: "e|...\nB|...\n..."
                row_dict = {}
                for line in cell.split("\n"):
                    if "|" in line:
                        sn, content = line.split("|", 1)
                        sn = sn.strip()
                        row_dict[sn] = content.strip()
                new_tab.append(row_dict)
            else:
                new_tab.append({})
        s.layers = {
            "tab":    new_tab,
            "chords": raw.get("chords", [""] * s.measures),
            "notes":  raw.get("notes",  [""] * s.measures),
            "lyrics": raw.get("lyrics", [""] * s.measures),
        }
        return s

# ==============================================================================
#  MAIN APPLICATION
# ==============================================================================

class SongNotationApp(tk.Tk):
    """
    Main application window.

    Layout (v0.5 — two-row topbar, horizontal tab grid):

        ┌─ topbar row1: logo | Title | Artist | [Save][Open][Export][Theme] ─┐
        ├─ topbar row2: Key | BPM | Time | Instrument | [Drop D][Transpose] ─┤
        ├─ separator ─────────────────────────────────────────────────────────┤
        │ left panel (sections)  │  right panel (editor)                      │
        │  SECTIONS  [+]         │  [ Section name ]  Layers: [tab][chords]…  │
        │  ┌──────────────────┐  │  ┌──M1──────────┬──M2──────────┬── ···    │
        │  │ Intro  [4m]      │  │  │ G| - - - - - │ G| - - - - -│           │
        │  │ Verse  [8m] ×2   │  │  │ D| - - - 2 - │ D| - - - - -│           │
        │  └──────────────────┘  │  │ A| - 0 - - - │ A| - - - - -│           │
        │  [✏][⎘][⬆][⬇][✕]     │  │ E| - - - - - │ E| - 5 - - -│           │
        └────────────────────────┴──┴──────────────┴─────────────┘           │
    """

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1100x820")
        self.minsize(900, 650)

        # Theme
        self.current_theme = "dark"
        self.style = ttk.Style(self)
        self.style.theme_use("clam")   # 'clam' allows full colour overrides

        # Song-level metadata
        self.song_title  = tk.StringVar(value="Untitled Song")
        self.song_artist = tk.StringVar(value="")
        self.song_key    = tk.StringVar(value="")
        self.song_tempo  = tk.StringVar(value="")
        self.song_time   = tk.StringVar(value="4/4")
        self.instrument  = tk.StringVar(value="Guitar (6-string)")

        # Section / editor state
        self.sections:            list = []
        self.current_section_idx: int  = None
        # _section_widgets[m_idx] = {"tab": {string_name: Entry}, "chords": Entry, …}
        self._section_widgets:    dict = {}
        self._toggle_buttons:     dict = {}
        self._measure_cards:      list = []
        # Internal clipboard for measure copy/paste: dict {string_name: row_text}
        self._measure_clipboard:  dict = {}
        # Link group counter — each new link group gets a unique integer id
        self._next_link_id:       int  = 1
        # Wrap mode: when True, measures reflow into multiple rows
        self.wrap_mode = tk.BooleanVar(value=False)

        # Widget refs needed by apply_theme
        self._topbar_entries: list = []
        self._topbar_labels:  list = []

        self.dirty = False

        self._build_ui()
        self.apply_theme()
        self._bind_shortcuts()

    # ==========================================================================
    #  UI CONSTRUCTION
    # ==========================================================================

    def _build_ui(self):
        """Build the two topbar rows, separator, and paned main area."""

        # ── Row 1: logo | Title | Artist | action buttons ─────────────────────
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

        # ── Right side: hamburger menu (always visible) + text buttons ──────────
        # The ≡ menu is packed first (rightmost) so it's always accessible even
        # on very narrow windows. Text buttons are packed after and will be
        # hidden by the window edge if there's no room.
        self.btn_theme = ttk.Button(self.topbar, text="🌗 Theme",
                                     command=self.toggle_theme,
                                     style="Normal.TButton")
        self.btn_theme.pack(side="right", padx=3)
        ToolTip(self.btn_theme, "Toggle light / dark theme")

        # Hamburger menu — red Accent button, contains all actions including Theme
        def _show_hamburger_menu():
            m = tk.Menu(self, tearoff=0)
            t_theme = THEMES[self.current_theme]
            m.configure(bg=t_theme["card_bg"], fg=t_theme["fg"],
                        activebackground=t_theme["select_bg"],
                        activeforeground=t_theme["select_fg"],
                        font=FONT_TINY)
            m.add_command(label="💾  Save project",    command=self._save)
            m.add_command(label="📂  Open project",    command=self._open)
            m.add_separator()
            m.add_command(label="📋  Export PDF",      command=self._export_pdf)
            m.add_command(label="📄  Export TXT",      command=self._export_txt)
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
        ToolTip(self.btn_menu, "Menu: Save / Open / Export / Theme")

        # Text buttons — visible when there's room, redundant with ≡ menu
        for text, cmd, tip in [
            ("📄 Export TXT", self._export_txt,
             "Export as plain-text tab file  (Cmd/Ctrl+E)"),
            ("📋 Export PDF", self._export_pdf,
             "Export as PDF  (Cmd/Ctrl+P)"),
            ("📂 Open",       self._open,    "Open a .sng project file"),
            ("💾 Save",       self._save,    "Save project  (Cmd/Ctrl+S)"),
        ]:
            b = ttk.Button(self.topbar, text=text, command=cmd,
                           style="Accent.TButton")
            b.pack(side="right", padx=3)
            ToolTip(b, tip)

        # ── Row 2: Key | BPM | Time | Instrument | Drop D | Transpose ─────────
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

        lbl_i = tk.Label(self.topbar2, text="Instrument:", font=FONT_TINY)
        lbl_i.pack(side="left")
        self._topbar_labels.append(lbl_i)
        self.instr_cb = ttk.Combobox(
            self.topbar2, textvariable=self.instrument,
            values=list(INSTRUMENT_STRINGS.keys()),
            width=22, state="readonly", font=FONT_MAIN)
        self.instr_cb.pack(side="left", padx=(2, 12))

        self.btn_drop_d = ttk.Button(
            self.topbar2, text="Drop D",
            command=self._apply_drop_d, style="Normal.TButton")
        self.btn_drop_d.pack(side="left", padx=3)
        ToolTip(self.btn_drop_d,
                "Toggle the current section's instrument to its Drop D variant.")

        self.btn_transpose = ttk.Button(
            self.topbar2, text="Transpose ↕",
            command=self._transpose_dialog, style="Normal.TButton")
        self.btn_transpose.pack(side="left", padx=3)
        ToolTip(self.btn_transpose,
                "Shift all fret numbers in the current section by ±N semitones.")

        # ── Separator ─────────────────────────────────────────────────────────
        self.sep = tk.Frame(self, height=2)
        self.sep.pack(fill="x")

        # ── Main paned area ───────────────────────────────────────────────────
        # sashwidth=8 gives a wide enough grab target; sashcursor gives visual
        # feedback that it can be dragged. The sash colour is set in apply_theme.
        self.pane = tk.PanedWindow(self, orient="horizontal",
                                    sashwidth=8, sashrelief="raised",
                                    sashcursor="sb_h_double_arrow")
        self.pane.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        self.left_panel = tk.Frame(self.pane, width=260)
        self.pane.add(self.left_panel, minsize=160, stretch="never")
        self._build_section_panel(self.left_panel)

        self.right_panel = tk.Frame(self.pane)
        self.pane.add(self.right_panel, minsize=400, stretch="always")
        self._build_editor_panel(self.right_panel)

    def _build_section_panel(self, parent):
        """Left panel: SECTIONS header, scrollable listbox, action buttons."""

        self.sec_hdr = tk.Frame(parent)
        self.sec_hdr.pack(fill="x", padx=6, pady=(8, 4))
        self.lbl_sections = tk.Label(self.sec_hdr, text="SECTIONS", font=FONT_HEAD)
        self.lbl_sections.pack(side="left")
        self.btn_add_sec = ttk.Button(
            self.sec_hdr, text=" + ", width=3,
            command=self._add_section_dialog, style="Accent.TButton")
        self.btn_add_sec.pack(side="right")
        ToolTip(self.btn_add_sec, "Add new section  (Cmd/Ctrl+N)")

        lbframe = tk.Frame(parent)
        lbframe.pack(fill="both", expand=True, padx=6, pady=4)
        sb = tk.Scrollbar(lbframe, orient="vertical")
        self.section_listbox = tk.Listbox(
            lbframe, relief="flat", font=FONT_MAIN,
            yscrollcommand=sb.set, activestyle="none", borderwidth=0)
        sb.config(command=self.section_listbox.yview)
        sb.pack(side="right", fill="y")
        self.section_listbox.pack(fill="both", expand=True)
        self.section_listbox.bind("<<ListboxSelect>>", self._on_section_select)

        # Bottom action buttons — laid out in a grid that wraps to 2 rows
        # when the panel is too narrow for all 6 buttons on one line.
        # Row 0: Edit | Copy | Link
        # Row 1: ⬆   | ⬇   | ✕
        self.sec_btn_frame = tk.Frame(parent)
        self.sec_btn_frame.pack(fill="x", padx=6, pady=(0, 8))

        # Configure 3 equal-weight columns so buttons share available width
        for col in range(3):
            self.sec_btn_frame.grid_columnconfigure(col, weight=1)

        btn_defs = [
            # (row, col, text, cmd, tip)
            (0, 0, "✏ Edit", self._edit_section_dialog, "Edit selected section"),
            (0, 1, "⎘ Copy", self._copy_section_dialog, "Clone section"),
            (0, 2, "🔗 Link", self._link_section_dialog, "Link / unlink sections"),
            (1, 0, "⬆",      self._move_section_up,     "Move section up"),
            (1, 1, "⬇",      self._move_section_down,   "Move section down"),
            (1, 2, "✕",      self._delete_section,      "Delete section"),
        ]
        for row, col, text, cmd, tip in btn_defs:
            b = ttk.Button(self.sec_btn_frame, text=text, command=cmd,
                           style="Normal.TButton")
            b.grid(row=row, column=col, padx=2, pady=2, sticky="ew")
            ToolTip(b, tip)

    def _build_editor_panel(self, parent):
        """Right panel: section title bar, layer toggles, scrollable measure grid."""

        # Section info line
        self.editor_title = tk.Label(
            parent, text="← Add a section to begin",
            font=FONT_HEAD, anchor="w")
        self.editor_title.pack(fill="x", padx=10, pady=(8, 2))

        # Permanent hint for the tab layer
        self.tab_hint = tk.Label(
            parent,
            text="Tab: each string row shows beats as  -  or fret numbers.  "
                 "Type a fret number to replace a beat.  "
                 "0=open  x=mute  h=hammer  p=pull  /=slide  b=bend",
            font=("Courier New", 8), anchor="w")
        self.tab_hint.pack(fill="x", padx=10, pady=(0, 4))

        # Layer toggle buttons (rebuilt each time a section loads)
        # and Wrap Mode toggle + Beats selector — all in the same toolbar row
        toolbar = tk.Frame(parent)
        toolbar.pack(fill="x", padx=10, pady=(0, 6))

        self.toggle_frame = tk.Frame(toolbar)
        self.toggle_frame.pack(side="left", fill="x", expand=True)

        # ── Right side of toolbar: Beats selector + Wrap toggle ───────────────

        # Wrap mode toggle
        self.btn_wrap = ttk.Button(
            toolbar, text="↵ Wrap",
            command=self._toggle_wrap,
            style="Normal.TButton")
        self.btn_wrap.pack(side="right", padx=(4, 0))
        ToolTip(self.btn_wrap,
                "Wrap measures into multiple rows to fit the window width")

        # Beats-per-measure selector — small segmented control
        # Label
        lbl_beats = tk.Label(toolbar, text="Beats/m:", font=FONT_TINY)
        lbl_beats.pack(side="right", padx=(8, 2))
        self._beats_lbl = lbl_beats   # kept for apply_theme

        # Three buttons: 8 | 16 | 32
        self._beats_btn_frame = tk.Frame(toolbar)
        self._beats_btn_frame.pack(side="right")
        self._beats_buttons = {}   # beats_value → ttk.Button

        def _make_beats_setter(beats_val):
            def _set():
                idx = self.current_section_idx
                if idx is None:
                    return
                self._save_current_section()
                self.sections[idx].tab_beats = beats_val
                self._section_widgets = {}
                self.dirty = True
                self._load_section(idx)
                self._update_beats_buttons()
            return _set

        for bval in TAB_BEATS_OPTIONS:
            b = ttk.Button(self._beats_btn_frame,
                           text=str(bval),
                           width=3,
                           command=_make_beats_setter(bval),
                           style="Normal.TButton")
            b.pack(side="left", padx=1)
            ToolTip(b, f"Set {bval} beats per measure for the current section")
            self._beats_buttons[bval] = b

        # Scrollable canvas for measure columns.
        # macOS Aqua overrides ALL tk/ttk Scrollbar colours unconditionally,
        # making thumbs invisible on dark backgrounds regardless of bg/troughcolor.
        # Solution: replace tk.Scrollbar with two small tk.Canvas widgets that
        # we draw the track and thumb on ourselves — fully theme-aware.

        self.canvas_outer = tk.Frame(parent)
        self.canvas_outer.pack(fill="both", expand=True, padx=8, pady=(0, 4))

        # Custom horizontal indicator (10 px tall strip at the bottom)
        self.hbar = tk.Canvas(self.canvas_outer, height=10, highlightthickness=0)
        self.hbar.pack(side="bottom", fill="x", pady=(2, 0))

        # Custom vertical indicator (10 px wide strip on the right)
        self.vbar = tk.Canvas(self.canvas_outer, width=10, highlightthickness=0)
        self.vbar.pack(side="right", fill="y", padx=(2, 0))

        # Main scrollable canvas
        self.canvas = tk.Canvas(self.canvas_outer, highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)

        self.measure_frame = tk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.measure_frame, anchor="nw")
        self.measure_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>",        self._on_canvas_configure)

        # Track scroll position fractions (updated via canvas scroll callbacks)
        self._xfrac = (0.0, 1.0)
        self._yfrac = (0.0, 1.0)

        def _xscroll_cb(first, last):
            self._xfrac = (float(first), float(last))
            self._draw_hbar()

        def _yscroll_cb(first, last):
            self._yfrac = (float(first), float(last))
            self._draw_vbar()

        self.canvas.configure(xscrollcommand=_xscroll_cb,
                               yscrollcommand=_yscroll_cb)

        self.hbar.bind("<Configure>", lambda e: self._draw_hbar())
        self.vbar.bind("<Configure>", lambda e: self._draw_vbar())

        # Click / drag on the indicator bars to scroll
        def _hbar_seek(event):
            w = self.hbar.winfo_width()
            if w <= 0: return
            span  = self._xfrac[1] - self._xfrac[0]
            start = max(0.0, min(event.x / w - span / 2, 1.0 - span))
            self.canvas.xview_moveto(start)

        def _vbar_seek(event):
            h = self.vbar.winfo_height()
            if h <= 0: return
            span  = self._yfrac[1] - self._yfrac[0]
            start = max(0.0, min(event.y / h - span / 2, 1.0 - span))
            self.canvas.yview_moveto(start)

        self.hbar.bind("<Button-1>",  _hbar_seek)
        self.hbar.bind("<B1-Motion>", _hbar_seek)
        self.vbar.bind("<Button-1>",  _vbar_seek)
        self.vbar.bind("<B1-Motion>", _vbar_seek)

        # Mousewheel bindings — bound to every relevant widget so scrolling
        # works wherever the mouse happens to be
        def _scroll_y(event):
            delta = -1 * (event.delta // abs(event.delta)) if event.delta else 0
            self.canvas.yview_scroll(delta, "units")
            return "break"

        def _scroll_x(event):
            delta = -1 * (event.delta // abs(event.delta)) if event.delta else 0
            self.canvas.xview_scroll(delta, "units")
            return "break"

        def _scroll_linux_y(event):
            if event.num == 4:   self.canvas.yview_scroll(-1, "units")
            elif event.num == 5: self.canvas.yview_scroll( 1, "units")

        def _scroll_linux_x(event):
            if event.num == 6:   self.canvas.xview_scroll(-1, "units")
            elif event.num == 7: self.canvas.xview_scroll( 1, "units")

        for w in (self.canvas, self.canvas_outer, self.hbar, self.vbar):
            w.bind("<MouseWheel>",       _scroll_y)
            w.bind("<Shift-MouseWheel>", _scroll_x)
            w.bind("<Button-4>",  _scroll_linux_y)
            w.bind("<Button-5>",  _scroll_linux_y)
            w.bind("<Button-6>",  _scroll_linux_x)
            w.bind("<Button-7>",  _scroll_linux_x)

    # ==========================================================================
    #  THEME ENGINE  (ported from QLC+ Swiss Knife v0.4)
    # ==========================================================================

    def toggle_theme(self):
        """Flip between dark and light, reapply styles, reload active section."""
        self.current_theme = "light" if self.current_theme == "dark" else "dark"
        self.apply_theme()
        if self.current_section_idx is not None:
            self._load_section(self.current_section_idx)

    def apply_theme(self):
        """
        Push the current theme palette to:
          1. ttk.Style  (buttons, combobox, scrollbar)
          2. Named structural tk widgets
          3. Everything else via _apply_theme_recursive
        """
        t = THEMES[self.current_theme]
        self.configure(bg=t["bg"])

        s = self.style
        s.configure(".", background=t["bg"], foreground=t["fg"])

        # ── Normal.TButton — secondary actions (Edit, Move, Theme, Drop D …) ──
        # FIX v0.5: On macOS, ttk.Style.configure() foreground is ignored for
        # buttons at rest. The only reliable fix is to include ("!disabled", fg)
        # in the s.map() foreground list, which macOS Aqua DOES honour.
        s.configure("Normal.TButton",
                    background=t["btn_bg"],
                    foreground=t["fg"],
                    font=FONT_TINY, padding=4)
        s.map("Normal.TButton",
              background=[("active",    t["select_bg"]),
                           ("pressed",   t["select_bg"])],
              foreground=[("!disabled", t["fg"]),      # ← macOS fix: rest state
                           ("active",    t["select_fg"]),
                           ("pressed",   t["select_fg"])])

        # ── Accent.TButton — primary actions (Save, Export, +, OK …) ──────────
        s.configure("Accent.TButton",
                    background=t["accent"],
                    foreground="#ffffff",
                    font=("Courier New", 9, "bold"), padding=4)
        s.map("Accent.TButton",
              background=[("active",    t["select_bg"]),
                           ("pressed",   t["select_bg"])],
              foreground=[("!disabled", "#ffffff"),    # ← macOS fix
                           ("active",    t["select_fg"]),
                           ("pressed",   t["select_fg"])])

        # ── Per-layer toggle button styles (NEW v0.6) ─────────────────────────
        # Plain tk.Button is overridden by macOS Aqua and ignores bg/fg.
        # Solution: use ttk.Button with a named style per layer so the colour
        # is applied through ttk (which Aqua does respect for the clam theme).
        for layer in ("tab", "chords", "notes", "lyrics"):
            style_name = f"{layer.capitalize()}.TButton"
            active_col = LAYER_COLORS[layer][self.current_theme]
            inactive_col = t["btn_bg"]
            # We set both the active and inactive colours here; _load_section
            # picks the right style name depending on s.visible[layer].
            s.configure(f"Active{style_name}",
                        background=active_col, foreground="#ffffff",
                        font=FONT_TINY, padding=4)
            s.map(f"Active{style_name}",
                  background=[("active",    t["select_bg"]),
                               ("pressed",   t["select_bg"]),
                               ("!disabled", active_col)],
                  foreground=[("!disabled", "#ffffff"),
                               ("active",    t["select_fg"]),
                               ("pressed",   t["select_fg"])])
            s.configure(f"Inactive{style_name}",
                        background=inactive_col, foreground=t["fg"],
                        font=FONT_TINY, padding=4)
            s.map(f"Inactive{style_name}",
                  background=[("active",    t["select_bg"]),
                               ("pressed",   t["select_bg"]),
                               ("!disabled", inactive_col)],
                  foreground=[("!disabled", t["fg"]),
                               ("active",    t["select_fg"]),
                               ("pressed",   t["select_fg"])])

        # ── TCombobox ─────────────────────────────────────────────────────────
        s.configure("TCombobox",
                    fieldbackground=t["input_bg"], background=t["btn_bg"],
                    foreground=t["fg"], arrowcolor=t["fg"],
                    selectbackground=t["select_bg"],
                    selectforeground=t["select_fg"])
        s.map("TCombobox",
              fieldbackground=[("readonly", t["input_bg"])],
              foreground=[("readonly", t["fg"]),
                           ("!disabled", t["fg"])],   # macOS fix
              selectbackground=[("readonly", t["select_bg"])],
              selectforeground=[("readonly", t["select_fg"])])

        # The combobox dropdown list is a plain tk.Listbox — needs option_add
        self.option_add("*TCombobox*Listbox.background",       t["input_bg"])
        self.option_add("*TCombobox*Listbox.foreground",       t["fg"])
        self.option_add("*TCombobox*Listbox.selectBackground", t["select_bg"])
        self.option_add("*TCombobox*Listbox.selectForeground", t["select_fg"])

        s.configure("TScrollbar",
                    background=t["select_bg"],
                    troughcolor=t["surface"],
                    arrowcolor=t["fg"],
                    borderwidth=0, relief="flat")

        # ── Custom scroll indicator canvases (fully theme-aware, macOS-proof) ──
        try:
            self.hbar.configure(bg=t["surface"])
            self.vbar.configure(bg=t["surface"])
            self._draw_hbar()
            self._draw_vbar()
        except AttributeError:
            pass   # called before _build_editor_panel — safe to skip

        # Canvas outer frame and canvas itself
        self.canvas_outer.configure(bg=t["bg"])
        self.canvas.configure(bg=t["bg"])

        # ── Named structural widgets ──────────────────────────────────────────
        self.topbar.configure(bg=t["header_bg"])
        self.topbar2.configure(bg=t["header_bg"])
        self.sep.configure(bg=t["border"])
        # PanedWindow bg IS the sash colour — make it accent so it's easy to grab
        self.pane.configure(bg=t["select_bg"])
        self.left_panel.configure(bg=t["card_bg"])
        self.sec_hdr.configure(bg=t["card_bg"])
        self.sec_btn_frame.configure(bg=t["card_bg"])
        self.right_panel.configure(bg=t["bg"])
        self.toggle_frame.configure(bg=t["bg"])
        # toolbar is the parent of toggle_frame, btn_wrap, beats selector
        try:
            self.toggle_frame.master.configure(bg=t["bg"])
            self._beats_lbl.configure(bg=t["bg"], fg=t["lbl_gray"])
            self._beats_btn_frame.configure(bg=t["bg"])
        except Exception:
            pass
        self.editor_title.configure(bg=t["bg"], fg=t["fg"])
        self.tab_hint.configure(bg=t["bg"], fg=t["lbl_gray"])
        self.measure_frame.configure(bg=t["bg"])
        self.lbl_app_title.configure(bg=t["header_bg"], fg=t["accent"])
        self.lbl_sections.configure(bg=t["card_bg"],    fg=t["accent"])

        for lbl in self._topbar_labels:
            lbl.configure(bg=t["header_bg"], fg=t["lbl_gray"])
        for e in self._topbar_entries:
            e.configure(bg=t["input_bg"], fg=t["fg"], insertbackground=t["fg"])

        self.section_listbox.configure(
            bg=t["input_bg"], fg=t["fg"],
            selectbackground=t["select_bg"],
            selectforeground=t["select_fg"])

        self._retheme_toggle_buttons(t)
        self._apply_theme_recursive(self, t)

    def _retheme_toggle_buttons(self, t):
        """
        Re-colour the layer visibility toggle buttons after a theme switch.
        v0.6: uses ttk style swapping (ActiveXxx / InactiveXxx) instead of
        setting bg/fg directly, which macOS Aqua ignores on tk.Button.
        """
        idx = self.current_section_idx
        if idx is None or idx >= len(self.sections):
            return
        s = self.sections[idx]
        for layer, btn in self._toggle_buttons.items():
            active     = s.visible[layer]
            style_name = (f"Active{layer.capitalize()}.TButton"
                          if active else
                          f"Inactive{layer.capitalize()}.TButton")
            try:
                btn.configure(style=style_name)
            except tk.TclError:
                pass

    def _apply_theme_recursive(self, widget, t):
        """
        Walk the widget tree and recolour plain tk widgets not handled above.
        Uses exclusion sets to protect specially-coloured named widgets.
        """
        _named   = {self.lbl_app_title, self.lbl_sections,
                    self.editor_title,  self.section_listbox, self.tab_hint}
        _card_bg = {self.left_panel, self.sec_hdr, self.sec_btn_frame}
        # canvas_outer, hbar, vbar are handled explicitly above
        _skip_children = {self.canvas_outer, self.hbar, self.vbar}

        wtype = widget.winfo_class()
        try:
            if widget in _named:
                pass
            elif widget in _skip_children:
                pass   # handled explicitly — don't recurse into scrollbars
            elif wtype in ("Frame", "Tk", "Toplevel"):
                widget.configure(bg=t["card_bg"] if widget in _card_bg
                                  else t["bg"])
            elif wtype == "Label":
                widget.configure(bg=t["bg"], fg=t["fg"])
            elif wtype in ("Labelframe", "LabelFrame"):
                widget.configure(bg=t["bg"], fg=t["fg"])
            elif wtype == "PanedWindow":
                widget.configure(bg=t["border"])
        except tk.TclError:
            pass
        if widget not in _skip_children:
            for child in widget.winfo_children():
                self._apply_theme_recursive(child, t)

    # ==========================================================================
    #  DROP D / TRANSPOSE
    # ==========================================================================

    def _apply_drop_d(self):
        """Toggle the current section between standard and Drop D tuning."""
        idx = self.current_section_idx
        if idx is None:
            messagebox.showinfo("Drop D", "Select a section first.")
            return
        s = self.sections[idx]
        drop_d_map = {
            "Guitar (6-string)":        "Guitar (6-string) Drop D",
            "Guitar (6-string) Drop D": "Guitar (6-string)",
            "Bass (4-string)":          "Bass (4-string) Drop D",
            "Bass (4-string) Drop D":   "Bass (4-string)",
        }
        target = drop_d_map.get(s.instrument)
        if target is None:
            messagebox.showinfo("Drop D",
                                f"No Drop D variant for '{s.instrument}'.")
            return
        self._save_current_section()
        s.instrument = target
        self._refresh_listbox()
        self._load_section(idx)
        self.dirty = True

    def _transpose_dialog(self):
        """
        Open a dialog to shift the current section by ±N semitones.
        Affects: tab fret numbers, notes layer, chords layer.
        """
        idx = self.current_section_idx
        if idx is None:
            messagebox.showinfo("Transpose", "Select a section first.")
            return
        self._save_current_section()

        t   = THEMES[self.current_theme]
        dlg = tk.Toplevel(self)
        dlg.title("Transpose Section")
        dlg.configure(bg=t["bg"])
        dlg.resizable(False, False)
        dlg.grab_set()

        tk.Label(dlg,
                 text="Semitones to shift  (negative = down):",
                 bg=t["bg"], fg=t["lbl_gray"], font=FONT_TINY
                 ).pack(padx=20, pady=(14, 2), anchor="w")

        semitones_var = tk.StringVar(value="0")
        e = tk.Entry(dlg, textvariable=semitones_var, width=6,
                     bg=t["input_bg"], fg=t["fg"],
                     insertbackground=t["fg"], relief="flat", font=FONT_MAIN)
        e.pack(padx=20, pady=(0, 8))
        e.focus_set(); e.select_range(0, "end")

        # Checkboxes: which layers to transpose
        chk_frame = tk.Frame(dlg, bg=t["bg"])
        chk_frame.pack(fill="x", padx=20, pady=(0, 12))
        tk.Label(chk_frame, text="Apply to:",
                 bg=t["bg"], fg=t["lbl_gray"], font=FONT_TINY).pack(side="left")
        var_tab    = tk.BooleanVar(value=True)
        var_chords = tk.BooleanVar(value=True)
        var_notes  = tk.BooleanVar(value=True)
        for text, var in [("tab", var_tab), ("chords", var_chords), ("notes", var_notes)]:
            tk.Checkbutton(chk_frame, text=text, variable=var,
                           bg=t["bg"], fg=t["fg"],
                           selectcolor=t["input_bg"],
                           activebackground=t["bg"],
                           font=FONT_TINY).pack(side="left", padx=6)

        def apply_transpose():
            try:
                shift = int(semitones_var.get())
            except ValueError:
                messagebox.showerror("Transpose", "Please enter a whole number.")
                return
            if shift == 0:
                dlg.destroy(); return

            s = self.sections[idx]

            # ── Tab layer: smart transpose with string reassignment ──────────
            if var_tab.get():
                s.layers["tab"] = [
                    transpose_tab_cell(cell, shift, s.instrument)
                    for cell in s.layers["tab"]
                ]

            # ── Notes layer: transpose each note name ─────────────────────────
            if var_notes.get():
                s.layers["notes"] = [
                    transpose_note(v, shift) if v.strip() else v
                    for v in s.layers["notes"]
                ]

            # ── Chords layer: transpose root note, keep suffix ────────────────
            if var_chords.get():
                s.layers["chords"] = [
                    transpose_chord(v, shift) if v.strip() else v
                    for v in s.layers["chords"]
                ]

            self.dirty = True
            dlg.destroy()
            self._section_widgets = {}
            self._load_section(idx)

        br = tk.Frame(dlg, bg=t["bg"])
        br.pack(fill="x", padx=20, pady=(0, 14))
        ttk.Button(br, text="Cancel", command=dlg.destroy,
                   style="Normal.TButton").pack(side="right", padx=4)
        ttk.Button(br, text="Apply", command=apply_transpose,
                   style="Accent.TButton").pack(side="right")
        dlg.bind("<Return>", lambda e: apply_transpose())

    # ==========================================================================
    #  SECTION MANAGEMENT
    # ==========================================================================

    def _add_section_dialog(self):
        self._section_dialog()

    def _edit_section_dialog(self):
        idx = self._selected_idx()
        if idx is not None:
            self._section_dialog(edit_idx=idx)

    def _copy_section_dialog(self):
        """Clone any existing section and insert the copy after the current one."""
        if not self.sections:
            messagebox.showinfo("Copy Section", "No sections to copy yet.")
            return
        self._save_current_section()

        t   = THEMES[self.current_theme]
        dlg = tk.Toplevel(self)
        dlg.title("Copy Section")
        dlg.configure(bg=t["bg"])
        dlg.resizable(False, False)
        dlg.grab_set()

        tk.Label(dlg, text="Copy from:",
                 bg=t["bg"], fg=t["lbl_gray"], font=FONT_TINY
                 ).pack(padx=16, pady=(14, 2), anchor="w")
        names      = [s.name for s in self.sections]
        source_var = tk.StringVar()
        cb = ttk.Combobox(dlg, textvariable=source_var, values=names,
                          state="readonly", font=FONT_MAIN, width=26)
        cb.pack(padx=16, pady=(0, 8))
        cb.current(self.current_section_idx if self.current_section_idx is not None
                   else 0)

        tk.Label(dlg, text="New section name (optional):",
                 bg=t["bg"], fg=t["lbl_gray"], font=FONT_TINY
                 ).pack(padx=16, pady=(4, 2), anchor="w")
        name_var = tk.StringVar()
        tk.Entry(dlg, textvariable=name_var,
                 bg=t["input_bg"], fg=t["fg"],
                 insertbackground=t["fg"],
                 relief="flat", font=FONT_MAIN, width=26
                 ).pack(padx=16, pady=(0, 12))

        def confirm():
            src_name = source_var.get()
            if not src_name:
                messagebox.showwarning("Copy Section",
                                       "Please select a source section.")
                return
            src_idx  = names.index(src_name)
            clone      = copy.deepcopy(self.sections[src_idx])
            clone.name = name_var.get().strip() or f"{clone.name} (copy)"
            insert_at  = (self.current_section_idx + 1
                          if self.current_section_idx is not None
                          else len(self.sections))
            self.sections.insert(insert_at, clone)
            self._refresh_listbox()
            self.section_listbox.selection_clear(0, "end")
            self.section_listbox.selection_set(insert_at)
            self._load_section(insert_at)
            self.dirty = True
            dlg.destroy()

        br = tk.Frame(dlg, bg=t["bg"])
        br.pack(fill="x", padx=16, pady=(0, 14))
        ttk.Button(br, text="Cancel", command=dlg.destroy,
                   style="Normal.TButton").pack(side="right", padx=4)
        ttk.Button(br, text="  Copy  ", command=confirm,
                   style="Accent.TButton").pack(side="right")
        dlg.bind("<Return>", lambda e: confirm())

    def _section_dialog(self, edit_idx=None):
        """Shared dialog for creating (edit_idx=None) or editing a section."""
        t       = THEMES[self.current_theme]
        dlg     = tk.Toplevel(self)
        dlg.title("Edit Section" if edit_idx is not None else "New Section")
        dlg.configure(bg=t["bg"])
        dlg.resizable(False, False)
        dlg.grab_set()
        editing = self.sections[edit_idx] if edit_idx is not None else None

        def row(label, widget_fn):
            f = tk.Frame(dlg, bg=t["bg"])
            f.pack(fill="x", padx=16, pady=5)
            tk.Label(f, text=label, bg=t["bg"], fg=t["lbl_gray"],
                     font=FONT_TINY, width=14, anchor="w").pack(side="left")
            w = widget_fn(f); w.pack(side="left", fill="x", expand=True)
            return w

        type_var = tk.StringVar(value=editing.section_type if editing else "Verse")
        row("Type:", lambda f: ttk.Combobox(
            f, textvariable=type_var, values=SECTION_TYPES,
            state="readonly", font=FONT_MAIN, width=16))

        name_var = tk.StringVar(value=editing.name if editing else "")
        row("Name (optional):", lambda f: tk.Entry(
            f, textvariable=name_var, bg=t["input_bg"], fg=t["fg"],
            insertbackground=t["fg"], relief="flat", font=FONT_MAIN, width=20))

        meas_var = tk.StringVar(value=str(editing.measures if editing else 4))
        row("Measures:", lambda f: tk.Entry(
            f, textvariable=meas_var, bg=t["input_bg"], fg=t["fg"],
            insertbackground=t["fg"], relief="flat", font=FONT_MAIN, width=6))

        rep_var = tk.StringVar(value=str(editing.repeat if editing else 1))
        row("Repeat (×):", lambda f: tk.Entry(
            f, textvariable=rep_var, bg=t["input_bg"], fg=t["fg"],
            insertbackground=t["fg"], relief="flat", font=FONT_MAIN, width=4))

        instr_var = tk.StringVar(
            value=editing.instrument if editing else self.instrument.get())
        row("Instrument:", lambda f: ttk.Combobox(
            f, textvariable=instr_var,
            values=list(INSTRUMENT_STRINGS.keys()),
            state="readonly", font=FONT_MAIN, width=22))

        beats_var = tk.StringVar(
            value=str(editing.tab_beats if editing else TAB_BEATS_DEFAULT))
        row("Tab beats/measure:", lambda f: ttk.Combobox(
            f, textvariable=beats_var,
            values=[str(b) for b in TAB_BEATS_OPTIONS],
            state="readonly", font=FONT_MAIN, width=6))

        lyr_frame = tk.Frame(dlg, bg=t["bg"])
        lyr_frame.pack(fill="x", padx=16, pady=6)
        tk.Label(lyr_frame, text="Show layers:",
                 bg=t["bg"], fg=t["lbl_gray"], font=FONT_TINY).pack(side="left")
        lyr_vars = {}
        for layer in ("tab", "chords", "notes", "lyrics"):
            default = (editing.visible.get(layer, layer in ("tab", "chords"))
                       if editing else layer in ("tab", "chords"))
            v = tk.BooleanVar(value=default); lyr_vars[layer] = v
            tk.Checkbutton(lyr_frame, text=layer, variable=v,
                           bg=t["bg"], fg=t["fg"],
                           selectcolor=t["input_bg"],
                           activebackground=t["bg"],
                           font=FONT_TINY).pack(side="left", padx=4)

        def confirm():
            try:
                m = int(meas_var.get()); r = int(rep_var.get())
                assert m > 0 and r > 0
            except Exception:
                messagebox.showerror(
                    "Error", "Measures and Repeat must be positive integers.")
                return
            sec_type  = type_var.get()
            display   = name_var.get().strip() or sec_type
            instr     = instr_var.get()
            new_beats = int(beats_var.get())

            if edit_idx is not None:
                s = self.sections[edit_idx]
                s.name = display; s.section_type = sec_type
                s.instrument = instr; s.repeat = r; s.tab_beats = new_beats
                s.resize(m)
                for layer, v in lyr_vars.items():
                    s.visible[layer] = v.get()
                self._section_widgets = {}
                self._refresh_listbox()
                self._load_section(edit_idx)
            else:
                s = Section(display, sec_type, m, instr)
                s.repeat = r; s.tab_beats = new_beats
                for layer, v in lyr_vars.items():
                    s.visible[layer] = v.get()
                self.sections.append(s); self._refresh_listbox()
                new_idx = len(self.sections) - 1
                self.section_listbox.selection_clear(0, "end")
                self.section_listbox.selection_set(new_idx)
                self._load_section(new_idx)
            self.dirty = True; dlg.destroy()

        br = tk.Frame(dlg, bg=t["bg"])
        br.pack(fill="x", padx=16, pady=(8, 14))
        ttk.Button(br, text="Cancel", command=dlg.destroy,
                   style="Normal.TButton").pack(side="right", padx=4)
        ttk.Button(br, text="  OK  ", command=confirm,
                   style="Accent.TButton").pack(side="right")
        dlg.bind("<Return>", lambda e: confirm())

    def _refresh_listbox(self):
        """Rebuild the section listbox from self.sections."""
        self.section_listbox.delete(0, "end")
        for s in self.sections:
            rep  = f" ×{s.repeat}" if s.repeat > 1 else ""
            link = f" 🔗{s.link_id}" if s.link_id is not None else ""
            self.section_listbox.insert(
                "end", f"  {s.name}  [{s.measures}m]{rep}{link}")

    def _selected_idx(self):
        sel = self.section_listbox.curselection()
        return sel[0] if sel else None

    def _on_section_select(self, event=None):
        idx = self._selected_idx()
        if idx is not None:
            self._save_current_section()
            self._load_section(idx)

    def _move_section_up(self):
        idx = self._selected_idx()
        if idx is None or idx == 0: return
        self._save_current_section()
        self.sections[idx-1], self.sections[idx] = \
            self.sections[idx], self.sections[idx-1]
        self._refresh_listbox()
        self.section_listbox.selection_set(idx-1)
        self._load_section(idx-1)

    def _move_section_down(self):
        idx = self._selected_idx()
        if idx is None or idx >= len(self.sections)-1: return
        self._save_current_section()
        self.sections[idx], self.sections[idx+1] = \
            self.sections[idx+1], self.sections[idx]
        self._refresh_listbox()
        self.section_listbox.selection_set(idx+1)
        self._load_section(idx+1)

    def _delete_section(self):
        idx = self._selected_idx()
        if idx is None: return
        if not messagebox.askyesno(
                "Delete", f"Delete section '{self.sections[idx].name}'?"):
            return
        # If this section was in a link group, clean up orphaned singletons
        old_lid = self.sections[idx].link_id
        self.sections.pop(idx)
        if old_lid is not None:
            remaining = [s for s in self.sections if s.link_id == old_lid]
            if len(remaining) == 1:
                remaining[0].link_id = None   # a group of 1 is not a group
        self._refresh_listbox()
        self._clear_editor()
        self.dirty = True

    # ==========================================================================
    #  EDITOR — measure grid
    # ==========================================================================

    def _clear_editor(self):
        """Reset the editor to the welcome state."""
        self.current_section_idx = None
        t = THEMES[self.current_theme]
        self.editor_title.config(text="← Add a section to begin",
                                  fg=t["lbl_gray"])
        for w in self.toggle_frame.winfo_children(): w.destroy()
        for w in self.measure_frame.winfo_children(): w.destroy()
        self._section_widgets = {}
        self._toggle_buttons  = {}
        self._measure_cards   = []
        self._update_beats_buttons()

    def _load_section(self, idx):
        """
        Populate the editor with the section at self.sections[idx].

        Tab layer (NEW in v0.5):
          Each measure card gets one tk.Entry per instrument string, laid out
          vertically inside the card.  The entries are pre-filled with dash
          tokens ("- - - - - - - -") and labelled with the string name on the
          left.  The user replaces individual dashes with fret numbers.
        """
        self._save_current_section()
        self.current_section_idx = idx
        s = self.sections[idx]
        t = THEMES[self.current_theme]

        rep = f"  ×{s.repeat}" if s.repeat > 1 else ""
        self.editor_title.config(
            text=f"[ {s.name} ]   {s.measures} measures   {s.instrument}{rep}",
            fg=t["fg"])

        # ── Rebuild layer toggle buttons (v0.6: ttk.Button with named styles) ──
        for w in self.toggle_frame.winfo_children(): w.destroy()
        self._toggle_buttons = {}
        tk.Label(self.toggle_frame, text="Layers:",
                 bg=t["bg"], fg=t["lbl_gray"], font=FONT_TINY).pack(side="left")
        for layer in ("tab", "chords", "notes", "lyrics"):
            active     = s.visible[layer]
            style_name = (f"Active{layer.capitalize()}.TButton"
                          if active else
                          f"Inactive{layer.capitalize()}.TButton")
            btn = ttk.Button(
                self.toggle_frame, text=layer,
                style=style_name,
                command=lambda l=layer: self._toggle_layer(l))
            btn.pack(side="left", padx=3)
            self._toggle_buttons[layer] = btn

        # ── Rebuild measure columns ───────────────────────────────────────────
        for w in self.measure_frame.winfo_children(): w.destroy()
        self._section_widgets = {}
        strings = INSTRUMENT_STRINGS.get(s.instrument, ["e","B","G","D","A","E"])

        # Build all measure card frames first, then place them
        # (placement is deferred to _place_measures which handles both modes)
        self._measure_cards = []   # list of border frames, in order

        for m_idx in range(s.measures):
            bf = tk.Frame(self.measure_frame, bg=t["border"], padx=1, pady=1)
            self._measure_cards.append(bf)
            inner = tk.Frame(bf, bg=t["card_bg"])
            inner.pack(fill="both", expand=True)

            tk.Label(inner, text=f"M{m_idx+1}",
                     bg=t["card_bg"], fg=t["accent"],
                     font=FONT_TINY).pack(pady=(4, 0))

            # ── Per-measure beats picker (collapsed by default) ───────────────
            # Current effective beats for this measure
            m_beats = self._get_measure_beats(s, m_idx)
            # Show current beats in small dim text; click ▸ to expand picker
            beats_row = tk.Frame(inner, bg=t["card_bg"])
            beats_row.pack(fill="x", padx=4, pady=(0, 2))

            # Dim label showing current value
            beats_lbl = tk.Label(beats_row,
                                  text=f"{m_beats}b",
                                  bg=t["card_bg"], fg=t["lbl_gray"],
                                  font=("Courier New", 7))
            beats_lbl.pack(side="left")

            # Expand toggle ▸/▾
            expand_var  = tk.BooleanVar(value=False)
            picker_frame = tk.Frame(inner, bg=t["card_bg"])
            # (picker_frame not packed yet — shown on expand)

            def _toggle_beats_picker(pf=picker_frame, bv=expand_var,
                                      bl=beats_lbl, mi=m_idx,
                                      _s=s, _t=t):
                bv.set(not bv.get())
                if bv.get():
                    # Build picker buttons inside pf
                    for w in pf.winfo_children():
                        w.destroy()
                    for bval in TAB_BEATS_OPTIONS:
                        is_active = self._get_measure_beats(_s, mi) == bval
                        style = "Accent.TButton" if is_active else "Normal.TButton"

                        def _set_m_beats(v=bval, pf2=pf, bv2=bv,
                                          bl2=bl, mi2=mi, _s2=_s):
                            self._save_current_section()
                            if v == _s2.tab_beats:
                                # Reset to section default
                                _s2.measure_beats.pop(mi2, None)
                            else:
                                _s2.measure_beats[mi2] = v
                            self.dirty = True
                            self._section_widgets = {}
                            self._load_section(self.current_section_idx)

                        btn = ttk.Button(pf, text=str(bval), width=3,
                                         command=_set_m_beats,
                                         style=style)
                        btn.pack(side="left", padx=1)
                    pf.pack(fill="x", padx=4, pady=(0, 2))
                else:
                    pf.pack_forget()

            expand_btn = ttk.Button(beats_row, text="▸", width=2,
                                     command=_toggle_beats_picker,
                                     style="Normal.TButton")
            expand_btn.pack(side="left", padx=(2, 0))
            ToolTip(expand_btn,
                    f"Override beats for M{m_idx+1} (currently {m_beats})\n"
                    f"Section default: {s.tab_beats}")

            # ── Per-measure copy / paste buttons (NEW v0.6) ───────────────────
            # ⎘ copies this measure's tab data to the internal clipboard.
            # ⏎ pastes the clipboard into this measure (only shown when clipboard
            # has content and is from the same instrument).
            m_btn_row = tk.Frame(inner, bg=t["card_bg"])
            m_btn_row.pack(fill="x", padx=4, pady=(0, 2))

            # Capture m_idx in closure
            def _copy_measure(mi=m_idx):
                """
                Copy this measure's current tab content to the clipboard.
                Reads from the live Entry widgets (not the model) so unsaved
                edits are captured correctly.
                """
                si = self.current_section_idx
                if si is None: return
                sec = self.sections[si]
                # Read directly from the widgets for this measure so we capture
                # whatever the user has typed, even if not yet flushed to model
                tab_widgets = self._section_widgets.get(mi, {}).get("tab", {})
                if tab_widgets:
                    clipboard_data = {sn: e.get() for sn, e in tab_widgets.items()}
                else:
                    # Fall back to model if widgets aren't available
                    clipboard_data = dict(sec.layers["tab"][mi])
                self._measure_clipboard = {
                    "instrument": sec.instrument,
                    "data":       clipboard_data,
                }
                # Reload to show the ⏎ paste buttons on all other measures
                # First flush current state so nothing is lost
                self._save_current_section()
                self._load_section(si)

            def _paste_measure(mi=m_idx):
                """
                Paste the clipboard into this measure's tab layer.
                Saves all other measures first, writes the paste data directly
                into the model for measure mi, then reloads the section.
                The key fix: paste data is written AFTER _save_current_section
                so _load_section's internal save-on-entry doesn't overwrite it.
                """
                if not self._measure_clipboard: return
                si = self.current_section_idx
                if si is None: return
                sec = self.sections[si]
                if self._measure_clipboard.get("instrument") != sec.instrument:
                    messagebox.showwarning(
                        "Paste Measure",
                        "Clipboard was copied from a different instrument.\n"
                        "Change the instrument or copy from a compatible measure.")
                    return
                # Step 1: flush all current widget values to the model
                self._save_current_section()
                # Step 2: overwrite only measure mi in the model with paste data
                sec.layers["tab"][mi] = dict(self._measure_clipboard["data"])
                self.dirty = True
                # Step 3: reload — _load_section calls _save_current_section
                # internally, but since we've already flushed (step 1) and
                # then written the paste (step 2) directly to the model,
                # the internal save in _load_section will just re-save the
                # same values from widgets (which still show the old content
                # for measure mi). To prevent that overwrite, we temporarily
                # remove measure mi from _section_widgets so the internal save
                # skips it.
                self._section_widgets.pop(mi, None)
                self._load_section(si)

            ttk.Button(m_btn_row, text="⎘", width=2,
                       command=_copy_measure,
                       style="Normal.TButton").pack(side="left", padx=(0, 2))
            ToolTip(m_btn_row.winfo_children()[-1],
                    f"Copy M{m_idx+1} tab data to clipboard")

            if self._measure_clipboard:
                pb = ttk.Button(m_btn_row, text="⏎", width=2,
                                command=_paste_measure,
                                style="Normal.TButton")
                pb.pack(side="left")
                ToolTip(pb, f"Paste clipboard into M{m_idx+1}")

            widgets = {}

            for layer in ("tab", "chords", "notes", "lyrics"):
                if not s.visible[layer]:
                    continue

                badge_color = LAYER_COLORS[layer][self.current_theme]

                if layer == "tab":
                    # ── Horizontal tab grid (NEW v0.5) ────────────────────────
                    # One Entry per string, stacked vertically inside the card.
                    # Each Entry is a horizontal row: "- - - 2 - - 5 -"
                    # A coloured "T" badge sits above the string rows.
                    tab_header = tk.Frame(inner, bg=t["card_bg"])
                    tab_header.pack(fill="x", padx=4, pady=(2, 0))
                    tk.Label(tab_header, text="T",
                             bg=badge_color, fg="#ffffff",
                             font=FONT_TINY, width=2,
                             anchor="center").pack(side="left")

                    tab_cell_dict = s.layers["tab"][m_idx]
                    string_entries = {}   # string_name → Entry
                    # Use per-measure beats if set, else section default
                    m_beats_eff = self._get_measure_beats(s, m_idx)

                    for st in strings:
                        row_frame = tk.Frame(inner, bg=t["card_bg"])
                        row_frame.pack(fill="x", padx=4, pady=1)

                        # String name label (e.g. "G")
                        tk.Label(row_frame, text=f"{st}|",
                                 bg=t["card_bg"], fg=t["lbl_gray"],
                                 font=FONT_MONO, width=2,
                                 anchor="e").pack(side="left")

                        # Restore existing content or use blank dashes
                        existing_row = tab_cell_dict.get(st, "")
                        # Pad/trim to this measure's effective beat count
                        if existing_row:
                            toks = [tok for tok in existing_row.split() if tok]
                            while len(toks) < m_beats_eff:
                                toks.append("-")
                            toks = toks[:m_beats_eff]
                            display_val = "  ".join(toks)
                        else:
                            display_val = make_blank_tab_row(m_beats_eff)

                        e = tk.Entry(row_frame,
                                     width=max(24, m_beats_eff * 3 + 2),
                                     bg=t["input_bg"], fg=t["fg"],
                                     insertbackground=t["fg"],
                                     relief="flat", font=FONT_MONO)
                        e.insert(0, display_val)
                        e.pack(side="left", fill="x", expand=True)

                        # ── Protected-dash validator ───────────────────────────
                        # Fires after each keystroke. Preserves token count and
                        # replaces invalid tokens with "-". Does NOT reformat
                        # while the user is mid-typing a number (i.e. if the
                        # last token is a partial digit, leave it alone).
                        _beats = m_beats_eff

                        def _on_key(event, _e=e, _b=_beats):
                            # Skip navigation/modifier keys — they don't change content
                            if event.keysym in ("Left","Right","Home","End",
                                                "Shift_L","Shift_R",
                                                "Control_L","Control_R",
                                                "Alt_L","Alt_R","Tab"):
                                return
                            _e.after(1, lambda: _validate_tab_entry(_e, _b))

                        e.bind("<KeyRelease>", _on_key)

                        string_entries[st] = e

                    widgets["tab"] = string_entries   # dict of {str_name: Entry}

                else:
                    # ── Chord / Notes / Lyrics entry ──────────────────────────
                    lf = tk.Frame(inner, bg=t["card_bg"])
                    lf.pack(fill="x", padx=4, pady=1)
                    tk.Label(lf, text=layer[0].upper(),
                             bg=badge_color, fg="#ffffff",
                             font=FONT_TINY, width=2,
                             anchor="center").pack(side="left")

                    current_val = s.layers[layer][m_idx]

                    if layer in ("chords", "notes"):
                        # ── Root note picker + free suffix ────────────────────
                        # Parse current value into (root, suffix)
                        root_canon, suffix = _parse_root_suffix(current_val)
                        root_disp = NOTE_CANONICAL_TO_DISPLAY.get(
                            root_canon, "") if root_canon else ""

                        note_var   = tk.StringVar(value=root_disp)
                        suffix_var = tk.StringVar(value=suffix)

                        # Root note dropdown (compact width)
                        note_cb = ttk.Combobox(
                            lf, textvariable=note_var,
                            values=[""] + NOTE_DISPLAY,
                            width=7, font=FONT_MONO, state="normal")
                        note_cb.pack(side="left")

                        # Free-text suffix entry (m, dim, 7, maj7, sus4…)
                        sfx_e = tk.Entry(
                            lf, textvariable=suffix_var, width=6,
                            bg=t["input_bg"], fg=t["fg"],
                            insertbackground=t["fg"],
                            relief="flat", font=FONT_MONO)
                        sfx_e.pack(side="left", fill="x", expand=True)

                        # Store both widgets together so _save_current_section
                        # can reconstruct the combined value
                        widgets[layer] = (note_var, suffix_var)

                    else:
                        # Plain entry for lyrics
                        e = tk.Entry(lf, width=20,
                                     bg=t["input_bg"], fg=t["fg"],
                                     insertbackground=t["fg"],
                                     relief="flat", font=FONT_MONO)
                        e.insert(0, current_val)
                        e.pack(side="left", fill="x", expand=True)
                        widgets[layer] = e

            self._section_widgets[m_idx] = widgets

        # Place cards according to current wrap mode.
        # Use after_idle so the canvas has its final width before we calculate.
        self.measure_frame.update_idletasks()
        self._place_measures()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._update_beats_buttons()

    def _place_measures(self):
        """
        (Re-)place all measure cards in the measure_frame grid.

        Scroll mode  (wrap_mode=False):
            All cards in a single row (row=0, column=m_idx).
            The canvas scrolls horizontally.

        Wrap mode  (wrap_mode=True):
            Cards flow left-to-right, wrapping to the next row when they
            would exceed the canvas width.  The canvas scrolls vertically.
            Card width is estimated from the first card's requested width.
        """
        cards = getattr(self, "_measure_cards", [])
        if not cards:
            return

        # Forget all current grid positions
        for c in cards:
            c.grid_forget()

        if not self.wrap_mode.get():
            # ── Scroll mode: single row ───────────────────────────────────────
            for col, c in enumerate(cards):
                c.grid(row=0, column=col, padx=4, pady=4, sticky="n")
        else:
            # ── Wrap mode: reflow into rows ────────────────────────────────────
            # Get available width from the canvas (may be 0 before first render)
            avail = self.canvas.winfo_width()
            if avail < 50:
                # Canvas not yet rendered — fall back to window width estimate
                avail = self.winfo_width() - 300

            # Measure the natural width of a single card
            cards[0].update_idletasks()
            card_w = cards[0].winfo_reqwidth() + 8   # +8 for padx

            cols_per_row = max(1, avail // card_w)

            for i, c in enumerate(cards):
                row = i // cols_per_row
                col = i % cols_per_row
                c.grid(row=row, column=col, padx=4, pady=4, sticky="n")

        self.measure_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _toggle_layer(self, layer):
        """Toggle the visibility of a layer in the current section."""
        idx = self.current_section_idx
        if idx is None: return
        self._save_current_section()
        self.sections[idx].visible[layer] = not self.sections[idx].visible[layer]
        self._load_section(idx)

    def _toggle_wrap(self):
        """Toggle wrap mode and reload the current section."""
        self.wrap_mode.set(not self.wrap_mode.get())
        t = THEMES[self.current_theme]
        if self.wrap_mode.get():
            self.btn_wrap.configure(style="ActiveTab.TButton")
        else:
            self.btn_wrap.configure(style="Normal.TButton")
        if self.current_section_idx is not None:
            self._load_section(self.current_section_idx)

    def _update_beats_buttons(self):
        """Highlight the section-level beats button for the current section."""
        idx = self.current_section_idx
        active_beats = (self.sections[idx].tab_beats
                        if idx is not None and idx < len(self.sections)
                        else TAB_BEATS_DEFAULT)
        for bval, btn in self._beats_buttons.items():
            if bval == active_beats:
                btn.configure(style="Accent.TButton")
            else:
                btn.configure(style="Normal.TButton")

    def _get_measure_beats(self, s, m_idx):
        """Return the effective beat count for measure m_idx in section s."""
        return s.measure_beats.get(m_idx, s.tab_beats)

    # ==========================================================================
    #  SECTION LINKING  (new in v0.7)
    # ==========================================================================

    def _get_next_link_id(self):
        """Return a fresh link group id, incrementing the counter."""
        lid = self._next_link_id
        # Make sure we're above any existing link_id in loaded sections
        existing = [s.link_id for s in self.sections if s.link_id is not None]
        if existing:
            lid = max(lid, max(existing) + 1)
        self._next_link_id = lid + 1
        return lid

    def _propagate_link(self, source_idx):
        """
        Copy all layer data from sections[source_idx] to every other section
        that shares the same link_id, provided they have the same instrument
        and measure count (incompatible sections are silently skipped).
        """
        src = self.sections[source_idx]
        if src.link_id is None:
            return
        import copy as _copy
        for i, s in enumerate(self.sections):
            if i == source_idx:
                continue
            if s.link_id != src.link_id:
                continue
            if s.measures != src.measures or s.instrument != src.instrument:
                continue   # incompatible — skip silently
            # Deep copy each layer so the target is fully independent
            s.layers["tab"]    = _copy.deepcopy(src.layers["tab"])
            s.layers["chords"] = list(src.layers["chords"])
            s.layers["notes"]  = list(src.layers["notes"])
            s.layers["lyrics"] = list(src.layers["lyrics"])

    def _link_section_dialog(self):
        """
        Open the section linking dialog for the currently selected section.

        The dialog shows:
          • The current link status of the selected section
          • Checkboxes for every other compatible section
          • A "source" radio to choose which section's content wins on link
          • Unlink button to remove the selected section from its group
        """
        idx = self._selected_idx()
        if idx is None:
            messagebox.showinfo("Link Sections", "Select a section first.")
            return
        self._save_current_section()

        t   = THEMES[self.current_theme]
        src = self.sections[idx]
        dlg = tk.Toplevel(self)
        dlg.title("Link Sections")
        dlg.configure(bg=t["bg"])
        dlg.resizable(False, False)
        dlg.grab_set()

        # ── Header ────────────────────────────────────────────────────────────
        tk.Label(dlg,
                 text=f"Section:  {src.name}",
                 bg=t["bg"], fg=t["accent"], font=FONT_HEAD
                 ).pack(padx=16, pady=(14, 2), anchor="w")

        status = (f"Currently in link group  🔗{src.link_id}"
                  if src.link_id is not None else "Not linked to any section")
        tk.Label(dlg, text=status,
                 bg=t["bg"], fg=t["lbl_gray"], font=FONT_TINY
                 ).pack(padx=16, pady=(0, 10), anchor="w")

        # ── Section checklist ─────────────────────────────────────────────────
        tk.Label(dlg,
                 text="Link with (✓ = linked, compatible sections only):",
                 bg=t["bg"], fg=t["fg"], font=FONT_TINY
                 ).pack(padx=16, anchor="w")

        scroll_frame = tk.Frame(dlg, bg=t["bg"])
        scroll_frame.pack(fill="x", padx=16, pady=(4, 8))

        check_vars = {}   # other_idx → BooleanVar
        source_var = tk.IntVar(value=idx)   # which section is the copy source

        for i, s in enumerate(self.sections):
            if i == idx:
                continue
            # Compatibility check: same instrument and same measure count
            compatible = (s.measures == src.measures and
                          s.instrument == src.instrument)
            already_linked = (src.link_id is not None and
                              s.link_id == src.link_id)

            row = tk.Frame(scroll_frame, bg=t["bg"])
            row.pack(fill="x", pady=1)

            var = tk.BooleanVar(value=already_linked)
            check_vars[i] = var

            cb = tk.Checkbutton(
                row, text=f"  {s.name}  [{s.measures}m]  {s.instrument}",
                variable=var,
                bg=t["bg"], fg=t["fg"] if compatible else t["lbl_gray"],
                selectcolor=t["input_bg"],
                activebackground=t["bg"],
                font=FONT_TINY,
                state="normal" if compatible else "disabled")
            cb.pack(side="left")

            if not compatible:
                tk.Label(row, text="  (different instrument or measure count)",
                         bg=t["bg"], fg=t["lbl_gray"],
                         font=("Courier New", 7)).pack(side="left")

        # ── Source radio ──────────────────────────────────────────────────────
        sep = tk.Frame(dlg, bg=t["border"], height=1)
        sep.pack(fill="x", padx=16, pady=(4, 8))

        tk.Label(dlg,
                 text="When linking, copy content FROM:",
                 bg=t["bg"], fg=t["fg"], font=FONT_TINY
                 ).pack(padx=16, anchor="w")

        source_scroll = tk.Frame(dlg, bg=t["bg"])
        source_scroll.pack(fill="x", padx=16, pady=(4, 12))

        # The source must be one of: the current section, or a currently-checked peer
        def build_source_radios():
            for w in source_scroll.winfo_children():
                w.destroy()
            # Current section always available as source
            tk.Radiobutton(
                source_scroll,
                text=f"  {src.name}  (this section)",
                variable=source_var, value=idx,
                bg=t["bg"], fg=t["fg"],
                selectcolor=t["input_bg"],
                activebackground=t["bg"],
                font=FONT_TINY).pack(anchor="w")
            for i, var in check_vars.items():
                if var.get():
                    s = self.sections[i]
                    tk.Radiobutton(
                        source_scroll,
                        text=f"  {s.name}",
                        variable=source_var, value=i,
                        bg=t["bg"], fg=t["fg"],
                        selectcolor=t["input_bg"],
                        activebackground=t["bg"],
                        font=FONT_TINY).pack(anchor="w")

        # Rebuild radios whenever a checkbox changes
        for var in check_vars.values():
            var.trace_add("write", lambda *_: build_source_radios())

        build_source_radios()

        # ── Buttons ───────────────────────────────────────────────────────────
        def apply_link():
            linked_indices = [i for i, v in check_vars.items() if v.get()]

            if not linked_indices:
                # No boxes checked — just unlink this section
                src.link_id = None
                self._refresh_listbox()
                self.dirty = True
                dlg.destroy()
                return

            # Determine the group id:
            # Use the source section's existing link_id, or create a new one
            src_idx   = source_var.get()
            src_sec   = self.sections[src_idx]
            group_lid = (src_sec.link_id if src_sec.link_id is not None
                         else self._get_next_link_id())

            # All newly-linked sections + the current section get the group id
            all_in_group = set(linked_indices) | {idx}
            for i in all_in_group:
                self.sections[i].link_id = group_lid

            # Propagate from the chosen source to all others in the group
            self._propagate_link(src_idx)

            self._refresh_listbox()
            self.dirty = True
            dlg.destroy()

            # Reload editor if the current section was affected
            if self.current_section_idx in all_in_group - {src_idx}:
                self._load_section(self.current_section_idx)

        def unlink_all():
            """Remove this section from its link group entirely."""
            if src.link_id is None:
                dlg.destroy()
                return
            old_lid = src.link_id
            src.link_id = None
            # If only one section remains in the old group, unlink it too
            remaining = [s for s in self.sections
                         if s.link_id == old_lid]
            if len(remaining) == 1:
                remaining[0].link_id = None
            self._refresh_listbox()
            self.dirty = True
            dlg.destroy()

        br = tk.Frame(dlg, bg=t["bg"])
        br.pack(fill="x", padx=16, pady=(0, 14))

        ttk.Button(br, text="Cancel",    command=dlg.destroy,
                   style="Normal.TButton").pack(side="right", padx=4)
        ttk.Button(br, text="✓ Apply",   command=apply_link,
                   style="Accent.TButton").pack(side="right")
        if src.link_id is not None:
            ttk.Button(br, text="🔓 Unlink this section", command=unlink_all,
                       style="Normal.TButton").pack(side="left")

        dlg.bind("<Return>", lambda e: apply_link())

    def _save_current_section(self):
        """
        Read all widget values back into the Section data model.
        Tab layer: reads each per-string Entry into a dict {string_name: text}.
        Other layers: reads the single Entry value.
        After saving, propagates layer data to all sections in the same link group.
        Must be called before switching sections, saving, or exporting.
        """
        idx = self.current_section_idx
        if idx is None or idx >= len(self.sections): return
        s = self.sections[idx]

        for m_idx, widgets in self._section_widgets.items():
            # Guard: model may have been resized before widgets were rebuilt.
            for layer, w in widgets.items():
                if layer == "tab":
                    if m_idx >= len(s.layers["tab"]): continue
                    cell_dict = {sn: entry.get()
                                 for sn, entry in w.items()}
                    s.layers["tab"][m_idx] = cell_dict
                elif isinstance(w, tuple):
                    # Chord / notes picker: (note_var, suffix_var)
                    if m_idx >= len(s.layers[layer]): continue
                    note_var, suffix_var = w
                    disp   = note_var.get().strip()
                    canon  = NOTE_DISPLAY_TO_CANONICAL.get(disp, disp)
                    suffix = suffix_var.get().strip()
                    # Store combined value e.g. "F#" + "m7" → "F#m7"
                    s.layers[layer][m_idx] = (canon + suffix) if canon else suffix
                else:
                    if m_idx >= len(s.layers[layer]): continue
                    s.layers[layer][m_idx] = w.get()

        # ── Propagate to linked sections ──────────────────────────────────────
        if s.link_id is not None:
            self._propagate_link(idx)

    # ── Canvas scroll helpers ─────────────────────────────────────────────────

    def _on_frame_configure(self, event=None):
        """Update the canvas scroll region whenever measure_frame changes size."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event=None):
        """
        Called when the canvas widget is resized (e.g. by dragging the sash).
        In scroll mode: keep inner frame at least as wide as canvas.
        In wrap mode: re-flow measure cards to the new width.
        """
        if self.wrap_mode.get():
            # Re-flow cards at new width, then let canvas size to content
            self.after_idle(self._place_measures)
        else:
            natural_w = self.measure_frame.winfo_reqwidth()
            canvas_w  = event.width
            if canvas_w > natural_w:
                self.canvas.itemconfig(self.canvas_window, width=canvas_w)
            else:
                self.canvas.itemconfig(self.canvas_window, width=natural_w)

    def _draw_hbar(self):
        """
        Draw the custom horizontal scroll indicator.
        Track = full width of hbar canvas; thumb = proportional filled rect.
        Only drawn when there is actually something to scroll (xfrac span < 1).
        """
        t = THEMES[self.current_theme]
        c = self.hbar
        c.delete("all")
        w = c.winfo_width()
        h = c.winfo_height()
        if w <= 1: return
        # Track
        c.create_rectangle(0, 0, w, h, fill=t["surface"], outline="", width=0)
        first, last = self._xfrac
        if last - first >= 0.999:
            return   # nothing to scroll — hide thumb
        # Thumb
        x0 = int(first * w) + 1
        x1 = int(last  * w) - 1
        c.create_rectangle(x0, 1, x1, h - 1,
                            fill=t["select_bg"], outline="", width=0)

    def _draw_vbar(self):
        """
        Draw the custom vertical scroll indicator.
        Track = full height of vbar canvas; thumb = proportional filled rect.
        """
        t = THEMES[self.current_theme]
        c = self.vbar
        c.delete("all")
        w = c.winfo_width()
        h = c.winfo_height()
        if h <= 1: return
        # Track
        c.create_rectangle(0, 0, w, h, fill=t["surface"], outline="", width=0)
        first, last = self._yfrac
        if last - first >= 0.999:
            return   # nothing to scroll — hide thumb
        # Thumb
        y0 = int(first * h) + 1
        y1 = int(last  * h) - 1
        c.create_rectangle(1, y0, w - 1, y1,
                            fill=t["select_bg"], outline="", width=0)

    # ==========================================================================
    #  FILE I/O — Export TXT, Save .sng, Open .sng
    # ==========================================================================

    def _export_txt(self):
        """Open export options dialog then write a TXT file."""
        self._save_current_section()
        t = THEMES[self.current_theme]

        dlg = tk.Toplevel(self)
        dlg.title("Export TXT Options")
        dlg.configure(bg=t["bg"])
        dlg.resizable(False, False)
        dlg.grab_set()

        def sec_lbl(text):
            tk.Label(dlg, text=text, bg=t["bg"], fg=t["accent"],
                     font=FONT_TINY).pack(padx=16, pady=(10,2), anchor="w")

        def row_chk(text, var):
            tk.Checkbutton(dlg, text=text, variable=var,
                           bg=t["bg"], fg=t["fg"], selectcolor=t["input_bg"],
                           activebackground=t["bg"],
                           font=FONT_TINY).pack(anchor="w", padx=28)

        sec_lbl("Include layers:")
        vt = tk.BooleanVar(value=True)
        vc = tk.BooleanVar(value=True)
        vn = tk.BooleanVar(value=True)
        vl = tk.BooleanVar(value=True)
        row_chk("Tab (fret numbers)", vt)
        row_chk("Chords",             vc)
        row_chk("Notes",              vn)
        row_chk("Lyrics",             vl)

        all_instrs = sorted({s.instrument for s in self.sections})
        sec_lbl("Include instruments:")
        instr_vars = {}
        for instr in all_instrs:
            v = tk.BooleanVar(value=True)
            instr_vars[instr] = v
            row_chk(instr, v)

        def do_export():
            layers = {"tab": vt.get(), "chords": vc.get(),
                      "notes": vn.get(), "lyrics": vl.get()}
            instrs = {i for i, v in instr_vars.items() if v.get()}
            dlg.destroy()
            artist = self.song_artist.get().strip()
            title  = self.song_title.get().strip()
            default_name = (f"{artist} - {title}" if artist else title
                            ).replace(" ", "_") + ".txt"
            path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files","*.txt"),("All files","*.*")],
                initialfile=default_name)
            if not path: return
            lines = self._build_song_lines(layers=layers, instruments=instrs)
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
            messagebox.showinfo("Exported", f"Saved to:\n{path}")

        br = tk.Frame(dlg, bg=t["bg"])
        br.pack(fill="x", padx=16, pady=(10,14))
        ttk.Button(br, text="Cancel", command=dlg.destroy,
                   style="Normal.TButton").pack(side="right", padx=4)
        ttk.Button(br, text="Export TXT", command=do_export,
                   style="Accent.TButton").pack(side="right")
        dlg.bind("<Return>", lambda e: do_export())

    def _build_song_lines(self, layers=None, instruments=None):
        """
        Build the full song as a list of text lines for TXT export.
        Layer order (top→bottom): Chords → Notes → Tab → Lyrics.

        Parameters
        ----------
        layers      : dict {layer: bool}  (None = all True)
        instruments : set of instrument names to include  (None = all)
        """
        if layers is None:
            layers = {"tab": True, "chords": True, "notes": True, "lyrics": True}
        if instruments is None:
            instruments = None   # means all

        W   = 80
        div = lambda c="=": c * W
        lines = []

        artist = self.song_artist.get().strip()
        title  = self.song_title.get().strip()

        lines += [div("="), f"  {title.upper()}"]
        if artist: lines.append(f"  {artist}")
        meta = []
        if self.song_key.get():   meta.append(f"Key: {self.song_key.get()}")
        if self.song_tempo.get(): meta.append(f"BPM: {self.song_tempo.get()}")
        if self.song_time.get():  meta.append(f"Time: {self.song_time.get()}")
        if meta: lines.append("  " + "   ".join(meta))
        lines += [div("="), ""]

        TOKEN_W = 3   # chars per fret token in tab rows

        for s in self.sections:
            if instruments is not None and s.instrument not in instruments:
                continue

            rep = f"  (x{s.repeat})" if s.repeat > 1 else ""
            lines += [
                div("-"),
                f"  [{s.name}]{rep}   {s.measures} measures   {s.instrument}",
                div("-"),
            ]
            strings = INSTRUMENT_STRINGS.get(s.instrument, ["e","B","G","D","A","E"])

            # ── Measure number header ──────────────────────────────────────────
            has_tab = layers.get("tab") and any(s.layers["tab"])
            if has_tab:
                sn_pad = max(len(st) for st in strings) + 3
                hdr = " " * sn_pad
                for m in range(s.measures):
                    m_beats = s.measure_beats.get(m, s.tab_beats)
                    m_cell  = TOKEN_W * m_beats + 1
                    hdr    += f"{'M'+str(m+1):<{m_cell}}"
                lines += ["", hdr]

            # ── Chords ────────────────────────────────────────────────────────
            if layers.get("chords") and any(s.layers["chords"]):
                _cell = TOKEN_W * s.tab_beats + 1 if has_tab else 16
                lines += ["", "  Chords:"]
                row = "   "
                for m in range(s.measures):
                    m_beats = s.measure_beats.get(m, s.tab_beats)
                    m_cell  = TOKEN_W * m_beats + 1 if has_tab else 16
                    row += f"{(s.layers['chords'][m] or '-'):<{m_cell}}"
                lines.append(row)

            # ── Notes ─────────────────────────────────────────────────────────
            if layers.get("notes") and any(s.layers["notes"]):
                lines += ["", "  Notes:"]
                row = "   "
                for m in range(s.measures):
                    m_beats = s.measure_beats.get(m, s.tab_beats)
                    m_cell  = TOKEN_W * m_beats + 1 if has_tab else 16
                    row += f"{(s.layers['notes'][m] or '-'):<{m_cell}}"
                lines.append(row)

            # ── Tab ───────────────────────────────────────────────────────────
            if has_tab:
                for st in strings:
                    row = f"{st}| "
                    for m in range(s.measures):
                        m_beats = s.measure_beats.get(m, s.tab_beats)
                        cell    = s.layers["tab"][m]
                        raw     = (cell.get(st, "") if isinstance(cell, dict) else "")
                        tokens  = [tok for tok in raw.split() if tok] if raw else []
                        while len(tokens) < m_beats:
                            tokens.append("-")
                        tokens   = tokens[:m_beats]
                        cell_str = "".join(f"{tok:>{TOKEN_W}}" for tok in tokens) + "|"
                        row += cell_str
                    lines.append(row)

            # ── Lyrics ────────────────────────────────────────────────────────
            if layers.get("lyrics") and any(s.layers["lyrics"]):
                lines += ["", "  Lyrics:"]
                for m in range(s.measures):
                    v = s.layers["lyrics"][m]
                    if v: lines.append(f"  M{m+1}: {v}")

            lines.append("")

        lines += [div("="), f"  Generated by Song Notation Tool v{APP_VERSION}", div("=")]
        return lines

    # ==========================================================================
    #  PDF EXPORT
    # ==========================================================================

    def _export_pdf(self):
        """Open export options dialog then write a PDF."""
        self._save_current_section()
        t = THEMES[self.current_theme]

        dlg = tk.Toplevel(self)
        dlg.title("Export PDF Options")
        dlg.configure(bg=t["bg"])
        dlg.resizable(False, False)
        dlg.grab_set()

        def sec_lbl(text):
            tk.Label(dlg, text=text, bg=t["bg"], fg=t["accent"],
                     font=FONT_TINY).pack(padx=16, pady=(10,2), anchor="w")

        def row_chk(parent, text, var):
            tk.Checkbutton(parent, text=text, variable=var,
                           bg=t["bg"], fg=t["fg"], selectcolor=t["input_bg"],
                           activebackground=t["bg"],
                           font=FONT_TINY).pack(anchor="w", padx=28)

        # Orientation
        sec_lbl("Page orientation:")
        orient_var = tk.StringVar(value="landscape")
        for val, lbl in [("landscape","A4 Landscape  (842 x 595 pt, wider)"),
                          ("portrait", "A4 Portrait   (595 x 842 pt, taller)")]:
            tk.Radiobutton(dlg, text=lbl, variable=orient_var, value=val,
                           bg=t["bg"], fg=t["fg"], selectcolor=t["input_bg"],
                           activebackground=t["bg"],
                           font=FONT_TINY).pack(anchor="w", padx=28)

        # Layers
        sec_lbl("Include layers:")
        vt = tk.BooleanVar(value=True)
        vc = tk.BooleanVar(value=True)
        vn = tk.BooleanVar(value=True)
        vl = tk.BooleanVar(value=True)
        row_chk(dlg, "Tab (fret numbers)", vt)
        row_chk(dlg, "Chords",             vc)
        row_chk(dlg, "Notes",              vn)
        row_chk(dlg, "Lyrics",             vl)

        # Instruments
        all_instrs = sorted({s.instrument for s in self.sections})
        sec_lbl("Include instruments:")
        instr_vars = {}
        for instr in all_instrs:
            v = tk.BooleanVar(value=True)
            instr_vars[instr] = v
            row_chk(dlg, instr, v)

        def do_export():
            layers = {"tab": vt.get(), "chords": vc.get(),
                      "notes": vn.get(), "lyrics": vl.get()}
            instrs = {i for i, v in instr_vars.items() if v.get()}
            orient = orient_var.get()
            dlg.destroy()
            artist = self.song_artist.get().strip()
            title  = self.song_title.get().strip()
            default_name = (f"{artist} - {title}" if artist else title
                            ).replace(" ", "_") + ".pdf"
            path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files","*.pdf"),("All files","*.*")],
                initialfile=default_name)
            if not path: return
            pdf_bytes = self._build_pdf(layers=layers, instruments=instrs,
                                         orient=orient)
            with open(path, "wb") as f:
                f.write(pdf_bytes)
            messagebox.showinfo("Exported PDF", f"Saved to:\n{path}")

        br = tk.Frame(dlg, bg=t["bg"])
        br.pack(fill="x", padx=16, pady=(10,14))
        ttk.Button(br, text="Cancel", command=dlg.destroy,
                   style="Normal.TButton").pack(side="right", padx=4)
        ttk.Button(br, text="Export PDF", command=do_export,
                   style="Accent.TButton").pack(side="right")
        dlg.bind("<Return>", lambda e: do_export())

    def _build_pdf(self, layers=None, instruments=None, orient="landscape"):
        """
        Build a PDF with fixes for:
          - Non-latin1 chars (? glyphs) — safe ASCII substitution in _esc()
          - Measure numbers no longer overlap section headers
          - Tab tokens rendered per-token so 2-digit frets stay correct
          - Layer order: Chords -> Notes -> Tab -> Lyrics
          - Portrait / landscape orientation
          - Layer + instrument filtering
        """
        if layers is None:
            layers = {"tab": True, "chords": True, "notes": True, "lyrics": True}
        if instruments is None:
            instruments = set(s.instrument for s in self.sections)

        W, H = (842, 595) if orient == "landscape" else (595, 842)

        MARGIN   = 28
        LINE_H   = 12
        MONO_SZ  = 7.5
        HEAD_SZ  = 9
        TITLE_SZ = 13
        FOOTER_H = 18
        TOKEN_W  = 3      # chars per tab beat token ("  -", " 10", etc.)
        CHAR_W   = 4.6    # approx pts per monospace char at MONO_SZ
        SN_W     = 20     # string-name prefix width in pts ("G|  " or "Ch| ")

        artist   = self.song_artist.get().strip()
        title    = self.song_title.get().strip()
        doc_date = datetime.date.today().strftime("%Y-%m-%d")

        pages  = []
        cur_ln = []

        # ── Drawing helpers ────────────────────────────────────────────────────
        def _esc(s):
            """Escape a string for a PDF text literal — latin-1 safe."""
            s = str(s)
            # Replace common non-latin-1 unicode with safe ASCII equivalents
            for frm, to in [
                ("\u2014", "-"), ("\u2013", "-"),   # em/en dash
                ("\u00d7", "x"),                     # multiplication sign x
                ("\u00d8", "x"),                     # Ø
                ("\u2019", "'"), ("\u2018", "'"),     # smart quotes
                ("\u201c", '"'), ("\u201d", '"'),     # smart double quotes
                ("\u00e9", "e"), ("\u00e8", "e"),     # accented e
                ("\u00e0", "a"), ("\u00f4", "o"),     # accented vowels
                ("\u266a", ""), ("\u2665", ""),       # music / heart symbols
                ("\u00d6", "O"), ("\u00fc", "u"),     # German umlauts
            ]:
                s = s.replace(frm, to)
            s = s.encode("latin-1", errors="replace").decode("latin-1")
            return s.replace("\\","\\\\").replace("(","\\(").replace(")","\\)")

        def txt(x, y, s, sz=MONO_SZ, bold=False):
            font = "/F2" if bold else "/F1"
            cur_ln.append(
                f"BT {font} {sz} Tf {x:.1f} {y:.1f} Td ({_esc(s)}) Tj ET")

        def color(r, g, b):
            cur_ln.append(f"{r:.3f} {g:.3f} {b:.3f} rg")

        def hline(x1, y, x2, width=0.25, gray=0.72):
            cur_ln.append(
                f"{gray:.2f} G {width} w {x1:.1f} {y:.1f} m "
                f"{x2:.1f} {y:.1f} l S")

        def rfill(x, y, w, h, r, g, b):
            cur_ln.append(
                f"{r:.3f} {g:.3f} {b:.3f} rg "
                f"{x:.1f} {y:.1f} {w:.1f} {h:.1f} re f")

        def finish_page(pn):
            hline(MARGIN, FOOTER_H, W - MARGIN)
            color(0.4, 0.4, 0.4)
            txt(MARGIN, 5,
                f"Song Notation Tool v{APP_VERSION}  -  {doc_date}", sz=6.5)
            txt(W - MARGIN - 28, 5, f"Page {pn}", sz=6.5)
            if cur_ln:
                pages.append(zlib.compress(
                    "\n".join(cur_ln).encode("latin-1")))
                cur_ln.clear()

        # ── Measure column width: based on max beats across all sections ───────
        # Use a uniform width so columns are consistent within a page line.
        max_beats = max(
            (s.measure_beats.get(m, s.tab_beats)
             for s in self.sections if s.instrument in instruments
             for m in range(s.measures)),
            default=TAB_BEATS_DEFAULT,
        )
        COL_W = TOKEN_W * max_beats * CHAR_W + CHAR_W  # pts wide per measure col

        def mpl_for(n_measures):
            """Measures that fit on one line."""
            usable = W - 2 * MARGIN - SN_W
            return max(1, min(n_measures, int(usable / COL_W)))

        # ── Title header (first page) ──────────────────────────────────────────
        pn = 1
        cy = H - MARGIN

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

        for s in self.sections:
            if s.instrument not in instruments:
                continue

            strings = INSTRUMENT_STRINGS.get(s.instrument, ["e","B","G","D","A","E"])
            rep_str = f" (x{s.repeat})" if s.repeat > 1 else ""
            mpl     = mpl_for(s.measures)

            # Rough line count for page-break estimation
            has_tab = layers.get("tab") and any(s.layers["tab"])
            has_ch  = layers.get("chords") and any(s.layers["chords"])
            has_no  = layers.get("notes")  and any(s.layers["notes"])
            has_ly  = layers.get("lyrics") and any(s.layers["lyrics"])
            n_str   = len(strings) if has_tab else 0
            batches = (s.measures + mpl - 1) // mpl
            lns_per_batch = 1 + (1 if has_ch else 0) + (1 if has_no else 0) \
                            + n_str + (1 if has_ly else 0) + 1
            needed_h = (batches * lns_per_batch + 3) * LINE_H

            if cy - needed_h < FOOTER_H + LINE_H * 3:
                finish_page(pn); pn += 1; cy = H - MARGIN

            # ── Section header bar ─────────────────────────────────────────────
            sec_label = (f"[ {s.name} ]{rep_str}"
                         f"   {s.measures} measures   {s.instrument}")
            rfill(MARGIN, cy - LINE_H, W - 2*MARGIN, LINE_H + 2,
                  0.16, 0.24, 0.42)
            color(1, 1, 1)
            txt(MARGIN + 4, cy - LINE_H + 3, sec_label, sz=HEAD_SZ, bold=True)
            cy -= LINE_H + 6

            # ── Render each batch of measures ──────────────────────────────────
            for bs in range(0, s.measures, mpl):
                batch = list(range(bs, min(bs + mpl, s.measures)))

                if cy < FOOTER_H + LINE_H * 4:
                    finish_page(pn); pn += 1; cy = H - MARGIN

                # Helper: x-origin of batch member i
                def col_x(i):
                    return MARGIN + SN_W + sum(
                        TOKEN_W * s.measure_beats.get(batch[j], s.tab_beats)
                        * CHAR_W + CHAR_W
                        for j in range(i))

                # Measure number header (separate row, above content)
                color(0.40, 0.58, 0.82)
                for i, m in enumerate(batch):
                    txt(col_x(i), cy, f"M{m+1}", sz=7)
                cy -= LINE_H

                # ── Chords ────────────────────────────────────────────────────
                if has_ch:
                    vals = [s.layers["chords"][m] for m in batch]
                    if any(vals):
                        color(0.42, 0.14, 0.62)
                        txt(MARGIN, cy, "Ch|", sz=MONO_SZ, bold=True)
                        color(0, 0, 0)
                        for i, m in enumerate(batch):
                            txt(col_x(i), cy,
                                s.layers["chords"][m] or "-", sz=MONO_SZ)
                        cy -= LINE_H

                # ── Notes ─────────────────────────────────────────────────────
                if has_no:
                    vals = [s.layers["notes"][m] for m in batch]
                    if any(vals):
                        color(0.08, 0.46, 0.28)
                        txt(MARGIN, cy, "No|", sz=MONO_SZ, bold=True)
                        color(0, 0, 0)
                        for i, m in enumerate(batch):
                            txt(col_x(i), cy,
                                s.layers["notes"][m] or "-", sz=MONO_SZ)
                        cy -= LINE_H

                # ── Tab ───────────────────────────────────────────────────────
                if has_tab:
                    for st in strings:
                        color(0.16, 0.32, 0.58)
                        txt(MARGIN, cy, f"{st}|", sz=MONO_SZ)
                        for i, m in enumerate(batch):
                            m_beats = s.measure_beats.get(m, s.tab_beats)
                            cell    = s.layers["tab"][m]
                            raw     = (cell.get(st, "")
                                       if isinstance(cell, dict) else "")
                            tokens  = [tok for tok in raw.split() if tok]
                            while len(tokens) < m_beats:
                                tokens.append("-")
                            tokens   = tokens[:m_beats]
                            # Render token-by-token (fixes 2-digit fret display)
                            row_str  = "".join(
                                f"{tok:>{TOKEN_W}}" for tok in tokens) + "|"
                            color(0, 0, 0)
                            txt(col_x(i), cy, row_str, sz=MONO_SZ)
                        cy -= LINE_H
                        if cy < FOOTER_H + LINE_H:
                            finish_page(pn); pn += 1; cy = H - MARGIN

                # ── Lyrics ────────────────────────────────────────────────────
                if has_ly:
                    ly_vals = [(m, s.layers["lyrics"][m]) for m in batch
                               if s.layers["lyrics"][m]]
                    if ly_vals:
                        color(0.52, 0.26, 0.08)
                        txt(MARGIN, cy, "Ly|", sz=MONO_SZ, bold=True)
                        color(0, 0, 0)
                        for m, v in ly_vals:
                            txt(col_x(batch.index(m)), cy, v[:24], sz=MONO_SZ)
                        cy -= LINE_H

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
        Two built-in fonts: F1=Helvetica, F2=Helvetica-Bold.
        For monospaced content a third font F3=Courier is added.
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
        cid = 6   # start after catalog(1), pages(2), F1(3), F2(4), F3(5)

        for ps in pages:
            kids.append(f"{cid} 0 R")
            po.append(obj(cid,
                f"<< /Type /Page /Parent 2 0 R "
                f"/MediaBox [0 0 {W:.2f} {H:.2f}] "
                f"/Contents {cid+1} 0 R /Resources {font_res} >>"))
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

    def _save(self):
        """Serialise the whole project to a JSON .sng file."""
        self._save_current_section()
        artist = self.song_artist.get().strip()
        title  = self.song_title.get().strip()
        default_name = (f"{artist} - {title}" if artist else title
                        ).replace(" ", "_") + ".sng"

        path = filedialog.asksaveasfilename(
            defaultextension=".sng",
            filetypes=[("Song files", "*.sng"), ("All files", "*.*")],
            initialfile=default_name)
        if not path: return

        data = {
            "app_version": APP_VERSION,
            "title":       self.song_title.get(),
            "artist":      self.song_artist.get(),
            "key":         self.song_key.get(),
            "tempo":       self.song_tempo.get(),
            "time":        self.song_time.get(),
            "instrument":  self.instrument.get(),
            "sections":    [s.to_dict() for s in self.sections],
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        self.dirty = False
        messagebox.showinfo("Saved", f"Project saved to:\n{path}")

    def _open(self):
        """Load a .sng JSON project file."""
        path = filedialog.askopenfilename(
            filetypes=[("Song files", "*.sng"), ("All files", "*.*")])
        if not path: return
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.song_title.set( data.get("title",      ""))
        self.song_artist.set(data.get("artist",     ""))
        self.song_key.set(   data.get("key",        ""))
        self.song_tempo.set( data.get("tempo",      ""))
        self.song_time.set(  data.get("time",       "4/4"))
        self.instrument.set( data.get("instrument", "Guitar (6-string)"))
        self.sections = [Section.from_dict(d)
                         for d in data.get("sections", [])]
        # Recalculate the link id counter so new groups don't collide
        existing_ids = [s.link_id for s in self.sections if s.link_id is not None]
        self._next_link_id = (max(existing_ids) + 1) if existing_ids else 1
        self._refresh_listbox()
        self._clear_editor()
        self.dirty = False

    # ==========================================================================
    #  KEYBOARD SHORTCUTS & WINDOW CLOSE
    # ==========================================================================

    def _bind_shortcuts(self):
        """Global keyboard shortcuts for macOS (Cmd) and Win/Linux (Ctrl)."""
        self.bind("<Command-s>", lambda e: self._save())
        self.bind("<Control-s>", lambda e: self._save())
        self.bind("<Command-e>", lambda e: self._export_txt())
        self.bind("<Control-e>", lambda e: self._export_txt())
        self.bind("<Command-p>", lambda e: self._export_pdf())
        self.bind("<Control-p>", lambda e: self._export_pdf())
        self.bind("<Command-n>", lambda e: self._add_section_dialog())
        self.bind("<Control-n>", lambda e: self._add_section_dialog())
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        """Ask for confirmation if there are unsaved changes."""
        self._save_current_section()
        if self.dirty:
            if messagebox.askyesno(
                    "Quit", "You have unsaved changes. Quit anyway?"):
                self.destroy()
        else:
            self.destroy()


# ==============================================================================
#  ENTRY POINT
# ==============================================================================

if __name__ == "__main__":
    app = SongNotationApp()
    app.mainloop()

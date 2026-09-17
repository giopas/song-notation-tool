"""
constants.py — Song Notation Tool shared domain constants.

Pure data, no UI dependency. Split out of song_writer.py so the CLI,
export, and web-server layers don't have to import Tkinter just to read
an instrument's string list or the default tab beat count.
"""

from __future__ import annotations

APP_VERSION = "0.18"
APP_TITLE = f"Song Notation Tool  v{APP_VERSION}"

# Section-type presets shown in the New Section dialog / web UI.
SECTION_TYPES = [
    "Intro", "Verse", "Pre-Chorus", "Chorus", "Refrain",
    "Bridge", "Interlude", "Solo", "Breakdown", "Outro", "Custom",
]

# Maps instrument/tuning -> string list (high -> low pitch).
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


def default_export_name(doc: dict, ext: str) -> str:
    """'<Artist> - <Title>.ext', or just '<Title>.ext' with no artist."""
    meta = doc.get("meta", {})
    artist = (meta.get("artist") or "").strip()
    title = (meta.get("title") or "Untitled Song").strip()
    stem = f"{artist} - {title}" if artist else title
    return stem.replace(" ", "_") + f".{ext.lstrip('.')}"

"""
constants.py — Song Notation Tool shared domain constants.

Pure data, no UI dependency. Split out of song_writer.py so the CLI,
export, and web-server layers don't have to import Tkinter just to read
an instrument's string list or the default tab beat count.
"""

from __future__ import annotations

APP_VERSION = "0.24.3"
APP_TITLE = f"Song Notation Tool  v{APP_VERSION}"

# Printed in the footer of every export, so a chart handed to someone else
# says where it came from and where to get the program.
APP_URL = "https://github.com/giopas/song-notation-tool"

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

SECTION_LAYOUT_LABELS = {"banner": "Sections on top",
                         "gutter": "Sections on the left"}

# Where the song's lyrics print — one choice for the whole song, made on
# the main screen. Either every section's words gathered into one block
# (at the start, at the end, or down a column of their own on the left),
# or each section's words with its own chart, or nothing. "Beside" is
# the pick for reading while playing: the chart keeps a narrow column and
# the words run down the right of it — what the instrument does *during*
# those words, with no pretence of aligning chord to syllable.
LYRICS_LAYOUT_LABELS = {
    "none":   "Don't print",
    "start":  "All together — at the start",
    "end":    "All together — at the end",
    "side":   "All together — left column",
    "beside": "Each section — beside chart",
    "below":  "Each section — under chart",
}

# What a recalled lick ({Riff1}) prints as. "Tab" plays it back in full
# wherever it's used; "Name only" prints "Riff1 (x3)" and leaves the notes
# where the lick was written — how the handwritten charts do it, and the
# shorter page.
LICK_REF_LABELS = {"tab": "Tab again",
                   "name": "Name only — Riff1 (x3)"}

# A page of chord shapes — the voicings you had to work out — printed
# once at one end of the song rather than beside every chord symbol.
CHORD_SHEET_LABELS = {"none": "Don't print",
                      "start": "At the start",
                      "end": "At the end"}

# Print colours, one per section type, mirroring the accent each section
# card carries in the UI so screen and paper agree at a glance. These are
# darkened versions of the on-screen pastels: a colour that reads well on
# a dark editor background is washed out on white paper.
SECTION_COLORS = {
    "Intro":      (0.05, 0.47, 0.42),   # teal
    "Verse":      (0.13, 0.33, 0.70),   # blue
    "Pre-Chorus": (0.05, 0.42, 0.62),   # cyan
    "Chorus":     (0.72, 0.16, 0.28),   # red
    "Refrain":    (0.60, 0.20, 0.42),   # pink
    "Bridge":     (0.42, 0.22, 0.66),   # mauve
    "Interlude":  (0.20, 0.45, 0.28),   # green
    "Solo":       (0.78, 0.42, 0.05),   # amber
    "Breakdown":  (0.45, 0.32, 0.18),   # brown
    "Outro":      (0.30, 0.32, 0.60),   # periwinkle
    "Custom":     (0.25, 0.25, 0.30),   # slate
}
DEFAULT_SECTION_COLOR = (0.16, 0.24, 0.42)

COLOR_MODE_LABELS = {"color": "Colour", "bw": "Black & white"}

# Two things on a chart line are not chord symbols and shouldn't be read
# as if they were: a lick (tab, in among the chords) and a rest (the bar
# you don't play). Both get a colour of their own rather than a bigger
# label. Blue for the lick — it's the thing you look down at mid-song, so
# it has to be findable at a glance; grey for the rest, because it should
# recede rather than compete with the notes on either side of it.
LICK_RGB = (0.10, 0.35, 0.78)
REST_RGB = (0.52, 0.52, 0.55)

# A fret number is a hint about *where* to play the chord under it, not a
# second line of notes — so it prints as a figure over the symbol: small,
# raised, and a step back from the black the symbols are set in. Amber on
# colour; on a mono printer a light grey, which says the same thing the
# only way a photocopier can.
FRET_RGB = (0.72, 0.47, 0.02)
FRET_BW_RGB = (0.48, 0.48, 0.48)

# How the fret row is set against the symbol row below it: a fraction of
# the type size, and of the line height — small enough to read as a
# figure over the chord, close enough to belong to it.
FRET_SIZE_RATIO = 0.72
FRET_LINE_RATIO = 0.74

# The same colours as CSS hex, for the browser front end.
LICK_HEX = "#1a59c7"
REST_HEX = "#85858c"

# PDF sizing. "fit" scales the whole chart up until it fills the page
# without needing another one — the point being a chart you can read from
# a music stand, or off the floor.
PDF_SCALE_LABELS = {
    "fit": "Fit to page",
    "one": "Fit to one page",
    "1.0": "100%",
    "1.25": "125%",
    "1.5": "150%",
    "2.0": "200%",
}

# Two-column printing. Most chart lines are short — a few chord symbols,
# a bar of tab — so one column down an A4 page leaves half the sheet
# white and caps how far "fit to page" can grow the type. Two columns
# halve the height the chart needs, and the fit spends that height on
# type size instead: the same song, bigger and still on one sheet.
# "Auto" splits only when it actually buys a bigger chart.
PDF_COLUMN_LABELS = {
    "auto": "Auto (2 columns if they help)",
    "1": "One column",
    "2": "Two columns",
}
MAX_PDF_COLUMNS = 2
# Space between columns, in points, and the grey of the rule drawn down
# the middle of it — the split has to be visible or you lose your place
# halfway down the page.
PDF_COLUMN_GAP = 22
PDF_COLUMN_RULE_GRAY = 0.78


# Black-and-white printing fills the heading bands with a light grey and
# sets the text in black, rather than reversing white out of solid black.
# A full-width solid band per section is a lot of toner for something
# that only has to say "new section starts here" — grey says it just as
# clearly, and survives a photocopy better than reversed-out text.
BW_BAND_FILL = (0.86, 0.86, 0.86)
BW_TITLE_FILL = (0.80, 0.80, 0.80)
BW_BAND_TEXT = (0.0, 0.0, 0.0)


def heading_fill(section_type: str, color_mode: str = "color"):
    """(band_rgb, text_rgb) for a section's heading band."""
    if color_mode == "bw":
        return BW_BAND_FILL, BW_BAND_TEXT
    return section_color(section_type, color_mode), (1.0, 1.0, 1.0)


def title_fill(color_mode: str = "color"):
    """(band_rgb, text_rgb) for the song's title band at the top."""
    if color_mode == "bw":
        return BW_TITLE_FILL, BW_BAND_TEXT
    return (0.10, 0.12, 0.22), (1.0, 1.0, 1.0)


def section_color(section_type: str, color_mode: str = "color"):
    """RGB for a section's heading and its content. Black-and-white mode
    collapses every section to near-black so a mono printer (or a
    photocopy) doesn't turn the palette into indistinguishable greys."""
    if color_mode == "bw":
        return (0.0, 0.0, 0.0)
    return SECTION_COLORS.get(section_type, DEFAULT_SECTION_COLOR)


RENDER_MODE_LABELS = {"chart": "Chart", "tab": "Tab grid",
                      "both": "Chart + Tab", "free": "Free text"}

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

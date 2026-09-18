"""
model.py — Song Notation Tool v0.16 data model.

Pure data: schema, item kinds, (de)serialisation, v0.15 -> v0.16 migration.
No Tkinter dependency. See DESIGN_v0_16.md sections 3 and 6.
"""

from __future__ import annotations

FORMAT_VERSION = 2

ITEM_KINDS = ("token", "group", "block_ref", "section_ref", "measure", "mark")

MARK_NAMES = (
    "repeat_open", "repeat_close", "ending_1", "ending_2",
    "segno", "coda", "dc", "ds", "simile", "rest",
)

RENDER_MODES = ("chart", "tab", "both", "free")

# Where a section's name goes in the TXT/PDF output:
#   "banner" — a full-width header band above the section (the default)
#   "gutter" — a left-hand column beside the section's first line,
#              which fits far more of a song on one page
SECTION_LAYOUTS = ("banner", "gutter")


# ==============================================================================
#  Item constructors — thin helpers that guarantee the required keys exist.
# ==============================================================================

def make_token(symbol, fret=None):
    return {"kind": "token", "symbol": symbol, "fret": fret}


def make_group(items=None, repeat=1, note=""):
    d = {"kind": "group", "repeat": repeat, "items": list(items or [])}
    if note:
        d["note"] = note
    return d


def make_block_ref(block, repeat=1, transpose=0):
    return {"kind": "block_ref", "block": block, "repeat": repeat, "transpose": transpose}


def make_section_ref(section, repeat=1, all=False, transpose=0):
    return {"kind": "section_ref", "section": section, "repeat": repeat,
            "all": all, "transpose": transpose}


def make_measure(beats, strings):
    return {"kind": "measure", "beats": beats, "strings": dict(strings)}


def make_mark(mark):
    if mark not in MARK_NAMES:
        raise ValueError(f"unknown mark: {mark!r}")
    return {"kind": "mark", "mark": mark}


# ==============================================================================
#  Document / Section / Block scaffolding
# ==============================================================================

def new_document(title="", artist="", key="", time="4/4", bpm=""):
    return {
        "format": FORMAT_VERSION,
        "app_version": "0.16",
        "meta": {"title": title, "artist": artist, "key": key, "time": time, "bpm": bpm},
        "transpose": 0,
        "lyrics_text": "",
        "print_lyrics": False,
        "section_layout": "banner",
        "blocks": {},
        "sections": [],
    }


def new_block(block_id, name, bars=4, beats_per_bar=8, instrument="Bass (4-string)"):
    return {
        "id": block_id, "name": name, "bars": bars,
        "beats_per_bar": beats_per_bar, "instrument": instrument, "items": [],
    }


def new_section(section_id, name, section_type="Verse",
                 instrument="Bass (4-string)", repeat=1, render="chart",
                 free_text=""):
    return {
        "id": section_id, "name": name, "type": section_type,
        "instrument": instrument, "repeat": repeat, "transpose": 0,
        "render": render, "annotation": "", "lyrics_text": "",
        "print_lyrics": False, "free_text": free_text, "items": [],
    }


import re as _re

_BLOCK_NAME_REJECT_RE = _re.compile(r'^\d{0,2}[A-Ga-g]([#b])?')


def validate_block_name(name):
    """
    Section 4.4 — reject a block name that would be ambiguous with a token
    (e.g. "A1", "b2"). Returns True if the name is valid, False otherwise.
    """
    if not name:
        return False
    return not _BLOCK_NAME_REJECT_RE.match(name)


# ==============================================================================
#  v0.15 -> v0.16 migration  (section 6)
# ==============================================================================

def migrate_document(doc: dict) -> dict:
    """
    Given a raw loaded .sng dict, return a format-2 document.
    A missing "format" key is treated as format 1 (v0.15).
    Already-format-2 documents pass through unchanged.
    """
    fmt = doc.get("format", 1)
    if fmt >= FORMAT_VERSION:
        return doc
    return _migrate_v1_to_v2(doc)


def _migrate_v1_to_v2(doc: dict) -> dict:
    sections_in = doc.get("sections", [])

    # First pass: figure out link groups so we know which section in a
    # group becomes the "real" one and which become section_ref's.
    link_groups = {}
    for s in sections_in:
        lid = s.get("link_id")
        if lid is not None:
            link_groups.setdefault(lid, []).append(s)

    new_sections = []
    seen_link_ids = set()

    for s in sections_in:
        new_sections.append(_migrate_section(s, link_groups, seen_link_ids))

    out = new_document(
        title=doc.get("title", ""),
        artist=doc.get("artist", ""),
        key=doc.get("key", ""),
        time=doc.get("time", "4/4"),
        bpm=doc.get("bpm", ""),
    )
    out["app_version"] = doc.get("app_version", "0.15")
    out["sections"] = new_sections
    # v0.15 had no blocks / riff library — start empty.
    out["blocks"] = {}
    return out


def _migrate_section(s: dict, link_groups: dict, seen_link_ids: set) -> dict:
    lid = s.get("link_id")
    is_repeat_of_earlier = lid is not None and lid in seen_link_ids
    if lid is not None:
        seen_link_ids.add(lid)

    sec_id = s.get("id") or s.get("name", "section")
    render = _migrate_render_mode(s)

    out = new_section(
        section_id=sec_id,
        name=s.get("name", ""),
        section_type=s.get("type", s.get("name", "")),
        instrument=s.get("instrument", "Bass (4-string)"),
        repeat=s.get("repeat", 1),
        render=render,
    )

    if is_repeat_of_earlier:
        # "as before" — point back at the first section in this link group.
        first = link_groups[lid][0]
        first_id = first.get("id") or first.get("name", "section")
        out["items"] = [make_section_ref(first_id, repeat=s.get("repeat", 1))]
        return out

    layers = s.get("layers", {})
    tab_beats = s.get("tab_beats", 8)
    measure_beats = {int(k): v for k, v in s.get("measure_beats", {}).items()}
    n_measures = len(layers.get("tab", [])) or len(layers.get("chords", [])) \
        or len(layers.get("notes", [])) or 0

    items = []
    tab_layer = layers.get("tab", [])
    chords_layer = layers.get("chords", [])
    notes_layer = layers.get("notes", [])

    for m in range(n_measures):
        beats = measure_beats.get(m, tab_beats)
        tab_cell = tab_layer[m] if m < len(tab_layer) else {}
        chord_val = chords_layer[m] if m < len(chords_layer) else ""
        note_val = notes_layer[m] if m < len(notes_layer) else ""

        if isinstance(tab_cell, dict) and any(
                (v or "").strip().strip("-") for v in tab_cell.values()):
            items.append(make_measure(beats, tab_cell))

        # chords/notes merge into token items per the migration table:
        # notes only used if chords is empty for that measure, else kept
        # as a parallel token stream (we emit both, chords first).
        if chord_val:
            items.append(make_token(chord_val))
        elif note_val:
            items.append(make_token(note_val))

    out["items"] = items
    out["lyrics"] = list(layers.get("lyrics", []))
    out["tab_beats"] = tab_beats
    return out


def _migrate_render_mode(s: dict) -> str:
    visible = s.get("visible", {})
    tab_visible = visible.get("tab", True)
    layers = s.get("layers", {})
    has_tab = any(
        isinstance(c, dict) and any((v or "").strip().strip("-") for v in c.values())
        for c in layers.get("tab", [])
    )
    if has_tab:
        return "tab" if tab_visible else "both"
    return "chart"

#!/usr/bin/env python3
"""Build the golden-file test songs and their expected exports.

    python3 tools/make_golden.py            # write tests/golden/*
    python3 tools/make_golden.py --check    # compare only, write nothing

Every song here is invented: band, titles and lyrics are made up for the
tests (no real songs in the repository). Each one exercises a layout
feature, and the expected TXT and PDF are what export.py produces for it
with the footer date fixed to GOLDEN_DATE and the version to
GOLDEN_VERSION. tests/test_golden.py rebuilds
them and compares.

Run it only when an output change is intended, then look at the new PDFs
before committing them: a golden file records what was approved, so it
must be checked by eye the first time.
"""
from __future__ import annotations

import datetime
import json
import os
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, ROOT)

import export   # noqa: E402
import grammar  # noqa: E402
import model    # noqa: E402

GOLDEN_DIR = os.path.join(ROOT, "tests", "golden")
GOLDEN_DATE = datetime.date(2026, 1, 2)
BAND = "The Placeholders"


def _sec(sid, name, kind, line="", instrument="Bass (4-string)", repeat=1,
         render="chart", lyrics="", free_text=""):
    sec = model.new_section(sid, name, kind, instrument=instrument,
                            repeat=repeat, render=render, free_text=free_text)
    if line:
        items, annotation = grammar.parse(line)
        sec["items"] = items
        sec["annotation"] = annotation or ""
    sec["lyrics_text"] = lyrics
    return sec


def _doc(title, key="E", bpm="120", **settings):
    doc = model.new_document(title=title, artist=BAND, key=key, time="4/4", bpm=bpm)
    doc.update(settings)
    return doc


def song_basics():
    """Sections on top, colour, fit to page: frets, groups, marks, line
    breaks, an indent, an annotation, a reference and a free-text section."""
    doc = _doc("Paper Lanterns")
    doc["sections"] = [
        _sec("intro", "Intro", "Intro", "5A 5D [5A 5D]x2"),
        _sec("verse1", "Verse 1", "Verse",
             '5A 5D // 8F 5D 7E //> 5A 5A 8F "lightly"', repeat=2),
        _sec("chorus1", "Chorus 1", "Chorus", "|: 7B 4G# :| |1. 3G |2. 5D % rest"),
        _sec("verse2", "Verse 2", "Verse", "=verse1 x2"),
        _sec("bridge", "Bridge", "Bridge", render="free",
             free_text="Hold the last chord.\nDrums alone for four bars."),
        _sec("outro", "Outro", "Outro", "segno 5A 5D coda dc"),
    ]
    return doc


def song_gutter_bw_two_columns():
    """Sections on the left, black and white, two columns forced."""
    doc = _doc("Slow Tide", key="A", section_layout="gutter",
               color_mode="bw", pdf_columns="2")
    lines = ["5A 5D 7E 5A", "[3G 5D]x2 7E", "5A // 5D 7E // 5A 5D",
             "7E 7E 5A 5A", "[5A 7E]x4", "3G 3G 5D 5D"]
    kinds = ["Intro", "Verse", "Pre-Chorus", "Chorus", "Bridge", "Outro"]
    doc["sections"] = [_sec(f"s{i}", f"{k} {i}", k, ln)
                       for i, (k, ln) in enumerate(zip(kinds * 2, lines * 2), 1)]
    return doc


def song_licks_tab():
    """Inline licks, a named lick, recalls with a repeat, a group holding a
    lick, recalled licks printed as tab."""
    doc = _doc("Copper Wire", key="G")
    doc["sections"] = [
        _sec("intro", "Intro", "Intro", "C {G 5 7 5 | D - - 3} G Am"),
        _sec("verse1", "Verse", "Verse", "{Riff1 = G 5 7 5 | D - - 3} 5A 5D"),
        _sec("chorus1", "Chorus", "Chorus", "{Riff1}x3 5A 5D"),
        _sec("bridge", "Bridge", "Bridge",
             "[3G 2F# {Riff3 = G - - - - | D 4 - - - | A - 4 5 4 | E - - - -}]x4"),
    ]
    return doc


def song_licks_name_only():
    """The same song with recalled licks printed by name only."""
    doc = song_licks_tab()
    doc["meta"]["title"] = "Copper Wire (names)"
    doc["lick_refs"] = "name"
    return doc


def song_lyrics_beside():
    """Each section's words beside its chart."""
    doc = _doc("Harbour Lights", lyrics_layout="beside")
    doc["sections"] = [
        _sec("verse1", "Verse 1", "Verse", "5A 5D 8F 5D",
             lyrics="Salt on the window\nlights on the bay"),
        _sec("chorus1", "Chorus", "Chorus", "7B 4G# 3G 5D",
             lyrics="Hold the rope\nhold the line\nwe come in slow"),
        _sec("solo", "Solo", "Solo", "[5A 7E]x4"),
    ]
    return doc


def song_lyrics_left_column():
    """All the words together in a column on the left."""
    doc = song_lyrics_beside()
    doc["meta"]["title"] = "Harbour Lights (left column)"
    doc["lyrics_layout"] = "side"
    return doc


def song_lyrics_sheet_at_start():
    """Words gathered at the start, from an unsplit sheet with markers."""
    doc = _doc("Grey Morning", lyrics_layout="start")
    doc["lyrics_text"] = ("=== Verse 1 ===\nCoffee gone cold\nradio low\n\n"
                          "=== Chorus ===\nGrey morning\nlet it go")
    doc["sections"] = [
        _sec("verse1", "Verse 1", "Verse", "5A 5D"),
        _sec("chorus1", "Chorus", "Chorus", "7E 5A"),
    ]
    return doc


def song_chord_shapes_transposed():
    """Chord shapes printed at the start, with the whole song transposed."""
    doc = _doc("Open Strings", key="C", chord_sheet="start", transpose=2)
    guitar = "Guitar (6-string)"
    doc["chords"] = [
        model.make_chord("C", ["0", "1", "0", "2", "3", "x"], guitar),
        model.make_chord("G", ["3", "0", "0", "0", "2", "3"], guitar, note="full"),
        model.make_chord("F", ["1", "1", "2", "3", "3", "1"], guitar, note="barre"),
    ]
    doc["sections"] = [
        _sec("verse1", "Verse", "Verse", "C G F C", instrument=guitar),
        _sec("chorus1", "Chorus", "Chorus", "F G C C", instrument=guitar),
    ]
    return doc


def song_fit_one_page():
    """A long song squeezed onto one page."""
    doc = _doc("Long Road Home", pdf_scale="one")
    lines = ["5A 5D 7E 5A // 3G 5D 7E 5A", "[5A 7E]x2 // 3G 3G 5D 5D",
             "7E 5A 5D 3G // 5A 5D 7E 5A"]
    doc["sections"] = [_sec(f"s{i}", f"Part {i}", "Verse", lines[i % 3])
                       for i in range(1, 25)]
    return doc


def song_tab_grid_landscape():
    """A tab-grid section and a "both" section, exported in landscape."""
    doc = _doc("Fret Study", pdf_scale="1.5")
    grid = _sec("riff", "Riff", "Intro", render="tab")
    grid["items"] = [model.make_measure(8, {"D": "- - 3 - - - - -",
                                            "A": "4 4 - 4 - - - -"})]
    both = _sec("verse1", "Verse", "Verse", "5A 5D", render="both")
    both["items"].append(model.make_measure(8, {"G": "5 - - - 7 - - -"}))
    doc["sections"] = [grid, both]
    return doc


def song_marks_and_techniques():
    """Dynamics, accents, a palm-mute run, cresc and dim hairpins, ending
    brackets carried across a line break, a coda, and lick techniques."""
    doc = _doc("Signal Fires")
    doc["sections"] = [
        _sec("intro", "Intro", "Intro", "p [pm 5A 5A 5A 5A] mf 7D^ 7D"),
        _sec("verse1", "Verse", "Verse",
             "|: 5A 5D |1. 8F 7E :| |2. 5A // 3G 5D"),
        _sec("chorus1", "Chorus", "Chorus",
             "[cresc A D E] ff F^ [dim E D A] coda"),
        _sec("solo", "Solo", "Solo",
             "{G 5h7 7p5 5/7 7\\5 | D 7b9 7b9r7 7~ x}"),
    ]
    tab = _sec("fill", "Fill", "Bridge", render="tab")
    tab["items"] = [model.make_measure(8, {"G": "5h7 - 7p5 - 5/7 - - -",
                                           "D": "- - - - - - 7b9r7 -"})]
    doc["sections"].append(tab)
    return doc


def song_parts():
    """A guitar part under two bass sections, and a section of guitar only;
    the Layout picks both instruments."""
    doc = _doc("Two Hands", key="A")
    verse = _sec("verse1", "Verse", "Verse", "5A 5A 7D 7D")
    verse["parts"] = [model.make_part("p1", "Guitar (6-string)",
                                      grammar.parse("A A D [pm E E]")[0])]
    chorus = _sec("chorus1", "Chorus", "Chorus", "5D 5D 7E //", repeat=2)
    chorus["parts"] = [model.make_part("p2", "Guitar (6-string)",
                                       grammar.parse("D D E E^")[0])]
    solo = _sec("solo", "Solo", "Solo", "{B 5 7 8 - | G - - - 7}",
                instrument="Guitar (6-string)")
    doc["sections"] = [verse, chorus, solo]
    return doc


SONGS = {
    "basics": (song_basics, "portrait"),
    "gutter-bw-two-columns": (song_gutter_bw_two_columns, "portrait"),
    "licks-tab": (song_licks_tab, "portrait"),
    "licks-name-only": (song_licks_name_only, "portrait"),
    "lyrics-beside": (song_lyrics_beside, "portrait"),
    "lyrics-left-column": (song_lyrics_left_column, "portrait"),
    "lyrics-sheet-at-start": (song_lyrics_sheet_at_start, "portrait"),
    "chord-shapes-transposed": (song_chord_shapes_transposed, "portrait"),
    "fit-one-page": (song_fit_one_page, "portrait"),
    "tab-grid-landscape": (song_tab_grid_landscape, "landscape"),
    "marks-and-techniques": (song_marks_and_techniques, "portrait"),
    "parts": (song_parts, "portrait"),
}


GOLDEN_VERSION = "0.0.0"


def outputs(doc: dict, orient: str) -> dict:
    """The TXT and PDF for a golden song, with the footer date and the app
    version fixed, so a release does not change every golden file."""
    real = export.APP_VERSION
    export.APP_VERSION = GOLDEN_VERSION
    try:
        return {"txt": export.export_bytes(doc, "txt"),
                "pdf": export.export_bytes(doc, "pdf", orient=orient,
                                           date=GOLDEN_DATE)}
    finally:
        export.APP_VERSION = real


def main(argv) -> int:
    check = "--check" in argv
    os.makedirs(GOLDEN_DIR, exist_ok=True)
    changed = []
    for name, (build, orient) in SONGS.items():
        doc = build()
        files = {"sng": (json.dumps(doc, indent=2, ensure_ascii=False) + "\n").encode("utf-8")}
        files.update(outputs(doc, orient))
        for ext, data in files.items():
            path = os.path.join(GOLDEN_DIR, f"{name}.{ext}")
            old = open(path, "rb").read() if os.path.exists(path) else None
            if old != data:
                changed.append(os.path.relpath(path, ROOT))
                if not check:
                    export.write_export(path, data, overwrite=True)
    verb = "differ" if check else "written"
    print("\n".join(changed) if changed else "All golden files match.")
    if changed:
        print(f"{len(changed)} file(s) {verb}.")
    return 1 if (check and changed) else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))

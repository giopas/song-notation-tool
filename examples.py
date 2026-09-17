"""
examples.py — the built-in sample song behind "Open example".

Enhancement 1 of the UX spec: on first launch (or a fresh "New song")
a user should see a real, short, correctly-laid-out song rather than an
empty grid. Both webserver.py (web front end) and song_writer.py
(desktop app) call example_document() so the sample never drifts
between the two surfaces.
"""

from __future__ import annotations

import grammar
import model


def example_document() -> dict:
    doc = model.new_document(title="Example Song", artist="Song Notation Tool",
                              key="G", time="4/4", bpm="120")

    intro = model.new_section("intro", "Intro", "Intro", render="chart")
    intro["items"] = grammar.parse_items("G D Em C")

    verse = model.new_section("verse1", "Verse", "Verse", repeat=2, render="chart")
    verse_items, verse_annotation = grammar.parse('G D "keep it simple" Em C')
    verse["items"] = verse_items
    verse["annotation"] = verse_annotation or ""

    chorus = model.new_section("chorus1", "Chorus", "Chorus", render="chart")
    chorus["items"] = grammar.parse_items("[C D]x2 G")

    doc["sections"] = [intro, verse, chorus]
    return doc

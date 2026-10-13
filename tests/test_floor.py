"""The floor sheet (Size: Floor) and a section's own print size.

A floor sheet is read off the floor at your feet: the largest type that
fits on at most two pages, portrait or landscape, one or two columns,
with riff tab and tab grids at half the chart's size.
"""

import os
import re
import sys
import zlib

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import export
import grammar
import model
from constants import (FLOOR_MAX_PAGES, FLOOR_TAB_MIN, FLOOR_TAB_RATIO,
                       PDF_SCALE_LABELS, SECTION_SIZE_FACTORS)

RIFF = "{Hook = G - - - - | D - - 3 - | A 4 4 - 4 | E - - - -}"


def _section(sid, name, line, typ="Verse"):
    sec = model.new_section(sid, name, typ, instrument="Bass (4-string)")
    sec["items"] = grammar.parse_items(line)
    return sec


def _short_song():
    """Four short sections: the kind of song that fits one page easily."""
    doc = model.new_document(title="Paper Lanterns", artist="The Placeholders",
                             key="E", time="4/4")
    for i, line in enumerate(["3E 5E 3E", "3E 12E 3E 5E", "3E 12E", "3E 5E"]):
        doc["sections"].append(_section(f"s{i}", f"Part_{i}", line))
    return doc


def _long_song():
    """A dozen sections of short lines, two of them with a riff: a song
    that runs out of height long before it runs out of width."""
    doc = model.new_document(title="Tin Weather", artist="The Placeholders",
                             key="A", time="4/4")
    for i in range(12):
        line = (f"[C]x3 [F] // {RIFF} // [F] [C] [F] // {{Hook}} // 5A"
                if i % 3 == 1 else "rest 5A 5D // 5A 8F 8C // 5A 5D")
        doc["sections"].append(_section(f"s{i}", f"Part_{i}", line))
    return doc


def _font_sizes(pdf: bytes, text: str) -> list:
    """Every font size the PDF draws `text` at."""
    streams = b""
    for chunk in pdf.split(b"stream\n")[1:]:
        try:
            streams += zlib.decompress(chunk.split(b"\nendstream")[0])
        except zlib.error:
            pass
    body = streams.decode("latin-1")
    return [float(m.group(1)) for m in re.finditer(
        r"/F\d ([\d.]+) Tf [\d.\-]+ [\d.\-]+ Td \(" + re.escape(text) + r"\)", body)]


def _resolved(doc):
    ins = export.instruments_to_print(doc)
    orient = export.resolve_orient(doc, ins, "portrait")
    cols = export.resolve_columns(doc, ins, orient)
    scale = export.resolve_scale(doc, ins, orient, columns=cols)
    pages = export._build_pdf(doc, ins, orient, scale, cols)[1]
    return orient, cols, scale, pages


def test_floor_is_offered_as_a_size():
    assert "floor" in PDF_SCALE_LABELS


def test_floor_prints_bigger_than_fit_and_one_page():
    for make in (_short_song, _long_song):
        doc = make()
        sizes = {}
        for mode in ("fit", "one", "floor"):
            doc["pdf_scale"] = mode
            sizes[mode] = _resolved(doc)[2]
        assert sizes["floor"] > sizes["fit"], make.__name__
        assert sizes["floor"] > sizes["one"], make.__name__


def test_floor_never_needs_more_than_two_pages():
    for make in (_short_song, _long_song):
        doc = make()
        doc["pdf_scale"] = "floor"
        assert _resolved(doc)[3] <= FLOOR_MAX_PAGES


def test_floor_picks_the_orientation_with_the_bigger_type():
    for make in (_short_song, _long_song):
        doc = make()
        doc["pdf_scale"] = "floor"
        ins = export.instruments_to_print(doc)
        chosen = export.resolve_orient(doc, ins, "portrait")
        other = "landscape" if chosen == "portrait" else "portrait"
        best = export.resolve_scale(doc, ins, chosen,
                                    columns=export.resolve_columns(doc, ins, chosen))
        alt = export.resolve_scale(doc, ins, other,
                                   columns=export.resolve_columns(doc, ins, other))
        assert best >= alt


def test_other_sizes_keep_the_orientation_asked_for():
    doc = _short_song()
    doc["pdf_scale"] = "fit"
    assert export.resolve_orient(doc, None, "portrait") == "portrait"
    assert export.resolve_orient(doc, None, "landscape") == "landscape"


def test_floor_gives_a_short_song_the_full_width():
    doc = _short_song()
    doc["pdf_scale"] = "floor"
    doc["pdf_columns"] = "2"           # ignored: Floor decides
    assert _resolved(doc)[1] == 1


def test_floor_splits_a_long_song_into_two_columns():
    doc = _long_song()
    doc["pdf_scale"] = "floor"
    doc["pdf_columns"] = "1"           # ignored: Floor decides
    assert _resolved(doc)[1] == 2


def test_tab_scale_is_half_on_a_floor_sheet_with_a_floor_of_its_own():
    assert export.tab_scale(3.0, False) == 3.0
    assert export.tab_scale(3.0, True) == 3.0 * FLOOR_TAB_RATIO
    assert export.tab_scale(1.2, True) == FLOOR_TAB_MIN
    assert export.tab_scale(0.6, True) == 0.6       # never bigger than the chart


def test_floor_draws_riff_tab_at_half_the_chart_size():
    doc = _long_song()
    doc["pdf_scale"] = "floor"
    pdf = export.build_pdf(doc)
    chord = max(_font_sizes(pdf, "[F]"))
    tab = max(_font_sizes(pdf, "4"))
    assert tab < chord * 0.8


def test_riff_rows_stay_lined_up_when_drawn_smaller():
    """Every row of a riff is cut at the same columns, so its strings
    still line up under each other."""
    rows = ["Hook |G|---------|", "     |D|-----3---|"]
    assert [a for a, _ in export.lick_segments(rows[0], [])] == [0]
    assert [a for a, _ in export.lick_segments(rows[1], [])] == [0]
    # A riff sitting after chords on the line: cut where it starts.
    inline = "            |G|--5--|"
    assert export.lick_segments(inline, [12]) == [(12, "|G|--5--|")]


def test_a_larger_section_prints_bigger_than_its_neighbours():
    doc = _short_song()
    doc["pdf_scale"] = "1.0"
    doc["sections"][1]["print_size"] = "larger"
    doc["sections"][2]["print_size"] = "smaller"
    doc["sections"][1]["items"] = grammar.parse_items("Bb")
    doc["sections"][2]["items"] = grammar.parse_items("Gm")
    doc["sections"][3]["items"] = grammar.parse_items("Dm")
    pdf = export.build_pdf(doc)
    normal = _font_sizes(pdf, "Dm")[0]
    assert _font_sizes(pdf, "Bb")[0] == normal * SECTION_SIZE_FACTORS["larger"]
    assert _font_sizes(pdf, "Gm")[0] == normal * SECTION_SIZE_FACTORS["smaller"]


def test_a_larger_section_counts_wider_for_the_page_width():
    doc = _short_song()
    before = max(export._line_units(doc, None))
    doc["sections"][1]["print_size"] = "larger"
    assert max(export._line_units(doc, None)) > before


def test_section_size_defaults_to_normal():
    assert export.section_size_factor({}) == 1.0
    assert export.section_size_factor({"print_size": "nonsense"}) == 1.0
    assert export.section_size_factor({"print_size": "Larger"}) == SECTION_SIZE_FACTORS["larger"]


def test_floor_stays_inside_the_page_width():
    doc = _short_song()
    doc["sections"][0]["items"] = grammar.parse_items("3E 5E 3E 5E 7E 5E 3E 5E 3E 5E 7E 5E")
    doc["pdf_scale"] = "floor"
    orient, cols, scale, _ = _resolved(doc)
    longest = max(export._line_units(doc, None, scale))
    assert longest * export.MONO_CHAR_W * scale <= export.column_width(orient, cols) + 0.5

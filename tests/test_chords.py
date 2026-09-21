import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest

import chords
import model


def test_shape_reads_the_way_guitarists_write_one():
    """x32010 is C, lowest string first — and comes back highest first,
    the order it prints in."""
    assert chords.shape_from_text("x32010") == ["0", "1", "0", "2", "3", "x"]


def test_shape_accepts_spaced_two_digit_frets():
    assert chords.shape_from_text("x 0 5 5 5 x") == ["x", "5", "5", "5", "0", "x"]
    assert chords.shape_from_text("x-0-12-12-12-x")[1] == "12"


def test_shape_round_trips():
    for text in ("x32010", "320003"):
        frets = chords.shape_from_text(text)
        assert chords.shape_to_text(frets) == text


def test_shape_spaces_itself_out_when_a_fret_needs_two_digits():
    frets = chords.shape_from_text("x 0 12 12 12 x")
    assert chords.shape_to_text(frets) == "x 0 12 12 12 x"


def test_shape_must_have_one_fret_per_string():
    with pytest.raises(chords.ChordError):
        chords.shape_from_text("x3201")
    with pytest.raises(chords.ChordError):
        chords.shape_from_text("")


def test_bass_shape_has_four_frets():
    frets = chords.shape_from_text("5 7 7 x", "Bass (4-string)")
    assert frets == ["x", "7", "7", "5"]


def test_diagram_reads_as_tab():
    c = model.make_chord("C", chords.shape_from_text("x32010"))
    lines = chords.diagram_lines(c)
    assert lines[0].strip() == "C"
    assert lines[1].strip() == "e|-0-|"
    assert lines[-1].strip() == "E|-x-|"
    assert len({len(ln) for ln in lines}) == 1   # a column, not a ragged edge


def test_sheet_lays_diagrams_out_in_rows_that_fit():
    doc = {"chords": [model.make_chord(n, chords.shape_from_text("x32010"))
                      for n in ("C", "G", "Am", "F")]}
    wide = chords.sheet_lines(doc, width=100)
    narrow = chords.sheet_lines(doc, width=30)
    assert len(narrow) > len(wide)
    assert max(len(ln) for ln in narrow) <= 30


def test_sheet_groups_by_instrument():
    doc = {"chords": [
        model.make_chord("C", chords.shape_from_text("x32010")),
        model.make_chord("E", chords.shape_from_text("0 2 2 x", "Bass (4-string)"),
                          instrument="Bass (4-string)"),
    ]}
    out = "\n".join(chords.sheet_lines(doc))
    assert "Guitar (6-string)" in out and "Bass (4-string)" in out


def test_no_shapes_means_no_sheet():
    assert chords.sheet_lines({"chords": []}) == []
    assert chords.sheet_position({"chords": [], "chord_sheet": "end"}) == "none"


def test_sheet_position_defaults_to_none():
    doc = {"chords": [model.make_chord("C", chords.shape_from_text("x32010"))]}
    assert chords.sheet_position(doc) == "none"
    doc["chord_sheet"] = "start"
    assert chords.sheet_position(doc) == "start"


# ---------------------------------------------------------------------------
#  Coverage — shapes against the chords the chart plays
# ---------------------------------------------------------------------------

def _song(guitar_line="C G [Am F]x2", bass_line="5A 5D"):
    import grammar
    doc = model.new_document()
    g = model.new_section("v", "Verse", instrument="Guitar (6-string)")
    g["items"] = grammar.parse_items(guitar_line)
    b = model.new_section("b", "Bass", instrument="Bass (4-string)")
    b["items"] = grammar.parse_items(bass_line)
    doc["sections"] += [g, b]
    return doc


def test_used_symbols_reads_guitar_sections_and_looks_inside_groups():
    assert chords.used_symbols(_song()) == ["C", "G", "Am", "F"]


def test_a_bass_only_song_still_has_symbols():
    """No guitar sections at all means every section counts — otherwise
    a bass chart could never be given shapes for a guitarist."""
    doc = _song()
    doc["sections"] = doc["sections"][1:]
    assert chords.used_symbols(doc) == ["A", "D"]


def test_coverage_reports_both_directions():
    doc = _song()
    doc["chords"] = [model.make_chord("C", chords.shape_from_text("x32010")),
                     model.make_chord("E7", chords.shape_from_text("020100"))]
    cov = chords.coverage(doc)
    assert cov["missing"] == ["G", "Am", "F"]
    assert cov["unused"] == ["E7"]


def test_coverage_follows_the_printed_chart_after_a_transpose():
    """Moved up a tone, the page says D — so D is what needs a shape,
    and the C shape is for a chord nobody reads any more."""
    doc = _song()
    doc["chords"] = [model.make_chord("C", chords.shape_from_text("x32010"))]
    doc["transpose"] = 2
    cov = chords.coverage(doc)
    assert "D" in cov["missing"]
    assert cov["unused"] == ["C"]


def test_coverage_expands_references():
    import grammar
    doc = _song()
    ref = model.new_section("v2", "Verse 2", instrument="Guitar (6-string)")
    ref["items"] = grammar.parse_items("=Verse Em")
    doc["sections"].append(ref)
    assert "Em" in chords.used_symbols(doc)

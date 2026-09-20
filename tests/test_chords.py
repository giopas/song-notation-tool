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

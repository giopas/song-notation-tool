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
    """Moved up a tone, the chart says D and G says A — and the C shape
    moves with the song, so it covers the D rather than going stale."""
    doc = _song()
    doc["chords"] = [model.make_chord("C", chords.shape_from_text("x32010"))]
    doc["transpose"] = 2
    cov = chords.coverage(doc)
    assert "D" not in cov["missing"]
    assert "A" in cov["missing"]
    assert cov["unused"] == []


def test_coverage_expands_references():
    import grammar
    doc = _song()
    ref = model.new_section("v2", "Verse 2", instrument="Guitar (6-string)")
    ref["items"] = grammar.parse_items("=Verse Em")
    doc["sections"].append(ref)
    assert "Em" in chords.used_symbols(doc)


# ---------------------------------------------------------------------------
#  Transposing shapes — open to open, else barre; movable shapes slide
# ---------------------------------------------------------------------------

def _t(name, text, n, instrument="Guitar (6-string)"):
    c = model.make_chord(name, chords.shape_from_text(text, instrument),
                          instrument=instrument)
    out = chords.transpose_chord(c, n)
    return out["name"], chords.shape_to_text(out["frets"], instrument), out["voicing"]


def test_an_open_shape_goes_to_the_open_shape_of_the_new_chord():
    assert _t("C", "x32010", 2) == ("D", "xx0232", "open")
    assert _t("A", "x02220", -7) == ("D", "xx0232", "open")
    assert _t("Em", "022000", 5) == ("Am", "x02210", "open")


def test_an_open_shape_with_no_open_equivalent_becomes_a_barre():
    assert _t("C", "x32010", 5) == ("F", "133211", "barre")
    assert _t("Am", "x02210", 2) == ("Bm", "x24432", "barre")


def test_the_barre_is_the_lower_of_the_e_and_a_forms():
    """D# is fret 11 as an E-form and fret 6 as an A-form: the A-form."""
    assert _t("E", "022100", -1) == ("D#", "x68886", "barre")


def test_a_movable_shape_slides_and_keeps_its_voicing():
    assert _t("F", "133211", 2) == ("G", "355433", "moved")


def test_a_shape_that_would_fall_off_the_nut_goes_up_the_octave():
    name, shape, how = _t("F", "133211", -3)
    assert (name, how) == ("D", "moved")
    assert shape == "10 12 12 11 10 10"


def test_a_chord_with_no_template_is_slid_whole_like_a_capo():
    assert _t("Cadd9", "x32030", 1) == ("C#add9", "x43141", "capo")


def test_quality_spellings_fold_together():
    assert _t("Cmin", "x35543", 0)[2] == "as typed"
    assert _t("Emin7", "020000", 5) == ("Amin7", "x02010", "open")


def test_off_standard_tuning_shapes_only_slide():
    """The open and barre libraries are standard-tuning shapes; a bass
    or drop-D shape is slid instead of being swapped for one."""
    name, shape, how = _t("E", "0 2 2 x", 2, "Bass (4-string)")
    assert (name, how) == ("F#", "capo")
    assert shape == "244x"


def test_no_transpose_leaves_the_shape_alone():
    assert _t("C", "x32010", 0) == ("C", "x32010", "as typed")


def test_the_chord_sheet_prints_moved_with_the_song():
    doc = {"chords": [model.make_chord("C", chords.shape_from_text("x32010"))],
           "transpose": 2}
    out = "\n".join(chords.sheet_lines(doc))
    assert out.splitlines()[0].strip() == "D"
    assert "e|-2-|" in out and "A|-x-|" in out   # xx0232
    assert doc["chords"][0]["name"] == "C"       # the stored shape is untouched


def test_coverage_matches_shapes_as_printed():
    doc = _song(guitar_line="C G")
    doc["chords"] = [model.make_chord("C", chords.shape_from_text("x32010"))]
    doc["transpose"] = 2
    cov = chords.coverage(doc)
    assert "D" not in cov["missing"]   # the C shape prints as D
    assert cov["missing"] == ["A"]
    assert cov["unused"] == []


def test_stored_name_undoes_the_song_transpose():
    assert chords.stored_name("D", {"transpose": 2}) == "C"
    assert chords.stored_name("D", {"transpose": 0}) == "D"


def test_a_row_with_no_shape_yet_still_prints_under_the_moved_name():
    doc = {"chords": [{"name": "G", "instrument": "Guitar (6-string)",
                       "frets": [], "note": ""}], "transpose": 2}
    assert chords.printed_chords(doc)[0]["name"] == "A"


def test_from_chart_rows_cover_the_chart_once_added():
    """The round trip the dialog relies on: store each missing chord under
    its un-transposed name and the coverage report comes back clean."""
    doc = _song(guitar_line="C G Am")
    doc["transpose"] = 2
    for printed in chords.coverage(doc)["missing"]:
        doc.setdefault("chords", []).append(
            model.make_chord(chords.stored_name(printed, doc), []))
    assert chords.coverage(doc)["missing"] == []

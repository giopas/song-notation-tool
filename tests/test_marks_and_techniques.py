"""v0.30: playing techniques on licks and tab, dynamics, accents, runs
([pm ...], [cresc ...], [dim ...]), ending brackets and the coda."""
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import export  # noqa: E402
import grammar  # noqa: E402
import model  # noqa: E402
import render  # noqa: E402


def rows(line):
    items, _ = grammar.parse(line)
    return render.render_chart_rows(items)


def roundtrip(line):
    items, _ = grammar.parse(line)
    return grammar.unparse(items)


# ── grammar ──────────────────────────────────────────────────────────────

@pytest.mark.parametrize("line", [
    "mf 5A^ 7D",
    "pp A ff D",
    "[pm 5A 5A 5A] 7D",
    "[cresc A D E] [dim E D A]",
    "[pm 5A 5A]x2",
    "{G 5h7 7p5 5/7 7\\5 7b9 7b9r7 7~ x}",
])
def test_new_marks_survive_a_round_trip(line):
    assert roundtrip(line) == line


def test_dynamics_are_lower_case_so_f_stays_a_chord():
    items, _ = grammar.parse("F f")
    assert items[0]["kind"] == "token" and items[0]["symbol"] == "F"
    assert items[1] == model.make_mark("dyn_f")


def test_accent_is_stored_on_the_token():
    items, _ = grammar.parse("5A^ 7D")
    assert items[0]["accent"] is True
    assert "accent" not in items[1]


def test_a_run_is_a_group_with_a_span():
    items, _ = grammar.parse("[pm 5A 5A]")
    assert items[0]["kind"] == "group" and items[0]["span"] == "pm"
    assert len(items[0]["items"]) == 2


@pytest.mark.parametrize("line, words", [
    ("pm 5A", "first in brackets"),
    ("[pm]", "needs the notes"),
    ("5A^^", "cannot parse"),
    ("{ G 5q7 }", "not a fret"),
    ("{ G 30 }", "MAX_FRET"),
])
def test_mistakes_get_a_message(line, words):
    with pytest.raises(grammar.ParseError, match=words):
        grammar.parse(line)


@pytest.mark.parametrize("text, ok", [
    ("-", True), ("x", True), ("7", True), ("5h7", True), ("7b9r7", True),
    ("7~", True), ("12/14", True), ("25", False), ("5q", False), ("", False),
])
def test_valid_tab_position(text, ok):
    assert grammar.valid_tab_position(text) is ok


def test_dynamics_and_run_words_are_not_riff_names():
    for name in ("mf", "pp", "pm", "cresc", "dim"):
        assert not model.validate_block_name(name)


# ── rendering ────────────────────────────────────────────────────────────

def test_accent_prints_a_wedge_over_its_note():
    r = rows("mf 5A^ 7D")
    assert r[0]["role"] == render.ROLE_MARKS
    col = r[0]["text"].index(">")
    assert r[1]["text"][col] == "5"


def test_palm_mute_run_has_a_label_and_an_end_tick():
    r = rows("[pm 5A 5A 5A] 7D")
    assert r[0]["role"] == render.ROLE_MARKS
    assert r[0]["text"].strip().startswith("P.M.")
    assert r[0]["text"].rstrip().endswith("|")


def test_ending_brackets_carry_across_a_line_break():
    r = rows("|: A |1. D :| |2. E // F G")
    marks = [x for x in r if x["role"] == render.ROLE_MARKS]
    assert len(marks) == 2
    assert "1." in marks[0]["text"] and "2." in marks[0]["text"]
    # The second ending runs on to the end of the section.
    assert marks[1]["text"].strip().startswith("-")
    assert all(k == render.ROLE_ENDING for *_, k in marks[1]["spans"])


def test_coda_prints_as_a_boxed_label():
    r = rows("5A coda 7D")
    sym = [x for x in r if x["role"] == render.ROLE_SYM][0]
    assert "[+ Coda]" in sym["text"]
    assert any(k == render.ROLE_CODA for *_, k in sym["spans"])


def test_lick_techniques_print_as_typed():
    r = rows("{G 5h7 7b9r7 x}")
    lick = [x for x in r if x["role"] == render.ROLE_LICK][0]
    assert "5h7" in lick["text"] and "7b9r7" in lick["text"]


def test_transposing_a_tab_position_moves_every_fret():
    assert render.transpose_tab_position("5h7", 2) == "7h9"
    assert render.transpose_tab_position("x", 3) == "x"
    # A fret pushed past 24 leaves the position as it was.
    assert render.transpose_tab_position("23b24", 2) == "23b24"


def test_tab_token_width_grows_with_the_widest_cell():
    narrow = [model.make_measure(8, {"E": "5 - - - - - - -"})]
    wide = [model.make_measure(8, {"E": "7b9r7 - - - - - - -"})]
    assert render.tab_token_width([]) == 3
    assert render.tab_token_width(wide) > render.tab_token_width(narrow)


def test_tab_cells_that_are_not_frets_are_a_warning():
    doc = model.new_document("T")
    sec = model.new_section("s1", "Intro", render="tab")
    sec["items"] = [model.make_measure(8, {"E": "5 q - - - - - -"})]
    doc["sections"].append(sec)
    report = export.check_document(doc)
    assert report["ok"]
    assert any("not frets" in w for w in report["warnings"])


def test_the_pdf_draws_the_new_marks():
    doc = model.new_document("Marks")
    sec = model.new_section("s1", "Verse")
    sec["items"], _ = grammar.parse(
        "mf [pm 5A^ 5A] |: A |1. D :| |2. E coda [cresc A D]")
    doc["sections"].append(sec)
    pdf = export.build_pdf(doc)
    assert pdf.startswith(b"%PDF")
    text = export.build_song_lines(doc)
    assert any("P.M." in ln for ln in text)
    assert any("[+ Coda]" in ln for ln in text)

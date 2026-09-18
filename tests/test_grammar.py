import os
import sys
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from grammar import parse, parse_items, unparse, ParseError
from model import make_token, make_group, make_block_ref, make_section_ref, make_mark


# ==============================================================================
#  Section 4.3 — the fret must lead
# ==============================================================================

def test_fret_leads_5A():
    items = parse_items("5A")
    assert items == [make_token("A", 5)]


def test_no_fret_A5_power_chord():
    items = parse_items("A5")
    assert items == [make_token("A5", None)]


def test_fret_and_power_chord_5A5():
    items = parse_items("5A5")
    assert items == [make_token("A5", 5)]


def test_leading_zero_rejected():
    with pytest.raises(ParseError):
        parse_items("01E")


def test_open_string_zero_is_fine():
    items = parse_items("0E")
    assert items == [make_token("E", 0)]


def test_fret_over_max_rejected():
    with pytest.raises(ParseError):
        parse_items("25E")


def test_complex_chord_with_fret():
    items = parse_items("2F#m7")
    assert items == [make_token("F#m7", 2)]


# ==============================================================================
#  Examples straight from the design doc table (section 4.2)
# ==============================================================================

def test_seven_positioned_tokens():
    items = parse_items("5A 5D 5A 8F 8C 5A 5D")
    assert [it["symbol"] for it in items] == ["A", "D", "A", "F", "C", "A", "D"]
    assert [it["fret"] for it in items] == [5, 5, 5, 8, 8, 5, 5]


def test_bare_chord_no_position():
    items = parse_items("C")
    assert items == [make_token("C", None)]


def test_bracketed_group_with_repeat():
    items = parse_items("[5A 5D]x2")
    assert items == [make_group([make_token("A", 5), make_token("D", 5)], repeat=2)]


def test_block_ref_with_repeat():
    items = parse_items("riff1 x3")
    assert items == [make_block_ref("riff1", repeat=3)]


def test_block_ref_with_repeat_and_shift():
    items = parse_items("riff1 x3 +2")
    assert items == [make_block_ref("riff1", repeat=3, transpose=2)]


def test_section_ref_all_shorthand():
    items = parse_items("=verse1 x1 all")
    assert items == [make_section_ref("verse1", repeat=1, all=True)]


def test_endings():
    items = parse_items("|1. 5A |2. 3C")
    assert items[0]["kind"] == "mark" and items[0]["mark"] == "ending_1"
    assert items[1] == make_token("A", 5)
    assert items[2]["kind"] == "mark" and items[2]["mark"] == "ending_2"
    assert items[3] == make_token("C", 3)


def test_annotation_returned_separately():
    items, annotation = parse('"stacco sul secondo"')
    assert items == []
    assert annotation == "stacco sul secondo"


# ==============================================================================
#  Section 4.4 — block name ambiguity is a *creation-time* check, not parsing;
#  but a bare identifier still parses as a block_ref while a note-looking
#  word parses as a token, never a block.
# ==============================================================================

def test_note_looking_word_is_a_token_not_a_block():
    items = parse_items("A")
    assert items[0]["kind"] == "token"


def test_real_block_name_is_a_block_ref():
    items = parse_items("riff1")
    assert items[0]["kind"] == "block_ref"


# ==============================================================================
#  Round trip — the acceptance test.  parse(unparse(items)) == items
# ==============================================================================

ROUND_TRIP_FIXTURES = [
    [make_token("A", 5)],
    [make_token("A5", None)],
    [make_token("F#m7", 2)],
    [make_token("A", 5), make_token("D", 5), make_token("A", 5)],
    [make_group([make_token("A", 5), make_token("D", 5)], repeat=2)],
    [make_block_ref("riff1", repeat=3)],
    [make_block_ref("riff1", repeat=3, transpose=2)],
    [make_block_ref("riff1", repeat=1, transpose=-1)],
    [make_section_ref("verse1", repeat=1, all=True)],
    [make_section_ref("verse1", repeat=4)],
    [make_mark("repeat_open"), make_token("A", 5), make_mark("repeat_close")],
    [make_mark("ending_1"), make_token("A", 5), make_mark("ending_2"), make_token("C", 3)],
    [make_mark("coda")],
    [make_mark("segno")],
    [make_mark("simile")],
    [
        make_group([make_token("A", 5), make_group([make_token("D", 7)], repeat=2)], repeat=1),
    ],
]


@pytest.mark.parametrize("items", ROUND_TRIP_FIXTURES)
def test_round_trip(items):
    line = unparse(items)
    assert parse_items(line) == items


def test_round_trip_all_fixtures_together():
    for items in ROUND_TRIP_FIXTURES:
        assert parse_items(unparse(items)) == items


# ==============================================================================
#  Licks — inline tab figures — v0.20
# ==============================================================================

def test_lick_parses_into_string_lines():
    items = parse_items("C {G 5 7 5 | D - - 3} G")
    lick = items[1]
    assert lick["kind"] == "lick"
    assert [ln["string"] for ln in lick["lines"]] == ["G", "D"]
    assert lick["lines"][0]["frets"] == ["5", "7", "5"]
    assert lick["lines"][1]["frets"] == ["-", "-", "3"]


def test_lick_round_trips():
    line = "C {G 5 7 5 | D - - 3} G"
    assert unparse(parse_items(line)) == line


def test_lick_accepts_a_colon_after_the_string_name():
    a = parse_items("{G: 5 7 5}")
    b = parse_items("{G 5 7 5}")
    assert a == b


def test_lick_takes_muted_and_unplayed_positions():
    frets = parse_items("{A 5 x - 7}")[0]["lines"][0]["frets"]
    assert frets == ["5", "x", "-", "7"]


def test_lick_rejects_a_bad_fret():
    for bad in ("{G 5 99}", "{G 5 z}", "{G}", "{}"):
        with pytest.raises(ParseError):
            parse_items(bad)


def test_unmatched_lick_brace_is_an_error():
    with pytest.raises(ParseError):
        parse_items("C {G 5 7")


def test_curly_quotes_are_accepted_as_annotation_quotes():
    """macOS and every word processor substitute typed quotes; a line
    pasted from anywhere is likely to carry them."""
    for lo, hi in (("“", "”"), ("‘", "’"), ('"', '"')):
        items, ann = parse(f'C {lo}keep it simple{hi} G')
        assert ann == "keep it simple"
        assert [it["symbol"] for it in items] == ["C", "G"]

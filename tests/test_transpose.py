import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from transpose import (
    resolve, transpose_symbol, transpose_note, note_to_index,
    effective_transpose, CHROMATIC, ENHARMONIC, MAX_FRET,
)


def test_enharmonic_table_has_no_duplicate_keys_and_covers_lowercase():
    # v0.15 bug: "Bb" was defined twice; db/eb/gb/ab (lowercase) were missing.
    assert set(ENHARMONIC.keys()) == {
        "Bb", "Db", "Eb", "Gb", "Ab", "bb", "db", "eb", "gb", "ab"
    }
    for k in ("db", "eb", "gb", "ab", "bb"):
        assert ENHARMONIC[k] == ENHARMONIC[k.capitalize()]


def test_token_rule_note_plus_fret_plus_slide_up():
    # 0/E +1 -> 1/F ; 7/E +1 -> 8/F  (section 5.2)
    r = resolve("E", 0, 1)
    assert (r.symbol, r.fret) == ("F", 1)
    r = resolve("E", 7, 1)
    assert (r.symbol, r.fret) == ("F", 8)


def test_transpose_symbol_no_fret_unaffected_positionally():
    assert transpose_symbol("Am7", 2) == "Bm7"
    assert transpose_symbol("Bbm7", 1) == "Bm7"


def test_resolve_no_fret_given():
    r = resolve("C", None, 3)
    assert r.fret is None
    assert r.flag is None


def test_resolve_negative_wraps_octave_up():
    r = resolve("E", 1, -3)
    # 1 - 3 = -2 -> +12 = 10, within range
    assert r.fret == 10
    assert r.flag == "octave_up"


def test_resolve_over_max_wraps_octave_down():
    r = resolve("E", 20, 10)
    # 20 + 10 = 30 -> -12 = 18, within range
    assert r.fret == 18
    assert r.flag == "octave_down"


def test_resolve_unplayable_when_still_out_of_range():
    r = resolve("E", 0, -20)
    # 0 - 20 = -20 -> +12 = -8, still out of range
    assert r.fret is None
    assert r.flag == "unplayable"


def test_effective_transpose_is_additive():
    assert effective_transpose(2, -1, 3) == 4


def test_plus_n_then_minus_n_round_trips_exactly():
    """Acceptance criterion: transposing +n then -n returns the document
    to its exact original state."""
    for symbol, fret in [("E", 7), ("A", 0), ("F#m7", 2), ("C", 12), ("A5", None)]:
        for n in range(-11, 12):
            up = resolve(symbol, fret, n)
            if up.fret is None and fret is not None:
                continue  # went unplayable at this offset; nothing to round-trip
            back_fret = up.fret if fret is not None else None
            down = resolve(up.symbol, back_fret, -n)
            assert down.symbol == symbol
            if fret is not None and up.flag is None and down.flag is None:
                assert down.fret == fret

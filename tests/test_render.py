import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from render import (
    section_has_tab, string_has_content, active_strings,
    measures_per_line, render_chart_row,
)
from model import make_token, make_group, make_block_ref


def test_section_has_tab_false_for_all_dash_cells():
    cells = [{"G": "- - - - - - - -", "D": "- - - - - - - -",
              "A": "- - - - - - - -", "E": "- - - - - - - -"}] * 4
    assert section_has_tab(cells) is False


def test_section_has_tab_true_when_any_real_fret_present():
    cells = [{"G": "- - - - - - - -", "E": "- 7 - - - - - -"}]
    assert section_has_tab(cells) is True


def test_section_has_tab_false_for_empty_list():
    assert section_has_tab([]) is False
    assert section_has_tab(None) is False


def test_interlude_notes_only_section_has_no_tab():
    """The v0.15 bug fixture: an 'Interlude' section with note names and an
    empty tab must be has_tab=False for BOTH exporters (section 7.1)."""
    cells = [{"G": "-", "D": "-", "A": "-", "E": "-"} for _ in range(4)]
    assert section_has_tab(cells) is False


def test_string_has_content_and_active_strings_drop_empty():
    cells = [
        {"G": "- - - - - - - -", "D": "- - - - - - - -",
         "A": "- - - - 7 - 9 -", "E": "- 7 - - - - - -"},
    ]
    assert string_has_content(cells, "G") is False
    assert string_has_content(cells, "E") is True
    assert active_strings(cells, ["G", "D", "A", "E"]) == ["A", "E"]


def test_measures_per_line_fits_width():
    # 8 measures of 8 beats each, at TOKEN_W=3 -> col width 25 chars.
    n = measures_per_line(8, max_beats=8, width=100)
    assert 1 <= n <= 8
    # Narrower width must never fit more measures than a wider one.
    assert measures_per_line(8, max_beats=8, width=40) <= n


def test_measures_per_line_always_at_least_one():
    assert measures_per_line(8, max_beats=64, width=20) == 1


def test_render_chart_row_two_lines_aligned():
    items = [make_token("E", 0), make_token("E", 7), make_token("A", 5), make_token("F", 8)]
    lines = render_chart_row(items, label="Verse")
    assert len(lines) == 2
    fret_line, sym_line = lines
    assert "0" in fret_line and "7" in fret_line
    assert "E" in sym_line and "A" in sym_line and "F" in sym_line
    assert sym_line.startswith("Verse")


def test_render_chart_row_block_ref_shows_repeat():
    items = [make_block_ref("riff1", repeat=3)]
    _, sym_line = render_chart_row(items, label="Chorus")
    assert "riff1" in sym_line and "x3" in sym_line

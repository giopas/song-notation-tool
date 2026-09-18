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


# ==============================================================================
#  mark_spans / has_coda  (section 6 — endings and the coda box)
# ==============================================================================

from render import mark_spans, has_coda, estimate_section_lines, estimate_page_count
from model import make_mark


def test_mark_spans_empty_for_no_endings():
    items = [make_token("A"), make_token("B")]
    assert mark_spans(items) == []


def test_mark_spans_resolves_two_endings():
    items = [
        make_token("A"), make_token("B"),
        make_mark("ending_1"), make_token("C"),
        make_mark("ending_2"), make_token("D"),
    ]
    spans = mark_spans(items)
    assert spans == [("1", 3, 3), ("2", 5, 5)]


def test_mark_spans_last_ending_runs_to_end_of_section():
    items = [make_mark("ending_1"), make_token("A"), make_token("B")]
    assert mark_spans(items) == [("1", 1, 2)]


def test_mark_spans_ignores_empty_span():
    # Two ending marks back to back with nothing between them.
    items = [make_mark("ending_1"), make_mark("ending_2"), make_token("A")]
    assert mark_spans(items) == [("2", 2, 2)]


def test_has_coda_true_only_when_coda_mark_present():
    assert has_coda([make_token("A"), make_mark("coda")]) is True
    assert has_coda([make_token("A")]) is False
    assert has_coda([]) is False


# ==============================================================================
#  Page-count estimate (section 3.1)
# ==============================================================================

def test_estimate_section_lines_zero_for_empty_section():
    assert estimate_section_lines({"items": [], "annotation": "", "render": "chart"}) == 0


def test_estimate_section_lines_counts_chart_rows():
    sec = {"items": [make_token("A", 5), make_token("B", 7)],
           "annotation": "", "render": "chart"}
    # 1 label + 2 chart rows + 1 trailing blank
    assert estimate_section_lines(sec) == 4


def test_estimate_section_lines_includes_annotation():
    sec = {"items": [make_token("A", 5)], "annotation": "stacco sul secondo",
           "render": "chart"}
    assert estimate_section_lines(sec) == 5


def test_estimate_section_lines_ignores_lyrics_when_print_lyrics_false():
    sec = {"items": [make_token("A", 5)], "annotation": "", "render": "chart",
           "lyrics_text": "verse one\nverse two", "print_lyrics": False}
    # 1 label + 2 chart rows + 1 trailing blank — same as with no lyrics at all
    assert estimate_section_lines(sec) == 4


def test_estimate_section_lines_counts_lyrics_when_print_lyrics_true():
    sec = {"items": [make_token("A", 5)], "annotation": "", "render": "chart",
           "lyrics_text": "verse one\nverse two", "print_lyrics": True}
    # 1 label + 2 chart rows + (2 lyric lines + 1 lyric-block blank) + 1 trailing blank
    assert estimate_section_lines(sec) == 7


def test_estimate_page_count_at_least_one_for_nonempty_doc():
    doc = {"sections": [
        {"items": [make_token("A", 5)], "annotation": "", "render": "chart"},
    ]}
    assert estimate_page_count(doc) == 1


def test_estimate_page_count_zero_sections_is_one_page():
    assert estimate_page_count({"sections": []}) == 1


def test_estimate_page_count_grows_with_many_sections():
    sec = {"items": [make_token("A", 5)] * 20, "annotation": "", "render": "chart"}
    doc = {"sections": [dict(sec) for _ in range(40)]}
    assert estimate_page_count(doc, lines_per_page=10) > 1


# ==============================================================================
#  resolve_display_items — render-time, non-destructive transpose (section 5)
# ==============================================================================

from render import resolve_display_items
from model import make_token, make_group, make_block_ref


def test_resolve_display_items_noop_at_zero_offset():
    items = [make_token("A", 5)]
    out = resolve_display_items(items, 0)
    assert out == items
    assert out is not items or out == items  # no crash either way


def test_resolve_display_items_shifts_token_fret_and_symbol():
    items = [make_token("A", 5)]
    out = resolve_display_items(items, 2)
    assert out[0]["fret"] == 7
    assert out[0]["symbol"] == "B"


def test_resolve_display_items_does_not_mutate_input():
    items = [make_token("A", 5)]
    resolve_display_items(items, 2)
    assert items[0]["fret"] == 5
    assert items[0]["symbol"] == "A"


def test_resolve_display_items_recurses_into_groups():
    items = [make_group([make_token("A", 5), make_token("E", 7)])]
    out = resolve_display_items(items, 2)
    assert out[0]["items"][0]["fret"] == 7
    assert out[0]["items"][1]["fret"] == 9


def test_resolve_display_items_leaves_block_ref_unchanged():
    items = [make_block_ref("riff1", repeat=3)]
    out = resolve_display_items(items, 5)
    assert out == items


# ==============================================================================
#  Free-text line estimates — v0.20
# ==============================================================================

def test_free_text_line_count_counts_lines_plus_separator():
    from render import free_text_line_count
    assert free_text_line_count("") == 0
    assert free_text_line_count("   \n  ") == 0
    assert free_text_line_count("one") == 2
    assert free_text_line_count("one\ntwo\nthree") == 4


def test_estimate_section_lines_counts_free_text():
    from render import estimate_section_lines
    sec = {"render": "free", "free_text": "a\nb\nc", "items": []}
    # label + 3 lines + separator + trailing blank
    assert estimate_section_lines(sec) == 6


def test_estimate_section_lines_ignores_free_text_when_not_free_mode():
    from render import estimate_section_lines
    sec = {"render": "chart", "free_text": "a\nb\nc", "items": []}
    assert estimate_section_lines(sec) == 0

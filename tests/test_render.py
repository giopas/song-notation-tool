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
    # No item here carries a fret, so there is no fret row to emit — the
    # blank one used to be printed anyway (v0.20 fix).
    rows = render_chart_row(items, label="Chorus")
    assert len(rows) == 1
    sym_line = rows[0]
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


# ==============================================================================
#  Reference expansion — v0.20
# ==============================================================================

def _doc_with_ref(ref_key, **ref_kw):
    import model, grammar
    doc = model.new_document(title="T")
    # id minted as "chorus1", section later renamed "Interlude" — the exact
    # shape that made a reference read as a stale, wrong-looking id.
    target = model.new_section("chorus1", "Interlude", "Interlude")
    target["items"] = grammar.parse_items("3A 12D")
    src = model.new_section("s2", "Interlude (ref)", "Interlude")
    src["items"] = [model.make_section_ref(ref_key, **ref_kw)]
    doc["sections"] = [target, src]
    return doc, src


def test_section_ref_expands_to_the_target_items():
    from render import resolve_references
    doc, src = _doc_with_ref("chorus1")
    out = resolve_references(src["items"], doc)
    assert [it["symbol"] for it in out] == ["A", "D"]
    assert [it["fret"] for it in out] == [3, 12]


def test_section_ref_resolves_by_name_as_well_as_id():
    from render import resolve_references
    doc, src = _doc_with_ref("Interlude")
    out = resolve_references(src["items"], doc)
    assert [it["symbol"] for it in out] == ["A", "D"]


def test_repeated_section_ref_expands_into_a_repeated_group():
    from render import resolve_references
    doc, src = _doc_with_ref("chorus1", repeat=3)
    out = resolve_references(src["items"], doc)
    assert len(out) == 1
    assert out[0]["kind"] == "group" and out[0]["repeat"] == 3
    assert [it["symbol"] for it in out[0]["items"]] == ["A", "D"]


def test_shifted_section_ref_expands_transposed():
    from render import resolve_references
    doc, src = _doc_with_ref("chorus1", transpose=2)
    out = resolve_references(src["items"], doc)
    assert [it["symbol"] for it in out] == ["B", "E"]


def test_unresolvable_section_ref_is_left_as_a_reference():
    from render import resolve_references
    doc, src = _doc_with_ref("nope")
    out = resolve_references(src["items"], doc)
    assert out[0]["kind"] == "section_ref" and out[0]["section"] == "nope"


def test_reference_cycle_does_not_recurse_forever():
    import model
    from render import resolve_references, render_chart_row
    a = model.new_section("a", "A", "Verse"); a["items"] = [model.make_section_ref("b")]
    b = model.new_section("b", "B", "Verse"); b["items"] = [model.make_section_ref("a")]
    doc = model.new_document(title="Cycle"); doc["sections"] = [a, b]
    out = resolve_references(a["items"], doc)      # must terminate
    assert render_chart_row(out)


def test_block_ref_expands_from_the_riff_library():
    import model, grammar
    from render import resolve_references
    doc = model.new_document(title="T")
    doc["blocks"] = {"riff1": {"id": "riff1", "name": "riff1",
                               "items": grammar.parse_items("5A 7D")}}
    items = [model.make_block_ref("riff1")]
    out = resolve_references(items, doc)
    assert [it["fret"] for it in out] == [5, 7]


def test_reference_inside_a_group_is_expanded_too():
    import model, grammar
    from render import resolve_references
    doc, _ = _doc_with_ref("chorus1")
    items = [model.make_group([model.make_section_ref("chorus1")], repeat=2)]
    out = resolve_references(items, doc)
    assert [it["symbol"] for it in out[0]["items"]] == ["A", "D"]


def test_group_symbol_row_keeps_fret_numbers():
    import model, grammar
    from render import render_chart_row
    rows = render_chart_row([model.make_group(grammar.parse_items("3A 12D"), repeat=2)])
    assert "3A 12D" in "\n".join(rows)


def test_a_line_break_inside_a_group_actually_breaks_the_line():
    # [A B // C D]x4 used to collapse the "//" into literal text inside
    # the brackets instead of starting a new line — the group swallowed
    # the break the same way it used to swallow a lick's tab. It now
    # breaks properly AND gets a right-hand bracket spanning every chord
    # line it produced, with "(x4)" on the bracket rather than glued onto
    # whichever chord happens to end the last line — see
    # test_the_repeat_bracket_does_not_read_as_belonging_to_one_chord
    # below for why that distinction matters.
    #
    # A fret row (each line's own "7  4  3  5") gets no bracket segment
    # of its own — it's set as a superscript over the chord row beneath
    # it, its own colour and size, not a line in its own right — see
    # test_a_lick_inside_a_group_still_prints_its_tab for the same rule
    # on a group's very first line.
    import grammar
    from render import render_chart_row
    rows = render_chart_row(grammar.parse_items("[7B 4G# 3G 5D // 7B 4G# 5A]x4"))
    assert rows == [
        "  7  4   3  5",
        "  B  G#  G  D  | (x4)",
        "  7  4   5",
        "  B  G#  A     |",
    ]
    assert "//" not in "\n".join(rows)


def test_the_repeat_bracket_does_not_read_as_belonging_to_one_chord():
    # The whole point of the bracket: "(x4)" must not sit on the same
    # line as, or right after, any one chord — that reads as "this one
    # note repeats", which is exactly the ambiguity a bare trailing
    # "B  G#  A (x4)" used to create. With two chord lines to choose
    # from, it goes on the first rather than the last, for the same
    # reason.
    import grammar
    from render import render_chart_row
    rows = render_chart_row(grammar.parse_items("[7B 4G# 3G 5D // 7B 4G# 5A]x4"))
    last_chord_line = next(r for r in rows if r.rstrip().endswith("A     |"))
    assert "(x4)" not in last_chord_line
    # every chord line carries the bar (fret rows don't), so the span
    # still reads as one unit
    chord_lines = [r for r in rows if r.strip().startswith(("B", "G", "A", "D"))]
    assert all(r.rstrip().endswith("|") or "(x4)" in r for r in chord_lines)


def test_group_with_no_break_still_collapses_to_one_bracketed_line():
    # A group without its own "//" is unaffected by the flatten step —
    # it still renders as the single bracketed column it always did.
    import grammar
    from render import render_chart_row
    rows = render_chart_row(grammar.parse_items("[5A 7D]x2"))
    assert rows == ["  [5A 7D](x2)"]


def test_repeated_section_reference_with_a_break_also_breaks_the_line():
    # The pre-existing bug in _expanded(): a repeated section/riff
    # reference expands into the same "group" item a literal [ ]xN does,
    # so it inherited the exact same swallowed-"//" bug.
    import model, grammar
    from render import resolve_references, chart_body_lines

    doc = model.new_document(title="T")
    target = model.new_section("verse1", "Verse 1", "Verse",
                                instrument="Bass (4-string)")
    target["items"] = grammar.parse_items("7B 4G# 3G 5D // 7B 4G# 5A")
    ref_items = [model.make_section_ref("verse1", repeat=4)]
    doc["sections"] = [target, model.new_section("s2", "Verse 1 (ref)", "Verse")]

    resolved = resolve_references(ref_items, doc)
    lines = chart_body_lines(resolved, label="Verse1ref")
    assert lines == [
        "           7  4   3  5",
        "Verse1ref  B  G#  G  D  | (x4)",
        "           7  4   5",
        "           B  G#  A     |",
    ]
    assert "//" not in "\n".join(lines)


def test_broken_group_line_count_matches_the_page_estimate_heuristic():
    # estimate_section_lines' no-doc heuristic has to flatten the same
    # way real rendering does, or a section with a broken group inside a
    # repeat is undercounted and risks a mid-phrase page split.
    import grammar
    from render import estimate_section_lines, chart_body_lines
    items = grammar.parse_items("[7B 4G# 3G 5D // 7B 4G# 5A]x4")
    sec = {"render": "chart", "items": items}
    assert estimate_section_lines(sec) == 1 + len(chart_body_lines(items)) + 1


def test_unlabelled_chart_row_uses_the_export_body_indent():
    """A chart section and a free-text section under the same header must
    start in the same column; they were 4 and 2 apart."""
    from render import render_chart_row, BODY_INDENT
    import grammar
    fret, sym = render_chart_row(grammar.parse_items("3A 12D"))
    assert sym.startswith(" " * BODY_INDENT + "3")  is False   # fret row is separate
    assert sym[:BODY_INDENT] == " " * BODY_INDENT
    assert sym[BODY_INDENT] != " "


def test_labelled_chart_row_still_gets_its_own_gutter():
    from render import render_chart_row
    import grammar
    _, sym = render_chart_row(grammar.parse_items("3A 12D"), label="Verse")
    assert sym.startswith("Verse")


# ==============================================================================
#  A reference to a free-text section — v0.20
# ==============================================================================

def _doc_ref_to_free():
    import model, grammar
    doc = model.new_document(title="T")
    target = model.new_section("chorus1", "Interlude", "Interlude", render="free")
    target["free_text"] = "|--|3-|  |--|12-|"
    # Switching a section to Free leaves its old chart items in place on
    # purpose. Expanding a reference must NOT resurrect them.
    target["items"] = grammar.parse_items("[C D]x2 G")
    ref = model.new_section("s9", "Interlude (ref)", "Interlude")
    ref["items"] = [model.make_section_ref("chorus1")]
    doc["sections"] = [target, ref]
    return doc, ref


def test_reference_to_a_free_section_expands_its_text_not_its_dormant_items():
    from render import resolve_references, chart_body_lines
    doc, ref = _doc_ref_to_free()
    out = resolve_references(ref["items"], doc)
    assert [it["kind"] for it in out] == ["text"]
    lines = chart_body_lines(out)
    assert lines == ["  |--|3-|  |--|12-|"]
    assert "C" not in "\n".join(lines)


def test_repeated_reference_to_a_free_section_marks_the_repeat():
    import model
    from render import resolve_references, chart_body_lines
    doc, _ = _doc_ref_to_free()
    out = resolve_references([model.make_section_ref("chorus1", repeat=2)], doc)
    assert chart_body_lines(out)[-1].endswith("(x2)")


def test_reference_to_an_empty_free_section_stays_a_reference():
    import model
    from render import resolve_references
    doc, _ = _doc_ref_to_free()
    doc["sections"][0]["free_text"] = "   "
    out = resolve_references([model.make_section_ref("chorus1")], doc)
    assert out[0]["kind"] == "section_ref"


def test_multiline_free_text_keeps_its_own_indent_under_a_label():
    from render import chart_body_lines, make_text
    lines = chart_body_lines([make_text("one\ntwo")], label="Verse")
    assert lines[0].startswith("Verse")
    assert lines[0].endswith("one")
    assert lines[1] == " " * len(lines[0][:lines[0].index("one")]) + "two"


def test_chart_row_with_no_frets_emits_one_line_not_a_blank_one():
    import grammar
    from render import render_chart_row
    assert len(render_chart_row(grammar.parse_items("C G Am"))) == 1
    assert len(render_chart_row(grammar.parse_items("3A 12D"))) == 2


# ==============================================================================
#  Licks render in among the chords — v0.20
# ==============================================================================

def test_lick_stacks_under_its_own_column_in_the_chart_row():
    import grammar
    from render import render_chart_row
    rows = render_chart_row(grammar.parse_items("C {G 5 7 5 | D - - 3} G Am"))
    assert len(rows) == 3                       # symbol row + two string lines
    sym, g_row, d_row = rows
    assert "C" in sym and "G" in sym and "Am" in sym
    # written the way tab is written, one row per string
    assert g_row.strip() == "|G|-5-7-5-|"
    assert d_row.strip() == "|D|-----3-|"
    # the lick sits in its own column, after C and before the following G
    assert sym.index("C") < g_row.index("|G|")


def test_a_row_without_a_lick_has_no_extra_lines():
    import grammar
    from render import render_chart_row
    assert len(render_chart_row(grammar.parse_items("C G Am"))) == 1
    assert len(render_chart_row(grammar.parse_items("5A 7D"))) == 2


def test_a_lick_inside_a_group_still_prints_its_tab():
    # A group repeat wraps a whole phrase — chords and a riff together —
    # not just bare chords. Nesting a lick inside [ ]xN used to collapse
    # it down to its bare name ("[3G 2F# Riff3](x4)"); the tab itself,
    # the reason the lick was written out at all, has to survive that.
    #
    # A group holding a lick is unwrapped entirely (see
    # _flatten_bracket_groups) rather than kept as one bracketed column,
    # because a lick's tab needs rows of its own under it that a single
    # bracketed line has no room for. The chords and the tab both print
    # in full, and a bracket on the right — covering the chords AND the
    # tab — carries the repeat instead of any inline "[...](x4)" text.
    #
    # The fret row above the pickup ("3  2") gets no bracket of its own
    # either — see test_a_line_break_inside_a_group_actually_breaks_the_line
    # for the same rule when a group breaks across more than one line.
    import grammar
    from render import render_chart_row
    line = "[3G 2F# {Riff3 = G - - - | D 4 - - | A - 4 5 | E - - -}]x4"
    rows = render_chart_row(grammar.parse_items(line))
    assert rows == [
        "  3  2",
        "  G  F#                     |",
        "         Riff3 |G|-------|  |",
        "               |D|-4-----|  | (x4)",
        "               |A|---4-5-|  |",
        "               |E|-------|  |",
    ]


def test_the_bracket_never_lands_on_a_fret_row():
    # A fret row is drawn as a superscript over the chord row under it —
    # its own colour, its own smaller size, tighter line spacing — in
    # every renderer that reads render_chart_rows' "role" field (PDF
    # export in particular). A bracket "|" character landing there used
    # to inherit all of that: it showed up small and off-colour, and out
    # of line with the plain black bar on every other row. So a fret row
    # never gets one, whether it's the very first line of a bracketed
    # group (a pickup's fret numbers) or one reappearing after a "//".
    import grammar
    from render import render_chart_rows, ROLE_FRET
    for line in (
        "[3G 2F# {Riff3 = G - - - | D 4 - - | A - 4 5 | E - - -}]x4",
        "[7B 4G# 3G 5D // 7B 4G# 5A]x4",
    ):
        rows = render_chart_rows(grammar.parse_items(line))
        assert any(r["role"] == ROLE_FRET for r in rows)   # sanity: the case applies
        for r in rows:
            if r["role"] == ROLE_FRET:
                assert "|" not in r["text"]


def test_a_lick_only_group_gets_a_bracket_not_empty_brackets():
    # No chords means no "[...]" text makes sense at all (there'd be
    # nothing inside it) — the whole tab gets a right-hand bracket
    # instead, same as any other group needing more than one row.
    import grammar
    from render import render_chart_row
    rows = render_chart_row(grammar.parse_items("[{Riff3 = G 5 | D 3}]x4"))
    assert rows == [
        "                 |",
        "  Riff3 |G|-5-|  | (x4)",
        "        |D|-3-|  |",
        "                 |",
    ]
    assert "[]" not in "\n".join(rows)


def test_group_with_two_licks_plays_them_in_sequence_side_by_side():
    # A group holding two licks with nothing between them is unwrapped
    # (see _flatten_bracket_groups) into the same two consecutive
    # top-level items {Riff1} {Riff2} would be outside any group — so
    # they sit side by side, in playing order left to right, exactly as
    # top-level licks already do (see
    # test_licks_of_different_heights_share_the_row). A bracket wraps
    # the whole row, with the group's "(x2)" once on it.
    import grammar
    from render import render_chart_row
    line = "[{Riff1 = G 5} {Riff2 = D 3}]x2"
    rows = render_chart_row(grammar.parse_items(line))
    riff1_row = next(i for i, r in enumerate(rows) if "Riff1" in r)
    riff2_row = next(i for i, r in enumerate(rows) if "Riff2" in r)
    assert riff1_row == riff2_row
    assert rows[riff1_row].index("Riff1") < rows[riff1_row].index("Riff2")
    assert rows[riff1_row].endswith("(x2)")
    assert all(r.rstrip().endswith("|") for i, r in enumerate(rows) if i != riff1_row)


def test_licks_of_different_heights_share_the_row():
    import grammar
    from render import render_chart_row
    rows = render_chart_row(grammar.parse_items("{G 5} C {e 1 | B 2 | G 3}"))
    assert len(rows) == 4                       # symbol row + three string lines


def test_transposing_moves_a_licks_frets_not_its_strings():
    import grammar
    from render import resolve_display_items
    out = resolve_display_items(grammar.parse_items("{G 5 7 5 | D - - 3}"), 2)
    assert [ln["string"] for ln in out[0]["lines"]] == ["G", "D"]
    assert out[0]["lines"][0]["frets"] == ["7", "9", "7"]
    assert out[0]["lines"][1]["frets"] == ["-", "-", "5"]


def test_transposing_leaves_a_lick_fret_that_would_fall_off_the_neck():
    import grammar
    from render import resolve_display_items
    out = resolve_display_items(grammar.parse_items("{G 0 23}"), -2)
    assert out[0]["lines"][0]["frets"] == ["0", "21"]   # 0-2 would be negative


def test_lick_lines_are_counted_in_the_page_estimate():
    import grammar
    from render import estimate_section_lines
    plain = {"render": "chart", "items": grammar.parse_items("C G")}
    with_lick = {"render": "chart", "items": grammar.parse_items("C {G 5 | D 3}")}
    assert estimate_section_lines(with_lick) == estimate_section_lines(plain) + 2


def test_lick_inside_a_group_is_counted_in_the_page_estimate_too():
    # Same heuristic (no doc, so no real render) has to see a lick that's
    # nested inside a [ ]xN group, or the export's keep-a-section-whole
    # pagination undercounts a repeated riff and splits it across pages.
    import grammar
    from render import estimate_section_lines
    plain = {"render": "chart", "items": grammar.parse_items("C G")}
    grouped = {"render": "chart",
               "items": grammar.parse_items("C [G {R = G 5 | D 3}]x4")}
    assert estimate_section_lines(grouped) == estimate_section_lines(plain) + 2


# ==============================================================================
#  Line breaks — v0.20
# ==============================================================================

def test_line_break_splits_the_chart_row():
    import grammar
    from render import render_chart_row
    rows = render_chart_row(grammar.parse_items(
        "A(5) D(5) // F(8) D(5) E(7) // A(5) A(5) F(8)"))
    assert len(rows) == 3
    assert rows[0].split() == ["A(5)", "D(5)"]
    assert rows[1].split() == ["F(8)", "D(5)", "E(7)"]
    assert rows[2].split() == ["A(5)", "A(5)", "F(8)"]


def test_each_block_aligns_its_own_columns():
    """The point of breaking: a short line isn't padded out to the width
    of the longest one."""
    import grammar
    from render import render_chart_row
    rows = render_chart_row(grammar.parse_items("C G // Am Fmaj7 Bdim"))
    assert rows[0].rstrip() == rows[0].rstrip().rstrip()      # no trailing pad
    assert len(rows[0]) < len(rows[1])


def test_a_labelled_row_stacks_under_its_label():
    import grammar
    from render import render_chart_row
    rows = render_chart_row(grammar.parse_items("C G // Am F"), label="Solo")
    assert rows[0].startswith("Solo")
    body_col = rows[0].index("C")
    assert rows[1].index("Am") == body_col      # continuation lines line up


def test_line_break_plays_nothing():
    import grammar
    items = grammar.parse_items("C // G")
    assert [it.get("kind") for it in items] == ["token", "mark", "token"]
    assert items[1]["mark"] == "line_break"


def test_line_break_round_trips():
    import grammar
    line = "A(5) D(5) // F(8) E(7)"
    assert grammar.unparse(grammar.parse_items(line)) == line


def test_stray_breaks_do_not_make_empty_lines():
    import grammar
    from render import render_chart_row
    assert len(render_chart_row(grammar.parse_items("// C G //"))) == 1
    assert len(render_chart_row(grammar.parse_items("C // // G"))) == 2


def test_breaks_are_counted_in_the_page_estimate():
    import grammar
    from render import estimate_section_lines
    one = {"render": "chart", "items": grammar.parse_items("C G Am F")}
    two = {"render": "chart", "items": grammar.parse_items("C G // Am F")}
    assert estimate_section_lines(two) == estimate_section_lines(one) + 2


def test_a_break_works_alongside_licks():
    import grammar
    from render import render_chart_row
    rows = render_chart_row(grammar.parse_items("C {G 5 7} // D G"))
    assert len(rows) == 3           # symbols + one lick line, then the next line
    assert rows[1].strip() == "|G|-5-7-|"
    assert rows[2].split() == ["D", "G"]


def test_indented_break_pushes_its_line_in():
    import grammar
    from render import render_chart_row, INDENT_STEP
    rows = render_chart_row(grammar.parse_items(
        "A(5) D(5) //> F(8) D(5) E(7) // A(5) A(5) F(8)"))
    assert len(rows) == 3
    flush = rows[0].index("A(5)")
    assert rows[1].index("F(8)") == flush + INDENT_STEP
    assert rows[2].index("A(5)") == flush        # back out again


def test_indent_levels_stack():
    import grammar
    from render import render_chart_row, INDENT_STEP
    rows = render_chart_row(grammar.parse_items("C //> G //>> Am"))
    base = rows[0].index("C")
    assert rows[1].index("G") == base + INDENT_STEP
    assert rows[2].index("Am") == base + 2 * INDENT_STEP


def test_indent_is_relative_to_a_label_gutter():
    import grammar
    from render import render_chart_row, INDENT_STEP
    rows = render_chart_row(grammar.parse_items("C //> G"), label="Solo")
    assert rows[1].index("G") == rows[0].index("C") + INDENT_STEP


def test_indented_break_round_trips():
    import grammar
    for line in ("C //> G", "C //>> G", "C // G"):
        assert grammar.unparse(grammar.parse_items(line)) == line


def test_a_plain_break_carries_no_indent_field():
    import grammar
    plain = grammar.parse_items("C // G")[1]
    pushed = grammar.parse_items("C //> G")[1]
    assert "indent" not in plain
    assert pushed["indent"] == 1


def test_section_ref_keeps_the_targets_line_breaks():
    """A `//` added to the target must read on every section that
    references it — the reference is the target's content, breaks and all."""
    import model, grammar
    from render import resolve_references, render_chart_row
    doc = model.new_document(title="T")
    target = model.new_section("chorus1", "Chorus_1", "Chorus")
    target["items"] = grammar.parse_items("[C]x3 [F] // [F] [C] // A(5)")
    src = model.new_section("s2", "Chorus_2", "Chorus")
    src["items"] = [model.make_section_ref("chorus1")]
    doc["sections"] = [target, src]

    out = resolve_references(src["items"], doc)
    assert [it["kind"] for it in out].count("mark") == 2
    assert render_chart_row(out) == render_chart_row(target["items"])


def test_estimate_section_lines_measures_a_reference_by_its_target():
    """Without the document a `=Chorus_1` section is one item; with it,
    it is however many lines Chorus_1 actually prints. The export's
    keep-a-section-whole rule is only as good as this number."""
    import model, grammar
    from render import estimate_section_lines
    doc = model.new_document(title="T")
    target = model.new_section("chorus1", "Chorus_1", "Chorus")
    target["items"] = grammar.parse_items("[C]x3 [F] // [F] [C] // A(5)")
    src = model.new_section("s2", "Chorus_2", "Chorus")
    src["items"] = [model.make_section_ref("chorus1")]
    doc["sections"] = [target, src]

    assert estimate_section_lines(src) < estimate_section_lines(src, doc=doc)
    assert (estimate_section_lines(src, doc=doc)
            == estimate_section_lines(target, doc=doc))


# ==============================================================================
#  Row roles — what a front end needs to colour a row without re-parsing it
# ==============================================================================

def test_rows_are_tagged_fret_symbol_and_lick():
    import grammar
    from render import (render_chart_rows, ROLE_FRET, ROLE_SYM, ROLE_LICK)
    rows = render_chart_rows(grammar.parse_items("5A {G 5 7 | D - 3} 7D"))
    assert [r["role"] for r in rows] == [ROLE_FRET, ROLE_SYM, ROLE_LICK, ROLE_LICK]


def test_a_rest_is_tagged_where_it_sits_on_the_symbol_row():
    import grammar
    from render import render_chart_rows, ROLE_REST
    rows = render_chart_rows(grammar.parse_items("5A rest 7D"))
    sym = next(r for r in rows if r["role"] == "sym")
    assert len(sym["spans"]) == 1
    start, end, role = sym["spans"][0]
    assert role == ROLE_REST
    assert sym["text"][start:end] == "rest"


def test_a_row_with_no_rest_has_no_spans():
    import grammar
    from render import render_chart_rows
    rows = render_chart_rows(grammar.parse_items("5A 7D G"))
    assert all(not r["spans"] for r in rows)


def test_tagged_rows_and_plain_lines_agree():
    import grammar
    from render import render_chart_row, render_chart_rows
    items = grammar.parse_items("C rest {G 5 7 5} // D G")
    assert ([r["text"] for r in render_chart_rows(items)] ==
            render_chart_row(items))


def test_a_licks_columns_line_up_across_its_strings():
    import grammar
    from render import render_chart_rows
    rows = render_chart_rows(grammar.parse_items("{G 12 - 0 | D - - 3}"))
    licks = [r["text"].strip() for r in rows if r["role"] == "lick"]
    assert len({len(t) for t in licks}) == 1        # same width, so they align
    assert licks[0].startswith("|G|") and licks[1].startswith("|D|")


def test_a_standalone_tab_block_is_spaced_the_same_on_both_sides():
    """A lick with no chords beside it is set off above and below — a
    blank line on one side only makes the next chart line read as tab."""
    import grammar, render
    items = grammar.parse_items("[C] [F] // {D - - - 3 | A - 4 - - 4} // A(5)")
    lines = render.chart_body_lines(items)
    tab = [i for i, ln in enumerate(lines) if ln.strip().startswith("|D|")][0]
    assert not lines[tab - 1].strip()                   # space before
    assert not lines[tab + 2].strip()                   # and after the block
    assert lines[tab + 3].strip().startswith("A(5)")


def test_a_section_does_not_end_on_the_tab_block_s_blank_line():
    """The gap before the next section is already there."""
    import grammar, render
    items = grammar.parse_items("[C] [F] // {D - - - 3 | A - 4 - - 4}")
    lines = render.chart_body_lines(items)
    assert lines[-1].strip()


def test_a_lick_among_chords_still_sits_under_them():
    """The other case is untouched: written in among chords, a lick reads
    where it's played rather than in a block of its own."""
    import grammar, render
    items = grammar.parse_items("[C] {D - - - 3} [F]")
    lines = render.chart_body_lines(items)
    sym = [i for i, ln in enumerate(lines) if "[C]" in ln][0]
    assert "|D|" in lines[sym + 1]


# ---------------------------------------------------------------------------
#  Named licks — recalled by name, rendered as the notes
# ---------------------------------------------------------------------------

import model      # noqa: E402
import render     # noqa: E402


def _doc_with_named_lick():
    import grammar
    doc = model.new_document(title="X")
    sec = model.new_section("s1", "Intro")
    sec["items"] = grammar.parse_items("{Riff1 = G 5 7 5 | D - - 3}")
    doc["sections"].append(sec)
    return doc


def test_a_named_lick_prints_its_name_in_front_of_its_tab():
    """On the tab's own first line — "Riff1 |G|-5-7-|" — not on a row of
    its own above it; the other string lines are indented to match."""
    import grammar
    rows = render.chart_body_lines(grammar.parse_items("{Riff1 = G 5 7 | D - 3}"))
    first = next(r for r in rows if "|G|" in r)
    second = next(r for r in rows if "|D|" in r)
    assert first.strip().startswith("Riff1 |G|")
    assert first.index("|G|") == second.index("|D|")
    assert not any(r.strip() == "Riff1" for r in rows)


def test_a_lick_reference_resolves_to_the_notes():
    import grammar
    doc = _doc_with_named_lick()
    items = grammar.parse_items("A {Riff1}x3")
    out = "\n".join(render.chart_body_lines(render.resolve_references(items, doc)))
    assert "Riff1 (x3)" in out
    assert "|G|-5-7-5-|" in out


def test_an_unknown_lick_reference_stays_a_reference():
    import grammar
    items = grammar.parse_items("A {Nope}")
    out = "\n".join(render.chart_body_lines(
        render.resolve_references(items, model.new_document())))
    assert "{Nope}" in out


def test_a_lick_reference_matches_loosely_like_a_section_does():
    import grammar
    doc = _doc_with_named_lick()
    items = grammar.parse_items("{riff1}")
    out = "\n".join(render.chart_body_lines(render.resolve_references(items, doc)))
    assert "|G|-5-7-5-|" in out


# ---------------------------------------------------------------------------
#  Lyrics beside the chart
# ---------------------------------------------------------------------------

def test_compose_beside_puts_the_words_to_the_right():
    out = render.compose_beside(["  A  C", "  |G|-5-|"], ["one", "two"])
    assert out[0].startswith("  A  C")
    assert out[0].rstrip().endswith("one")
    assert out[1].rstrip().endswith("two")
    # Both lines start their words in the same column.
    assert out[0].index("one") == out[1].index("two")


def test_compose_beside_carries_an_overhanging_lyric():
    out = render.compose_beside(["  A"], ["one", "two", "three"])
    assert len(out) == 3
    assert out[2].strip() == "three"


def test_compose_beside_with_no_words_changes_nothing():
    chart = ["  A  C"]
    assert render.compose_beside(chart, []) == chart


def test_printable_lyrics_drops_section_markers():
    assert render.printable_lyrics("=== Verse 1 ===\nsing it") == ["sing it"]


def test_beside_layout_costs_the_taller_side_only():
    import grammar
    doc = model.new_document(title="X")
    sec = model.new_section("s1", "Verse 1")
    sec["items"] = grammar.parse_items("5A 5D 8F 5C")
    sec["lyrics_text"] = "one\ntwo"
    sec["print_lyrics"] = True
    doc["sections"].append(sec)

    doc["lyrics_layout"] = "beside"
    beside = render.estimate_section_lines(sec, None, doc)
    doc["lyrics_layout"] = "below"
    below = render.estimate_section_lines(sec, None, doc)
    assert beside < below


def test_recalled_lick_by_name_only_prints_no_tab():
    import grammar
    doc = _doc_with_named_lick()
    doc["lick_refs"] = "name"
    items = grammar.parse_items("A {riff1}x3")
    out = "\n".join(render.chart_body_lines(render.resolve_references(items, doc)))
    assert "Riff1 (x3)" in out          # spelled as defined, not as typed
    assert "|G|" not in out and "{" not in out


def test_an_unknown_lick_keeps_its_braces_in_either_mode():
    import grammar
    doc = _doc_with_named_lick()
    doc["lick_refs"] = "name"
    out = "\n".join(render.chart_body_lines(
        render.resolve_references(grammar.parse_items("{Nope}"), doc)))
    assert "{Nope}" in out


def test_lick_names_carry_the_lick_colour():
    """As tab, the name sits on the tab's own line, which is all lick
    blue; by name only, it's a blue span on the chord row."""
    import grammar
    doc = _doc_with_named_lick()
    doc["lick_refs"] = "tab"
    rows = render.chart_body_rows(
        render.resolve_references(grammar.parse_items("A {Riff1}"), doc))
    named = next(r for r in rows if "Riff1" in r["text"])
    assert named["role"] == render.ROLE_LICK and "|G|" in named["text"]

    doc["lick_refs"] = "name"
    rows = render.chart_body_rows(
        render.resolve_references(grammar.parse_items("A {Riff1}"), doc))
    sym = next(r for r in rows if "Riff1" in r["text"])
    start = sym["text"].index("Riff1")
    assert any(s <= start < e and role == render.ROLE_LICK
               for s, e, role in sym["spans"])

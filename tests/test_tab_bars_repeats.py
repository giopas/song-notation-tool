"""v0.36: bar lines inside tab measures, tab repeats, named tab grids,
copies that bring their parts and grid, and tab kept as text moved into
the grid."""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import export  # noqa: E402
import grammar  # noqa: E402
import model  # noqa: E402
import render  # noqa: E402
import songmap  # noqa: E402
import tabtext  # noqa: E402

BASS, GUITAR = "Bass (4-string)", "Guitar (6-string)"


def measure(beats, **strings):
    m = {"kind": "measure", "beats": beats,
         "strings": {st: strings.get(st, "") for st in ("G", "D", "A", "E")}}
    return m


def tab_song(*measures, name="", mode="tab"):
    doc = model.new_document("Tab Song", "The Testers")
    sec = model.new_section("intro", "Intro", instrument=BASS, render=mode)
    sec["items"] = list(measures)
    if name:
        sec["tab_name"] = name
    doc["sections"].append(sec)
    return doc


# ── bar lines inside a measure ──────────────────────────────────────────

def test_an_inner_bar_prints_after_its_beat():
    m = measure(8, E="3 - - - 5 - - -")
    m["bars"] = [4]
    assert render.measure_text(m, "E") == "  3  -  -  -|  5  -  -  -|"


def test_bars_out_of_range_or_repeated_are_ignored():
    m = measure(4)
    m["bars"] = [0, 2, 2, 4, 9, "x"]
    assert render.measure_bars(m) == [2]


def test_bars_widen_the_measure_and_keep_strings_in_line():
    m = measure(6, G="- - 5 - - 7", E="3 - - - - -")
    m["bars"] = [3]
    doc = tab_song(m)
    lines = export.build_song_lines(doc)
    rows = [ln for ln in lines if ln.startswith(("G|", "E|"))]
    assert len(rows) == 2
    assert rows[0].index("|", 3) == rows[1].index("|", 3)
    assert render.measure_width(m) == 6 * 3 + 2


def test_no_bars_prints_as_before():
    m = measure(8, E="3 - - - - - - -")
    assert render.measure_text(m, "E") == "".join(f"{t:>3}" for t in
                                                  "3 - - - - - - -".split()) + "|"


# ── repeats ─────────────────────────────────────────────────────────────

def test_a_repeated_measure_prints_repeat_signs_and_the_count():
    m = measure(4, E="3 - 5 -")
    m["repeat_open"] = True
    m["repeat_close"] = 3
    assert render.measure_text(m, "E") == ":  3  -  5  -:|"
    head = render.measure_header(m, 0)
    assert head.startswith("M1") and head.rstrip().endswith("x3")
    text = "\n".join(export.build_song_lines(tab_song(m)))
    assert "x3" in text and ":|" in text


def test_a_repeat_can_span_measures():
    a, b, c = measure(4, E="3 - - -"), measure(4, E="5 - - -"), measure(4, E="7 - - -")
    a["repeat_open"] = True
    b["repeat_close"] = 2
    lines = render.tab_grid_lines([a, b, c], BASS)
    e_row = [ln for ln in lines if ln.startswith("E|")][0]
    assert e_row.count(":") == 2
    assert e_row.index(":") < e_row.index("5") < e_row.rindex(":") < e_row.index("7")


def test_a_repeat_count_below_two_is_no_repeat():
    m = measure(2)
    m["repeat_close"] = 1
    assert render.measure_repeat(m) == 0
    assert render.measure_text(m, "E").endswith("-|")
    assert not render.measure_text(m, "E").endswith(":|")


def test_pdf_builds_with_bars_repeats_and_a_name():
    m = measure(8, E="3 - - - 5 - - -", A="- - 5 - - - 7 -")
    m["bars"] = [4]
    m["repeat_open"] = True
    m["repeat_close"] = 4
    doc = tab_song(m, name="RiffA")
    pdf = export.build_pdf(doc)
    assert pdf.startswith(b"%PDF")
    assert len(pdf) > len(export.build_pdf(tab_song(measure(8))))


# ── named grids ─────────────────────────────────────────────────────────

def test_a_named_grid_prints_its_name_over_it():
    doc = tab_song(measure(4, E="3 - 5 -"), name="Hook")
    lines = export.build_song_lines(doc)
    i = lines.index("Hook")
    assert lines[i + 1].lstrip().startswith("M1")


def test_a_named_grid_plays_where_it_is_recalled():
    doc = tab_song(measure(4, E="3 - 5 -"), name="Hook")
    verse = model.new_section("v1", "Verse", instrument=BASS)
    verse["items"], _ = grammar.parse("5A {Hook}x2")
    doc["sections"].append(verse)
    lick = songmap.find_lick(doc, "hook")
    assert lick and lick["name"] == "Hook"
    text = "\n".join(export.build_song_lines(doc))
    assert "Hook (x2)" in text
    assert "|E|-3---5---|" in text


def test_grid_lick_keeps_bar_lines_in_columns():
    a, b = measure(2, E="3 -"), measure(2, A="- 5")
    a["bars"] = [1]
    lick = render.measures_to_lick([a, b], ["G", "D", "A", "E"], name="X")
    lines = render._lick_lines(lick)
    assert len({ln.index("|", 3) for ln in lines}) == 1
    assert lines == ["|A|--|--|---5-|", "|E|-3|--|-----|"]


def test_grid_names_are_listed_in_chart_order():
    doc = tab_song(measure(2), name="One")
    sec = model.new_section("s2", "Solo", instrument=GUITAR, render="tab")
    sec["items"] = [{"kind": "measure", "beats": 2, "strings": {"e": "1 -"}}]
    sec["tab_name"] = "Two"
    doc["sections"].append(sec)
    assert songmap.grid_names(doc) == ["One", "Two"]


# ── copies bring parts and grid ─────────────────────────────────────────

def song_with_parts_and_tab():
    doc = model.new_document("Linked", "The Testers")
    src = model.new_section("v1", "Verse", instrument=BASS, render="both")
    src["items"], _ = grammar.parse("5A 7D")
    src["items"].append(measure(4, E="3 - 5 -"))
    src["parts"] = [model.make_part("p1", GUITAR, grammar.parse("Am C")[0]),
                    model.make_part("p2", "Keys", [], render="free",
                                    free_text="pad on C")]
    doc["sections"].append(src)
    copy = model.new_section("v2", "Verse (ref)", instrument=BASS, render="both")
    copy["items"] = [model.make_section_ref("v1")]
    copy["parts"] = [songmap.make_part_ref("q1", "v1", "p1", GUITAR),
                     songmap.make_part_ref("q2", "v1", "p2", "Keys")]
    doc["sections"].append(copy)
    return doc


def test_a_copied_section_prints_every_part_and_its_grid():
    doc = song_with_parts_and_tab()
    text = "\n".join(export.build_song_lines(doc))
    assert text.count("Am") == 2
    assert text.count("pad on C") == 2
    assert text.count("M1") == 2


def test_a_part_reference_follows_the_part_it_plays():
    doc = song_with_parts_and_tab()
    doc["sections"][0]["parts"][0]["items"], _ = grammar.parse("Dm F")
    text = "\n".join(export.build_song_lines(doc))
    assert text.count("Dm") == 2 and "Am" not in text


def test_expand_links_leaves_the_song_alone():
    doc = song_with_parts_and_tab()
    before = repr(doc)
    out = songmap.expand_links(doc)
    assert repr(doc) == before
    assert out is not doc
    plain = tab_song(measure(2))
    assert songmap.expand_links(plain) is plain


def test_a_reference_to_a_missing_part_prints_nothing_and_does_not_fail():
    doc = song_with_parts_and_tab()
    doc["sections"][1]["parts"][0]["ref"]["part"] = "gone"
    export.build_song_lines(doc)
    export.build_pdf(doc)


def test_a_linked_tab_copy_repeats_the_whole_grid():
    doc = tab_song(measure(4, E="3 - 5 -"), measure(4, E="7 - - -"))
    copy = model.new_section("i2", "Intro (ref)", instrument=BASS, render="tab")
    copy["items"] = [model.make_section_ref("intro", repeat=2)]
    doc["sections"].append(copy)
    linked = songmap.linked_measures(doc, copy)
    assert linked[0]["repeat_open"] and linked[-1]["repeat_close"] == 2
    assert "repeat_open" not in doc["sections"][0]["items"][0]


def test_a_chart_reference_to_a_tab_section_shows_its_tab():
    doc = tab_song(measure(4, E="3 - 5 -"))
    items = render.resolve_references([model.make_section_ref("intro")], doc)
    assert items and items[0]["kind"] == "lick"


def test_the_saved_song_keeps_part_references():
    doc = song_with_parts_and_tab()
    data = export.song_bytes(doc)
    assert b'"ref"' in data


# ── tab kept as text, moved into the grid ───────────────────────────────

TEXT_TAB = """Hook
G|-----5---7-|---------|
D|---7-------|-5-------|
A|-----------|---------|
E|-3---------|-----3---|
x4
"""


def test_text_tab_reads_as_measures():
    measures, other = tabtext.parse_text_tab(TEXT_TAB, BASS)
    assert len(measures) == 2
    assert measures[0]["strings"]["E"].split()[0] == "3"
    assert measures[0]["repeat_open"] and measures[-1]["repeat_close"] == 4
    assert other == ["Hook"]


def test_text_that_is_not_tab_is_left_alone():
    assert tabtext.parse_text_tab("G: play it soft\nthen D", BASS) is None
    assert not tabtext.is_tab_row("G: play it soft")
    assert tabtext.is_tab_row("e|--3--|")


def test_guitar_tab_with_two_e_strings_maps_by_position():
    text = "\n".join(f"{s}|--{n}--|" for s, n in
                     zip("EBGDAE", ["0", "1", "0", "2", "3", "-"]))
    measures, _ = tabtext.parse_text_tab(text, GUITAR)
    assert measures[0]["strings"]["e"].split()[0] == "0"
    assert measures[0]["strings"]["A"].split()[0] == "3"


def test_a_free_section_of_tab_becomes_a_grid_and_keeps_its_text():
    doc = model.new_document("Old", "The Testers")
    sec = model.new_section("s1", "Riff", instrument=BASS, render="free",
                            free_text=TEXT_TAB)
    doc["sections"].append(sec)
    changed = tabtext.upgrade_tab_layout(doc)
    assert changed == [{"section": "Riff", "part": "", "from": "text"}]
    assert sec["render"] == "tab"
    assert sec["free_text"] == TEXT_TAB
    assert sec["tab_name"] == "Hook"
    assert songmap.measure_items(sec)
    # Once only: switched back to Free on purpose, it stays Free.
    sec["render"] = "free"
    assert tabtext.upgrade_tab_layout(doc) == []
    assert sec["render"] == "free"


def test_a_section_of_only_licks_becomes_a_grid():
    doc = model.new_document("Old", "The Testers")
    sec = model.new_section("s1", "Intro", instrument=BASS)
    sec["items"], _ = grammar.parse("[{Riff1 = G 5 7 5 | D - - 3}]x4")
    doc["sections"].append(sec)
    other = model.new_section("s2", "Verse", instrument=BASS)
    other["items"], _ = grammar.parse("5A {Riff1}")
    doc["sections"].append(other)
    assert tabtext.upgrade_tab_layout(doc)[0]["from"] == "licks"
    assert sec["render"] == "tab"
    (m,) = songmap.measure_items(sec)
    assert m["name"] == "Riff1"
    assert m["strings"]["G"] == "5 7 5" and m["repeat_close"] == 4
    # The verse that recalls it still plays it.
    assert "|G|-5-7-5-|" in "\n".join(export.build_song_lines(doc))


def test_licks_among_chords_stay_licks():
    doc = model.new_document("Old", "The Testers")
    sec = model.new_section("s1", "Verse", instrument=BASS)
    sec["items"], _ = grammar.parse("C {G 5 7 5 | D - - 3} G")
    doc["sections"].append(sec)
    assert tabtext.upgrade_tab_layout(doc) == []
    assert sec["render"] == "chart"


def test_a_free_part_of_tab_is_upgraded_too():
    doc = model.new_document("Old", "The Testers")
    sec = model.new_section("s1", "Verse", instrument=BASS)
    sec["items"], _ = grammar.parse("5A")
    sec["parts"] = [model.make_part("p1", BASS, render="free",
                                    free_text="G|--5--7--|\nD|--------|")]
    doc["sections"].append(sec)
    changed = tabtext.upgrade_tab_layout(doc)
    assert changed and changed[0]["part"] == BASS
    assert sec["parts"][0]["render"] == "tab"
    assert songmap.measure_items(sec["parts"][0])


def test_a_very_long_line_of_text_tab_is_cut_into_measures():
    text = "G|" + "-5" * 80 + "|"
    measures, _ = tabtext.parse_text_tab(text, BASS)
    assert all(m["beats"] <= tabtext.MAX_BEATS for m in measures)
    assert len(measures) >= 2


def test_upgraded_songs_still_pass_the_check():
    doc = model.new_document("Old", "The Testers")
    doc["sections"].append(model.new_section("s1", "Riff", instrument=BASS,
                                             render="free", free_text=TEXT_TAB))
    tabtext.upgrade_tab_layout(doc)
    assert not export.check_document(doc)["errors"]


def test_tab_names_follow_the_lick_rules():
    doc = tab_song(measure(2, E="3 -"), name="Hook")
    owner = doc["sections"][0]
    assert songmap.check_tab_name(doc, "Hook", owner) == ""
    assert songmap.check_tab_name(doc, "Hook", None)      # taken by the grid
    assert songmap.check_tab_name(doc, "A1")
    assert songmap.check_tab_name(doc, "two words")
    owner["tab_name"] = "Am"
    assert any("cannot be recalled" in w
               for w in export.check_document(doc)["warnings"])


# ── riffs named inside a tab (Resonance, 14 November) ───────────────────

def resonance_like():
    """Rests, then two named licks on one line, recalled elsewhere: the
    shape of a real song written before the grid could grow."""
    doc = model.new_document("Echo Room", "The Testers", time="3/4")
    intro = model.new_section("i", "Intro", instrument=BASS)
    intro["items"], _ = grammar.parse(
        "rest rest {RiffA = A - - 9 10 | E 7 10 - -} {RiffB = A - 9 | E 7 -}")
    verse = model.new_section("v", "Verse", instrument=BASS)
    verse["items"], _ = grammar.parse("{RiffA}x3 {RiffB}")
    doc["sections"] = [intro, verse]
    return doc


def test_rests_and_two_named_licks_become_one_grid():
    doc = resonance_like()
    changed = tabtext.upgrade_tab_layout(doc)
    assert changed == [{"section": "Intro", "part": "", "from": "licks"}]
    intro, verse = doc["sections"]
    ms = songmap.measure_items(intro)
    assert [m["beats"] for m in ms] == [3, 3, 4, 2]     # rests: a 3/4 bar each
    assert [m.get("name") for m in ms] == [None, None, "RiffA", "RiffB"]
    assert verse["render"] == "chart"                   # recalls stay recalls
    text = "\n".join(export.build_song_lines(doc))
    assert "M3 RiffA" in text and "RiffB" in text.split("Verse")[0]
    assert "RiffA (x3) |A|-------9--10-|" in text
    assert "RiffB |A|---9-|" in text


def test_a_line_of_only_rests_is_left_alone():
    doc = model.new_document("Quiet", "The Testers")
    sec = model.new_section("o", "Outro", instrument=BASS)
    sec["items"], _ = grammar.parse("rest rest")
    doc["sections"].append(sec)
    assert tabtext.upgrade_tab_layout(doc) == []


def test_a_song_opened_by_an_older_upgrade_gets_the_new_pass():
    doc = resonance_like()
    doc["tab_layout"] = 2
    assert tabtext.upgrade_tab_layout(doc)


def test_a_riff_name_spans_its_measures():
    a, b, c = measure(2, E="3 -"), measure(2, E="5 -"), measure(2, E="7 -")
    a["name"], a["span"] = "Hook", 2
    doc = tab_song(a, b, c)
    lick = songmap.find_lick(doc, "Hook")
    assert lick["lines"][0]["frets"] == ["3", "-", "|", "5", "-"]
    assert songmap.grid_names(doc) == ["Hook"]


def test_riff_names_are_checked_against_each_other():
    a, b = measure(2, E="3 -"), measure(2, E="5 -")
    a["name"], b["name"] = "Hook", "Hook"
    doc = tab_song(a, b)
    assert songmap.check_tab_name(doc, "Hook", a)
    assert any("riff name 'Hook'" in w for w in export.check_document(doc)["warnings"])


def test_a_linked_copy_does_not_print_the_riff_name_twice():
    a = measure(2, E="3 -")
    a["name"] = "Hook"
    doc = tab_song(a)
    copy = model.new_section("i2", "Again", instrument=BASS, render="tab")
    copy["items"] = [model.make_section_ref("intro")]
    doc["sections"].append(copy)
    text = "\n".join(export.build_song_lines(doc))
    assert text.count("Hook") == 1
    assert "name" not in songmap.linked_measures(doc, copy)[0]


def test_the_header_stays_as_wide_as_its_measure():
    m = measure(2, E="3 -")
    m["name"] = "AVeryLongRiffName"
    m["repeat_close"] = 2
    assert len(render.measure_header(m, 0)) == render.measure_width(m)
    assert render.measure_header(m, 0).endswith("x2")

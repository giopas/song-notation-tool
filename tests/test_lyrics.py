import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import lyrics
import model


SHEET = """We passed upon the stairs
We spoke of was and when

Oh no, not me
We never lost control

I laughed and shook his hand

Oh no, not me
We never lost control
"""


def _song(types):
    doc = model.new_document(title="The Man Who Sold The World")
    doc["sections"] = [model.new_section(f"s{i}", f"{t}_{i}", t)
                       for i, t in enumerate(types)]
    return doc


def test_a_sheet_splits_on_blank_lines():
    assert lyrics.split_blocks(SHEET)[0].startswith("We passed upon the stairs")
    assert len(lyrics.split_blocks(SHEET)) == 3      # the chorus folded into one


def test_a_chorus_written_twice_is_one_block_not_two():
    """Three identical entries to choose between is no choice at all."""
    blocks = lyrics.split_blocks(SHEET)
    assert len(blocks) == len(set(blocks))
    assert lyrics.repeated_blocks(SHEET) == ["Oh no, not me\nWe never lost control"]


def test_the_repeated_block_goes_to_every_chorus():
    """A lyric sheet marks its chorus by repeating it — so a chorus played
    three times gets the words all three times, not just the first."""
    doc = _song(["Intro", "Verse", "Chorus", "Bridge", "Verse", "Chorus",
                 "Solo", "Chorus"])
    blocks = lyrics.split_blocks(SHEET)
    got = lyrics.suggest(blocks, doc["sections"], lyrics.repeated_blocks(SHEET))
    chorus_block = blocks.index("Oh no, not me\nWe never lost control")
    assert [got[2], got[5], got[7]] == [chorus_block] * 3


def test_an_instrumental_section_is_stepped_over():
    """The old split put verse one on the intro and knocked every block
    after it one section out of place."""
    doc = _song(["Intro", "Verse", "Chorus", "Bridge", "Verse", "Chorus",
                 "Solo", "Chorus"])
    blocks = lyrics.split_blocks(SHEET)
    got = lyrics.suggest(blocks, doc["sections"], lyrics.repeated_blocks(SHEET))
    assert got[0] is None and got[6] is None             # Intro, Solo
    assert blocks[got[1]].startswith("We passed")        # Verse_1
    assert blocks[got[4]].startswith("I laughed")        # Verse_2
    assert got[3] is None                                # Bridge: no words left


def test_a_song_of_custom_sections_still_gets_a_suggestion():
    doc = _song(["Custom", "Custom"])
    got = lyrics.suggest(["one", "two"], doc["sections"])
    assert got == [0, 1]


def test_without_repeats_the_sheet_is_read_in_playing_order():
    doc = _song(["Verse", "Chorus", "Verse", "Chorus"])
    got = lyrics.suggest(["v1", "ch", "v2"], doc["sections"], [])
    assert got == [0, 1, 2, 1]        # the last chorus repeats its own type


def test_applying_keeps_the_sheet_but_stops_it_printing_twice():
    """The sheet is the source the blocks came from — and the reason a
    section added next week can still be given one."""
    doc = _song(["Verse", "Chorus"])
    doc["lyrics_text"] = SHEET
    doc["print_lyrics"] = True
    blocks = lyrics.split_blocks(SHEET)
    n = lyrics.apply_assignment(doc, blocks, [0, 1], print_lyrics=True)
    assert n == 2
    assert doc["lyrics_text"] == SHEET
    assert doc["print_lyrics"] is False
    assert doc["sections"][0]["lyrics_text"] == blocks[0]
    assert doc["sections"][0]["print_lyrics"] is True


def test_keep_leaves_a_section_alone_and_none_clears_it():
    doc = _song(["Verse", "Chorus"])
    doc["sections"][0]["lyrics_text"] = "hand-typed"
    doc["sections"][1]["lyrics_text"] = "to be cleared"
    lyrics.apply_assignment(doc, ["a block"], ["keep", None])
    assert doc["sections"][0]["lyrics_text"] == "hand-typed"
    assert doc["sections"][1]["lyrics_text"] == ""


def test_reopening_shows_what_was_assigned_rather_than_a_fresh_guess():
    doc = _song(["Verse", "Chorus"])
    blocks = lyrics.split_blocks(SHEET)
    lyrics.apply_assignment(doc, blocks, [2, 1])
    assert lyrics.current_assignment(blocks, doc["sections"]) == [2, 1]


def test_the_split_endpoint_answers_with_blocks_and_a_proposal():
    import webserver
    doc = _song(["Intro", "Verse", "Chorus"])
    doc["lyrics_text"] = SHEET
    blocks = lyrics.split_blocks(SHEET)
    # The route's own body handling is HTTP plumbing; this is the payload
    # it builds, which is what the front ends actually read.
    payload = {
        "blocks": blocks,
        "suggested": lyrics.suggest(blocks, doc["sections"],
                                     lyrics.repeated_blocks(SHEET)),
        "current": lyrics.current_assignment(blocks, doc["sections"]),
    }
    assert payload["suggested"][0] is None
    assert payload["current"] == [None, None, None]
    assert hasattr(webserver, "lyrics_mod")

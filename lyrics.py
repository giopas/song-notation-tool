"""
lyrics.py — splitting a pasted lyric sheet across a song's sections.

A lyric sheet arrives as one block of text: verses, choruses and the rest,
separated by blank lines, in the order they're sung. A chart is a list of
sections in the order they're played. The two are the same song told twice,
and lining them up by hand — copy a verse, pick the section, paste, repeat —
is the tedious part this module exists to remove.

What it does *not* do is guess silently. `suggest()` proposes a mapping and
the front end shows it as one row per section for the user to correct, so a
song whose sheet doesn't run in playing order (an instrumental intro, a
chorus written out once and played three times) is fixed in the dialog
rather than discovered later on paper.

    split_blocks(text)                  -> ["verse one…", "chorus…", …]
    repeated_blocks(text)               -> the ones written out twice
    suggest(blocks, sections, repeats)  -> [block index or None, per section]
    current_assignment(blocks, secs)    -> what each section already holds
    apply_assignment(doc, assignment)   -> writes it into the document

The document's own `lyrics_text` is left alone by `apply_assignment`: it is
the sheet, and it stays the sheet. Add a section next week and the same
blocks are still there to assign to it — which is the whole point of
keeping it.
"""

from __future__ import annotations

import re

# Sections that carry words. The rest of the type list — Intro, Interlude,
# Solo, Breakdown, Outro — is where nobody is singing, so the suggestion
# steps over them rather than pushing the first verse onto the intro and
# knocking every later block one section out of place.
SINGABLE_TYPES = ("Verse", "Pre-Chorus", "Chorus", "Refrain", "Bridge")

_BLANK_LINE = re.compile(r"\n\s*\n+")


def split_blocks(text: str) -> list:
    """The lyric sheet's blocks: the stretches between blank lines, with
    repeats folded into one.

    A blank line is how every lyric sheet in the world separates one verse
    from the next, which makes it the one piece of structure a pasted sheet
    can be relied on to carry. A chorus printed out three times is one
    block, not three: the same words assigned to three sections is exactly
    what the split is for, and three identical entries to choose between
    is no choice at all.
    """
    blocks = [b.strip() for b in _BLANK_LINE.split(text or "") if b.strip()]
    return list(dict.fromkeys(blocks))


def repeated_blocks(text: str) -> list:
    """The blocks the sheet writes out more than once, in first-appearance
    order — the refrain, in other words. A lyric sheet marks its chorus by
    repeating it, which is the only structural hint it gives beyond the
    blank lines."""
    seen, repeats = {}, []
    for b in (b.strip() for b in _BLANK_LINE.split(text or "") if b.strip()):
        seen[b] = seen.get(b, 0) + 1
        if seen[b] == 2:
            repeats.append(b)
    return repeats


def singable_sections(sections) -> list:
    """The indices of the sections that get words.

    A song written entirely in Custom sections has no singable types at
    all; rather than suggest nothing, every section counts in that case.
    """
    idx = [i for i, sec in enumerate(sections or [])
           if sec.get("type") in SINGABLE_TYPES]
    return idx if idx else list(range(len(sections or [])))


def suggest(blocks, sections, refrains=()) -> list:
    """A block index (or None) for each section, in section order.

    Two shapes, because lyric sheets come in two shapes.

    When the sheet repeats a block, that block is the chorus — that
    repetition is the sheet saying so — and it goes to every Chorus and
    Refrain section at once. What's left goes to the verses in order, and
    then to whatever other singable sections remain, which keeps a bridge
    from swallowing verse two and pushing everything after it one section
    out of place.

    When nothing repeats, the sheet is read as written in playing order:
    blocks go to the singable sections in turn, and a section that finds
    the blocks exhausted repeats the last one given to a section of its
    own type — a chorus written once and played three times being the
    ordinary shape of a song.
    """
    out = [None] * len(sections or [])
    if not blocks:
        return out
    idx = singable_sections(sections)
    chorus_secs = [i for i in idx
                   if (sections[i] or {}).get("type") in ("Chorus", "Refrain")]
    refrain_idx = next((blocks.index(r) for r in refrains if r in blocks), None)

    if refrain_idx is not None and chorus_secs:
        for i in chorus_secs:
            out[i] = refrain_idx
        rest = [j for j in range(len(blocks)) if j != refrain_idx]
        others = [i for i in idx if i not in chorus_secs]
        # Verses first, in their own order, then the rest of the singable
        # sections — a sheet's leftover blocks are nearly always verses.
        verses = [i for i in others if (sections[i] or {}).get("type") == "Verse"]
        ordered = verses + [i for i in others if i not in verses]
        for i, j in zip(ordered, rest):
            out[i] = j
        return out

    last_of_type = {}
    nxt = 0
    for i in idx:
        stype = (sections[i] or {}).get("type")
        if nxt < len(blocks):
            out[i] = nxt
            last_of_type[stype] = nxt
            nxt += 1
        elif stype in last_of_type:
            out[i] = last_of_type[stype]
    return out


def current_assignment(blocks, sections) -> list:
    """Which block each section is already holding, if any.

    Re-opening the dialog should show what was assigned last time rather
    than a fresh guess — otherwise every visit silently proposes undoing
    the corrections made on the previous one.
    """
    by_text = {}
    for i, b in enumerate(blocks or []):
        by_text.setdefault(b.strip(), i)
    return [by_text.get((sec.get("lyrics_text") or "").strip())
            if (sec.get("lyrics_text") or "").strip() else None
            for sec in (sections or [])]


def apply_assignment(doc: dict, blocks, assignment, print_lyrics=None) -> int:
    """Write `assignment` into `doc`'s sections; returns how many got words.

    `assignment` is one entry per section: a block index, None to clear
    that section's lyrics, or the string "keep" to leave it untouched —
    which is how a section holding words that came from somewhere else
    survives a split it wasn't part of.

    The document's own lyric sheet is kept (it is the source the blocks
    came from, and the reason a section added later can still be given
    one), but it stops printing: the sheet and the split-up sections are
    the same words, and printing both puts them on the chart twice.
    """
    sections = doc.get("sections", [])
    assigned = 0
    for sec, choice in zip(sections, assignment or []):
        if choice == "keep":
            continue
        if choice is None:
            sec["lyrics_text"] = ""
            sec["print_lyrics"] = False
            continue
        if not (0 <= int(choice) < len(blocks)):
            continue
        sec["lyrics_text"] = blocks[int(choice)]
        if print_lyrics is not None:
            sec["print_lyrics"] = bool(print_lyrics)
        assigned += 1
    if assigned:
        doc["print_lyrics"] = False
    return assigned

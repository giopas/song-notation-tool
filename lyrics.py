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


# ==============================================================================
#  Section markers — structure written into the sheet itself
#
#  Splitting a sheet on blank lines and then choosing "block 3" from a
#  dropdown asks you to hold the song's shape in your head while you read
#  a list of first lines. Writing the shape into the sheet instead —
#
#      === Intro ===
#      === Verse 1 ===
#      I'm the one…
#
#  — says the same thing where you can see it, survives re-editing the
#  words, and leaves the sheet self-describing for next time. The markers
#  are structure, not lyrics: they show in the editor and never print.
# ==============================================================================

MARKER_RE = re.compile(r'^[ \t]*={2,}[ \t]*(.+?)[ \t]*={2,}[ \t]*$')


def marker_line(name: str) -> str:
    """The marker for `name`, the way the editor writes it."""
    return f"=== {(name or '').strip()} ==="


def is_marker(line: str) -> bool:
    return bool(MARKER_RE.match(line or ""))


def marker_name(line: str) -> str:
    """The name inside a marker line, or "" if it isn't one."""
    m = MARKER_RE.match(line or "")
    return m.group(1).strip() if m else ""


def has_markers(text: str) -> bool:
    return any(is_marker(ln) for ln in (text or "").splitlines())


def strip_markers(text: str) -> str:
    """`text` with its marker lines removed — what actually prints.

    A marker sits *above* the words it labels, so removing it usually
    leaves a blank line where the separation used to be; those are
    collapsed at the edges so a section's lyrics don't start with an empty
    line on paper.
    """
    out, prev_blank = [], False
    for ln in (text or "").splitlines():
        if is_marker(ln):
            continue
        blank = not ln.strip()
        # Taking a marker out leaves the blank line that sat above it as
        # well as the one below; one blank line between verses is the
        # separation, two is a hole in the page.
        if blank and prev_blank:
            continue
        out.append(ln)
        prev_blank = blank
    while out and not out[0].strip():
        out.pop(0)
    while out and not out[-1].strip():
        out.pop()
    return "\n".join(out)


def split_marked(text: str) -> list:
    """The sheet as its marked segments: [{"name", "text"}, …].

    Anything before the first marker is a segment with an empty name — a
    title line, a note to self, the half of the sheet you haven't got to
    yet. It is kept rather than dropped so nothing is lost by marking up
    only part of a sheet.
    """
    segments, name, body = [], "", []

    def flush():
        joined = "\n".join(body).strip("\n")
        if name or joined.strip():
            segments.append({"name": name, "text": joined})

    for ln in (text or "").splitlines():
        if is_marker(ln):
            flush()
            name, body = marker_name(ln), []
        else:
            body.append(ln)
    flush()
    return segments


def _key(name: str) -> str:
    """Matching key for a marker name against a section: case-folded, and
    spaces, dashes and underscores all the same thing — so a marker typed
    "=== Chorus_1 ===" finds the section called "Chorus 1"."""
    return re.sub(r'[\s_\-]+', "", (name or "").strip().lower())


def match_sections(doc: dict, segments) -> list:
    """The section id each segment belongs to, or None where no section
    of that name exists — one entry per segment, in order.

    A segment with no name (the preamble) never matches. A name that
    matches nothing is what the front end offers to create a section for.
    """
    sections = doc.get("sections", []) or []
    by_key = {}
    for sec in sections:
        by_key.setdefault(_key(sec.get("name", "")), sec.get("id"))
        by_key.setdefault(_key(sec.get("id", "")), sec.get("id"))
    return [by_key.get(_key(seg.get("name", ""))) if seg.get("name") else None
            for seg in (segments or [])]


def unmatched_names(doc: dict, text: str) -> list:
    """The marker names in `text` with no section to go to, in order and
    without duplicates — the list the dialog turns into "create these?"."""
    segs = split_marked(text)
    ids = match_sections(doc, segs)
    out = []
    for seg, sid in zip(segs, ids):
        name = (seg.get("name") or "").strip()
        if name and sid is None and name not in out:
            out.append(name)
    return out


def apply_marked(doc: dict, text: str, print_lyrics=None) -> dict:
    """Write a marked-up sheet into the document's sections.

    Every section named by a marker gets the words under it; a section the
    sheet never names is left exactly as it is, because the sheet saying
    nothing about a section is not the same as the sheet saying it has no
    words. Segments whose names match nothing are reported back rather
    than dropped silently.

    Returns {"assigned": n, "unmatched": [name, …]}.
    """
    segments = split_marked(text)
    ids = match_sections(doc, segments)
    by_id = {sec.get("id"): sec for sec in doc.get("sections", []) or []}

    assigned, unmatched, seen = 0, [], {}
    for seg, sid in zip(segments, ids):
        name = (seg.get("name") or "").strip()
        if not name:
            continue
        if sid is None:
            if name not in unmatched:
                unmatched.append(name)
            continue
        sec = by_id.get(sid)
        if sec is None:
            continue
        body = (seg.get("text") or "").strip("\n")
        # The same section marked twice (a chorus written out where it is
        # sung) keeps both halves rather than the last one winning.
        if sid in seen and body.strip():
            sec["lyrics_text"] = (sec["lyrics_text"].rstrip("\n") + "\n\n" + body)
        else:
            sec["lyrics_text"] = body
            seen[sid] = True
        if print_lyrics is not None:
            sec["print_lyrics"] = bool(print_lyrics)
        if body.strip():
            assigned += 1

    if assigned:
        # Same rule as the block split: the sheet stays as the source, and
        # stops printing so the words don't land on the chart twice.
        doc["print_lyrics"] = False
    return {"assigned": assigned, "unmatched": unmatched}

"""
chords.py — chord shapes, and the sheet of them a chart can carry.

A chart says which chords are played. It doesn't say how to hold one, and
for most of a set it doesn't need to — but there is always the one voicing
you had to work out, or the open shape the song wants rather than the barre
your hands default to. Writing that on the back of the sheet is what this
is: a small block of diagrams, printed once at the top or the bottom of the
song rather than repeated beside every chord symbol.

The diagrams are drawn as tab, not as the dotted grid a chord book uses,
because tab is what the rest of this program prints and a fret number is
unambiguous where a dot on a grid needs a fret marker beside it:

      C
    e|-0-|
    B|-1-|
    G|-0-|
    D|-2-|
    A|-3-|
    E|-x-|

    shape_from_text("x32010")            -> ["0", "1", "0", "2", "3", "x"]
    diagram_lines(chord, strings)        -> the block above, as lines
    sheet_lines(doc, width)              -> every chord, laid out in rows
"""

from __future__ import annotations

from constants import INSTRUMENT_STRINGS

DEFAULT_INSTRUMENT = "Guitar (6-string)"
COLUMN_GAP = 4          # blank columns between two diagrams on a row
MAX_FRET = 24


class ChordError(ValueError):
    pass


def strings_for(instrument: str) -> list:
    """The instrument's strings, highest first — the order a shape's frets
    are stored in, and the order they print down the page."""
    return INSTRUMENT_STRINGS.get(instrument or DEFAULT_INSTRUMENT,
                                  INSTRUMENT_STRINGS[DEFAULT_INSTRUMENT])


def shape_from_text(text: str, instrument: str = DEFAULT_INSTRUMENT) -> list:
    """Parse a chord shape written the way guitarists write one.

    "x32010" is C: one character per string, lowest string first, which is
    the convention every chord chart and every forum post uses. Frets above
    nine need separating, so spaces (or dashes) are accepted too —
    "x 10 12 12 10 x" — and read in the same low-to-high order.

    Returned high-string-first, matching INSTRUMENT_STRINGS, because that
    is the order the diagram prints in and the order tab reads.
    """
    strings = strings_for(instrument)
    raw = (text or "").strip()
    if not raw:
        raise ChordError("no shape given")

    if any(ch in raw for ch in " -,"):
        parts = [p for p in raw.replace(",", " ").replace("-", " ").split() if p]
    else:
        parts = list(raw)

    if len(parts) != len(strings):
        raise ChordError(
            f"{len(parts)} fret(s) for a {len(strings)}-string instrument — "
            f"write one per string, lowest first (e.g. x32010)")

    out = []
    for p in parts:
        if p.lower() in ("x", "-"):
            out.append("x")
            continue
        if not p.isdigit():
            raise ChordError(f"{p!r} is not a fret (0-{MAX_FRET}, or x)")
        if int(p) > MAX_FRET:
            raise ChordError(f"fret {p} is past the end of the neck")
        out.append(str(int(p)))
    out.reverse()   # low-to-high in, high-to-low out
    return out


def shape_to_text(frets, instrument: str = DEFAULT_INSTRUMENT) -> str:
    """The inverse of shape_from_text — what the editor puts back in the
    box. Spaced out when any fret needs two digits, so it reads back."""
    low_first = list(frets or [])[::-1]
    if any(len(f) > 1 and f != "x" for f in low_first):
        return " ".join(low_first)
    return "".join(low_first)


def diagram_lines(chord: dict) -> list:
    """One chord shape as printable lines: its name, then a tab row per
    string. Every row is the same width, so a row of diagrams lines up."""
    strings = strings_for(chord.get("instrument"))
    frets = list(chord.get("frets") or [])
    frets += ["x"] * (len(strings) - len(frets))
    w = max((len(f) for f in frets), default=1)
    name_w = max((len(st) for st in strings), default=1)
    rows = [f"{st:<{name_w}}|-{f.rjust(w)}-|" for st, f in zip(strings, frets)]
    width = max((len(r) for r in rows), default=0)
    head = (chord.get("name") or "").strip()
    note = (chord.get("note") or "").strip()
    out = [f"{head:<{width}}"] + [f"{r:<{width}}" for r in rows]
    if note:
        out.append(f"{note[:width]:<{width}}")
    return out


def _block_width(chord: dict) -> int:
    return max((len(ln) for ln in diagram_lines(chord)), default=0)


def sheet_lines(doc: dict, width: int = 100, indent: int = 2) -> list:
    """Every chord shape in the document, laid out in rows that fit
    `width` columns. Returns [] when the song has no shapes.

    Shapes are grouped by instrument: a bass fingering and a guitar one
    have different numbers of strings, and a row mixing them reads as one
    wrong diagram rather than two right ones.
    """
    chords = [c for c in (doc.get("chords") or []) if (c.get("name") or "").strip()]
    if not chords:
        return []

    by_instrument = {}
    for c in chords:
        by_instrument.setdefault(c.get("instrument") or DEFAULT_INSTRUMENT,
                                  []).append(c)

    out = []
    multi = len(by_instrument) > 1
    for instrument, group in by_instrument.items():
        if multi:
            out += [" " * indent + instrument, ""]
        col_w = max(_block_width(c) for c in group) + COLUMN_GAP
        per_row = max(1, (width - indent) // col_w)
        for start in range(0, len(group), per_row):
            row = group[start:start + per_row]
            blocks = [diagram_lines(c) for c in row]
            height = max(len(b) for b in blocks)
            for i in range(height):
                line = " " * indent
                for b in blocks:
                    cell = b[i] if i < len(b) else ""
                    line += f"{cell:<{col_w}}"
                out.append(line.rstrip())
            out.append("")
    while out and not out[-1].strip():
        out.pop()
    return out


def sheet_position(doc: dict) -> str:
    """Where this song's chord sheet prints: "start", "end", or "none".
    A document with no shapes never prints one, whatever it says."""
    pos = (doc.get("chord_sheet") or "none").strip().lower()
    if pos not in ("start", "end"):
        return "none"
    if not [c for c in (doc.get("chords") or []) if (c.get("name") or "").strip()]:
        return "none"
    return pos


# ==============================================================================
#  Coverage — the shapes against the chords the chart actually plays
#
#  A sheet of shapes drifts out of step with the song the moment the song
#  changes: a chord added to the bridge with no shape for it, a shape kept
#  for a chord the arrangement dropped. Neither is an error — plenty of
#  chords need no diagram — so this reports rather than refuses.
# ==============================================================================

def _is_guitar(instrument: str) -> bool:
    return (instrument or "").lower().startswith("guitar")


def _walk(items):
    for it in items or []:
        yield it
        if it.get("kind") == "group":
            yield from _walk(it.get("items", []))


def used_symbols(doc: dict, guitar_only: bool = None) -> list:
    """Every chord symbol the printed chart shows, in order of first use.

    *Printed* is the operative word: references are expanded and the
    song's and each section's transpose applied, because a shape is for
    the chord you read on the page, not the one you typed before moving
    the song up a tone.

    `guitar_only` limits it to guitar sections — on a bass chart "A" is a
    root note, not a chord to finger. None (the default) means: guitar
    sections if the song has any, otherwise all of them.
    """
    import render
    import songmap
    import transpose

    sections = doc.get("sections", []) or []
    if guitar_only is None:
        guitar_only = any(_is_guitar(s.get("instrument")) for s in sections)

    out = []
    for sec in sections:
        if guitar_only and not _is_guitar(sec.get("instrument")):
            continue
        if sec.get("render", "chart") not in ("chart", "both"):
            continue
        eff = transpose.effective_transpose(doc.get("transpose", 0),
                                            sec.get("transpose", 0))
        items = render.resolve_display_items(
            render.resolve_references(songmap.chart_items(sec), doc), eff)
        for it in _walk(items):
            if it.get("kind") == "token":
                sym = (it.get("symbol") or "").strip()
                if sym and sym not in out:
                    out.append(sym)
    return out


def coverage(doc: dict) -> dict:
    """{"missing": [...], "unused": [...]}

    `missing` — chords the chart plays that have no shape.
    `unused`  — shapes for chords the chart never plays.

    Names compare exactly, case included: "Am" and "AM" are different
    chords, and guessing that one was meant for the other is how a sheet
    ends up printing the wrong diagram.
    """
    names = [(c.get("name") or "").strip() for c in (doc.get("chords") or [])]
    names = [n for n in names if n]
    used = used_symbols(doc)
    # Anything played anywhere counts as used, so a shape isn't flagged
    # just because the song's only guitar part is written as a bass line.
    played_anywhere = set(used) | set(used_symbols(doc, guitar_only=False))
    return {
        "missing": [u for u in used if u not in names],
        "unused": [n for n in names if n not in played_anywhere],
    }

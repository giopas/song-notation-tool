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
    chords = printed_chords(doc)
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
#  Transposing a shape
#
#  Chord symbols transpose by arithmetic; chord *shapes* don't. Move a song
#  up a tone and the C you were holding becomes a D — but the hand doesn't
#  slide x32010 up two frets (that's x54232, which nobody plays); it goes
#  to the open D, xx0232. So, in order:
#
#    1. A shape with no open strings is already movable (a barre, a power
#       chord, a jazz voicing): slide it. Same hand, different fret —
#       which is what the player chose it for.
#    2. An open shape goes to the open shape of the new chord, if there
#       is one in standard tuning.
#    3. If there isn't — there's no open F, no open B minor — it goes to
#       the E-form or A-form barre of that chord, whichever sits lower on
#       the neck.
#    4. A chord with no barre template either (an add9, a slash chord) is
#       slid as a whole, open strings included — the capo answer. Always
#       playable, if not always pretty.
#
#  Only step 1 and step 4 apply off standard six-string tuning: the open
#  and barre libraries are shapes for E A D G B E and nothing else.
#
#  Like every other transpose in this program it happens at render time:
#  the shape you typed is what's stored, and what prints follows the song.
# ==============================================================================

STANDARD_GUITAR = "Guitar (6-string)"

# Standard open shapes, written low string first as chord charts write them.
# Keyed by (root, quality) with sharps as the canonical spelling, the same
# spelling transpose.py produces.
_OPEN_TEXT = {
    ("C", ""): "x32010",  ("C", "7"): "x32310",  ("C", "maj7"): "x32000",
    ("C", "add9"): "x32030",
    ("D", ""): "xx0232",  ("D", "m"): "xx0231",  ("D", "7"): "xx0212",
    ("D", "m7"): "xx0211", ("D", "maj7"): "xx0222", ("D", "sus2"): "xx0230",
    ("D", "sus4"): "xx0233", ("D", "5"): "xx023x",
    ("E", ""): "022100",  ("E", "m"): "022000",  ("E", "7"): "020100",
    ("E", "m7"): "020000", ("E", "maj7"): "021100", ("E", "sus4"): "022200",
    ("E", "5"): "022xxx",
    ("F", "maj7"): "xx3210",
    ("G", ""): "320003",  ("G", "7"): "320001",  ("G", "maj7"): "320002",
    ("A", ""): "x02220",  ("A", "m"): "x02210",  ("A", "7"): "x02020",
    ("A", "m7"): "x02010", ("A", "maj7"): "x02120", ("A", "sus2"): "x02200",
    ("A", "sus4"): "x02230", ("A", "5"): "x022xx",
    ("B", "7"): "x21202",
}

# Movable barre templates at the nut (fret 0 = root on the open string),
# low string first. "E" roots on the low E string, "A" on the A string.
_BARRE_TEXT = {
    "":     {"E": "022100", "A": "x02220"},
    "m":    {"E": "022000", "A": "x02210"},
    "7":    {"E": "020100", "A": "x02020"},
    "m7":   {"E": "020000", "A": "x02010"},
    "maj7": {"E": "0x110x", "A": "x02120"},
    "sus4": {"E": "022200", "A": "x02230"},
    "sus2": {"A": "x02200"},
    "5":    {"E": "022xxx", "A": "x022xx"},
    "6":    {"E": "022120", "A": "x02222"},
    "m6":   {"E": "022020", "A": "x02212"},
}

# Spellings of the same quality, folded to the one the tables use.
_QUALITY_ALIASES = {
    "maj": "", "M": "", "min": "m", "-": "m", "mi": "m",
    "M7": "maj7", "Maj7": "maj7", "ma7": "maj7", "Δ": "maj7", "Δ7": "maj7",
    "min7": "m7", "-7": "m7", "mi7": "m7", "sus": "sus4", "dom7": "7",
    "min6": "m6", "-6": "m6",
}

_E_INDEX = 7   # CHROMATIC index of E  (transpose.CHROMATIC starts at A)
_A_INDEX = 0


def split_chord_name(name: str):
    """(root_index, quality) for a chord name, or (None, name) if it
    doesn't start with a note. The quality is folded to the spelling the
    shape tables use, so "Amin7", "A-7" and "Am7" all find the same shape."""
    import transpose
    root, suffix = transpose._parse_root_suffix((name or "").strip())
    if root is None:
        return None, name
    idx = transpose.note_to_index(root)
    return idx, _QUALITY_ALIASES.get(suffix, suffix)


def _open_shape(root_idx: int, quality: str):
    import transpose
    text = _OPEN_TEXT.get((transpose.CHROMATIC[root_idx], quality))
    return shape_from_text(text, STANDARD_GUITAR) if text else None


def _barre_shape(root_idx: int, quality: str):
    """The lower-sitting of the E-form and A-form barres for this chord,
    or None if the quality has no template. Lower wins because it's the
    easier reach from wherever the song's other chords are; a tie goes to
    the E form, which is the one most hands know first."""
    forms = _BARRE_TEXT.get(quality)
    if not forms:
        return None
    best = None
    for form, text in forms.items():
        fret = (root_idx - (_E_INDEX if form == "E" else _A_INDEX)) % 12
        frets = [f if f == "x" else str(int(f) + fret)
                 for f in shape_from_text(text, STANDARD_GUITAR)]
        cand = (fret, 0 if form == "E" else 1, frets)
        if best is None or cand[:2] < best[:2]:
            best = cand
    return best[2]


def _shift(frets, semitones: int):
    """Slide a whole shape along the neck — up or down, whichever octave
    keeps every sounded string on the fretboard and the hand lowest."""
    best = None
    for n in (semitones, semitones - 12, semitones + 12):
        out = []
        for f in frets:
            if f == "x":
                out.append("x")
                continue
            v = int(f) + n
            if not 0 <= v <= MAX_FRET:
                break
            out.append(str(v))
        else:
            top = max((int(f) for f in out if f != "x"), default=0)
            if best is None or top < best[0]:
                best = (top, out)
    return best[1] if best else list(frets)


def is_open_shape(frets) -> bool:
    """True if any sounded string is open — the shape is tied to the nut
    and can't simply slide."""
    return any(f == "0" for f in (frets or []))


def transpose_chord(chord: dict, semitones: int) -> dict:
    """`chord` as it prints with the song moved `semitones`: name
    transposed, shape re-voiced by the rules above. Adds `"voicing"` —
    "as typed", "moved", "open", "barre" or "capo" — so a front end can
    say what happened. The stored chord is never modified."""
    import transpose
    out = dict(chord)
    out["frets"] = list(chord.get("frets") or [])
    if not semitones:
        out["voicing"] = "as typed"
        return out

    # The name moves with the song even before a shape has been typed —
    # a row with no frets yet still has to say which chord it's for.
    out["name"] = transpose.transpose_symbol(chord.get("name", ""), semitones)
    if not out["frets"]:
        out["voicing"] = "as typed"
        return out
    frets = out["frets"]
    standard = (chord.get("instrument") or DEFAULT_INSTRUMENT) == STANDARD_GUITAR

    if not is_open_shape(frets):
        out["frets"], out["voicing"] = _shift(frets, semitones), "moved"
        return out

    root, quality = split_chord_name(out["name"])
    if standard and root is not None:
        shape = _open_shape(root, quality)
        if shape:
            out["frets"], out["voicing"] = shape, "open"
            return out
        shape = _barre_shape(root, quality)
        if shape:
            out["frets"], out["voicing"] = shape, "barre"
            return out

    out["frets"], out["voicing"] = _shift(frets, semitones), "capo"
    return out


def printed_chords(doc: dict) -> list:
    """The document's chord shapes as they print: moved with the song's
    own transpose. (A section's transpose doesn't apply — the sheet
    belongs to the whole song, not to any one section of it.)"""
    n = int(doc.get("transpose", 0) or 0)
    return [transpose_chord(c, n) for c in (doc.get("chords") or [])
            if (c.get("name") or "").strip()]


def stored_name(printed: str, doc: dict) -> str:
    """The name to store for a chord that should *print* as `printed` —
    the reverse of the song's transpose, so a row added for the D on a
    chart moved up a tone is stored as C and prints as D."""
    import transpose
    n = int(doc.get("transpose", 0) or 0)
    return transpose.transpose_symbol(printed, -n) if n else printed


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
    # Compared as printed: with the song moved up a tone, a shape stored
    # as C prints as D, and D is what it has to match on the chart.
    names = [(c.get("name") or "").strip() for c in printed_chords(doc)]
    names = [n for n in names if n]
    used = used_symbols(doc)
    # Anything played anywhere counts as used, so a shape isn't flagged
    # just because the song's only guitar part is written as a bass line.
    played_anywhere = set(used) | set(used_symbols(doc, guitar_only=False))
    return {
        "missing": [u for u in used if u not in names],
        "unused": [n for n in names if n not in played_anywhere],
    }

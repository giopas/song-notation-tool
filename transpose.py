"""
transpose.py — Song Notation Tool v0.16 render-time transposition.

Non-destructive: effective = document.transpose + section.transpose + ref.transpose,
applied at render time only. See DESIGN_v0_16.md section 5.

This module is the single source of truth for the chromatic scale and its
enharmonic aliases, fixing the v0.15 bug where "Bb" was defined twice in
ENHARMONIC and the lowercase flat variants (db/eb/gb/ab) were missing.
"""

from __future__ import annotations
from collections import namedtuple

MAX_FRET = 24

CHROMATIC = ["A", "A#", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#"]

ENHARMONIC = {
    "Bb": "A#", "Db": "C#", "Eb": "D#", "Gb": "F#", "Ab": "G#",
    "bb": "A#", "db": "C#", "eb": "D#", "gb": "F#", "ab": "G#",
}

Resolved = namedtuple("Resolved", ["symbol", "fret", "flag"])


def note_to_index(note_str: str):
    n = (note_str or "").strip()
    if n in CHROMATIC:
        return CHROMATIC.index(n)
    if n in ENHARMONIC:
        return CHROMATIC.index(ENHARMONIC[n])
    nu = n[0].upper() + n[1:] if len(n) > 1 else n.upper()
    if nu in CHROMATIC:
        return CHROMATIC.index(nu)
    if nu in ENHARMONIC:
        return CHROMATIC.index(ENHARMONIC[nu])
    return None


def transpose_note(note_str: str, semitones: int) -> str:
    idx = note_to_index(note_str)
    if idx is None:
        return note_str
    return CHROMATIC[(idx + semitones) % 12]


def _parse_root_suffix(value: str):
    v = (value or "").strip()
    if not v:
        return None, ""
    for length in (2, 1):
        root = v[:length]
        canon = None
        if root in CHROMATIC:
            canon = root
        elif root in ENHARMONIC:
            canon = ENHARMONIC[root]
        if canon is not None:
            return canon, v[length:]
    return None, v


def transpose_symbol(symbol: str, semitones: int) -> str:
    """Transpose a chord/note symbol such as 'Am', 'F#7', 'Bbm7', 'A5'."""
    if not symbol or semitones == 0:
        return symbol
    root, suffix = _parse_root_suffix(symbol)
    if root is None:
        return symbol
    return transpose_note(root, semitones) + suffix


def effective_transpose(document_transpose: int, section_transpose: int,
                         ref_transpose: int = 0) -> int:
    return document_transpose + section_transpose + ref_transpose


def resolve(symbol: str, fret, eff: int) -> Resolved:
    """
    Resolve a token for display at effective offset `eff`.
    fret may be None (no position given) or an int.
    Rule: note + n semitones, fret + n (same string, slide up) — section 5.2.
    """
    new_symbol = transpose_symbol(symbol, eff)

    if fret is None:
        return Resolved(new_symbol, None, None)

    new_fret = fret + eff

    if new_fret < 0:
        corrected = new_fret + 12
        if 0 <= corrected <= MAX_FRET:
            return Resolved(new_symbol, corrected, "octave_up")
        return Resolved(new_symbol, None, "unplayable")

    if new_fret > MAX_FRET:
        corrected = new_fret - 12
        if 0 <= corrected <= MAX_FRET:
            return Resolved(new_symbol, corrected, "octave_down")
        return Resolved(new_symbol, None, "unplayable")

    return Resolved(new_symbol, new_fret, None)

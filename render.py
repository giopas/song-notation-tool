"""
render.py — Song Notation Tool v0.16 shared rendering helpers.

Used by both the TXT and PDF exporters so they can no longer disagree
(section 7.1 — the has_tab predicate bug) and so measure wrapping is
computed the same way in both places (section 7.4).
"""

from __future__ import annotations

from grammar import _MARK_TO_TEXT  # noqa: F401  (re-exported for callers)

DEFAULT_TXT_WIDTH = 100
TOKEN_W = 3


def _row_tokens(raw: str):
    return [t for t in (raw or "").split() if t]


def _cell_has_content(raw: str) -> bool:
    return any(t != "-" for t in _row_tokens(raw))


def section_has_tab(tab_cells) -> bool:
    """
    True if any measure in `tab_cells` (a list of {string: row_text} dicts,
    the v0.15/measure-item shape) has real content — not just an empty or
    all-dash row. This is the single predicate both exporters must call;
    previously each computed it separately (inline dict-truthiness) and
    disagreed on sections like an all-notes "Interlude" with an empty tab.
    """
    for cell in (tab_cells or []):
        if not isinstance(cell, dict):
            continue
        if any(_cell_has_content(raw) for raw in cell.values()):
            return True
    return False


def string_has_content(tab_cells, string_name: str) -> bool:
    for cell in (tab_cells or []):
        if not isinstance(cell, dict):
            continue
        if _cell_has_content(cell.get(string_name, "")):
            return True
    return False


def active_strings(tab_cells, strings):
    """Compact-mode helper (section 7.3): drop strings with no content
    anywhere in the section, in the given instrument's string order."""
    return [st for st in strings if string_has_content(tab_cells, st)]


def measures_per_line(n_measures: int, max_beats: int,
                       width: int = DEFAULT_TXT_WIDTH,
                       token_w: int = TOKEN_W, prefix_w: int = 4) -> int:
    """
    How many measures fit on one output line at `width` columns.
    Mirrors the PDF builder's mpl_for() geometry (usable-space / column-width)
    so TXT and PDF wrap on the same logical rule (section 7.4).
    """
    usable = max(width - prefix_w, token_w * max_beats + 1)
    col_w = token_w * max_beats + 1
    return max(1, min(n_measures, usable // col_w))


# ==============================================================================
#  Chart-row rendering (section 7.2) — fret line above symbol line.
# ==============================================================================

def _fret_str(item: dict) -> str:
    if item["kind"] == "token":
        f = item.get("fret")
        return str(f) if f is not None else ""
    return ""


def _symbol_str(item: dict) -> str:
    k = item["kind"]
    if k == "token":
        return item["symbol"]
    if k == "block_ref":
        s = item["block"]
        if item.get("repeat", 1) != 1:
            s += f" (x{item['repeat']})"
        return s
    if k == "section_ref":
        s = f"={item['section']}"
        if item.get("repeat", 1) != 1 or item.get("all"):
            tag = f"(x{item.get('repeat', 1)}"
            tag += " all)" if item.get("all") else ")"
            s += f" {tag}"
        return s
    if k == "group":
        inner = " ".join(_symbol_str(x) for x in item["items"])
        s = f"[{inner}]"
        if item.get("repeat", 1) != 1:
            s += f"(x{item['repeat']})"
        return s
    if k == "mark":
        return _MARK_TO_TEXT.get(item["mark"], item["mark"])
    return ""


def render_chart_row(items, label: str = ""):
    """
    Render one section's chart items as two column-aligned text lines:
    fret numbers above, symbols below. Groups/refs/marks render in the
    symbol row only (no fret line) — section 7.2.
    """
    if not items:
        return [label.rstrip()] if label else []

    frets = [_fret_str(it) for it in items]
    symbols = [_symbol_str(it) for it in items]
    label_w = max(len(label) + 2, 4)
    cols = [max(len(f), len(s)) + 2 for f, s in zip(frets, symbols)]

    fret_line = " " * label_w + "".join(f"{f:<{w}}" for f, w in zip(frets, cols))
    sym_line = f"{label:<{label_w}}" + "".join(f"{s:<{w}}" for s, w in zip(symbols, cols))
    return [fret_line.rstrip(), sym_line.rstrip()]


# ==============================================================================
#  Mark spans (section 6) — resolve ending_1/ending_2 into bracket spans.
# ==============================================================================

def mark_spans(items):
    """
    Resolve 1st/2nd-ending marks into closed spans over the item indices
    they bracket, for rendering a numbered bracket above a run (section 6:
    "endings need a span, not a point"). An ending span opens at an
    ending_1/ending_2 mark and runs to just before the next ending_* mark,
    or to the end of the sequence. The data itself (`items`) stays a flat
    sequence; only this resolver — and its callers — deal in spans.

    Returns a list of (label, start_idx, end_idx) tuples: `label` is "1"
    or "2", and [start_idx, end_idx] (inclusive) indexes the items the
    bracket covers (the mark item itself is not included).
    """
    spans = []
    open_label = None
    open_start = None
    items = items or []
    for idx, it in enumerate(items):
        if it.get("kind") != "mark":
            continue
        m = it.get("mark")
        if m in ("ending_1", "ending_2"):
            if open_label is not None and open_start <= idx - 1:
                spans.append((open_label, open_start, idx - 1))
            open_label = "1" if m == "ending_1" else "2"
            open_start = idx + 1
    if open_label is not None and open_start <= len(items) - 1:
        spans.append((open_label, open_start, len(items) - 1))
    return spans


def has_coda(items) -> bool:
    """True if a 'coda' mark appears anywhere in `items` (not nested in
    groups — the coda mark always sits at the top level of a section)."""
    return any(it.get("kind") == "mark" and it.get("mark") == "coda"
               for it in (items or []))


# ==============================================================================
#  Page-count estimate (section 3.1 — the header's live page indicator).
#
#  This mirrors the geometry the PDF builder uses (measures_per_line) so the
#  number shown while typing tracks the real export, but it is a line-budget
#  estimate, not a run of the actual paginator — the live indicator is meant
#  to be cheap enough to recompute on every edit.
# ==============================================================================

LINES_PER_PAGE = 58


def lyrics_block_line_count(text: str) -> int:
    """Line count a lyrics_text block will occupy when printed — one line
    per source line (no word-wrap accounting, same "rough estimate" spirit
    as the rest of this module), plus a trailing blank separator. 0 for
    blank/whitespace-only text."""
    text = (text or "")
    if not text.strip():
        return 0
    return len(text.splitlines()) + 1  # + trailing blank line


def estimate_section_lines(section: dict, strings=None) -> int:
    """Rough line count `section` will occupy in the TXT/PDF export."""
    items = section.get("items", [])
    annotation = section.get("annotation", "")
    lyrics = section.get("lyrics_text", "") if section.get("print_lyrics") else ""
    if not items and not annotation and not lyrics.strip():
        return 0

    lines = 1  # section label line
    render_mode = section.get("render", "chart")

    if render_mode in ("chart", "both"):
        chart_items = [it for it in items if it.get("kind") != "measure"]
        if chart_items:
            lines += 2  # fret row + symbol row

    if render_mode in ("tab", "both"):
        measures = [it for it in items if it.get("kind") == "measure"]
        if measures:
            max_beats = max((m.get("beats", 8) for m in measures), default=8)
            mpl = measures_per_line(len(measures), max_beats)
            n_rows = len(strings) if strings else 4
            n_groups = -(-len(measures) // mpl)  # ceil division
            lines += n_groups * n_rows

    if annotation:
        lines += 1

    lines += lyrics_block_line_count(lyrics)

    return lines + 1  # trailing blank line between sections


def estimate_page_count(doc: dict, instrument_strings: dict = None,
                         lines_per_page: int = LINES_PER_PAGE) -> int:
    """Live page-count estimate for the song-map header."""
    total = 0
    if doc.get("print_lyrics"):
        total += lyrics_block_line_count(doc.get("lyrics_text", "")) + 1  # + label line
    for sec in doc.get("sections", []):
        strings = None
        if instrument_strings:
            strings = instrument_strings.get(sec.get("instrument"))
        total += estimate_section_lines(sec, strings)
    if total == 0:
        return 1
    return max(1, -(-total // lines_per_page))


# ==============================================================================
#  Render-time transpose resolution — shared by the song map, TXT, PDF, and
#  stage view (section 7.5: "one renderer with two scales, not two code
#  paths"). transpose.py never mutates stored items; this walks a copy.
# ==============================================================================

def resolve_display_items(items, eff_semitones: int):
    """
    Return a shallow-copied item list with every `token`'s symbol/fret
    resolved for display at `eff_semitones` (transpose.resolve). Groups
    are resolved recursively; block_ref/section_ref/measure/mark items
    pass through unchanged — a reference always displays by its own
    name, not by expanding the riff's transposed notes inline.
    """
    import transpose  # local import: keeps render.py's only hard
                       # dependency (grammar) unconditional at module load

    if not eff_semitones:
        return list(items or [])

    out = []
    for it in (items or []):
        kind = it.get("kind")
        if kind == "token":
            r = transpose.resolve(it["symbol"], it.get("fret"), eff_semitones)
            new_it = dict(it)
            new_it["symbol"] = r.symbol
            new_it["fret"] = r.fret
            out.append(new_it)
        elif kind == "group":
            new_it = dict(it)
            new_it["items"] = resolve_display_items(it.get("items", []), eff_semitones)
            out.append(new_it)
        else:
            out.append(it)
    return out

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

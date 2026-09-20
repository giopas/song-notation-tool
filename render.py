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
        # Frets belong inline here: the group renders on the symbol row
        # only (there's no fret row to align a bracketed run against), so
        # dropping them would silently lose "3A 12D" down to "A D" — which
        # is exactly what an expanded x2 reference looks like.
        inner = " ".join(_fret_str(x) + _symbol_str(x) for x in item["items"])
        s = f"[{inner}]"
        if item.get("repeat", 1) != 1:
            s += f"(x{item['repeat']})"
        return s
    if k == "mark":
        return _MARK_TO_TEXT.get(item["mark"], item["mark"])
    return ""


BODY_INDENT = 2   # every body line in the TXT/PDF export starts here


def _lick_lines(item: dict):
    """A lick's printable lines: 'G 5 7 5', one per string, as typed."""
    out = []
    for ln in item.get("lines", []):
        frets = ln.get("frets", [])
        width = max((len(f) for f in frets), default=1)
        cells = " ".join(f"{f:>{width}}" for f in frets)
        out.append(f"{ln.get('string', '')} {cells}")
    return out


def _gutter_width(label: str, indent: int) -> int:
    """Width of the left column: a label's own gutter, or the plain body
    indent when there's no label. Matches render_chart_row exactly."""
    return max(len(label) + 2, 4) if label else indent


# How far one ">" pushes a line in. Wide enough to read as deliberate at a
# glance from a music stand, narrow enough that a couple of levels don't
# eat the page.
INDENT_STEP = 6


def split_on_line_breaks(items):
    """
    `items` split into the runs between `line_break` marks, empties
    dropped. Returns (indent_level, run) pairs — one per printed line,
    where the level is the count of ">" on the break that started it.
    """
    runs, current, level = [], [], 0
    for it in (items or []):
        if it.get("kind") == "mark" and it.get("mark") == "line_break":
            if current:
                runs.append((level, current))
            current, level = [], int(it.get("indent", 0))
        else:
            current.append(it)
    if current:
        runs.append((level, current))
    return runs


def render_chart_row(items, label: str = "", indent: int = BODY_INDENT):
    """
    Render one section's chart items as column-aligned text lines: fret
    numbers above, symbols below, plus a line per string for any lick.

    A `line_break` mark starts a new block, so a long section can be
    grouped into the phrases you'd write out by hand. Each block aligns
    its own columns — that's the point of breaking, rather than having one
    grid stretch across the whole section — and only the first block
    carries the label.

    With no `label`, the rows start at `indent` — the same column the
    export indents its section headers, annotations and free text to, so a
    chart section and a free-text section under the same header line up
    instead of sitting two columns apart. A `label` (the desktop app's
    song map, which prints the section name in a gutter) sizes its own
    column as before.
    """
    if not items:
        return [label.rstrip()] if label else []

    runs = split_on_line_breaks(items)
    if len(runs) != 1:
        # Continuation blocks indent to the label's gutter, so every line
        # of the section stacks under the first one rather than sliding
        # back to the left margin beneath the name.
        cont_indent = _gutter_width(label, indent)
        out = []
        for i, (level, run) in enumerate(runs):
            step = level * INDENT_STEP
            if i == 0:
                out.extend(_render_one_row(run, label, indent + step))
            else:
                out.extend(_render_one_row(run, "", cont_indent + step))
        return out or ([label.rstrip()] if label else [])
    _, items = runs[0]

    return _render_one_row(items, label, indent)


def _render_one_row(items, label: str, indent: int):
    """One block of a chart row — see render_chart_row."""
    if not items:
        return []

    frets = [_fret_str(it) for it in items]
    symbols = [_symbol_str(it) for it in items]

    # A lick occupies the same column as any other item, but stacks its
    # string lines underneath the chord row — so it reads where it's
    # played, in among the chords, rather than in a separate grid.
    licks = [_lick_lines(it) if it.get("kind") == "lick" else [] for it in items]
    n_lick_rows = max((len(l) for l in licks), default=0)

    label_w = max(len(label) + 2, 4) if label else indent
    cols = [max(len(f), len(s), *(len(x) for x in lk) if lk else (0,)) + 2
            for f, s, lk in zip(frets, symbols, licks)]

    fret_line = " " * label_w + "".join(f"{f:<{w}}" for f, w in zip(frets, cols))
    sym_line = f"{label:<{label_w}}" + "".join(f"{s:<{w}}" for s, w in zip(symbols, cols))

    lick_rows = []
    for r in range(n_lick_rows):
        cells = [(lk[r] if r < len(lk) else "") for lk in licks]
        lick_rows.append((" " * label_w +
                          "".join(f"{c:<{w}}" for c, w in zip(cells, cols))).rstrip())

    # Chords with no fret numbers (or a lone group / reference) produce an
    # all-blank fret row. Emitting it anyway put an empty line above the
    # chart in every such section — visible in the export as one section
    # sitting a line lower than its neighbours.
    fret_line = fret_line.rstrip()
    sym_line = sym_line.rstrip()
    rows = ([fret_line, sym_line] if fret_line else [sym_line])
    return rows + [r for r in lick_rows if r.strip()]


# ==============================================================================
#  Reference expansion — show what a reference actually plays.
#
#  A section_ref / block_ref is stored as a pointer (that's what makes
#  "edit the riff once, every user updates" work), but a player reading the
#  chart needs the notes, not the pointer: "=chorus1" tells them nothing,
#  least of all when the id was minted before the section was renamed. So
#  the pointer stays in the data and gets expanded at render time, here.
#
#  Guarded against cycles (A refs B refs A) by carrying the set of keys
#  currently being expanded; a cycle renders as the bare reference rather
#  than recursing forever.
# ==============================================================================

MAX_REF_DEPTH = 8


def _ref_placeholder(item: dict, doc: dict):
    """What to show when a reference can't be expanded — the target's
    name if we can find it, otherwise the key exactly as written."""
    import songmap
    if item.get("kind") == "section_ref":
        new_it = dict(item)
        new_it["section"] = songmap.section_display_name(doc, item.get("section", ""))
        return new_it
    return item


def make_text(text: str, repeat: int = 1):
    """
    A render-time-only item: literal lines, already formatted, that pass
    through to the output untouched.

    Not a stored item kind — nothing ever writes one to a .sng. It exists
    so that expanding a reference to a *free-text* section can carry that
    section's text, which by definition isn't chart items at all.
    """
    return {"kind": "text", "text": text or "", "repeat": repeat}


def _expanded(items, repeat, shift):
    """Wrap an expanded body so its repeat count still reads on the chart:
    a bare run for x1, a bracketed group otherwise."""
    from model import make_group
    body = list(items)
    if shift:
        body = resolve_display_items(body, shift)
    if repeat and repeat != 1:
        return [make_group(body, repeat=repeat)]
    return body


def resolve_references(items, doc: dict, _seen=None, _depth=0):
    """
    Return `items` with every section_ref / block_ref replaced by the items
    it points at, recursively. Anything that can't be resolved — a missing
    target, a cycle, or nesting deeper than MAX_REF_DEPTH — is left as a
    reference so the chart still says *something* rather than silently
    dropping a section's entire content.
    """
    import songmap

    doc = doc or {}
    _seen = _seen or frozenset()
    out = []

    for it in (items or []):
        kind = it.get("kind")

        if kind == "group":
            new_it = dict(it)
            new_it["items"] = resolve_references(
                it.get("items", []), doc, _seen, _depth + 1)
            out.append(new_it)
            continue

        if kind == "section_ref":
            key = it.get("section", "")
            target = songmap.find_section(doc, key)
            marker = ("section", target.get("id") if target else key)
            if target is None or marker in _seen or _depth >= MAX_REF_DEPTH:
                out.append(_ref_placeholder(it, doc))
                continue
            # A section in free-text mode renders as its text, so a
            # reference to it must too. Its `items` may still hold the
            # chart it had before the switch — switching to Free is
            # deliberately non-destructive — and expanding *those* would
            # show content the target itself no longer displays.
            if target.get("render") == "free":
                free = (target.get("free_text") or "").strip()
                if not free:
                    out.append(_ref_placeholder(it, doc))
                    continue
                out.append(make_text(target.get("free_text", ""),
                                     repeat=it.get("repeat", 1)))
                continue

            body = resolve_references(
                [i for i in target.get("items", []) if i.get("kind") != "measure"],
                doc, _seen | {marker}, _depth + 1)
            if not body:
                out.append(_ref_placeholder(it, doc))
                continue
            out.extend(_expanded(body, it.get("repeat", 1), it.get("transpose", 0)))
            continue

        if kind == "block_ref":
            key = it.get("block", "")
            block = (doc.get("blocks") or {}).get(key)
            marker = ("block", key)
            if not block or marker in _seen or _depth >= MAX_REF_DEPTH:
                out.append(it)
                continue
            body = resolve_references(
                [i for i in block.get("items", []) if i.get("kind") != "measure"],
                doc, _seen | {marker}, _depth + 1)
            if not body:
                out.append(it)
                continue
            out.extend(_expanded(body, it.get("repeat", 1), it.get("transpose", 0)))
            continue

        out.append(it)

    return out


def chart_body_lines(items, label: str = "", indent: int = BODY_INDENT):
    """
    Render a section's (already reference-resolved) items to output lines.

    Chart items are column-aligned by render_chart_row; a `text` item —
    which only ever comes from expanding a reference to a free-text
    section — passes through as its own lines, because free text is not
    column data and aligning it would corrupt it. Consecutive chart items
    are grouped so a mixed run still aligns within each stretch, and only
    the first line of the section carries the label.
    """
    lines = []
    run = []
    label_used = [False]

    def take_label():
        if label and not label_used[0]:
            label_used[0] = True
            return label
        return ""

    def flush():
        if run:
            lines.extend(render_chart_row(run, take_label(), indent))
            run.clear()

    for it in (items or []):
        if it.get("kind") != "text":
            run.append(it)
            continue

        flush()
        text_lines = (it.get("text") or "").splitlines() or [""]
        this_label = take_label()
        width = _gutter_width(label, indent)
        lines.append(f"{this_label:<{width}}" + text_lines[0])
        lines.extend(" " * width + ln for ln in text_lines[1:])
        if it.get("repeat", 1) != 1:
            lines[-1] += f"  (x{it['repeat']})"

    flush()
    return lines


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


def free_text_line_count(text: str) -> int:
    """Line count a section's verbatim free-text block occupies when
    printed — one line per source line, plus a trailing blank separator.
    0 for blank/whitespace-only text."""
    text = (text or "")
    if not text.strip():
        return 0
    return len(text.splitlines()) + 1


def estimate_section_lines(section: dict, strings=None, doc: dict = None) -> int:
    """Line count `section` will occupy in the TXT/PDF export.

    With `doc`, references are expanded and the chart is counted by
    rendering it — the only way to be exact, and the export's
    keep-a-section-whole rule is only as good as this number. A section
    whose entire content is `=Chorus_1` counts as one item without the
    document and as the seven lines Chorus_1 actually prints with it.
    """
    items = section.get("items", [])
    annotation = section.get("annotation", "")
    lyrics = section.get("lyrics_text", "") if section.get("print_lyrics") else ""
    render_mode = section.get("render", "chart")
    free = section.get("free_text", "") if render_mode == "free" else ""
    if not items and not annotation and not lyrics.strip() and not free.strip():
        return 0

    lines = 1  # section label line

    if render_mode == "free":
        lines += free_text_line_count(free)

    if render_mode in ("chart", "both"):
        chart_items = [it for it in items if it.get("kind") != "measure"]
        if doc is not None:
            chart_items = resolve_references(chart_items, doc)
        if chart_items:
            if doc is not None:
                lines += len(chart_body_lines(chart_items, "", 0))
            else:
                # one fret + symbol pair per block, plus however many string
                # lines the tallest lick in each block needs
                for _level, run in split_on_line_breaks(chart_items) or [(0, [])]:
                    lines += 2
                    lines += max((len(it.get("lines", [])) for it in run
                                  if it.get("kind") == "lick"), default=0)

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
        total += estimate_section_lines(sec, strings, doc)
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
        elif kind == "lick":
            # A lick is fret positions on named strings, so transposing it
            # is arithmetic on the frets — the strings don't move. Anything
            # that would fall off either end of the neck is left alone
            # rather than silently clamped to a fret you'd actually play.
            new_lines = []
            for ln in it.get("lines", []):
                frets = []
                for f in ln.get("frets", []):
                    if f in ("-", "x", "X"):
                        frets.append(f)
                        continue
                    shifted = int(f) + eff_semitones
                    frets.append(str(shifted) if 0 <= shifted <= 24 else f)
                new_lines.append({"string": ln.get("string", ""), "frets": frets})
            new_it = dict(it)
            new_it["lines"] = new_lines
            out.append(new_it)
        elif kind == "group":
            new_it = dict(it)
            new_it["items"] = resolve_display_items(it.get("items", []), eff_semitones)
            out.append(new_it)
        else:
            out.append(it)
    return out

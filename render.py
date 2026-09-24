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
        #
        # A group reaching this point is always the simple case: plain
        # chords (and/or a nested group of plain chords) that fit on one
        # line. A group that holds a lick or a line break is unwrapped
        # before rendering ever gets here — see _flatten_bracket_groups —
        # because neither a lick's tab nor a "//" survives being flattened
        # into inline bracket text.
        inner = " ".join(_fret_str(x) + _symbol_str(x) for x in item["items"])
        s = f"[{inner}]"
        if item.get("repeat", 1) != 1:
            s += f"(x{item['repeat']})"
        return s
    if k == "lick":
        # A named lick prints its name over its tab, so the same figure
        # recalled later as "{Riff1}" is recognisable as the thing you
        # already read. An unnamed one has nothing to say here.
        s = item.get("name", "") or ""
        if s and item.get("repeat", 1) != 1:
            s += f" (x{item['repeat']})"
        return s
    if k == "lick_ref":
        # A reference that found its lick prints as the bare name; one
        # that didn't keeps its braces, so a typo is visible on paper.
        s = (item.get("lick", "") if item.get("resolved")
             else "{" + item.get("lick", "") + "}")
        if item.get("repeat", 1) != 1:
            s += f" (x{item['repeat']})"
        return s
    if k == "mark":
        return _MARK_TO_TEXT.get(item["mark"], item["mark"])
    return ""


BODY_INDENT = 2   # every body line in the TXT/PDF export starts here


def _lick_lines(item: dict):
    """
    A lick's printable lines, one per string, written the way tab is
    written rather than as a bare list of numbers:

        |G|-5--7--5-|
        |D|-------3-|

    Every position is the same width across every line of the lick, so
    the columns line up vertically and a player reads down the stack the
    way they would read a tab staff. A fret wider than one digit widens
    every cell in that lick, not just its own.
    """
    lines = item.get("lines", [])
    if not lines:
        return []
    fret_w = max((len(f) for ln in lines for f in ln.get("frets", [])),
                 default=1) or 1
    name_w = max((len(str(ln.get("string", ""))) for ln in lines), default=1)
    out = []
    for ln in lines:
        cells = "".join("-" + str(f).ljust(fret_w, "-")
                        for f in ln.get("frets", []))
        out.append(f"|{str(ln.get('string', '')):<{name_w}}|{cells}-|")
    return out


def _gutter_width(label: str, indent: int) -> int:
    """Width of the left column: a label's own gutter, or the plain body
    indent when there's no label. Matches render_chart_row exactly."""
    return max(len(label) + 2, 4) if label else indent


def _column_lick_height(it: dict) -> int:
    """A bare lick's own height in extra tab lines. A group never reaches
    here holding one — see _flatten_bracket_groups — so this only ever
    needs to look at the item itself."""
    if it.get("kind") == "lick":
        return len(it.get("lines", []))
    return 0


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


def _needs_bracket(items) -> bool:
    """True when a group's own content can't render as the single
    bracketed line _symbol_str builds — either it holds a lick (whose tab
    needs rows of its own underneath) or it breaks ("//"), recursing into
    any group nested inside it. Plain chords, however many, are always
    fine on one line."""
    for x in (items or []):
        k = x.get("kind")
        if k == "lick":
            return True
        if k == "mark" and x.get("mark") == "line_break":
            return True
        if k == "group" and _needs_bracket(x.get("items")):
            return True
    return False


def _flatten_bracket_groups(items):
    """A group whose content needs more than one printed row — a lick, a
    "//", or both — is unwrapped here, before rendering, into the flat
    sequence it stands for. Collapsing either one into the group's usual
    single bracketed line is wrong: a lick nested in a group used to
    print its bare name instead of its tab, and a "//" nested in a group
    used to print as the literal text "//" instead of starting a new
    line. Unwrapped, a lick prints its tab and a "//" breaks the line
    exactly as either would at the top level — because, once unwrapped,
    that's exactly what they are.

    The group's own repeat count, when it has one, would have nowhere
    correct to go as inline text once its content might span several
    rows (which row would "(xN)" belong to?), so it isn't typed in here
    at all. Instead the first and last item of what the group expands to
    are tagged (`_group_bracket_start` / `_group_bracket_end`) —
    render_chart_rows turns that into an actual bracket, drawn down the
    right of every row the group ends up printing on, once it knows
    where each one landed (see _apply_group_brackets).

    A group needing none of this — plain chords, however many — is left
    exactly alone; it still renders as the single bracketed column
    _symbol_str builds."""
    out = []
    for it in (items or []):
        if it.get("kind") == "group" and _needs_bracket(it.get("items")):
            inner = _flatten_bracket_groups(it.get("items", []))
            rep = it.get("repeat", 1)
            if inner and rep != 1:
                if len(inner) == 1:
                    inner = [dict(inner[0], _group_bracket_start=True,
                                  _group_bracket_end=rep)]
                else:
                    inner = ([dict(inner[0], _group_bracket_start=True)]
                              + inner[1:-1]
                              + [dict(inner[-1], _group_bracket_end=rep)])
            out.extend(inner)
            continue
        out.append(it)
    return out


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
    return [r["text"] for r in render_chart_rows(items, label, indent)]


# A row (and, inside the symbol row, a single item) carries a role, so the
# PDF and the browser preview can colour what they draw without re-parsing
# the text they were handed. Roles, not colours: the palette belongs to
# whoever is drawing.
ROLE_FRET = "fret"
ROLE_SYM = "sym"
ROLE_LICK = "lick"
ROLE_REST = "rest"
ROLE_TEXT = "text"


def _row(text: str, role: str, spans=None):
    return {"text": text, "role": role, "spans": list(spans or [])}


def _apply_group_brackets(rows, run_ranges):
    """Turn a _flatten_bracket_groups tag into an actual bracket: a "|" on
    the right of every printed row a tagged group produced, padded to a
    common width so they line up in one column, with the group's "(xN)"
    once, on the vertically-centred row among them — so it reads as
    belonging to the whole bracketed run, not to whichever chord, lick or
    line happens to sit on the last row of it.

    `run_ranges` is a list of (row_start, row_end, run) — the slice of
    `rows` each processed run produced, and the items that made it, in
    the same order render_chart_rows built them in. A group with no tag
    anywhere (the overwhelmingly common case) leaves `rows` untouched.
    """
    pending, spans = [], []
    for row_start, row_end, run in run_ranges:
        for it in run:
            if it.get("_group_bracket_start"):
                pending.append(row_start)
            rep = it.get("_group_bracket_end")
            if rep:
                s = pending.pop() if pending else row_start
                spans.append((s, row_end, rep))
    if not spans:
        return rows
    for s, e, rep in spans:
        # A fret row is set as a superscript figure over the row below it
        # — its own colour, a smaller size, and tighter line spacing, in
        # every renderer that draws one — not a line in its own right.
        # Giving it a bracket segment fought that styling instead of
        # sitting inside it: a stray mark in the fret colour, at fret
        # size, out of step with the bar around it. A broken group can
        # have more than one — every line it breaks onto gets its own —
        # so every row in the span is checked, not just the first.
        idxs = [i for i in range(s, e + 1) if rows[i]["role"] != ROLE_FRET]
        if not idxs:
            idxs = list(range(s, e + 1))
        # Centred on the rows that actually carry the bar; rounded down
        # on a tie so a two-line break puts "(xN)" on the first line, not
        # the last — landing there reads as "this last chord repeats"
        # instead of "this whole phrase does".
        mid = idxs[(len(idxs) - 1) // 2]
        width = max((len(rows[i]["text"]) for i in idxs), default=0)
        for i in idxs:
            pad = " " * (width - len(rows[i]["text"]))
            tail = "  |" + (f" (x{rep})" if i == mid else "")
            rows[i]["text"] += pad + tail
    return rows


def render_chart_rows(items, label: str = "", indent: int = BODY_INDENT):
    """
    render_chart_row's rows, each as a dict:

        {"text": str, "role": "fret"|"sym"|"lick"|"text",
         "spans": [(start_col, end_col, role), ...]}

    `spans` marks stretches *within* a row that a drawing front end should
    treat differently — a rest, which prints grey, is the only one so far.
    Column indices are into `text`, and the text is monospace everywhere it
    is drawn, so a span converts to an x offset by multiplying.
    """
    if not items:
        return [_row(label.rstrip(), ROLE_SYM)] if label else []

    items = _flatten_bracket_groups(items)
    runs = split_on_line_breaks(items)
    if len(runs) != 1:
        # Continuation blocks indent to the label's gutter, so every line
        # of the section stacks under the first one rather than sliding
        # back to the left margin beneath the name.
        cont_indent = _gutter_width(label, indent)
        out, run_ranges = [], []
        for i, (level, run) in enumerate(runs):
            step = level * INDENT_STEP
            before = len(out)
            if i == 0:
                out.extend(_render_one_row(run, label, indent + step))
            else:
                out.extend(_render_one_row(run, "", cont_indent + step))
            run_ranges.append((before, len(out) - 1, run))
        out = _apply_group_brackets(out, run_ranges)
        return out or ([_row(label.rstrip(), ROLE_SYM)] if label else [])
    _, one_run = runs[0]

    rows = _render_one_row(one_run, label, indent)
    return _apply_group_brackets(rows, [(0, len(rows) - 1, one_run)])


def _render_one_row(items, label: str, indent: int):
    """One block of a chart row — see render_chart_row."""
    if not items:
        return []

    frets = [_fret_str(it) for it in items]
    symbols = [_symbol_str(it) for it in items]

    # A lick occupies the same column as any other item, but stacks its
    # string lines underneath the chord row — so it reads where it's
    # played, in among the chords, rather than in a separate grid. (A
    # group never reaches here holding one of its own — see
    # _flatten_bracket_groups — so only a bare lick item needs this.)
    licks = [_lick_lines(it) if it.get("kind") == "lick" else [] for it in items]

    # A named lick's name goes on the tab itself, in front of its first
    # string line, rather than on the chord row above it: on a row of its
    # own the name read as a separate line of the chart, one line away
    # from the notes it names. The other string lines are indented by the
    # same amount, so the tab stays a grid.
    for i, (it, lk) in enumerate(zip(items, licks)):
        if lk and symbols[i]:
            head = symbols[i] + " "
            licks[i] = [head + lk[0]] + [" " * len(head) + ln for ln in lk[1:]]
            symbols[i] = ""
    n_lick_rows = max((len(l) for l in licks), default=0)

    label_w = max(len(label) + 2, 4) if label else indent
    cols = [max(len(f), len(s), *(len(x) for x in lk) if lk else (0,)) + 2
            for f, s, lk in zip(frets, symbols, licks)]

    fret_line = " " * label_w + "".join(f"{f:<{w}}" for f, w in zip(frets, cols))
    sym_line = f"{label:<{label_w}}" + "".join(f"{s:<{w}}" for s, w in zip(symbols, cols))

    # Where each item starts on the symbol row, so a rest can be drawn in
    # its own colour without the drawing code having to find it in the text.
    sym_spans, col_x = [], label_w
    for it, sym, w in zip(items, symbols, cols):
        if sym and it.get("kind") == "mark" and it.get("mark") == "rest":
            sym_spans.append((col_x, col_x + len(sym), ROLE_REST))
        # A lick's name, or a reference to one, is the same thing as the
        # tab it stands for, so it prints in the same blue — findable at
        # a glance among the chord symbols, which is the point of naming it.
        elif sym and it.get("kind") in ("lick", "lick_ref"):
            sym_spans.append((col_x, col_x + len(sym), ROLE_LICK))
        col_x += w

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
    rows = ([_row(fret_line, ROLE_FRET)] if fret_line else [])
    rows.append(_row(sym_line, ROLE_SYM, sym_spans))
    rows += [_row(r, ROLE_LICK) for r in lick_rows if r.strip()]
    # A lick with nothing beside it is a tab block standing on its own,
    # and the blank symbol row above is what sets it apart from the chart
    # line before it. Close it on the other side too: without that, the
    # next line of chords butts straight up against the tab and reads as
    # part of it — the space has to be the same either side or it stops
    # saying "this bit is tab".
    if lick_rows and not sym_line.strip():
        rows.append(_row("", ROLE_SYM))
    return rows


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

        if kind == "lick_ref":
            name = it.get("lick", "")
            target = songmap.find_lick(doc, name)
            marker = ("lick", songmap._ident(name))
            if target is None or marker in _seen or _depth >= MAX_REF_DEPTH:
                out.append(it)
                continue
            if (doc.get("lick_refs") or "tab") == "name":
                # Compact: the name and the count, the way a handwritten
                # chart says "RIFF 1 (x3)" — the notes print once, where
                # the lick was written. Spelled as the definition spells
                # it, whatever case the reference was typed in.
                out.append({"kind": "lick_ref", "lick": target.get("name", name),
                            "repeat": it.get("repeat", 1), "resolved": True})
                continue
            # A lick resolves to the notes themselves, carrying its name
            # and this reference's repeat count — "Riff1 (x3)" over the
            # tab, which is how the handwritten charts say it.
            new_it = dict(target)
            if it.get("repeat", 1) != 1:
                new_it["repeat"] = it["repeat"]
            else:
                new_it.pop("repeat", None)
            out.append(new_it)
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
    return [r["text"] for r in chart_body_rows(items, label, indent)]


def chart_body_rows(items, label: str = "", indent: int = BODY_INDENT):
    """chart_body_lines' output, tagged — see render_chart_rows."""
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
            lines.extend(render_chart_rows(run, take_label(), indent))
            run.clear()

    for it in (items or []):
        if it.get("kind") != "text":
            run.append(it)
            continue

        flush()
        text_lines = (it.get("text") or "").splitlines() or [""]
        this_label = take_label()
        width = _gutter_width(label, indent)
        lines.append(_row(f"{this_label:<{width}}" + text_lines[0], ROLE_TEXT))
        lines.extend(_row(" " * width + ln, ROLE_TEXT) for ln in text_lines[1:])
        if it.get("repeat", 1) != 1:
            lines[-1]["text"] += f"  (x{it['repeat']})"

    flush()
    # A tab block closes with a blank line, but not when it ends the
    # section — the gap before the next section is already there.
    while lines and not lines[-1]["text"].strip():
        lines.pop()
    return lines


# ==============================================================================
#  Lyrics on the page
#
#  Words under the chart is the obvious layout and the wrong one: it puts
#  the thing you glance at (what the bass is doing here) and the thing you
#  read (the words) in the same column, so a verse pushes the next
#  section's chart off the bottom of the sheet.
#
#  Beside it instead: the chart keeps a narrow left column, the words run
#  down the right of it. Nothing is aligned chord-to-syllable — that isn't
#  what this is for, and pretending otherwise would be a lie about where
#  the changes fall. It says: *during these words, this is what you play*.
# ==============================================================================

LYRIC_GAP = 4          # blank columns between the chart and the words
MIN_LYRIC_W = 24       # narrower than this and the words wrap to nothing


def printable_lyrics(text: str) -> list:
    """A lyric block's printable lines — its section markers removed,
    because those are structure for the editor, not words to sing."""
    import lyrics as lyrics_mod
    text = lyrics_mod.strip_markers(text or "")
    return text.splitlines() if text.strip() else []


# Where the words go. One setting for the whole song, chosen on the main
# screen: either every section's words gathered into one block (at the
# start, at the end, or down a column of their own), or each section's
# words with its chart (beside it or under it) — or none at all.
LYRICS_MODES = ("none", "start", "end", "side", "beside", "below")
LYRICS_GATHERED = ("start", "end", "side")
LYRICS_PER_SECTION = ("beside", "below")


def lyrics_mode(doc: dict) -> str:
    """This document's lyric placement, from LYRICS_MODES."""
    # Missing means a song the loader hasn't seen (a bare dict in a test,
    # a hand-built document): nothing prints unless asked for, which is
    # how lyrics have behaved since they were added.
    mode = ((doc or {}).get("lyrics_layout") or "none").strip().lower()
    return mode if mode in LYRICS_MODES else "none"


def lyrics_beside(doc: dict) -> bool:
    """True when each section prints its words beside its chart."""
    return lyrics_mode(doc) == "beside"


def section_lyric_lines(doc: dict, sec: dict) -> list:
    """The words this section prints *with its chart* — none unless the
    song's lyrics are placed per section. In the gathered modes the same
    words print once, together, somewhere else."""
    if doc is not None and lyrics_mode(doc) not in LYRICS_PER_SECTION:
        return []
    return printable_lyrics(sec.get("lyrics_text", ""))


def lyric_blocks(doc: dict) -> list:
    """Every lyric the song has, as [{"name", "lines"}] in song order.

    Sections' own words come first in priority: they're what the sheet
    was split into. Only when no section has any does the whole-song
    sheet stand in — with its "=== Verse 1 ===" markers promoted from
    structure to headings, which is the one place they earn ink.
    """
    blocks = []
    for sec in (doc or {}).get("sections", []) or []:
        lines = printable_lyrics(sec.get("lyrics_text", ""))
        if lines:
            blocks.append({"name": sec.get("name", ""), "lines": lines})
    if blocks:
        return blocks
    import lyrics as lyrics_mod
    for seg in lyrics_mod.split_marked((doc or {}).get("lyrics_text", "")):
        lines = printable_lyrics(seg.get("text", ""))
        if lines:
            blocks.append({"name": seg.get("name", ""), "lines": lines})
    return blocks


def gathered_lyrics(doc: dict):
    """(position, blocks) for the words that print as one block — or
    (None, []) when there's nothing to gather.

    Position is "start", "end" or "side". A per-section layout with no
    section lyrics but a filled-in sheet gathers the sheet at the start
    instead, so words the song has are never silently left off the page.
    """
    mode = lyrics_mode(doc)
    if mode == "none":
        return None, []
    if mode in LYRICS_GATHERED:
        blocks = lyric_blocks(doc)
        return (mode, blocks) if blocks else (None, [])
    has_section_words = any(printable_lyrics(sec.get("lyrics_text", ""))
                            for sec in (doc or {}).get("sections", []) or [])
    if has_section_words:
        return None, []
    blocks = lyric_blocks(doc)
    return ("start", blocks) if blocks else (None, [])


def lyric_block_items(blocks, width: int = None) -> list:
    """Gathered lyric blocks as [(kind, text)] — kind "head" for a section
    heading, "line" for a line of words (indented two), "gap" between
    blocks. Lines longer than `width` wrap: a column of words is narrower
    than a verse was typed for, and running off the page is not an option.
    """
    import textwrap
    out = []
    for i, b in enumerate(blocks or []):
        if i:
            out.append(("gap", ""))
        if b.get("name"):
            out.append(("head", b["name"]))
        for ln in b.get("lines", []):
            if not ln.strip():
                out.append(("line", ""))
            elif width and len(ln) + 2 > width:
                # A wrapped line hangs its continuation further in, so it
                # reads as the rest of the line above rather than a new one.
                parts = textwrap.wrap(ln, max(8, width - 4))
                out += [("line", ("  " if j == 0 else "    ") + w)
                        for j, w in enumerate(parts)]
            else:
                out.append(("line", "  " + ln))
    return out


def lyric_block_lines(blocks, width: int = None, indent: int = 0) -> list:
    """The same, as plain text lines — for the TXT export."""
    pad = " " * indent
    return [pad + text if text else "" for _kind, text in
            lyric_block_items(blocks, None if width is None else width - indent)]


def lyric_column_x(chart_rows, indent: int = BODY_INDENT) -> int:
    """The column the words start in: clear of the widest chart line, and
    never so far right that there's no room left to read them."""
    widest = max((len(r["text"] if isinstance(r, dict) else r)
                  for r in (chart_rows or [])), default=0)
    return max(widest + LYRIC_GAP, indent + LYRIC_GAP)


def shared_lyric_column(doc: dict, instruments=None) -> int:
    """The one column every section's words start in, in the "beside"
    layout — measured from the start of the chart, not the page edge.

    Each section placing its words just past its *own* chart put a narrow
    chorus's words in one column and a wide verse's in another, so down
    the page they zigzagged. One column for the whole song, clear of the
    widest chart that has words next to it, lines them all up — a column
    you can read down, which is what makes it a column.

    Sections without words don't count: a long instrumental line has
    nothing beside it to push along.
    """
    import songmap
    import transpose
    widest = 0
    for sec in (doc or {}).get("sections", []) or []:
        if instruments is not None and sec.get("instrument") not in instruments:
            continue
        if not section_lyric_lines(doc, sec):
            continue
        if sec.get("render", "chart") not in ("chart", "both"):
            continue
        eff = transpose.effective_transpose(doc.get("transpose", 0),
                                            sec.get("transpose", 0))
        chart = resolve_references(songmap.chart_items(sec), doc)
        if not chart:
            continue
        rows = chart_body_lines(resolve_display_items(chart, eff), "", 0)
        widest = max(widest, max((len(r) for r in rows), default=0))
    return widest + LYRIC_GAP


def compose_beside(chart_lines, lyric_lines, col_x: int = None,
                    indent: int = BODY_INDENT) -> list:
    """`chart_lines` with `lyric_lines` laid out in a column to their
    right, as plain text lines.

    Whichever side is taller decides how many lines come back; the chart
    is never truncated to fit the words, nor the other way round.
    """
    chart_lines = [ln for ln in (chart_lines or [])]
    lyric_lines = [ln for ln in (lyric_lines or [])]
    if not lyric_lines:
        return chart_lines
    if col_x is None:
        col_x = lyric_column_x(chart_lines, indent)
    out = []
    for i in range(max(len(chart_lines), len(lyric_lines))):
        left = chart_lines[i] if i < len(chart_lines) else ""
        right = lyric_lines[i] if i < len(lyric_lines) else ""
        if not right:
            out.append(left.rstrip())
            continue
        out.append(f"{left:<{col_x}}" + right)
    return out


def beside_line_count(chart_lines, lyric_lines) -> int:
    """How many lines the pair occupies side by side."""
    return max(len(chart_lines or []), len(lyric_lines or []))


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
    lines = printable_lyrics(text)
    return len(lines) + 1 if lines else 0  # + trailing blank line


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
    if doc is not None:
        lyrics = "\n".join(section_lyric_lines(doc, section))
    else:
        lyrics = section.get("lyrics_text", "") if section.get("print_lyrics") else ""
    beside = lyrics_beside(doc) if doc is not None else True
    chart_lines_drawn = 0
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
                chart_lines_drawn = len(chart_body_lines(chart_items, "", 0))
                lines += chart_lines_drawn
            else:
                # one fret + symbol pair per block, plus however many string
                # lines the tallest lick in each block needs. Flatten first
                # — a group holding a lick or a "//" unwraps into real
                # top-level rows here too (see _flatten_bracket_groups), or
                # this would undercount it the same way a real render no
                # longer does.
                flat = _flatten_bracket_groups(chart_items)
                for _level, run in split_on_line_breaks(flat) or [(0, [])]:
                    lines += 2
                    lines += max((_column_lick_height(it) for it in run),
                                 default=0)

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

    if beside and chart_lines_drawn:
        # The words share the chart's lines; only the overhang costs more.
        lyric_lines = len(printable_lyrics(lyrics))
        lines += max(0, lyric_lines - chart_lines_drawn) + (1 if lyric_lines else 0)
    else:
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

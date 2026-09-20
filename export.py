"""
export.py — Song Notation Tool export engine (core, no UI dependency).

Builds the TXT lines and PDF bytes for a document dict shaped like
model.py's schema. This used to live as SongNotationApp methods in
song_writer.py, reading title/artist/key/etc. from live Tkinter
StringVars; it now reads the same values from doc["meta"] so it can be
called headlessly (cli.py) or from a web request (webserver.py) with no
Tkinter import at all — see the "browser and headless use" enhancement
in the UX & Multi-Platform spec, and ROADMAP.md's "architecture fork"
item, which this closes: the pure-data layers (model/grammar/transpose/
render) already didn't care about the UI; this makes the export layer
join them.

    build_song_lines(doc, instruments=None)              -> list[str]
    build_pdf(doc, instruments=None, orient="portrait")   -> bytes

Both take an optional `instruments` set/iterable to filter which
sections are included (None = all instruments present in the doc).
"""

from __future__ import annotations

import datetime
import zlib

import render
import songmap
import transpose
from constants import (APP_VERSION, APP_URL, INSTRUMENT_STRINGS, LICK_RGB, REST_RGB,
                        TAB_BEATS_DEFAULT, MAX_PDF_COLUMNS, PDF_COLUMN_GAP,
                        PDF_COLUMN_RULE_GRAY, section_color, heading_fill,
                        title_fill)

DEFAULT_TXT_WIDTH = 100


def _meta(doc):
    return doc.get("meta", {}) or {}


def _all_instruments(doc):
    return {s.get("instrument", "") for s in doc.get("sections", [])}


# ==============================================================================
#  TXT export
# ==============================================================================

def uniform_instrument(doc: dict, instruments=None):
    """The one instrument every printed section uses, or None if they
    differ. A song played entirely on one bass doesn't need that repeated
    on every section heading — it's a fact about the song, so it belongs
    once, in the header, next to the key and the tempo."""
    names = {sec.get("instrument", "") for sec in doc.get("sections", [])
             if instruments is None or sec.get("instrument") in instruments}
    names.discard("")
    return next(iter(names)) if len(names) == 1 else None


# Helvetica's average advance is around 0.55 em for mixed-case text and a
# little wider for the all-caps title; 0.58 is a deliberately pessimistic
# single number, since guessing narrow here means text off the page.
_PROPORTIONAL_EM = 0.58


def _fit_text_size(text: str, max_width: float, desired: float) -> float:
    """`desired`, or smaller if the text wouldn't fit in `max_width`."""
    n = len(text or "")
    if n <= 0:
        return desired
    return min(desired, max_width / (n * _PROPORTIONAL_EM))


def _mix(rgb, toward_grey: float):
    """`rgb` faded toward mid-grey — for secondary text (annotations, tab
    string labels) that should read as part of its section without
    competing with the notes themselves."""
    return tuple(c + (0.45 - c) * toward_grey for c in rgb)


# ==============================================================================
#  Page geometry
#
#  A page is margins, a footer, and one or more text columns between them.
#  Everything that measures the chart — the width bound the fit works
#  against, the tab grid's measures-per-line — asks for a *column* width
#  rather than a page width, so the same code serves one column or two.
# ==============================================================================

PAGE_MARGIN = 28

# The body is set in Courier, whose advance is exactly 0.6 em — and that
# advance *is* the column arithmetic: every x offset in a chart row, the
# tab grid's measure columns, the width bound the fit works against. Size
# and character width therefore move together; change one without the
# other and the text stops landing where the geometry says it does.
#
# 9pt is the size at 100%. Courier sets small for its point size, so the
# old 7.5 read a good deal smaller than the headings next to it — and a
# chart is read at arm's length, off a stand, which is the whole argument
# for the largest type the page will take.
MONO_SIZE = 9.0
MONO_CHAR_W = MONO_SIZE * 0.6
# Width of the string-label column beside a tab grid ("G|"), at 100%.
STRING_LABEL_W = 20.0


def page_size(orient: str):
    """(width, height) of an A4 sheet in points, in the given orientation."""
    return (842, 595) if orient == "landscape" else (595, 842)


def column_width(orient: str, columns: int = 1) -> float:
    """Width of a single text column: the page less its margins, less the
    gutters between columns, divided by however many columns there are."""
    n = max(1, min(MAX_PDF_COLUMNS, int(columns or 1)))
    W = page_size(orient)[0]
    return (W - 2 * PAGE_MARGIN - PDF_COLUMN_GAP * (n - 1)) / n


# ==============================================================================
#  Fit-to-page
#
#  The chart is monospace and its geometry is linear in the type size, so
#  the largest scale that still fits is found by: (a) an exact width bound
#  — the longest line must stay inside the margins — and (b) a bisection on
#  height, where "fits" means "needs no more pages than it did at 100%".
#  A build is a few hundred microseconds, so a dozen of them is nothing.
# ==============================================================================

MAX_FIT_SCALE = 4.0
# Fit is allowed to shrink, not only grow: a chart whose longest line
# already runs off the edge of the paper is not "fitted" by leaving it
# there. The floor stops it shrinking into illegibility — past that point
# the honest answer is landscape, or fewer measures per line.
MIN_FIT_SCALE = 0.6
# "Fit to one page" is an explicit instruction to get the whole chart onto
# one sheet, so it is allowed to shrink further than plain fit would — but
# not without limit: below this the chart is unreadable, and the honest
# answer is landscape or fewer sections.
MIN_ONE_PAGE_SCALE = 0.3
_FIT_ITERATIONS = 10


def _body_lines_for_width(doc: dict, instruments) -> list[str]:
    """Every monospace line the PDF will draw, for measuring purposes.
    Tab grids are excluded: their column count adapts to the page width
    on its own (mpl_for), so they can't overflow it."""
    out = []
    one = uniform_instrument(doc, instruments)
    gutter = doc.get("section_layout") == "gutter"
    gutter_w = 0
    if gutter:
        gutter_w = max((len(_gutter_label(s)) for s in doc.get("sections", [])),
                        default=0) + 2
    for sec in doc.get("sections", []):
        if instruments is not None and sec.get("instrument") not in instruments:
            continue
        pad = " " * gutter_w
        if sec.get("render") == "free":
            out += [pad + ln for ln in (sec.get("free_text") or "").splitlines()]
        else:
            eff = transpose.effective_transpose(
                doc.get("transpose", 0), sec.get("transpose", 0))
            chart = render.resolve_references(songmap.chart_items(sec), doc)
            if chart:
                out += [pad + ln for ln in render.chart_body_lines(
                    render.resolve_display_items(chart, eff), "", 0)]
        if sec.get("annotation"):
            out.append(pad + f'"{sec["annotation"]}"')
        if sec.get("print_lyrics"):
            out += [pad + ln for ln in (sec.get("lyrics_text") or "").splitlines()]
        if not gutter:
            rep = sec.get("repeat", 1)
            rep_str = f" (x{rep})" if rep and rep != 1 else ""
            instr = "" if one else f"   {sec.get('instrument', '')}"
            out.append(f"{sec.get('name', '')}{rep_str}{instr}")
    return out


def _width_limited_scale(doc: dict, instruments, orient: str,
                          columns: int = 1) -> float:
    usable = column_width(orient, columns)
    longest = max((len(ln) for ln in _body_lines_for_width(doc, instruments)), default=0)
    if longest <= 0:
        return MAX_FIT_SCALE
    return usable / (longest * MONO_CHAR_W)


def _tab_block_units(doc: dict, instruments) -> float:
    """The narrowest a tab block can be drawn, in points at 100%.

    A tab grid wraps to the width it is given — but not below one measure
    per line, so a wide measure sets a floor the column has to clear. The
    numbers mirror the drawing code: SN_W for the string labels, then
    TOKEN_W characters per beat plus the closing bar line.
    """
    beats = max(
        (m.get("beats", TAB_BEATS_DEFAULT)
         for sec in doc.get("sections", [])
         if instruments is None or sec.get("instrument") in instruments
         if sec.get("render", "chart") in ("tab", "both")
         for m in songmap.measure_items(sec)),
        default=0,
    )
    return 0.0 if not beats else STRING_LABEL_W + (3 * beats + 1) * MONO_CHAR_W


def _fit_scale(doc: dict, instruments, orient: str, columns: int = 1) -> float:
    """Largest scale that keeps the chart on the same number of pages it
    needed at 100%, and inside the horizontal margins."""
    width_cap = _width_limited_scale(doc, instruments, orient, columns)
    if width_cap < 1.0:
        # Too wide even at 100% — shrink to the width that fits.
        return max(MIN_FIT_SCALE, width_cap)

    _, base_pages = _build_pdf(doc, instruments, orient, 1.0, columns)
    hi_cap = min(MAX_FIT_SCALE, width_cap)
    if hi_cap <= 1.0:
        return 1.0

    lo, hi = 1.0, hi_cap
    if _build_pdf(doc, instruments, orient, hi, columns)[1] <= base_pages:
        return hi
    for _ in range(_FIT_ITERATIONS):
        mid = (lo + hi) / 2
        if _build_pdf(doc, instruments, orient, mid, columns)[1] <= base_pages:
            lo = mid
        else:
            hi = mid
    return lo


def _one_page_scale(doc: dict, instruments, orient: str,
                     columns: int = 1) -> float:
    """Largest scale that gets the whole chart onto a single page.

    Unlike `_fit_scale`, which keeps the page count it already had and
    grows into the space left over, this one is prepared to shrink: the
    setting says "one page", so the page count is the constraint and the
    type size is what gives. If even the floor needs two pages, the floor
    is what comes back — the chart is as small as it is allowed to get,
    and the export says two pages rather than pretending otherwise.
    """
    hi = min(MAX_FIT_SCALE,
             max(_width_limited_scale(doc, instruments, orient, columns),
                 MIN_ONE_PAGE_SCALE))
    if _build_pdf(doc, instruments, orient, hi, columns)[1] <= 1:
        return hi
    lo = MIN_ONE_PAGE_SCALE
    if _build_pdf(doc, instruments, orient, lo, columns)[1] > 1:
        return lo
    for _ in range(_FIT_ITERATIONS):
        mid = (lo + hi) / 2
        if _build_pdf(doc, instruments, orient, mid, columns)[1] <= 1:
            lo = mid
        else:
            hi = mid
    return lo


def resolve_scale(doc: dict, instruments=None, orient: str = "portrait",
                   columns=None) -> float:
    """The scale a document's `pdf_scale` setting asks for. Unknown or
    missing values fall back to fit, which is never worse than 100%.

    Scale and column count are decided together — half a page's width
    fits a different size of type than a whole one — so `columns` is
    resolved from the document unless the caller has already done it.
    """
    if columns is None:
        columns = resolve_columns(doc, instruments, orient)
    raw = doc.get("pdf_scale", "fit")
    if isinstance(raw, (int, float)):
        return max(0.5, float(raw))
    raw = str(raw).strip().lower()
    if raw in ("", "fit", "auto"):
        return _fit_scale(doc, instruments, orient, columns)
    if raw in ("one", "one page", "onepage"):
        return _one_page_scale(doc, instruments, orient, columns)
    if raw in ("normal", "100%"):
        return 1.0
    try:
        return max(0.5, float(raw.rstrip("%")) / (100 if raw.endswith("%") else 1))
    except ValueError:
        return _fit_scale(doc, instruments, orient, columns)


# ==============================================================================
#  Columns
#
#  A chart is mostly short lines — a handful of chord symbols, a bar of
#  tab — so one column down an A4 page leaves half the sheet white and
#  caps how big "fit to page" can set the type. Printing two columns
#  halves the height the chart needs, which is height the fit can spend
#  on type size instead. It only pays when the lines are short enough to
#  live in half a page, which is what "auto" is for.
# ==============================================================================

# Two columns have to buy a visibly bigger chart to be worth the split.
_AUTO_COLUMN_GAIN = 1.05


def _auto_columns(doc: dict, instruments, orient: str) -> int:
    """Two columns when they earn their keep, one when they don't.

    Splitting the page is not free: each line has half the width to live
    in, so a chart with long lines comes out *smaller* across two
    columns, not bigger. So the two layouts are built and compared on
    what actually matters at a music stand, in this order: how many
    sheets you have to turn, then how big the type is. A page saved
    beats a point of type size — reaching for the next sheet mid-song
    costs more than a slightly smaller chord symbol.
    """
    if MAX_PDF_COLUMNS < 2 or not doc.get("sections"):
        return 1
    one = resolve_scale(doc, instruments, orient, columns=1)
    two = resolve_scale(doc, instruments, orient, columns=2)

    # A column narrower than the chart is not a layout, it's an overflow.
    if two > _width_limited_scale(doc, instruments, orient, 2) + 1e-9:
        return 1
    if _tab_block_units(doc, instruments) * two > column_width(orient, 2):
        return 1

    pages_one = _build_pdf(doc, instruments, orient, one, 1)[1]
    pages_two = _build_pdf(doc, instruments, orient, two, 2)[1]
    if pages_two != pages_one:
        return 2 if pages_two < pages_one else 1
    return 2 if two >= one * _AUTO_COLUMN_GAIN else 1


def resolve_columns(doc: dict, instruments=None, orient: str = "portrait") -> int:
    """How many text columns a document's `pdf_columns` setting asks for.
    Unknown or missing values fall back to auto."""
    raw = doc.get("pdf_columns", "auto")
    if isinstance(raw, (int, float)):
        return max(1, min(MAX_PDF_COLUMNS, int(raw)))
    raw = str(raw).strip().lower()
    if raw in ("", "auto", "fit"):
        return _auto_columns(doc, instruments, orient)
    if raw in ("1", "one", "single"):
        return 1
    if raw in ("2", "two", "double"):
        return min(2, MAX_PDF_COLUMNS)
    try:
        return max(1, min(MAX_PDF_COLUMNS, int(float(raw))))
    except ValueError:
        return _auto_columns(doc, instruments, orient)


def build_pdf(doc: dict, instruments=None, orient: str = "portrait") -> bytes:
    """Render `doc` to PDF bytes at whatever scale its settings ask for."""
    if instruments is None:
        instruments = _all_instruments(doc)
    columns = resolve_columns(doc, instruments, orient)
    scale = resolve_scale(doc, instruments, orient, columns=columns)
    return _build_pdf(doc, instruments, orient, scale, columns)[0]


def _gutter_label(sec: dict) -> str:
    """'Verse 1 (x2)' — the section's name for the left-hand column."""
    rep = sec.get("repeat", 1)
    return f"{sec.get('name', '')}{f' (x{rep})' if rep and rep != 1 else ''}"


def build_song_lines(doc: dict, instruments=None) -> list[str]:
    """Render `doc` into the one-page TXT chart layout (design section 7)."""
    W = DEFAULT_TXT_WIDTH
    div = lambda c="=": c * W
    lines = []

    meta = _meta(doc)
    artist = (meta.get("artist") or "").strip()
    title = (meta.get("title") or "").strip()
    lines += [div("="), f"  {title.upper()}"]
    if artist:
        lines.append(f"  {artist}")
    meta_parts = []
    if meta.get("key"):
        meta_parts.append(f"Key: {meta['key']}")
    if meta.get("bpm"):
        meta_parts.append(f"BPM: {meta['bpm']}")
    if meta.get("time"):
        meta_parts.append(f"Time: {meta['time']}")
    one_instrument = uniform_instrument(doc, instruments)
    if one_instrument:
        meta_parts.append(one_instrument)
    if meta_parts:
        lines.append("  " + "   ".join(meta_parts))
    lines += [div("="), ""]

    if doc.get("print_lyrics") and (doc.get("lyrics_text") or "").strip():
        lines += [div("-"), "  LYRICS", div("-"), ""]
        lines += [f"  {ln}" for ln in doc["lyrics_text"].splitlines()]
        lines.append("")

    TOKEN_W = 3
    # "gutter" puts each section's name in a left-hand column beside its
    # first line instead of a full-width banner above it — three lines
    # saved per section, which is the difference between a one-page chart
    # and a two-page one for most songs.
    gutter = doc.get("section_layout") == "gutter"
    gutter_w = 0
    if gutter:
        gutter_w = max((len(_gutter_label(s)) for s in doc.get("sections", [])),
                        default=0) + 2

    for sec in doc.get("sections", []):
        if instruments is not None and sec.get("instrument") not in instruments:
            continue

        label = ""
        if gutter:
            label = _gutter_label(sec)
        else:
            rep = sec.get("repeat", 1)
            rep_str = f"  (x{rep})" if rep and rep != 1 else ""
            instr = "" if one_instrument else f"   {sec.get('instrument', '')}"
            lines += [div("-"),
                      f"  {sec['name']}{rep_str}{instr}",
                      div("-"), ""]

        body_indent = gutter_w if gutter else render.BODY_INDENT

        if sec.get("render") == "free":
            free = sec.get("free_text", "") or ""
            if free.strip():
                free_lines = free.splitlines()
                lines.append(f"{label:<{body_indent}}" + free_lines[0])
                lines += [" " * body_indent + ln for ln in free_lines[1:]]
                lines.append("")
                label = ""

        eff = transpose.effective_transpose(
            doc.get("transpose", 0), sec.get("transpose", 0))
        chart = render.resolve_references(songmap.chart_items(sec), doc)
        if sec.get("render", "chart") in ("chart", "both") and chart:
            resolved = render.resolve_display_items(chart, eff)
            lines += render.chart_body_lines(resolved, label, body_indent)
            lines.append("")
            label = ""

        # A gutter section with nothing in it still needs its name printed.
        if label:
            lines += [label, ""]

        measures = songmap.measure_items(sec)
        if sec.get("render", "chart") in ("tab", "both") and measures:
            all_strings = INSTRUMENT_STRINGS.get(
                sec.get("instrument"), ["e", "B", "G", "D", "A", "E"])
            strings = render.active_strings(
                [m.get("strings", {}) for m in measures], all_strings) or all_strings
            max_beats = max((m.get("beats", TAB_BEATS_DEFAULT) for m in measures),
                             default=TAB_BEATS_DEFAULT)
            mpl = render.measures_per_line(len(measures), max_beats, width=W)

            for chunk_start in range(0, len(measures), mpl):
                chunk = list(range(chunk_start, min(chunk_start + mpl, len(measures))))
                sn_pad = max(len(st) for st in strings) + 3
                hdr = " " * sn_pad
                for m_idx in chunk:
                    beats = measures[m_idx].get("beats", TAB_BEATS_DEFAULT)
                    cell_w = TOKEN_W * beats + 1
                    hdr += f"{'M' + str(m_idx + 1):<{cell_w}}"
                lines.append(hdr)
                for st in strings:
                    row = f"{st}| "
                    for m_idx in chunk:
                        beats = measures[m_idx].get("beats", TAB_BEATS_DEFAULT)
                        raw = measures[m_idx].get("strings", {}).get(st, "")
                        tokens = [tok for tok in raw.split() if tok]
                        while len(tokens) < beats:
                            tokens.append("-")
                        tokens = tokens[:beats]
                        row += "".join(f"{tok:>{TOKEN_W}}" for tok in tokens) + "|"
                    lines.append(row)
                lines.append("")

        if sec.get("annotation"):
            lines.append(" " * body_indent + f'"{sec["annotation"]}"')
            lines.append("")

        if sec.get("print_lyrics") and (sec.get("lyrics_text") or "").strip():
            lines += [" " * body_indent + ln
                      for ln in sec["lyrics_text"].splitlines()]
            lines.append("")

    lines += [div("="),
              f"  Generated by Song Notation Tool v{APP_VERSION}",
              f"  {APP_URL}",
              div("=")]
    return lines


# ==============================================================================
#  PDF export — byte-level writer ported from QLC+ Swiss Knife, no external
#  libraries (keeps the project's "zero external dependencies" rule).
# ==============================================================================

def _build_pdf(doc: dict, instruments, orient: str, scale: float,
                columns=None):
    """Render the chart at `scale` and return (pdf_bytes, page_count).

    Everything that sets the size of the type — line height, font sizes,
    the monospace character width the column maths depends on — is
    multiplied by `scale`. The margins and the footer stay put, so scaling
    up genuinely fills the page rather than just inflating the whole sheet.

    `columns` splits the area between the margins into that many text
    columns, filled top to bottom and then left to right; None asks the
    document what it wants.
    """
    if instruments is None:
        instruments = _all_instruments(doc)
    if columns is None:
        columns = resolve_columns(doc, instruments, orient)

    W, H = page_size(orient)
    S = max(0.5, float(scale or 1.0))
    MARGIN, FOOTER_H = PAGE_MARGIN, 18
    NCOLS = max(1, min(MAX_PDF_COLUMNS, int(columns or 1)))
    COL_GAP = PDF_COLUMN_GAP
    COL_W_TEXT = column_width(orient, NCOLS)
    LINE_H = 12 * S
    MONO_SZ, HEAD_SZ, TITLE_SZ = MONO_SIZE * S, 10.5 * S, 14 * S
    TOKEN_W, CHAR_W, SN_W = 3, MONO_CHAR_W * S, STRING_LABEL_W * S

    colors = doc.get("color_mode", "color")
    def sec_rgb(sec):
        return section_color(sec.get("type", ""), colors)

    meta = _meta(doc)
    artist = (meta.get("artist") or "").strip()
    title = (meta.get("title") or "").strip()
    doc_date = datetime.date.today().strftime("%Y-%m-%d")

    pages = []
    cur_ln = []

    def _esc(s):
        s = str(s)
        for frm, to in [
            ("\u2014", "-"), ("\u2013", "-"), ("\u00d7", "x"), ("\u00d8", "x"),
            ("\u2019", "'"), ("\u2018", "'"), ("\u201c", '"'), ("\u201d", '"'),
            ("\u00e9", "e"), ("\u00e8", "e"), ("\u00e0", "a"), ("\u00f4", "o"),
            ("\u266a", ""), ("\u2665", ""), ("\u00d6", "O"), ("\u00fc", "u"),
            ("\u25b2", "^"), ("\u25bc", "v"), ("\u25c6", "*"), ("\u203a", ">"),
            ("\u00b7", "."),
        ]:
            s = s.replace(frm, to)
        s = s.encode("latin-1", errors="replace").decode("latin-1")
        return s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

    def txt(x, y, s, sz=MONO_SZ, bold=False):
        font = "/F2" if bold else "/F1"
        cur_ln.append(f"BT {font} {sz} Tf {x:.1f} {y:.1f} Td ({_esc(s)}) Tj ET")

    def color(r, g, b):
        cur_ln.append(f"{r:.3f} {g:.3f} {b:.3f} rg")

    def role_rgb(role):
        """Colour for a chart row, or for a stretch inside one.

        A lick is tab sitting in a line of chord symbols; printed in the
        same black it reads as more symbols until you look twice, so it
        gets its own colour. A rest is the opposite case — it is the
        absence of playing, and grey says that without it competing with
        the notes around it. Black-and-white mode keeps the distinction
        the only way a mono printer can: tone, not hue.
        """
        if role == render.ROLE_LICK:
            return (0.30, 0.30, 0.30) if colors == "bw" else LICK_RGB
        if role == render.ROLE_REST:
            return REST_RGB
        return (0, 0, 0)

    def draw_chart_row(x, y, row):
        """One rendered chart row, in the colours its roles ask for.

        The text is monospace, so a span's column index is an x offset:
        the row is drawn as the runs between its spans rather than drawn
        once and overpainted, which in PDF would leave both layers visible.
        """
        text, spans = row["text"], row.get("spans") or []
        if not spans:
            color(*role_rgb(row.get("role")))
            txt(x, y, text, sz=MONO_SZ)
            return
        base = role_rgb(row.get("role"))
        pos = 0
        for start, end, role in sorted(spans):
            start, end = max(start, pos), min(end, len(text))
            if end <= start:
                continue
            if start > pos:
                color(*base)
                txt(x + pos * CHAR_W, y, text[pos:start], sz=MONO_SZ)
            color(*role_rgb(role))
            txt(x + start * CHAR_W, y, text[start:end], sz=MONO_SZ)
            pos = end
        if pos < len(text):
            color(*base)
            txt(x + pos * CHAR_W, y, text[pos:], sz=MONO_SZ)

    def hline(x1, y, x2, width=0.25, gray=0.72):
        cur_ln.append(f"{gray:.2f} G {width} w {x1:.1f} {y:.1f} m {x2:.1f} {y:.1f} l S")

    def rfill(x, y, w, h, r, g, b):
        cur_ln.append(f"{r:.3f} {g:.3f} {b:.3f} rg {x:.1f} {y:.1f} {w:.1f} {h:.1f} re f")

    pn_holder = [1]
    cy_holder = [0.0]
    col_holder = [0]            # which column the writing head is in
    top_holder = [H - MARGIN]   # y the current page's columns start at

    def col_x0(i=None):
        """Left edge of a column — the x every body line is measured from."""
        i = col_holder[0] if i is None else i
        return MARGIN + i * (COL_W_TEXT + COL_GAP)

    def col_x1(i=None):
        return col_x0(i) + COL_W_TEXT

    def vline(x, y1, y2, width=0.4, gray=PDF_COLUMN_RULE_GRAY):
        cur_ln.append(f"{gray:.2f} G {width} w {x:.1f} {y1:.1f} m {x:.1f} {y2:.1f} l S")

    def finish_page():
        # A rule down each gutter. Two columns of chart with nothing
        # between them is two charts you have to guess the edges of;
        # the line says where one ends and the next begins, which is
        # the whole point of reading off a stand mid-song.
        for i in range(1, NCOLS):
            vline(col_x0(i) - COL_GAP / 2, FOOTER_H + 8, top_holder[0] + LINE_H * 0.8)
        hline(MARGIN, FOOTER_H, W - MARGIN)
        color(0.4, 0.4, 0.4)
        txt(MARGIN, 5,
            f"Song Notation Tool v{APP_VERSION}  -  {doc_date}  -  {APP_URL}",
            sz=6.5)
        txt(W - MARGIN - 28, 5, f"Page {pn_holder[0]}", sz=6.5)
        if cur_ln:
            pages.append(zlib.compress("\n".join(cur_ln).encode("latin-1")))
            cur_ln.clear()

    def next_column():
        """Move the writing head to the top of the next column, or to the
        next page when the last column on this one is full. Returns the
        y to carry on at."""
        if col_holder[0] + 1 < NCOLS:
            col_holder[0] += 1
        else:
            finish_page()
            pn_holder[0] += 1
            col_holder[0] = 0
            top_holder[0] = H - MARGIN
        return top_holder[0]

    sections = doc.get("sections", [])
    max_beats = max(
        (m.get("beats", TAB_BEATS_DEFAULT)
         for s in sections if s.get("instrument") in instruments
         for m in songmap.measure_items(s)),
        default=TAB_BEATS_DEFAULT,
    )
    MEAS_W = TOKEN_W * max_beats * CHAR_W + CHAR_W

    def mpl_for(n_measures):
        usable = COL_W_TEXT - SN_W
        return max(1, min(n_measures, int(usable / MEAS_W)))

    band_h = 46 * S
    band_rgb, band_text = title_fill(colors)
    rfill(0, H - band_h, W, band_h, *band_rgb)
    color(*band_text)
    hdr_txt = title.upper() + (f"  -  {artist}" if artist else "")
    # The title is set in a proportional face, so it isn't covered by the
    # monospace width bound that sizes the body. A long title at a high
    # scale would simply run off the right edge — so it gets its own cap
    # and shrinks on its own rather than dragging the whole chart down
    # with it.
    txt(MARGIN, H - 28 * S, hdr_txt,
        sz=_fit_text_size(hdr_txt, W - 2 * MARGIN, TITLE_SZ), bold=True)
    meta_parts = []
    if meta.get("key"):
        meta_parts.append(f"Key: {meta['key']}")
    if meta.get("bpm"):
        meta_parts.append(f"BPM: {meta['bpm']}")
    if meta.get("time"):
        meta_parts.append(f"Time: {meta['time']}")
    one_instrument = uniform_instrument(doc, instruments)
    if one_instrument:
        meta_parts.append(one_instrument)
    if meta_parts:
        meta_line = "   |   ".join(meta_parts)
        txt(MARGIN, H - 41 * S, meta_line,
            sz=_fit_text_size(meta_line, W - 2 * MARGIN, 7.5 * S))
    # The banner spans the sheet, so both columns start below it.
    top_holder[0] = H - 54 * S
    cy_holder[0] = top_holder[0]

    if doc.get("print_lyrics") and (meta_lyrics := (doc.get("lyrics_text") or "")).strip():
        # Not bound by the "never split a section" rule below — this sits
        # before any section, so unlike section-level lyrics it's allowed
        # to run onto a second page if it's long.
        cy = cy_holder[0]
        color(0.16, 0.24, 0.42)
        txt(col_x0(), cy, "LYRICS", sz=HEAD_SZ, bold=True)
        cy -= LINE_H + 2 * S
        color(0.15, 0.15, 0.15)
        for ln in meta_lyrics.splitlines():
            if cy < FOOTER_H + LINE_H:
                cy = next_column()
            txt(col_x0(), cy, ln, sz=MONO_SZ)
            cy -= LINE_H
        cy -= LINE_H
        cy_holder[0] = cy

    gutter = doc.get("section_layout") == "gutter"
    gutter_w = 0
    if gutter:
        gutter_w = max((len(_gutter_label(s)) for s in sections), default=0) + 2
    def body_x():
        """Where a body line starts: the column's left edge, plus the
        gutter the section names sit in when the layout asks for one."""
        return col_x0() + gutter_w * CHAR_W

    for sec in sections:
        if sec.get("instrument") not in instruments:
            continue
        cy = cy_holder[0]

        all_strings = INSTRUMENT_STRINGS.get(
            sec.get("instrument"), ["e", "B", "G", "D", "A", "E"])
        eff = transpose.effective_transpose(
            doc.get("transpose", 0), sec.get("transpose", 0))
        # The section's render mode decides what gets printed, exactly as
        # it decides what the editor shows. A section left in Chart mode
        # can still hold tab measures from before the switch — the editor
        # hides that grid, and printing it anyway put a tab block on the
        # page that the card gave no sign of.
        mode = sec.get("render", "chart")
        free_mode = mode == "free"
        chart = render.resolve_references(songmap.chart_items(sec), doc)
        chart_rows = (render.chart_body_rows(
                          render.resolve_display_items(chart, eff), "", 0)
                      if chart and mode in ("chart", "both") else [])
        measures = songmap.measure_items(sec) if mode in ("tab", "both") else []
        strings = (render.active_strings(
                    [m.get("strings", {}) for m in measures], all_strings)
                   if measures else [])
        mpl = mpl_for(len(measures)) if measures else 1

        # Acceptance criterion (design section 10): never split a section
        # across a page break.
        # ... and the estimate has to see the same document the drawing
        # does, or a section that is only "=Chorus_1" measures one line
        # and splits across the break anyway. The pad covers the vertical
        # space the heading band and the inter-section gap add.
        needed_h = (render.estimate_section_lines(sec, strings or all_strings, doc)
                    * LINE_H + 34 * S)
        if cy - needed_h < FOOTER_H + LINE_H * 2:
            cy = next_column()

        rep = sec.get("repeat", 1)
        rep_str = f" (x{rep})" if rep and rep != 1 else ""
        rgb = sec_rgb(sec)
        if gutter:
            # Name in a left column beside the section's first line rather
            # than a banner above it; body text shifts right to clear it.
            color(*rgb)
            txt(col_x0(), cy, _gutter_label(sec), sz=MONO_SZ, bold=True)
        else:
            instr = "" if one_instrument else f"   {sec.get('instrument', '')}"
            sec_label = f"{sec['name']}{rep_str}{instr}"
            fill_rgb, label_rgb = heading_fill(sec.get("type", ""), colors)
            rfill(col_x0(), cy - LINE_H, COL_W_TEXT, LINE_H + 2 * S, *fill_rgb)
            color(*label_rgb)
            txt(col_x0() + 4 * S, cy - LINE_H + 3 * S, sec_label,
                sz=HEAD_SZ, bold=True)
            cy -= LINE_H + 14 * S

        if free_mode and (sec.get("free_text") or "").strip():
            # The notes are always black. Colour is a label — it tells you
            # which section you're in; tinting the notes themselves just
            # makes them harder to read, which is the opposite of the point.
            color(0, 0, 0)
            for ln in sec["free_text"].splitlines():
                txt(body_x(), cy, ln, sz=MONO_SZ)
                cy -= LINE_H
                if cy < FOOTER_H + LINE_H:
                    cy = next_column()
            cy -= 2 * S

        if chart_rows:
            for row in chart_rows:
                draw_chart_row(body_x(), cy, row)
                cy -= LINE_H
            cy -= 2 * S

        if sec.get("annotation"):
            color(0.3, 0.3, 0.3)
            txt(body_x(), cy, f'"{sec["annotation"]}"', sz=MONO_SZ)
            cy -= LINE_H

        if sec.get("print_lyrics") and (sec.get("lyrics_text") or "").strip():
            color(0.15, 0.15, 0.15)
            for ln in sec["lyrics_text"].splitlines():
                txt(body_x(), cy, ln, sz=MONO_SZ)
                cy -= LINE_H
            cy -= 2 * S

        for bs in range(0, len(measures), mpl):
            batch = list(range(bs, min(bs + mpl, len(measures))))

            def col_x(i, batch=batch):
                return col_x0() + SN_W + sum(
                    TOKEN_W * measures[batch[j]].get("beats", TAB_BEATS_DEFAULT)
                    * CHAR_W + CHAR_W
                    for j in range(i))

            color(*_mix(rgb, 0.5))
            for i, m_idx in enumerate(batch):
                txt(col_x(i), cy, f"M{m_idx + 1}", sz=7 * S)
            cy -= LINE_H

            for st in strings:
                color(*_mix(rgb, 0.35))
                txt(col_x0(), cy, f"{st}|", sz=MONO_SZ)
                for i, m_idx in enumerate(batch):
                    beats = measures[m_idx].get("beats", TAB_BEATS_DEFAULT)
                    raw = measures[m_idx].get("strings", {}).get(st, "")
                    tokens = [tok for tok in raw.split() if tok]
                    while len(tokens) < beats:
                        tokens.append("-")
                    tokens = tokens[:beats]
                    row_str = "".join(f"{tok:>{TOKEN_W}}" for tok in tokens) + "|"
                    color(0, 0, 0)
                    txt(col_x(i), cy, row_str, sz=MONO_SZ)
                cy -= LINE_H
                if cy < FOOTER_H + LINE_H:
                    cy = next_column()

            hline(col_x0(), cy, col_x1(), gray=0.82)
            cy -= 3 * S

        cy -= 18 * S
        cy_holder[0] = cy

    finish_page()
    return _assemble_pdf(pages, W, H), len(pages)


def _assemble_pdf(pages, W, H) -> bytes:
    """Assemble raw PDF bytes from a list of zlib-compressed page streams.
    Ported directly from QLC+ Swiss Knife v0.4 — no external libraries."""
    raw = "%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"
    offsets = []

    def add(s):
        nonlocal raw
        offsets.append(len(raw))
        raw += s

    def obj(n, c):
        return f"{n} 0 obj\n{c}\nendobj\n"

    def sobj(n, data):
        body = data.decode("latin-1")
        return obj(n, (f"<< /Length {len(data)} /Filter /FlateDecode >>\n"
                       f"stream\n{body}\nendstream"))

    font_res = "<< /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R >> >>"
    kids, po, so = [], [], []
    cid = 6

    for ps in pages:
        kids.append(f"{cid} 0 R")
        po.append(obj(cid,
            f"<< /Type /Page /Parent 2 0 R "
            f"/MediaBox [0 0 {W:.2f} {H:.2f}] "
            f"/Contents {cid + 1} 0 R /Resources {font_res} >>"))
        so.append(sobj(cid + 1, ps))
        cid += 2

    add(obj(1, "<< /Type /Catalog /Pages 2 0 R >>"))
    add(obj(2, f"<< /Type /Pages /Kids [{' '.join(kids)}] /Count {len(pages)} >>"))
    add(obj(3, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica "
               "/Encoding /WinAnsiEncoding >>"))
    add(obj(4, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold "
               "/Encoding /WinAnsiEncoding >>"))
    add(obj(5, "<< /Type /Font /Subtype /Type1 /BaseFont /Courier "
               "/Encoding /WinAnsiEncoding >>"))
    for p, s in zip(po, so):
        add(p); add(s)

    n = cid - 1
    xoff = len(raw)
    raw += f"xref\n0 {n + 1}\n0000000000 65535 f \n"
    for o in offsets:
        raw += f"{o:010d} 00000 n \n"
    raw += f"trailer\n<< /Size {n + 1} /Root 1 0 R >>\nstartxref\n{xoff}\n%%EOF\n"
    return raw.encode("latin-1")

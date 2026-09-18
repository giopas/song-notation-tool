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
from constants import (APP_VERSION, INSTRUMENT_STRINGS, TAB_BEATS_DEFAULT,
                        section_color)

DEFAULT_TXT_WIDTH = 100


def _meta(doc):
    return doc.get("meta", {}) or {}


def _all_instruments(doc):
    return {s.get("instrument", "") for s in doc.get("sections", [])}


# ==============================================================================
#  TXT export
# ==============================================================================

def _mix(rgb, toward_grey: float):
    """`rgb` faded toward mid-grey — for secondary text (annotations, tab
    string labels) that should read as part of its section without
    competing with the notes themselves."""
    return tuple(c + (0.45 - c) * toward_grey for c in rgb)


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
_FIT_ITERATIONS = 10


def _body_lines_for_width(doc: dict, instruments) -> list[str]:
    """Every monospace line the PDF will draw, for measuring purposes.
    Tab grids are excluded: their column count adapts to the page width
    on its own (mpl_for), so they can't overflow it."""
    out = []
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
            out.append(f"[ {sec.get('name', '')} ]{rep_str}   {sec.get('instrument', '')}")
    return out


def _width_limited_scale(doc: dict, instruments, orient: str) -> float:
    W = 842 if orient == "landscape" else 595
    longest = max((len(ln) for ln in _body_lines_for_width(doc, instruments)), default=0)
    if longest <= 0:
        return MAX_FIT_SCALE
    return (W - 2 * 28) / (longest * 4.6)


def _fit_scale(doc: dict, instruments, orient: str) -> float:
    """Largest scale that keeps the chart on the same number of pages it
    needed at 100%, and inside the horizontal margins."""
    width_cap = _width_limited_scale(doc, instruments, orient)
    if width_cap < 1.0:
        # Too wide even at 100% — shrink to the width that fits.
        return max(MIN_FIT_SCALE, width_cap)

    _, base_pages = _build_pdf(doc, instruments, orient, 1.0)
    hi_cap = min(MAX_FIT_SCALE, width_cap)
    if hi_cap <= 1.0:
        return 1.0

    lo, hi = 1.0, hi_cap
    if _build_pdf(doc, instruments, orient, hi)[1] <= base_pages:
        return hi
    for _ in range(_FIT_ITERATIONS):
        mid = (lo + hi) / 2
        if _build_pdf(doc, instruments, orient, mid)[1] <= base_pages:
            lo = mid
        else:
            hi = mid
    return lo


def resolve_scale(doc: dict, instruments=None, orient: str = "portrait") -> float:
    """The scale a document's `pdf_scale` setting asks for. Unknown or
    missing values fall back to fit, which is never worse than 100%."""
    raw = doc.get("pdf_scale", "fit")
    if isinstance(raw, (int, float)):
        return max(0.5, float(raw))
    raw = str(raw).strip().lower()
    if raw in ("", "fit", "auto"):
        return _fit_scale(doc, instruments, orient)
    if raw in ("normal", "100%"):
        return 1.0
    try:
        return max(0.5, float(raw.rstrip("%")) / (100 if raw.endswith("%") else 1))
    except ValueError:
        return _fit_scale(doc, instruments, orient)


def build_pdf(doc: dict, instruments=None, orient: str = "portrait") -> bytes:
    """Render `doc` to PDF bytes at whatever scale its settings ask for."""
    if instruments is None:
        instruments = _all_instruments(doc)
    scale = resolve_scale(doc, instruments, orient)
    return _build_pdf(doc, instruments, orient, scale)[0]


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
            lines += [div("-"),
                      f"  [{sec['name']}]{rep_str}   {sec.get('instrument', '')}",
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
        if sec.get("render") != "free" and chart:
            resolved = render.resolve_display_items(chart, eff)
            lines += render.chart_body_lines(resolved, label, body_indent)
            lines.append("")
            label = ""

        # A gutter section with nothing in it still needs its name printed.
        if label:
            lines += [label, ""]

        measures = songmap.measure_items(sec)
        if sec.get("render") != "free" and measures:
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

    lines += [div("="), f"  Generated by Song Notation Tool v{APP_VERSION}", div("=")]
    return lines


# ==============================================================================
#  PDF export — byte-level writer ported from QLC+ Swiss Knife, no external
#  libraries (keeps the project's "zero external dependencies" rule).
# ==============================================================================

def _build_pdf(doc: dict, instruments, orient: str, scale: float):
    """Render the chart at `scale` and return (pdf_bytes, page_count).

    Everything that sets the size of the type — line height, font sizes,
    the monospace character width the column maths depends on — is
    multiplied by `scale`. The margins and the footer stay put, so scaling
    up genuinely fills the page rather than just inflating the whole sheet.
    """
    if instruments is None:
        instruments = _all_instruments(doc)

    W, H = (842, 595) if orient == "landscape" else (595, 842)
    S = max(0.5, float(scale or 1.0))
    MARGIN, FOOTER_H = 28, 18
    LINE_H = 12 * S
    MONO_SZ, HEAD_SZ, TITLE_SZ = 7.5 * S, 9 * S, 13 * S
    TOKEN_W, CHAR_W, SN_W = 3, 4.6 * S, 20 * S

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

    def hline(x1, y, x2, width=0.25, gray=0.72):
        cur_ln.append(f"{gray:.2f} G {width} w {x1:.1f} {y:.1f} m {x2:.1f} {y:.1f} l S")

    def rfill(x, y, w, h, r, g, b):
        cur_ln.append(f"{r:.3f} {g:.3f} {b:.3f} rg {x:.1f} {y:.1f} {w:.1f} {h:.1f} re f")

    pn_holder = [1]
    cy_holder = [0.0]

    def finish_page():
        hline(MARGIN, FOOTER_H, W - MARGIN)
        color(0.4, 0.4, 0.4)
        txt(MARGIN, 5, f"Song Notation Tool v{APP_VERSION}  -  {doc_date}", sz=6.5)
        txt(W - MARGIN - 28, 5, f"Page {pn_holder[0]}", sz=6.5)
        if cur_ln:
            pages.append(zlib.compress("\n".join(cur_ln).encode("latin-1")))
            cur_ln.clear()

    sections = doc.get("sections", [])
    max_beats = max(
        (m.get("beats", TAB_BEATS_DEFAULT)
         for s in sections if s.get("instrument") in instruments
         for m in songmap.measure_items(s)),
        default=TAB_BEATS_DEFAULT,
    )
    COL_W = TOKEN_W * max_beats * CHAR_W + CHAR_W

    def mpl_for(n_measures):
        usable = W - 2 * MARGIN - SN_W
        return max(1, min(n_measures, int(usable / COL_W)))

    band_h = 46 * S
    band_rgb = (0.0, 0.0, 0.0) if colors == "bw" else (0.10, 0.12, 0.22)
    rfill(0, H - band_h, W, band_h, *band_rgb)
    color(1, 1, 1)
    hdr_txt = title.upper() + (f"  -  {artist}" if artist else "")
    txt(MARGIN, H - 28 * S, hdr_txt, sz=TITLE_SZ, bold=True)
    meta_parts = []
    if meta.get("key"):
        meta_parts.append(f"Key: {meta['key']}")
    if meta.get("bpm"):
        meta_parts.append(f"BPM: {meta['bpm']}")
    if meta.get("time"):
        meta_parts.append(f"Time: {meta['time']}")
    if meta_parts:
        txt(MARGIN, H - 41 * S, "   |   ".join(meta_parts), sz=7.5 * S)
    cy_holder[0] = H - 54 * S

    if doc.get("print_lyrics") and (meta_lyrics := (doc.get("lyrics_text") or "")).strip():
        # Not bound by the "never split a section" rule below — this sits
        # before any section, so unlike section-level lyrics it's allowed
        # to run onto a second page if it's long.
        cy = cy_holder[0]
        color(0.16, 0.24, 0.42)
        txt(MARGIN, cy, "LYRICS", sz=HEAD_SZ, bold=True)
        cy -= LINE_H + 2 * S
        color(0.15, 0.15, 0.15)
        for ln in meta_lyrics.splitlines():
            if cy < FOOTER_H + LINE_H:
                finish_page(); pn_holder[0] += 1; cy = H - MARGIN
            txt(MARGIN, cy, ln, sz=MONO_SZ)
            cy -= LINE_H
        cy -= LINE_H
        cy_holder[0] = cy

    gutter = doc.get("section_layout") == "gutter"
    gutter_w = 0
    if gutter:
        gutter_w = max((len(_gutter_label(s)) for s in sections), default=0) + 2
    BODY_X = MARGIN + gutter_w * CHAR_W

    for sec in sections:
        if sec.get("instrument") not in instruments:
            continue
        cy = cy_holder[0]

        all_strings = INSTRUMENT_STRINGS.get(
            sec.get("instrument"), ["e", "B", "G", "D", "A", "E"])
        eff = transpose.effective_transpose(
            doc.get("transpose", 0), sec.get("transpose", 0))
        free_mode = sec.get("render") == "free"
        chart = render.resolve_references(songmap.chart_items(sec), doc)
        chart_rows = (render.chart_body_lines(
                          render.resolve_display_items(chart, eff), "", 0)
                      if chart and not free_mode else [])
        measures = [] if free_mode else songmap.measure_items(sec)
        strings = (render.active_strings(
                    [m.get("strings", {}) for m in measures], all_strings)
                   if measures else [])
        mpl = mpl_for(len(measures)) if measures else 1

        # Acceptance criterion (design section 10): never split a section
        # across a page break.
        needed_h = render.estimate_section_lines(sec, strings or all_strings) * LINE_H
        if cy - needed_h < FOOTER_H + LINE_H * 2:
            finish_page(); pn_holder[0] += 1; cy = H - MARGIN

        rep = sec.get("repeat", 1)
        rep_str = f" (x{rep})" if rep and rep != 1 else ""
        rgb = sec_rgb(sec)
        if gutter:
            # Name in a left column beside the section's first line rather
            # than a banner above it; body text shifts right to clear it.
            color(*rgb)
            txt(MARGIN, cy, _gutter_label(sec), sz=MONO_SZ, bold=True)
        else:
            sec_label = f"[ {sec['name']} ]{rep_str}   {sec.get('instrument', '')}"
            rfill(MARGIN, cy - LINE_H, W - 2 * MARGIN, LINE_H + 2 * S, *rgb)
            color(1, 1, 1)
            txt(MARGIN + 4 * S, cy - LINE_H + 3 * S, sec_label,
                sz=HEAD_SZ, bold=True)
            cy -= LINE_H + 6 * S

        if free_mode and (sec.get("free_text") or "").strip():
            color(*rgb)
            for ln in sec["free_text"].splitlines():
                txt(BODY_X, cy, ln, sz=MONO_SZ)
                cy -= LINE_H
                if cy < FOOTER_H + LINE_H:
                    finish_page(); pn_holder[0] += 1; cy = H - MARGIN
            cy -= 2 * S

        if chart_rows:
            color(*rgb)
            for ln in chart_rows:
                txt(BODY_X, cy, ln, sz=MONO_SZ)
                cy -= LINE_H
            cy -= 2 * S

        if sec.get("annotation"):
            color(*_mix(rgb, 0.45))
            txt(BODY_X, cy, f'"{sec["annotation"]}"', sz=MONO_SZ)
            cy -= LINE_H

        if sec.get("print_lyrics") and (sec.get("lyrics_text") or "").strip():
            color(0.15, 0.15, 0.15)
            for ln in sec["lyrics_text"].splitlines():
                txt(BODY_X, cy, ln, sz=MONO_SZ)
                cy -= LINE_H
            cy -= 2 * S

        for bs in range(0, len(measures), mpl):
            batch = list(range(bs, min(bs + mpl, len(measures))))

            def col_x(i, batch=batch):
                return MARGIN + SN_W + sum(
                    TOKEN_W * measures[batch[j]].get("beats", TAB_BEATS_DEFAULT)
                    * CHAR_W + CHAR_W
                    for j in range(i))

            color(*_mix(rgb, 0.45))
            for i, m_idx in enumerate(batch):
                txt(col_x(i), cy, f"M{m_idx + 1}", sz=7 * S)
            cy -= LINE_H

            for st in strings:
                color(*_mix(rgb, 0.3))
                txt(MARGIN, cy, f"{st}|", sz=MONO_SZ)
                for i, m_idx in enumerate(batch):
                    beats = measures[m_idx].get("beats", TAB_BEATS_DEFAULT)
                    raw = measures[m_idx].get("strings", {}).get(st, "")
                    tokens = [tok for tok in raw.split() if tok]
                    while len(tokens) < beats:
                        tokens.append("-")
                    tokens = tokens[:beats]
                    row_str = "".join(f"{tok:>{TOKEN_W}}" for tok in tokens) + "|"
                    color(*rgb)
                    txt(col_x(i), cy, row_str, sz=MONO_SZ)
                cy -= LINE_H
                if cy < FOOTER_H + LINE_H:
                    finish_page(); pn_holder[0] += 1; cy = H - MARGIN

            hline(MARGIN, cy, W - MARGIN, gray=0.82)
            cy -= 3 * S

        cy -= 8 * S
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

"""
tabtext.py — tab written as text, read into tab grid measures.

Two kinds of older song keep their tab as text rather than in the tab grid
(the measures you can lengthen, repeat and name in the browser app):

* a Free section (or part) holding tab typed by hand, the usual way:

      G|-----5---7-|---------|
      D|---7-------|-5-------|

* a chart line holding nothing but licks and rests,
  `rest {Riff1 = G 5 7 5 | D - - 3} {Riff2 = ...}`, which is what the
  **{ tab }** button gave before the grid could grow. Each named lick keeps
  its name on its measures, so `{Riff1}` elsewhere still plays it.

`upgrade_tab_layout()` turns both into tab grid measures when a song is
opened. Nothing is lost: a Free section keeps its text (switch it back to
Free and it is there), and a lick section is only converted when every
item in it is a lick the grid can hold.

Pure functions, no user interface: see DEVELOPMENT.md, "Pure modules
first".
"""

from __future__ import annotations

import re

from constants import INSTRUMENT_STRINGS

# The version of the tab layout a song was last upgraded to. A song that
# carries it is never upgraded again, so a section switched back to Free
# or Chart on purpose stays as it was left.
# 3 (14 November): rests and several named licks on one line convert too;
# a song opened with 2 gets the new pass once.
TAB_LAYOUT_VERSION = 3

MAX_BEATS = 64          # the browser grid's longest measure
_SPLIT_BEATS = 16       # a longer run of text is cut into measures this long

# A text tab row: a string name, then a bar or a colon, then the tab.
_ROW_RE = re.compile(r'^\s*([A-Ga-g][#b]?)\s*([|:])(.*)$')


def _strings_for(instrument: str):
    return list(INSTRUMENT_STRINGS.get(instrument, ["e", "B", "G", "D", "A", "E"]))


def is_tab_row(line: str) -> bool:
    """True for a line of text tab: "G|--5--7--|", "e:---3---".
    It needs a run of dashes, so "G: play it soft" is not one."""
    m = _ROW_RE.match(line or "")
    if not m:
        return False
    body = m.group(3)
    if "---" not in body and not re.search(r'-\d|\d-', body):
        return False
    return bool(re.fullmatch(r'[-0-9xXhpbr/\\~^|:.()* ]*', body))


def _map_strings(names, instrument):
    """Which grid string each text row is: by position when the text has
    as many rows as the instrument has strings (guitar tab often names
    both E strings "E"), otherwise by name. None when they do not fit."""
    strings = _strings_for(instrument)
    if len(names) == len(strings):
        return list(strings)
    out = []
    for n in names:
        hit = [s for s in strings if s == n] or \
              [s for s in strings if s.lower() == n.lower()]
        if len(hit) != 1 or hit[0] in out:
            return None
        out.append(hit[0])
    return out


def _measure(beats, strings, cells, bars=None):
    """A measure item from {string: [token, ...]}; every row padded."""
    out = {"kind": "measure", "beats": beats, "strings": {}}
    for st in strings:
        toks = list(cells.get(st, []))[:beats]
        while len(toks) < beats:
            toks.append("-")
        out["strings"][st] = " ".join(toks)
    if bars:
        out["bars"] = sorted(set(bars))
    return out


def _quantise(segment_rows, width):
    """The beats of one text measure. Every note starts a beat; the gaps
    between them become empty beats in proportion, measured by the
    narrowest gap, so "--5---7--" keeps its feel of space."""
    starts = {}
    for st, text in segment_rows.items():
        for m in re.finditer(r'[^-| ]+', text):
            starts.setdefault(m.start(), {})[st] = m.group(0)
    if not starts:
        return max(1, min(MAX_BEATS, round(width / 3) or 1)), {}
    cols = sorted(starts)
    gaps = [b - a for a, b in zip(cols, cols[1:])]
    # Hand-typed tab gives a beat two or three characters; a wider gap is
    # rest, not a slower beat.
    unit = max(1, min(min(gaps), 3)) if gaps else 3
    # A dash or two before the first note is the usual padding, not a beat.
    base = cols[0] if cols[0] <= unit else 0
    index = {c: (c - base + unit // 2) // unit for c in cols}
    last = index[cols[-1]]
    tail = max(0, (width - cols[-1]) // unit - 1)
    beats = last + 1 + tail
    cells = {}
    for c in cols:
        for st, tok in starts[c].items():
            row = cells.setdefault(st, [])
            while len(row) < index[c]:
                row.append("-")
            row.append(tok)
    return beats, cells


def _split_long(beats, cells, strings):
    """A run longer than the grid allows, cut into measures."""
    if beats <= MAX_BEATS:
        return [_measure(beats, strings, cells)]
    out = []
    for start in range(0, beats, _SPLIT_BEATS):
        n = min(_SPLIT_BEATS, beats - start)
        part = {st: [t for t in (cells.get(st, []) + ["-"] * beats)[start:start + n]]
                for st in strings}
        out.append(_measure(n, strings, part))
    return out


def parse_text_tab(text: str, instrument: str):
    """Text tab as measures, or None when `text` is not text tab.

    Returns (measures, other_lines): `other_lines` are the lines that were
    not tab (a riff's name, "x4"), in order, for the caller to keep."""
    lines = (text or "").splitlines()
    systems, other, current = [], [], []
    for ln in lines:
        if is_tab_row(ln):
            current.append(ln)
            continue
        if current:
            systems.append(current)
            current = []
        if ln.strip():
            other.append(ln.strip())
    if current:
        systems.append(current)
    if not systems:
        return None

    measures = []
    for rows in systems:
        names, bodies = [], []
        for ln in rows:
            m = _ROW_RE.match(ln)
            names.append(m.group(1))
            bodies.append(m.group(3).rstrip())
        strings = _map_strings(names, instrument)
        if strings is None:
            return None
        width = max(len(b) for b in bodies)
        bodies = [b.ljust(width, "-") for b in bodies]
        # A bar line is a column with "|" on the rows that reach it.
        bar_cols = [c for c in range(width)
                    if any(b[c] == "|" for b in bodies)]
        edges = [-1] + bar_cols + [width]
        for a, b in zip(edges, edges[1:]):
            seg = {st: body[a + 1:b] for st, body in zip(strings, bodies)}
            if not any(s.strip("-: ") for s in seg.values()) and \
                    all(len(s.strip(": ")) < 2 for s in seg.values()):
                continue            # "||" or the space before the first bar
            opens = all(s.startswith(":") for s in seg.values() if s)
            closes = all(s.endswith(":") for s in seg.values() if s)
            seg = {st: s.strip(":") for st, s in seg.items()}
            seg_w = max(len(s) for s in seg.values())
            beats, cells = _quantise(seg, seg_w)
            new = _split_long(beats, cells, strings)
            if opens:
                new[0]["repeat_open"] = True
            if closes:
                new[-1]["repeat_close"] = 2
            measures += new
    if not measures:
        return None
    # "x4" on a line of its own, after a repeat: how many times it plays.
    times = [int(m.group(1)) for ln in other
             for m in [re.fullmatch(r'\(?[xX×]\s*(\d+)\)?', ln)] if m]
    if times:
        other = [ln for ln in other if not re.fullmatch(r'\(?[xX×]\s*(\d+)\)?', ln)]
        closing = [m for m in measures if m.get("repeat_close")]
        if closing:
            closing[-1]["repeat_close"] = max(2, times[0])
        else:
            measures[0]["repeat_open"] = True
            measures[-1]["repeat_close"] = max(2, times[0])
    return measures, other


def _lick_measures(lick: dict, instrument: str):
    """One lick as a measure (or several when it is very long), or None
    when its strings do not belong to the instrument."""
    lines = lick.get("lines") or []
    if not lines:
        return None
    strings = _strings_for(instrument)
    names = [str(ln.get("string", "")) for ln in lines]
    mapped = []
    for n in names:
        hit = [s for s in strings if s == n] or \
              [s for s in strings if s.lower() == n.lower()]
        if len(hit) != 1 or hit[0] in mapped:
            # Guitar tab with both E strings written "E": by position.
            if len(names) == len(strings):
                mapped = list(strings)
                break
            return None
        mapped.append(hit[0])
    beats = max(len(ln.get("frets") or []) for ln in lines) or 1
    cells = {st: [str(f) for f in ln.get("frets") or []]
             for st, ln in zip(mapped, lines)}
    return _split_long(beats, cells, strings)


def _licks_as_measures(items, instrument, rest_beats=4):
    """Measures for a chart line that holds only licks (alone, repeated, or
    in a repeated group) and rests, or None when something else is on the
    line. Each named lick keeps its name on its own measures (`name` and
    `span`), so `{Riff1}` elsewhere still plays just that riff; a rest
    becomes an empty measure of `rest_beats` beats."""
    measures = []

    def add(item):
        new = _lick_measures(item, instrument)
        if new is None:
            return False
        if item.get("name"):
            new[0]["name"] = item["name"]
            if len(new) > 1:
                new[0]["span"] = len(new)
        n = int(item.get("repeat", 1) or 1)
        if n >= 2:
            new[0]["repeat_open"] = True
            new[-1]["repeat_close"] = n
        measures.extend(new)
        return True

    def rest():
        measures.append(_measure(rest_beats, _strings_for(instrument), {}))

    def skip(x):
        return x.get("kind") == "mark" and x.get("mark") == "line_break"

    for it in items or []:
        kind = it.get("kind")
        if skip(it):
            continue
        if kind == "mark" and it.get("mark") == "rest":
            rest()
            continue
        if kind == "lick":
            if not add(it):
                return None
            continue
        if kind == "group" and not it.get("span"):
            inner = [x for x in it.get("items", []) if not skip(x)]
            if not inner or not all(
                    x.get("kind") == "lick"
                    or (x.get("kind") == "mark" and x.get("mark") == "rest")
                    for x in inner):
                return None
            start = len(measures)
            for x in inner:
                if x.get("kind") == "lick":
                    if not add(x):
                        return None
                else:
                    rest()
            rep = int(it.get("repeat", 1) or 1)
            if rep >= 2:
                measures[start]["repeat_open"] = True
                if measure_close(measures[-1]) and len(inner) == 1:
                    measures[-1]["repeat_close"] = measure_close(measures[-1]) * rep
                else:
                    measures[-1]["repeat_close"] = rep
            continue
        return None
    if not any(m["strings"] and any(t != "-" for v in m["strings"].values()
                                    for t in v.split()) for m in measures):
        return None             # only rests: nothing to put in a grid
    return measures


def measure_close(m: dict) -> int:
    try:
        return int(m.get("repeat_close") or 0)
    except (TypeError, ValueError):
        return 0


def _upgrade_owner(owner: dict, instrument: str, rest_beats: int = 4) -> str:
    """Upgrade one section or part in place. Returns what happened:
    "text", "licks" or ""."""
    items = owner.get("items") or []
    has_measures = any(i.get("kind") == "measure" for i in items)
    if has_measures:
        return ""
    mode = owner.get("render", "chart")
    if mode == "free":
        parsed = parse_text_tab(owner.get("free_text") or "", instrument)
        if not parsed:
            return ""
        measures, other = parsed
        # The text stays in free_text: switching back to Free shows it.
        owner["items"] = [i for i in items if i.get("kind") != "measure"] + measures
        owner["render"] = "tab"
        if other and len(other) <= 2 and not owner.get("tab_name") and \
                _good_name(other[0]):
            owner["tab_name"] = other.pop(0)
        if other and "annotation" in owner and not owner.get("annotation"):
            owner["annotation"] = " · ".join(other)
        return "text"
    if mode in ("chart", "both"):
        chart = [i for i in items if i.get("kind") != "measure"]
        measures = _licks_as_measures(chart, instrument, rest_beats)
        if not measures:
            return ""
        owner["items"] = measures
        owner["render"] = "tab"
        return "licks"
    return ""


def _good_name(text: str) -> bool:
    from model import validate_block_name
    return bool(re.fullmatch(r'[A-Za-z][A-Za-z0-9_]*', text or "")) and \
        validate_block_name(text)


def upgrade_tab_layout(doc: dict) -> list:
    """Move tab kept as text into the tab grid, once per song.

    Changes `doc` in place and returns a list of what changed, one
    {"section", "part", "from"} each ("from" is "text" or "licks"), empty
    when nothing did. A song already upgraded is left alone."""
    if int(doc.get("tab_layout", 0) or 0) >= TAB_LAYOUT_VERSION:
        return []
    changed = []
    # A rest is a bar: as many beats as the time signature's top number.
    try:
        rest_beats = int(str((doc.get("meta") or {}).get("time") or "4/4").split("/")[0])
    except ValueError:
        rest_beats = 4
    rest_beats = min(max(rest_beats, 1), MAX_BEATS)
    for sec in doc.get("sections", []) or []:
        how = _upgrade_owner(sec, sec.get("instrument"), rest_beats)
        if how:
            changed.append({"section": sec.get("name", ""), "part": "", "from": how})
        for part in sec.get("parts") or []:
            how = _upgrade_owner(part, part.get("instrument") or sec.get("instrument"),
                                 rest_beats)
            if how:
                changed.append({"section": sec.get("name", ""),
                                "part": part.get("instrument", ""), "from": how})
    doc["tab_layout"] = TAB_LAYOUT_VERSION
    return changed

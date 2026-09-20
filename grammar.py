"""
grammar.py — Song Notation Tool v0.16 chart-line grammar.

Parses one typed line (the "chart row") into a list of items per the
model.py schema, and regenerates the line from items (the unparser).
Pure functions, fully unit-testable. See DESIGN_v0_16.md section 4.

    parse(line)            -> (items, annotation_or_None)
    parse_items(line)      -> items          (annotation dropped)
    unparse(items)         -> line
"""

from __future__ import annotations
import re

from model import (make_token, make_group, make_block_ref, make_section_ref,
                    make_mark, make_lick, make_lick_ref, validate_block_name)

MAX_FRET = 24


class ParseError(ValueError):
    pass


# A quote is a quote. macOS (and every word processor) turns a typed " into
# a curly one, and a line pasted from anywhere is likely to carry them — so
# accept all four rather than rejecting the line with a message that points
# at the text instead of at the character.
_QUOTES = '"\u201c\u201d\u2018\u2019'


def _strip_quotes(word: str) -> str:
    return word.strip(_QUOTES)


# ==============================================================================
#  Lexer — whitespace-separated, but bracket- and quote-aware.
# ==============================================================================

def _lex(line: str):
    tokens = []
    i, n = 0, len(line)
    while i < n:
        c = line[i]
        if c.isspace():
            i += 1
            continue
        if c in _QUOTES:
            j = i + 1
            while j < n and line[j] not in _QUOTES:
                j += 1
            end = min(j + 1, n)
            tokens.append(line[i:end])
            i = end
            continue
        if c == '{':
            j = i + 1
            while j < n and line[j] != '}':
                j += 1
            if j >= n:
                raise ParseError(f"unmatched '{{' in: {line!r}")
            # consume any suffix glued to the closing brace ("}x3")
            k = j + 1
            while k < n and not line[k].isspace():
                k += 1
            tokens.append(line[i:k])
            i = k
            continue
        if c == '[':
            depth = 1
            j = i + 1
            while j < n and depth > 0:
                if line[j] == '[':
                    depth += 1
                elif line[j] == ']':
                    depth -= 1
                j += 1
            if depth != 0:
                raise ParseError(f"unmatched '[' in: {line!r}")
            # consume any suffix directly glued to the closing bracket
            k = j
            while k < n and not line[k].isspace():
                k += 1
            tokens.append(line[i:k])
            i = k
            continue
        j = i
        while j < n and not line[j].isspace():
            j += 1
        tokens.append(line[i:j])
        i = j
    return tokens


# ==============================================================================
#  Marks
# ==============================================================================

_LITERAL_MARKS = {
    "|:": "repeat_open",
    ":|": "repeat_close",
    "%": "simile",
    # A line break. "//" rather than a bare "/" because a slash already
    # lives inside chord symbols (C/G) — as a whole word it's unambiguous,
    # but doubling it keeps it obvious at a glance which is meant.
    "//": "line_break",
}
_WORD_MARKS = {
    "coda": "coda", "segno": "segno", "dc": "dc", "ds": "ds",
    "rest": "rest",
}
_ENDING_RE = re.compile(r'^\|(\d)\.$')
# "//" breaks to a new line; each ">" after it pushes that line in one
# step, so a phrase can sit visibly inside the one above it.
_BREAK_RE = re.compile(r'^//(>*)$')


def _match_mark(word: str):
    if word in _LITERAL_MARKS:
        return _LITERAL_MARKS[word]
    m = _ENDING_RE.match(word)
    if m:
        n = m.group(1)
        if n == "1":
            return "ending_1"
        if n == "2":
            return "ending_2"
        return None
    if word.lower() in _WORD_MARKS:
        return _WORD_MARKS[word.lower()]
    return None


_MARK_TO_TEXT = {v: k for k, v in _LITERAL_MARKS.items()}
_MARK_TO_TEXT["ending_1"] = "|1."
_MARK_TO_TEXT["ending_2"] = "|2."
_MARK_TO_TEXT.update({v: k for k, v in _WORD_MARKS.items()})


# ==============================================================================
#  Standalone repeat / all / shift markers (merge onto the previous item)
# ==============================================================================

_STANDALONE_REPEAT_RE = re.compile(r'^[xX](\d+)$')
_STANDALONE_ALL_RE = re.compile(r'^all$', re.IGNORECASE)
_STANDALONE_SHIFT_RE = re.compile(r'^([+-]\d+)$')


def _apply_repeat(items, n):
    if not items:
        raise ParseError(f"'x{n}' has no preceding item to repeat")
    last = items[-1]
    if last["kind"] not in ("group", "block_ref", "section_ref",
                            "lick", "lick_ref"):
        raise ParseError(f"repeat is not valid on a {last['kind']} item")
    last["repeat"] = n


def _apply_all(items):
    if not items or items[-1]["kind"] != "section_ref":
        raise ParseError("'all' is only valid after a section reference")
    items[-1]["all"] = True


def _apply_shift(items, n):
    if not items or items[-1]["kind"] not in ("block_ref", "section_ref"):
        raise ParseError("a +/- shift is only valid after a block or section reference")
    items[-1]["transpose"] = n


# ==============================================================================
#  Token / block-ref word classification
# ==============================================================================

_TOKEN_RE = re.compile(
    r'^(?P<fret>0|[1-9][0-9]?)?(?P<note>[A-G])(?P<acc>[#b])?(?P<qual>[A-Za-z0-9/+\-()]*)$'
)
_IDENT_RE = re.compile(r'^[A-Za-z][A-Za-z0-9_]*$')
_LEADING_ZERO_RE = re.compile(r'^0\d')


def _parse_token_or_block(word: str):
    if _LEADING_ZERO_RE.match(word):
        raise ParseError(f"leading zero not allowed in fret: {word!r}")

    m = _TOKEN_RE.match(word)
    if m:
        fret = m.group("fret")
        fret_val = int(fret) if fret is not None else None
        if fret_val is not None and fret_val > MAX_FRET:
            raise ParseError(f"fret {fret_val} exceeds MAX_FRET ({MAX_FRET}): {word!r}")
        symbol = m.group("note") + (m.group("acc") or "") + (m.group("qual") or "")
        return make_token(symbol, fret_val)

    if _IDENT_RE.match(word):
        return make_block_ref(word)

    raise ParseError(f"cannot parse item: {word!r}")


def _consume_suffix(tail: str):
    """Parse an optional 'xN[all]' and/or '+/-N' glued directly onto a
    group's closing bracket, e.g. ']x2', ']x2+3'."""
    if tail == "":
        return 1, False, 0
    m = re.fullmatch(r'(?:[xX](\d+)(all)?)?([+-]\d+)?', tail)
    if not m:
        raise ParseError(f"cannot parse suffix: {tail!r}")
    repeat = int(m.group(1)) if m.group(1) else 1
    all_flag = bool(m.group(2))
    shift = int(m.group(3)) if m.group(3) else 0
    return repeat, all_flag, shift


def _find_matching_bracket(word: str) -> int:
    depth = 0
    for idx, ch in enumerate(word):
        if ch == '[':
            depth += 1
        elif ch == ']':
            depth -= 1
            if depth == 0:
                return idx
    raise ParseError(f"unmatched '[' in: {word!r}")


def _parse_group(word: str):
    close = _find_matching_bracket(word)
    inner = word[1:close]
    tail = word[close + 1:]
    repeat, all_flag, shift = _consume_suffix(tail)
    if all_flag or shift:
        raise ParseError("'all' and +/- shift are not valid on a group")
    return make_group(parse_items(inner), repeat=repeat)


# ==============================================================================
#  Licks — a short tab figure written inline, where it's played
#
#      {G 5 7 5 | D - - 3}
#
#  Each |-separated line is one string: its name, then its frets. Positions
#  align by index across the lines, and "-" means that string isn't played
#  there. An optional ":" after the string name reads more naturally for
#  some people, so both "G 5 7 5" and "G: 5 7 5" are accepted.
# ==============================================================================

_LICK_FRET_RE = re.compile(r'^(?:[-x]|0|[1-9][0-9]?)$', re.IGNORECASE)
# "{Riff1 = G 5 7 5 | D - - 3}" names the lick; "{Riff1}" recalls it.
_LICK_NAME_RE = re.compile(r'^\s*([A-Za-z][A-Za-z0-9_]*)\s*=\s*(.*)$', re.DOTALL)


def _parse_lick(word: str):
    inner = word[1:-1].strip()
    if not inner:
        raise ParseError("empty lick: {}")

    # A bare identifier inside the braces is a reference to a lick named
    # somewhere else — "{Riff1}". It can't be anything else: a lick line
    # needs a string name *and* frets, so "{Riff1}" was an error before.
    if _IDENT_RE.match(inner):
        return make_lick_ref(inner)

    name = ""
    m = _LICK_NAME_RE.match(inner)
    if m:
        name, inner = m.group(1), m.group(2).strip()
        if not validate_block_name(name):
            raise ParseError(
                f"{name!r} would read as a chord on a chart line — "
                f"name the lick something that doesn't start like a note")
        if not inner:
            raise ParseError(f"lick {name!r} has no tab after the '='")

    lines = []
    for raw in inner.split("|"):
        parts = raw.replace(":", " ").split()
        if not parts:
            continue
        st_name, frets = parts[0], parts[1:]
        if not st_name:
            raise ParseError(f"lick line has no string name: {raw.strip()!r}")
        if not frets:
            raise ParseError(f"lick line {st_name!r} has no frets")
        for f in frets:
            if not _LICK_FRET_RE.match(f):
                raise ParseError(
                    f"{f!r} is not a fret in lick line {st_name!r} "
                    f"(0-{MAX_FRET}, '-' for not played, or 'x' for muted)")
            if f not in ("-", "x", "X") and int(f) > MAX_FRET:
                raise ParseError(f"fret {f} exceeds MAX_FRET ({MAX_FRET}) in lick")
        lines.append({"string": st_name, "frets": frets})

    if not lines:
        raise ParseError("empty lick: {}")
    return make_lick(lines, name=name)


def _unparse_lick(item: dict) -> str:
    body = " | ".join(f"{ln['string']} " + " ".join(ln["frets"])
                      for ln in item.get("lines", []))
    name = item.get("name")
    if name:
        body = f"{name} = {body}"
    out = "{" + body + "}"
    if item.get("repeat", 1) != 1:
        out += f"x{item['repeat']}"
    return out


def _parse_section_ref(word: str):
    ident = word[1:]
    if not _IDENT_RE.match(ident):
        raise ParseError(f"invalid section reference: {word!r}")
    return make_section_ref(ident)


# ==============================================================================
#  Public API
# ==============================================================================

def parse_items(text: str):
    """Parse a chart line into a list of items. Quoted annotations are
    dropped (see parse() to retrieve them)."""
    items, _ = parse(text)
    return items


def parse(text: str):
    """
    Parse a chart line.
    Returns (items, annotation) where annotation is the text of the last
    quoted string on the line, or None if there wasn't one.
    """
    tokens = _lex(text or "")
    items = []
    annotation = None

    for tok in tokens:
        if _STANDALONE_REPEAT_RE.match(tok):
            _apply_repeat(items, int(_STANDALONE_REPEAT_RE.match(tok).group(1)))
            continue
        if _STANDALONE_ALL_RE.match(tok):
            _apply_all(items)
            continue
        if _STANDALONE_SHIFT_RE.match(tok):
            _apply_shift(items, int(tok))
            continue
        if tok[0] in _QUOTES:
            annotation = _strip_quotes(tok)
            continue
        if tok.startswith('{'):
            close = tok.rfind('}')
            if close < 0:
                raise ParseError(f"unmatched '{{' in: {tok!r}")
            item = _parse_lick(tok[:close + 1])
            repeat, all_flag, shift = _consume_suffix(tok[close + 1:])
            if all_flag or shift:
                raise ParseError("'all' and +/- shift are not valid on a lick")
            if repeat != 1:
                item["repeat"] = repeat
            items.append(item)
            continue
        brk = _BREAK_RE.match(tok)
        if brk:
            items.append(make_mark("line_break", indent=len(brk.group(1))))
            continue
        mk = _match_mark(tok)
        if mk:
            items.append(make_mark(mk))
            continue
        if tok.startswith('['):
            items.append(_parse_group(tok))
            continue
        if tok.startswith('='):
            items.append(_parse_section_ref(tok))
            continue
        items.append(_parse_token_or_block(tok))

    return items, annotation


def _unparse_item(it: dict) -> str:
    k = it["kind"]
    if k == "token":
        fret = it.get("fret")
        return (str(fret) if fret is not None else "") + it["symbol"]
    if k == "lick":
        return _unparse_lick(it)
    if k == "lick_ref":
        s = "{" + it["lick"] + "}"
        if it.get("repeat", 1) != 1:
            s += f"x{it['repeat']}"
        return s
    if k == "mark":
        if it["mark"] == "line_break":
            return "//" + ">" * int(it.get("indent", 0))
        return _MARK_TO_TEXT[it["mark"]]
    if k == "group":
        inner = " ".join(_unparse_item(x) for x in it["items"])
        s = f"[{inner}]"
        if it.get("repeat", 1) != 1:
            s += f"x{it['repeat']}"
        return s
    if k == "block_ref":
        parts = [it["block"]]
        if it.get("repeat", 1) != 1:
            parts.append(f"x{it['repeat']}")
        if it.get("transpose", 0):
            parts.append(f"{it['transpose']:+d}")
        return " ".join(parts)
    if k == "section_ref":
        parts = [f"={it['section']}"]
        if it.get("repeat", 1) != 1 or it.get("all"):
            parts.append(f"x{it.get('repeat', 1)}")
            if it.get("all"):
                parts.append("all")
        if it.get("transpose", 0):
            parts.append(f"{it['transpose']:+d}")
        return " ".join(parts)
    raise ValueError(f"{k!r} items have no chart-line representation")


def unparse(items) -> str:
    return " ".join(_unparse_item(it) for it in items)

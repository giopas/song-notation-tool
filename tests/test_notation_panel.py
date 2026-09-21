"""The in-app Notation reference only shows syntax that does what it says.

Its line-break examples were once written as "A(5) D(5)". That parses —
as a chord literally named "A(5)" — so nothing failed, but it printed
"A(5)" rather than fret 5 over A, teaching a form that silently doesn't
work. Every example in the reference table must parse, and none may
produce a symbol with brackets or a stray fret in its name.
"""
import html
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import grammar

INDEX = os.path.join(os.path.dirname(__file__), "..", "web", "index.html")


def _table_examples():
    page = open(INDEX, encoding="utf-8").read()
    table = page[page.index("<th>Example</th>"):]
    table = table[:table.index("</table>")]
    out = []
    for row in re.findall(r"<tr>(.*?)</tr>", table, re.S):
        cells = re.findall(r"<td>(.*?)</td>", row, re.S)
        if len(cells) == 3:
            out += [html.unescape(c) for c in re.findall(r"<code>(.*?)</code>", cells[2])]
    return out


def test_the_reference_has_examples():
    assert len(_table_examples()) > 15


def test_every_example_is_real_syntax():
    for ex in _table_examples():
        items = grammar.parse_items(ex)          # raises on bad syntax
        for it in items:
            sym = it.get("symbol") or ""
            assert "(" not in sym and ")" not in sym, f"{ex!r} makes a chord named {sym!r}"


def test_licks_can_be_named_and_recalled_in_the_reference():
    ex = _table_examples()
    assert any("=" in e and e.startswith("{") for e in ex)     # {Riff1 = …}
    assert "{Riff1}" in ex and "{Riff1}x3" in ex

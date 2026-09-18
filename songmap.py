"""
songmap.py — Song Notation Tool v0.17 song-map / riff-library helpers.

Pure functions only (no Tk dependency), so they're unit-testable the same
way model.py / grammar.py / render.py / transpose.py are. See
DESIGN_v0_17.md sections 3, 5, and 8.
"""

from __future__ import annotations

import re

import grammar
import model


# ==============================================================================
#  Riff library — usage tracking (section 5)
# ==============================================================================

def _walk_items(items):
    """Yield every item in a flat sequence, descending into group items
    (a block or section reference inside a group still counts as usage)."""
    for it in items or []:
        yield it
        if it.get("kind") == "group":
            yield from _walk_items(it.get("items", []))


def block_refs_in(items):
    """Every block id referenced anywhere in `items`, groups included."""
    return [it["block"] for it in _walk_items(items) if it.get("kind") == "block_ref"]


def block_usage(doc: dict) -> dict:
    """
    {block_id: [section_name, ...]} for every block referenced by at least
    one section. A block referenced twice within the same section is
    listed once, in section order. Used for the riff strip's usage badge
    and to block deletion of a referenced riff (section 5).
    """
    usage: dict = {}
    for sec in doc.get("sections", []):
        seen = set()
        for bid in block_refs_in(sec.get("items", [])):
            if bid not in seen:
                usage.setdefault(bid, []).append(sec.get("name") or sec.get("id", ""))
                seen.add(bid)
    return usage


def referencing_sections(doc: dict, block_id: str):
    return block_usage(doc).get(block_id, [])


def can_delete_block(doc: dict, block_id: str):
    """(True, []) if nothing references `block_id`; otherwise
    (False, [section_name, ...]) naming every referencing section."""
    refs = referencing_sections(doc, block_id)
    return (len(refs) == 0), refs


# ==============================================================================
#  Section-name / block-id references (renaming, deleting a section)
# ==============================================================================

def _same_section_key(doc: dict, ref_key, section_id) -> bool:
    if ref_key == section_id:
        return True
    target = find_section(doc, ref_key) if ref_key else None
    return bool(target) and target.get("id") == section_id


def section_refs_to(doc: dict, section_id: str):
    """Section names that hold a section_ref pointing at `section_id`."""
    out = []
    for sec in doc.get("sections", []):
        for it in _walk_items(sec.get("items", [])):
            if it.get("kind") == "section_ref" and _same_section_key(
                    doc, it.get("section"), section_id):
                out.append(sec.get("name") or sec.get("id", ""))
                break
    return out


def _ident(text: str) -> str:
    """Normalised key for matching a reference against a section name:
    case-folded, spaces and dashes as underscores."""
    return (text or "").strip().lower().replace(" ", "_").replace("-", "_")


def find_section(doc: dict, key: str):
    """
    The section a `=key` reference points at, or None.

    Ids are the stable identity (renaming a section never changes its id,
    so existing references keep working), but a reference may also be
    written with the section's *name* — `=Interlude` rather than the
    `=chorus1` that section happened to be created as. Id wins on a tie.
    """
    sections = doc.get("sections", [])
    for sec in sections:
        if sec.get("id") == key:
            return sec
    k = _ident(key)
    for sec in sections:
        if _ident(sec.get("name", "")) == k:
            return sec
    return None


_REF_IDENT_RE = re.compile(r'^[A-Za-z][A-Za-z0-9_]*$')


def display_ref_key(doc: dict, key: str) -> str:
    """
    How a `=key` reference should be *spelled on screen*.

    Ids are the truth — minted once, never changed, so a reference keeps
    working when its target is renamed. But an id is meaningless to read:
    a section created as "Chorus" and later renamed "Interlude" still has
    id `chorus1`, and `=chorus1` on a chart looks like a mistake. So the
    chart line displays the target's current name whenever that name can
    be written as a reference (spaces and dashes become underscores, which
    `find_section` matches back), and falls back to the id when it can't.

    Nothing is stored in this form: `canonicalise_refs` turns whatever was
    typed back into ids before it reaches the document.
    """
    target = find_section(doc, key)
    if target is None:
        return key
    candidate = (target.get("name") or "").strip().replace(" ", "_").replace("-", "_")
    if not _REF_IDENT_RE.match(candidate):
        return target.get("id") or key
    # Ambiguous names stay as ids — two sections called "Verse" would both
    # answer to "=Verse", and silently picking the first is worse than a
    # spelling nobody loves.
    matches = sum(1 for sec in doc.get("sections", [])
                  if _ident(sec.get("name", "")) == _ident(candidate))
    return candidate if matches == 1 else (target.get("id") or key)


def _map_section_refs(items, fn):
    """Copy `items`, rewriting every section_ref key with `fn`. Groups are
    walked recursively; every other item passes through untouched."""
    out = []
    for it in items or []:
        kind = it.get("kind")
        if kind == "section_ref":
            new_it = dict(it)
            new_it["section"] = fn(it.get("section", ""))
            out.append(new_it)
        elif kind == "group":
            new_it = dict(it)
            new_it["items"] = _map_section_refs(it.get("items", []), fn)
            out.append(new_it)
        else:
            out.append(it)
    return out


def display_items(items, doc: dict):
    """`items` with section_ref keys spelled for a human to read."""
    if not doc:
        return list(items or [])
    return _map_section_refs(items, lambda k: display_ref_key(doc, k))


def canonicalise_refs(items, doc: dict):
    """`items` with section_ref keys turned back into stable ids. A key
    that doesn't resolve is left exactly as typed, so a reference to a
    section that doesn't exist yet isn't silently rewritten."""
    if not doc:
        return list(items or [])

    def to_id(key):
        target = find_section(doc, key)
        return (target.get("id") or key) if target else key

    return _map_section_refs(items, to_id)


def section_display_name(doc: dict, key: str) -> str:
    """The name to show for a `=key` reference — the target's current name
    when it resolves, otherwise the key as written."""
    sec = find_section(doc, key)
    return (sec.get("name") or sec.get("id") or key) if sec else key


# ==============================================================================
#  Promote-to-riff (section 5, "Promote to riff" / Ctrl+R)
# ==============================================================================

def promote_to_riff(line: str, sel_start: int, sel_end: int, block_id: str):
    """
    Given a chart-line string and a [sel_start, sel_end) character range
    the user selected in the editor bar, parse the selected substring as
    its own item list (the new riff's content) and return
    (new_line, riff_items):

    - new_line   — `line` with the selection replaced by `block_id`
    - riff_items — the parsed items to store as the new block's `items`

    Raises grammar.ParseError if the selection doesn't parse on its own,
    or ValueError if the selection is empty/whitespace.
    """
    selected_text = line[sel_start:sel_end]
    if not selected_text.strip():
        raise ValueError("nothing selected")
    riff_items = grammar.parse_items(selected_text)
    if not riff_items:
        raise ValueError("selection did not produce any items")

    before, after = line[:sel_start], line[sel_end:]
    # Guard the replacement with whitespace on each side it borders
    # non-space text, so the new block id never fuses with a neighbouring
    # token (e.g. "...MainRiff15D" instead of "...MainRiff1 5D").
    lead = "" if (not before or before[-1].isspace()) else " "
    trail = "" if (not after or after[0].isspace()) else " "
    new_line = before + lead + block_id + trail + after
    return new_line, riff_items


def unique_block_id(doc: dict, stem: str = "riff") -> str:
    """First `{stem}N` not already used as a block id, N starting at 1."""
    existing = set(doc.get("blocks", {}).keys())
    n = 1
    while f"{stem}{n}" in existing:
        n += 1
    return f"{stem}{n}"


# ==============================================================================
#  Section reordering (section 3.2 — drag by the gutter, or Alt+Up/Alt+Down)
# ==============================================================================

def move_section(doc: dict, index: int, delta: int):
    """
    Move sections[index] by `delta` positions (delta=-1 up, +1 down),
    clamped to stay in range. Returns the item's new index (unchanged if
    already at the boundary it was pushed against). Mutates doc in place.
    """
    sections = doc.get("sections", [])
    if not (0 <= index < len(sections)):
        raise IndexError(index)
    new_index = max(0, min(len(sections) - 1, index + delta))
    if new_index == index:
        return index
    item = sections.pop(index)
    sections.insert(new_index, item)
    return new_index


def reorder_sections(doc: dict, old_index: int, new_index: int):
    """Move sections[old_index] to `new_index` (used by drag-reorder,
    where the drop target is a row rather than a +1/-1 step)."""
    sections = doc.get("sections", [])
    n = len(sections)
    if not (0 <= old_index < n):
        raise IndexError(old_index)
    new_index = max(0, min(n - 1, new_index))
    item = sections.pop(old_index)
    sections.insert(new_index, item)
    return new_index


# ==============================================================================
#  Chart items vs. measure items (section 4, last paragraph)
#
#  A section's `items` list can hold both chart-grammar items (token, group,
#  block_ref, section_ref, mark) and `measure` items from the v0.15 tab
#  grid. The chart editor bar only ever edits the former; the tab grid only
#  ever edits the latter. These helpers keep the two apart so committing an
#  edit in one surface never clobbers the other (render:"both" sections use
#  both surfaces on the same section).
# ==============================================================================

def chart_items(section: dict):
    return [it for it in section.get("items", []) if it.get("kind") != "measure"]


def measure_items(section: dict):
    return [it for it in section.get("items", []) if it.get("kind") == "measure"]


def set_chart_items(section: dict, new_chart_items):
    section["items"] = list(new_chart_items) + measure_items(section)


def set_measure_items(section: dict, new_measure_items):
    section["items"] = chart_items(section) + list(new_measure_items)


# ==============================================================================
#  Duplicate-as-reference (Ctrl+D — section 8: a reference, never a copy)
# ==============================================================================

def duplicate_as_reference(doc: dict, index: int, new_id: str, new_name: str):
    """
    Insert a new section right after sections[index] whose only content
    is a section_ref back at the original — "same as Verse 1" rather than
    a byte-for-byte copy. This is what removes the duplicate-page failure
    mode the v0.15 copy dialog produced (design section 8).
    """
    sections = doc.get("sections", [])
    if not (0 <= index < len(sections)):
        raise IndexError(index)
    source = sections[index]
    new_sec = model.new_section(
        section_id=new_id, name=new_name,
        section_type=source.get("type", "Verse"),
        instrument=source.get("instrument", "Bass (4-string)"),
        render=source.get("render", "chart"),
    )
    new_sec["items"] = [model.make_section_ref(source["id"])]
    sections.insert(index + 1, new_sec)
    return new_sec

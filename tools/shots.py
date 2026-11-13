#!/usr/bin/env python3
"""Retake the README screenshots in one command.

    python3 tools/shots.py                 # retake all of them into screenshots/
    python3 tools/shots.py --dry           # into a scratch folder; the repo is not touched
    python3 tools/shots.py --only editor section-copy
                                           # write only these (all are still taken)
    python3 tools/shots.py --list          # the pictures and what each one shows

It starts webserver.py on a free port with an empty scratch songs folder and
a scratch home folder, so your songs and your settings are never read or
changed, and opens the built-in example song. Every picture is taken at
1280 x 800 and saved at 1.5x (1920 x 1200), dark theme, as the browser shows
it (no window frame).

After a run it says, for each picture, whether it changed against the one in
the repo and by how much, and writes contact.png (every new picture, small)
so one image shows them all. Look at it before committing. "Nearly the
same" is a new version number or a one-pixel shift of the text in a field,
which can happen between two runs on the same computer.

Run it before every release: the release checklist in DEVELOPMENT.md asks
for it. Needs Playwright with Chromium (`pip install playwright`, then
`python3 -m playwright install chromium`); Pillow is optional and only used
for the comparison and the contact sheet. Neither is needed to run the app.
"""
from __future__ import annotations

import argparse
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SHOTS = os.path.join(ROOT, "screenshots")
WIDTH, HEIGHT, SCALE = 1280, 800, 1.5

# Name, then what it shows. The order is the order they are taken in.
PICTURES = [
    ("start-here.png", "no song open: New song, Open example, Import"),
    ("editor.png", "the example song in the editor"),
    ("section-copy.png", "the Copy to… list opened from a section's ⧉"),
    ("section-drag.png", "a section halfway through a drag, cards folded"),
    ("layout-panel.png", "⚙ Layout open"),
    ("lyrics-dialog.png", "the Lyrics dialog"),
    ("chord-shapes.png", "the Chords dialog after + From chart"),
    ("notation-reference.png", "the Notation ⌘ reference"),
    ("help-strip.png", "the How this works strip"),
    ("export-menu.png", "the Export menu open"),
    ("tab-grid.png", "a section in Tab: +4, bar lines, a repeat and a name"),
    ("collection.png", "a setlist in the song list, a song's ⋯ menu open"),
    ("delete-confirm.png", "the question before a song is deleted"),
]


# ── the app ──────────────────────────────────────────────────────────────────

def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


class App:
    """webserver.py on a free port, with scratch songs and home folders."""

    def __enter__(self):
        self.tmp = tempfile.mkdtemp(prefix="snt-shots-app-")
        songs = os.path.join(self.tmp, "songs")
        home = os.path.join(self.tmp, "home")
        os.makedirs(songs)
        os.makedirs(home)
        port = _free_port()
        self.url = f"http://127.0.0.1:{port}"
        env = dict(os.environ, HOME=home, USERPROFILE=home, APPDATA=home,
                   XDG_CONFIG_HOME=os.path.join(home, ".config"), SNT_NO_REEXEC="1")
        self.proc = subprocess.Popen(
            [sys.executable, os.path.join(ROOT, "webserver.py"), "--dir", songs,
             "--no-open", "--port", str(port)],
            cwd=ROOT, env=env, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
        for _ in range(100):
            try:
                urllib.request.urlopen(self.url + "/", timeout=1)
                return self
            except OSError:
                if self.proc.poll() is not None:
                    raise SystemExit("webserver.py stopped before it was ready")
                time.sleep(0.1)
        self.__exit__()
        raise SystemExit("webserver.py did not answer within 10 seconds")

    def __exit__(self, *_a):
        self.proc.terminate()
        try:
            self.proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            self.proc.kill()
        shutil.rmtree(self.tmp, ignore_errors=True)


# ── the pictures ─────────────────────────────────────────────────────────────

def take(url: str, out: str) -> list[str]:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise SystemExit("Playwright is missing: pip install playwright, "
                         "then python3 -m playwright install chromium")
    made = []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": WIDTH, "height": HEIGHT},
                                device_scale_factor=SCALE, color_scheme="dark")
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))

        def shot(name, wait=500, park=True):
            page.wait_for_timeout(wait)
            page.wait_for_load_state("networkidle")
            page.evaluate("document.fonts.ready.then(() => 0)")
            # The scratch folder's path means nothing to a reader: show the
            # default one. And no toast left over from the step before.
            page.evaluate("""() => {
                const p = document.querySelector('.songs-folder-path');
                if (p) p.textContent = '~/Documents/Song Notation Tool';
                document.getElementById('toast').classList.add('hidden');
                const a = document.activeElement;
                if (a && a.matches('input, select, textarea')) a.blur();
            }""")
            if park:                     # no hover highlight in the picture
                page.mouse.move(WIDTH - 5, HEIGHT - 5)
            path = os.path.join(out, name)
            page.screenshot(path=path)
            made.append(path)
            print("  ✓", name)

        def reopen():
            # Back to the editor with nothing open over it.
            page.goto(url + "/")
            page.wait_for_timeout(600)
            page.click("#song-list li >> nth=0")
            page.wait_for_selector(".section-card")
            # The rendered chart under each line comes from the server a
            # moment later; wait for it, or pictures differ run to run.
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(300)

        page.goto(url + "/")
        page.wait_for_selector("#start-here:not(.hidden)")
        shot("start-here.png")

        page.click("#start-example")
        page.wait_for_selector(".section-card")
        shot("editor.png", 800)

        # Copy the verse: the list opens with "after this one" highlighted.
        page.click(".section-card:nth-child(2) .sec-more")
        page.click(".section-card:nth-child(2) .sec-dup")
        shot("section-copy.png", 300)
        page.click(".copy-picker-row >> text=After 3.")    # a second verse after the chorus
        page.wait_for_timeout(1200)                        # let the flash fade

        # Hold the last section halfway up the list.
        handle = page.locator(".section-card:last-child .drag-handle")
        handle.scroll_into_view_if_needed()
        box = handle.bounding_box()
        x, y = box["x"] + box["width"] / 2, box["y"] + box["height"] / 2
        page.mouse.move(x, y)
        page.mouse.down()
        page.mouse.move(x, y - 10, steps=3)
        target = page.locator(".section-card:nth-child(2)").bounding_box()
        page.mouse.move(x + 40, target["y"] + 4, steps=8)
        shot("section-drag.png", 300, park=False)
        page.keyboard.press("Escape")
        page.mouse.up()

        # Save, so reopen() finds the song with its four sections.
        page.click("#btn-save")
        page.wait_for_timeout(500)

        reopen()
        page.click("#btn-layout")
        shot("layout-panel.png")

        reopen()
        page.click("#btn-lyrics")
        shot("lyrics-dialog.png")

        reopen()
        page.click("#btn-chords")
        page.wait_for_timeout(300)
        page.click("#chords-from-chart")
        page.wait_for_timeout(300)
        # Open shapes for the four chords the example plays.
        shapes = {"G": "320003", "D": "xx0232", "Em": "022000", "C": "x32010"}
        for row in page.locator("#chords-rows .chord-row").all():
            name = row.locator(".chord-name").input_value()
            if name in shapes:
                row.locator(".chord-shape").fill(shapes[name])
        shot("chord-shapes.png", 700)

        reopen()
        page.click("#btn-notation-ref")
        shot("notation-reference.png")

        reopen()
        page.click("#btn-help")
        shot("help-strip.png")

        reopen()
        page.click("#btn-export")
        shot("export-menu.png", 300)

        # A section in Tab: a bass line in the first measure, then +4 on
        # the second, with the new beats still tinted.
        reopen()
        card = page.locator(".section-card").nth(0)
        card.locator('.sec-render button[data-mode="tab"]').click()
        card.locator(".tab-add-measure").click()
        card.locator(".tab-add-measure").click()
        rows = card.locator(".tab-grid-table tbody tr")
        for row, beat, fret in [(3, 0, "3"), (3, 4, "3"), (2, 2, "5"), (2, 6, "5"),
                                (3, 8, "3"), (2, 10, "5"), (1, 13, "7"), (3, 12, "5")]:
            rows.nth(row).locator("input.tab-cell").nth(beat).fill(fret)
        card.locator(".tab-size-btn >> text=+4").nth(1).click()
        # A bar line every four beats in the first measure, the second one
        # repeated, and a name to play it again elsewhere.
        card.locator(".tab-measure-more").nth(0).click()
        page.locator("#tab-measure-menu .tab-menu-row button", has_text="4").first.click()
        rows.nth(3).locator("input.tab-cell").nth(9).click()
        card.locator(".tab-repeat").click()
        card.locator(".tab-name").fill("Hook")
        page.mouse.click(5, 5)
        page.wait_for_timeout(300)
        card.scroll_into_view_if_needed()
        shot("tab-grid.png", 250)

        # A setlist: three more songs (invented names) and the example, in
        # a collection in their running order, with one song's menu open.
        page.goto(url + "/")
        page.wait_for_timeout(600)
        page.evaluate("""async () => {
            const post = (u, b) => fetch(u, {method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(b)}).then(r => r.json());
            const names = [];
            for (const t of ['Night Drive', 'Paper Lanterns', 'Slow Burn']) {
                names.push((await post('/api/songs', {title: t, artist: 'The Invented',
                    key: 'E', time: '4/4'})).filename);
            }
            const songs = await (await fetch('/api/songs')).json();
            const example = songs.find(s => !names.includes(s.filename)).filename;
            await post('/api/library', {op: 'create', name: 'Friday set',
                songs: [names[1], example, names[0], names[2]]});
            await post('/api/library', {op: 'create', name: 'Album demos',
                songs: [names[0]]});
        }""")
        page.goto(url + "/")
        page.wait_for_selector("#song-list li")
        page.select_option("#song-view", label="Friday set (4)")
        page.wait_for_timeout(400)
        page.click("#song-list li >> nth=1 >> .song-item")
        page.wait_for_selector(".section-card")
        page.wait_for_load_state("networkidle")
        item = page.locator("#song-list li").nth(2)
        item.hover()
        item.locator(".song-more").click()
        shot("collection.png", 300, park=False)

        page.click("#song-menu >> text=Delete…")
        shot("delete-confirm.png", 300)
        page.click("#confirm-cancel")

        browser.close()
        if errors:
            print("\n  the page reported errors while the pictures were taken:")
            for e in errors:
                print("   ", e)
    return made


# ── comparing and the contact sheet ──────────────────────────────────────────

def changed(new: str, old: str) -> str:
    if not os.path.exists(old):
        return "new"
    try:
        from PIL import Image, ImageChops
    except ImportError:
        return "(install Pillow to compare)"
    a, b = Image.open(new).convert("RGB"), Image.open(old).convert("RGB")
    if a.size != b.size:
        return f"size {b.size[0]}×{b.size[1]} → {a.size[0]}×{a.size[1]}"
    diff = ImageChops.difference(a, b).convert("L").point(lambda v: 255 if v > 24 else 0)
    n = diff.histogram()[255]
    if n == 0:
        return "same"
    pct = 100 * n / (a.size[0] * a.size[1])
    # Under 0.1 %: a new version number, or the text in a field sitting a
    # pixel higher, which Chromium does now and then between two runs.
    return f"nearly the same ({n} pixels)" if pct < 0.1 else f"{pct:.1f} % changed"


def contact_sheet(paths: list[str], out: str, cols: int = 3, w: int = 640) -> str | None:
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return None
    if not paths:
        return None
    h = int(w * HEIGHT / WIDTH)
    rows = (len(paths) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * w, rows * (h + 22)), (13, 17, 23))
    d = ImageDraw.Draw(sheet)
    for i, p in enumerate(paths):
        x, y = (i % cols) * w, (i // cols) * (h + 22)
        im = Image.open(p).convert("RGB")
        im.thumbnail((w, h))
        sheet.paste(im, (x + (w - im.width) // 2, y + 22))
        d.text((x + 6, y + 5), os.path.basename(p), fill=(232, 236, 245))
    sheet.save(out)
    return out


# ── the run ──────────────────────────────────────────────────────────────────

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--dry", action="store_true",
                    help="leave the repo alone; the pictures stay in a scratch folder")
    ap.add_argument("--only", nargs="*", default=None,
                    help="write only the pictures whose names start with these")
    ap.add_argument("--out", help="write the pictures to this folder instead of screenshots/")
    ap.add_argument("--list", action="store_true", help="list the pictures and stop")
    a = ap.parse_args(argv)

    if a.list:
        for name, what in PICTURES:
            print(f"  {name:<26} {what}")
        return 0

    scratch = tempfile.mkdtemp(prefix="snt-shots-")
    with App() as app:
        made = take(app.url, scratch)

    names = [os.path.basename(m) for m in made]
    if a.only:
        names = [n for n in names if any(n.startswith(o) for o in a.only)]
    print("\n  picture                     against the repo")
    for n in names:
        print(f"  {n:<27} {changed(os.path.join(scratch, n), os.path.join(SHOTS, n))}")

    dest = a.out or (None if a.dry else SHOTS)
    if dest:
        os.makedirs(dest, exist_ok=True)
        for n in names:
            shutil.copy2(os.path.join(scratch, n), os.path.join(dest, n))
    sheet = contact_sheet([os.path.join(scratch, n) for n in names],
                          os.path.join(scratch, "contact.png"))
    print(f"\n  written to {dest or scratch}")
    if sheet:
        print(f"  all of them, small: {sheet}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

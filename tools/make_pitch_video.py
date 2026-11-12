#!/usr/bin/env python3
"""
Song Notation Tool pitch video: a vertical (9:16) MP4 that shows guitar and
bass players what the app does in under two minutes, ready to drop into a
WhatsApp chat (or an Instagram / TikTok story). It can also save the same
slides as PNGs and a PDF deck.

What it makes, from the repo's own screenshots and exported charts:

  * 10 slides with the real screens and real PDF output, animated (desktop
    screenshots pan and zoom in on the part that matters, charts slide in),
    a story-style progress bar on top and a little chart cell in the corner
    (fret number over chord symbol, with four beat dots) that follows the
    backing track's chords bar by bar;
  * a spoken voice-over (macOS `say`, `espeak-ng`, or `edge-tts`), with
    burned-in captions so it also works with the sound off;
  * a 120 BPM backing groove (Am - F - C - G, plucked bass, pads, soft kick
    and synthesised percussion; no samples, nothing to license), ducked
    smoothly under the voice. Slide changes land on a beat.

Usage (from anywhere; paths are found from this file):

    python3 tools/make_pitch_video.py                 # dist/marketing/Song_Notation_Tool_Pitch.mp4
    python3 tools/make_pitch_video.py --stills        # also one PNG per slide + a PDF deck
    python3 tools/make_pitch_video.py --stills-only   # just the PNGs and the PDF (no ffmpeg needed)
    python3 tools/make_pitch_video.py --hd            # 1080x1920 instead of 720x1280
    python3 tools/make_pitch_video.py --only 5        # just slide 5, for a quick look
    python3 tools/make_pitch_video.py --list-voices   # voices `say` has on this Mac
    python3 tools/make_pitch_video.py --voice "Daniel" --rate 175
    python3 tools/make_pitch_video.py --tts edge      # Microsoft neural voice (online), much less robotic
    python3 tools/make_pitch_video.py --tts edge --voice en-GB-RyanNeural
    python3 -m edge_tts --list-voices                 # all edge voices
    python3 tools/make_pitch_video.py --refresh-pages # re-render the PDF chart pictures (needs pdftoppm)

Needs Pillow and ffmpeg (the app itself needs neither):

    python3 -m pip install --user Pillow       # or a venv, see the error text
    brew install ffmpeg                         # macOS; apt install ffmpeg on Linux

Optional, for a more natural voice: `pip install edge-tts`, then --tts edge
(needs internet while generating; the app itself never needs it). Every
voiced slide is kept in dist/marketing/.voice-cache/, so a rerun only asks
for the lines that changed, and an edge-tts run that stops halfway (Microsoft
drops requests now and then) carries on where it stopped instead of mixing in
another voice. On a Mac,
better `say` voices are a free download: System Settings > Accessibility >
Read & Speak (Spoken Content before macOS 26) > System voice menu > Manage
Voices... (e.g. "Zoe (Premium)").

Where the pictures come from:

  * screenshots/ - the README screenshots (retake them with tools/shots.py).
    They are 1920x1200 desktop shots, too small to read on a phone, so each
    one slowly zooms in (and pans) towards the point given in SLIDES as
    "name.png@x,y,zoom" (x and y from 0 to 1 across the picture).
  * tools/pitch_media/ - pictures made for the video: paper-*.png are pages
    of real PDFs the app exported from the invented golden-test songs in
    tests/golden/, cropped to the chart (re-render them after a change to
    the PDF output with --refresh-pages, which needs pdftoppm from poppler:
    `brew install poppler`); paper-handwritten.jpg is a photo of a chart
    written by hand, the kind this app replaces.

A missing picture is skipped with a warning, so the video still builds. To
change what the video says, edit SLIDES below.
"""

from __future__ import annotations

import argparse
import array
import copy
import hashlib
import math
import os
import random
import re
import shutil
import subprocess
import sys
import tempfile
import time
import wave

try:
    from PIL import Image, ImageDraw, ImageFilter, ImageFont
except ImportError:
    sys.exit(
        "This tool needs Pillow (Song Notation Tool itself doesn't, only this marketing tool).\n\n"
        "Install it with:\n"
        "    python3 -m pip install --user Pillow\n\n"
        "If pip refuses with \"externally-managed-environment\" (Homebrew's Python), either:\n"
        "    python3 -m pip install --user --break-system-packages Pillow\n"
        "or use a virtual environment, once:\n"
        "    python3 -m venv .venv && .venv/bin/pip install Pillow\n"
        "and then run:  .venv/bin/python tools/make_pitch_video.py"
    )

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SHOTS = os.path.join(ROOT, "screenshots")
PITCH = os.path.join(ROOT, "tools", "pitch_media")
MEDIA_DIRS = [SHOTS, PITCH]
ICON = os.path.join(ROOT, "web", "logo", "icon-256.png")
OUT_DIR = os.path.join(ROOT, "dist", "marketing")
# Finished voice clips, kept between runs: a rerun with the same text, voice
# and rate reuses them instead of asking the voice engine again.
TTS_CACHE = os.path.join(OUT_DIR, ".voice-cache")
NAME = "Song Notation Tool"
FILE_STEM = "Song_Notation_Tool_Pitch"
TAGLINE = "One-page charts for guitar & bass"

RATE = 22050
BPM = 120
BEAT = 60.0 / BPM     # 0.5 s
FPS = 30
XFADE = 0.35          # crossfade between slides, seconds

# The browser app's colours (web/style.css, dark theme, since v0.34.0).
# DIM is a little brighter than --lbl-gray, for a subtitle on a phone.
BG = (18, 19, 25)          # --bg
CARD = (38, 41, 56)        # --surface
LINE = (58, 62, 80)        # --border-strong
TEXT = (230, 232, 239)     # --fg
DIM = (176, 181, 197)
WHITE = (255, 255, 255)
BLUE = (122, 162, 255)     # --blue
MAUVE = (196, 161, 255)    # --mauve
AMBER = (249, 200, 107)    # --amber
GREEN = (166, 227, 161)    # --green
RED = (255, 138, 107)      # --accent (coral; the name stays for the slides that use it)

URL = "github.com/giopas/song-notation-tool"

# Text the voice says differently from how the captions write it.
SPOKEN = {
    URL: "github dot com slash gio pass slash song notation tool",
    ".sng": "S N G",
    "PDF": "P D F",
    "PDFs": "P D Fs",
}

SLIDES = [
    {
        "kicker": "BY A BASS & GUITAR PLAYER",
        "color": RED,
        "title": NAME,
        "sub": "The chart you scribble on paper, as a clean one-page PDF.",
        "media": ["paper-handwritten.jpg@0.45,0.45,1.25", "paper-basics.png@0.3,0.3,1.5"],
        "bullets": [
            "Fret numbers over chords, the way you'd write it by hand",
            "Riffs named once, repeats and endings: no bar written twice",
            "Free for Mac, Windows and Linux; songs stay on your computer",
        ],
        "vo": "Hi! I play bass and guitar, and for years I wrote my songs on paper: fret numbers over chords, "
              "repeats, riffs. So I built Song Notation Tool, to turn those scribbles into a clean "
              "one-page chart.",
    },
    {
        "kicker": "1 · WRITE IT AS YOU PLAY",
        "color": BLUE,
        "title": "One Line per Section",
        "sub": "Type 5A 5D [5A 5D]x2 and the chart draws itself as you type.",
        "media": ["editor.png@0.42,0.3,1.9", "notation-reference.png@0.45,0.5,1.8"],
        "bullets": [
            "Repeats, endings, segno and coda from a quick-insert palette",
            "Type // to break a long section into phrases, as on paper",
            "A typo shows up in its row and never wipes what you wrote",
        ],
        "vo": "Each section is one short line. Type frets and chords the way you'd jot them down, "
              "and the chart draws itself as you type.",
    },
    {
        "kicker": "2 · RIFFS & LICKS",
        "color": MAUVE,
        "title": "Name a Riff Once, Reuse It",
        "sub": "Write a lick as a little tab figure, name it, call it anywhere.",
        "media": ["paper-licks-tab.png@0.3,0.3,1.5", "paper-marks-and-techniques.png@0.3,0.45,1.5"],
        "bullets": [
            "Edit a riff once and every place that plays it follows",
            "Hammer-ons, pull-offs, slides, bends and vibrato, in tab letters",
            "A repeated chorus points to the first one, not a second copy",
        ],
        "vo": "Riffs get a name, just like on paper. Change a riff once, and every place that plays it follows. "
              "Licks print as little tab blocks, with slides and bends.",
    },
    {
        "kicker": "3 · SHAPE THE SONG",
        "color": AMBER,
        "title": "Drag, Copy, Rearrange",
        "sub": "Move a section by its handle; a copy plays what the original holds.",
        "media": ["section-drag.png@0.35,0.3,1.9", "section-copy.png@0.85,0.76,1.9"],
        "bullets": [
            "Dragging folds every card to its name: the whole song fits",
            "Copy a section to the top, or after any other section",
            "Add a bass line under the guitar chords, in the same section",
        ],
        "vo": "Rearranging is easy. Drag a section by its handle, or copy a chorus to where it comes back. "
              "You can even put a bass line under the guitar chords.",
    },
    {
        "kicker": "4 · TRANSPOSE & CHORDS",
        "color": GREEN,
        "title": "Change Key in One Click",
        "sub": "Shift the whole song or one section; what you typed stays as is.",
        "media": ["chord-shapes.png@0.48,0.5,1.8", "paper-chord-shapes-transposed.png@0.3,0.3,1.4"],
        "bullets": [
            "Up two, then down two, and you're back exactly where you were",
            "Chord shapes like x32010 print as small diagrams",
            "Shapes re-voice when you transpose: open C up a tone is open D",
        ],
        "vo": "Singer wants it a tone lower? Transpose the whole song, or just one section. "
              "Your chord shapes move too, the way a guitarist would play them.",
    },
    {
        "kicker": "5 · LYRICS ALONGSIDE",
        "color": RED,
        "title": "Words Next to the Chart",
        "sub": "Paste or import the lyrics, then split them across your sections.",
        "media": ["lyrics-dialog.png@0.5,0.28,1.8", "paper-lyrics-beside.png@0.4,0.3,1.4"],
        "bullets": [
            "Mark blocks with === Verse 1 === and split them in one go",
            "Print the words beside each section, under it, or as one block",
            "Nothing prints until you choose where the words go",
        ],
        "vo": "Need the words? Paste the lyrics, split them across your sections, "
              "and print them right beside what you play.",
    },
    {
        "kicker": "6 · PRINT IT YOUR WAY",
        "color": BLUE,
        "title": "One Page on the Stand",
        "sub": "Fit to one page, two columns, colour by section, or black and white.",
        "media": ["layout-panel.png@0.18,0.45,1.9", "paper-gutter-bw-two-columns.png@0.3,0.3,1.4"],
        "bullets": [
            "Fit to page grows the type until the sheet is full",
            "Two columns when that means bigger type or one sheet fewer",
            "Each section type has its colour: every Verse looks alike",
        ],
        "vo": "One panel decides how it prints: all on one sheet, two columns, "
              "a colour for each kind of section, or black and white for the photocopier.",
    },
    {
        "kicker": "7 · FLOOR SHEETS",
        "color": MAUVE,
        "title": "Read It From Your Feet",
        "sub": "The biggest type that fits on one page, to tape down by your pedals.",
        "media": ["paper-floor.png@0.3,0.3,1.3", "export-menu.png@0.87,0.15,2.0"],
        "bullets": [
            "Tries portrait and landscape, one or two columns, keeps the biggest",
            "Any section can print larger or smaller than the rest",
            "A riff played again can print as just its name",
        ],
        "vo": "For the gig, there's the floor sheet: the biggest type that fits on one page, "
              "to read from your feet.",
    },
    {
        "kicker": "8 · NOTHING GETS LOST",
        "color": AMBER,
        "title": "Every PDF Carries Its Song",
        "sub": "Open an exported PDF in the app and keep editing the song.",
        "media": ["start-here.png@0.59,0.45,1.8", "paper-basics.png@0.3,0.3,1.5"],
        "bullets": [
            "Send a friend the PDF: they can import it and keep editing",
            "Songs are plain .sng files named Artist - Title, in Documents",
            "Export TXT or PDF, or print with your system's preview",
        ],
        "vo": "Every PDF you export carries the song inside it. Send a chart to a friend, "
              "and they can open it in the app and keep editing.",
    },
    {
        "kicker": "9 · YOUR TURN",
        "color": GREEN,
        "title": "Write Up Your Next Song!",
        "sub": "Free and open source, made for how players write songs down.",
        "media": ["editor.png@0.42,0.3,1.6", "paper-handwritten.jpg@0.45,0.45,1.25"],
        "bullets": [
            "1. Open " + URL,
            "2. Download it for Mac, Windows or Linux from Releases",
            "3. Write up one song you know, and print it",
        ],
        "vo": "Have a look at " + URL + ", download it, and write up one song you know. "
              "I'd love to hear what you think!",
        "footer": NAME + " {version} · free & open source (MIT) · github.com/giopas",
    },
]

# Pages for --refresh-pages: picture name -> (golden song, settings to change).
PAGES = {
    "paper-basics.png": ("basics.sng", {}),
    "paper-licks-tab.png": ("licks-tab.sng", {}),
    "paper-marks-and-techniques.png": ("marks-and-techniques.sng", {}),
    "paper-chord-shapes-transposed.png": ("chord-shapes-transposed.sng", {}),
    "paper-lyrics-beside.png": ("lyrics-beside.sng", {}),
    "paper-gutter-bw-two-columns.png": ("gutter-bw-two-columns.sng", {}),
    "paper-floor.png": ("licks-name-only.sng", {"pdf_scale": "floor"}),
}


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def app_version() -> str:
    try:
        with open(os.path.join(ROOT, "constants.py"), encoding="utf-8") as f:
            m = re.search(r"APP_VERSION\s*=\s*[\"'](\d+\.\d+\.\d+)[\"']", f.read())
        return "v" + m.group(1) if m else ""
    except OSError:
        return ""


def ease_out(x: float) -> float:
    x = max(0.0, min(1.0, x))
    return 1 - (1 - x) ** 3


def run(cmd: list, **kw) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, check=True, **kw)


def read_wav_floats(path: str) -> list:
    with wave.open(path, "rb") as w:
        data = array.array("h", w.readframes(w.getnframes()))
    if sys.byteorder == "big":
        data.byteswap()
    return [v / 32768.0 for v in data]


def write_wav(path: str, samples: list) -> None:
    pcm = array.array("h", (max(-32767, min(32767, int(v * 32767))) for v in samples))
    if sys.byteorder == "big":
        pcm.byteswap()
    with wave.open(path, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(RATE)
        w.writeframes(pcm.tobytes())


def parse_media(spec: str):
    """"name.png@x,y,zoom" -> (name, (x, y), zoom). x, y, zoom are optional."""
    name, _, opts = spec.partition("@")
    focus, zoom = (0.5, 0.45), None
    if opts:
        parts = [float(p) for p in opts.split(",") if p.strip()]
        if len(parts) >= 2:
            focus = (parts[0], parts[1])
        if len(parts) >= 3:
            zoom = parts[2]
    return name, focus, zoom


def find_media(name: str):
    for d in MEDIA_DIRS:
        p = os.path.join(d, name)
        if os.path.exists(p):
            return p
    return None


# ---------------------------------------------------------------------------
# --refresh-pages: real PDF output, rasterised and cropped to the chart
# ---------------------------------------------------------------------------

def refresh_pages() -> None:
    if not shutil.which("pdftoppm"):
        sys.exit("--refresh-pages needs pdftoppm, from poppler. On a Mac: brew install poppler "
                 "(Linux: sudo apt install poppler-utils).")
    sys.path.insert(0, ROOT)
    try:
        import export  # noqa: E402  (the app's own PDF writer)
    except Exception as e:
        sys.exit(f"Couldn't load the app's export.py ({e}).")
    os.makedirs(PITCH, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="snt-pages-") as tmp:
        for out_name, (song, changes) in PAGES.items():
            doc = export.load_song_file(os.path.join(ROOT, "tests", "golden", song))
            doc = copy.deepcopy(doc)
            doc.update(changes)
            pdf = os.path.join(tmp, "page.pdf")
            with open(pdf, "wb") as f:
                f.write(export.build_pdf(doc))
            run(["pdftoppm", "-r", "150", "-png", "-f", "1", "-l", "1", "-singlefile", pdf,
                 os.path.join(tmp, "page")])
            im = Image.open(os.path.join(tmp, "page.png")).convert("RGB")
            # Crop to the ink (leave the footer out), with a margin.
            body = im.crop((0, 0, im.width, int(im.height * 0.94)))
            inv = Image.eval(body.convert("L"), lambda v: 255 if v < 235 else 0)
            box = inv.getbbox() or (0, 0, im.width, im.height)
            m = int(im.width * 0.03)
            box = (max(0, box[0] - m), max(0, box[1] - m), min(im.width, box[2] + m),
                   min(body.height, box[3] + m))
            # Keep the top of the page, about as wide as the video's picture
            # area (about 8:7): a whole A4 page would be too small to read there.
            box = (box[0], box[1], box[2], min(box[3], box[1] + int((box[2] - box[0]) / 1.15)))
            im.crop(box).save(os.path.join(PITCH, out_name), optimize=True)
            print(f"   {out_name}  <- tests/golden/{song}")
    print(f"Pages written to {PITCH}")


# ---------------------------------------------------------------------------
# Voice-over
# ---------------------------------------------------------------------------

# Nicer built-in macOS voices first; whichever is installed wins.
PREFERRED_SAY_VOICES = [
    "Zoe (Premium)", "Ava (Premium)", "Evan (Premium)", "Jamie (Premium)",
    "Zoe (Enhanced)", "Ava (Enhanced)", "Evan (Enhanced)", "Samantha (Enhanced)", "Daniel (Enhanced)",
    "Samantha", "Daniel", "Alex",
]


def say_voices() -> list:
    """[(name, locale)] of the voices `say` knows."""
    if not shutil.which("say"):
        return []
    out = subprocess.run(["say", "-v", "?"], capture_output=True, text=True).stdout
    voices = []
    for line in out.splitlines():
        m = re.match(r"^(.+?)\s{2,}([a-z]{2}[_-][A-Za-z]{2,})", line)
        if m:
            voices.append((m.group(1).strip(), m.group(2)))
    return voices


def pick_say_voice(requested: str = None):
    names = [n for n, loc in say_voices()]
    if requested:
        if requested not in names:
            print(f"   ! `say` has no voice called {requested!r}; try --list-voices. Using the default.")
            return None
        return requested
    english = {n for n, loc in say_voices() if loc.lower().startswith("en")}
    for v in PREFERRED_SAY_VOICES:
        if v in english:
            return v
    return None


def spoken(text: str) -> str:
    # Longest first, so "PDFs" isn't caught by "PDF".
    for written in sorted(SPOKEN, key=len, reverse=True):
        text = re.sub(r"(?<![\w.])" + re.escape(written) + r"(?!\w)", SPOKEN[written], text)
    return text


def edge_tts_command():
    """How to run edge-tts: the command if it's on the PATH, otherwise as a
    module of this Python (`pip install --user` puts the command in a folder,
    e.g. ~/Library/Python/3.x/bin, that often isn't on the PATH)."""
    if shutil.which("edge-tts"):
        return ["edge-tts"]
    try:
        import importlib.util
        if importlib.util.find_spec("edge_tts"):
            return [sys.executable, "-m", "edge_tts"]
    except (ImportError, ValueError):
        pass
    return None


class Voice:
    """Turns a line of text into a 22050 Hz mono WAV with one engine."""

    def __init__(self, engine: str, voice: str = None, rate: int = 180):
        self.engine, self.voice, self.rate = engine, voice, rate
        self.label = engine
        self.done = 0  # slides this engine has voiced (cached or not)

    def cache_path(self, said: str) -> str:
        key = hashlib.sha1(f"{self.engine}|{self.voice}|{self.rate}|{said}".encode("utf-8")).hexdigest()[:16]
        return os.path.join(TTS_CACHE, f"{self.engine}-{key}.wav")

    @classmethod
    def choose(cls, tts: str, voice: str, rate: int) -> "Voice":
        if tts == "none":
            return cls("none")
        if tts == "edge":
            cmd = edge_tts_command()
            if cmd:
                v = cls("edge", voice or "en-US-AndrewNeural", rate)
                v.edge_cmd = cmd
                v.label = f"edge-tts ({v.voice})"
                return v
            print("   ! edge-tts isn't installed for this Python "
                  f"({sys.executable} -m pip install --user edge-tts); falling back.")
        if tts in ("say", "auto", "edge") and shutil.which("say"):
            v = cls("say", pick_say_voice(voice if tts == "say" else None), rate)
            v.label = f"macOS say ({v.voice or 'system voice'})"
            return v
        if tts in ("espeak", "auto", "edge", "say") and shutil.which("espeak-ng"):
            v = cls("espeak", voice if tts == "espeak" and voice else "en-us", rate)
            v.label = f"espeak-ng ({v.voice})"
            return v
        print("   ! No text-to-speech found: the video gets captions and music, no voice.")
        return cls("none")

    def render(self, text: str, out_wav: str, tmp: str) -> bool:
        if self.engine == "none":
            return False
        said = spoken(text)
        cached = self.cache_path(said)
        if os.path.exists(cached):
            shutil.copyfile(cached, out_wav)
            self.done += 1
            return True
        # edge-tts goes over the internet to Microsoft, which sometimes drops
        # a request or two in a row: try again before giving up on it.
        tries = 3 if self.engine == "edge" else 1
        for attempt in range(tries):
            try:
                self._synthesise(said, out_wav, tmp)
                try:
                    os.makedirs(TTS_CACHE, exist_ok=True)
                    shutil.copyfile(out_wav, cached)
                except OSError:
                    pass
                self.done += 1
                return True
            except (subprocess.CalledProcessError, subprocess.TimeoutExpired, OSError) as e:
                reason = failure_reason(e)
                if attempt + 1 < tries:
                    wait = 4 * (attempt + 1)
                    print(f"   ! {self.label} didn't answer ({reason}); trying again in {wait} s...")
                    time.sleep(wait)
        if self.engine == "edge":
            print(f"   ! {self.label} failed: {reason}")
            if self.done:
                # Some slides already have the edge voice: switching to another
                # voice now would give a video with two different speakers.
                sys.exit(
                    f"\nedge-tts stopped answering after {self.done} slide(s). Rather than mix two voices,\n"
                    "the video was not made. The slides already voiced are saved, so run the same\n"
                    "command again in a minute or two and it carries on from here.\n"
                    "Or use the Mac's own voice for the whole video:  --tts say")
            # Nothing voiced yet: say why once, and use the local voice for all of it.
            print("     Check with:  " + " ".join(self.edge_cmd) + ' --text "Hello" --write-media /tmp/hello.mp3')
            print(f"     Often fixed by updating it:  {sys.executable} -m pip install --user -U edge-tts")
            fallback = Voice.choose("auto", None, self.rate)
            if fallback.engine == "none":
                self.engine, self.label = "none", "no voice"
                return False
            print(f"     Using {fallback.label} for the whole video instead.")
            self.__dict__.update(fallback.__dict__)
            return self.render(text, out_wav, tmp)
        print(f"   ! {self.label} failed ({reason}); this slide gets no voice.")
        return False

    def _synthesise(self, said: str, out_wav: str, tmp: str) -> None:
        raw = os.path.join(tmp, "tts_raw")
        if self.engine == "say":
            raw += ".aiff"
            cmd = ["say", "-r", str(self.rate), "-o", raw]
            if self.voice:
                cmd += ["-v", self.voice]
            run(cmd + [said], capture_output=True, text=True)
        elif self.engine == "espeak":
            raw += ".wav"
            run(["espeak-ng", "-v", self.voice, "-s", str(int(self.rate * 0.92)), "-w", raw, said],
                capture_output=True, text=True)
        elif self.engine == "edge":
            raw += ".mp3"
            if os.path.exists(raw):
                os.remove(raw)
            pct = round((self.rate / 180.0 - 1) * 100)
            run(self.edge_cmd + ["--voice", self.voice, f"--rate={pct:+d}%", "--text", said,
                 "--write-media", raw], capture_output=True, text=True, timeout=90)
            if not os.path.exists(raw) or os.path.getsize(raw) < 1000:
                raise OSError("edge-tts returned no audio")
        # Clean, consistent level: trim silence at both ends, gentle compression, normalise.
        run(["ffmpeg", "-loglevel", "error", "-y", "-i", raw, "-af",
             "silenceremove=start_periods=1:start_threshold=-50dB,"
             "areverse,silenceremove=start_periods=1:start_threshold=-50dB,areverse,"
             "highpass=f=80,acompressor=threshold=-18dB:ratio=3:attack=5:release=120,"
             "loudnorm=I=-16:TP=-1.5:LRA=7",
             "-ar", str(RATE), "-ac", "1", "-sample_fmt", "s16", out_wav],
            capture_output=True, text=True)


def failure_reason(e: Exception) -> str:
    """One line on why a command failed: the last thing it printed, if anything."""
    if isinstance(e, subprocess.TimeoutExpired):
        return f"no answer after {int(e.timeout)} s"
    if isinstance(e, subprocess.CalledProcessError):
        out = "\n".join(x for x in (e.stderr, e.stdout) if isinstance(x, str))
        lines = [ln.strip() for ln in out.splitlines() if ln.strip()]
        what = os.path.basename(str(e.cmd[0])) if e.cmd else "a command"
        return lines[-1][:200] if lines else f"{what} exited with {e.returncode}"
    return str(e)


# ---------------------------------------------------------------------------
# Timing: slide lengths snapped to the beat
# ---------------------------------------------------------------------------

def plan_timing(slides: list, voice_lens: list) -> list:
    """Per slide: start, dur, voice_at (seconds, absolute). Every slide starts on a beat."""
    plan, t = [], 0.0
    for i, (slide, vo) in enumerate(zip(slides, voice_lens)):
        lead = 2.0 if i == 0 else 0.6            # one bar of groove before the first word
        tail = 3.0 if i == len(slides) - 1 else 0.7
        dur = max(5.0, lead + vo + tail)
        dur = math.ceil(dur / BEAT - 1e-9) * BEAT
        plan.append({"start": t, "dur": dur, "voice_at": t + lead, "voice_len": vo})
        t += dur
    return plan


def estimate_speech(text: str) -> float:
    return max(2.5, len(text.split()) / 2.7)


# ---------------------------------------------------------------------------
# Soundtrack: bass + pads + kick + synthesised percussion, ducked under the voice
# ---------------------------------------------------------------------------

def synth_kit() -> dict:
    """The app makes no sounds of its own, so the percussion is synthesised:
    a side-stick click, a closed hi-hat, a wood block and a shaker."""
    rnd = random.Random(7)

    def noise(n):
        return [rnd.uniform(-1, 1) for _ in range(n)]

    def highpass(x, k=0.85):
        out, prev_x, prev_y = [], 0.0, 0.0
        for v in x:
            y = k * (prev_y + v - prev_x)
            out.append(y)
            prev_x, prev_y = v, y
        return out

    n = int(0.06 * RATE)
    rim_noise = highpass(noise(n), 0.7)
    rim = [(0.55 * rim_noise[i] + 0.8 * math.sin(2 * math.pi * 1750 * i / RATE)) * math.exp(-i / (0.012 * RATE))
           for i in range(n)]
    n = int(0.05 * RATE)
    hat_noise = highpass(highpass(noise(n), 0.6), 0.6)
    hat = [hat_noise[i] * math.exp(-i / (0.010 * RATE)) for i in range(n)]
    n = int(0.12 * RATE)
    wood = [(math.sin(2 * math.pi * 1020 * i / RATE) + 0.4 * math.sin(2 * math.pi * 2610 * i / RATE))
            * math.exp(-i / (0.025 * RATE)) for i in range(n)]
    n = int(0.09 * RATE)
    sh_noise = highpass(noise(n), 0.5)
    shaker = [sh_noise[i] * min(1.0, i / (0.03 * RATE)) * math.exp(-max(0, i - 0.03 * RATE) / (0.02 * RATE))
              for i in range(n)]
    return {"rim": rim, "hat": hat, "wood": wood, "shaker": shaker}


NOTE = {"A2": 110.0, "C3": 130.81, "F2": 87.31, "G2": 98.0}
# Am - F - C - G, one bar each: (bass root, pad triad)
PROGRESSION = [
    (NOTE["A2"], (220.0, 261.63, 329.63)),
    (NOTE["F2"], (174.61, 220.0, 261.63)),
    (NOTE["C3"], (196.0, 261.63, 329.63)),
    (NOTE["G2"], (196.0, 246.94, 293.66)),
]
# The same four bars as the corner motif shows them: (fret, chord), where a
# bass player would play the root (A on the E string at 5, F on the A string
# at 8, C on the E string at 8, G on the E string at 3).
MOTIF = [("5", "Am"), ("8", "F"), ("8", "C"), ("3", "G")]


def _chord_bar(root: float, triad: tuple) -> list:
    """One bar (2 s) of bass + pad + soft kick, as floats."""
    n = int(4 * BEAT * RATE)
    out = [0.0] * n
    # Bass: beat 1, the "and" of 2, beat 3 (a little push), plucked and warm.
    for beat_pos, vel in ((0.0, 1.0), (1.5, 0.55), (2.0, 0.85)):
        s0 = int(beat_pos * BEAT * RATE)
        length = int(0.9 * BEAT * RATE * (2 if beat_pos == 2.0 else 1))
        for j in range(min(length, n - s0)):
            t = j / RATE
            env = min(1.0, t / 0.006) * math.exp(-3.5 * t) * vel
            w = 2 * math.pi * root * t
            out[s0 + j] += 0.30 * env * (math.sin(w) + 0.25 * math.sin(2 * w) + 0.08 * math.sin(3 * w))
    # Pad: the triad, slow swell, two slightly detuned sines per note.
    for j in range(n):
        t = j / RATE
        env = min(1.0, t / 0.35) * min(1.0, (n - j) / (0.25 * RATE))
        v = 0.0
        for f in triad:
            v += math.sin(2 * math.pi * f * t) + math.sin(2 * math.pi * f * 1.004 * t)
        out[j] += 0.012 * env * v
    # Soft kick on 1 and 3: a pitch-dropping sine.
    for beat in (0, 2):
        s0 = int(beat * BEAT * RATE)
        ph = 0.0
        for j in range(int(0.22 * RATE)):
            t = j / RATE
            f = 48 + 90 * math.exp(-t * 28)
            ph += 2 * math.pi * f / RATE
            out[s0 + j] += 0.42 * math.sin(ph) * math.exp(-t * 14)
    return out


def build_music(total: float, kit: dict) -> list:
    n = int(total * RATE)
    bars = [_chord_bar(r, tri) for r, tri in PROGRESSION]
    music = [0.0] * n
    bar_n = len(bars[0])
    for b in range(0, n // bar_n + 1):
        src, s0 = bars[b % 4], b * bar_n
        m = min(bar_n, n - s0)
        if m <= 0:
            break
        music[s0:s0 + m] = [a + c for a, c in zip(music[s0:s0 + m], src[:m])]

    def hit(sample: list, at: float, gain: float) -> None:
        s0 = int(at * RATE)
        for j in range(min(len(sample), n - s0)):
            music[s0 + j] += sample[j] * gain

    eighth = BEAT / 2
    for k in range(int(total / eighth)):
        at = k * eighth
        beat_in_bar = (k // 2) % 4
        bar = k // 8
        if k % 2 == 0:
            if beat_in_bar in (1, 3):
                hit(kit["rim"], at, 0.28)                     # side stick backbeat
            if beat_in_bar == 0 and bar % 4 == 0:
                hit(kit["wood"], at, 0.12)                    # wood block marks every 4 bars
        hit(kit["hat"], at, 0.10 if k % 2 == 0 else 0.06)     # closed hat 8ths, quiet
        if beat_in_bar == 3 and k % 2 == 1 and bar % 2 == 1:
            hit(kit["shaker"], at, 0.10)                      # a little shaker pickup

    # Fade in over a beat, out over the last 2.5 s.
    for j in range(min(n, int(BEAT * RATE))):
        music[j] *= j / (BEAT * RATE)
    fade = int(2.5 * RATE)
    for j in range(max(0, n - fade), n):
        music[j] *= (n - j) / fade
    return music


def mix_soundtrack(total: float, plan: list, vo_wavs: list, music_on: bool, out_wav: str) -> None:
    n = int(total * RATE)
    voice = [0.0] * n
    for p, wav in zip(plan, vo_wavs):
        if not wav:
            continue
        s = read_wav_floats(wav)
        s0 = int(p["voice_at"] * RATE)
        m = min(len(s), n - s0)
        voice[s0:s0 + m] = s[:m]

    if not music_on:
        write_wav(out_wav, voice)
        return

    music = build_music(total, synth_kit())
    # Ducking: follow the voice's level in 10 ms blocks (fast attack, slow
    # release), so the music dips smoothly while someone talks instead of
    # pumping on every syllable.
    block = RATE // 100
    gains, env = [], 0.0
    for b0 in range(0, n, block):
        seg = voice[b0:b0 + block]
        rms = math.sqrt(sum(v * v for v in seg) / max(1, len(seg)))
        env = max(rms, env * 0.965)  # ~280 ms release
        duck = min(1.0, env / 0.04)
        gains.append(0.85 - 0.6 * duck)  # 0.85 alone, 0.25 under the voice
    out = [0.0] * n
    for i in range(n):
        b = i // block
        g0 = gains[b]
        g1 = gains[b + 1] if b + 1 < len(gains) else g0
        g = g0 + (g1 - g0) * ((i % block) / block)
        out[i] = voice[i] + music[i] * g
    # Peak at most -1.4 dBFS: AAC overshoots a little, and the result must
    # stay at or under -1 dB.
    peak = max(1e-9, max(abs(v) for v in out))
    if peak > 0.85:
        k = 0.85 / peak
        out = [v * k for v in out]
    write_wav(out_wav, out)


# ---------------------------------------------------------------------------
# Drawing
# ---------------------------------------------------------------------------

FONT_FILES = {
    False: ["/System/Library/Fonts/Supplemental/Arial.ttf", "/Library/Fonts/Arial.ttf",
            "C:/Windows/Fonts/arial.ttf", "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"],
    True: ["/System/Library/Fonts/Supplemental/Arial Bold.ttf", "/Library/Fonts/Arial Bold.ttf",
           "C:/Windows/Fonts/arialbd.ttf", "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
           "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
           "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"],
}
_fonts: dict = {}


def font(size: int, bold: bool = False):
    key = (size, bold)
    if key not in _fonts:
        for path in FONT_FILES[bold]:
            if os.path.exists(path):
                _fonts[key] = ImageFont.truetype(path, size)
                break
        else:
            _fonts[key] = ImageFont.load_default()
    return _fonts[key]


def wrap(draw, text: str, f, max_w: int) -> list:
    lines, cur = [], ""
    for word in text.split():
        trial = (cur + " " + word).strip()
        if draw.textlength(trial, font=f) <= max_w or not cur:
            cur = trial
        else:
            lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return lines


def rounded_mask(size: tuple, radius: int) -> Image.Image:
    m = Image.new("L", size, 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, size[0] - 1, size[1] - 1], radius=radius, fill=255)
    return m


def with_alpha(im: Image.Image, a: float) -> Image.Image:
    if a >= 0.999:
        return im
    out = im.copy()
    out.putalpha(im.getchannel("A").point(lambda v: int(v * a)))
    return out


def smoothstep(x: float) -> float:
    x = max(0.0, min(1.0, x))
    return x * x * (3 - 2 * x)


class Renderer:
    ZOOM = 1.32  # default for how far a wide picture zooms in over its turn

    def __init__(self, width: int, captions: bool, version: str):
        self.W, self.H = width, width * 16 // 9
        self.S = width / 720.0
        self.captions = captions
        self.version = version
        self.icon = None
        if os.path.exists(ICON):
            self.icon = Image.open(ICON).convert("RGBA").resize((self.s(46), self.s(46)), Image.LANCZOS)
        self._cap_cache: dict = {}

    def s(self, v: float) -> int:
        return int(round(v * self.S))

    # -- static parts of a slide ------------------------------------------------

    def prepare(self, slide: dict, idx: int, total: int) -> dict:
        s, W, H = self.s, self.W, self.H
        base = Image.new("RGB", (W, H), BG)
        d = ImageDraw.Draw(base)
        color = slide["color"]

        # A soft glow of the slide's colour behind the title.
        glow = Image.new("L", (W, H), 0)
        ImageDraw.Draw(glow).ellipse([-s(200), -s(260), W + s(200), s(420)], fill=40)
        glow = glow.filter(ImageFilter.GaussianBlur(s(90)))
        base.paste(Image.new("RGB", (W, H), color), (0, 0), glow)

        # Header: icon, name, tagline (the chart-cell motif is drawn per frame)
        if self.icon:
            base.paste(self.icon, (s(28), s(34)), self.icon)
        d.text((s(86), s(33)), NAME, fill=WHITE, font=font(s(27), True))
        d.text((s(87), s(64)), TAGLINE, fill=DIM, font=font(s(15)))

        # Kicker pill
        y = s(112)
        fk = font(s(17), True)
        kw = int(d.textlength(slide["kicker"], font=fk)) + s(32)
        d.rounded_rectangle([s(28), y, s(28) + kw, y + s(36)], radius=s(18), fill=CARD, outline=color, width=max(2, s(2)))
        d.text((s(44), y + s(8)), slide["kicker"], fill=color, font=fk)
        y += s(52)

        ft = font(s(44) if len(slide["title"]) > 12 else s(64), True)
        lh = int(ft.size * 1.16)
        for line in wrap(d, slide["title"], ft, W - s(56)):
            d.text((s(28), y), line, fill=WHITE, font=ft)
            y += lh
        y += s(6)
        fs = font(s(22))
        for line in wrap(d, slide["sub"], fs, W - s(56)):
            d.text((s(28), y), line, fill=DIM, font=fs)
            y += s(30)
        vis_top = y + s(22)

        # Footer line (last slide)
        bottom_limit = H - s(26)
        if slide.get("footer"):
            ff = font(s(15))
            text = slide["footer"].format(version=self.version).replace("  ", " ")
            tw = d.textlength(text, font=ff)
            d.text(((W - tw) / 2, H - s(42)), text, fill=DIM, font=ff)
            bottom_limit = H - s(56)

        # Bullet cards, bottom-up, above the caption band.
        cap_top = H - s(178) if self.captions else bottom_limit
        cap_top = min(cap_top, bottom_limit)
        fb = font(s(21))
        cards = []
        for text in slide["bullets"]:
            lines = wrap(d, text, fb, W - s(56) - s(44))
            h = s(28) + len(lines) * s(28)
            card = Image.new("RGBA", (W - s(56), h), (0, 0, 0, 0))
            cd = ImageDraw.Draw(card)
            cd.rounded_rectangle([0, 0, card.width - 1, h - 1], radius=s(14), fill=CARD + (255,), outline=LINE, width=1)
            cd.rounded_rectangle([0, 0, s(6), h - 1], radius=s(3), fill=color)
            ty = s(14)
            for ln in lines:
                cd.text((s(26), ty), ln, fill=TEXT, font=fb)
                ty += s(28)
            cards.append(card)
        gap = s(12)
        total_h = sum(c.height for c in cards) + gap * (len(cards) - 1)
        by = cap_top - s(14) - total_h
        placed = []
        for c in cards:
            placed.append((c, by))
            by += c.height + gap
        vis_bottom = placed[0][1] - s(24) if placed else cap_top - s(24)

        media = self.load_media(slide, W - s(56), vis_bottom - vis_top)
        return {"base": base, "cards": placed, "media": media, "vis": (vis_top, vis_bottom),
                "color": color, "idx": idx, "total": total, "cap_y": cap_top}

    # -- screenshots, pages, GIFs -------------------------------------------------

    def frame_card(self, im: Image.Image, box_w: int, box_h: int) -> Image.Image:
        """A picture fitted into the box: a phone frame for tall screenshots,
        rounded corners and a soft shadow for everything else."""
        s = self.s
        im = im.convert("RGB")
        phone = im.height / im.width > 1.3
        pad = s(18)  # room for the shadow
        if phone:
            bezel = s(10)
            scale = min((box_w - 2 * pad - 2 * bezel) / im.width, (box_h - 2 * pad - 2 * bezel) / im.height)
            sw, sh = max(1, int(im.width * scale)), max(1, int(im.height * scale))
            screen = im.resize((sw, sh), Image.LANCZOS)
            cw, ch = sw + 2 * bezel, sh + 2 * bezel
            body = Image.new("RGBA", (cw, ch), (0, 0, 0, 0))
            ImageDraw.Draw(body).rounded_rectangle([0, 0, cw - 1, ch - 1], radius=s(34), fill=(8, 9, 11, 255),
                                                   outline=(70, 74, 84, 255), width=max(1, s(2)))
            body.paste(screen, (bezel, bezel), rounded_mask((sw, sh), s(24)))
            radius = s(34)
        else:
            scale = min((box_w - 2 * pad) / im.width, (box_h - 2 * pad) / im.height)
            cw, ch = max(1, int(im.width * scale)), max(1, int(im.height * scale))
            body = im.resize((cw, ch), Image.LANCZOS).convert("RGBA")
            body.putalpha(rounded_mask((cw, ch), s(14)))
            ImageDraw.Draw(body).rounded_rectangle([0, 0, cw - 1, ch - 1], radius=s(14), outline=LINE, width=1)
            radius = s(14)
        out = Image.new("RGBA", (cw + 2 * pad, ch + 2 * pad), (0, 0, 0, 0))
        shadow = Image.new("L", out.size, 0)
        ImageDraw.Draw(shadow).rounded_rectangle([pad, pad + s(8), pad + cw, pad + ch + s(8)], radius=radius, fill=150)
        shadow = shadow.filter(ImageFilter.GaussianBlur(s(12)))
        out.paste(Image.new("RGBA", out.size, (0, 0, 0, 255)), (0, 0), shadow)
        out.alpha_composite(body, (pad, pad))
        return out

    def load_media(self, slide: dict, box_w: int, box_h: int) -> list:
        """[(kind, payload, weight)]: still -> card image; gif -> [(card, seconds)]; zoom -> dict."""
        specs = [m for m in slide.get("media", []) if find_media(parse_media(m)[0])]
        if not specs:
            specs = [m for m in slide.get("fallback", []) if find_media(parse_media(m)[0])]
        items = []
        for spec in specs[:3]:
            name, focus, zoom = parse_media(spec)
            im = Image.open(find_media(name))
            paper = name.startswith("paper-")
            if getattr(im, "n_frames", 1) > 1:
                frames = []
                for k in range(im.n_frames):
                    im.seek(k)
                    secs = (im.info.get("duration") or 100) / 1000.0
                    frames.append((self.frame_card(im.convert("RGB"), box_w, box_h), secs))
                items.append(("gif", frames, 2.0))
            elif im.height / im.width > 1.3 and not paper:
                items.append(("still", self.frame_card(im, box_w, box_h), 1.0))
            else:
                items.append(("zoom", self.zoom_source(im, box_w, box_h, focus, zoom or self.ZOOM), 1.0))
        return items

    def zoom_source(self, im: Image.Image, box_w: int, box_h: int, focus: tuple, zoom: float) -> dict:
        """A desktop screenshot (or a printed page) is small on a phone, so it
        slowly zooms in (Ken Burns) towards `focus` instead of sitting still:
        everything here is built once, each frame only crops and scales."""
        s = self.s
        im = im.convert("RGB")
        pad = s(18)
        scale = min((box_w - 2 * pad) / im.width, (box_h - 2 * pad) / im.height)
        cw, ch = max(1, int(im.width * scale)), max(1, int(im.height * scale))
        sw = min(im.width, int(cw * zoom * 1.5))
        sh = max(1, int(im.height * sw / im.width))
        src = im.resize((sw, sh), Image.LANCZOS)
        canvas = Image.new("RGBA", (cw + 2 * pad, ch + 2 * pad), (0, 0, 0, 0))
        shadow = Image.new("L", canvas.size, 0)
        ImageDraw.Draw(shadow).rounded_rectangle([pad, pad + s(8), pad + cw, pad + ch + s(8)], radius=s(14), fill=150)
        canvas.paste(Image.new("RGBA", canvas.size, (0, 0, 0, 255)), (0, 0), shadow.filter(ImageFilter.GaussianBlur(s(12))))
        edge = Image.new("RGBA", (cw, ch), (0, 0, 0, 0))
        ImageDraw.Draw(edge).rounded_rectangle([0, 0, cw - 1, ch - 1], radius=s(14), outline=LINE, width=1)
        return {"src": src, "size": (cw, ch), "pad": pad, "canvas": canvas, "focus": focus, "zoom": zoom,
                "mask": rounded_mask((cw, ch), s(14)), "edge": edge, "cache": {}}

    def zoom_card(self, z: dict, amount: float) -> Image.Image:
        """The card at zoom 1 + (zoom-1)*amount, drifting from the centre to the focus point."""
        key = round(amount, 3)
        if key in z["cache"]:
            return z["cache"][key]
        cw, ch = z["size"]
        sw, sh = z["src"].size
        f = 1 + (z["zoom"] - 1) * amount
        w, h = sw / f, sh / f
        fx, fy = z["focus"]
        cx = sw * (0.5 + (fx - 0.5) * amount)
        cy = sh * (0.5 + (fy - 0.5) * amount)
        x0 = min(max(0, cx - w / 2), sw - w)
        y0 = min(max(0, cy - h / 2), sh - h)
        body = z["src"].resize((cw, ch), Image.BILINEAR, box=(x0, y0, x0 + w, y0 + h)).convert("RGBA")
        body.putalpha(z["mask"])
        body.alpha_composite(z["edge"])
        card = z["canvas"].copy()
        card.alpha_composite(body, (z["pad"], z["pad"]))
        if len(z["cache"]) < 4:
            z["cache"][key] = card
        return card

    # -- per frame --------------------------------------------------------------

    def caption_image(self, text: str) -> Image.Image:
        if text in self._cap_cache:
            return self._cap_cache[text]
        s = self.s
        f = font(s(27), True)
        tmp = ImageDraw.Draw(Image.new("RGB", (1, 1)))
        lines = wrap(tmp, text, f, self.W - s(110))
        lh = s(36)
        w = int(max(tmp.textlength(ln, font=f) for ln in lines)) + s(40)
        h = len(lines) * lh + s(24)
        im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        d.rounded_rectangle([0, 0, w - 1, h - 1], radius=s(16), fill=(0, 0, 0, 178))
        y = s(12)
        for ln in lines:
            lw = d.textlength(ln, font=f)
            d.text(((w - lw) / 2, y), ln, fill=WHITE, font=f)
            y += lh
        self._cap_cache[text] = im
        return im

    def motif(self, d: ImageDraw.ImageDraw, gt: float, color: tuple) -> None:
        """The app's own notation as a beat counter: a chart cell (fret number
        over chord symbol) that follows the backing track's chords bar by bar,
        a barline before it, and four beat dots under it, beat 1 lit white."""
        s, W = self.s, self.W
        beat = int(gt / BEAT)
        phase = (gt / BEAT) - beat
        pulse = math.exp(-phase * 3.2)
        bar = (beat // 4) % 4
        fret, chord = MOTIF[bar]
        right = W - s(28)
        cell_w = s(76)
        cx = right - cell_w / 2
        # Barline (brighter on beat 1)
        on_one = beat % 4 == 0
        bl = tuple(int(LINE[j] + (WHITE[j] - LINE[j]) * (0.6 * pulse if on_one else 0.0)) for j in range(3))
        d.line([right - cell_w - s(4), s(26), right - cell_w - s(4), s(68)], fill=bl, width=max(2, s(2)))
        # Fret number over chord symbol, as on the printed chart.
        ffret, fchord = font(s(16), True), font(s(26), True)
        tw = d.textlength(fret, font=ffret)
        d.text((cx - tw / 2, s(22)), fret, fill=DIM, font=ffret)
        cc = color if not on_one else tuple(int(color[j] + (WHITE[j] - color[j]) * 0.5 * pulse) for j in range(3))
        tw = d.textlength(chord, font=fchord)
        d.text((cx - tw / 2, s(40)), chord, fill=cc, font=fchord)
        # Four beat dots under the cell.
        r = s(4)
        for k in range(4):
            x = cx + (k - 1.5) * s(15)
            y = s(80)
            if (beat % 4) == k:
                c = WHITE if k == 0 else color
                rr = r + int(s(2) * pulse)
                fill = tuple(int(BG[j] + (c[j] - BG[j]) * (0.55 + 0.45 * pulse)) for j in range(3))
                d.ellipse([x - rr, y - rr, x + rr, y + rr], fill=fill)
            else:
                d.ellipse([x - r, y - r, x + r, y + r], fill=LINE)

    def frame(self, prep: dict, lt: float, dur: float, gt: float, caption: tuple = None,
              progress: list = None) -> Image.Image:
        """Slide `prep` at local time `lt` (of `dur`), global time `gt`."""
        s, W = self.s, self.W
        img = prep["base"].copy()
        d = ImageDraw.Draw(img)
        color = prep["color"]

        # Story-style progress bar
        if progress:
            n = len(progress)
            gap = s(6)
            seg_w = (W - s(56) - gap * (n - 1)) / n
            for k, frac in enumerate(progress):
                x0 = s(28) + k * (seg_w + gap)
                d.rounded_rectangle([x0, s(14), x0 + seg_w, s(18)], radius=s(2), fill=(69, 71, 90))
                if frac > 0:
                    d.rounded_rectangle([x0, s(14), x0 + max(s(4), seg_w * frac), s(18)], radius=s(2), fill=WHITE)

        self.motif(d, gt, color)

        # Media: items share the slide; each slides in and crossfades to the next.
        media = prep["media"]
        if media:
            vt, vb = prep["vis"]
            total_w = sum(w for _, _, w in media)
            edges, acc = [], 0.0
            for _, _, w in media:
                edges.append(acc)
                acc += dur * w / total_w
            k = max(i for i, e in enumerate(edges) if lt >= e) if lt >= 0 else 0
            for idx in ([k - 1, k] if k > 0 else [k]):
                kind, payload, w = media[idx]
                it = lt - edges[idx]
                share = dur * w / total_w
                if idx == k:
                    a = ease_out(it / 0.4)
                    rise = int(s(26) * (1 - a))
                else:
                    a = 1.0 - ease_out((lt - edges[k]) / 0.4)
                    rise = 0
                    if a <= 0:
                        continue
                if kind == "gif":
                    frames = payload
                    loop = sum(sec for _, sec in frames)
                    speed = max(1.0, loop / max(0.5, share - 0.3))
                    speed = min(speed, max(1.0, min(sec for _, sec in frames) / 0.16))
                    tt = (it * speed) % loop if it > 0 else 0
                    card = frames[-1][0]
                    for c, sec in frames:
                        if tt < sec:
                            card = c
                            break
                        tt -= sec
                elif kind == "zoom":
                    card = self.zoom_card(payload, smoothstep(it / max(0.5, share)))
                else:
                    card = payload
                x = (W - card.width) // 2
                y = vt + (vb - vt - card.height) // 2 + rise
                img.paste(with_alpha(card, a), (x, y), with_alpha(card, a))

        # Bullet cards, staggered in.
        for i, (card, y) in enumerate(prep["cards"]):
            a = ease_out((lt - 0.35 - 0.22 * i) / 0.35)
            if a <= 0:
                continue
            c = with_alpha(card, a)
            img.paste(c, (s(28), y + int(s(18) * (1 - a))), c)

        # Caption
        if caption:
            text, alpha = caption
            cim = self.caption_image(text)
            c = with_alpha(cim, alpha)
            cy = prep["cap_y"] + (self.H - prep["cap_y"] - cim.height) // 2 - s(4)
            img.paste(c, ((W - cim.width) // 2, cy), c)
        return img


# ---------------------------------------------------------------------------
# Captions
# ---------------------------------------------------------------------------

def caption_chunks(text: str, max_words: int = 7) -> list:
    words, chunks, cur = text.split(), [], []
    for w in words:
        cur.append(w)
        end = w[-1] in ".!?"
        soft = w[-1] in ",:;" and len(cur) >= 3
        if end or soft or len(cur) >= max_words:
            chunks.append(" ".join(cur))
            cur = []
    if cur:
        if chunks and len(cur) <= 2:
            chunks[-1] += " " + " ".join(cur)
        else:
            chunks.append(" ".join(cur))
    return chunks


def caption_track(text: str, start: float, length: float) -> list:
    """[(t0, t1, text)] spread over the speech by length, with pauses at punctuation."""
    chunks = caption_chunks(text)
    weights = [len(spoken(c)) + (8 if c[-1] in ".!?" else 4 if c[-1] in ",:;" else 0) for c in chunks]
    total = sum(weights) or 1
    out, t = [], start
    for c, w in zip(chunks, weights):
        d = length * w / total
        out.append((t, t + d, c))
        t += d
    return out


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def draw_poster(path: str, width: int = 1280, height: int = 720) -> None:
    """The picture at the top of the README that links to the video: the icon,
    the name, the tagline and a play button, in the app's dark colours. It has
    no version and no screenshot on it, so it never needs to be made again."""
    img = Image.new("RGB", (width, height), BG)
    d = ImageDraw.Draw(img)
    for y in range(height):                                   # a soft glow at the top
        k = max(0.0, 1.0 - y / (height * 0.75)) * 0.12
        d.line([(0, y), (width, y)], fill=tuple(int(BG[c] * (1 - k) + RED[c] * k) for c in range(3)))
    cx, cy, r = width // 2, int(height * 0.64), 74
    if os.path.exists(ICON):
        icon = Image.open(ICON).convert("RGBA").resize((120, 120), Image.LANCZOS)
        img.paste(icon, (cx - 60, 64), icon)

    def centred(text, y, f, fill):
        w = d.textlength(text, font=f)
        d.text((cx - w / 2, y), text, font=f, fill=fill)
    centred(NAME, 206, font(64, True), WHITE)
    centred(TAGLINE, 288, font(32), DIM)
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=RED)
    d.polygon([(cx - 22, cy - 36), (cx - 22, cy + 36), (cx + 40, cy)], fill=BG)
    centred("Watch the video", cy + r + 26, font(30, True), TEXT)
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    img.save(path, optimize=True)


def main() -> None:
    ap = argparse.ArgumentParser(description=f"Make {NAME}'s vertical pitch video (and stills).")
    ap.add_argument("--out", help="output .mp4 (default dist/marketing/%s.mp4)" % FILE_STEM)
    ap.add_argument("--hd", action="store_true", help="1080x1920 instead of 720x1280")
    ap.add_argument("--tts", choices=("auto", "say", "espeak", "edge", "none"), default="auto",
                    help="voice engine (auto: macOS say, else espeak-ng)")
    ap.add_argument("--voice", help="voice name (say: see --list-voices; edge: e.g. en-GB-RyanNeural)")
    ap.add_argument("--rate", type=int, default=180, help="speaking rate, words per minute (default 180)")
    ap.add_argument("--list-voices", action="store_true", help="list the voices macOS `say` has, and stop")
    ap.add_argument("--no-music", action="store_true", help="voice only, no backing groove")
    ap.add_argument("--no-captions", action="store_true", help="no burned-in captions")
    ap.add_argument("--stills", action="store_true", help="also save one PNG per slide and a PDF deck")
    ap.add_argument("--stills-only", action="store_true", help="only the PNGs and the PDF, no video")
    ap.add_argument("--only", type=int, metavar="N", help="render just slide N (1-based), for a quick check")
    ap.add_argument("--crf", type=int, help="H.264 quality, lower is better/bigger (default 24, 22 with --hd)")
    ap.add_argument("--strict-voice", action="store_true",
                    help="stop with an error when the asked voice is missing, changes half-way or a slide "
                         "gets no voice (the pitch workflow uses it, so a video without a voice is never "
                         "published)")
    ap.add_argument("--poster", metavar="PNG",
                    help="only draw the README picture that links to the video (16:9, no version on it), "
                         "and stop")
    ap.add_argument("--refresh-pages", action="store_true",
                    help="re-render tools/pitch_media/paper-*.png from the golden songs (needs pdftoppm), and stop")
    args = ap.parse_args()

    if args.refresh_pages:
        refresh_pages()
        return

    if args.poster:
        draw_poster(args.poster)
        print(f"Poster: {args.poster}")
        return

    if args.list_voices:
        voices = say_voices()
        if not voices:
            sys.exit("`say` isn't available here (it's macOS only).")
        for name, loc in voices:
            star = " *" if name in PREFERRED_SAY_VOICES else ""
            print(f"  {name:<28} {loc}{star}")
        print("\n* = picked automatically when installed (first one wins).")
        return

    slides = SLIDES
    if args.only:
        if not 1 <= args.only <= len(SLIDES):
            sys.exit(f"--only takes 1 to {len(SLIDES)}.")
        slides = [SLIDES[args.only - 1]]
    want_video = not args.stills_only
    if want_video and not shutil.which("ffmpeg"):
        sys.exit("ffmpeg is needed to make the video. On a Mac: brew install ffmpeg "
                 "(Linux: sudo apt install ffmpeg). Or use --stills-only.")

    missing = sorted({parse_media(m)[0] for sl in slides for m in sl.get("media", [])
                      if not find_media(parse_media(m)[0])})
    if missing:
        print("   ! Not in screenshots/ or tools/pitch_media/ (skipped): " + ", ".join(missing))
    version = app_version()
    if not version:
        print("   ! Couldn't read APP_VERSION from constants.py; the footer goes without it.")
    if not os.path.exists(ICON):
        print(f"   ! No app icon at {os.path.relpath(ICON, ROOT)}; the header goes without it.")

    # --only gets its own file names, so a quick check never replaces the full video or deck.
    stem = FILE_STEM + (f"_slide{args.only:02d}" if args.only else "")
    if not args.out:
        args.out = os.path.join(OUT_DIR, stem + ".mp4")
    width = 1080 if args.hd else 720
    renderer = Renderer(width, captions=not args.no_captions and want_video, version=version)
    out_path = os.path.abspath(args.out)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    t_start = time.time()

    if args.stills or args.stills_only:
        stills_dir = os.path.join(os.path.dirname(out_path), "stills")
        os.makedirs(stills_dir, exist_ok=True)
        still_r = Renderer(width, captions=False, version=renderer.version)
        pages = []
        for i, sl in enumerate(slides):
            prep = still_r.prepare(sl, i, len(slides))
            # Everything settled in, first picture showing, half zoomed.
            img = still_r.frame(prep, 4.0, 16.0, 0.0,
                                progress=[1.0 if k <= i else 0.0 for k in range(len(slides))])
            p = os.path.join(stills_dir, f"slide-{(args.only or i + 1):02d}.png")
            img.save(p)
            pages.append(img)
        pdf = os.path.join(os.path.dirname(out_path), stem + ".pdf")
        pages[0].save(pdf, save_all=True, append_images=pages[1:], resolution=144)
        print(f"Stills: {stills_dir}/  ·  deck: {pdf}")
        if args.stills_only:
            return

    with tempfile.TemporaryDirectory(prefix="snt-pitch-") as tmp:
        voice = Voice.choose(args.tts, args.voice, args.rate)
        if args.strict_voice and args.tts not in ("auto", "none") and voice.engine != args.tts:
            sys.exit(f"--strict-voice: asked for {args.tts}, got {voice.label}.")
        print(f"1/3 Voice-over with {voice.label}...")
        vo_wavs, vo_lens = [], []
        for i, sl in enumerate(slides):
            wav = os.path.join(tmp, f"vo_{i:02d}.wav")
            if voice.render(sl["vo"], wav, tmp):
                with wave.open(wav, "rb") as w:
                    vo_lens.append(w.getnframes() / float(w.getframerate()))
                vo_wavs.append(wav)
            else:
                if args.strict_voice:
                    sys.exit(f"--strict-voice: slide {i + 1} got no voice.")
                vo_lens.append(estimate_speech(sl["vo"]))
                vo_wavs.append(None)
        # edge-tts may hand over to a local voice when it fails on the first
        # slide: fine on a Mac, never in the workflow.
        if args.strict_voice and args.tts not in ("auto", "none") and voice.engine != args.tts:
            sys.exit(f"--strict-voice: {args.tts} failed and {voice.label} took over.")
        plan = plan_timing(slides, vo_lens)
        total = plan[-1]["start"] + plan[-1]["dur"]

        print(f"2/3 Backing groove ({BPM} BPM) and mix, {total:.0f} s...")
        audio = os.path.join(tmp, "soundtrack.wav")
        mix_soundtrack(total, plan, vo_wavs, not args.no_music, audio)

        caps = []
        if renderer.captions:
            for sl, p in zip(slides, plan):
                caps += caption_track(sl["vo"], p["voice_at"], p["voice_len"])

        print(f"3/3 Rendering {renderer.W}x{renderer.H} at {FPS} fps and encoding...")
        preps = [renderer.prepare(sl, i, len(slides)) for i, sl in enumerate(slides)]
        crf = args.crf if args.crf is not None else (22 if args.hd else 24)
        enc = subprocess.Popen([
            "ffmpeg", "-loglevel", "error", "-y",
            "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{renderer.W}x{renderer.H}", "-r", str(FPS), "-i", "-",
            "-i", audio,
            "-map", "0:v", "-map", "1:a",
            "-c:v", "libx264", "-preset", "medium", "-crf", str(crf), "-pix_fmt", "yuv420p",
            "-profile:v", "high", "-g", str(FPS * 2),
            "-c:a", "aac", "-b:a", "128k", "-ar", "44100",
            "-movflags", "+faststart", "-shortest", out_path,
        ], stdin=subprocess.PIPE)

        n_frames = int(math.ceil(total * FPS))
        ci, last_pct = 0, -1
        for f in range(n_frames):
            gt = f / FPS
            si = max(i for i, p in enumerate(plan) if gt >= p["start"] - 1e-9)
            p = plan[si]
            lt = gt - p["start"]
            progress = [1.0 if k < si else (min(1.0, lt / p["dur"]) if k == si else 0.0) for k in range(len(plan))]

            cap = None
            while ci < len(caps) and caps[ci][1] <= gt:
                ci += 1
            if ci < len(caps) and caps[ci][0] <= gt:
                t0, t1, text = caps[ci]
                cap = (text, min(1.0, (gt - t0) / 0.12, (t1 - gt) / 0.08 + 0.3))

            img = renderer.frame(preps[si], lt, p["dur"], gt, cap, progress)
            if si > 0 and lt < XFADE:
                q = plan[si - 1]
                prev = renderer.frame(preps[si - 1], q["dur"], q["dur"], gt, None, progress)
                img = Image.blend(prev, img, ease_out(lt / XFADE))
            enc.stdin.write(img.tobytes())

            pct = int(100 * (f + 1) / n_frames)
            if pct % 10 == 0 and pct != last_pct:
                print(f"    {pct}%", flush=True)
                last_pct = pct
        enc.stdin.close()
        if enc.wait() != 0:
            sys.exit("ffmpeg failed to encode the video (see above).")

    mb = os.path.getsize(out_path) / (1024 * 1024)
    print(f"Done in {time.time() - t_start:.0f} s: {out_path}  ({mb:.1f} MB, {total:.0f} s)")
    if mb > 15.5:
        print("   ! Over ~16 MB: WhatsApp may squeeze it hard. Try --crf 27, or send it as a document.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nStopped.")

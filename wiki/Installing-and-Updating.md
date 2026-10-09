# Installing and updating

## The packages

Each release has a package for macOS (Apple silicon and Intel), Windows
and Linux, built by the release workflow from the same sources. A package
is the native window (`webserver.py` with pywebview), with Python and
everything else it needs inside. The Tkinter desktop app (`song_writer.py`)
and the command line (`cli.py`) are not in the packages; they run from the
sources.

| File | What it is for |
|---|---|
| `…-macos-arm64.dmg`, `…-macos-x86_64.dmg` | first install on a Mac: drag the app to Applications |
| `…-macos-arm64.tar.gz`, `…-macos-x86_64.tar.gz` | what **Update and restart** downloads on a Mac |
| `…-windows-x64-setup.exe` | Windows installer, per user (no administrator), with a Start menu entry and an uninstaller |
| `…-windows-x64.zip` | the same app without an installer, and what **Update and restart** downloads on Windows |
| `…-linux-x64.tar.gz` | Linux: unpack, run `Song-Notation-Tool`; `install-desktop-entry.sh` adds it to the applications menu for your user |
| `SHA256SUMS` | the SHA-256 of every file, used to check a download |

The apps are not signed or notarised (a decision of 8 October 2026), so
the system asks once before the first start:

- macOS: **System Settings** → **Privacy & Security** → **Open Anyway**
  next to the message about Song Notation Tool. On macOS 14 and earlier,
  right-clicking the app and choosing **Open** also works. In Terminal:
  `xattr -dr com.apple.quarantine "/Applications/Song Notation Tool.app"`.
- Windows: SmartScreen shows "Windows protected your PC". Click
  **More info**, then **Run anyway**.

On Linux the package opens in your browser: the native window needs GTK or
Qt with WebKit from the system, which a package cannot carry. When the
window cannot open, the app says so and uses the browser instead. From the
sources with `pywebview` and its GTK or Qt parts installed, you get the
window.

If port 8420 is taken (a second copy of the app, or another program), the
app uses any free port instead.

## The update check

This is the only network call the app makes by itself, and only the
installed app makes it:

- When the app starts, it asks the GitHub Releases API once for the latest
  release of `giopas/song-notation-tool`
  (`https://api.github.com/repos/giopas/song-notation-tool/releases/latest`).
- The request carries the app's name and version in its User-Agent and
  nothing else: no songs, no file names, no settings.
- The answer is remembered for 12 hours, so the app does not ask again on
  every start.
- Offline, nothing happens and nothing is shown.

If the release is newer than the app, a badge appears next to the version
number. The badge and the version number both open the About card, with
the release notes and the buttons.

**Update and restart** (installed app only):

1. saves the song you have open;
2. downloads this computer's file from the release;
3. checks its SHA-256 against the release's `SHA256SUMS`, and stops if it
   does not match or the release has no `SHA256SUMS`;
4. unpacks it next to the installed app, refusing any path that would land
   outside;
5. closes the app, swaps the old version for the new one, and starts it
   again. The old version is kept until the new one has started, then
   removed.

Your songs (in the songs folder) and settings (in the config folder) live
outside the app and are never part of an update.

**Check for updates when the app starts** in the About card switches the
check off. The choice is saved in `update.json`, next to `config.json`:

```
~/Library/Application Support/Song Notation Tool/update.json   # macOS
%APPDATA%\Song Notation Tool\update.json                       # Windows
$XDG_CONFIG_HOME/song-notation-tool/update.json                # Linux
```

**Check now** in the same card asks straight away, in the installed app
and from the sources. From the sources the card says to update with
`git pull`.

## Building the packages yourself

```bash
pip install pyinstaller pywebview certifi
pyinstaller --noconfirm packaging/songnotation.spec
"dist/Song-Notation-Tool/Song-Notation-Tool" --smoke     # Linux; see the spec for the others
python packaging/make_archive.py                          # -> release/
```

`--smoke` starts the server on a free port with a temporary songs folder,
loads the page and its files, asks a local update feed, exports the
example song as a PDF, and exits with 0 when all of that worked. The
release workflow runs it on every package, and on Windows again after a
silent install.

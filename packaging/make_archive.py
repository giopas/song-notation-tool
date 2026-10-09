#!/usr/bin/env python3
"""Pack the PyInstaller output (dist/) into the release files for this system.

    python packaging/make_archive.py    # -> release/Song-Notation-Tool-<ver>-<os>-<arch>.<ext>

Run it after:  pyinstaller --noconfirm packaging/songnotation.spec

The names are the ones update.py looks for. macOS gets a .tar.gz (what
"Update and restart" downloads) and a .dmg for the first install; Windows
a .zip (for updates) and, when Inno Setup is installed, the -setup.exe
installer; Linux a .tar.gz with a script that adds the app to the
applications menu. Adapted from QLC+ Swiss Knife's packaging/make_archive.py.
"""
import os
import platform
import shutil
import subprocess
import sys
import tarfile
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, ROOT)
from constants import APP_VERSION  # noqa: E402
from update import platform_key     # noqa: E402

APP = "Song Notation Tool"
LINUX_APP = "Song-Notation-Tool"


def asset_name(ext: str, os_name: str = None, arch: str = None,
               version: str = APP_VERSION) -> str:
    o, a = (os_name, arch) if os_name else platform_key()
    return f"Song-Notation-Tool-{version}-{o}-{a}.{ext}"


def main() -> int:
    dist, out = os.path.join(ROOT, "dist"), os.path.join(ROOT, "release")
    os.makedirs(out, exist_ok=True)
    system = platform.system()
    if system == "Darwin":
        app = os.path.join(dist, f"{APP}.app")
        # A .tar.gz keeps the .app's symlinks and file modes; Python's zipfile
        # writes symlinks as text files and the updated app would not start
        # (learned in Swiss Knife 2.8.4).
        with tarfile.open(os.path.join(out, asset_name("tar.gz")), "w:gz") as t:
            t.add(app, arcname=f"{APP}.app")
        stage = os.path.join(out, "dmg-stage")
        shutil.rmtree(stage, ignore_errors=True)
        os.makedirs(stage)
        subprocess.check_call(["ditto", app, os.path.join(stage, f"{APP}.app")])
        os.symlink("/Applications", os.path.join(stage, "Applications"))
        subprocess.check_call(["hdiutil", "create", "-volname", APP, "-srcfolder", stage,
                               "-ov", "-format", "UDZO", os.path.join(out, asset_name("dmg"))])
        shutil.rmtree(stage)
    elif system == "Windows":
        folder = os.path.join(dist, APP)
        with zipfile.ZipFile(os.path.join(out, asset_name("zip")), "w",
                             zipfile.ZIP_DEFLATED) as zf:
            for d, _dirs, files in os.walk(folder):
                for f in files:
                    p = os.path.join(d, f)
                    zf.write(p, os.path.join(APP, os.path.relpath(p, folder)))
        iscc = shutil.which("ISCC") or next((c for c in (
            r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
            r"C:\Program Files\Inno Setup 6\ISCC.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"))
            if os.path.exists(c)), None)
        if iscc:                       # the installer (per user, with an uninstaller)
            subprocess.check_call([iscc, f"/DAppVersion={APP_VERSION}",
                                   os.path.join(ROOT, "packaging", "windows", "installer.iss")])
        else:
            print("Inno Setup (ISCC) not found: no installer built.")
    else:
        folder = os.path.join(dist, LINUX_APP)
        linux = os.path.join(ROOT, "packaging", "linux")
        for name in ("install-desktop-entry.sh", "song-notation-tool.desktop"):
            shutil.copy2(os.path.join(linux, name), os.path.join(folder, name))
        shutil.copy2(os.path.join(ROOT, "packaging", "icons", "icon-512.png"),
                     os.path.join(folder, "song-notation-tool.png"))
        with tarfile.open(os.path.join(out, asset_name("tar.gz")), "w:gz") as t:
            t.add(folder, arcname=LINUX_APP)
    for f in sorted(os.listdir(out)):
        print(f, os.path.getsize(os.path.join(out, f)))
    return 0


if __name__ == "__main__":
    sys.exit(main())

# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for Song Notation Tool. One-folder build of the native
# window (webserver.py with pywebview), the way QLC+ Swiss Knife is built:
#
#     pip install pyinstaller pywebview certifi
#     pyinstaller --noconfirm packaging/songnotation.spec
#
# -> dist/Song Notation Tool/ (Windows), dist/Song Notation Tool.app (macOS),
#    dist/Song-Notation-Tool/ (Linux). Run it with --smoke to check the bundle.
# The Tkinter desktop app (song_writer.py) is not part of the package; it
# stays available from the sources.
import os
import sys

ROOT = os.path.abspath(os.path.join(SPECPATH, '..'))
sys.path.insert(0, ROOT)
from constants import APP_VERSION  # noqa: E402

ICONS = os.path.join(ROOT, 'packaging', 'icons')

a = Analysis(
    [os.path.join(ROOT, 'webserver.py')],
    pathex=[ROOT],
    binaries=[],
    datas=[(os.path.join(ROOT, 'web'), 'web')],
    hiddenimports=['webview', 'certifi'],
    excludes=['tkinter', 'pytest', 'tests', 'song_writer'],
    noarchive=False,
)
pyz = PYZ(a.pure)

NAME = 'Song-Notation-Tool' if sys.platform.startswith('linux') else 'Song Notation Tool'

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name=NAME,
    console=sys.platform.startswith('linux'),
    upx=False,
    icon=(os.path.join(ICONS, 'icon.ico') if sys.platform == 'win32' else None),
)
coll = COLLECT(exe, a.binaries, a.datas, name=NAME, upx=False)

if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='Song Notation Tool.app',
        icon=os.path.join(ICONS, 'icon.icns'),
        bundle_identifier='io.github.giopas.song-notation-tool',
        info_plist={
            'CFBundleShortVersionString': APP_VERSION,
            'CFBundleVersion': APP_VERSION,
            'NSHighResolutionCapable': True,
            'LSMinimumSystemVersion': '11.0',
        },
    )

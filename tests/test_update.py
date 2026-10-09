"""The update check, the verified download and swap, the smoke test and
the packaging files. Adapted from QLC+ Swiss Knife's tests/test_update.py.
Nothing here reaches the network: SNT_UPDATE_FEED points at local files."""
import hashlib
import importlib.util
import io
import json
import os
import stat
import sys
import tarfile
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, ROOT)

import pytest

import update
from constants import APP_VERSION


@pytest.fixture(autouse=True)
def _isolated(tmp_path, monkeypatch):
    monkeypatch.setenv("SNT_UPDATE", str(tmp_path / "update.json"))
    monkeypatch.delenv("SNT_UPDATE_FEED", raising=False)


def _feed(tmp_path, tag="v99.0.0", assets=None):
    f = tmp_path / "feed.json"
    f.write_text(json.dumps({"tag_name": tag,
                             "html_url": "https://github.com/giopas/song-notation-tool/releases/tag/" + tag,
                             "body": "notes", "assets": assets or []}))
    return str(f)


def _tar(tmp_path, name, files):
    p = tmp_path / name
    with tarfile.open(p, "w:gz") as t:
        for n, data in files.items():
            ti = tarfile.TarInfo(n)
            ti.size = len(data)
            t.addfile(ti, io.BytesIO(data))
    return str(p)


def _load(rel):
    spec = importlib.util.spec_from_file_location("m_" + os.path.basename(rel)[:-3],
                                                  os.path.join(ROOT, rel))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def test_versions():
    assert update.parse_version("v0.29.0") == (0, 29, 0)
    assert update.parse_version("garbage") == ()
    assert update.is_newer("v0.30.0", "0.29.9") and not update.is_newer("0.29.0", "0.29.0")
    assert not update.is_newer("nonsense", "0.29.0")


def test_pick_asset_by_platform():
    a = [{"name": f"Song-Notation-Tool-0.29.0-{o}-{r}.{e}"} for o, r, e in
         (("macos", "arm64", "tar.gz"), ("macos", "arm64", "dmg"), ("macos", "x86_64", "tar.gz"),
          ("windows", "x64", "zip"), ("linux", "x64", "tar.gz"))] + [{"name": "SHA256SUMS"}]
    assert update.pick_asset(a, "macos", "arm64")["name"].endswith("arm64.tar.gz")   # not the .dmg
    assert update.pick_asset(a, "macos", "x86_64")["name"].endswith("x86_64.tar.gz")
    assert update.pick_asset(a, "windows", "x64")["name"].endswith("windows-x64.zip")
    assert update.pick_asset(a, "linux", "x64")["name"].endswith("linux-x64.tar.gz")
    assert update.pick_asset(a, "linux", "arm64") is None
    assert update.pick_asset([{"name": "Song-Notation-Tool-0.29.0-windows-x64-setup.exe"}],
                             "windows", "x64") is None


def test_check_newer_and_cache(tmp_path, monkeypatch):
    monkeypatch.setenv("SNT_UPDATE_FEED", _feed(tmp_path))
    r = update.check()
    assert r["newer"] and r["latest"] == "99.0.0" and r["current"] == APP_VERSION
    assert not r["frozen"] and not r["can_install"]
    monkeypatch.setenv("SNT_UPDATE_FEED", str(tmp_path / "gone.json"))
    assert update.check()["latest"] == "99.0.0"           # cached for 12 hours
    assert update.check(force=True)["error"]


def test_same_version_is_not_newer(tmp_path, monkeypatch):
    monkeypatch.setenv("SNT_UPDATE_FEED", _feed(tmp_path, "v" + APP_VERSION))
    assert update.check()["newer"] is False


def test_off_switch(tmp_path, monkeypatch):
    monkeypatch.setenv("SNT_UPDATE_FEED", _feed(tmp_path))
    update.set_enabled(False)
    r = update.check()
    assert not r["enabled"] and "newer" not in r
    update.set_enabled(True)
    assert update.check()["newer"]


def test_settings_live_next_to_config(monkeypatch):
    import userpaths
    monkeypatch.delenv("SNT_UPDATE", raising=False)
    assert update._file() == os.path.join(userpaths.config_dir(), "update.json")


def test_offline_is_calm(monkeypatch):
    monkeypatch.setenv("SNT_UPDATE_FEED", "/no/such/file.json")
    r = update.check()
    assert r["error"] and not r.get("newer")


def test_only_https_or_a_local_file():
    with pytest.raises(ValueError):
        update._read("http://example.com/feed.json")


def test_parse_sums():
    h = "a" * 64
    assert update.parse_sums(f"{h}  Song-Notation-Tool-1.zip\n{h.upper()} *dir/x.tar.gz\njunk") == {
        "Song-Notation-Tool-1.zip": h, "x.tar.gz": h}


def test_safe_extract_refuses_escape(tmp_path):
    z = tmp_path / "bad.zip"
    with zipfile.ZipFile(z, "w") as zf:
        zf.writestr("../evil.txt", "x")
    out = tmp_path / "out"
    out.mkdir()
    with pytest.raises(ValueError):
        update._safe_extract(str(z), str(out))
    with pytest.raises(ValueError):
        update._safe_extract(_tar(tmp_path, "bad.tar.gz", {"../evil.txt": b"x"}), str(out))
    assert not (tmp_path / "evil.txt").exists()


@pytest.mark.skipif(sys.platform == "win32",
                    reason="symlinks are a macOS .app concern; Windows needs a privilege for them")
def test_zip_symlinks_are_kept(tmp_path):
    z = tmp_path / "a.zip"
    with zipfile.ZipFile(z, "w") as zf:
        zf.writestr("App.app/Contents/Frameworks/lib.dylib", "x")
        info = zipfile.ZipInfo("App.app/Contents/Resources/lib.dylib")
        info.external_attr = (stat.S_IFLNK | 0o755) << 16
        zf.writestr(info, "../Frameworks/lib.dylib")
    out = tmp_path / "out"
    out.mkdir()
    update._safe_extract(str(z), str(out))
    link = out / "App.app/Contents/Resources/lib.dylib"
    assert link.is_symlink() and link.read_text() == "x"


def test_install_verifies_and_stages(tmp_path, monkeypatch):
    """A packaged install: download, SHA-256, unpack, swap script (not run here)."""
    inst = tmp_path / "apps" / "Song-Notation-Tool"
    inst.mkdir(parents=True)
    exe = inst / "Song-Notation-Tool"
    exe.write_text("old")
    monkeypatch.setattr(update, "is_frozen", lambda: True)
    monkeypatch.setattr(update.sys, "executable", str(exe))
    monkeypatch.setattr(update.platform, "system", lambda: "Linux")
    monkeypatch.setattr(update, "platform_key", lambda: ("linux", "x64"))
    ran = []
    monkeypatch.setattr(update.subprocess, "Popen", lambda *a, **k: ran.append(a))
    name = "Song-Notation-Tool-99.0.0-linux-x64.tar.gz"
    arc = _tar(tmp_path, name, {"Song-Notation-Tool/Song-Notation-Tool": b"new"})
    sums = tmp_path / "SHA256SUMS"
    sums.write_text(f"{hashlib.sha256(open(arc, 'rb').read()).hexdigest()}  {name}\n")
    monkeypatch.setenv("SNT_UPDATE_FEED", _feed(tmp_path, assets=[
        {"name": name, "browser_download_url": arc, "size": 1},
        {"name": "SHA256SUMS", "browser_download_url": str(sums)}]))
    r = update.install()
    assert r["ok"] and r["restart"], r
    assert ran and "snt-update.sh" in ran[0][0][-1]
    assert exe.read_text() == "old"                      # swapped only after quitting
    ran.clear()
    sums.write_text(f"{'0' * 64}  {name}\n")
    r = update.install(update.check(force=True))
    assert not r["ok"] and "SHA-256" in r["message"] and not ran


def test_install_from_sources_only_explains(tmp_path, monkeypatch):
    monkeypatch.setenv("SNT_UPDATE_FEED", _feed(tmp_path))
    r = update.install()
    assert not r["ok"] and "git pull" in r["message"]


def test_helper_scripts():
    _n, sh = update.helper_script(4242, "/opt/App", "/tmp/new", "Linux", "App")
    assert "kill -0 4242" in sh and "mv '/opt/App' '/opt/App.old'" in sh and "'/opt/App/App' &" in sh
    _n, mac = update.helper_script(7, "/Applications/Song Notation Tool.app", "/tmp/n", "Darwin")
    assert "xattr -dr com.apple.quarantine" in mac and "open '/Applications/Song Notation Tool.app'" in mac
    n, bat = update.helper_script(9, r"C:\Apps\SNT", r"C:\tmp\new", "Windows")
    assert n == "snt-update.bat" and "tasklist" in bat and "Song Notation Tool.exe" in bat


def test_windows_update_keeps_the_uninstaller(tmp_path):
    old, new = tmp_path / "old", tmp_path / "new"
    old.mkdir(); new.mkdir()
    (old / "unins000.exe").write_text("u"); (old / "unins000.dat").write_text("d")
    (old / "app.exe").write_text("a")
    update.keep_uninstaller(str(old), str(new))
    assert sorted(os.listdir(new)) == ["unins000.dat", "unins000.exe"]


# ── the web server ────────────────────────────────────────────────────────

def test_smoke_passes(monkeypatch):
    import webserver
    monkeypatch.setattr(webserver.Handler, "store", webserver.Handler.store)
    assert webserver.smoke() == 0


def test_update_routes(tmp_path, monkeypatch):
    import http.server, threading, urllib.request
    import webserver
    monkeypatch.setenv("SNT_UPDATE_FEED", _feed(tmp_path))
    monkeypatch.setattr(webserver.Handler, "store", webserver.SongStore(str(tmp_path / "songs")))
    monkeypatch.setattr(webserver.webbrowser, "open", lambda url: True)
    httpd = http.server.ThreadingHTTPServer(("127.0.0.1", 0), webserver.Handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{httpd.server_address[1]}"

    def post(path, body):
        req = urllib.request.Request(base + path, data=json.dumps(body).encode(),
                                     headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=10) as r:
                return r.status, json.loads(r.read())
        except urllib.error.HTTPError as e:
            return e.code, json.loads(e.read())
    try:
        with urllib.request.urlopen(base + "/api/update/check", timeout=10) as r:
            d = json.loads(r.read())
        assert d["newer"] and not d["can_install"]
        status, d = post("/api/update/install", {})
        assert status == 400 and "git pull" in d["message"]          # from sources
        status, d = post("/api/update/settings", {"check": False})
        assert d["enabled"] is False
        status, d = post("/api/update/open", {"url": "https://evil.example/"})
        assert d["url"] == "https://github.com/giopas/song-notation-tool/releases/latest"
    finally:
        httpd.shutdown()
        httpd.server_close()


def test_a_taken_port_falls_back_to_a_free_one():
    import webserver
    first = webserver._make_server("127.0.0.1")
    try:
        second = webserver._make_server("127.0.0.1")
        assert second.server_address[1] != first.server_address[1]
        second.server_close()
    finally:
        first.server_close()


# ── the packaging files ───────────────────────────────────────────────────

def test_asset_names_match_what_the_updater_looks_for():
    m = _load("packaging/make_archive.py")
    for ext, o, a in (("tar.gz", "macos", "arm64"), ("tar.gz", "macos", "x86_64"),
                      ("dmg", "macos", "arm64"), ("zip", "windows", "x64"),
                      ("tar.gz", "linux", "x64")):
        assert update._ASSET_RE.match(m.asset_name(ext, o, a))
    mk = open(os.path.join(ROOT, "packaging", "make_archive.py"), encoding="utf-8").read()
    mac = mk.split('system == "Darwin"')[1].split('elif system == "Windows"')[0]
    assert 'asset_name("tar.gz")' in mac            # the updater's macOS archive keeps symlinks


def test_spec_bundles_the_web_files_and_icons():
    spec = open(os.path.join(ROOT, "packaging", "songnotation.spec"), encoding="utf-8").read()
    for want in ("webserver.py", "'web'", "icon.icns", "icon.ico", "'webview'",
                 "io.github.giopas.song-notation-tool"):
        assert want in spec
    for icon in ("icon.icns", "icon.ico", "icon-512.png"):
        assert os.path.getsize(os.path.join(ROOT, "packaging", "icons", icon)) > 100


def test_windows_installer_is_per_user():
    iss = open(os.path.join(ROOT, "packaging", "windows", "installer.iss"), encoding="utf-8").read()
    assert "PrivilegesRequired=lowest" in iss and "windows-x64-setup" in iss
    assert "Song Notation Tool.exe" in iss


def test_linux_desktop_entry_is_filled_in_by_the_script():
    desktop = open(os.path.join(ROOT, "packaging", "linux", "song-notation-tool.desktop")).read()
    script = open(os.path.join(ROOT, "packaging", "linux", "install-desktop-entry.sh")).read()
    assert "@DIR@/Song-Notation-Tool" in desktop and "s|@DIR@|$DIR|g" in script


def test_release_workflow_when_present():
    wf = os.path.join(ROOT, ".github", "workflows", "release.yml")
    if not os.path.exists(wf):
        pytest.skip("release.yml is copied into .github/workflows by hand")
    text = open(wf, encoding="utf-8").read()
    for want in ("macos-14", "macos-15-intel", "windows-latest", "ubuntu-22.04",
                 "SHA256SUMS", "--smoke", "packaging/songnotation.spec"):
        assert want in text

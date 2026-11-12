"""The song list's collections, hide and delete in the web front end (v0.35.0).

The clicks run in the browser. These tests keep the page and the script in
step: every element the script looks up is on the page, deleting always
asks first, and the confirm dialog puts Cancel before the delete button.
"""
import os
import re

HERE = os.path.dirname(__file__)
INDEX = open(os.path.join(HERE, "..", "web", "index.html"), encoding="utf-8").read()
APP = open(os.path.join(HERE, "..", "web", "app.js"), encoding="utf-8").read()


def test_every_element_the_song_list_script_uses_is_on_the_page():
    start = APP.index("//  Rendering: song list, collections")
    end = APP.index("function escapeHtml(s) {")
    ids = set(re.findall(r'getElementById\("([\w-]+)"\)', APP[start:end]))
    assert {"song-view", "btn-view-more", "song-menu", "view-menu", "confirm-ok",
            "confirm-cancel", "name-input", "btn-empty-trash"} <= ids
    for i in ids:
        assert f'id="{i}"' in INDEX, i


def test_deleting_always_asks_first():
    for fn in ("deleteSong", "deleteCollection", "emptyTrash"):
        body = APP[APP.index(f"async function {fn}("):]
        body = body[:body.index("\n}\n")]
        assert "await confirmDialog(" in body, fn
        assert body.index("await confirmDialog(") < body.index("if (!ok) return;"), fn
    purge = APP[APP.index('API.trashOp("purge"') - 600:APP.index('API.trashOp("purge"')]
    assert "await confirmDialog(" in purge


def test_cancel_comes_first_in_the_confirm_dialog():
    start = INDEX.index('id="confirm-modal-backdrop"')
    dialog = INDEX[start:INDEX.index("</div>\n  </div>", start)]
    assert dialog.index('id="confirm-cancel"') < dialog.index('id="confirm-ok"')
    assert 'data-close="confirm-cancel"' in INDEX[start - 80:start + 120]


def test_a_hidden_song_is_not_opened_at_start():
    assert "SONGS.find((s) => !s.hidden)" in APP

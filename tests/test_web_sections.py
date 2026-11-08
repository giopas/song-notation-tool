"""Moving and copying sections in the web front end (v0.32.0).

The drag itself runs in the browser and is checked there. These tests
keep the page and the script in step: the ↑ ↓ buttons are gone, the drag
handle can be reached from the keyboard, and the "Copy to…" picker the ⧉
button opens is on the page with everything the script expects in it.
"""
import os
import re

HERE = os.path.dirname(__file__)
INDEX = open(os.path.join(HERE, "..", "web", "index.html"), encoding="utf-8").read()
APP = open(os.path.join(HERE, "..", "web", "app.js"), encoding="utf-8").read()


def _template():
    start = INDEX.index('<template id="section-template">')
    return INDEX[start:INDEX.index("</template>", start)]


def test_the_up_and_down_buttons_are_gone():
    tpl = _template()
    assert "sec-up" not in tpl and "sec-down" not in tpl
    assert "sec-up" not in APP and "sec-down" not in APP


def test_the_drag_handle_can_be_focused():
    handle = re.search(r'<span class="drag-handle"[^>]*>', _template()).group(0)
    assert 'tabindex="0"' in handle
    assert "aria-label" in handle


def test_the_copy_button_is_still_on_every_card():
    assert 'class="btn icon-btn sec-dup"' in _template()


def test_the_copy_picker_has_what_the_script_fills_in():
    start = INDEX.index('id="copy-picker"')
    picker = INDEX[start:INDEX.index("</div>\n  </div>", start)]
    for cls in ("copy-picker-title", "copy-picker-list", "copy-picker-cancel"):
        assert cls in picker, cls
    for cls in ("copy-picker-title", "copy-picker-list", "copy-picker-cancel",
                "copy-picker-row"):
        assert f".{cls}" in APP, cls


def test_the_script_wires_both():
    assert "wireDragHandle(node, sec.id)" in APP
    assert "openCopyPicker(sec.id" in APP
    # A copy can be put anywhere, and still goes right after the original
    # when no place is given.
    assert re.search(r"function duplicateSection\(id, at\)", APP)

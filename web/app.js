/* app.js — Song Notation Tool browser front end.
 *
 * Talks to webserver.py's JSON API. No build step, no framework — the
 * doc object mirrors model.py's schema exactly (same keys), so it can
 * be PUT back to the server byte-for-byte compatible with the desktop
 * app's .sng files.
 */
"use strict";

const API = {
  meta: () => fetchJSON("/api/meta"),
  songs: () => fetchJSON("/api/songs"),
  loadSong: (name) => fetchJSON(`/api/songs/${encodeURIComponent(name)}`),
  saveSong: (name, doc) => fetchJSON(`/api/songs/${encodeURIComponent(name)}`, {
    method: "PUT", body: JSON.stringify(doc),
  }),
  newSong: (fields) => fetchJSON("/api/songs", {
    method: "POST", body: JSON.stringify(fields),
  }),
  openExample: () => fetchJSON("/api/songs/example", { method: "POST" }),
  deleteSong: (name) => fetchJSON(`/api/songs/${encodeURIComponent(name)}`, { method: "DELETE" }),
  // `doc` is optional context so the server can expand =section / riff
  // references into the items they point at for the preview row.
  parseLine: (line, doc, items) => fetchJSON("/api/parse", {
    method: "POST",
    body: JSON.stringify(
      items ? { line, doc, items } : (doc ? { line, doc } : { line })),
  }),
  render: (doc, instruments) => fetchJSON("/api/render", {
    method: "POST", body: JSON.stringify({ doc, instruments }),
  }),
  splitLyrics: (doc, text) => fetchJSON("/api/lyrics/split", {
    method: "POST", body: JSON.stringify({ doc, text }),
  }),
  markedLyrics: (doc, text) => fetchJSON("/api/lyrics/marked", {
    method: "POST", body: JSON.stringify({ doc, text }),
  }),
  chordCoverage: (doc) => fetchJSON("/api/chords/coverage", {
    method: "POST", body: JSON.stringify({ doc }),
  }),
  chordShape: (text, instrument) => fetchJSON("/api/chords/shape", {
    method: "POST", body: JSON.stringify({ text, instrument }),
  }),
  quit: () => fetchJSON("/api/quit", { method: "POST" }),
};

function fetchJSON(url, opts) {
  return fetch(url, Object.assign({ headers: { "Content-Type": "application/json" } }, opts))
    .then(async (res) => {
      const body = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(body.error || `${res.status} ${res.statusText}`);
      return body;
    });
}

function debounce(fn, ms) {
  let t;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}

function toast(msg) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.classList.remove("hidden");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.classList.add("hidden"), 2200);
}

// ---------------------------------------------------------------------------
//  State
// ---------------------------------------------------------------------------
let META = { section_types: [], render_modes: {}, instruments: {}, app_version: "" };
let currentFilename = null;
let currentDoc = null;
let previewOpen = false;

// ---------------------------------------------------------------------------
//  Section helpers (client-side mirror of model.py / songmap.py)
// ---------------------------------------------------------------------------
function chartItems(section) {
  return (section.items || []).filter((it) => it.kind !== "measure");
}
function measureItems(section) {
  return (section.items || []).filter((it) => it.kind === "measure");
}
function setChartItems(section, newItems) {
  section.items = [...newItems, ...measureItems(section)];
}
function setMeasureItems(section, newItems) {
  section.items = [...chartItems(section), ...newItems];
}
function stringsForInstrument(instrument) {
  return (META.instruments || {})[instrument] || ["G", "D", "A", "E"];
}
/**
 * An empty lick for `instrument`: every string of it, `slots` positions
 * wide, all unplayed — "{G - - - - - - | D - - - - - - | ...}". Typing
 * over a dash is the whole interaction; nothing has to be deleted first.
 */
function emptyLick(instrument, slots) {
  const n = Math.max(1, slots || 6);
  const body = stringsForInstrument(instrument)
    .map((st) => st + " " + Array(n).fill("-").join(" "))
    .join(" | ");
  return "{" + body + "}";
}
/**
 * Every lick named anywhere in this song, in chart order.
 *
 * Derived rather than stored, exactly as the server derives it
 * (songmap.lick_index): the name lives on the lick item itself, so
 * renaming or deleting one needs no second list kept in step.
 */
function namedLicks() {
  const out = [];
  const walk = (items) => (items || []).forEach((it) => {
    if (it.kind === "group") walk(it.items);
    if (it.kind === "lick" && it.name && !out.includes(it.name)) out.push(it.name);
  });
  ((currentDoc && currentDoc.sections) || []).forEach((s) => walk(s.items));
  return out;
}

/** A name for the next lick in this song: Riff1, Riff2, … */
function nextLickName() {
  const taken = new Set(namedLicks().map((n) => n.toLowerCase()));
  for (let i = 1; i < 999; i += 1) {
    if (!taken.has(`riff${i}`)) return `Riff${i}`;
  }
  return "Riff";
}

function newSectionId() {
  return "section_" + Date.now().toString(36) + Math.floor(Math.random() * 1000);
}
function makeSection(name, type, instrument) {
  return {
    id: newSectionId(), name, type, instrument: instrument || defaultInstrument(),
    repeat: 1, transpose: 0, render: "chart", annotation: "",
    lyrics_text: "", print_lyrics: false, free_text: "", items: [],
  };
}
function defaultInstrument() {
  const names = Object.keys(META.instruments || {});
  return names.length ? names[0] : "Bass (4-string)";
}

// ---------------------------------------------------------------------------
//  Rendering: song list
// ---------------------------------------------------------------------------
async function refreshSongList() {
  const songs = await API.songs();
  const ul = document.getElementById("song-list");
  ul.innerHTML = "";
  for (const s of songs) {
    const li = document.createElement("li");
    li.className = s.filename === currentFilename ? "active" : "";
    li.innerHTML = `<span class="song-title">${escapeHtml(s.title || s.filename)}</span>` +
      (s.artist ? `<span class="song-artist">${escapeHtml(s.artist)}</span>` : "");
    li.addEventListener("click", () => openSong(s.filename));
    ul.appendChild(li);
  }
}

function escapeHtml(s) {
  const d = document.createElement("div");
  d.textContent = s == null ? "" : String(s);
  return d.innerHTML;
}

// ---------------------------------------------------------------------------
//  Opening / creating songs
// ---------------------------------------------------------------------------
async function openSong(filename) {
  const doc = await API.loadSong(filename);
  currentFilename = filename;
  currentDoc = doc;
  document.getElementById("start-here").classList.add("hidden");
  document.getElementById("editor").classList.remove("hidden");
  renderEditor();
  refreshSongList();
  schedulePreviewUpdate();
}

async function createSong(fields) {
  const { filename, doc } = await API.newSong(fields);
  currentFilename = filename;
  currentDoc = doc;
  document.getElementById("start-here").classList.add("hidden");
  document.getElementById("editor").classList.remove("hidden");
  renderEditor();
  refreshSongList();
  toast(`Created ${filename}`);
}

async function openExampleSong() {
  const { filename, doc } = await API.openExample();
  currentFilename = filename;
  currentDoc = doc;
  document.getElementById("start-here").classList.add("hidden");
  document.getElementById("editor").classList.remove("hidden");
  renderEditor();
  refreshSongList();
}

// ---------------------------------------------------------------------------
//  Editor rendering
// ---------------------------------------------------------------------------
function renderEditor() {
  document.getElementById("meta-title").value = currentDoc.meta.title || "";
  document.getElementById("meta-artist").value = currentDoc.meta.artist || "";
  document.getElementById("meta-key").value = currentDoc.meta.key || "";
  document.getElementById("meta-time").value = currentDoc.meta.time || "";
  document.getElementById("meta-bpm").value = currentDoc.meta.bpm || "";
  document.getElementById("meta-layout").value = currentDoc.section_layout || "banner";
  document.getElementById("meta-color").value = currentDoc.color_mode || "color";
  document.getElementById("meta-scale").value = String(currentDoc.pdf_scale || "fit");
  document.getElementById("meta-columns").value = String(currentDoc.pdf_columns || "auto");
  document.getElementById("meta-lyrics-layout").value =
    currentDoc.lyrics_layout || "beside";
  document.getElementById("meta-chord-sheet").value =
    currentDoc.chord_sheet || "none";
  document.getElementById("meta-lick-refs").value =
    currentDoc.lick_refs || "tab";

  const list = document.getElementById("section-list");
  list.innerHTML = "";
  currentDoc.sections.forEach((sec, idx) => {
    list.appendChild(buildSectionCard(sec, idx));
  });
}

function buildSectionCard(sec, idx) {
  const tpl = document.getElementById("section-template");
  const node = tpl.content.firstElementChild.cloneNode(true);
  node.dataset.id = sec.id;
  node.dataset.type = sec.type;

  const typeSel = node.querySelector(".sec-type");
  META.section_types.forEach((t) => {
    const opt = document.createElement("option");
    opt.value = t; opt.textContent = t;
    if (t === sec.type) opt.selected = true;
    typeSel.appendChild(opt);
  });

  const instrSel = node.querySelector(".sec-instrument");
  Object.keys(META.instruments).forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name; opt.textContent = name;
    if (name === sec.instrument) opt.selected = true;
    instrSel.appendChild(opt);
  });

  node.querySelector(".sec-name").value = sec.name || "";
  node.querySelector(".sec-repeat").value = sec.repeat || 1;

  const renderSel = node.querySelector(".sec-render");
  renderSel.value = sec.render || "chart";

  const lineInput = node.querySelector(".sec-chart-line");
  lineInput.value = sec.chart_line || "";

  node.querySelector(".sec-free-text").value = sec.free_text || "";
  node.querySelector(".sec-annotation").value = sec.annotation || "";

  // Populate tab beats dropdown
  const beatsSelect = node.querySelector(".tab-beats-select");
  (META.tab_beats_options || [8, 16, 32, 64]).forEach((b) => {
    const opt = document.createElement("option");
    opt.value = b; opt.textContent = b;
    if (b === (META.tab_beats_default || 8)) opt.selected = true;
    beatsSelect.appendChild(opt);
  });

  wireSectionEvents(node, sec, idx);
  updateSectionPreview(node, sec);
  updateTabGridVisibility(node, sec);

  // A freshly built card has no "last render" to preserve, so the
  // leave-it-alone path in updateSectionPreview would leave the chart row
  // blank until the first keystroke. Fetch the render once, in the
  // background — failures are silent, since this is only a display nicety
  // and the typed line is already on screen.
  if (lineInput.value.trim() && sec.render !== "free") {
    API.parseLine(lineInput.value, currentDoc)
      .then((result) => { if (result.ok) updateSectionPreview(node, sec, result); })
      .catch(() => {});
  }

  return node;
}

function wireSectionEvents(node, sec, idx) {
  const byId = () => currentDoc.sections.find((s) => s.id === sec.id);

  node.querySelector(".sec-name").addEventListener("input", (e) => {
    byId().name = e.target.value;
    refreshReferencingPreviews(sec.id);
    schedulePreviewUpdate();
  });
  node.querySelector(".sec-type").addEventListener("change", (e) => {
    byId().type = e.target.value;
    node.dataset.type = e.target.value;
    schedulePreviewUpdate();
  });
  node.querySelector(".sec-instrument").addEventListener("change", (e) => {
    byId().instrument = e.target.value;
    schedulePreviewUpdate();
  });
  node.querySelector(".sec-repeat").addEventListener("change", (e) => {
    byId().repeat = Math.max(1, parseInt(e.target.value, 10) || 1);
    schedulePreviewUpdate();
  });

  node.querySelector(".sec-render").addEventListener("change", (e) => {
    const s = byId();
    s.render = e.target.value;
    updateTabGridVisibility(node, s);
    refreshReferencingPreviews(sec.id);
    schedulePreviewUpdate();
  });

  node.querySelector(".tab-add-measure").addEventListener("click", () => {
    const s = byId();
    const beats = parseInt(node.querySelector(".tab-beats-select").value, 10) || 8;
    const strings = stringsForInstrument(s.instrument);
    const stringMap = {};
    strings.forEach((st) => { stringMap[st] = ""; });
    const measures = measureItems(s);
    measures.push({ kind: "measure", beats, strings: stringMap });
    setMeasureItems(s, measures);
    rebuildTabGrid(node, s);
    schedulePreviewUpdate();
  });

  node.querySelector(".tab-beats-select").addEventListener("change", () => {
    // Just changes default for new measures; existing measures keep their beats
  });

  const lineInput = node.querySelector(".sec-chart-line");
  const onLineChange = debounce(async () => {
    const line = lineInput.value;
    const s = byId();
    if (!line.trim()) {
      setChartItems(s, []);
      updateSectionPreview(node, s, { items: [], rendered: [] });
      refreshReferencingPreviews(s.id);
      schedulePreviewUpdate();
      return;
    }
    try {
      const result = await API.parseLine(line, currentDoc);
      if (result.ok) {
        setChartItems(s, result.items);
        s.chart_line = result.unparsed;
        if (result.annotation) {
          // grammar.unparse() drops the quoted text from the line, so show
          // it in the field instead of losing track of it.
          s.annotation = result.annotation;
          node.querySelector(".sec-annotation").value = result.annotation;
        }
        updateSectionPreview(node, s, result);
        refreshReferencingPreviews(s.id);
      } else {
        updateSectionPreview(node, s, { error: result.error });
      }
    } catch (err) {
      updateSectionPreview(node, s, { error: String(err) });
    }
    schedulePreviewUpdate();
  }, 150);
  lineInput.addEventListener("input", onLineChange);

  node.querySelector(".sec-annotation").addEventListener("input", (e) => {
    byId().annotation = e.target.value;
    schedulePreviewUpdate();
  });

  // Free-text mode: stored verbatim, never parsed. No debounce on the
  // model write (it's a plain string assignment); only the export preview
  // recompute is throttled.
  node.querySelector(".sec-free-text").addEventListener("input", (e) => {
    byId().free_text = e.target.value;
    refreshReferencingPreviews(sec.id);
    schedulePreviewUpdate();
  });

  // Quick-insert palette. Inserting through the same input event the user
  // would have produced by typing means the existing debounced parse,
  // error display and preview refresh all run unchanged.
  renderLickChips(node);
  node.querySelector(".sec-insert-strip").addEventListener("click", (e) => {
    const btn = e.target.closest(".ins-btn");
    if (!btn) return;
    if (btn.dataset.insLickRef) {
      insertAtCursor(node.querySelector(".sec-chart-line"),
                      `{${btn.dataset.insLickRef}}`, 0);
      return;
    }
    if (btn.dataset.insLickNamed) {
      // A named lick is worth nothing if naming it is a chore, so the
      // name is filled in for you — Riff1, Riff2 — and the caret lands
      // on the first fret position, same as the unnamed button.
      const body = emptyLick(byId().instrument,
                              parseInt(btn.dataset.insLickNamed, 10) || 6);
      const text = "{" + nextLickName() + " = " + body.slice(1);
      insertAtCursor(node.querySelector(".sec-chart-line"), text,
                      text.length - text.indexOf("-"));
      return;
    }
    // The lick button builds its insert from the section's own instrument
    // rather than dropping a fixed example: an empty grid with every
    // string already named and room for six positions is a thing to fill
    // in, where "{G 5 7 5 | D - - 3}" is a thing to delete first.
    const text = btn.dataset.insLick
      ? emptyLick(byId().instrument, parseInt(btn.dataset.insLick, 10) || 6)
      : btn.dataset.ins;
    const back = btn.dataset.insLick
      ? text.length - text.indexOf("-")   // caret on the first position
      : parseInt(btn.dataset.caret, 10) || 0;
    insertAtCursor(node.querySelector(".sec-chart-line"), text, back);
  });

  node.querySelector(".sec-up").addEventListener("click", () => moveSection(sec.id, -1));
  node.querySelector(".sec-down").addEventListener("click", () => moveSection(sec.id, 1));
  node.querySelector(".sec-dup").addEventListener("click", () => duplicateSection(sec.id));
  node.querySelector(".sec-del").addEventListener("click", () => deleteSection(sec.id));
}

/**
 * The licks this song has already named, one button each, so recalling
 * one is a click rather than remembering what you called it. The button
 * inserts the *reference* — `{Riff1}` — not the notes: that is the whole
 * point of having named it.
 */
function renderLickChips(node) {
  const box = node.querySelector(".sec-lick-refs");
  if (!box) return;
  box.innerHTML = "";
  namedLicks().forEach((name) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "ins-btn lick-ref-btn";
    b.dataset.insLickRef = name;
    b.textContent = `{${name}}`;
    b.title = `Play the lick named ${name} here. Add x3 after it to repeat it.`;
    box.appendChild(b);
  });
}

/**
 * Drop `text` into `input` at the caret, space-separated from whatever is
 * already there (the chart-line grammar is whitespace-delimited, so a
 * missing space is the one way an insert could corrupt a valid line).
 * `caretBack` leaves the caret that many characters from the end of the
 * inserted text — 4 for "[ ]x2" puts it between the brackets, 1 for a
 * quote pair puts it between the quotes.
 */
function insertAtCursor(input, text, caretBack) {
  if (!input || !text) return;
  const value = input.value;
  const start = input.selectionStart == null ? value.length : input.selectionStart;
  const end = input.selectionEnd == null ? value.length : input.selectionEnd;
  const before = value.slice(0, start);
  const after = value.slice(end);

  const needsLeadingSpace = before.length > 0 && !/\s$/.test(before);
  const needsTrailingSpace = after.length > 0 && !/^\s/.test(after);
  const insert = (needsLeadingSpace ? " " : "") + text + (needsTrailingSpace ? " " : "");

  input.value = before + insert + after;
  const caret = before.length + insert.length -
    (needsTrailingSpace ? 1 : 0) - (caretBack || 0);
  input.focus();
  input.setSelectionRange(caret, caret);
  input.dispatchEvent(new Event("input", { bubbles: true }));
}

/**
 * The rows after the first fret/symbol pair: a lick's string lines, and
 * the continuation lines of a section broken with //.
 */
function setExtraPreviewRows(node, rows) {
  const box = node.querySelector(".sec-preview");
  [...box.querySelectorAll(".sec-preview-more")].forEach((el) => el.remove());
  rows.forEach((row) => {
    const div = document.createElement("div");
    const role = (row && row.role) || "";
    div.className = "sec-preview-more" + (role ? " sec-row-" + role : "");
    paintRow(div, row);
    box.appendChild(div);
  });
}

/**
 * Write one rendered row into `el`, colouring the stretches the server
 * tagged — a rest, so far, which prints grey here exactly as it does on
 * paper. Plain text when there's nothing tagged, so the common row costs
 * no DOM.
 */
function paintRow(el, row) {
  const text = (row && row.text !== undefined) ? row.text : (row || "");
  const spans = (row && row.spans) || [];
  if (!spans.length) { el.textContent = text; return; }
  el.textContent = "";
  let pos = 0;
  [...spans].sort((a, b) => a[0] - b[0]).forEach(([start, end, role]) => {
    const from = Math.max(start, pos), to = Math.min(end, text.length);
    if (to <= from) return;
    if (from > pos) el.appendChild(document.createTextNode(text.slice(pos, from)));
    const mark = document.createElement("span");
    mark.className = "tok-" + role;
    mark.textContent = text.slice(from, to);
    el.appendChild(mark);
    pos = to;
  });
  if (pos < text.length) el.appendChild(document.createTextNode(text.slice(pos)));
}

/**
 * A card whose line is a reference shows the *target's* content, so it goes
 * stale the moment the target is edited — and nothing in the DOM records
 * that dependency. After any edit that could change what a reference
 * resolves to, re-parse every other card that holds one.
 */
function sectionHasRef(sec) {
  const scan = (items) => (items || []).some(
    (it) => it.kind === "section_ref" || it.kind === "block_ref" ||
            (it.kind === "group" && scan(it.items)));
  return scan(sec.items);
}

const refreshReferencingPreviews = debounce((excludeId) => {
  if (!currentDoc) return;
  currentDoc.sections.forEach((sec) => {
    if (sec.id === excludeId || sec.render === "free" || !sectionHasRef(sec)) return;
    const node = document.querySelector(`.section-card[data-id="${sec.id}"]`);
    if (!node) return;
    const line = node.querySelector(".sec-chart-line").value;
    if (!line.trim()) return;
    // Render from the stored items, not the typed line: the items hold the
    // target's id, so a reference survives the target being renamed.
    API.parseLine(line, currentDoc, sec.items || [])
      .then((result) => {
        if (!result.ok) return;
        updateSectionPreview(node, sec, result);
        // A reference is spelled with the target's *name*, so renaming the
        // target also restyles the line that points at it. Never while the
        // user is typing in that field.
        const input = node.querySelector(".sec-chart-line");
        if (result.unparsed && result.unparsed !== input.value &&
            document.activeElement !== input) {
          input.value = result.unparsed;
          sec.chart_line = result.unparsed;
        }
      })
      .catch(() => {});
  });
}, 200);

function updateSectionPreview(node, sec, parseResult) {
  const fretRow = node.querySelector(".sec-preview-fret");
  const symRow = node.querySelector(".sec-preview-sym");
  const errEl = node.querySelector(".sec-error");
  const emptyEl = node.querySelector(".sec-empty-hint");

  const items = (parseResult && parseResult.items) || chartItems(sec);
  const rendered = (parseResult && parseResult.rendered) ||
    (items.length ? null : []); // null = "not recomputed, leave last render"

  if (parseResult && parseResult.error) {
    errEl.textContent = parseResult.error;
    errEl.classList.remove("hidden");
    return; // leave the last valid render on screen — never clear it
  }
  errEl.classList.add("hidden");

  if (rendered && rendered.length) {
    // A row is 1..N lines now: an optional fret row, the symbol row, and
    // one line per string for any lick. The first two elements are reused
    // so the fret/symbol colouring stays; anything beyond is appended.
    const hasFret = parseResult ? !!parseResult.fret_row : rendered.length > 1;
    const roles = (parseResult && parseResult.roles) || [];
    const spans = (parseResult && parseResult.spans) || [];
    const row = (i) => ({ text: rendered[i] || "", role: roles[i], spans: spans[i] });
    fretRow.textContent = hasFret ? rendered[0] : "";
    paintRow(symRow, hasFret ? row(1) : row(0));
    setExtraPreviewRows(node, rendered.slice(hasFret ? 2 : 1)
      .map((_, i) => row(i + (hasFret ? 2 : 1))));
    emptyEl.classList.add("hidden");
  } else if (!items.length && !measureItems(sec).length) {
    fretRow.textContent = "";
    symRow.textContent = "";
    setExtraPreviewRows(node, []);
    emptyEl.classList.remove("hidden");
  } else if (rendered) {
    fretRow.textContent = "";
    symRow.textContent = "";
    setExtraPreviewRows(node, []);
    emptyEl.classList.add("hidden");
  }
}

// ---------------------------------------------------------------------------
//  Tab grid — render-mode toggle and measure editing
// ---------------------------------------------------------------------------
function updateTabGridVisibility(node, sec) {
  const mode = sec.render || "chart";
  const lineRow = node.querySelector(".sec-line-row");
  const freeText = node.querySelector(".sec-free-text");
  const chartPreview = node.querySelector(".sec-preview");
  const tabGrid = node.querySelector(".sec-tab-grid");
  const emptyHint = node.querySelector(".sec-empty-hint");
  const errEl = node.querySelector(".sec-error");

  // "free" is the odd one out: the whole chart apparatus — typed line,
  // insert palette, rendered preview, tab grid, parse errors — is replaced
  // by one textarea. The section's items are left alone rather than
  // cleared, so switching back to Chart brings the chart line back intact.
  if (mode === "free") {
    lineRow.classList.add("hidden");
    chartPreview.classList.add("hidden");
    tabGrid.classList.add("hidden");
    emptyHint.classList.add("hidden");
    errEl.classList.add("hidden");
    freeText.classList.remove("hidden");
    return;
  }
  freeText.classList.add("hidden");

  if (mode === "chart") {
    lineRow.classList.remove("hidden");
    chartPreview.classList.remove("hidden");
    tabGrid.classList.add("hidden");
  } else if (mode === "tab") {
    lineRow.classList.add("hidden");
    chartPreview.classList.add("hidden");
    tabGrid.classList.remove("hidden");
    emptyHint.classList.add("hidden");
    rebuildTabGrid(node, sec);
  } else { // "both"
    lineRow.classList.remove("hidden");
    chartPreview.classList.remove("hidden");
    tabGrid.classList.remove("hidden");
    rebuildTabGrid(node, sec);
  }
}

function rebuildTabGrid(node, sec) {
  const body = node.querySelector(".tab-grid-body");
  const measures = measureItems(sec);
  const strings = stringsForInstrument(sec.instrument);

  if (!measures.length) {
    body.innerHTML = '<p style="color:var(--lbl-gray);font-size:12px;font-style:italic;margin:4px 0;">' +
      'No measures yet — click "+ Measure" to add one.</p>';
    return;
  }

  const table = document.createElement("table");
  table.className = "tab-grid-table";

  // Header row: measure numbers with delete buttons
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  headerRow.appendChild(document.createElement("th")); // empty corner
  measures.forEach((m, mi) => {
    const th = document.createElement("th");
    th.colSpan = m.beats || 8;
    const hdr = document.createElement("span");
    hdr.className = "tab-measure-header";
    hdr.innerHTML = 'M' + (mi + 1) + ' ';
    const delBtn = document.createElement("button");
    delBtn.className = "tab-del-measure";
    delBtn.title = "Delete measure " + (mi + 1);
    delBtn.textContent = "×";
    delBtn.addEventListener("click", () => {
      const allMeasures = measureItems(sec);
      allMeasures.splice(mi, 1);
      setMeasureItems(sec, allMeasures);
      rebuildTabGrid(node, sec);
      schedulePreviewUpdate();
    });
    hdr.appendChild(delBtn);
    th.appendChild(hdr);
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  // One row per string
  const tbody = document.createElement("tbody");
  strings.forEach((st) => {
    const tr = document.createElement("tr");
    const label = document.createElement("td");
    label.className = "tab-string-label";
    label.textContent = st + "|";
    tr.appendChild(label);

    measures.forEach((m, mi) => {
      const beats = m.beats || 8;
      const raw = (m.strings || {})[st] || "";
      const tokens = raw.split(/\s+/).filter(Boolean);
      while (tokens.length < beats) tokens.push("-");

      for (let bi = 0; bi < beats; bi++) {
        const td = document.createElement("td");
        const input = document.createElement("input");
        input.type = "text";
        input.className = "tab-cell";
        input.value = tokens[bi] || "-";
        input.maxLength = 3;
        input.addEventListener("input", () => {
          commitTabCell(sec, mi, st, beats, node);
        });
        input.addEventListener("focus", () => input.select());
        // Arrow key navigation between cells
        input.addEventListener("keydown", (e) => {
          if (e.key === "ArrowRight" || (e.key === "Tab" && !e.shiftKey)) {
            const next = td.nextElementSibling;
            if (next && next.querySelector("input")) {
              e.preventDefault();
              next.querySelector("input").focus();
            }
          } else if (e.key === "ArrowLeft" || (e.key === "Tab" && e.shiftKey)) {
            const prev = td.previousElementSibling;
            if (prev && prev.querySelector("input")) {
              e.preventDefault();
              prev.querySelector("input").focus();
            }
          } else if (e.key === "ArrowDown") {
            const rowIdx = Array.from(tr.parentElement.children).indexOf(tr);
            const colIdx = Array.from(tr.children).indexOf(td);
            const nextRow = tr.parentElement.children[rowIdx + 1];
            if (nextRow && nextRow.children[colIdx]) {
              const nextInput = nextRow.children[colIdx].querySelector("input");
              if (nextInput) { e.preventDefault(); nextInput.focus(); }
            }
          } else if (e.key === "ArrowUp") {
            const rowIdx = Array.from(tr.parentElement.children).indexOf(tr);
            const colIdx = Array.from(tr.children).indexOf(td);
            const prevRow = tr.parentElement.children[rowIdx - 1];
            if (prevRow && prevRow.children[colIdx]) {
              const prevInput = prevRow.children[colIdx].querySelector("input");
              if (prevInput) { e.preventDefault(); prevInput.focus(); }
            }
          }
        });
        td.appendChild(input);
        tr.appendChild(td);
      }
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  body.innerHTML = "";
  body.appendChild(table);
}

function commitTabCell(sec, measureIdx, stringName, beats, node) {
  const body = node.querySelector(".tab-grid-body");
  const table = body.querySelector("table");
  if (!table) return;

  const strings = stringsForInstrument(sec.instrument);
  const measures = measureItems(sec);
  const m = measures[measureIdx];
  if (!m) return;

  // Find the row for this string
  const strIdx = strings.indexOf(stringName);
  const tbody = table.querySelector("tbody");
  const row = tbody.children[strIdx];
  if (!row) return;

  // Collect all beat values for this string in this measure
  // Cells start at column 1 (col 0 is the string label)
  // We need to figure out the column offset for this measure
  let colOffset = 1; // skip string label
  for (let i = 0; i < measureIdx; i++) {
    colOffset += (measures[i].beats || 8);
  }

  const tokens = [];
  for (let bi = 0; bi < beats; bi++) {
    const cell = row.children[colOffset + bi];
    const input = cell ? cell.querySelector("input") : null;
    let val = (input ? input.value.trim() : "-") || "-";
    tokens.push(val);
  }

  if (!m.strings) m.strings = {};
  m.strings[stringName] = tokens.join(" ");
  setMeasureItems(sec, measures);
  schedulePreviewUpdate();
}

// ---------------------------------------------------------------------------
//  Section CRUD
// ---------------------------------------------------------------------------
function addSection() {
  const sec = makeSection("New section", META.section_types[1] || "Verse");
  currentDoc.sections.push(sec);
  renderEditor();
  schedulePreviewUpdate();
}

function deleteSection(id) {
  const sec = currentDoc.sections.find((s) => s.id === id);
  if (!confirm(`Delete section "${sec ? sec.name : id}"?`)) return;
  currentDoc.sections = currentDoc.sections.filter((s) => s.id !== id);
  renderEditor();
  schedulePreviewUpdate();
}

function moveSection(id, delta) {
  const secs = currentDoc.sections;
  const i = secs.findIndex((s) => s.id === id);
  if (i < 0) return;
  const j = Math.max(0, Math.min(secs.length - 1, i + delta));
  if (j === i) return;
  const [item] = secs.splice(i, 1);
  secs.splice(j, 0, item);
  renderEditor();
  schedulePreviewUpdate();
}

/**
 * How a reference to `sec` is spelled on screen (mirrors
 * songmap.display_ref_key on the server): the target's name, with spaces
 * and dashes as underscores, when that's a bare identifier and no other
 * section answers to it — otherwise the id, which is always unambiguous.
 * Only ever a display form; the document stores the id.
 */
function refKeyFor(sec) {
  const candidate = (sec.name || "").trim().replace(/[ -]/g, "_");
  if (!/^[A-Za-z][A-Za-z0-9_]*$/.test(candidate)) return sec.id;
  const norm = (s) => (s || "").trim().toLowerCase().replace(/[ -]/g, "_");
  const clashes = currentDoc.sections.filter(
    (s) => norm(s.name) === norm(candidate)).length;
  return clashes === 1 ? candidate : sec.id;
}

function duplicateSection(id) {
  const i = currentDoc.sections.findIndex((s) => s.id === id);
  if (i < 0) return;
  const source = currentDoc.sections[i];
  const newSec = {
    id: newSectionId(), name: source.name + " (ref)", type: source.type,
    instrument: source.instrument, repeat: 1, transpose: 0,
    render: source.render === "free" ? "chart" : source.render, annotation: "",
    free_text: "",
    // The id is what's stored (rename-safe); the name is what's shown.
    items: [{ kind: "section_ref", section: source.id, repeat: 1, all: false, transpose: 0 }],
    chart_line: `=${refKeyFor(source)}`,
  };
  currentDoc.sections.splice(i + 1, 0, newSec);
  renderEditor();
  schedulePreviewUpdate();
}

// ---------------------------------------------------------------------------
//  Meta form
// ---------------------------------------------------------------------------
function wireMetaForm() {
  const fields = ["title", "artist", "key", "time", "bpm"];
  fields.forEach((f) => {
    document.getElementById(`meta-${f}`).addEventListener("input", (e) => {
      if (!currentDoc) return;
      currentDoc.meta[f] = e.target.value;
      schedulePreviewUpdate();
    });
  });
}

// ---------------------------------------------------------------------------
//  Save / export
// ---------------------------------------------------------------------------
async function saveCurrent() {
  if (!currentDoc || !currentFilename) return;
  await API.saveSong(currentFilename, currentDoc);
  toast(`Saved ${currentFilename}`);
  refreshSongList();
}

async function saveAndClose() {
  if (currentDoc && currentFilename) {
    try {
      await saveCurrent();
    } catch (err) {
      if (!confirm(`Couldn't save ${currentFilename}: ${err}\n\nClose anyway without saving?`)) {
        return;
      }
    }
  }
  showClosedOverlay("Closing…");
  try {
    await API.quit();
  } catch (err) {
    // The server may tear down before this response makes it back —
    // that's expected, not a failure.
  }
  setTimeout(() => showClosedOverlay("Closed — you can close this window now."), 600);
}

function showClosedOverlay(message) {
  let overlay = document.getElementById("closed-overlay");
  if (!overlay) {
    overlay = document.createElement("div");
    overlay.id = "closed-overlay";
    document.body.appendChild(overlay);
  }
  overlay.textContent = message;
  overlay.classList.remove("hidden");
}

/**
 * Print through the OS. In the native window the PDF is built and handed
 * to the system viewer, which brings up the real print panel — printer,
 * paper size, scaling, page range — rather than a silent job. In a
 * browser tab the same PDF opens in a new tab, where Cmd/Ctrl+P does the
 * same thing.
 */
async function printCurrent() {
  const api = nativeApi();
  if (api && api.print_document) {
    const res = await api.print_document(currentDoc, "portrait");
    toast(res && res.ok
      ? "Opened in your PDF viewer — print from there"
      : `Print failed: ${(res && res.error) || "unknown error"}`);
    return;
  }
  const res = await fetch("/api/export.pdf", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ doc: currentDoc, orient: "portrait" }),
  });
  if (!res.ok) {
    toast("Print failed: could not build the PDF");
    return;
  }
  const url = URL.createObjectURL(await res.blob());
  const win = window.open(url, "_blank");
  if (!win) {
    toast("Allow pop-ups to print, or export the PDF and print that");
  } else {
    toast("Opened in a new tab — print from there");
  }
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

/** True in the native window, where pywebview's js_api bridge exists. */
function nativeApi() {
  return (window.pywebview && window.pywebview.api) || null;
}

async function exportCurrent(kind) {
  if (!currentDoc) return;
  const fmt = kind === "txt" ? "txt" : "pdf";
  const orient = kind === "pdf-landscape" ? "landscape" : "portrait";

  if (kind === "print") return printCurrent();

  // Native window: ask the OS where to put it, so the file lands somewhere
  // the user picked and we can say exactly where. A browser tab has no such
  // dialog available to us — there the download below is the right answer,
  // and the browser's own settings decide the folder.
  const api = nativeApi();
  if (api && api.export_document) {
    const res = await api.export_document(currentDoc, fmt, orient);
    if (res && res.ok) {
      toast(`Saved to ${res.path}`);
    } else if (res && res.cancelled) {
      // user closed the panel — say nothing
    } else {
      toast(`Export failed: ${(res && res.error) || "unknown error"}`);
    }
    return;
  }

  let url, filename, body;
  if (kind === "txt") {
    url = "/api/export.txt";
    body = { doc: currentDoc };
  } else {
    url = "/api/export.pdf";
    body = { doc: currentDoc, orient: kind === "pdf-landscape" ? "landscape" : "portrait" };
  }
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    toast(`Export failed: ${err.error || res.statusText}`);
    return;
  }
  const disposition = res.headers.get("Content-Disposition") || "";
  const m = /filename="([^"]+)"/.exec(disposition);
  filename = m ? m[1] : `export.${kind === "txt" ? "txt" : "pdf"}`;
  const blob = await res.blob();
  const dlUrl = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = dlUrl; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(dlUrl);
  toast(`Downloaded ${filename} — check your browser's downloads folder`);
}

// ---------------------------------------------------------------------------
//  Songs folder — shown in the sidebar so "where is my file?" is answerable
//  without leaving the app.
// ---------------------------------------------------------------------------
function renderSongsFolder() {
  const box = document.getElementById("songs-folder");
  if (!box) return;
  const dir = META.songs_dir || "";
  box.querySelector(".songs-folder-path").textContent = dir;
  box.querySelector(".songs-folder-path").title = dir;

  const api = nativeApi();
  const changeBtn = box.querySelector("#btn-songs-folder");
  const revealBtn = box.querySelector("#btn-reveal-songs");
  // Both actions need the native bridge; in a browser tab there's no way to
  // open a folder picker or a Finder window, so say what to do instead.
  const native = !!(api && api.choose_songs_folder);
  changeBtn.classList.toggle("hidden", !native);
  revealBtn.classList.toggle("hidden", !native);
  box.querySelector(".songs-folder-hint").classList.toggle("hidden", native);
}

async function changeSongsFolder() {
  const api = nativeApi();
  if (!api || !api.choose_songs_folder) return;
  const res = await api.choose_songs_folder();
  if (!res || !res.ok) {
    if (res && !res.cancelled) toast(`Could not change folder: ${res.error || "?"}`);
    return;
  }
  META.songs_dir = res.path;
  renderSongsFolder();
  currentFilename = null;
  currentDoc = null;
  document.getElementById("editor").classList.add("hidden");
  document.getElementById("start-here").classList.remove("hidden");
  await refreshSongList();
  toast(`Songs folder is now ${res.path}`);
}

// ---------------------------------------------------------------------------
//  Live preview pane
// ---------------------------------------------------------------------------
const schedulePreviewUpdate = debounce(async () => {
  if (!previewOpen || !currentDoc) return;
  try {
    const { lines } = await API.render(currentDoc);
    document.getElementById("preview-text").textContent = lines.join("\n");
  } catch (err) {
    document.getElementById("preview-text").textContent = `(preview error: ${err})`;
  }
}, 250);

function togglePreview(open) {
  previewOpen = open !== undefined ? open : !previewOpen;
  document.getElementById("preview-pane").classList.toggle("hidden", !previewOpen);
  if (previewOpen) schedulePreviewUpdate();
}

// ---------------------------------------------------------------------------
//  New-song modal
// ---------------------------------------------------------------------------
function openNewSongModal() {
  ["title", "artist", "key", "bpm"].forEach((f) => (document.getElementById(`new-${f}`).value = ""));
  document.getElementById("new-time").value = "4/4";
  document.getElementById("modal-backdrop").classList.remove("hidden");
}
function closeNewSongModal() {
  document.getElementById("modal-backdrop").classList.add("hidden");
}

// ---------------------------------------------------------------------------
//  Transpose modal — whole song or one section, mirrors the desktop app's
//  Transpose dialog. Non-destructive: sets doc.transpose / section.transpose,
//  same fields the desktop app and render-time engine already use.
// ---------------------------------------------------------------------------
function scopeOptionsHtml(marked) {
  const songOpt = `<option value="__song__">Whole song</option>`;
  const sectionOpts = currentDoc.sections
    .map((s) => {
      // In the Lyrics dialog the point of the list is *which sections still
      // need words*, so each one says whether it has any. Transpose doesn't
      // care, and asks for the plain names.
      const dot = marked && (s.lyrics_text || "").trim() ? "• " : "";
      const type = s.type ? `  (${s.type})` : "";
      return `<option value="${s.id}">${dot}${escapeHtml(s.name || s.id)}${escapeHtml(type)}</option>`;
    })
    .join("");
  return songOpt + sectionOpts;
}

function openTransposeModal() {
  if (!currentDoc) { toast("Open a song first"); return; }
  document.getElementById("transpose-scope").innerHTML = scopeOptionsHtml();
  document.getElementById("transpose-amount").value = "0";
  document.getElementById("transpose-modal-backdrop").classList.remove("hidden");
}
function closeTransposeModal() {
  document.getElementById("transpose-modal-backdrop").classList.add("hidden");
}
function applyTranspose() {
  const scope = document.getElementById("transpose-scope").value;
  const n = parseInt(document.getElementById("transpose-amount").value, 10);
  if (!Number.isInteger(n)) { toast("Enter a whole number of semitones."); return; }
  if (n === 0) { closeTransposeModal(); return; }
  if (scope === "__song__") {
    currentDoc.transpose = (currentDoc.transpose || 0) + n;
  } else {
    const sec = currentDoc.sections.find((s) => s.id === scope);
    if (sec) sec.transpose = (sec.transpose || 0) + n;
  }
  closeTransposeModal();
  schedulePreviewUpdate();
  const label = scope === "__song__" ? "whole song" : "section";
  toast(`Transposed ${label} by ${n > 0 ? "+" : ""}${n} semitone(s) — remember to Save`);
}

// ---------------------------------------------------------------------------
//  Lyrics modal — stored non-destructively on the document or a section as
//  `lyrics_text`, never parsed or aligned to the chart. Printed in the
//  TXT/PDF export and the Preview pane only when its own `print_lyrics`
//  flag is on (off by default). "Search online" opens a browser search in
//  a new tab; it never fetches or auto-pastes lyrics in, by design
//  (copyright + accuracy).
// ---------------------------------------------------------------------------
function lyricsTarget() {
  const scope = document.getElementById("lyrics-scope").value;
  if (scope === "__song__") return currentDoc;
  return currentDoc.sections.find((s) => s.id === scope);
}
// The song's sections as a row of chips inside the Lyrics dialog. The Scope
// dropdown can only show you where you are; this shows the whole song at
// once, with a filled dot on every section that already has words — so the
// gaps are visible without opening each section in turn, and getting to one
// is a click rather than a hunt through a list.
function renderLyricsChips() {
  const box = document.getElementById("lyrics-chips");
  const scope = document.getElementById("lyrics-scope").value;
  const marking = scope === "__song__";
  box.innerHTML = "";
  const chip = (value, label, hasLyrics) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "lyrics-chip" + (value === scope ? " active" : "")
      + (hasLyrics ? " has-lyrics" : "");
    const dot = document.createElement("span");
    dot.className = "chip-dot";
    b.appendChild(dot);
    b.appendChild(document.createTextNode(label));
    // On the whole-song sheet the chips mark the sheet up; on a single
    // section's words there is nothing to mark, so they navigate as
    // before.
    if (marking && value !== "__song__") {
      b.classList.add("draggable");
      b.draggable = true;
      b.title = `Drag into the sheet (or click) to mark where ${label}'s words start`;
      // The browser inserts dropped plain text at the drop caret on its
      // own; all this has to do is say what the text is.
      b.addEventListener("dragstart", (e) => {
        e.dataTransfer.setData("text/plain", markerFor(label) + "\n");
        e.dataTransfer.effectAllowed = "copy";
      });
      b.addEventListener("click", () => insertMarkerAtCursor(label));
    } else {
      b.title = hasLyrics ? "Has lyrics — click to edit"
        : "No lyrics yet — click to add";
      b.addEventListener("click", () => {
        document.getElementById("lyrics-scope").value = value;
        loadLyricsForScope();
      });
    }
    box.appendChild(b);
  };
  chip("__song__", "Whole song", !!(currentDoc.lyrics_text || "").trim());
  currentDoc.sections.forEach((sec) => {
    chip(sec.id, sec.name || sec.id, !!(sec.lyrics_text || "").trim());
  });
}

// ---------------------------------------------------------------------------
//  Section markers — the sheet carries its own structure
//
//  Splitting a sheet on blank lines and then picking "block 3" from a
//  dropdown asks you to hold the shape of the song in your head while you
//  read a list of first lines. Writing the shape into the sheet instead
//  says the same thing where you can see it, survives re-editing the
//  words, and leaves the sheet self-describing next time you open it.
//
//  The markers are structure, not words: they're drawn in colour by the
//  overlay behind the textarea, and stripped before anything prints.
// ---------------------------------------------------------------------------
const MARKER_RE = /^[ \t]*={2,}[ \t]*(.+?)[ \t]*={2,}[ \t]*$/;

function markerFor(name) {
  return `=== ${String(name || "").trim()} ===`;
}

function lyricsBox() {
  return document.getElementById("lyrics-textarea");
}

// Typing writes back to whatever the Scope box is pointing at. Debounced
// because the chip rail re-renders from it, not because the assignment
// costs anything.
const writeLyricsBack = debounce(() => {
  const t = lyricsTarget();
  if (t) t.lyrics_text = lyricsBox().value;
  renderLyricsChips();
}, 200);

/** Drop a marker on its own line at the caret, and keep the overlay in
 *  step. A marker mid-line would be a marker no parser can see, so it
 *  always lands at the start of a line of its own. */
function insertMarkerAtCursor(name) {
  const box = lyricsBox();
  const value = box.value;
  let at = box.selectionStart == null ? value.length : box.selectionStart;
  // Snap to the start of the line the caret is on.
  while (at > 0 && value[at - 1] !== "\n") at -= 1;
  const before = value.slice(0, at);
  const after = value.slice(at);
  const text = markerFor(name) + "\n";
  box.value = before + text + after;
  const caret = before.length + text.length;
  box.focus();
  box.setSelectionRange(caret, caret);
  box.dispatchEvent(new Event("input", { bubbles: true }));
}

function hasMarkers(text) {
  return (text || "").split("\n").some((ln) => MARKER_RE.test(ln));
}

/** Paint the sheet into the overlay behind the textarea, markers in
 *  colour. A textarea can't style its own content and a contenteditable
 *  breaks paste, undo and drag-and-drop — which are the three things this
 *  box is for — so the text is drawn twice and the top copy is
 *  transparent. */
function refreshLyricsOverlay() {
  const overlay = document.getElementById("lyrics-overlay");
  const box = lyricsBox();
  if (!overlay || !box) return;
  overlay.innerHTML = (box.value || "").split("\n").map((ln) => {
    const esc = escapeHtml(ln) || "&nbsp;";
    return MARKER_RE.test(ln) ? `<span class="lyric-marker">${esc}</span>` : esc;
  }).join("\n");
  overlay.scrollTop = box.scrollTop;
  overlay.scrollLeft = box.scrollLeft;
}

/** Hand every marked block to the section it names. Sections the sheet
 *  says nothing about are left exactly as they are — silence about a
 *  section is not the same as saying it has no words. */
async function applyLyricMarkers() {
  const text = lyricsBox().value;
  if (!hasMarkers(text)) {
    toast("No === Section === markers in the sheet yet — drag a section in.");
    return;
  }
  currentDoc.lyrics_text = text;
  let info;
  try {
    info = await API.markedLyrics(currentDoc, text);
  } catch (err) { toast(String(err)); return; }

  const unmatched = info.unmatched || [];
  if (unmatched.length) {
    const ok = confirm(
      `The sheet names ${unmatched.length} section(s) this song doesn't have `
      + `yet:\n\n  ${unmatched.join("\n  ")}\n\nCreate them?`);
    if (ok) unmatched.forEach((name) => {
      currentDoc.sections.push(makeSection(name, guessSectionType(name)));
    });
  }

  const print = document.getElementById("lyrics-print").checked;
  const segments = (info.segments || []);
  // Applied on the client so the document in the browser stays the one
  // source of truth — the server only ever told us where the blocks go.
  let assigned = 0;
  const seen = new Set();
  segments.forEach((seg) => {
    const name = (seg.name || "").trim();
    if (!name) return;
    const sec = findSectionByName(name);
    if (!sec) return;
    const body = (seg.text || "").replace(/^\n+|\n+$/g, "");
    if (seen.has(sec.id) && body.trim()) {
      sec.lyrics_text = `${(sec.lyrics_text || "").replace(/\n+$/, "")}\n\n${body}`;
    } else {
      sec.lyrics_text = body;
      seen.add(sec.id);
    }
    sec.print_lyrics = print;
    if (body.trim()) assigned += 1;
  });
  // The sheet stays — it's the source these blocks came from, and the
  // reason a section added next week can still be given one. It just
  // stops printing, so the same words don't land on the chart twice.
  if (assigned) currentDoc.print_lyrics = false;

  renderEditor();
  schedulePreviewUpdate();
  document.getElementById("lyrics-scope").innerHTML = scopeOptionsHtml(true);
  document.getElementById("lyrics-scope").value = "__song__";
  loadLyricsForScope();
  toast(`Lyrics assigned to ${assigned} section(s) — remember to Save`);
}

/** The section a marker name refers to: case, spaces, dashes and
 *  underscores all the same thing, matching lyrics.py's rule so the
 *  browser and the server never disagree about where a block goes. */
function findSectionByName(name) {
  const key = (s) => String(s || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const k = key(name);
  return currentDoc.sections.find((s) => key(s.name) === k || key(s.id) === k);
}

/** A type for a section created from a marker name, so "Chorus 2" lands
 *  as a Chorus rather than as whatever the dropdown defaults to. */
function guessSectionType(name) {
  const n = String(name || "").toLowerCase();
  const hit = (META.section_types || []).find(
    (t) => n.startsWith(t.toLowerCase()));
  return hit || (META.section_types[1] || "Verse");
}

/** Create a section from inside the Lyrics dialog and mark the sheet for
 *  it in one go — the section appears in the editor behind the dialog. */
function addSectionFromLyrics() {
  const name = (prompt("Name for the new section (e.g. Verse 3)") || "").trim();
  if (!name) return;
  if (findSectionByName(name)) {
    toast(`There's already a section called ${name}`);
  } else {
    currentDoc.sections.push(makeSection(name, guessSectionType(name)));
    renderEditor();
    schedulePreviewUpdate();
  }
  document.getElementById("lyrics-scope").innerHTML = scopeOptionsHtml(true);
  document.getElementById("lyrics-scope").value = "__song__";
  loadLyricsForScope();
  insertMarkerAtCursor(name);
}

function loadLyricsForScope() {
  const t = lyricsTarget();
  document.getElementById("lyrics-textarea").value = (t && t.lyrics_text) || "";
  document.getElementById("lyrics-print").checked = !!(t && t.print_lyrics);
  const isSongScope = document.getElementById("lyrics-scope").value === "__song__";
  ["lyrics-split", "lyrics-apply-markers", "lyrics-add-section"].forEach((id) => {
    document.getElementById(id).classList.toggle("hidden", !isSongScope);
  });
  document.getElementById("lyrics-marker-hint")
    .classList.toggle("hidden", !isSongScope);
  refreshLyricsOverlay();
  renderLyricsChips();
}
function openLyricsModal() {
  if (!currentDoc) { toast("Open a song first"); return; }
  document.getElementById("lyrics-scope").innerHTML = scopeOptionsHtml(true);
  document.getElementById("lyrics-scope").value = "__song__";
  loadLyricsForScope();
  document.getElementById("lyrics-modal-backdrop").classList.remove("hidden");
}
function closeLyricsModal() {
  document.getElementById("lyrics-modal-backdrop").classList.add("hidden");
}

// ---------------------------------------------------------------------------
//  Chord shapes
//
//  A chart says which chords are played; it doesn't say how to hold one,
//  and for most of a set it doesn't need to. But there is always the
//  voicing you had to work out, or the open shape the song wants rather
//  than the barre your hands default to — and writing that on the back of
//  the sheet is what this is. Printed once, at one end of the chart,
//  rather than repeated beside every chord symbol.
// ---------------------------------------------------------------------------
function chordList() {
  if (!currentDoc.chords) currentDoc.chords = [];
  return currentDoc.chords;
}

function openChordsModal() {
  if (!currentDoc) { toast("Open a song first"); return; }
  document.getElementById("chords-where").value =
    currentDoc.chord_sheet || "none";
  renderChordRows();
  document.getElementById("chords-modal-backdrop").classList.remove("hidden");
}

function closeChordsModal() {
  document.getElementById("chords-modal-backdrop").classList.add("hidden");
  schedulePreviewUpdate();
}

function renderChordRows() {
  const box = document.getElementById("chords-rows");
  box.innerHTML = "";
  const chords = chordList();
  if (!chords.length) {
    const empty = document.createElement("p");
    empty.className = "modal-hint";
    empty.textContent = "No shapes yet — \u201c+ Chord\u201d adds one.";
    box.appendChild(empty);
  }
  chords.forEach((chord, i) => box.appendChild(chordRow(chord, i)));
  renderChordPreview();
}

function chordRow(chord, index) {
  const row = document.createElement("div");
  row.className = "chord-row";

  const name = document.createElement("input");
  name.type = "text";
  name.className = "chord-name";
  name.placeholder = "C";
  name.value = chord.name || "";
  name.title = "What you call this chord on the chart line";
  name.addEventListener("input", () => {
    chord.name = name.value;
    renderChordPreview();
  });

  const instrument = document.createElement("select");
  instrument.className = "chord-instrument";
  Object.keys(META.instruments || {}).forEach((key) => {
    const opt = document.createElement("option");
    opt.value = key; opt.textContent = key;
    instrument.appendChild(opt);
  });
  instrument.value = chord.instrument || defaultInstrument();
  instrument.title = "How many strings the shape has";

  const shape = document.createElement("input");
  shape.type = "text";
  shape.className = "chord-shape";
  shape.placeholder = "x32010";
  shape.value = chord.shape_text || "";
  shape.title = "One fret per string, lowest string first. x means don't sound it.";

  const err = document.createElement("span");
  err.className = "chord-error";

  // With the song transposed, a row is the shape as typed and prints as
  // something else — say what, and how it was re-voiced.
  const printsAs = document.createElement("span");
  printsAs.className = "chord-prints-as";
  printsAs.dataset.index = String(index);

  const readShape = async () => {
    chord.instrument = instrument.value;
    chord.shape_text = shape.value;
    if (!shape.value.trim()) { chord.frets = []; err.textContent = ""; renderChordPreview(); return; }
    let res;
    try {
      res = await API.chordShape(shape.value, instrument.value);
    } catch (e) { err.textContent = String(e); return; }
    if (!res.ok) {
      err.textContent = res.error;
      chord.frets = [];
    } else {
      err.textContent = "";
      chord.frets = res.frets;
    }
    renderChordPreview();
  };
  shape.addEventListener("input", debounce(readShape, 250));
  instrument.addEventListener("change", readShape);

  const note = document.createElement("input");
  note.type = "text";
  note.className = "chord-note";
  note.placeholder = "note (optional)";
  note.value = chord.note || "";
  note.title = "A word printed under the diagram — \u201cbarre\u201d, \u201cthumb\u201d, whatever you need to remember";
  note.addEventListener("input", () => { chord.note = note.value; renderChordPreview(); });

  const del = document.createElement("button");
  del.type = "button";
  del.className = "btn icon-btn";
  del.textContent = "\u{1F5D1}";
  del.title = "Remove this shape";
  del.addEventListener("click", () => {
    chordList().splice(index, 1);
    renderChordRows();
  });

  [name, instrument, shape, note, del, printsAs, err].forEach((el) => row.appendChild(el));
  return row;
}

/** The shapes as they will print, drawn from the same diagram code the
 *  export uses — asking the server to render it is how the dialog and the
 *  paper stay the same thing. */
/** Say, under the rows, which chords the chart plays with no shape and
 *  which shapes it never plays. Reported, not enforced: plenty of chords
 *  need no diagram. */
// Debounced: the preview redraws on every keystroke in a chord's name,
// and the coverage only needs to catch up once the typing stops.
const renderChordCoverage = debounce(() => renderChordCoverageNow(), 250);
function renderChordCoverageNow() {
  const line = document.getElementById("chords-coverage");
  if (!line) return;
  API.chordCoverage(currentDoc).then((cov) => {
    const n = cov.transpose || 0;
    const head = document.getElementById("chords-transposed");
    head.classList.toggle("hidden", !n);
    head.textContent = n
      ? `The song is transposed ${n > 0 ? "+" : ""}${n}. Shapes are stored as `
        + "you typed them and print moved with it: an open shape goes to the "
        + "open shape of the new chord, or its barre if there isn't one; a "
        + "barre slides."
      : "";
    const how = { open: "open shape", barre: "barre", moved: "moved",
                  capo: "slid, open strings too" };
    document.querySelectorAll("#chords-rows .chord-prints-as").forEach((el) => {
      const p = (cov.printed || [])[Number(el.dataset.index)];
      el.textContent = n && p && p.name && p.voicing !== "as typed"
        ? `prints as ${p.name} · ${how[p.voicing] || p.voicing}` : "";
    });
    const bits = [];
    if ((cov.missing || []).length) bits.push(`No shape yet: ${cov.missing.join(", ")}`);
    if ((cov.unused || []).length) bits.push(`Not played in this song: ${cov.unused.join(", ")}`);
    line.textContent = bits.join("  ·  ");
  }).catch(() => { line.textContent = ""; });
}

/** A row for every chord the chart plays that has no shape yet. The
 *  names come from the printed chart, so a song transposed up a tone
 *  asks for D, not the C you typed. */
async function addChordsFromChart() {
  let cov;
  try { cov = await API.chordCoverage(currentDoc); }
  catch (e) { toast(String(e)); return; }
  const missing = cov.missing || [];
  if (!missing.length) {
    toast((cov.used || []).length ? "Every chord on the chart already has a shape"
      : "No chord symbols on the chart yet");
    return;
  }
  const guitar = Object.keys(META.instruments || {})
    .find((k) => k.toLowerCase().startsWith("guitar")) || defaultInstrument();
  // Stored un-transposed, so each row prints as the chord the chart
  // shows: on a song moved up a tone, the chart's D is stored as C.
  const stored = cov.missing_stored || missing;
  stored.forEach((name) => chordList().push(
    { name, instrument: guitar, frets: [], shape_text: "", note: "" }));
  if ((currentDoc.chord_sheet || "none") === "none") {
    currentDoc.chord_sheet = "end";
    document.getElementById("chords-where").value = "end";
    const sel = document.getElementById("meta-chord-sheet");
    if (sel) sel.value = "end";
  }
  renderChordRows();
  toast((cov.transpose
    ? `Added ${missing.length} chord(s), as they read before the transpose — type a shape for each`
    : `Added ${missing.length} chord(s) — type a shape for each`));
}

function renderChordPreview() {
  renderChordCoverage();
  const box = document.getElementById("chords-preview");
  const chords = chordList().filter((c) => (c.name || "").trim() && (c.frets || []).length);
  if (!chords.length) { box.textContent = ""; return; }
  API.render({ meta: { title: "" }, sections: [], chords,
                chord_sheet: "end", format: 2 })
    .then((res) => {
      const lines = res.lines || [];
      const start = lines.findIndex((ln) => ln.trim() === "CHORDS");
      if (start < 0) { box.textContent = ""; return; }
      const body = [];
      for (let i = start + 2; i < lines.length; i += 1) {
        if (lines[i].startsWith("=")) break;
        body.push(lines[i]);
      }
      box.textContent = body.join("\n").replace(/^\n+|\n+$/g, "");
    })
    .catch(() => { box.textContent = ""; });
}

function addChord() {
  chordList().push({ name: "", instrument: defaultInstrument(), frets: [],
                      shape_text: "", note: "" });
  if ((currentDoc.chord_sheet || "none") === "none") {
    // Adding a shape is the whole reason to print one; asking for it a
    // second time in a dropdown is a step with no decision in it.
    currentDoc.chord_sheet = "end";
    document.getElementById("chords-where").value = "end";
    const sel = document.getElementById("meta-chord-sheet");
    if (sel) sel.value = "end";
  }
  renderChordRows();
  const rows = document.querySelectorAll("#chords-rows .chord-name");
  if (rows.length) rows[rows.length - 1].focus();
}

// Split the whole-song lyric sheet across the sections.
//
// The old version of this assigned block N to section N and hoped: a song
// that opens on an instrumental intro had its first verse land on the
// intro and every block after it one section out of place, and a chorus
// played three times only ever reached the first of them. So the split is
// now a proposal you correct — one row per section, one block per row,
// and the same block free to go to as many sections as sing it.
//
// The server does the splitting and the guessing (lyrics.py), so the
// desktop app's dialog makes exactly the same proposal from the same text.
let splitBlocks = [];

const KEEP = "keep";   // a section holding words that came from elsewhere

function blockLabel(block, i) {
  const first = block.split("\n")[0].trim();
  const rest = block.split("\n").length - 1;
  const tail = rest ? ` … (+${rest} line${rest > 1 ? "s" : ""})` : "";
  return `${i + 1} · ${first.slice(0, 46)}${first.length > 46 ? "…" : ""}${tail}`;
}

function splitRow(sec, choice) {
  const row = document.createElement("div");
  row.className = "split-row";

  const name = document.createElement("div");
  name.className = "split-row-name";
  name.textContent = sec.name || sec.id;
  const type = document.createElement("span");
  type.className = "split-row-type";
  type.textContent = sec.type || "";
  name.appendChild(type);

  const sel = document.createElement("select");
  sel.dataset.sectionId = sec.id;
  const opt = (value, label) => {
    const o = document.createElement("option");
    o.value = value; o.textContent = label;
    sel.appendChild(o);
  };
  opt("", "— no lyrics —");
  // Words this section already holds that aren't in the sheet came from
  // somewhere else — a hand-typed line, an older split. Offer to leave
  // them alone rather than quietly overwriting them.
  const held = (sec.lyrics_text || "").trim();
  const heldIsBlock = splitBlocks.some((b) => b.trim() === held);
  if (held && !heldIsBlock) opt(KEEP, "— keep what's here —");
  splitBlocks.forEach((b, i) => opt(String(i), blockLabel(b, i)));

  sel.value = choice === null || choice === undefined ? "" : String(choice);
  row.appendChild(name);
  row.appendChild(sel);
  return row;
}

async function splitLyricsIntoSections() {
  const text = document.getElementById("lyrics-textarea").value;
  if (!text.trim()) { toast("Nothing to split — paste some lyrics first."); return; }
  if (!currentDoc.sections.length) { toast("Add a section first"); return; }
  // Whatever is in the box is the sheet being split, typed or pasted a
  // moment ago and not yet written back to the document.
  currentDoc.lyrics_text = text;

  let result;
  try {
    result = await API.splitLyrics(currentDoc, text);
  } catch (err) {
    toast(String(err));
    return;
  }
  splitBlocks = result.blocks || [];
  if (!splitBlocks.length) { toast("Nothing to split — paste some lyrics first."); return; }

  // What a section already holds wins over a fresh guess: re-opening this
  // shouldn't propose undoing the corrections made last time.
  const current = result.current || [];
  const suggested = result.suggested || [];
  const anyCurrent = current.some((c) => c !== null && c !== undefined);
  const rows = document.getElementById("lyrics-split-rows");
  rows.innerHTML = "";
  currentDoc.sections.forEach((sec, i) => {
    let choice = anyCurrent ? current[i] : suggested[i];
    if ((choice === null || choice === undefined) && (sec.lyrics_text || "").trim()
        && !splitBlocks.some((b) => b.trim() === (sec.lyrics_text || "").trim())) {
      choice = KEEP;
    }
    rows.appendChild(splitRow(sec, choice));
  });
  document.getElementById("lyrics-split-print").checked =
    currentDoc.sections.some((s) => s.print_lyrics) || !!currentDoc.print_lyrics;
  document.getElementById("lyrics-split-backdrop").classList.remove("hidden");
}

function closeSplitModal() {
  document.getElementById("lyrics-split-backdrop").classList.add("hidden");
}

function applyLyricsSplit() {
  const print = document.getElementById("lyrics-split-print").checked;
  let assigned = 0;
  [...document.querySelectorAll("#lyrics-split-rows select")].forEach((sel) => {
    const sec = currentDoc.sections.find((s) => s.id === sel.dataset.sectionId);
    if (!sec) return;
    if (sel.value === KEEP) return;
    if (sel.value === "") {
      sec.lyrics_text = "";
      sec.print_lyrics = false;
      return;
    }
    sec.lyrics_text = splitBlocks[Number(sel.value)] || "";
    sec.print_lyrics = print;
    assigned += 1;
  });
  // The sheet stays — it's the source these blocks came from, and the
  // reason a section added next week can still be given one. It just
  // stops printing, so the same words don't land on the chart twice.
  if (assigned) currentDoc.print_lyrics = false;

  closeSplitModal();
  renderEditor();
  schedulePreviewUpdate();
  document.getElementById("lyrics-scope").innerHTML = scopeOptionsHtml(true);
  document.getElementById("lyrics-scope").value = "__song__";
  loadLyricsForScope();
  toast(`Lyrics assigned to ${assigned} section(s) — remember to Save`);
}

function openSearchLyricsModal() {
  const meta = (currentDoc && currentDoc.meta) || {};
  document.getElementById("lyrics-search-title").value = meta.title || "";
  document.getElementById("lyrics-search-artist").value = meta.artist || "";
  document.getElementById("lyrics-search-modal-backdrop").classList.remove("hidden");
}
function closeSearchLyricsModal() {
  document.getElementById("lyrics-search-modal-backdrop").classList.add("hidden");
}
function runLyricsSearch() {
  const title = document.getElementById("lyrics-search-title").value.trim();
  const artist = document.getElementById("lyrics-search-artist").value.trim();
  const query = [artist, title, "lyrics"].filter((p) => p).join(" ") || "song lyrics";
  // DuckDuckGo rather than Google — no account/consent wall, friendlier
  // default for an open-source tool.
  const url = "https://duckduckgo.com/?q=" + encodeURIComponent(query);
  openInBrowser(url);
  closeSearchLyricsModal();
}

// In the native app window (pywebview), window.open()/target="_blank" is a
// no-op — there's no browser-tab concept inside that webview to open it
// in. webserver.py exposes window.pywebview.api.open_url(), which hands
// the URL to the OS's real default browser instead; plain browser mode
// (no pywebview) doesn't have window.pywebview at all, so window.open()
// is used there.
function openInBrowser(url) {
  if (window.pywebview && window.pywebview.api && window.pywebview.api.open_url) {
    window.pywebview.api.open_url(url).then((ok) => {
      if (!ok) toast("Couldn't open the browser for that search.");
    }).catch(() => toast("Couldn't open the browser for that search."));
    return;
  }
  const win = window.open(url, "_blank", "noopener");
  if (!win) {
    toast("Your browser blocked the new tab — allow pop-ups for this page, or copy the link.");
  }
}
function importLyricsFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    const textarea = document.getElementById("lyrics-textarea");
    if (textarea.value.trim() && !confirm("Replace the current lyrics text with the file's contents?")) {
      return;
    }
    textarea.value = String(reader.result || "");
    const t = lyricsTarget();
    if (t) t.lyrics_text = textarea.value;
  };
  reader.readAsText(file);
}

// ---------------------------------------------------------------------------
//  Wire-up
// ---------------------------------------------------------------------------
async function init() {
  META = await API.meta();
  document.getElementById("app-version").textContent = `v${META.app_version}`;
  wireMetaForm();

  // Layout, Colour and Size are all properties of the *song*, not of this
  // machine — they travel with the .sng so a chart prints the same way
  // wherever it's opened.
  [["meta-layout", "section_layout", META.section_layouts, "banner"],
   ["meta-color", "color_mode", META.color_modes, "color"],
   ["meta-scale", "pdf_scale", META.pdf_scales, "fit"],
   ["meta-columns", "pdf_columns", META.pdf_columns, "auto"],
   ["meta-lyrics-layout", "lyrics_layout", META.lyrics_layouts, "beside"],
   ["meta-chord-sheet", "chord_sheet", META.chord_sheets, "none"],
   ["meta-lick-refs", "lick_refs", META.lick_refs, "tab"],
  ].forEach(([id, key, labels, fallback]) => {
    const sel = document.getElementById(id);
    Object.entries(labels || {}).forEach(([value, label]) => {
      const opt = document.createElement("option");
      opt.value = value; opt.textContent = label;
      sel.appendChild(opt);
    });
    if (!sel.options.length) {
      const opt = document.createElement("option");
      opt.value = fallback; opt.textContent = fallback;
      sel.appendChild(opt);
    }
    sel.addEventListener("change", (e) => {
      if (!currentDoc) return;
      currentDoc[key] = e.target.value;
      schedulePreviewUpdate();
    });
  });

  renderSongsFolder();
  document.getElementById("btn-songs-folder").addEventListener("click",
    () => changeSongsFolder().catch((e) => toast(String(e))));
  document.getElementById("btn-reveal-songs").addEventListener("click", () => {
    const api = nativeApi();
    if (api && api.reveal) api.reveal(META.songs_dir || "");
  });

  document.getElementById("btn-help").addEventListener("click", () => {
    document.getElementById("help-strip").classList.toggle("hidden");
  });
  document.getElementById("btn-notation-ref").addEventListener("click", () => {
    document.getElementById("notation-ref-backdrop").classList.remove("hidden");
  });
  document.getElementById("notation-ref-close").addEventListener("click", () => {
    document.getElementById("notation-ref-backdrop").classList.add("hidden");
  });
  document.getElementById("notation-ref-backdrop").addEventListener("click", (e) => {
    if (e.target === e.currentTarget) e.target.classList.add("hidden");
  });
  document.getElementById("btn-preview-toggle").addEventListener("click", () => togglePreview());
  document.getElementById("btn-preview-close").addEventListener("click", () => togglePreview(false));
  document.getElementById("btn-save").addEventListener("click", () => saveCurrent().catch((e) => toast(String(e))));
  document.getElementById("btn-close").addEventListener("click", () => saveAndClose().catch((e) => toast(String(e))));
  document.getElementById("btn-add-section").addEventListener("click", addSection);

  document.getElementById("btn-export").addEventListener("click", () => {
    document.getElementById("export-menu").classList.toggle("hidden");
  });
  document.getElementById("export-menu").addEventListener("click", (e) => {
    const fmt = e.target.dataset.fmt;
    if (!fmt) return;
    document.getElementById("export-menu").classList.add("hidden");
    exportCurrent(fmt).catch((err) => toast(String(err)));
  });
  document.addEventListener("click", (e) => {
    if (!e.target.closest("#topbar-right .dropdown")) {
      document.getElementById("export-menu").classList.add("hidden");
    }
  });

  document.getElementById("btn-transpose").addEventListener("click", openTransposeModal);
  document.getElementById("transpose-cancel").addEventListener("click", closeTransposeModal);
  document.getElementById("transpose-apply").addEventListener("click", applyTranspose);

  document.getElementById("btn-lyrics").addEventListener("click", openLyricsModal);
  document.getElementById("btn-chords").addEventListener("click", openChordsModal);
  document.getElementById("chords-add").addEventListener("click", addChord);
  document.getElementById("chords-from-chart").addEventListener("click",
    () => addChordsFromChart().catch((e) => toast(String(e))));
  document.getElementById("chords-close").addEventListener("click", closeChordsModal);
  document.getElementById("chords-where").addEventListener("change", (e) => {
    if (!currentDoc) return;
    currentDoc.chord_sheet = e.target.value;
    const sel = document.getElementById("meta-chord-sheet");
    if (sel) sel.value = e.target.value;
    schedulePreviewUpdate();
  });
  Object.entries(META.chord_sheets || {}).forEach(([value, label]) => {
    const opt = document.createElement("option");
    opt.value = value; opt.textContent = label;
    document.getElementById("chords-where").appendChild(opt);
  });
  document.getElementById("lyrics-scope").addEventListener("change", loadLyricsForScope);
  document.getElementById("lyrics-textarea").addEventListener("input", () => {
    refreshLyricsOverlay();
    writeLyricsBack();
  });
  // A drop lands as an input event on some browsers and not others; the
  // overlay has to follow either way or the markers paint in the wrong
  // place until the next keystroke.
  ["drop", "scroll", "keyup", "click"].forEach((ev) => {
    document.getElementById("lyrics-textarea")
      .addEventListener(ev, () => setTimeout(refreshLyricsOverlay, 0));
  });
  document.getElementById("lyrics-print").addEventListener("change", (e) => {
    const t = lyricsTarget();
    if (t) t.print_lyrics = e.target.checked;
    schedulePreviewUpdate();
  });
  document.getElementById("lyrics-split").addEventListener("click",
    () => splitLyricsIntoSections().catch((e) => toast(String(e))));
  document.getElementById("lyrics-apply-markers").addEventListener("click",
    () => applyLyricMarkers().catch((e) => toast(String(e))));
  document.getElementById("lyrics-add-section").addEventListener("click",
    addSectionFromLyrics);
  document.getElementById("lyrics-split-cancel").addEventListener("click", closeSplitModal);
  document.getElementById("lyrics-split-apply").addEventListener("click", applyLyricsSplit);
  document.getElementById("lyrics-import").addEventListener("click", () =>
    document.getElementById("lyrics-file-input").click());
  document.getElementById("lyrics-file-input").addEventListener("change", (e) => {
    importLyricsFile(e.target.files && e.target.files[0]);
    e.target.value = "";
  });
  document.getElementById("lyrics-search").addEventListener("click", openSearchLyricsModal);
  document.getElementById("lyrics-close").addEventListener("click", closeLyricsModal);

  document.getElementById("lyrics-search-cancel").addEventListener("click", closeSearchLyricsModal);
  document.getElementById("lyrics-search-go").addEventListener("click", runLyricsSearch);
  ["lyrics-search-title", "lyrics-search-artist"].forEach((id) => {
    document.getElementById(id).addEventListener("keydown", (e) => {
      if (e.key === "Enter") runLyricsSearch();
    });
  });

  document.getElementById("btn-new-song").addEventListener("click", openNewSongModal);
  document.getElementById("start-new").addEventListener("click", openNewSongModal);
  document.getElementById("new-song-cancel").addEventListener("click", closeNewSongModal);
  document.getElementById("new-song-create").addEventListener("click", async () => {
    const fields = {
      title: document.getElementById("new-title").value.trim() || "Untitled Song",
      artist: document.getElementById("new-artist").value.trim(),
      key: document.getElementById("new-key").value.trim(),
      time: document.getElementById("new-time").value.trim() || "4/4",
      bpm: document.getElementById("new-bpm").value.trim(),
    };
    try {
      await createSong(fields);
      closeNewSongModal();
    } catch (err) {
      toast(String(err));
    }
  });

  document.getElementById("btn-open-example").addEventListener("click", () =>
    openExampleSong().catch((e) => toast(String(e))));
  document.getElementById("start-example").addEventListener("click", () =>
    openExampleSong().catch((e) => toast(String(e))));

  window.addEventListener("keydown", (e) => {
    const mod = e.metaKey || e.ctrlKey;
    if (mod && e.key === "s") { e.preventDefault(); saveCurrent().catch((err) => toast(String(err))); }
    if (mod && e.key === "e") { e.preventDefault(); exportCurrent("txt").catch((err) => toast(String(err))); }
  });

  await refreshSongList();
  const songs = await API.songs();
  if (songs.length) {
    await openSong(songs[0].filename);
  } else {
    document.getElementById("start-here").classList.remove("hidden");
  }
}

document.addEventListener("DOMContentLoaded", () => {
  init().catch((err) => toast(`Failed to start: ${err}`));
});

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
  parseLine: (line) => fetchJSON("/api/parse", { method: "POST", body: JSON.stringify({ line }) }),
  render: (doc, instruments) => fetchJSON("/api/render", {
    method: "POST", body: JSON.stringify({ doc, instruments }),
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
function newSectionId() {
  return "section_" + Date.now().toString(36) + Math.floor(Math.random() * 1000);
}
function makeSection(name, type, instrument) {
  return {
    id: newSectionId(), name, type, instrument: instrument || defaultInstrument(),
    repeat: 1, transpose: 0, render: "chart", annotation: "", items: [],
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

  const lineInput = node.querySelector(".sec-chart-line");
  lineInput.value = sec.chart_line || "";

  wireSectionEvents(node, sec, idx);
  updateSectionPreview(node, sec);
  return node;
}

function wireSectionEvents(node, sec, idx) {
  const byId = () => currentDoc.sections.find((s) => s.id === sec.id);

  node.querySelector(".sec-name").addEventListener("input", (e) => {
    byId().name = e.target.value;
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

  const lineInput = node.querySelector(".sec-chart-line");
  const onLineChange = debounce(async () => {
    const line = lineInput.value;
    const s = byId();
    if (!line.trim()) {
      setChartItems(s, []);
      updateSectionPreview(node, s, { items: [], rendered: [] });
      schedulePreviewUpdate();
      return;
    }
    try {
      const result = await API.parseLine(line);
      if (result.ok) {
        setChartItems(s, result.items);
        s.chart_line = result.unparsed;
        updateSectionPreview(node, s, result);
      } else {
        updateSectionPreview(node, s, { error: result.error });
      }
    } catch (err) {
      updateSectionPreview(node, s, { error: String(err) });
    }
    schedulePreviewUpdate();
  }, 150);
  lineInput.addEventListener("input", onLineChange);

  node.querySelector(".sec-up").addEventListener("click", () => moveSection(sec.id, -1));
  node.querySelector(".sec-down").addEventListener("click", () => moveSection(sec.id, 1));
  node.querySelector(".sec-dup").addEventListener("click", () => duplicateSection(sec.id));
  node.querySelector(".sec-del").addEventListener("click", () => deleteSection(sec.id));
}

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
    fretRow.textContent = rendered[0] || "";
    symRow.textContent = rendered[1] || "";
    emptyEl.classList.add("hidden");
  } else if (!items.length && !measureItems(sec).length) {
    fretRow.textContent = "";
    symRow.textContent = "";
    emptyEl.classList.remove("hidden");
  } else if (rendered) {
    fretRow.textContent = "";
    symRow.textContent = "";
    emptyEl.classList.add("hidden");
  }
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

function duplicateSection(id) {
  const i = currentDoc.sections.findIndex((s) => s.id === id);
  if (i < 0) return;
  const source = currentDoc.sections[i];
  const newSec = {
    id: newSectionId(), name: source.name + " (ref)", type: source.type,
    instrument: source.instrument, repeat: 1, transpose: 0,
    render: source.render, annotation: "",
    items: [{ kind: "section_ref", section: source.id, repeat: 1, all: false, transpose: 0 }],
    chart_line: `=${source.id}`,
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

async function exportCurrent(kind) {
  if (!currentDoc) return;
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
//  Wire-up
// ---------------------------------------------------------------------------
async function init() {
  META = await API.meta();
  document.getElementById("app-version").textContent = `v${META.app_version}`;
  wireMetaForm();

  document.getElementById("btn-help").addEventListener("click", () => {
    document.getElementById("help-strip").classList.toggle("hidden");
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

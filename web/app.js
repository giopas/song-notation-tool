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
  importSong: (filename, data) => fetchJSON("/api/import", {
    method: "POST", body: JSON.stringify({ filename, data }),
  }),
  deleteSong: (name) => fetchJSON(`/api/songs/${encodeURIComponent(name)}`, { method: "DELETE" }),
  library: () => fetchJSON("/api/library"),
  libraryOp: (op, fields) => fetchJSON("/api/library", {
    method: "POST", body: JSON.stringify(Object.assign({ op }, fields || {})),
  }),
  trash: () => fetchJSON("/api/trash"),
  trashOp: (op, fields) => fetchJSON("/api/trash", {
    method: "POST", body: JSON.stringify(Object.assign({ op }, fields || {})),
  }),
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
  riffs: (op, fields) => fetch("/api/riffs", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify(Object.assign({ op, doc: currentDoc }, fields || {})),
  }).then((r) => r.json()),
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

function toast(msg, long) {
  // `long`: a message worth reading in full, such as why a file could
  // not be imported, stays up longer.
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.classList.remove("hidden");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.classList.add("hidden"), long ? 6000 : 2200);
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
//  Rendering: song list, collections, hidden songs and the trash
//
//  The list shows one view at a time: all songs (hidden ones left out), a
//  collection in its own order, the hidden songs, or the deleted ones.
//  Hidden songs and collections live in library.json in the songs folder;
//  a deleted song sits in "Deleted songs" there until the trash is emptied.
// ---------------------------------------------------------------------------
let SONGS = [];
let LIBRARY = { collections: [], hidden: [], trash_count: 0 };
// "all", "hidden", "trash" or "col:<id>". Remembered per browser only.
let songView = "all";
try { songView = localStorage.getItem("snt.songView") || "all"; } catch (e) { /* storage off */ }

function setSongView(v) {
  songView = v;
  try { localStorage.setItem("snt.songView", v); } catch (e) { /* storage off */ }
  return refreshSongList();
}

function viewCollection() {
  if (!songView.startsWith("col:")) return null;
  return LIBRARY.collections.find((c) => "col:" + c.id === songView) || null;
}

function songLabel(s) {
  return s.title || s.filename.replace(/\.sng$/i, "");
}

async function refreshSongList() {
  const [songs, lib] = await Promise.all([API.songs(), API.library()]);
  SONGS = songs;
  LIBRARY = lib;
  if (songView.startsWith("col:") && !viewCollection()) songView = "all";
  if (songView === "trash" && !lib.trash_count) songView = "all";
  renderViewPicker();

  const ul = document.getElementById("song-list");
  const note = document.getElementById("song-view-note");
  ul.innerHTML = "";
  document.getElementById("btn-empty-trash").classList.toggle("hidden", songView !== "trash");
  if (songView === "trash") return renderTrash(ul, note);

  const col = viewCollection();
  let list;
  if (col) {
    const byName = new Map(songs.map((s) => [s.filename, s]));
    list = col.songs.map((f) => byName.get(f)).filter(Boolean);
  } else if (songView === "hidden") {
    list = songs.filter((s) => s.hidden);
  } else {
    list = songs.filter((s) => !s.hidden);
  }
  ul.classList.toggle("is-collection", !!col);

  let msg = "";
  if (col && !list.length) {
    msg = "No songs here yet. In All songs, use a song's ⋯ menu to add it.";
  } else if (col && list.length > 1) {
    msg = "Drag a song to change the order, or select it and press Alt with ↑ or ↓.";
  } else if (songView === "hidden" && !list.length) {
    msg = "No hidden songs.";
  } else if (songView === "hidden") {
    msg = "Hidden songs are kept, but left out of All songs.";
  }
  note.textContent = msg;
  note.classList.toggle("hidden", !msg);

  list.forEach((s, i) => ul.appendChild(songListItem(s, i, list.length, col)));
}

function songListItem(s, index, count, col) {
  const li = document.createElement("li");
  li.className = s.filename === currentFilename ? "active" : "";
  li.dataset.filename = s.filename;
  // A real button, so the list can be used from the keyboard.
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "song-item";
  if (s.filename === currentFilename) btn.setAttribute("aria-current", "true");
  const tag = s.hidden && songView !== "hidden"
    ? `<span class="song-tag" title="Hidden from All songs">hidden</span>` : "";
  btn.innerHTML = (col ? `<span class="song-pos">${index + 1}</span>` : "") +
    `<span class="song-text"><span class="song-title">${escapeHtml(songLabel(s))}${tag}</span>` +
    (s.artist ? `<span class="song-artist">${escapeHtml(s.artist)}</span>` : "") + "</span>";
  btn.addEventListener("click", () => openSong(s.filename).catch((e) => toast(String(e))));
  li.appendChild(btn);

  const more = document.createElement("button");
  more.type = "button";
  more.className = "icon-btn song-more";
  more.setAttribute("aria-label", `More for ${songLabel(s)}: collections, hide, delete`);
  more.setAttribute("aria-haspopup", "menu");
  more.setAttribute("aria-expanded", "false");
  more.title = "Collections, hide, delete";
  more.innerHTML = '<svg class="ico" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><circle cx="5" cy="12" r="1.7" fill="currentColor" stroke="none"/><circle cx="12" cy="12" r="1.7" fill="currentColor" stroke="none"/><circle cx="19" cy="12" r="1.7" fill="currentColor" stroke="none"/></svg>';
  more.addEventListener("click", (e) => {
    e.stopPropagation();
    openSongMenu(s, more, index, count, col);
  });
  li.appendChild(more);

  if (col) {
    // The running order: drag, or Alt + arrow keys on the song.
    li.draggable = true;
    li.addEventListener("dragstart", (e) => {
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", s.filename);
      li.classList.add("dragging");
    });
    li.addEventListener("dragend", () => {
      li.classList.remove("dragging");
      document.querySelectorAll("#song-list .drop-before, #song-list .drop-after")
        .forEach((n) => n.classList.remove("drop-before", "drop-after"));
    });
    li.addEventListener("dragover", (e) => {
      e.preventDefault();
      const r = li.getBoundingClientRect();
      const after = e.clientY > r.top + r.height / 2;
      li.classList.toggle("drop-after", after);
      li.classList.toggle("drop-before", !after);
    });
    li.addEventListener("dragleave", () => li.classList.remove("drop-before", "drop-after"));
    li.addEventListener("drop", (e) => {
      e.preventDefault();
      const moving = e.dataTransfer.getData("text/plain");
      const after = li.classList.contains("drop-after");
      li.classList.remove("drop-before", "drop-after");
      if (!moving || moving === s.filename) return;
      const order = col.songs.filter((f) => f !== moving);
      let to = order.indexOf(s.filename);
      if (to < 0) return;
      if (after) to += 1;
      moveInCollection(col, moving, to);
    });
    btn.addEventListener("keydown", (e) => {
      if (!e.altKey || (e.key !== "ArrowUp" && e.key !== "ArrowDown")) return;
      e.preventDefault();
      const to = index + (e.key === "ArrowUp" ? -1 : 1);
      if (to < 0 || to >= count) return;
      moveInCollection(col, s.filename, to, true);
    });
  }
  return li;
}

async function moveInCollection(col, filename, to, refocus) {
  await libraryOp("move", { id: col.id, song: filename, to });
  if (refocus) {
    const li = [...document.querySelectorAll("#song-list li")]
      .find((n) => n.dataset.filename === filename);
    if (li) li.querySelector(".song-item").focus();
  }
}

/** The view picker over the list: All songs, each collection, Hidden, Deleted. */
function renderViewPicker() {
  const sel = document.getElementById("song-view");
  sel.innerHTML = "";
  const add = (parent, value, text) => {
    const o = document.createElement("option");
    o.value = value; o.textContent = text;
    parent.appendChild(o);
  };
  const visible = SONGS.filter((s) => !s.hidden).length;
  add(sel, "all", `All songs (${visible})`);
  if (LIBRARY.collections.length) {
    const g = document.createElement("optgroup");
    g.label = "Collections";
    LIBRARY.collections.forEach((c) => add(g, "col:" + c.id, `${c.name} (${c.songs.length})`));
    sel.appendChild(g);
  }
  if (LIBRARY.hidden.length || songView === "hidden") {
    add(sel, "hidden", `Hidden songs (${LIBRARY.hidden.length})`);
  }
  if (LIBRARY.trash_count) add(sel, "trash", `Deleted songs (${LIBRARY.trash_count})`);
  sel.value = songView;
  if (sel.value !== songView) { songView = "all"; sel.value = "all"; }
}

async function libraryOp(op, fields) {
  const res = await API.libraryOp(op, fields);
  if (!res.ok) toast(res.error || "That didn't work", true);
  await refreshSongList();
  return res;
}

// ── the song's ⋯ menu ────────────────────────────────────────────────────
function closeSongMenu() {
  const menu = document.getElementById("song-menu");
  if (menu.classList.contains("hidden")) return;
  menu.classList.add("hidden");
  if (menu._anchor) {
    menu._anchor.setAttribute("aria-expanded", "false");
    if (menu.contains(document.activeElement)) menu._anchor.focus();
  }
  menu._anchor = null;
}

/** Fill and place a menu next to `anchor`. `items` are
 *  {label, run, danger, checked, disabled} or {sep:true} or {heading}. */
function showMenu(menu, anchor, items) {
  menu.innerHTML = "";
  for (const it of items) {
    if (it.sep) { const d = document.createElement("div"); d.className = "menu-sep"; menu.appendChild(d); continue; }
    if (it.heading) { const d = document.createElement("div"); d.className = "menu-label"; d.textContent = it.heading; menu.appendChild(d); continue; }
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = it.label;
    b.setAttribute("role", it.checked === undefined ? "menuitem" : "menuitemcheckbox");
    if (it.checked !== undefined) b.setAttribute("aria-checked", String(!!it.checked));
    if (it.danger) b.classList.add("danger");
    if (it.disabled) b.disabled = true;
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      closeSongMenu(); closeViewMenu();
      Promise.resolve(it.run()).catch((err) => toast(String(err), true));
    });
    menu.appendChild(b);
  }
  menu.classList.remove("hidden");
  return menu;
}

function placeMenu(menu, anchor) {
  const r = anchor.getBoundingClientRect();
  const w = menu.offsetWidth, h = menu.offsetHeight;
  let top = r.bottom + 4;
  if (top + h > window.innerHeight - 8) top = Math.max(8, r.top - h - 4);
  const left = Math.max(8, Math.min(r.right - w, window.innerWidth - w - 8));
  menu.style.top = `${top}px`;
  menu.style.left = `${left}px`;
}

function wireMenuKeys(menu, close) {
  menu.addEventListener("keydown", (e) => {
    const items = [...menu.querySelectorAll("button:not(:disabled)")];
    const i = items.indexOf(document.activeElement);
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); close(); }
    else if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      const n = items.length;
      if (n) items[(i + (e.key === "ArrowDown" ? 1 : n - 1) + n) % n].focus();
    } else if (e.key === "Tab") { close(); }
  });
}

function openSongMenu(s, anchor, index, count, col) {
  const menu = document.getElementById("song-menu");
  if (!menu.classList.contains("hidden") && menu._anchor === anchor) return closeSongMenu();
  closeSongMenu(); closeViewMenu();
  const items = [];
  if (col) {
    items.push({ heading: col.name });
    items.push({ label: "Move up", disabled: index === 0,
                 run: () => moveInCollection(col, s.filename, index - 1) });
    items.push({ label: "Move down", disabled: index >= count - 1,
                 run: () => moveInCollection(col, s.filename, index + 1) });
    items.push({ label: "Remove from this collection",
                 run: () => libraryOp("remove", { id: col.id, song: s.filename }) });
    items.push({ sep: true });
  }
  items.push({ heading: "In collections" });
  for (const c of LIBRARY.collections) {
    const isIn = (s.collections || []).includes(c.id);
    items.push({ label: c.name, checked: isIn,
                 run: () => libraryOp(isIn ? "remove" : "add", { id: c.id, song: s.filename }) });
  }
  items.push({ label: "New collection with this song…", run: () => newCollection([s.filename]) });
  items.push({ sep: true });
  items.push(s.hidden
    ? { label: "Show in All songs", run: () => libraryOp("hide", { song: s.filename, hidden: false }) }
    : { label: "Hide from All songs", run: () => hideSong(s) });
  items.push({ label: "Delete…", danger: true, run: () => deleteSong(s) });
  showMenu(menu, anchor, items);
  menu._anchor = anchor;
  anchor.setAttribute("aria-expanded", "true");
  placeMenu(menu, anchor);
  const first = menu.querySelector("button:not(:disabled)");
  if (first) first.focus();
}

async function hideSong(s) {
  await libraryOp("hide", { song: s.filename, hidden: true });
  toast(`"${songLabel(s)}" is hidden. Find it under Hidden songs.`);
}

async function deleteSong(s) {
  const isOpen = s.filename === currentFilename;
  const unsaved = isOpen && currentDoc && JSON.stringify(currentDoc) !== savedSnapshot;
  const ok = await confirmDialog({
    title: "Delete this song?",
    message: `"${songLabel(s)}" moves to Deleted songs. You can restore it from there ` +
      "until you empty the trash." +
      (unsaved ? " Its changes that are not saved yet are lost." : ""),
    confirm: "Delete",
  });
  if (!ok) return;
  await API.deleteSong(s.filename);
  if (isOpen) closeCurrentSong();
  await refreshSongList();
  toast(`"${songLabel(s)}" moved to Deleted songs.`);
}

/** The open song went away (deleted): back to the start screen. */
function closeCurrentSong() {
  currentFilename = null;
  currentDoc = null;
  markSaved();
  document.getElementById("editor").classList.add("hidden");
  document.getElementById("start-here").classList.remove("hidden");
  setSongToolsEnabled(false);
}

// ── collections: the ⋯ menu beside the view picker ─────────────────────
function closeViewMenu() {
  const menu = document.getElementById("view-menu");
  if (menu.classList.contains("hidden")) return;
  menu.classList.add("hidden");
  const btn = document.getElementById("btn-view-more");
  btn.setAttribute("aria-expanded", "false");
  if (menu.contains(document.activeElement)) btn.focus();
}

function openViewMenu() {
  const menu = document.getElementById("view-menu");
  const btn = document.getElementById("btn-view-more");
  if (!menu.classList.contains("hidden")) return closeViewMenu();
  closeSongMenu();
  const col = viewCollection();
  const items = [{ label: "New collection…", run: () => newCollection([]) }];
  if (col) {
    const empty = !col.songs.length;
    items.push({ sep: true }, { heading: col.name });
    items.push({ label: "Export as one PDF", disabled: empty, run: () => exportCollection(col, "portrait") });
    items.push({ label: "Export as one PDF, landscape", disabled: empty, run: () => exportCollection(col, "landscape") });
    items.push({ label: "Rename…", run: () => renameCollection(col) });
    items.push({ label: "Delete collection…", danger: true, run: () => deleteCollection(col) });
  }
  showMenu(menu, btn, items);
  btn.setAttribute("aria-expanded", "true");
  placeMenu(menu, btn);
  const first = menu.querySelector("button:not(:disabled)");
  if (first) first.focus();
}

async function newCollection(songs) {
  const name = await nameDialog({
    title: "New collection",
    hint: "A setlist, an album, or any group of songs. A song can be in several.",
    value: "", confirm: "Create",
  });
  if (!name) return;
  const res = await API.libraryOp("create", { name, songs });
  if (!res.ok) return toast(res.error, true);
  await setSongView("col:" + res.collection.id);
  toast(songs.length ? `Added to "${res.collection.name}"`
                     : `"${res.collection.name}" is empty. Add songs from All songs with their ⋯ menu.`, true);
}

async function renameCollection(col) {
  const name = await nameDialog({ title: "Rename collection", value: col.name, confirm: "Rename" });
  if (!name || name === col.name) return;
  await libraryOp("rename", { id: col.id, name });
}

async function deleteCollection(col) {
  const ok = await confirmDialog({
    title: "Delete this collection?",
    message: `"${col.name}" goes away. Its ${col.songs.length} song` +
      `${col.songs.length === 1 ? "" : "s"} stay in your songs folder, and in any other collection.`,
    confirm: "Delete collection",
  });
  if (!ok) return;
  await libraryOp("delete", { id: col.id });
  await setSongView("all");
}

async function exportCollection(col, orient) {
  // The PDF is made from the saved files: save the open song first if it
  // is in this collection and has changes, so the PDF shows them.
  if (currentDoc && col.songs.includes(currentFilename)
      && JSON.stringify(currentDoc) !== savedSnapshot) {
    await saveCurrent();
  }
  const api = nativeApi();
  if (api && api.export_collection) {
    const res = await api.export_collection(col.id, orient);
    if (res && res.ok) toast(`Saved to ${res.path}` + exportWarningText(res.warnings), true);
    else if (!(res && res.cancelled)) toast(`Export failed: ${(res && res.error) || "unknown error"}`, true);
    return;
  }
  const res = await fetch(`/api/collections/${encodeURIComponent(col.id)}/export.pdf?orient=${orient}`);
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    toast(`Export failed: ${err.error || res.statusText}`, true);
    return;
  }
  const m = /filename="([^"]+)"/.exec(res.headers.get("Content-Disposition") || "");
  const filename = m ? m[1] : "Collection.pdf";
  const url = URL.createObjectURL(await res.blob());
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
  let warnings = [];
  try { warnings = JSON.parse(res.headers.get("X-Export-Warnings") || "[]"); } catch (e) { warnings = []; }
  toast(`Downloaded ${filename}. It is in your browser's downloads folder.` + exportWarningText(warnings), true);
}

// ── the trash ───────────────────────────────────────────────────────────
async function renderTrash(ul, note) {
  ul.classList.remove("is-collection");
  const trash = await API.trash();
  note.textContent = trash.length
    ? "Deleted songs stay here until you empty the trash. Restore puts a song back in its collections."
    : "The trash is empty.";
  note.classList.remove("hidden");
  for (const t of trash) {
    const li = document.createElement("li");
    li.className = "trash-item";
    const when = t.deleted_at ? new Date(t.deleted_at) : null;
    const whenText = when && !isNaN(when) ? "Deleted " + when.toLocaleDateString() : "";
    li.innerHTML = `<div class="song-text"><span class="song-title">${escapeHtml(t.title || t.original.replace(/\.sng$/i, ""))}</span>` +
      `<span class="song-artist">${escapeHtml([t.artist, whenText].filter(Boolean).join(" · "))}</span></div>`;
    const actions = document.createElement("div");
    actions.className = "trash-actions";
    const restore = document.createElement("button");
    restore.type = "button"; restore.className = "btn subtle"; restore.textContent = "Restore";
    restore.addEventListener("click", async () => {
      const res = await API.trashOp("restore", { name: t.name });
      if (!res.ok) return toast(res.error, true);
      await refreshSongList();
      toast(`Restored "${t.title || res.filename}".`);
    });
    const purge = document.createElement("button");
    purge.type = "button"; purge.className = "btn subtle danger-text"; purge.textContent = "Delete forever";
    purge.addEventListener("click", async () => {
      const ok = await confirmDialog({
        title: "Delete forever?",
        message: `"${t.title || t.original}" is deleted from your computer. This can't be undone.`,
        confirm: "Delete forever",
      });
      if (!ok) return;
      const res = await API.trashOp("purge", { name: t.name });
      if (!res.ok) toast(res.error, true);
      await refreshSongList();
    });
    actions.append(restore, purge);
    li.appendChild(actions);
    ul.appendChild(li);
  }
}

async function emptyTrash() {
  const n = LIBRARY.trash_count;
  const ok = await confirmDialog({
    title: "Empty the trash?",
    message: `${n} deleted song${n === 1 ? " is" : "s are"} deleted from your computer for good. This can't be undone.`,
    confirm: "Empty trash",
  });
  if (!ok) return;
  const res = await API.trashOp("empty");
  if (!res.ok) return toast(res.error, true);
  await setSongView("all");
  toast(`Deleted ${res.deleted} song${res.deleted === 1 ? "" : "s"} for good.`);
}

// ── the two small dialogs: confirm, and a name ─────────────────────────
function confirmDialog({ title, message, confirm = "OK", danger = true }) {
  return new Promise((resolve) => {
    const bd = document.getElementById("confirm-modal-backdrop");
    const ok = document.getElementById("confirm-ok");
    const cancel = document.getElementById("confirm-cancel");
    document.getElementById("confirm-title").textContent = title;
    document.getElementById("confirm-message").textContent = message;
    ok.textContent = confirm;
    ok.classList.toggle("danger", danger);
    ok.classList.toggle("accent", !danger);
    const done = (v) => {
      ok.onclick = cancel.onclick = null;
      bd.classList.add("hidden");
      resolve(v);
    };
    ok.onclick = () => done(true);
    cancel.onclick = () => done(false);
    bd.classList.remove("hidden");
  });
}

function nameDialog({ title, hint = "", value = "", confirm = "OK" }) {
  return new Promise((resolve) => {
    const bd = document.getElementById("name-modal-backdrop");
    const input = document.getElementById("name-input");
    const ok = document.getElementById("name-ok");
    const cancel = document.getElementById("name-cancel");
    document.getElementById("name-title").textContent = title;
    const h = document.getElementById("name-hint");
    h.textContent = hint;
    h.classList.toggle("hidden", !hint);
    input.value = value;
    ok.textContent = confirm;
    const done = (v) => {
      ok.onclick = cancel.onclick = input.onkeydown = null;
      bd.classList.add("hidden");
      resolve(v);
    };
    ok.onclick = () => { const v = input.value.trim(); if (v) done(v); else input.focus(); };
    cancel.onclick = () => done(null);
    input.onkeydown = (e) => { if (e.key === "Enter") { e.preventDefault(); ok.click(); } };
    bd.classList.remove("hidden");
    setTimeout(() => input.select(), 0);
  });
}

function wireSongLibrary() {
  document.getElementById("song-view").addEventListener("change", (e) =>
    setSongView(e.target.value).catch((err) => toast(String(err))));
  const viewBtn = document.getElementById("btn-view-more");
  viewBtn.addEventListener("click", (e) => { e.stopPropagation(); openViewMenu(); });
  wireMenuKeys(document.getElementById("view-menu"), closeViewMenu);
  wireMenuKeys(document.getElementById("song-menu"), closeSongMenu);
  document.getElementById("btn-empty-trash").addEventListener("click", () =>
    emptyTrash().catch((err) => toast(String(err), true)));
  document.addEventListener("click", (e) => {
    if (!e.target.closest("#song-menu")) closeSongMenu();
    if (!e.target.closest("#view-menu")) closeViewMenu();
  });
  // A menu placed against the list would drift as it scrolls: close it.
  document.getElementById("sidebar").addEventListener("scroll", () => { closeSongMenu(); closeViewMenu(); });
  window.addEventListener("resize", () => { closeSongMenu(); closeViewMenu(); });
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
  showEditor();
  renderEditor();
  markSaved();
  refreshSongList();
  schedulePreviewUpdate();
}

async function createSong(fields) {
  const { filename, doc } = await API.newSong(fields);
  currentFilename = filename;
  currentDoc = doc;
  showEditor();
  renderEditor();
  markSaved();
  refreshSongList();
  toast(`Created ${filename}`);
}

async function openExampleSong() {
  const { filename, doc } = await API.openExample();
  currentFilename = filename;
  currentDoc = doc;
  showEditor();
  renderEditor();
  markSaved();
  refreshSongList();
}

// Import: a .sng file, or a PDF this app exported (it carries the song).
// The server always makes a new song file, so nothing is overwritten.
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(",", 2)[1] || "");
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

async function importSongFile(file) {
  try {
    const { filename, doc } = await API.importSong(file.name, await fileToBase64(file));
    currentFilename = filename;
    currentDoc = doc;
    document.getElementById("start-here").classList.add("hidden");
    document.getElementById("editor").classList.remove("hidden");
    renderEditor();
    refreshSongList();
    schedulePreviewUpdate();
    toast(`Imported ${file.name} as ${filename}`);
  } catch (err) {
    toast(err.message, true);
  }
}

function wireImport() {
  const input = document.getElementById("import-file");
  const pick = () => { input.value = ""; input.click(); };
  document.getElementById("btn-import").addEventListener("click", pick);
  document.getElementById("start-import").addEventListener("click", pick);
  input.addEventListener("change", () => {
    if (input.files && input.files[0]) importSongFile(input.files[0]);
  });
}

// ---------------------------------------------------------------------------
//  Layout panel
//
//  Everything about how the song comes out on paper, in one place at the
//  bottom of the sidebar: out of the way of the song itself, and one spot
//  to look instead of a row of dropdowns that kept growing.
// ---------------------------------------------------------------------------
const LYRICS_HINTS = {
  none: "Nothing prints. The words stay in the Lyrics dialog for reference.",
  start: "Every section's words in one block before the chart, each under its section's name.",
  end: "Every section's words in one block after the chart, each under its section's name.",
  side: "Every section's words down a column on the left, with the chart in one column beside them. This sets the columns, so Columns is greyed out.",
  beside: "Each section's words in a column to the right of its own chart, next to what you play under them. Chords are not lined up with syllables.",
  below: "Each section's words printed under its own chart.",
};

function optionLabel(id) {
  const sel = document.getElementById(id);
  const opt = sel && sel.options[sel.selectedIndex];
  return opt ? opt.textContent : "";
}

/** Keep the panel's dependent bits in step with its settings: the lyrics
 *  hint, Columns stepping aside in "side" layout, and the one-line summary
 *  under the sidebar button. */
function refreshLayoutPanel() {
  if (!currentDoc) return;
  const lyr = currentDoc.lyrics_layout || "none";
  document.getElementById("layout-lyrics-hint").textContent = LYRICS_HINTS[lyr] || "";

  // In "side" layout the page is the words' column plus one chart column,
  // whatever Columns says — so say that, rather than leave a control on
  // screen that silently does nothing.
  const cols = document.getElementById("meta-columns");
  const side = lyr === "side";
  const floor = (currentDoc.pdf_scale || "fit") === "floor";
  cols.disabled = side || floor;
  document.getElementById("layout-columns-hint").textContent = side
    ? "Set by the lyrics: their column on the left, one chart column beside it."
    : floor
    ? "Set by Floor: one column across the page, or two when a long song gets bigger type that way. Portrait or landscape, whichever is bigger."
    : "Two columns halve the height a chart needs, and the fit uses the space for bigger type. Auto only splits the page when that helps.";

  renderLayoutInstruments();

  const bits = [optionLabel("meta-scale")];
  if (!side && !floor) bits.push(optionLabel("meta-columns").replace(/\s*\(.*\)$/, ""));
  bits.push(lyr === "none" ? "no lyrics" : `lyrics: ${({
    start: "at the start", end: "at the end", side: "left column",
    beside: "beside", below: "under" })[lyr]}`);
  document.getElementById("layout-summary").textContent = bits.filter(Boolean).join(" · ");
}

/** Every instrument the song uses, in sections and parts. */
function songInstruments() {
  const out = [];
  (currentDoc.sections || []).forEach((s) => {
    [s.instrument, ...((s.parts || []).map((p) => p.instrument))].forEach((i) => {
      if (i && !out.includes(i)) out.push(i);
    });
  });
  return out;
}

/** Layout → Instruments: a box per instrument; all ticked means all print
 *  (stored as an empty list, so a new instrument prints without asking). */
function renderLayoutInstruments() {
  const box = document.getElementById("layout-instruments");
  if (!box || !currentDoc) return;
  const all = songInstruments();
  const chosen = (currentDoc.print_instruments || []).filter((i) => all.includes(i));
  box.innerHTML = "";
  all.forEach((name) => {
    const label = document.createElement("label");
    label.className = "layout-instrument";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = !chosen.length || chosen.includes(name);
    cb.addEventListener("change", () => {
      const ticked = [...box.querySelectorAll("input")]
        .map((el, i) => (el.checked ? all[i] : null)).filter(Boolean);
      // None ticked would print nothing: keep at least this one.
      if (!ticked.length) { cb.checked = true; return; }
      currentDoc.print_instruments = ticked.length === all.length ? [] : ticked;
      schedulePreviewUpdate();
    });
    label.appendChild(cb);
    label.appendChild(document.createTextNode(" " + name));
    box.appendChild(label);
  });
}

function openLayoutPanel() {
  if (!currentDoc) return;
  refreshLayoutPanel();
  document.getElementById("layout-panel").classList.remove("hidden");
  document.getElementById("btn-layout").classList.add("active");
}

function closeLayoutPanel() {
  document.getElementById("layout-panel").classList.add("hidden");
  document.getElementById("btn-layout").classList.remove("active");
}

function wireLayoutPanel() {
  const panel = document.getElementById("layout-panel");
  const btn = document.getElementById("btn-layout");
  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    if (panel.classList.contains("hidden")) openLayoutPanel();
    else closeLayoutPanel();
  });
  document.getElementById("layout-close").addEventListener("click", closeLayoutPanel);
  // A popover, not a modal: it closes when you click back into the song,
  // and stays open while you change several settings in a row — the
  // preview keeps updating underneath it.
  document.addEventListener("mousedown", (e) => {
    if (panel.classList.contains("hidden")) return;
    if (panel.contains(e.target) || btn.contains(e.target)) return;
    closeLayoutPanel();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !panel.classList.contains("hidden")) closeLayoutPanel();
  });
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
    currentDoc.lyrics_layout || "none";
  document.getElementById("meta-chord-sheet").value =
    currentDoc.chord_sheet || "none";
  document.getElementById("meta-lick-refs").value =
    currentDoc.lick_refs || "tab";
  document.getElementById("btn-layout").disabled = false;
  refreshLayoutPanel();

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
  showRepeatChip(node, sec);

  showRenderMode(node, sec.render || "chart");
  showPrintSize(node, sec.print_size || "normal");

  const lineInput = node.querySelector(".sec-chart-line");
  lineInput.value = sec.chart_line || "";

  node.querySelector(".sec-free-text").value = sec.free_text || "";
  node.querySelector(".sec-annotation").value = sec.annotation || "";

  // Populate tab beats dropdown
  fillBeatsSelect(node.querySelector(".tab-beats-select"));

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

/** The length a new tab measure starts with: "8 beats", "16 beats"… */
function fillBeatsSelect(sel) {
  (META.tab_beats_options || [8, 16, 32, 64]).forEach((b) => {
    const opt = document.createElement("option");
    opt.value = b; opt.textContent = `${b} beats`;
    if (b === (META.tab_beats_default || 8)) opt.selected = true;
    sel.appendChild(opt);
  });
}

/** Chart / Tab / Both / Free: one pressed button in the section header. */
function showRenderMode(node, mode) {
  node.querySelectorAll(":scope > .section-head .sec-render button").forEach((b) => {
    b.setAttribute("aria-pressed", String(b.dataset.mode === mode));
  });
}

const PRINT_SIZE_LABELS = { smaller: "Prints smaller", larger: "Prints larger" };

/** The ⋯ menu's size items, and a badge in the header when it is not the song's size. */
function showPrintSize(node, size) {
  node.querySelectorAll(":scope > .section-head .sec-menu [data-size]").forEach((b) => {
    b.setAttribute("aria-checked", String(b.dataset.size === size));
  });
  const badge = node.querySelector(":scope > .section-head .sec-size-badge");
  badge.textContent = PRINT_SIZE_LABELS[size] || "";
  badge.classList.toggle("hidden", !PRINT_SIZE_LABELS[size]);
}

/** Open or close a section's ⋯ menu. Only one is open at a time. */
function toggleSectionMenu(node, open) {
  const menu = node.querySelector(":scope > .section-head .sec-menu");
  const btn = node.querySelector(":scope > .section-head .sec-more");
  if (open === undefined) open = menu.classList.contains("hidden");
  document.querySelectorAll(".sec-menu").forEach((m) => {
    if (m !== menu) m.classList.add("hidden");
  });
  document.querySelectorAll(".sec-more").forEach((b) => {
    if (b !== btn) b.setAttribute("aria-expanded", "false");
  });
  menu.classList.toggle("hidden", !open);
  btn.setAttribute("aria-expanded", String(open));
  if (open) {
    const first = menu.querySelector('[aria-checked="true"]') || menu.querySelector("button");
    if (first) first.focus();
  }
}

/**
 * The card you are working in. Its insert buttons show; the others keep
 * them hidden, so a long song is not a wall of identical button rows.
 */
function setActiveCard(node) {
  document.querySelectorAll(".section-card.is-active").forEach((c) => {
    if (c !== node) c.classList.remove("is-active");
  });
  if (node) node.classList.add("is-active");
}

/**
 * Repeats live in the chart line ([G D]x2). A section from an older song
 * can still carry a whole-section repeat, printed as (xN) after its name;
 * it shows as a chip that removes it, instead of a number box on every
 * card that read as a mystery "1".
 */
function showRepeatChip(node, sec) {
  const chip = node.querySelector(".sec-repeat-chip");
  const n = parseInt(sec.repeat, 10) || 1;
  chip.textContent = `×${n}  ✕`;
  chip.classList.toggle("hidden", n <= 1);
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
  node.querySelector(".sec-repeat-chip").addEventListener("click", () => {
    const s = byId();
    s.repeat = 1;
    showRepeatChip(node, s);
    schedulePreviewUpdate();
  });

  node.querySelector(".sec-render").addEventListener("click", (e) => {
    const b = e.target.closest("button[data-mode]");
    if (!b) return;
    const s = byId();
    s.render = b.dataset.mode;
    showRenderMode(node, s.render);
    updateTabGridVisibility(node, s);
    refreshReferencingPreviews(sec.id);
    schedulePreviewUpdate();
  });

  // The ⋯ menu: print size, copy, delete.
  const head = node.querySelector(".section-head");
  head.querySelector(".sec-more").addEventListener("click", (e) => {
    e.stopPropagation();
    toggleSectionMenu(node);
  });
  const menu = head.querySelector(".sec-menu");
  menu.addEventListener("keydown", (e) => {
    const items = [...menu.querySelectorAll("button")];
    const i = items.indexOf(document.activeElement);
    if (e.key === "Escape") {
      e.preventDefault(); e.stopPropagation();
      toggleSectionMenu(node, false);
      head.querySelector(".sec-more").focus();
    } else if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      const n = items.length;
      items[(i + (e.key === "ArrowDown" ? 1 : n - 1)) % n].focus();
    }
  });
  // Printed size of this section against the rest of the song. "Normal"
  // is stored as no key at all, so a song that never used it is unchanged.
  menu.querySelectorAll("[data-size]").forEach((b) => b.addEventListener("click", () => {
    const s = byId();
    if (b.dataset.size === "normal") delete s.print_size;
    else s.print_size = b.dataset.size;
    showPrintSize(node, b.dataset.size);
    toggleSectionMenu(node, false);
    schedulePreviewUpdate();
  }));
  node.addEventListener("focusin", () => setActiveCard(node));
  node.addEventListener("mousedown", () => setActiveCard(node));

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
  attachAutocomplete(lineInput);

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
    if (btn.dataset.promote) {
      promoteSelection(node.querySelector(".sec-chart-line"), byId());
      return;
    }
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

  // Parts: another instrument's line in this section.
  renderParts(node, byId());
  node.querySelector(".sec-add-part").addEventListener("click", () => {
    const s = byId();
    s.parts = s.parts || [];
    const others = Object.keys(META.instruments).filter((i) => i !== s.instrument);
    s.parts.push({ id: `p${Date.now().toString(36)}`, instrument: others[0] || s.instrument,
                   items: [], chart_line: "" });
    renderParts(node, s);
    const inputs = node.querySelectorAll(".part-line");
    if (inputs.length) inputs[inputs.length - 1].focus();
  });

  wireDragHandle(node, sec.id);
  // The copy picker opens beside the ⋯ button: the menu item it was
  // chosen from is gone by then.
  node.querySelector(".sec-dup").addEventListener("click", () => {
    toggleSectionMenu(node, false);
    openCopyPicker(sec.id, node.querySelector(".sec-more"));
  });
  node.querySelector(".sec-del").addEventListener("click", () => {
    toggleSectionMenu(node, false);
    deleteSection(sec.id);
  });
}

/**
 * A section's parts, one row each: the instrument, its chart line (parsed
 * and previewed exactly like the section's own), and a remove button.
 */
function renderParts(node, sec) {
  const box = node.querySelector(".sec-parts");
  if (!box) return;
  box.innerHTML = "";
  (sec.parts || []).forEach((part) => {
    const row = document.createElement("div");
    row.className = "sec-part";
    const head = document.createElement("div");
    head.className = "sec-part-head";
    const sel = document.createElement("select");
    sel.className = "part-instrument";
    sel.setAttribute("aria-label", "Instrument for this part");
    sel.title = "The instrument this line is for";
    Object.keys(META.instruments).forEach((name) => {
      const opt = document.createElement("option");
      opt.value = name; opt.textContent = name;
      if (name === part.instrument) opt.selected = true;
      sel.appendChild(opt);
    });
    // The same four modes as a section: what this part shows and prints.
    const mode = document.createElement("select");
    mode.className = "part-render";
    mode.setAttribute("aria-label", "Show this part as");
    mode.title = "What this part holds: Chart (a chart line), Tab (a tab grid), Both, or Free (any text, kept as typed)";
    [["chart", "Chart"], ["tab", "Tab"], ["both", "Both"], ["free", "Free"]].forEach(([v, t]) => {
      const opt = document.createElement("option");
      opt.value = v; opt.textContent = t;
      if (v === (part.render || "chart")) opt.selected = true;
      mode.appendChild(opt);
    });
    const input = document.createElement("input");
    input.type = "text";
    input.className = "part-line sec-chart-line-like";
    input.spellcheck = false;
    input.setAttribute("aria-label", "Chart line for this part");
    input.placeholder = "This instrument's chart line, e.g. A D E7";
    input.value = part.chart_line || "";
    const spacer = document.createElement("span");
    spacer.className = "part-spacer";
    const del = document.createElement("button");
    del.type = "button";
    del.className = "icon-btn";
    del.innerHTML = '<svg class="ico" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M6 6l12 12M18 6L6 18"/></svg>';
    del.title = "Remove this part";
    del.setAttribute("aria-label", "Remove this part");
    head.append(sel, mode, input, spacer, del);
    const err = document.createElement("div");
    err.className = "sec-error hidden";
    const preview = document.createElement("div");
    preview.className = "sec-preview part-preview is-empty";
    // Tab: the same grid editor as a section's, with its own toolbar.
    const grid = document.createElement("div");
    grid.className = "part-tab-grid";
    grid.innerHTML = '<div class="tab-grid-toolbar">'
      + '<button type="button" class="btn part-add-measure" title="Add a measure to this part\'s tab grid">+ Measure</button>'
      + '<label class="tab-beats-label">New measures start with <select class="part-beats-select"></select></label>'
      + '</div><div class="tab-grid-body"></div>';
    const beatsSel = grid.querySelector(".part-beats-select");
    fillBeatsSelect(beatsSel);
    const gridBody = grid.querySelector(".tab-grid-body");
    // Free: a text box kept as typed.
    const free = document.createElement("textarea");
    free.className = "part-free-text";
    free.rows = 4;
    free.spellcheck = false;
    free.placeholder = "Free text for this instrument, kept and printed exactly as typed.";
    free.value = part.free_text || "";
    err.setAttribute("role", "alert");
    row.append(head, err, preview, grid, free);
    box.appendChild(row);

    const show = () => {
      const m = part.render || "chart";
      const chart = m === "chart" || m === "both";
      const tab = m === "tab" || m === "both";
      input.classList.toggle("hidden", !chart);
      spacer.classList.toggle("hidden", chart);
      preview.classList.toggle("hidden", !chart);
      if (!chart) err.classList.add("hidden");
      grid.classList.toggle("hidden", !tab);
      free.classList.toggle("hidden", m !== "free");
      if (tab) rebuildTabGrid(row, part, gridBody);
    };

    sel.addEventListener("change", () => {
      part.instrument = sel.value;
      renderLayoutInstruments();
      if ((part.render || "chart") !== "chart") rebuildTabGrid(row, part, gridBody);
      schedulePreviewUpdate();
    });
    mode.addEventListener("change", () => {
      // Stored only when it is not Chart, as older songs have it.
      if (mode.value === "chart") delete part.render;
      else part.render = mode.value;
      show();
      schedulePreviewUpdate();
    });
    grid.querySelector(".part-add-measure").addEventListener("click", () => {
      const beats = parseInt(beatsSel.value, 10) || 8;
      const stringMap = {};
      stringsForInstrument(part.instrument).forEach((st) => { stringMap[st] = ""; });
      const measures = measureItems(part);
      measures.push({ kind: "measure", beats, strings: stringMap });
      setMeasureItems(part, measures);
      rebuildTabGrid(row, part, gridBody);
      schedulePreviewUpdate();
    });
    free.addEventListener("input", () => {
      part.free_text = free.value;
      schedulePreviewUpdate();
    });
    del.addEventListener("click", () => {
      sec.parts = (sec.parts || []).filter((p) => p !== part);
      renderParts(node, sec);
      schedulePreviewUpdate();
    });
    const paint = (result) => {
      if (!result || !result.ok) {
        err.textContent = result && result.error ? friendlyError(result.error) : "";
        err.classList.toggle("hidden", !(result && result.error));
        preview.classList.toggle("is-stale", !!(result && result.error));
        return;
      }
      err.classList.add("hidden");
      preview.classList.remove("is-stale");
      preview.classList.toggle("is-empty", !(result.rendered || []).length);
      preview.innerHTML = "";
      (result.rendered || []).forEach((text, i) => {
        const div = document.createElement("div");
        const role = (result.roles || [])[i] || "";
        div.className = "sec-preview-more" + (role ? " sec-row-" + role : "");
        paintRow(div, { text, role, spans: (result.spans || [])[i] });
        preview.appendChild(div);
      });
      applyBracketRuns(preview, [...preview.querySelectorAll(".sec-preview-more")],
                       result.brackets);
    };
    // The chart line sets the part's chart items only; its tab measures stay.
    const onLine = debounce(async () => {
      if (!input.value.trim()) {
        setChartItems(part, []); part.chart_line = ""; preview.innerHTML = "";
        preview.classList.add("is-empty"); preview.classList.remove("is-stale");
        err.classList.add("hidden"); schedulePreviewUpdate(); return;
      }
      try {
        const result = await API.parseLine(input.value, currentDoc);
        if (result.ok) { setChartItems(part, result.items); part.chart_line = result.unparsed; }
        paint(result);
      } catch (e) { paint({ ok: false, error: String(e) }); }
      schedulePreviewUpdate();
    }, 150);
    input.addEventListener("input", onLine);
    attachAutocomplete(input);
    show();
    if (input.value.trim()) {
      API.parseLine(input.value, currentDoc).then(paint).catch(() => {});
    }
  });
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
    // A group's repeat bracket is drawn afterwards as a real border, by
    // applyBracketRuns — the "|"/"(xN)" text this span marks is only
    // there for a front end with no vector graphics of its own (TXT
    // export, the desktop app), and would only double up with it here.
    if (role !== "bracket") {
      const mark = document.createElement("span");
      mark.className = "tok-" + role;
      mark.textContent = text.slice(from, to);
      el.appendChild(mark);
    }
    pos = to;
  });
  if (pos < text.length) el.appendChild(document.createTextNode(text.slice(pos)));
}

/**
 * The real line a group's repeat bracket is drawn as, in place of the
 * "|"/"(xN)" characters paintRow leaves out for a "bracket"-tagged span
 * (see there) — mirrors export.py's draw_bracket_runs, the PDF's
 * equivalent, including why: a fret row is set smaller, in its own
 * colour, on tighter line spacing (see .sec-preview-fret in the CSS),
 * not a line in its own right, so measuring this in pixels rather than
 * character columns is what lets the line cross one without a kink —
 * character-column math would have to fight that row's own font-size
 * override to land in the same place.
 *
 * `container` is the rows' common positioned ancestor (.sec-preview).
 * `rowEls[i]` must be the element row index i was painted into — the
 * caller builds that list in the same order the server's row indices
 * count in (an optional fret row, then the symbol row, then any extra
 * rows). `brackets` is parseResult.brackets as the server sent it, one
 * entry per group: {start, end, col, rep, label_row}.
 */
function applyBracketRuns(container, rowEls, brackets) {
  [...container.querySelectorAll(":scope > .sec-bracket-line, :scope > .sec-bracket-label")]
    .forEach((el) => el.remove());
  (brackets || []).forEach(({ start, end, col, rep }) => {
    const first = rowEls[start], last = rowEls[end];
    if (!first || !last) return;
    const top = first.offsetTop, bottom = last.offsetTop + last.offsetHeight;
    const line = document.createElement("div");
    line.className = "sec-bracket-line";
    line.style.left = col + "ch";
    line.style.top = top + "px";
    line.style.height = Math.max(0, bottom - top) + "px";
    container.appendChild(line);
    if (rep) {
      // Centered on the run's full span, like the PDF's draw_bracket_runs
      // (its (y_top + y_bot) / 2) - not the "(xN)"-carrying row's own
      // center, which sits off-center whenever the run has more rows
      // below that row than above it (or vice versa).
      const label = document.createElement("span");
      label.className = "sec-bracket-label";
      label.textContent = `(x${rep})`;
      label.style.left = `calc(${col}ch + 8px)`;
      label.style.top = ((top + bottom) / 2) + "px";
      container.appendChild(label);
    }
  });
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

/**
 * The parser's message, starting with a capital letter and without the
 * "ParseError:" a thrown error carries. The wording itself is the
 * parser's (grammar.py), so the CLI's lint says the same thing.
 */
function friendlyError(msg) {
  const text = String(msg || "").replace(/^(Error|ParseError):\s*/, "").trim();
  return text ? text[0].toUpperCase() + text.slice(1) : "This line could not be read.";
}

function updateSectionPreview(node, sec, parseResult) {
  const fretRow = node.querySelector(".sec-preview-fret");
  const symRow = node.querySelector(".sec-preview-sym");
  const errEl = node.querySelector(".sec-error");
  const emptyEl = node.querySelector(".sec-empty-hint");

  const items = (parseResult && parseResult.items) || chartItems(sec);
  const rendered = (parseResult && parseResult.rendered) ||
    (items.length ? null : []); // null = "not recomputed, leave last render"

  const previewBox = node.querySelector(".sec-preview");
  const lineEl = node.querySelector(".sec-chart-line");
  if (parseResult && parseResult.error) {
    errEl.textContent = friendlyError(parseResult.error);
    errEl.classList.remove("hidden");
    lineEl.setAttribute("aria-invalid", "true");
    // The last chart that worked stays on screen, dimmed and labelled, so
    // it is never mistaken for what the line says now.
    previewBox.classList.add("is-stale");
    return; // leave the last valid render on screen — never clear it
  }
  errEl.classList.add("hidden");
  lineEl.removeAttribute("aria-invalid");
  previewBox.classList.remove("is-stale");

  // Shown before painting: the bracket lines are measured from the rows.
  if (rendered) previewBox.classList.toggle("is-empty", !rendered.length);
  if (rendered && rendered.length) {
    // A row is 1..N lines now: an optional fret row, the symbol row, and
    // one line per string for any lick. The first two elements are reused
    // so the fret/symbol colouring stays; anything beyond is appended.
    const roles = (parseResult && parseResult.roles) || [];
    const spans = (parseResult && parseResult.spans) || [];
    const row = (i) => ({ text: rendered[i] || "", role: roles[i], spans: spans[i] });
    if (roles.length === rendered.length) {
      // Every row painted by its own role: the marks row (endings, runs,
      // accents), frets, symbols, tab. The two fixed rows stay empty.
      fretRow.textContent = "";
      symRow.textContent = "";
      setExtraPreviewRows(node, rendered.map((_, i) => row(i)));
    } else {
      const hasFret = parseResult ? !!parseResult.fret_row : rendered.length > 1;
      fretRow.textContent = hasFret ? rendered[0] : "";
      paintRow(symRow, hasFret ? row(1) : row(0));
      setExtraPreviewRows(node, rendered.slice(hasFret ? 2 : 1)
        .map((_, i) => row(i + (hasFret ? 2 : 1))));
    }
    // rowEls[i] is whatever row index i was just painted into, in the
    // same order the server counted rows in — applyBracketRuns measures
    // these to draw a group's repeat bracket as a real line.
    const rowEls = roles.length === rendered.length
      ? [...node.querySelectorAll(".sec-preview-more")]
      : [...(fretRow.textContent ? [fretRow] : []), symRow,
         ...node.querySelectorAll(".sec-preview-more")];
    applyBracketRuns(previewBox, rowEls, parseResult && parseResult.brackets);
    emptyEl.classList.add("hidden");
  } else if (!items.length && !measureItems(sec).length) {
    fretRow.textContent = "";
    symRow.textContent = "";
    setExtraPreviewRows(node, []);
    applyBracketRuns(previewBox, [], []);
    emptyEl.classList.remove("hidden");
  } else if (rendered) {
    fretRow.textContent = "";
    symRow.textContent = "";
    setExtraPreviewRows(node, []);
    applyBracketRuns(previewBox, [], []);
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

/** The section's own grid body (not one inside a part). */
function sectionGridBody(node) {
  return node.querySelector(":scope > .sec-tab-grid .tab-grid-body");
}

function rebuildTabGrid(node, sec, body) {
  body = body || sectionGridBody(node);
  const measures = measureItems(sec);
  const strings = stringsForInstrument(sec.instrument);

  if (!measures.length) {
    body.innerHTML = '<p class="tab-empty">No measures yet. Click <strong>+ Measure</strong> to add one.</p>';
    return;
  }
  // Beats just added by +2 / +4: tinted for a moment, so you see where they went.
  const fresh = body._fresh || null;
  body._fresh = null;

  const table = document.createElement("table");
  table.className = "tab-grid-table";

  // Header row: measure numbers with delete buttons
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  headerRow.appendChild(document.createElement("th")); // empty corner
  measures.forEach((m, mi) => {
    const th = document.createElement("th");
    th.colSpan = m.beats || 8;
    const beats = m.beats || 8;
    const hdr = document.createElement("span");
    hdr.className = "tab-measure-header";
    const name = document.createElement("span");
    name.className = "tab-measure-name";
    name.textContent = `M${mi + 1} · ${beats}`;
    name.title = `Measure ${mi + 1}: ${beats} beats`;
    hdr.appendChild(name);
    // Room for more notes: +2 / +4 add empty beats to every string of this
    // measure only, so the other measures keep their spacing; − takes
    // empty beats off the end again.
    const mk = (label, title, fn, disabled) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "tab-size-btn";
      b.textContent = label;
      b.title = title;
      b.setAttribute("aria-label", title);
      b.disabled = !!disabled;
      b.addEventListener("click", fn);
      hdr.appendChild(b);
    };
    mk("+2", `Add 2 beats to measure ${mi + 1}`,
       () => resizeMeasure(node, sec, body, mi, 2), beats >= TAB_BEATS_MAX);
    mk("+4", `Add 4 beats to measure ${mi + 1}`,
       () => resizeMeasure(node, sec, body, mi, 4), beats >= TAB_BEATS_MAX);
    mk("−", `Remove empty beats from the end of measure ${mi + 1}`,
       () => resizeMeasure(node, sec, body, mi, -2), emptyTail(m, strings) < 1 || beats <= 1);
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "tab-del-measure";
    delBtn.title = "Delete measure " + (mi + 1);
    delBtn.setAttribute("aria-label", "Delete measure " + (mi + 1));
    delBtn.innerHTML = '<svg class="ico" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M6 6l12 12M18 6L6 18"/></svg>';
    delBtn.addEventListener("click", () => {
      const allMeasures = measureItems(sec);
      allMeasures.splice(mi, 1);
      setMeasureItems(sec, allMeasures);
      rebuildTabGrid(node, sec, body);
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
        input.className = "tab-cell" + (bi === beats - 1 ? " tab-cell-last" : "");
        if (fresh && fresh.mi === mi && bi >= fresh.from) input.classList.add("tab-cell-new");
        input.setAttribute("aria-label", `${st} string, measure ${mi + 1}, beat ${bi + 1}`);
        input.value = tokens[bi] || "-";
        // A cell with a technique ("7b9r7") grows to show all of it.
        const fit = () => { input.style.width = Math.max(32, input.value.length * 8 + 12) + "px"; };
        fit();
        input.addEventListener("input", fit);
        // Room for a technique: 5h7, 7b9r7, 12h14^.
        input.maxLength = 8;
        input.title = "A fret, - for not played, x for muted; techniques: "
          + "5h7 hammer-on, 7p5 pull-off, 5/7 or 7\\5 slide, 7b9 bend, 7b9r7 "
          + "bend and release, 7~ vibrato, 5^ accent";
        input.addEventListener("input", () => {
          commitTabCell(sec, mi, st, beats, body);
        });
        input.addEventListener("focus", () => input.select());
        // Arrow key navigation between cells
        input.addEventListener("keydown", (e) => {
          if (e.key === "ArrowRight" || (e.key === "Tab" && !e.shiftKey)) {
            const next = td.nextElementSibling;
            if (next && next.querySelector("input")) {
              e.preventDefault();
              next.querySelector("input").focus();
            } else if (e.key === "ArrowRight" && mi === measures.length - 1
                       && beats < TAB_BEATS_MAX) {
              // → past the very last cell: the last measure grows by two
              // beats and the caret moves into the first new one.
              e.preventDefault();
              const row = strings.indexOf(st);
              resizeMeasure(node, sec, body, mi, 2);
              const cells = body.querySelectorAll("tbody tr")[row]
                ?.querySelectorAll("input.tab-cell-new");
              if (cells && cells.length) cells[0].focus();
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

const TAB_BEATS_MAX = 64;

/** A measure's tokens for one string, padded with "-" to its length. */
function measureTokens(m, st) {
  const beats = m.beats || 8;
  const tokens = ((m.strings || {})[st] || "").split(/\s+/).filter(Boolean);
  while (tokens.length < beats) tokens.push("-");
  return tokens.slice(0, beats);
}

/** How many beats at the end of a measure are empty on every string. */
function emptyTail(m, strings) {
  const beats = m.beats || 8;
  let n = 0;
  for (let bi = beats - 1; bi >= 0; bi--) {
    const empty = strings.every((st) => {
      const t = measureTokens(m, st)[bi];
      return !t || t === "-";
    });
    if (!empty) break;
    n += 1;
  }
  return n;
}

/**
 * Make one measure longer (delta > 0: empty beats on every string at its
 * end) or shorter (delta < 0: only beats that are empty on every string
 * are taken off, so no note is ever lost). The other measures keep their
 * length. Tokens are written out in full so every string stays aligned.
 */
function resizeMeasure(node, sec, body, mi, delta) {
  const measures = measureItems(sec);
  const m = measures[mi];
  if (!m) return;
  const strings = stringsForInstrument(sec.instrument);
  const beats = m.beats || 8;
  let next = beats + delta;
  if (delta < 0) next = Math.max(beats - Math.min(-delta, emptyTail(m, strings)), 1);
  next = Math.min(Math.max(next, 1), TAB_BEATS_MAX);
  if (next === beats) return;
  const out = {};
  strings.forEach((st) => {
    const tokens = measureTokens(m, st);
    while (tokens.length < next) tokens.push("-");
    out[st] = tokens.slice(0, next).join(" ");
  });
  m.beats = next;
  m.strings = Object.assign({}, m.strings || {}, out);
  setMeasureItems(sec, measures);
  if (next > beats) body._fresh = { mi, from: beats };
  rebuildTabGrid(node, sec, body);
  schedulePreviewUpdate();
}

function commitTabCell(sec, measureIdx, stringName, beats, body) {
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

/** Scroll a section's card into view and flash it, so you see where it went. */
function showSectionCard(id) {
  const card = document.querySelector(`.section-card[data-id="${CSS.escape(id)}"]`);
  if (!card) return null;
  card.scrollIntoView({ block: "nearest" });
  card.classList.remove("just-moved");
  void card.offsetWidth;            // restart the animation
  card.classList.add("just-moved");
  return card;
}

// ---------------------------------------------------------------------------
//  Reordering by dragging the ⠿ handle
// ---------------------------------------------------------------------------
//  Pointer events rather than HTML drag-and-drop: they behave the same in a
//  browser tab and in the native window, work with a trackpad or a finger,
//  and leave the text fields inside each card alone.
//
//  While you drag, every card folds down to its header row, so the whole
//  song fits on screen and a section can go from the end to the top in one
//  move. The card you hold moves through the list as you go; letting go
//  keeps the new order, Escape puts everything back.
let sectionDrag = null;

function wireDragHandle(node, id) {
  const handle = node.querySelector(".drag-handle");
  if (!handle) return;
  handle.addEventListener("pointerdown", (e) => {
    if (e.button !== 0 || sectionDrag) return;
    e.preventDefault();
    sectionDrag = { id, node, startY: e.clientY, lastY: e.clientY, started: false,
                    before: currentDoc.sections.map((s) => s.id), raf: 0 };
    window.addEventListener("pointermove", onSectionDragMove);
    window.addEventListener("pointerup", onSectionDragUp);
    window.addEventListener("pointercancel", onSectionDragCancel);
    window.addEventListener("keydown", onSectionDragKey, true);
  });
  // The keyboard way: focus the handle, then up and down arrows.
  handle.addEventListener("keydown", (e) => {
    if (e.key !== "ArrowUp" && e.key !== "ArrowDown") return;
    e.preventDefault();
    moveSection(id, e.key === "ArrowUp" ? -1 : 1);
    const card = showSectionCard(id);
    const again = card && card.querySelector(".drag-handle");
    if (again) again.focus();
  });
}

function startSectionDrag() {
  const d = sectionDrag;
  const list = document.getElementById("section-list");
  const main = document.getElementById("main");
  d.started = true;
  document.body.classList.add("dragging-section");
  list.classList.add("reordering");
  d.node.classList.add("lifted");
  // Folding the cards moves everything; scroll so the held card is back
  // under the pointer.
  const r = d.node.getBoundingClientRect();
  main.scrollTop += (r.top + r.height / 2) - d.lastY;
  const tick = () => {
    if (!sectionDrag || !sectionDrag.started) return;
    autoScrollWhileDragging();
    placeDraggedSection();
    sectionDrag.raf = requestAnimationFrame(tick);
  };
  d.raf = requestAnimationFrame(tick);
}

function onSectionDragMove(e) {
  const d = sectionDrag;
  if (!d) return;
  d.lastY = e.clientY;
  if (!d.started && Math.abs(e.clientY - d.startY) > 4) startSectionDrag();
}

// Near the top or bottom edge of the editor, keep scrolling, faster the
// closer the pointer gets.
function autoScrollWhileDragging() {
  const main = document.getElementById("main");
  const r = main.getBoundingClientRect();
  const edge = 48, y = sectionDrag.lastY;
  if (y < r.top + edge) main.scrollTop -= Math.ceil((r.top + edge - y) / 4);
  else if (y > r.bottom - edge) main.scrollTop += Math.ceil((y - (r.bottom - edge)) / 4);
}

// Put the held card before the first other card whose middle is below the
// pointer, or last if there is none.
function placeDraggedSection() {
  const d = sectionDrag;
  const list = document.getElementById("section-list");
  const others = [...list.children].filter((c) => c !== d.node);
  const target = others.find((c) => {
    const r = c.getBoundingClientRect();
    return d.lastY < r.top + r.height / 2;
  }) || null;
  if (target === null) {
    if (list.lastElementChild !== d.node) list.appendChild(d.node);
  } else if (d.node.nextElementSibling !== target) {
    list.insertBefore(d.node, target);
  }
}

function onSectionDragUp() { endSectionDrag(true); }
function onSectionDragCancel() { endSectionDrag(false); }
function onSectionDragKey(e) {
  if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); endSectionDrag(false); }
}

function endSectionDrag(keep) {
  const d = sectionDrag;
  if (!d) return;
  sectionDrag = null;
  cancelAnimationFrame(d.raf);
  window.removeEventListener("pointermove", onSectionDragMove);
  window.removeEventListener("pointerup", onSectionDragUp);
  window.removeEventListener("pointercancel", onSectionDragCancel);
  window.removeEventListener("keydown", onSectionDragKey, true);
  if (!d.started) return;          // a click on the handle, not a drag
  document.body.classList.remove("dragging-section");
  const list = document.getElementById("section-list");
  list.classList.remove("reordering");
  d.node.classList.remove("lifted");
  const after = [...list.children].map((c) => c.dataset.id);
  const changed = keep && after.join("\n") !== d.before.join("\n");
  if (changed) {
    const byIdMap = new Map(currentDoc.sections.map((s) => [s.id, s]));
    currentDoc.sections = after.map((id) => byIdMap.get(id)).filter(Boolean);
    schedulePreviewUpdate();
  }
  renderEditor();                  // also puts the cards back on Escape
  showSectionCard(d.id);
}

// ---------------------------------------------------------------------------
//  "Copy to…" picker
// ---------------------------------------------------------------------------
//  Copy as a reference… (a card's ⋯ menu) opens a small list of every place the copy can go: the top
//  of the song, or after any section. One click puts it there. The place
//  right after the original is highlighted, so Enter does what the button
//  used to do.
function openCopyPicker(id, anchor) {
  const secs = currentDoc.sections;
  const i = secs.findIndex((s) => s.id === id);
  if (i < 0) return;
  const picker = document.getElementById("copy-picker");
  const label = (s, n) => `${n + 1}. ${s.name || "(no name)"}`;
  picker.querySelector(".copy-picker-title").textContent =
    `Copy "${secs[i].name || "(no name)"}" to…`;
  const list = picker.querySelector(".copy-picker-list");
  list.innerHTML = "";
  const addRow = (text, at, isDefault) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "copy-picker-row" + (isDefault ? " default" : "");
    b.textContent = text;
    b.addEventListener("click", () => {
      closeCopyPicker();
      const sec = duplicateSection(id, at);
      if (sec) toast(at === 0 ? `Added "${sec.name}" at the top`
                              : `Added "${sec.name}" after "${secs[at - 1].name}"`);
    });
    list.appendChild(b);
    return b;
  };
  addRow("At the top", 0, false);
  let focusMe = null;
  secs.forEach((s, n) => {
    const own = n === i;
    const row = addRow(`After ${label(s, n)}${own ? "  (this one)" : ""}`, n + 1, own);
    if (own) focusMe = row;
  });

  picker.querySelector(".copy-picker-cancel").onclick = closeCopyPicker;
  picker.classList.remove("hidden");
  // Next to the button, kept inside the window.
  const a = anchor.getBoundingClientRect();
  const w = picker.offsetWidth, h = picker.offsetHeight;
  let left = Math.min(a.right - w, window.innerWidth - w - 8);
  let top = a.bottom + 6;
  if (top + h > window.innerHeight - 8) top = Math.max(8, a.top - h - 6);
  picker.style.left = `${Math.max(8, left)}px`;
  picker.style.top = `${top}px`;
  if (focusMe) { focusMe.focus(); focusMe.scrollIntoView({ block: "nearest" }); }

  setTimeout(() => {
    document.addEventListener("pointerdown", onCopyPickerOutside, true);
    document.addEventListener("keydown", onCopyPickerKey, true);
  }, 0);
}

function closeCopyPicker() {
  document.getElementById("copy-picker").classList.add("hidden");
  document.removeEventListener("pointerdown", onCopyPickerOutside, true);
  document.removeEventListener("keydown", onCopyPickerKey, true);
}

function onCopyPickerOutside(e) {
  if (!document.getElementById("copy-picker").contains(e.target)) closeCopyPicker();
}

function onCopyPickerKey(e) {
  if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closeCopyPicker(); return; }
  if (e.key === "ArrowDown" || e.key === "ArrowUp") {
    const rows = [...document.querySelectorAll("#copy-picker .copy-picker-row")];
    const k = rows.indexOf(document.activeElement);
    const next = rows[Math.max(0, Math.min(rows.length - 1, k + (e.key === "ArrowDown" ? 1 : -1)))];
    if (next) { e.preventDefault(); next.focus(); }
  }
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

/**
 * Add a reference to section `id`. `at` is where it goes in the list
 * (0 = top); left out, it goes right after the original, as before.
 */
function duplicateSection(id, at) {
  const i = currentDoc.sections.findIndex((s) => s.id === id);
  if (i < 0) return;
  const source = currentDoc.sections[i];
  const where = (at === undefined || at === null)
    ? i + 1 : Math.max(0, Math.min(currentDoc.sections.length, at));
  const newSec = {
    id: newSectionId(), name: source.name + " (ref)", type: source.type,
    instrument: source.instrument, repeat: 1, transpose: 0,
    render: source.render === "free" ? "chart" : source.render, annotation: "",
    free_text: "",
    // The id is what's stored (rename-safe); the name is what's shown.
    items: [{ kind: "section_ref", section: source.id, repeat: 1, all: false, transpose: 0 }],
    chart_line: `=${refKeyFor(source)}`,
  };
  currentDoc.sections.splice(where, 0, newSec);
  renderEditor();
  schedulePreviewUpdate();
  showSectionCard(newSec.id);
  return newSec;
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
  // The server saves under the song's own name — "Artist - Title.sng" —
  // and renames the file when the title or artist has changed, so the
  // name to use from here on is the one it answers with.
  const res = await API.saveSong(currentFilename, currentDoc);
  const saved = (res && res.filename) || currentFilename;
  const renamed = saved !== currentFilename;
  currentFilename = saved;
  markSaved();
  toast(renamed ? `Saved. The file is now "${saved}"` : `Saved ${saved}`);
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
  setTimeout(() => showClosedOverlay("Closed. You can close this window now."), 600);
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
      ? "Opened in your PDF viewer. Print from there."
      : `Print failed: ${(res && res.error) || "unknown error"}`);
    return;
  }
  const res = await fetch("/api/export.pdf", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ doc: currentDoc, orient: "portrait" }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    toast(`Print failed: ${err.error || "could not build the PDF"}`);
    return;
  }
  const url = URL.createObjectURL(await res.blob());
  const win = window.open(url, "_blank");
  if (!win) {
    toast("Allow pop-ups to print, or export the PDF and print that.");
  } else {
    toast("Opened in a new tab. Print from there.");
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
      toast(`Saved to ${res.path}` + exportWarningText(res.warnings));
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
  let warnings = [];
  try { warnings = JSON.parse(res.headers.get("X-Export-Warnings") || "[]"); }
  catch (e) { warnings = []; }
  toast(`Downloaded ${filename}. It is in your browser's downloads folder.`
        + exportWarningText(warnings));
}

/** The check's warnings, as a sentence to add to an export message. */
function exportWarningText(warnings) {
  return (warnings && warnings.length) ? ` Note: ${warnings.join("; ")}.` : "";
}

// ---------------------------------------------------------------------------
//  Songs folder — shown in the sidebar so "where is my file?" is answerable
//  without leaving the app.
// ---------------------------------------------------------------------------
function renderSongsFolder() {
  const box = document.getElementById("songs-folder");
  if (!box) return;
  const dir = META.songs_dir || "";
  // Break only after a "/" or "\\", never inside a folder name.
  const pathEl = box.querySelector(".songs-folder-path");
  pathEl.textContent = "";
  dir.split(/(?<=[\/\\])/).forEach((part, i) => {
    if (i) pathEl.appendChild(document.createElement("wbr"));
    pathEl.appendChild(document.createTextNode(part));
  });
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

let folderPlan = null;

async function changeSongsFolder() {
  const api = nativeApi();
  if (!api || !api.choose_songs_folder) return;
  const res = await api.choose_songs_folder();
  if (!res || !res.ok) {
    if (res && !res.cancelled) toast(`Could not change folder: ${res.error || "?"}`);
    return;
  }
  const plan = res.plan || {};
  if (plan.same) { toast("That's already your songs folder"); return; }
  if (!plan.writable) { toast(`Can't use that folder: ${plan.error}`); return; }
  folderPlan = plan;

  // Say what will happen before anything does: how many songs move, and
  // which ones stay because a song of the same name is already there.
  document.getElementById("folder-from").textContent = plan.current;
  document.getElementById("folder-to").textContent = plan.new;
  const n = (plan.to_move || []).length;
  const there = plan.already_there || 0;
  document.getElementById("folder-summary").textContent =
    (n ? `${n} song${n === 1 ? "" : "s"} will move to the new folder.`
       : "There are no songs to move.")
    + (there ? ` It already has ${there} song${there === 1 ? "" : "s"} of its own, which will show up too.` : "");
  const clashes = plan.clashes || [];
  const cl = document.getElementById("folder-clashes");
  cl.classList.toggle("hidden", !clashes.length);
  cl.textContent = clashes.length
    ? `Staying where they are, because a song with the same name is already in the new folder: ${clashes.join(", ")}`
    : "";
  const moveBtn = document.getElementById("folder-move");
  moveBtn.textContent = n ? `Move ${n} song${n === 1 ? "" : "s"}` : "Use this folder";
  document.getElementById("folder-use").classList.toggle("hidden", !n);
  document.getElementById("folder-modal-backdrop").classList.remove("hidden");
}

function closeFolderModal() {
  document.getElementById("folder-modal-backdrop").classList.add("hidden");
  folderPlan = null;
}

async function applySongsFolder(move) {
  const api = nativeApi();
  const plan = folderPlan;
  closeFolderModal();
  if (!api || !plan) return;
  const res = await api.set_songs_folder(plan.new, move);
  if (!res || !res.ok) {
    toast(`Could not change folder: ${(res && res.error) || "?"}`);
    return;
  }
  META.songs_dir = res.path;
  renderSongsFolder();

  // The open song stays open if it came along: same file name, new
  // folder, and any unsaved edits still in the editor. If it didn't come
  // (a clash, or "just use this folder"), close it rather than let a
  // later Save quietly write a copy into the new folder.
  const moved = res.moved || [];
  if (!(currentFilename && moved.includes(currentFilename))) {
    currentFilename = null;
    currentDoc = null;
    document.getElementById("editor").classList.add("hidden");
    document.getElementById("start-here").classList.remove("hidden");
  }
  await refreshSongList();
  const skipped = res.skipped || [];
  toast(move
    ? `Moved ${moved.length} song${moved.length === 1 ? "" : "s"} to ${res.path}`
      + (skipped.length ? `; ${skipped.length} left where they were` : "")
    : `Songs folder is now ${res.path}`);
}

// ---------------------------------------------------------------------------
//  Live preview pane
// ---------------------------------------------------------------------------
const schedulePreviewUpdate = debounce(async () => {
  updateSaveStatus();
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
  document.getElementById("btn-preview-toggle").setAttribute("aria-pressed", String(previewOpen));
  if (previewOpen) schedulePreviewUpdate();
}

// ---------------------------------------------------------------------------
//  Song tools, saved state, dialogs, theme
// ---------------------------------------------------------------------------

/** The buttons that act on a song do nothing without one, so they say so. */
function setSongToolsEnabled(on) {
  document.querySelectorAll("#topbar .song-tool").forEach((b) => { b.disabled = !on; });
  if (!on) togglePreview(false);
}

/** A song was opened or created: the editor shows, the tools wake up. */
function showEditor() {
  document.getElementById("start-here").classList.add("hidden");
  document.getElementById("editor").classList.remove("hidden");
  setSongToolsEnabled(true);
  closeSidebarOnNarrow();
}

// "Saved" against "Not saved": the song as it was last read or written,
// compared with the song now. A song is small, so a string compare after
// each edit (debounced) costs nothing worth measuring.
let savedSnapshot = null;
function markSaved() {
  savedSnapshot = currentDoc ? JSON.stringify(currentDoc) : null;
  updateSaveStatus();
}
function updateSaveStatus() {
  const el = document.getElementById("save-status");
  if (!el) return;
  if (!currentDoc) { el.textContent = ""; el.classList.remove("unsaved"); return; }
  const unsaved = JSON.stringify(currentDoc) !== savedSnapshot;
  el.textContent = unsaved ? "Not saved" : "Saved";
  el.classList.toggle("unsaved", unsaved);
}
const scheduleSaveStatus = debounce(updateSaveStatus, 300);

/**
 * Every dialog behaves the same way: Escape closes it (through its own
 * Cancel or Done button, so whatever that button does still happens),
 * focus moves into it when it opens and back where it was when it closes.
 */
function topmostDialog() {
  const open = [...document.querySelectorAll(".modal-backdrop")]
    .filter((b) => !b.classList.contains("hidden"));
  if (!open.length) return null;
  const z = (b) => parseInt(getComputedStyle(b).zIndex, 10) || 0;
  return open.reduce((a, b) => (z(b) >= z(a) ? b : a));
}
function wireDialogs() {
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    const top = topmostDialog();
    if (!top) return;
    const btn = top.dataset.close && document.getElementById(top.dataset.close);
    if (!btn) return;
    e.preventDefault();
    btn.click();
  });
  document.querySelectorAll(".modal-backdrop").forEach((backdrop) => {
    let returnTo = null;
    new MutationObserver(() => {
      const open = !backdrop.classList.contains("hidden");
      if (open && !backdrop._open) {
        returnTo = document.activeElement;
        const first = backdrop.querySelector(
          "input:not([type=hidden]):not(.hidden):not([tabindex='-1']), textarea, select:not(.hidden), button");
        if (first) setTimeout(() => first.focus(), 0);
      } else if (!open && backdrop._open) {
        if (returnTo && document.contains(returnTo) && !returnTo.disabled) returnTo.focus();
        returnTo = null;
      }
      backdrop._open = open;
    }).observe(backdrop, { attributes: true, attributeFilter: ["class"] });
  });
  // Tab stays inside an open dialog.
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Tab") return;
    const top = topmostDialog();
    if (!top) return;
    const items = [...top.querySelectorAll(
      "button, input, select, textarea, [tabindex]:not([tabindex='-1'])")]
      .filter((el) => !el.disabled && el.offsetParent !== null && el.tabIndex !== -1);
    if (!items.length) return;
    const first = items[0], last = items[items.length - 1];
    if (!top.contains(document.activeElement)) { e.preventDefault(); first.focus(); return; }
    if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
  });
}

/** Light or dark: the system's until you choose, then yours (this computer only). */
const THEME_KEY = "snt-theme";
function applyTheme(theme) {
  if (theme === "light" || theme === "dark") document.documentElement.dataset.theme = theme;
  else delete document.documentElement.dataset.theme;
}
function currentTheme() {
  const set = document.documentElement.dataset.theme;
  if (set) return set;
  return window.matchMedia && window.matchMedia("(prefers-color-scheme: light)").matches
    ? "light" : "dark";
}
function wireTheme() {
  let saved = null;
  try { saved = localStorage.getItem(THEME_KEY); } catch (e) { /* private window */ }
  applyTheme(saved);
  document.getElementById("btn-theme").addEventListener("click", () => {
    const next = currentTheme() === "light" ? "dark" : "light";
    applyTheme(next);
    try { localStorage.setItem(THEME_KEY, next); } catch (e) { /* not kept */ }
  });
}

/** On a phone-sized window the song list is a panel the menu button opens. */
function closeSidebarOnNarrow() {
  document.body.classList.remove("sidebar-open");
  const b = document.getElementById("btn-sidebar");
  if (b) b.setAttribute("aria-expanded", "false");
}
function wireSidebarToggle() {
  const b = document.getElementById("btn-sidebar");
  b.addEventListener("click", (e) => {
    e.stopPropagation();
    const open = !document.body.classList.contains("sidebar-open");
    document.body.classList.toggle("sidebar-open", open);
    b.setAttribute("aria-expanded", String(open));
  });
  document.addEventListener("mousedown", (e) => {
    if (!document.body.classList.contains("sidebar-open")) return;
    if (e.target.closest("#sidebar") || e.target.closest("#btn-sidebar")) return;
    if (e.target.closest(".layout-panel")) return;
    closeSidebarOnNarrow();
  });
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
  toast(`Transposed ${label} by ${n > 0 ? "+" : ""}${n} semitone(s). Remember to save.`);
}

// ---------------------------------------------------------------------------
//  Lyrics modal — stored non-destructively on the document or a section as
//  `lyrics_text`, never parsed or aligned to the chart. Whether and where
//  it prints is one song-wide setting, `lyrics_layout`, in the Layout
//  panel (off by default). "Search online" opens a browser search in
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
      b.title = hasLyrics ? "Has lyrics. Click to edit."
        : "No lyrics yet. Click to add.";
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
    toast("No === Section === markers in the sheet yet. Drag a section in.");
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

  // Placement is the Layout panel's business now; the flag is still set
  // so a file opened in an older version prints what this one would.
  const print = (currentDoc.lyrics_layout || "none") !== "none";
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
  toast(`Lyrics assigned to ${assigned} section(s). Remember to save.`);
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
  const lyr = currentDoc.lyrics_layout || "none";
  document.getElementById("lyrics-where-text").textContent = lyr === "none"
    ? "Lyrics don't print for this song yet."
    : `Printing: ${optionLabel("meta-lyrics-layout").toLowerCase()}.`;
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
    empty.textContent = "No shapes yet. \u201c+ Chord\u201d adds one.";
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
  shape.title = "One fret per string, lowest string first. x means the string is not played.";

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
  note.title = "A word printed under the diagram, such as \u201cbarre\u201d or \u201cthumb\u201d: whatever you need to remember";
  note.addEventListener("input", () => { chord.note = note.value; renderChordPreview(); });

  const del = document.createElement("button");
  del.type = "button";
  del.className = "icon-btn";
  del.innerHTML = '<svg class="ico" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/></svg>';
  del.title = "Remove this shape";
  del.setAttribute("aria-label", "Remove this shape");
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
        + "open shape of the new chord, or to its barre if there is none, and "
        + "a barre slides."
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
    ? `Added ${missing.length} chord(s), as they read before the transpose. Type a shape for each.`
    : `Added ${missing.length} chord(s). Type a shape for each.`));
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
  opt("", "(no lyrics)");
  // Words this section already holds that aren't in the sheet came from
  // somewhere else — a hand-typed line, an older split. Offer to leave
  // them alone rather than quietly overwriting them.
  const held = (sec.lyrics_text || "").trim();
  const heldIsBlock = splitBlocks.some((b) => b.trim() === held);
  if (held && !heldIsBlock) opt(KEEP, "(keep what's here)");
  splitBlocks.forEach((b, i) => opt(String(i), blockLabel(b, i)));

  sel.value = choice === null || choice === undefined ? "" : String(choice);
  row.appendChild(name);
  row.appendChild(sel);
  return row;
}

async function splitLyricsIntoSections() {
  const text = document.getElementById("lyrics-textarea").value;
  if (!text.trim()) { toast("Nothing to split. Paste some lyrics first."); return; }
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
  if (!splitBlocks.length) { toast("Nothing to split. Paste some lyrics first."); return; }

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
  document.getElementById("lyrics-split-backdrop").classList.remove("hidden");
}

function closeSplitModal() {
  document.getElementById("lyrics-split-backdrop").classList.add("hidden");
}

function applyLyricsSplit() {
  // Where the words print is Layout's call; the per-section flag is
  // still written so an older version opening the file agrees with it.
  const print = (currentDoc.lyrics_layout || "none") !== "none";
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
  toast(`Lyrics assigned to ${assigned} section(s). Remember to save.`);
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
    toast("Your browser blocked the new tab. Allow pop-ups for this page, or copy the link.");
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
// ---------------------------------------------------------------------------
//  Riffs: the library of named runs, kept in the song (doc.blocks). Every
//  change goes through the server (songmap), which returns the whole song,
//  so the rules (names, references, what is still in use) live in one place.
// ---------------------------------------------------------------------------

async function riffOp(op, fields) {
  const msg = document.getElementById("riffs-msg");
  const res = await API.riffs(op, fields);
  if (!res.ok) {
    if (msg) msg.textContent = res.error || "That did not work.";
    toast(res.error || "That did not work.");
    return null;
  }
  if (msg) msg.textContent = "";
  currentDoc = res.doc;
  renderEditor();
  renderRiffRows(res.riffs || []);
  schedulePreviewUpdate();
  return res;
}

function renderRiffRows(riffs) {
  const box = document.getElementById("riffs-rows");
  box.innerHTML = "";
  if (!riffs.length) {
    box.innerHTML = '<p class="modal-hint">No riffs yet.</p>';
    return;
  }
  riffs.forEach((r) => {
    const row = document.createElement("div");
    row.className = "riff-row";
    const name = document.createElement("input");
    name.type = "text"; name.value = r.id; name.className = "riff-name";
    name.title = "The riff's name, as typed on a chart line. Change it to rename the riff everywhere.";
    name.spellcheck = false;
    const line = document.createElement("input");
    line.type = "text"; line.value = r.line; line.className = "riff-line";
    line.title = "What the riff plays";
    line.spellcheck = false;
    const used = document.createElement("span");
    used.className = "riff-used";
    used.textContent = r.used_in.length ? `in ${r.used_in.join(", ")}` : "not used yet";
    const del = document.createElement("button");
    del.type = "button"; del.className = "icon-btn"; del.innerHTML = '<svg class="ico" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/></svg>';
    del.setAttribute("aria-label", `Delete the riff ${r.id}`);
    del.disabled = r.used_in.length > 0;
    del.title = r.used_in.length
      ? `Still used in ${r.used_in.join(", ")}: take it out of those lines first`
      : "Delete this riff";
    name.addEventListener("change", () => {
      if (name.value.trim() && name.value.trim() !== r.id) {
        riffOp("rename", { id: r.id, name: name.value.trim() });
      }
    });
    line.addEventListener("change", () => riffOp("set_line", { id: r.id, line: line.value }));
    del.addEventListener("click", () => riffOp("delete", { id: r.id }));
    row.append(name, line, used, del);
    box.appendChild(row);
  });
}

async function openRiffsModal() {
  if (!currentDoc) { toast("Open a song first"); return; }
  document.getElementById("riffs-msg").textContent = "";
  const res = await API.riffs("list");
  renderRiffRows((res && res.riffs) || []);
  document.getElementById("riffs-modal-backdrop").classList.remove("hidden");
}

function closeRiffsModal() {
  document.getElementById("riffs-modal-backdrop").classList.add("hidden");
}

async function addRiff() {
  const name = document.getElementById("riff-new-name");
  const line = document.getElementById("riff-new-line");
  const res = await riffOp("create", { name: name.value, line: line.value });
  if (res) { name.value = ""; line.value = ""; }
}

/** "→ riff": the selected part of a chart line becomes a riff, named for
 *  you (riff1, riff2…; rename it in Riffs), and its name takes its place. */
async function promoteSelection(input, sec) {
  if (!input || !sec) return;
  const start = input.selectionStart, end = input.selectionEnd;
  if (start == null || end == null || start === end) {
    toast("Select the part of the chart line to turn into a riff first.");
    return;
  }
  const res = await riffOp("promote", { section_id: sec.id, line: input.value, start, end });
  if (res) toast("Made a riff from the selection. Rename it in Riffs if you like.");
}

function wireRiffs() {
  document.getElementById("btn-riffs").addEventListener("click", openRiffsModal);
  document.getElementById("riffs-close").addEventListener("click", closeRiffsModal);
  document.getElementById("riff-add").addEventListener("click", addRiff);
}

// ---------------------------------------------------------------------------
//  Autocomplete on chart lines: "=" offers the song's sections, "{" its
//  named licks, and a plain word the riffs. Up/Down to choose, Enter or Tab
//  to take one, Esc to close.
// ---------------------------------------------------------------------------

function refSpelling(name) {
  return (name || "").trim().replace(/[\s-]+/g, "_");
}

function autocompleteOptions(word) {
  if (!currentDoc || !word) return [];
  const lower = word.toLowerCase();
  let pool = [];
  if (word.startsWith("=")) {
    pool = currentDoc.sections.map((s) => "=" + refSpelling(s.name))
      .filter((w) => /^=[A-Za-z][A-Za-z0-9_]*$/.test(w));
  } else if (word.startsWith("{")) {
    pool = namedLicks().map((n) => `{${n}}`);
  } else if (/^[A-Za-z][A-Za-z0-9_]*$/.test(word)) {
    pool = Object.keys(currentDoc.blocks || {});
  }
  return [...new Set(pool)]
    .filter((w) => w.toLowerCase().startsWith(lower) && w.toLowerCase() !== lower)
    .slice(0, 8);
}

function attachAutocomplete(input) {
  if (!input || input.dataset.autocomplete) return;
  input.dataset.autocomplete = "1";
  const list = document.createElement("div");
  list.className = "autocomplete hidden";
  input.insertAdjacentElement("afterend", list);
  let options = [], active = 0;

  const currentWord = () => {
    const caret = input.selectionStart ?? input.value.length;
    const before = input.value.slice(0, caret);
    const m = /(\S*)$/.exec(before);
    return { word: m ? m[1] : "", start: caret - (m ? m[1].length : 0), caret };
  };
  const close = () => { list.classList.add("hidden"); options = []; };
  const show = () => {
    const { word } = currentWord();
    options = autocompleteOptions(word);
    if (!options.length) { close(); return; }
    active = 0;
    list.innerHTML = "";
    options.forEach((opt, i) => {
      const item = document.createElement("div");
      item.className = "autocomplete-item" + (i === active ? " active" : "");
      item.textContent = opt;
      item.addEventListener("mousedown", (e) => { e.preventDefault(); take(i); });
      list.appendChild(item);
    });
    list.style.left = input.offsetLeft + "px";
    list.style.top = (input.offsetTop + input.offsetHeight) + "px";
    list.classList.remove("hidden");
  };
  const take = (i) => {
    const { start, caret } = currentWord();
    const text = options[i];
    input.value = input.value.slice(0, start) + text + input.value.slice(caret);
    const pos = start + text.length;
    input.setSelectionRange(pos, pos);
    close();
    input.dispatchEvent(new Event("input", { bubbles: true }));
  };
  input.addEventListener("input", (e) => { if (e.isTrusted) show(); });
  input.addEventListener("blur", () => setTimeout(close, 100));
  input.addEventListener("keydown", (e) => {
    if (list.classList.contains("hidden") || !options.length) return;
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      active = (active + (e.key === "ArrowDown" ? 1 : options.length - 1)) % options.length;
      [...list.children].forEach((el, i) => el.classList.toggle("active", i === active));
    } else if (e.key === "Enter" || e.key === "Tab") {
      e.preventDefault();
      take(active);
    } else if (e.key === "Escape") {
      close();
    }
  });
}

// ---------------------------------------------------------------------------
//  About and updates. The installed app asks GitHub once at start whether a
//  newer release is out (see update.py; it can be switched off here). Run
//  from sources, it never checks by itself: "Check now" still works.
// ---------------------------------------------------------------------------

let UPDATE = null;

async function updateFetch(force) {
  const res = await fetch(`/api/update/check${force ? "?force=1" : ""}`);
  return res.json();
}

function updateApply(d) {
  UPDATE = d;
  const badge = document.getElementById("update-badge");
  badge.hidden = !d.newer;
  if (d.newer) badge.textContent = `⬆ v${d.latest} available`;
  document.getElementById("about-check-enabled").checked = !!d.enabled;
}

async function updateAtStart() {
  try {
    const d = await updateFetch(false);
    // Remember the setting either way; only the installed app checks by itself.
    document.getElementById("about-check-enabled").checked = !!d.enabled;
    if (d.frozen) updateApply(d);
  } catch (e) { /* offline: say nothing */ }
}

/**
 * Release notes are a CHANGELOG.md section: markdown, wrapped at 72
 * columns. Shown as typed they broke mid-sentence and showed "###" and
 * "**". This joins the wrapped lines and draws headings, bullets, bold
 * and code. The text is escaped first, so nothing in it becomes markup
 * except these few marks.
 */
function notesHtml(md) {
  const esc = (t) => t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const inline = (t) => esc(t)
    .replace(/`([^`]+)`/g, "<code>$1</code>")
    .replace(/\*\*([^*]+)\*\*/g, "<b>$1</b>")
    .replace(/\[([^\]]+)\]\((https:\/\/[^)\s]+)\)/g, "$1");
  const out = [];
  let para = null, list = null;
  const flush = () => {
    if (para) { out.push("<p>" + inline(para.join(" ")) + "</p>"); para = null; }
    if (list) { out.push("<ul>" + list.map((li) => "<li>" + inline(li) + "</li>").join("") + "</ul>"); list = null; }
  };
  String(md).split(/\r?\n/).forEach((raw) => {
    const line = raw.trim();
    if (!line || /^-{3,}$/.test(line)) { flush(); return; }
    const h = line.match(/^#{1,6}\s+(.*)$/);
    if (h) { flush(); out.push("<h4>" + inline(h[1]) + "</h4>"); return; }
    const li = line.match(/^[-*]\s+(.*)$/);
    if (li) { if (para) flush(); list = list || []; list.push(li[1]); return; }
    if (list && /^\s{2,}/.test(raw)) { list[list.length - 1] += " " + line; return; }
    if (list) flush();
    para = para || [];
    para.push(line);
  });
  flush();
  return out.join("");
}

function openAbout() {
  const d = UPDATE || {};
  const current = (META && META.app_version) || d.current || "";
  document.getElementById("about-title").textContent = d.newer
    ? `Song Notation Tool v${d.latest} is out` : `Song Notation Tool v${current}`;
  let sub = `You have v${current}.`;
  if (d.newer) {
    sub += d.can_install
      ? " Update and restart downloads the new version, checks it against the"
        + " release's SHA-256 list and starts it again. Your songs and settings"
        + " are not touched."
      : (d.frozen ? " Download the new version from the release page."
                  : " You run it from the sources: update with  git pull,"
                    + " then restart it, or download the release.");
  }
  document.getElementById("about-sub").textContent = sub;
  const notes = document.getElementById("about-notes");
  notes.innerHTML = d.newer ? notesHtml(d.notes || "") : "";
  notes.classList.toggle("hidden", !(d.newer && d.notes));
  document.getElementById("about-install").hidden = !(d.newer && d.can_install && d.asset);
  document.getElementById("about-msg").textContent = d.error || "";
  document.getElementById("about-modal-backdrop").classList.remove("hidden");
}

function closeAbout() {
  document.getElementById("about-modal-backdrop").classList.add("hidden");
}

async function aboutCheckNow() {
  const msg = document.getElementById("about-msg");
  msg.textContent = "Checking…";
  let d;
  try { d = await updateFetch(true); }
  catch (e) { msg.textContent = "Could not check for updates."; return; }
  updateApply(d);
  openAbout();
  if (!d.error && !d.newer) {
    msg.textContent = `You have the latest version (v${d.current}).`;
  }
}

async function aboutInstall() {
  const msg = document.getElementById("about-msg");
  const btn = document.getElementById("about-install");
  btn.disabled = true;
  try {
    if (currentDoc) await saveCurrent();   // the app restarts: keep the song
  } catch (e) {
    msg.textContent = `Could not save the song first: ${e}. Nothing was updated.`;
    btn.disabled = false;
    return;
  }
  msg.textContent = "Downloading…";
  try {
    const res = await fetch("/api/update/install", { method: "POST" });
    const d = await res.json();
    msg.textContent = d.message || (res.ok ? "Done." : "Failed.");
    if (!res.ok) btn.disabled = false;
  } catch (e) {
    msg.textContent = "The app is restarting…";
  }
}

async function aboutSetEnabled(on) {
  try {
    const res = await fetch("/api/update/settings", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ check: !!on }),
    });
    const d = await res.json();
    if (!on) document.getElementById("update-badge").hidden = true;
    UPDATE = Object.assign(UPDATE || {}, { enabled: d.enabled });
  } catch (e) { /* stays as it was */ }
}

async function aboutReleasePage() {
  const url = (UPDATE && UPDATE.url) || "";
  try {
    await fetch("/api/update/open", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ url }),
    });
  } catch (e) { toast("Could not open the release page."); }
}

function wireAbout() {
  document.getElementById("app-version").addEventListener("click", openAbout);
  document.getElementById("update-badge").addEventListener("click", openAbout);
  document.getElementById("about-close").addEventListener("click", closeAbout);
  document.getElementById("about-check-now").addEventListener("click", aboutCheckNow);
  document.getElementById("about-install").addEventListener("click", aboutInstall);
  document.getElementById("about-release-page").addEventListener("click", aboutReleasePage);
  document.getElementById("about-check-enabled").addEventListener(
    "change", (e) => aboutSetEnabled(e.target.checked));
  setTimeout(updateAtStart, 1500);
}

async function init() {
  META = await API.meta();
  document.getElementById("app-version").textContent = `v${META.app_version}`;
  wireAbout();
  wireRiffs();
  wireImport();
  wireMetaForm();

  // Layout, Colour and Size are all properties of the *song*, not of this
  // machine — they travel with the .sng so a chart prints the same way
  // wherever it's opened.
  [["meta-layout", "section_layout", META.section_layouts, "banner"],
   ["meta-color", "color_mode", META.color_modes, "color"],
   ["meta-scale", "pdf_scale", META.pdf_scales, "fit"],
   ["meta-columns", "pdf_columns", META.pdf_columns, "auto"],
   ["meta-lyrics-layout", "lyrics_layout", META.lyrics_layouts, "none"],
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
      refreshLayoutPanel();
      schedulePreviewUpdate();
    });
  });
  wireLayoutPanel();
  wireDialogs();
  wireTheme();
  wireSidebarToggle();
  setSongToolsEnabled(false);

  renderSongsFolder();
  // pywebview injects its bridge *after* the page has loaded, and says so
  // with this event. Checking only at startup found no bridge, so the
  // native window showed the browser-tab hint and hid Change… and Reveal.
  window.addEventListener("pywebviewready", renderSongsFolder);
  document.getElementById("btn-songs-folder").addEventListener("click",
    () => changeSongsFolder().catch((e) => toast(String(e))));
  document.getElementById("folder-cancel").addEventListener("click", closeFolderModal);
  document.getElementById("folder-move").addEventListener("click",
    () => applySongsFolder(true).catch((e) => toast(String(e))));
  document.getElementById("folder-use").addEventListener("click",
    () => applySongsFolder(false).catch((e) => toast(String(e))));
  document.getElementById("btn-reveal-songs").addEventListener("click", () => {
    const api = nativeApi();
    if (api && api.reveal) api.reveal(META.songs_dir || "");
  });

  document.getElementById("btn-help").addEventListener("click", (e) => {
    const strip = document.getElementById("help-strip");
    strip.classList.toggle("hidden");
    e.currentTarget.setAttribute("aria-expanded", String(!strip.classList.contains("hidden")));
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

  const exportMenu = document.getElementById("export-menu");
  const exportBtn = document.getElementById("btn-export");
  const setExportOpen = (open) => {
    exportMenu.classList.toggle("hidden", !open);
    exportBtn.setAttribute("aria-expanded", String(open));
  };
  exportBtn.addEventListener("click", () => {
    const open = exportMenu.classList.contains("hidden");
    setExportOpen(open);
    if (open) exportMenu.querySelector("button").focus();
  });
  exportMenu.addEventListener("click", (e) => {
    const fmt = e.target.closest("[data-fmt]") && e.target.closest("[data-fmt]").dataset.fmt;
    if (!fmt) return;
    setExportOpen(false);
    exportCurrent(fmt).catch((err) => toast(String(err)));
  });
  exportMenu.addEventListener("keydown", (e) => {
    const items = [...exportMenu.querySelectorAll("button")];
    const i = items.indexOf(document.activeElement);
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); setExportOpen(false); exportBtn.focus(); }
    else if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      items[(i + (e.key === "ArrowDown" ? 1 : items.length - 1)) % items.length].focus();
    }
  });
  document.addEventListener("click", (e) => {
    if (!e.target.closest("#topbar-right .dropdown")) setExportOpen(false);
    // A section's ⋯ menu closes when you click anywhere else.
    if (!e.target.closest(".sec-more-wrap")) {
      document.querySelectorAll(".sec-menu").forEach((m) => m.classList.add("hidden"));
      document.querySelectorAll(".sec-more").forEach((b) => b.setAttribute("aria-expanded", "false"));
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
  document.getElementById("lyrics-open-layout").addEventListener("click", (e) => {
    e.stopPropagation();
    closeLyricsModal();
    openLayoutPanel();
    document.getElementById("meta-lyrics-layout").focus();
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

  wireSongLibrary();
  await refreshSongList();
  // The first song that shows in All songs: a hidden one is not opened.
  const first = SONGS.find((s) => !s.hidden);
  if (first) {
    await openSong(first.filename);
  } else {
    document.getElementById("start-here").classList.remove("hidden");
  }
}

document.addEventListener("DOMContentLoaded", () => {
  init().catch((err) => toast(`Failed to start: ${err}`));
});

/*
 * ChartLite datasheet: an Excel-style grid that opens as a popup next to PowerPoint.
 * Row 1 = category names, column A = series names, values in between.
 * Every change is sent to the ChartLite panel, which redraws the chart live.
 *
 * Messages to the panel:   {type: "ready"} | {type: "data", categories, series:[{name, raw}], pasted}
 * Messages from the panel: {type: "load", chartType, title, categories, series:[{name, raw}]}
 * Transport: Office dialog API inside PowerPoint; window.postMessage in a normal browser.
 */
(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const grid = $("grid"), wrap = $("wrap"), editor = $("editor"), fx = $("fx");

  // ------------------------------------------------------------------ model
  let cells = [[""]];               // jagged: cells[r][c] (strings)
  let R = 20, C = 10;               // rendered size (always a bit larger than the data)
  let anchor = { r: 1, c: 1 };      // active cell
  let focus = { r: 1, c: 1 };       // other corner of the selection
  let editing = null;               // null | "enter" (typed over) | "edit" (F2 / double-click)
  let undoStack = [], redoStack = [];
  let dragging = false;

  const get = (r, c) => (cells[r] && cells[r][c] !== undefined ? cells[r][c] : "");
  function set(r, c, v) {
    while (cells.length <= r) cells.push([]);
    const row = cells[r];
    while (row.length <= c) row.push("");
    row[c] = v;
  }
  const colName = (c) => { let s = ""; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
  const rect = () => ({
    r0: Math.min(anchor.r, focus.r), r1: Math.max(anchor.r, focus.r),
    c0: Math.min(anchor.c, focus.c), c1: Math.max(anchor.c, focus.c),
  });
  const isNumeric = (v) => /^\(?\s*[-−+]?\s*(Rp\.?|IDR|\$|€)?\s*\d[\d.,]*\s*%?\s*\)?$/i.test(String(v).trim());
  const isNeg = (v) => /^\s*(\(|[-−])/.test(String(v));

  function usedSize() {
    let rows = 0, cols = 0;
    cells.forEach((row, r) => row.forEach((v, c) => { if (String(v).trim()) { rows = Math.max(rows, r + 1); cols = Math.max(cols, c + 1); } }));
    return { rows, cols };
  }
  function ensureSize() {
    const u = usedSize();
    const nr = Math.max(20, u.rows + 6, anchor.r + 3, focus.r + 3);
    const nc = Math.max(10, u.cols + 3, anchor.c + 2, focus.c + 2);
    if (nr !== R || nc !== C) { R = nr; C = nc; renderAll(); return true; }
    return false;
  }

  // ------------------------------------------------------------------ rendering
  function renderAll() {
    let h = '<thead><tr><th class="rowh"></th>';
    for (let c = 0; c < C; c++) h += `<th data-hc="${c}" style="${c === 0 ? "width:130px;min-width:130px;max-width:130px" : ""}">${colName(c)}</th>`;
    h += "</tr></thead><tbody>";
    for (let r = 0; r < R; r++) {
      h += `<tr><td class="rowh" data-hr="${r}">${r + 1}</td>`;
      for (let c = 0; c < C; c++) {
        if (r === 0 && c === 0) {
          h += `<td class="c corner r0 c0" data-r="0" data-c="0" style="width:130px;min-width:130px;max-width:130px">` +
            `<svg preserveAspectRatio="none" viewBox="0 0 100 100"><line x1="0" y1="0" x2="100" y2="100" stroke="#d9a07f" stroke-width="1.5" vector-effect="non-scaling-stroke"/></svg>` +
            `<span class="cat">Category</span><span class="ser">Series</span></td>`;
        } else {
          h += `<td class="c${r === 0 ? " r0" : ""}${c === 0 ? " c0" : ""}" data-r="${r}" data-c="${c}"${c === 0 ? ' style="width:130px;min-width:130px;max-width:130px"' : ""}></td>`;
        }
      }
      h += "</tr>";
    }
    grid.innerHTML = h + "</tbody>";
    for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) paintCell(r, c);
    paintSelection();
  }
  const td = (r, c) => grid.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
  function paintCell(r, c) {
    if (r === 0 && c === 0) return;
    const el = td(r, c); if (!el) return;
    const v = get(r, c);
    el.textContent = v;
    const numeric = r > 0 && c > 0 && isNumeric(v);
    el.classList.toggle("num", numeric);
    el.classList.toggle("neg", numeric && isNeg(v));
  }
  let lastPainted = [];
  function paintSelection() {
    lastPainted.forEach((el) => el.classList.remove("insel", "active", "hl"));
    lastPainted = [];
    const q = rect();
    for (let r = q.r0; r <= q.r1; r++) for (let c = q.c0; c <= q.c1; c++) {
      const el = td(r, c); if (!el) continue;
      if (q.r0 !== q.r1 || q.c0 !== q.c1) el.classList.add("insel");
      lastPainted.push(el);
    }
    const a = td(anchor.r, anchor.c);
    if (a) { a.classList.add("active"); lastPainted.push(a); }
    for (let c = q.c0; c <= q.c1; c++) { const h = grid.querySelector(`th[data-hc="${c}"]`); if (h) { h.classList.add("hl"); lastPainted.push(h); } }
    for (let r = q.r0; r <= q.r1; r++) { const h = grid.querySelector(`td[data-hr="${r}"]`); if (h) { h.classList.add("hl"); lastPainted.push(h); } }
    $("ref").textContent = colName(anchor.c) + (anchor.r + 1);
    if (document.activeElement !== fx) fx.value = anchor.r === 0 && anchor.c === 0 ? "" : get(anchor.r, anchor.c);
  }
  function scrollToActive() {
    const el = td(anchor.r, anchor.c);
    if (!el) return;
    const wr = wrap.getBoundingClientRect(), er = el.getBoundingClientRect();
    const top = wr.top + 26, left = wr.left + 41;
    if (er.top < top) wrap.scrollTop -= top - er.top;
    else if (er.bottom > wr.bottom - 14) wrap.scrollTop += er.bottom - wr.bottom + 14;
    if (er.left < left) wrap.scrollLeft -= left - er.left;
    else if (er.right > wr.right - 14) wrap.scrollLeft += er.right - wr.right + 14;
  }

  // ------------------------------------------------------------------ undo
  function snapshot() { undoStack.push(JSON.stringify(cells)); if (undoStack.length > 200) undoStack.shift(); redoStack = []; }
  function undo() {
    if (!undoStack.length) return;
    redoStack.push(JSON.stringify(cells));
    cells = JSON.parse(undoStack.pop());
    if (!ensureSize()) renderAll();
    emit();
  }
  function redo() {
    if (!redoStack.length) return;
    undoStack.push(JSON.stringify(cells));
    cells = JSON.parse(redoStack.pop());
    if (!ensureSize()) renderAll();
    emit();
  }

  // ------------------------------------------------------------------ messaging
  let officeReady = false;
  const canOffice = () => officeReady && Office.context && Office.context.ui && typeof Office.context.ui.messageParent === "function";
  window.__sent = [];
  function send(msg) {
    window.__sent.push(msg);
    if (canOffice()) { try { Office.context.ui.messageParent(JSON.stringify(msg)); return; } catch (e) { /* fall through */ } }
    if (window.opener && !window.opener.closed) window.opener.postMessage({ source: "chartlite-sheet", msg }, "*");
  }
  let emitTimer = null, pastedFlag = false;
  function emit(pasted) {
    if (pasted) pastedFlag = true;
    $("sync").textContent = "Updating…"; $("sync").className = "sync";
    clearTimeout(emitTimer);
    emitTimer = setTimeout(() => {
      send(Object.assign({ type: "data", pasted: pastedFlag }, extract()));
      pastedFlag = false;
      $("sync").textContent = "✓ Chart updated"; $("sync").className = "sync ok";
    }, 220);
  }
  function extract() {
    let lastCol = 0, lastRow = 0;
    cells.forEach((row, r) => row.forEach((v, c) => {
      if (!String(v).trim() || (r === 0 && c === 0)) return;
      if (c >= 1) lastCol = Math.max(lastCol, c);
      if (r >= 1) lastRow = Math.max(lastRow, r);
    }));
    const categories = [];
    for (let c = 1; c <= lastCol; c++) categories.push(String(get(0, c)).trim());
    const series = [];
    for (let r = 1; r <= lastRow; r++) {
      const vals = [];
      for (let c = 1; c <= lastCol; c++) vals.push(String(get(r, c)).trim());
      const name = String(get(r, 0)).trim();
      if (!name && !vals.some(Boolean)) continue;
      series.push({ name, raw: vals });
    }
    return { categories, series };
  }
  const TYPE_NAMES = { stacked: "Stacked", stacked100: "100%", clustered: "Clustered", waterfall: "Waterfall", line: "Line", combo: "Combination" };
  function load(p) {
    if (!p || p.type !== "load") return;
    cancelEdit();
    cells = [[""].concat((p.categories || []).map(String))];
    (p.series || []).forEach((s) => cells.push([String(s.name || "")].concat((s.raw || []).map((v) => (v === null || v === undefined ? "" : String(v))))));
    undoStack = []; redoStack = [];
    $("title").textContent = "ChartLite · " + (p.title || ((TYPE_NAMES[p.chartType] || "Chart") + " datasheet"));
    anchor = { r: 1, c: 1 }; focus = { r: 1, c: 1 };
    if (!ensureSize()) renderAll();
    $("sync").textContent = "Changes update the chart live"; $("sync").className = "sync";
  }

  // ------------------------------------------------------------------ editing
  function placeEditor() {
    const el = td(anchor.r, anchor.c); if (!el) return;
    const wr = wrap.getBoundingClientRect(), er = el.getBoundingClientRect();
    editor.style.left = er.left - wr.left + wrap.scrollLeft - 1 + "px";
    editor.style.top = er.top - wr.top + wrap.scrollTop - 1 + "px";
    editor.style.width = Math.max(er.width + 2, 120) + "px";
    editor.style.height = er.height + 2 + "px";
    editor.style.textAlign = el.classList.contains("num") ? "right" : "left";
  }
  function startEdit(initial, mode) {
    if (anchor.r === 0 && anchor.c === 0) return;
    scrollToActive();
    editing = mode;
    editor.value = initial;
    placeEditor();
    editor.style.display = "block";
    editor.focus();
    const n = editor.value.length; editor.setSelectionRange(n, n);
  }
  function commitEdit() {
    if (!editing) return;
    const v = editor.value;
    editing = null;
    editor.style.display = "none";
    if (v !== get(anchor.r, anchor.c)) {
      snapshot(); set(anchor.r, anchor.c, v); paintCell(anchor.r, anchor.c);
      ensureSize(); emit();
    }
    paintSelection();
    wrap.focus({ preventScroll: true });
  }
  function cancelEdit() {
    if (!editing) return;
    editing = null; editor.style.display = "none";
    wrap.focus({ preventScroll: true });
  }
  function move(dr, dc, extend) {
    const t = extend ? focus : anchor;
    const nr = Math.max(0, t.r + dr), nc = Math.max(0, t.c + dc);
    if (extend) focus = { r: nr, c: nc };
    else { anchor = { r: nr, c: nc }; focus = { r: nr, c: nc }; }
    ensureSize(); paintSelection(); scrollToActive();
  }
  function clearRange() {
    const q = rect();
    snapshot();
    for (let r = q.r0; r <= q.r1; r++) for (let c = q.c0; c <= q.c1; c++) { if (r === 0 && c === 0) continue; set(r, c, ""); paintCell(r, c); }
    emit();
  }

  // ------------------------------------------------------------------ clipboard
  function splitTSV(text) {
    text = String(text || "").replace(/\r\n?/g, "\n");
    const rows = []; let row = [], cell = "", q = false;
    for (let k = 0; k < text.length; k++) {
      const ch = text[k];
      if (q) {
        if (ch === '"' && text[k + 1] === '"') { cell += '"'; k++; }
        else if (ch === '"') q = false;
        else cell += ch === "\n" ? " " : ch;
      } else if (ch === '"' && cell === "") q = true;
      else if (ch === "\t") { row.push(cell); cell = ""; }
      else if (ch === "\n") { row.push(cell); rows.push(row); row = []; cell = ""; }
      else cell += ch;
    }
    if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
    while (rows.length && rows[rows.length - 1].every((c) => !c.trim())) rows.pop();
    return rows.map((r) => r.map((c) => c.replace(/ /g, " ").trim()));
  }
  function pasteBlock(text) {
    const rows = splitTSV(text);
    if (!rows.length) return;
    const q = rect();
    snapshot();
    rows.forEach((row, i) => row.forEach((v, j) => {
      const r = q.r0 + i, c = q.c0 + j;
      if (r === 0 && c === 0) return;
      set(r, c, v);
    }));
    anchor = { r: q.r0, c: q.c0 };
    focus = { r: q.r0 + rows.length - 1, c: q.c0 + Math.max(...rows.map((r) => r.length)) - 1 };
    if (!ensureSize()) { for (let r = q.r0; r <= focus.r; r++) for (let c = q.c0; c <= focus.c; c++) paintCell(r, c); paintSelection(); }
    emit(true);
  }
  function copyText() {
    const q = rect(), lines = [];
    for (let r = q.r0; r <= q.r1; r++) {
      const vals = [];
      for (let c = q.c0; c <= q.c1; c++) vals.push(r === 0 && c === 0 ? "" : get(r, c));
      lines.push(vals.join("\t"));
    }
    return lines.join("\r\n");
  }
  document.addEventListener("copy", (e) => { if (editing || document.activeElement === fx) return; e.preventDefault(); e.clipboardData.setData("text/plain", copyText()); });
  document.addEventListener("cut", (e) => { if (editing || document.activeElement === fx) return; e.preventDefault(); e.clipboardData.setData("text/plain", copyText()); clearRange(); });
  document.addEventListener("paste", (e) => {
    if (document.activeElement === fx) return;
    const text = (e.clipboardData || window.clipboardData).getData("text");
    if (editing && !/[\t\n]/.test(text.trim())) return;   // single value typed into a cell
    e.preventDefault();
    cancelEdit();
    pasteBlock(text);
  });

  // ------------------------------------------------------------------ mouse
  const cellOf = (el) => { const t = el && el.closest && el.closest("td.c"); return t ? { r: +t.dataset.r, c: +t.dataset.c } : null; };
  grid.addEventListener("mousedown", (e) => {
    const p = cellOf(e.target);
    if (!p) {
      const h = e.target.closest && e.target.closest("th[data-hc], td[data-hr]");
      if (!h) return;
      commitEdit();
      if (h.dataset.hc !== undefined) { anchor = { r: 0, c: +h.dataset.hc }; focus = { r: R - 1, c: +h.dataset.hc }; }
      else { anchor = { r: +h.dataset.hr, c: 0 }; focus = { r: +h.dataset.hr, c: C - 1 }; }
      paintSelection(); wrap.focus({ preventScroll: true });
      e.preventDefault();
      return;
    }
    commitEdit();
    if (e.shiftKey) focus = p; else { anchor = p; focus = p; }
    dragging = true;
    paintSelection();
    wrap.focus({ preventScroll: true });
    e.preventDefault();
  });
  grid.addEventListener("mouseover", (e) => {
    if (!dragging) return;
    const p = cellOf(e.target);
    if (p && (p.r !== focus.r || p.c !== focus.c)) { focus = p; paintSelection(); }
  });
  document.addEventListener("mouseup", () => { dragging = false; });
  grid.addEventListener("dblclick", (e) => { const p = cellOf(e.target); if (p) { anchor = p; focus = p; paintSelection(); startEdit(get(p.r, p.c), "edit"); } });

  // ------------------------------------------------------------------ keyboard
  wrap.addEventListener("keydown", (e) => {
    if (editing) return;
    const k = e.key, ctrl = e.ctrlKey || e.metaKey;
    if (ctrl && (k === "z" || k === "Z")) { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
    if (ctrl && (k === "y" || k === "Y")) { e.preventDefault(); redo(); return; }
    if (ctrl && (k === "a" || k === "A")) {
      e.preventDefault();
      const u = usedSize(); anchor = { r: 0, c: 0 }; focus = { r: Math.max(0, u.rows - 1), c: Math.max(0, u.cols - 1) }; paintSelection(); return;
    }
    if (ctrl) return;   // let the browser raise copy / cut / paste events
    const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (moves[k]) { e.preventDefault(); move(moves[k][0], moves[k][1], e.shiftKey); return; }
    if (k === "Enter") { e.preventDefault(); move(e.shiftKey ? -1 : 1, 0, false); return; }
    if (k === "Tab") { e.preventDefault(); move(0, e.shiftKey ? -1 : 1, false); return; }
    if (k === "Delete" || k === "Backspace") { e.preventDefault(); clearRange(); return; }
    if (k === "F2") { e.preventDefault(); startEdit(get(anchor.r, anchor.c), "edit"); return; }
    if (k.length === 1 && !e.altKey) { e.preventDefault(); startEdit(k, "enter"); }
  });
  editor.addEventListener("keydown", (e) => {
    e.stopPropagation();   // the editor sits inside the grid: don't let the grid handle the same key again
    const k = e.key;
    if (k === "Escape") { e.preventDefault(); cancelEdit(); return; }
    if (k === "Enter") { e.preventDefault(); commitEdit(); move(e.shiftKey ? -1 : 1, 0, false); return; }
    if (k === "Tab") { e.preventDefault(); commitEdit(); move(0, e.shiftKey ? -1 : 1, false); return; }
    const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (editing === "enter" && moves[k]) { e.preventDefault(); commitEdit(); move(moves[k][0], moves[k][1], false); }
  });
  editor.addEventListener("input", () => { fx.value = editor.value; });
  editor.addEventListener("blur", () => { if (editing) commitEdit(); });
  wrap.addEventListener("scroll", () => { if (editing) placeEditor(); });

  // formula bar edits the active cell
  fx.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === "Tab") {
      e.preventDefault();
      if (fx.value !== get(anchor.r, anchor.c) && !(anchor.r === 0 && anchor.c === 0)) {
        snapshot(); set(anchor.r, anchor.c, fx.value); paintCell(anchor.r, anchor.c); ensureSize(); emit();
      }
      wrap.focus({ preventScroll: true });
      e.key === "Enter" ? move(1, 0, false) : move(0, 1, false);
    } else if (e.key === "Escape") { fx.value = get(anchor.r, anchor.c); wrap.focus({ preventScroll: true }); }
  });

  // ------------------------------------------------------------------ toolbar
  $("btnUndo").addEventListener("click", () => { undo(); wrap.focus(); });
  $("btnClear").addEventListener("click", () => {
    snapshot(); cells = [[""]]; anchor = { r: 1, c: 1 }; focus = { r: 1, c: 1 };
    if (!ensureSize()) renderAll(); emit(); wrap.focus();
  });
  $("btnTranspose").addEventListener("click", () => {
    const u = usedSize();
    snapshot();
    const t = [];
    for (let c = 0; c < u.cols; c++) { t.push([]); for (let r = 0; r < u.rows; r++) t[c].push(get(r, c)); }
    cells = t.length ? t : [[""]];
    if (cells[0]) cells[0][0] = "";
    if (!ensureSize()) renderAll();
    emit(); wrap.focus();
  });

  // ------------------------------------------------------------------ boot
  window.addEventListener("message", (e) => {
    const d = e.data;
    if (d && d.source === "chartlite-pane") load(d.msg);
  });
  try {
    if (location.hash.startsWith("#d=")) load(JSON.parse(decodeURIComponent(location.hash.slice(3))));
  } catch (e) { /* the panel will send the data instead */ }
  if (!grid.innerHTML) renderAll();
  wrap.focus();

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(() => {
      officeReady = true;
      try {
        Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
          try { load(JSON.parse(arg.message)); } catch (e) { /* ignore malformed */ }
        });
      } catch (e) { /* older Office: data arrives through the URL only */ }
      send({ type: "ready" });
    });
  } else {
    send({ type: "ready" });
  }

  // test hooks
  window.__sheet = { load, extract, get cells() { return cells; } };
})();

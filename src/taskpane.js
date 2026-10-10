/* ChartLite task pane: datasheet, options, preview and PowerPoint wiring. */
(function () {
  "use strict";
  const E = window.ChartEngine;
  const $ = (id) => document.getElementById(id);

  // ------------------------------------------------------------------ state
  const state = {
    type: "stacked",
    categories: ["Jan", "Feb", "Mar", "Apr"],
    series: [
      { name: "Product A", color: "", kind: "bar", axis: "primary", raw: ["100", "110", "105", "120"] },
      { name: "Product B", color: "", kind: "bar", axis: "primary", raw: ["80", "85", "90", "95"] },
      { name: "Product C", color: "", kind: "bar", axis: "primary", raw: ["50", "55", "52", "60"] },
    ],
    deltas: [],
  };
  let locale = "en";
  let target = null;            // {id, slideId} of the chart being edited on the slide
  let inOffice = false;
  let busy = false, pending = false;
  let ignoreSelUntil = 0;
  let liveTimer = null, previewTimer = null;

  // ------------------------------------------------------------------ helpers
  const toRaw = (v, loc) => {
    if (v === "e") return "e";
    if (v === null || v === undefined || v === "") return "";
    const s = String(v);
    return loc === "id" ? s.replace(".", ",") : s;
  };

  function setStatus(msg, kind) {
    const el = $("status");
    el.textContent = msg || "";
    el.className = "status" + (kind ? " " + kind : "");
  }

  function numOpt(id) {
    const v = E.parseCell($(id).value, locale);
    return typeof v === "number" ? v : null;
  }

  function readOptions() {
    const v = (id) => $(id).value;
    const c = (id) => $(id).checked;
    const cf = v("cagrFrom"), ct = v("cagrTo");
    return {
      decimals: Math.max(0, Math.min(4, parseInt(v("decimals"), 10) || 0)),
      fontSize: Math.max(6, Math.min(24, parseFloat(v("fontSize")) || 10)),
      prefix: v("prefix"),
      unit: v("unit"),
      locale,
      negStyle: v("negStyle"),
      fontName: v("fontName").trim(),
      showLabels: c("showLabels"),
      showTotals: c("showTotals"),
      showLegend: c("showLegend"),
      pct100Labels: c("pct100Labels"),
      barWidth: parseFloat(v("barWidth")),
      cagr: cf !== "" && ct !== "" ? { from: +cf, to: +ct } : null,
      deltas: state.deltas.map((d) => ({ from: d.from, to: d.to, mode: d.mode })),
      wf: { start: v("wfTotal"), total: v("wfTotal"), up: v("wfUp"), down: v("wfDown") },
      connectors: c("connectors"),
      axisBreak: c("axisBreak"),
      yMin: numOpt("yMin"), yMax: numOpt("yMax"),
      y2Min: numOpt("y2Min"), y2Max: numOpt("y2Max"),
      smooth: c("smooth"),
      markerSize: numOpt("markerSize"),
      markerFill: v("markerFill"),
    };
  }

  function writeOptions(o) {
    const set = (id, val) => { if (val !== undefined && val !== null) $(id).value = val; };
    const chk = (id, val) => { if (val !== undefined) $(id).checked = !!val; };
    set("decimals", o.decimals); set("fontSize", o.fontSize); set("prefix", o.prefix);
    set("unit", o.unit); set("negStyle", o.negStyle); set("fontName", o.fontName);
    set("barWidth", o.barWidth);
    locale = o.locale || "en"; $("locale").value = locale;
    chk("showLabels", o.showLabels); chk("showTotals", o.showTotals); chk("showLegend", o.showLegend);
    chk("pct100Labels", o.pct100Labels); chk("connectors", o.connectors); chk("axisBreak", o.axisBreak);
    if (o.wf) { set("wfTotal", o.wf.total); set("wfUp", o.wf.up); set("wfDown", o.wf.down); }
    ["yMin", "yMax", "y2Min", "y2Max", "markerSize"].forEach((id) => { $(id).value = o[id] === null || o[id] === undefined ? "" : toRaw(o[id], locale); });
    chk("smooth", o.smooth);
    set("markerFill", o.markerFill || "solid");
    state.deltas = (o.deltas || []).map((d) => ({ from: d.from, to: d.to, mode: d.mode || "pct" }));
    renderArrowControls(o.cagr || null);
  }

  function buildSpec() {
    const type = state.type;
    return {
      version: 1,
      type,
      categories: state.categories.slice(),
      series: state.series.map((s) => {
        const out = {
          name: s.name,
          values: s.raw.map((r) => {
            const v = E.parseCell(r, locale);
            return v === "e" ? (type === "waterfall" ? "e" : null) : v;
          }),
        };
        if (s.color) out.color = s.color;
        if (type === "combo") { out.kind = s.kind || "bar"; out.axis = s.axis || "primary"; }
        if (s.unit) out.unit = s.unit;
        return out;
      }),
      options: readOptions(),
    };
  }

  function loadSpec(spec) {
    state.type = spec.type || "stacked";
    $("chartType").value = state.type;
    const loc = (spec.options && spec.options.locale) || "en";
    state.categories = (spec.categories || []).map(String);
    state.series = (spec.series || []).map((s) => ({
      name: s.name || "",
      color: s.color || "",
      kind: s.kind || "bar",
      axis: s.axis || "primary",
      unit: s.unit,
      raw: (s.values || []).map((v) => toRaw(v, loc)),
    }));
    writeOptions(spec.options || {});
    renderSheet();
    refresh({ live: false });
  }

  // ------------------------------------------------------------------ datasheet
  function renderSheet() {
    const t = $("sheet");
    const n = state.categories.length;
    const combo = state.type === "combo";
    const wf = state.type === "waterfall";
    let html = `<tr><th class="corner" colspan="2">Series ↓ · Category →</th>`;
    state.categories.forEach((c, j) => { html += `<th><input data-cat="${j}" value="${attr(c)}" /></th>`; });
    html += "</tr>";
    state.series.forEach((s, i) => {
      const dim = wf && i > 0 ? ' style="opacity:.45" title="Waterfall uses the first series only"' : "";
      html += `<tr${dim}><td class="sname"><input data-sname="${i}" value="${attr(s.name)}" /></td><td class="smeta">`;
      html += `<input type="color" class="swatch" data-color="${i}" value="${s.color || E.DEFAULT_PALETTE[i % E.DEFAULT_PALETTE.length]}" title="Series color" />`;
      if (combo) {
        html += `<select data-kind="${i}"><option value="bar"${s.kind !== "line" ? " selected" : ""}>Bar</option><option value="line"${s.kind === "line" ? " selected" : ""}>Line</option></select>`;
        html += `<select data-axis="${i}" title="Secondary axis draws the line above the bars on its own scale"><option value="primary"${s.axis !== "secondary" ? " selected" : ""}>Axis 1</option><option value="secondary"${s.axis === "secondary" ? " selected" : ""}>Axis 2</option></select>`;
      }
      html += `</td>`;
      for (let j = 0; j < n; j++) {
        const r = s.raw[j] === undefined ? "" : s.raw[j];
        html += `<td><input data-r="${i}" data-c="${j}" value="${attr(r)}" class="${isNeg(r) ? "neg" : ""}" /></td>`;
      }
      html += "</tr>";
    });
    t.innerHTML = html;
    $("wfHint").hidden = !wf;
    document.querySelectorAll(".y2").forEach((el) => { el.style.display = combo ? "" : "none"; });
    $("wfPanel").style.display = wf ? "" : "none";
    renderArrowControls();
  }

  function attr(s) { return String(s).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;"); }
  function isNeg(r) { const v = E.parseCell(r, locale); return typeof v === "number" && v < 0; }

  $("sheet").addEventListener("input", (ev) => {
    const el = ev.target, d = el.dataset;
    if (d.cat !== undefined) { state.categories[+d.cat] = el.value; renderArrowControls(); }
    else if (d.sname !== undefined) state.series[+d.sname].name = el.value;
    else if (d.color !== undefined) state.series[+d.color].color = el.value;
    else if (d.kind !== undefined) state.series[+d.kind].kind = el.value;
    else if (d.axis !== undefined) state.series[+d.axis].axis = el.value;
    else if (d.r !== undefined) {
      state.series[+d.r].raw[+d.c] = el.value;
      el.classList.toggle("neg", isNeg(el.value));
    }
    refresh();
  });
  $("sheet").addEventListener("change", (ev) => { if (ev.target.tagName === "SELECT") refresh(); });

  // Excel-style: pasting a block into any value cell fills the grid from there
  $("sheet").addEventListener("paste", (ev) => {
    const el = ev.target, d = el.dataset;
    const text = (ev.clipboardData || window.clipboardData).getData("text");
    if (!/[\t\n]/.test(text.trim())) return;
    ev.preventDefault();
    const rows = splitTSV(text);
    const det = detectLocale(rows);
    if (det && det !== locale) {
      // convert what is already in the grid, then switch
      state.series.forEach((s) => { s.raw = s.raw.map((r) => toRaw(E.parseCell(r, locale), det)); });
      locale = det; $("locale").value = det;
    }
    if (d.r !== undefined) {
      const r0 = +d.r, c0 = +d.c;
      rows.forEach((row, i) => {
        const si = r0 + i;
        while (state.series.length <= si) addSeries();
        row.forEach((cell, j) => {
          const cj = c0 + j;
          while (state.categories.length <= cj) addCategory();
          state.series[si].raw[cj] = cell.trim();
        });
      });
      renderSheet(); refresh();
      setStatus(`Pasted ${rows.length}×${rows[0].length} cells.`, "ok");
    } else {
      pasteBlock(text);
    }
  });

  // Enter moves down like Excel
  $("sheet").addEventListener("keydown", (ev) => {
    if (ev.key !== "Enter") return;
    const d = ev.target.dataset;
    if (d.r === undefined) return;
    ev.preventDefault();
    const next = $("sheet").querySelector(`input[data-r="${+d.r + (ev.shiftKey ? -1 : 1)}"][data-c="${d.c}"]`);
    if (next) { next.focus(); next.select(); }
  });

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
    // drop fully empty trailing columns (Excel selections often include one)
    let w = Math.max(0, ...rows.map((r) => r.length));
    while (w > 1 && rows.every((r) => !(r[w - 1] || "").trim())) w--;
    return rows.map((r) => r.slice(0, w).map((c) => c.replace(/\u00a0/g, " ")));
  }

  // Match Excel's displayed formatting: decimal places and a % unit
  function adoptExcelFormat(rows) {
    let maxDec = 0, nums = 0, pct = 0;
    const ds = locale === "id" ? "," : ".";
    rows.forEach((r) => r.forEach((c) => {
      const v = E.parseCell(c, locale);
      if (typeof v !== "number") return;
      nums++;
      if (/%\s*\)?$/.test(c.trim())) pct++;
      const t = c.replace(/[^\d.,]/g, "");
      const k = t.lastIndexOf(ds);
      if (k >= 0) maxDec = Math.max(maxDec, t.length - k - 1);
    }));
    if (!nums) return;
    $("decimals").value = Math.min(2, maxDec);
    if (pct === nums) $("unit").value = "%";
    else if ($("unit").value === "%") $("unit").value = "";
  }

  // Guess whether pasted numbers use 1,234.5 (en) or 1.234,5 (id). Returns null if unsure.
  function detectLocale(rows) {
    let en = 0, id = 0;
    rows.forEach((r) => r.forEach((c) => {
      const t = c.replace(/Rp|IDR|[%\s()+\-−]/gi, "");
      if (!/^[\d.,]+$/.test(t) || !/\d/.test(t)) return;
      const hasDot = t.includes("."), hasCom = t.includes(",");
      if (hasDot && hasCom) { t.lastIndexOf(",") > t.lastIndexOf(".") ? id++ : en++; }
      else if (hasCom) { if ((t.match(/,/g) || []).length > 1) en++; else if (/,\d{1,2}$|,\d{4,}$/.test(t)) id++; }
      else if (hasDot) { if ((t.match(/\./g) || []).length > 1) id++; else if (/\.\d{1,2}$|\.\d{4,}$/.test(t)) en++; }
    }));
    return id > en ? "id" : en > id ? "en" : null;
  }

  function applyBlock(rows) {
    if (!rows.length) return;
    const isNum = (s) => { const v = E.parseCell(s, locale); return typeof v === "number" || v === "e"; };
    const width = Math.max(...rows.map((r) => r.length));
    rows = rows.map((r) => { while (r.length < width) r.push(""); return r; });
    // Excel layout: blank top-left corner => header row + name column, even if headers are years
    const corner = rows.length > 1 && width > 1 && !rows[0][0].trim();
    const hasHeader = corner || (rows.length > 1 && rows[0].slice(1).some((c) => c.trim() && !isNum(c)));
    const hasNames = corner || (width > 1 && rows.slice(hasHeader ? 1 : 0).some((r) => r[0].trim() && !isNum(r[0])));
    const body = rows.slice(hasHeader ? 1 : 0);
    const off = hasNames ? 1 : 0;
    state.categories = hasHeader
      ? rows[0].slice(off).map((c) => c.trim())
      : Array.from({ length: width - off }, (_, j) => "Cat " + (j + 1));
    const old = state.series;
    state.series = body.map((r, i) => ({
      name: hasNames ? r[0].trim() : "Series " + (i + 1),
      color: old[i] ? old[i].color : "",
      kind: old[i] ? old[i].kind : "bar",
      axis: old[i] ? old[i].axis : "primary",
      raw: r.slice(off).map((c) => c.trim()),
    }));
    state.deltas = state.deltas.filter((d) => d.from < state.categories.length && d.to < state.categories.length);
  }

  function addCategory() {
    state.categories.push("Cat " + (state.categories.length + 1));
    state.series.forEach((s) => s.raw.push(""));
  }
  function addSeries() {
    state.series.push({ name: "Series " + (state.series.length + 1), color: "", kind: "bar", axis: "primary", raw: state.categories.map(() => "") });
  }

  document.querySelector(".toolbar").addEventListener("click", (ev) => {
    const act = ev.target.dataset.act;
    if (!act) return;
    if (act === "addCat") addCategory();
    if (act === "delCat" && state.categories.length > 1) {
      state.categories.pop(); state.series.forEach((s) => s.raw.pop());
      state.deltas = state.deltas.filter((d) => d.from < state.categories.length && d.to < state.categories.length);
    }
    if (act === "addSer") addSeries();
    if (act === "delSer" && state.series.length > 1) state.series.pop();
    if (act === "transpose") {
      const cats = state.series.map((s) => s.name);
      const sers = state.categories.map((c, j) => ({ name: c, color: "", kind: "bar", axis: "primary", raw: state.series.map((s) => s.raw[j] || "") }));
      state.categories = cats; state.series = sers; state.deltas = [];
    }
    renderSheet(); refresh();
  });

  // ---- Paste zone + Ctrl+V anywhere in the pane
  function pasteBlock(text, sourceLabel) {
    const rows = splitTSV(text);
    if (!rows.length || (rows.length === 1 && rows[0].length === 1 && !rows[0][0].trim())) return false;
    const det = detectLocale(rows);
    if (det && det !== locale) { locale = det; $("locale").value = det; }
    applyBlock(rows);
    adoptExcelFormat(rows);
    renderSheet(); refresh();
    const nNum = state.series.reduce((a, s) => a + s.raw.filter((r) => E.parseCell(r, locale) !== null).length, 0);
    setStatus(`Pasted ${state.series.length} series × ${state.categories.length} categories (${nNum} values)` +
      (det ? ", number style " + (det === "id" ? "1.234,5" : "1,234.5") : "") + ". " +
      (state.categories.length === 1 && state.series.length > 1 ? "Looks sideways — try ⇄ Transpose." : ""), "ok");
    const z = $("pasteZone"); z.classList.add("flash"); setTimeout(() => z.classList.remove("flash"), 600);
    return true;
  }

  $("pasteZone").addEventListener("paste", (ev) => {
    ev.preventDefault();
    pasteBlock((ev.clipboardData || window.clipboardData).getData("text"));
    $("pasteZone").value = "";
  });
  // typed/dropped text in the zone (fallback if the clipboard event is unavailable)
  $("pasteZone").addEventListener("input", () => {
    const v = $("pasteZone").value;
    if (/[\t\n]/.test(v)) { pasteBlock(v); $("pasteZone").value = ""; }
  });
  // Ctrl+V while nothing editable is focused
  document.addEventListener("paste", (ev) => {
    const t = ev.target;
    if (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "SELECT")) return;
    ev.preventDefault();
    pasteBlock((ev.clipboardData || window.clipboardData).getData("text"));
  });

  // ------------------------------------------------------------------ arrows
  function renderArrowControls(cagr) {
    const n = state.categories.length;
    const opts = (sel) => `<option value="">—</option>` + state.categories.map((c, j) =>
      `<option value="${j}"${sel === j ? " selected" : ""}>${attr(c || "Cat " + (j + 1))}</option>`).join("");
    const cur = cagr !== undefined ? cagr : (() => {
      const f = $("cagrFrom").value, t = $("cagrTo").value;
      return f !== "" && t !== "" ? { from: +f, to: +t } : null;
    })();
    const ok = cur && cur.from < n && cur.to < n;
    $("cagrFrom").innerHTML = opts(ok ? cur.from : -1);
    $("cagrTo").innerHTML = opts(ok ? cur.to : -1);
    $("deltaList").innerHTML = state.deltas.map((d, k) => `
      <div class="delta" data-k="${k}">
        <select data-f="from">${opts(d.from).replace('<option value="">—</option>', "")}</select>
        <select data-f="to">${opts(d.to).replace('<option value="">—</option>', "")}</select>
        <select data-f="mode"><option value="pct"${d.mode === "pct" ? " selected" : ""}>% change</option><option value="abs"${d.mode === "abs" ? " selected" : ""}>Absolute</option></select>
        <button data-f="del" title="Remove">×</button>
      </div>`).join("");
  }
  $("deltaList").addEventListener("change", (ev) => {
    const row = ev.target.closest(".delta"); if (!row) return;
    const d = state.deltas[+row.dataset.k], f = ev.target.dataset.f;
    d[f] = f === "mode" ? ev.target.value : +ev.target.value;
    refresh();
  });
  $("deltaList").addEventListener("click", (ev) => {
    if (ev.target.dataset.f !== "del") return;
    state.deltas.splice(+ev.target.closest(".delta").dataset.k, 1);
    renderArrowControls(); refresh();
  });
  $("addDelta").addEventListener("click", () => {
    const n = state.categories.length;
    if (n < 2) return;
    state.deltas.push({ from: Math.max(0, n - 2), to: n - 1, mode: "pct" });
    renderArrowControls(); refresh();
  });

  // ------------------------------------------------------------------ scale & dots
  function allValues(filter) {
    const out = [];
    const spec = buildSpec();
    spec.series.forEach((s, i) => {
      if (filter && !filter(s, i)) return;
      s.values.forEach((v) => { if (typeof v === "number") out.push(v); });
    });
    return out;
  }
  // Round a number down/up to a "nice" step (1, 2, 2.5, 5 x 10^k)
  function niceStep(span) {
    const p = Math.pow(10, Math.floor(Math.log10(span || 1)));
    const f = span / p;
    return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p;
  }
  $("btnFitScale").addEventListener("click", () => {
    const sec = (s) => state.type === "combo" && s.kind === "line" && s.axis === "secondary";
    let vals = allValues((s) => !sec(s));
    if (state.type === "stacked" || state.type === "combo") {
      // stacked bars: the scale must fit the stack totals, not single segments
      const tot = state.categories.map((_, c) => buildSpec().series.filter((s) => !sec(s) && s.kind !== "line")
        .reduce((a, s) => a + (typeof s.values[c] === "number" ? s.values[c] : 0), 0));
      const lines = allValues((s) => !sec(s) && s.kind === "line" && state.type === "combo");
      vals = tot.concat(lines);
    }
    if (!vals.length) return;
    const lo = Math.min(...vals), hi = Math.max(...vals);
    const step = niceStep((hi - lo) / 4 || Math.abs(lo) * 0.1);
    $("yMin").value = toRaw(Math.floor((lo - step * 0.5) / step) * step, locale);
    $("yMax").value = "";
    setStatus("Scale now starts at " + $("yMin").value + ". Adjust Axis min/max to fine-tune.", "ok");
    refresh();
  });
  $("btnAutoScale").addEventListener("click", () => {
    ["yMin", "yMax", "y2Min", "y2Max"].forEach((id) => { $(id).value = ""; });
    refresh();
  });
  function stepDot(delta) {
    const fs = parseFloat($("fontSize").value) || 10;
    const cur = numOpt("markerSize");
    const base = cur === null ? Math.round(fs * 0.55) : cur;
    $("markerSize").value = String(Math.max(0, Math.min(20, base + delta)));
    refresh();
  }
  $("dotMinus").addEventListener("click", () => stepDot(-1));
  $("dotPlus").addEventListener("click", () => stepDot(1));

  // ------------------------------------------------------------------ options
  document.querySelectorAll(".panel input, .panel select").forEach((el) => {
    if (el.closest("#sheet") || el.closest("#deltaList")) return;
    el.addEventListener("input", () => refresh());
    el.addEventListener("change", () => refresh());
  });
  $("locale").addEventListener("change", () => {
    const old = locale, nw = $("locale").value;
    state.series.forEach((s) => { s.raw = s.raw.map((r) => toRaw(E.parseCell(r, old), nw)); });
    ["yMin", "yMax", "y2Min", "y2Max", "markerSize"].forEach((id) => {
      const v = E.parseCell($(id).value, old);
      $(id).value = typeof v === "number" ? toRaw(v, nw) : "";
    });
    locale = nw;
    renderSheet(); refresh();
  });
  $("chartType").addEventListener("change", () => {
    state.type = $("chartType").value;
    if (state.type === "waterfall" && !state.series[0].raw.some((r) => /^e$/i.test(String(r).trim()))) {
      setStatus("Tip: type e in the last cell for an end total.", "");
    }
    renderSheet(); refresh();
  });

  // ------------------------------------------------------------------ preview + live
  function refresh(opts) {
    clearTimeout(previewTimer);
    previewTimer = setTimeout(drawPreview, 60);
    if ((!opts || opts.live !== false) && target && $("liveUpdate").checked) {
      clearTimeout(liveTimer);
      liveTimer = setTimeout(() => pushUpdate(), 700);
    }
  }

  function drawPreview() {
    try {
      const L = E.layoutChart(buildSpec(), { x: 0, y: 0, w: 600, h: 330 });
      $("preview").innerHTML = E.toSVG(L, { fontName: $("fontName").value.trim() });
    } catch (e) {
      $("preview").innerHTML = `<p style="padding:8px;color:#c00000">Preview error: ${attr(e.message)}</p>`;
    }
  }

  // ------------------------------------------------------------------ PowerPoint actions
  async function guarded(fn) {
    if (!pptReady()) { cannotConnect("PowerPoint API not loaded"); return; }
    if (busy) { pending = true; return; }
    busy = true; ignoreSelUntil = Date.now() + 1500;
    try { await fn(); }
    catch (e) { setStatus(e.message || String(e), "err"); }
    finally {
      busy = false; ignoreSelUntil = Date.now() + 600;
      if (pending) { pending = false; pushUpdate(); }
    }
  }

  function pushUpdate() {
    if (!target) return;
    return guarded(async () => {
      const r = await ChartPPT.updateChart(target, buildSpec());
      target = r;
      setStatus("Chart updated.", "ok");
    });
  }

  $("btnInsert").addEventListener("click", () => guarded(async () => {
    const r = await ChartPPT.insertChart(buildSpec());
    target = { id: r.id, slideId: r.slideId };
    $("btnUpdate").disabled = false;
    setStatus("Chart inserted" + r.note + ". Edits now apply to it" + ($("liveUpdate").checked ? " live." : " when you click Update."), "ok");
  }));

  $("btnUpdate").addEventListener("click", async () => {
    if (!pptReady()) { cannotConnect("PowerPoint API not loaded"); return; }
    if (!target) {
      const r = await ChartPPT.readSelected().catch(() => null);
      if (!r) { setStatus("Select a ChartLite chart on the slide first.", "err"); return; }
      target = { id: r.id, slideId: r.slideId };
    }
    pushUpdate();
  });

  let selTimer = null;
  function onSelectionChanged() {
    clearTimeout(selTimer);
    selTimer = setTimeout(async () => {
      if (busy || Date.now() < ignoreSelUntil) return;
      let r = null;
      try { r = await ChartPPT.readSelected(); } catch (e) { r = null; }
      if (r) {
        if (!target || target.id !== r.id) {
          target = { id: r.id, slideId: r.slideId };
          loadSpec(r.spec);
          setStatus("Editing the selected chart.", "ok");
        }
        $("btnUpdate").disabled = false;
      } else if (target) {
        target = null;
        $("btnUpdate").disabled = true;
        setStatus("No chart selected — Insert creates a new chart from this data.");
      }
    }, 150);
  }

  // ------------------------------------------------------------------ boot
  renderSheet();
  writeOptions(E.mergeOptions({ options: {} }));
  drawPreview();

  const VERSION = "1.1.0";
  $("ver").textContent = "ChartLite v" + VERSION;
  $("btnReload").addEventListener("click", () => location.reload());

  // True only once Office has finished loading PowerPoint's API (PowerPoint.run exists).
  const pptReady = () => typeof window.PowerPoint !== "undefined" && typeof window.PowerPoint.run === "function";

  function cannotConnect(detail) {
    $("btnInsert").disabled = true;
    $("btnUpdate").disabled = true;
    $("btnReload").hidden = false;
    setStatus("ChartLite can't reach PowerPoint's add-in library from Microsoft (" + detail +
      "). Check your internet/VPN, then click Reload. If it keeps happening, close the panel and open it again from the Home tab.", "err");
  }

  let booted = false;
  function boot(info) {
    booted = true;
    inOffice = info && info.host === Office.HostType.PowerPoint;
    if (!inOffice) {
      $("btnInsert").disabled = true;
      setStatus("Preview mode — open this pane inside PowerPoint to insert charts.");
      return;
    }
    if (!pptReady()) {
      // Office answered but its PowerPoint part didn't load yet: wait a little before giving up
      let tries = 0;
      const t = setInterval(() => {
        if (pptReady()) { clearInterval(t); boot(info); }
        else if (++tries > 40) { clearInterval(t); cannotConnect("PowerPoint API missing"); }
      }, 250);
      return;
    }
    if (!ChartPPT.supports("1.8")) {
      $("btnInsert").disabled = true;
      setStatus("This PowerPoint is too old for ChartLite (needs Microsoft 365 Version 2504 or newer).", "err");
      return;
    }
    $("btnInsert").disabled = false;
    $("btnReload").hidden = true;
    setStatus(ChartPPT.supports("1.10") ? "Ready — paste data and click Insert chart." :
      "Ready. Note: update Office to Version 2601+ for solid rising lines (older builds draw them dotted).", "ok");
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    onSelectionChanged();
  }

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(boot);
    // office.js loaded but never finished starting (slow or blocked network)
    setTimeout(() => { if (!booted) cannotConnect("timed out"); }, 12000);
  } else if (/_host_Info=/i.test(location.search) || window.external && "GetContext" in window.external) {
    cannotConnect("office.js did not load");
  } else {
    boot(null);
  }
})();

/*
 * ChartLite engine — turns a chart spec into drawing primitives (points).
 * Pure JS, no dependencies. Used by the task-pane preview (SVG) and by the
 * PowerPoint renderer (native shapes), so what you preview is what you get.
 *
 * Primitive types:
 *   rect  {x,y,w,h, fill, stroke, strokeW, round}
 *   line  {x1,y1,x2,y2, color, w, dash}
 *   tri   {x,y,w,h, dir:'down'|'up', fill}
 *   dot   {x,y,w,h, fill, stroke, strokeW}
 *   text  {x,y,w,h, text, size, color, bold, align, wrap, fill, stroke}
 */
(function (root) {
  "use strict";

  const NEG_RED = "#C00000";

  const DEFAULT_PALETTE = [
    "#1F3B57", "#3E6E9E", "#7FA7CF", "#B9D0E8",
    "#E0A458", "#5B8C5A", "#8C8C8C", "#C9C9C9",
  ];

  const DEFAULTS = {
    decimals: 0,
    prefix: "",
    unit: "",
    locale: "en",          // 'en' => 1,234.5   'id' => 1.234,5
    negStyle: "minus",     // 'minus' => -12    'paren' => (12)
    fontSize: 10,
    fontName: "",
    showTotals: true,
    showLabels: true,
    showLegend: true,
    connectors: true,
    yMin: null, yMax: null,      // manual value-axis range (blank = auto)
    y2Min: null, y2Max: null,    // combo Axis 2 range
    smooth: false,               // smooth (curved) lines
    markerSize: null,            // dot diameter in pt (null = auto, 0 = no dots)
    markerFill: "solid",         // solid | white | none
    legendPos: "auto",           // auto | top | start (left of bars) | end (line ends / right of bars)
    labelPos: "auto",            // auto | above | center | below | none (series can override)
    markerShape: "circle",       // circle | square | diamond | triangle | none
    showAxis: false,             // value axis with tick labels
    lineWidth: 1.75,
    lineDash: "solid",           // solid | dash | longDash | roundDot | dashDot (series can override)
    holeSize: 0.55,              // donut: inner radius as a share of the outer radius
    labelContent: "auto",        // auto | value | pct | valuePct | pctValue | category | series | seriesValue | seriesPct
    scalePow: 0,                 // show numbers divided by 10^scalePow (negative = multiplied, e.g. -2 = x100)
    barWidth: 0.62,        // fraction of the category slot used by bars
    pct100Labels: true,    // 100% chart: label segments with share instead of value
    cagr: null,            // {from, to}
    deltas: [],            // [{from, to, mode:'pct'|'abs'}]
    palette: DEFAULT_PALETTE,
    wf: { start: "#1F3B57", up: "#5B8C5A", down: NEG_RED, total: "#1F3B57" },
    textColor: "#262626",
    axisColor: "#7F7F7F",
  };

  // ---------------------------------------------------------------- numbers
  function formatNumber(v, o, extra) {
    extra = extra || {};
    if (v === null || v === undefined || isNaN(v)) return "";
    // number scale (thousands, millions, ...): applies to values, never to percentages
    const sp = o.scalePow || 0;
    if (sp && !extra.noScale && extra.unit !== "%") v = v / Math.pow(10, sp);
    const dec = extra.decimals !== undefined ? extra.decimals : o.decimals;
    const neg = v < 0;
    let s = Math.abs(v).toFixed(dec);
    let [ip, fp] = s.split(".");
    const th = o.locale === "id" ? "." : ",";
    const ds = o.locale === "id" ? "," : ".";
    ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, th);
    s = fp ? ip + ds + fp : ip;
    const unit = extra.unit !== undefined ? extra.unit : o.unit;
    const prefix = extra.prefix !== undefined ? extra.prefix : o.prefix;
    s = prefix + s + unit;
    if (neg && Number(Math.abs(v).toFixed(dec)) !== 0) {
      s = o.negStyle === "paren" ? "(" + s + ")" : "-" + s;
    } else if (extra.sign && v > 0) {
      s = "+" + s;
    }
    return s;
  }

  // Parses a datasheet cell. Returns number, 'e' (waterfall total) or null.
  function parseCell(raw, locale) {
    if (raw === null || raw === undefined) return null;
    if (typeof raw === "number") return isFinite(raw) ? raw : null;
    let s = String(raw).replace(/\u00a0/g, " ").replace(/\u2212/g, "-").trim();
    if (!s || /^[-–—]$/.test(s) || /^#/.test(s)) return null;   // blank, dash, #N/A, #DIV/0!
    if (/^e$/i.test(s)) return "e";
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    s = s.replace(/[%\s]|Rp\.?|IDR|\$/gi, "");
    if (/^-\(.*\)$/.test(s)) s = s.slice(2, -1);
    if (s.startsWith("-")) { neg = !neg; s = s.slice(1); }
    if (locale === "id") s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
    if (s.startsWith("+")) s = s.slice(1);
    const n = Number(s);
    if (!isFinite(n) || s === "") return null;
    return neg ? -n : n;
  }

  // ---------------------------------------------------------------- helpers
  function luminance(hex) {
    const h = hex.replace("#", "");
    const r = parseInt(h.substr(0, 2), 16) / 255;
    const g = parseInt(h.substr(2, 2), 16) / 255;
    const b = parseInt(h.substr(4, 2), 16) / 255;
    const f = (c) => (c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4));
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b);
  }
  const onColor = (fill) => (luminance(fill) < 0.4 ? "#FFFFFF" : "#262626");
  const textW = (s, fs) => String(s).length * fs * 0.56 + 6;

  // dash names used in specs -> PowerPoint line dash styles, and the SVG equivalent (in line widths)
  const DASH_PPT = { solid: null, dash: "Dash", longDash: "LongDash", roundDot: "RoundDot", dashDot: "DashDot" };
  const DASH_SVG = { Dash: [4, 3], LongDash: [8, 3], RoundDot: [0.01, 2], DashDot: [4, 3, 1, 3], SquareDot: [1, 1] };
  function dashAttrs(pptDash, w, solidCap) {
    if (!pptDash || !DASH_SVG[pptDash]) return ` stroke-linecap="${solidCap || "butt"}"`;
    const arr = DASH_SVG[pptDash].map((k) => +(k * Math.max(w, 0.75)).toFixed(2)).join(" ");
    return ` stroke-dasharray="${arr}" stroke-linecap="${pptDash === "RoundDot" ? "round" : "butt"}"`;
  }
  // Cuts a polyline into the visible pieces of a dash pattern (for Office builds without vector curves)
  function dashPolyline(path, pptDash, w) {
    const pat = (DASH_SVG[pptDash] || [1]).map((k) => Math.max(k * w, pptDash === "RoundDot" && k < 0.1 ? w * 0.6 : 0.3));
    const out = [];
    let pi = 0, left = pat[0], on = true;
    for (let k = 0; k < path.length - 1; k++) {
      let ax = path[k].x, ay = path[k].y;
      const bx = path[k + 1].x, by = path[k + 1].y;
      let seg = Math.hypot(bx - ax, by - ay);
      while (seg > 1e-6) {
        const step = Math.min(seg, left), t = step / seg;
        const cx2 = ax + (bx - ax) * t, cy2 = ay + (by - ay) * t;
        if (on) out.push([ax, ay, cx2, cy2]);
        ax = cx2; ay = cy2; seg -= step; left -= step;
        if (left <= 1e-6) { pi = (pi + 1) % pat.length; left = pat[pi]; on = pi % 2 === 0; }
      }
    }
    // merge consecutive "on" pieces that continue each other (fewer shapes)
    const merged = [];
    out.forEach((q) => {
      const last = merged[merged.length - 1];
      if (last && Math.abs(last[last.length - 2] - q[0]) < 1e-6 && Math.abs(last[last.length - 1] - q[1]) < 1e-6) last.push(q[2], q[3]);
      else merged.push(q.slice());
    });
    const pieces = [];
    merged.forEach((q) => { for (let k = 0; k + 3 < q.length; k += 2) pieces.push([q[k], q[k + 1], q[k + 2], q[k + 3]]); });
    return pieces;
  }

  function mergeOptions(spec) {
    const o = Object.assign({}, DEFAULTS, spec.options || {});
    o.wf = Object.assign({}, DEFAULTS.wf, (spec.options || {}).wf || {});
    o.palette = (o.palette && o.palette.length) ? o.palette : DEFAULT_PALETTE;
    o.deltas = o.deltas || [];
    return o;
  }

  function seriesColor(spec, o, i) {
    const s = spec.series[i];
    return (s && s.color) || o.palette[i % o.palette.length];
  }

  function num(v) { return typeof v === "number" && isFinite(v) ? v : 0; }

  // ---------------------------------------------------------------- layout
  function layoutChart(spec, box) {
    const o = mergeOptions(spec);
    const fs = o.fontSize;
    const P = [];
    const cats = spec.categories || [];
    const n = cats.length;
    if (!n || !spec.series || !spec.series.length) {
      P.push({ type: "text", x: box.x, y: box.y, w: box.w, h: box.h, text: "No data", size: fs, color: o.textColor, align: "center" });
      return { prims: P, box };
    }

    const type = spec.type || "stacked";
    if (type === "donut" || type === "pie") return layoutDonut(spec, box, o);
    const isWF = type === "waterfall";
    const barIdx = [], lineIdx = [];
    spec.series.forEach((s, i) => {
      const kind = type === "line" ? "line" : type === "combo" ? (s.kind || "bar") : "bar";
      (kind === "line" ? lineIdx : barIdx).push(i);
    });
    if (isWF) { barIdx.length = 0; barIdx.push(0); lineIdx.length = 0; }

    // ---- per-category geometry in value space
    const cat = [];              // {pos, neg, total, segs:[{i,v,lo,hi}], top, bottom}
    let wfRun = 0;
    for (let c = 0; c < n; c++) {
      const entry = { segs: [], pos: 0, neg: 0, total: 0, kind: "bar" };
      if (isWF) {
        const raw = spec.series[0].values[c];
        if (raw === "e") {
          entry.segs.push({ i: 0, v: wfRun, lo: Math.min(0, wfRun), hi: Math.max(0, wfRun), wfKind: "total" });
          entry.total = wfRun; entry.kind = "total";
        } else {
          const v = num(raw);
          if (c === 0) {
            entry.segs.push({ i: 0, v, lo: Math.min(0, v), hi: Math.max(0, v), wfKind: "start" });
            wfRun = v; entry.kind = "start";
          } else {
            const a = wfRun, b = wfRun + v;
            entry.segs.push({ i: 0, v, lo: Math.min(a, b), hi: Math.max(a, b), wfKind: v >= 0 ? "up" : "down" });
            wfRun = b; entry.kind = v >= 0 ? "up" : "down";
          }
          entry.total = wfRun;
        }
        entry.top = Math.max(...entry.segs.map((s) => s.hi));
        entry.bottom = Math.min(...entry.segs.map((s) => s.lo));
      } else if (type === "clustered") {
        barIdx.forEach((i) => {
          const v = num(spec.series[i].values[c]);
          entry.segs.push({ i, v, lo: Math.min(0, v), hi: Math.max(0, v) });
          entry.total += v;
        });
        entry.top = Math.max(0, ...entry.segs.map((s) => s.hi));
        entry.bottom = Math.min(0, ...entry.segs.map((s) => s.lo));
      } else {
        let p = 0, q = 0, sumAbs = 0, tot = 0;
        barIdx.forEach((i) => { const v = num(spec.series[i].values[c]); sumAbs += Math.abs(v); tot += v; });
        barIdx.forEach((i) => {
          let v = num(spec.series[i].values[c]);
          const shown = v;
          if (type === "stacked100") v = sumAbs ? (v / sumAbs) * 100 : 0;
          if (v >= 0) { entry.segs.push({ i, v: shown, share: v, lo: p, hi: p + v }); p += v; }
          else { entry.segs.push({ i, v: shown, share: v, lo: q + v, hi: q }); q += v; }
        });
        entry.pos = p; entry.neg = q; entry.total = tot;
        entry.top = p; entry.bottom = q;
      }
      cat.push(entry);
    }

    // ---- value range
    let vmin = 0, vmax = 0;
    if (barIdx.length) cat.forEach((e) => { vmin = Math.min(vmin, e.bottom); vmax = Math.max(vmax, e.top); });
    const secondary = lineIdx.filter((i) => type === "combo" && spec.series[i].axis === "secondary");
    const primaryLines = lineIdx.filter((i) => secondary.indexOf(i) < 0);
    let lmin = Infinity, lmax = -Infinity;
    primaryLines.forEach((i) => spec.series[i].values.forEach((v) => {
      if (typeof v === "number") { lmin = Math.min(lmin, v); lmax = Math.max(lmax, v); }
    }));
    if (primaryLines.length && isFinite(lmin)) {
      if (!barIdx.length) {
        // pure line chart: zoom to data range instead of forcing zero
        const span = (lmax - lmin) || Math.abs(lmax) || 1;
        vmin = lmin >= 0 && lmin - span * 0.6 < 0 ? 0 : lmin - span * 0.25;
        vmax = lmax + span * 0.1;
      } else { vmin = Math.min(vmin, lmin); vmax = Math.max(vmax, lmax); }
    }
    // waterfall axis break: when every floating bar sits far above zero, cut the scale
    let broken = false;
    if (isWF && o.axisBreak !== false && vmin === 0) {
      let low = Infinity;
      cat.forEach((e) => e.segs.forEach((s) => {
        low = Math.min(low, s.wfKind === "up" || s.wfKind === "down" ? s.lo : s.hi);
      }));
      if (low > 0 && (vmax - low) / vmax < 0.5) { vmin = Math.max(0, low - (vmax - low) * 1.2); broken = vmin > 0; }
    }
    const isNum = (v) => typeof v === "number" && isFinite(v);
    if (isNum(o.yMin)) { vmin = o.yMin; broken = false; }
    if (isNum(o.yMax)) vmax = o.yMax;
    if (vmax <= vmin) vmax = vmin + (Math.abs(vmin) || 1) * 0.1;

    // ---- vertical bands (top → bottom)
    const multiSeries = !isWF && spec.series.length > 1;
    const pureLine = type === "line";
    const stackType = type === "stacked" || type === "stacked100";
    const legendPos = o.legendPos === "auto" || !o.legendPos
      ? (pureLine ? "end" : stackType && barIdx.length > 1 ? (o.showAxis ? "end" : "start") : "top") : o.legendPos;
    const endLegend = o.showLegend && pureLine && legendPos === "end";
    const sideLegend = o.showLegend && multiSeries && stackType && (legendPos === "start" || legendPos === "end") ? legendPos : null;
    const legendH = o.showLegend && multiSeries && !endLegend && !sideLegend ? fs * 2.0 : 0;
    const seriesName = (i) => spec.series[i].name || "Series " + (i + 1);
    const sideNameW = sideLegend ? Math.max(...barIdx.map((i) => textW(seriesName(i), fs))) + fs * 0.6 : 0;

    // value axis: "nice" ticks inside [vmin, vmax]
    const ticks = [];
    let tickDec = 0;
    if (o.showAxis) {
      const raw = (vmax - vmin) / 5;
      const p10 = Math.pow(10, Math.floor(Math.log10(raw || 1)));
      const f = raw / p10;
      const step = (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p10;
      while (tickDec < 6 && Math.abs(step * Math.pow(10, tickDec) - Math.round(step * Math.pow(10, tickDec))) > 1e-6) tickDec++;
      for (let t = Math.ceil(vmin / step - 1e-9) * step; t <= vmax + step * 1e-6; t += step) ticks.push(Math.abs(t) < step * 1e-9 ? 0 : t);
    }
    const tickLabel = (t) => formatNumber(t, o, { decimals: tickDec });
    const axisW = ticks.length ? Math.max(...ticks.map((t) => textW(tickLabel(t), fs))) + fs * 0.5 : 0;
    const arrows = [];
    if (o.cagr && o.cagr.from !== undefined && o.cagr.to !== undefined && o.cagr.from !== o.cagr.to && !isWF)
      arrows.push({ from: +o.cagr.from, to: +o.cagr.to, mode: "cagr" });
    (o.deltas || []).forEach((d) => { if (d.from !== d.to) arrows.push({ from: +d.from, to: +d.to, mode: d.mode || "pct" }); });
    const valid = arrows.filter((a) => a.from >= 0 && a.to >= 0 && a.from < n && a.to < n);
    const arrowRowH = fs * 2.3;
    const arrowH = valid.length * arrowRowH;
    const topPad = fs * 1.7;
    const hasNeg = vmin < 0;
    const bottomPad = hasNeg ? fs * 1.6 : 0;
    const catH = fs * 2.8;

    const plotTop = box.y + legendH + arrowH + topPad;
    const plotBottom = box.y + box.h - catH - bottomPad;
    const plotH = Math.max(10, plotBottom - plotTop);
    const secondaryBand = type === "combo" && lineIdx.some((i) => spec.series[i].axis === "secondary");
    const barH = secondaryBand ? plotH * 0.62 : plotH;   // primary scale height
    const Y = (v) => plotBottom - ((v - vmin) / (vmax - vmin)) * barH;
    // clamp to the plot so values outside a manual axis range don't spill over labels
    const Yc = (v) => Math.min(plotBottom, Math.max(plotTop - fs * 0.5, Y(v)));
    const yBase = Yc(Math.max(0, vmin));

    let smin = 0, smax = 1;
    if (secondary.length) {
      smin = Infinity; smax = -Infinity;
      secondary.forEach((i) => spec.series[i].values.forEach((v) => {
        if (typeof v === "number") { smin = Math.min(smin, v); smax = Math.max(smax, v); }
      }));
      const span = (smax - smin) || Math.abs(smax) || 1;
      smin -= span * 0.05; smax += span * 0.05;
      if (isNum(o.y2Min)) smin = o.y2Min;
      if (isNum(o.y2Max)) smax = o.y2Max;
      if (smax <= smin) smax = smin + (Math.abs(smin) || 1) * 0.1;
    }
    // secondary lines live in the upper part of the plot so they don't sit on the bars
    const Y2 = (v) => plotTop + (1 - (v - smin) / (smax - smin)) * plotH * 0.26;

    const msEarly = isNum(o.markerSize) ? Math.max(0, o.markerSize) : fs * 0.5;
    const endW = endLegend ? Math.max(...lineIdx.map((i) => textW(spec.series[i].name || "Series " + (i + 1), fs))) + msEarly / 2 + fs * 1.6 : 0;
    const leftW = axisW + (sideLegend === "start" ? sideNameW : 0);
    const rightW = endW + (sideLegend === "end" ? sideNameW : 0);
    const px0 = box.x + leftW;
    const slot = (box.w - leftW - rightW) / n;
    const pxEnd = px0 + slot * n;
    const barW = slot * o.barWidth;
    const cx = (c) => px0 + slot * (c + 0.5);
    const labelTop = [];   // per category: highest y used (for arrows)
    const ms = o.markerShape === "none" ? 0 : msEarly;
    const shape = o.markerShape || "circle";
    const posFor = (i) => {
      const sp = spec.series[i] && spec.series[i].labelPos;
      return sp && sp !== "auto" ? sp : (o.labelPos || "auto");
    };
    const dotStyle = (col, size) => o.markerFill === "none"
      ? { fill: null, stroke: col, strokeW: Math.max(1, size * 0.18) }
      : o.markerFill === "white"
        ? { fill: "#FFFFFF", stroke: col, strokeW: Math.max(1, size * 0.18) }
        : { fill: col, stroke: "#FFFFFF", strokeW: 0.75 };

    // ---- legend
    if (legendH) {
      let lx = box.x;
      const ly = box.y + legendH * 0.15;
      spec.series.forEach((s, i) => {
        const col = seriesColor(spec, o, i);
        const isLine = lineIdx.indexOf(i) >= 0;
        if (isLine) {
          P.push({ type: "line", x1: lx, y1: ly + fs * 0.6, x2: lx + fs * 1.4, y2: ly + fs * 0.6, color: col, w: s.lineWeight || o.lineWidth || 1.75, dash: DASH_PPT[s.dash || o.lineDash || "solid"] || undefined });
          if (ms > 0) P.push(Object.assign({ type: "dot", shape, x: lx + fs * 0.45, y: ly + fs * 0.35, w: fs * 0.5, h: fs * 0.5 }, dotStyle(col, fs * 0.5)));
        } else {
          P.push({ type: "rect", x: lx + fs * 0.2, y: ly + fs * 0.1, w: fs, h: fs, fill: col });
        }
        const nm = s.name || "Series " + (i + 1);
        const tw = textW(nm, fs);
        P.push({ type: "text", x: lx + fs * 1.6, y: ly - fs * 0.2, w: tw, h: fs * 1.6, text: nm, size: fs, color: o.textColor, align: "left" });
        lx += fs * 1.6 + tw + fs;
      });
    }

    // ---- baseline
    if (barIdx.length) {
      P.push({ type: "line", x1: px0, y1: yBase, x2: pxEnd, y2: yBase, color: o.axisColor, w: 0.75 });
    }

    // ---- label content: absolute, %, both, category or series name
    const barDen = cats.map((_, c) => barIdx.reduce((a, i) => a + Math.abs(num(spec.series[i].values[c])), 0));
    const lineDen = cats.map((_, c) => lineIdx.reduce((a, i) => { const v = spec.series[i].values[c]; return a + (typeof v === "number" ? Math.abs(v) : 0); }, 0));
    const wfDen = isWF ? Math.abs(num(spec.series[0].values[0])) : 0;
    const labelText = (i, c, v, den, dflt, extra) => contentText(spec, o, i, c, v, den, dflt, extra);

    // ---- bar labels at a chosen position: above (outside end), center, below (inside base)
    const lhB = fs * 1.4;
    const barLabel = (pos, a) => {
      // a: {x, w, y1, y2, lbl, col, neg, wideX, wideW, stackedTop}
      if (pos === "none") return null;
      const h = a.y2 - a.y1;
      let y, inside;
      if (pos === "above") {
        if (a.stackedTop && h >= lhB) { y = a.y1 + 0.5; inside = true; }        // stacked: inside top of segment
        else { y = a.neg ? a.y2 + fs * 0.1 : a.y1 - lhB - fs * 0.05; inside = false; }
      } else if (pos === "below") {
        if (h >= lhB) { y = a.y2 - lhB - 0.5; inside = true; } else { y = a.y2 + fs * 0.1; inside = false; }
      } else {
        if (h < fs * 1.15) return null;                                          // center: only if it fits
        y = a.y1 + h / 2 - lhB / 2; inside = true;
      }
      P.push({
        type: "text", x: inside ? a.x : a.wideX, y, w: inside ? a.w : a.wideW, h: lhB, text: a.lbl, size: fs,
        color: a.neg ? NEG_RED : inside ? onColor(a.col) : o.textColor, align: "center",
        fill: inside && a.neg && luminance(a.col) < 0.4 ? "#FFFFFF" : undefined,
      });
      return y;
    };

    // ---- bars
    if (barIdx.length) {
      for (let c = 0; c < n; c++) {
        const e = cat[c];
        let top = Infinity;
        if (type === "clustered") {
          const k = e.segs.length;
          const bw = barW / k;
          e.segs.forEach((s, j) => {
            const x = cx(c) - barW / 2 + j * bw;
            const y1 = Yc(s.hi), y2 = Yc(s.lo);
            if (y2 - y1 < 0.3 && s.v !== 0) return;   // entirely outside a manual axis range
            const col = seriesColor(spec, o, s.i);
            P.push({ type: "rect", x: x + bw * 0.04, y: y1, w: bw * 0.92, h: Math.max(0.5, y2 - y1), fill: col, role: "bar" });
            if (o.showLabels) {
              const pos = posFor(s.i);
              const ly = barLabel(pos === "auto" ? "above" : pos, { x: x + bw * 0.04, w: bw * 0.92, y1, y2, lbl: labelText(s.i, c, s.v, barDen[c], formatNumber(s.v, o)), col, neg: s.v < 0, wideX: x - fs, wideW: bw + fs * 2 });
              if (ly !== null) top = Math.min(top, ly);
            }
            top = Math.min(top, y1);
          });
        } else {
          e.segs.forEach((s) => {
            const y1 = Yc(s.hi), y2 = Yc(s.lo);
            if (y2 - y1 < 0.3 && s.v !== 0) return;
            let col = seriesColor(spec, o, s.i);
            if (isWF) col = o.wf[s.wfKind] || col;
            const x = cx(c) - barW / 2;
            const h = Math.max(0.5, y2 - y1);
            P.push({ type: "rect", x, y: y1, w: barW, h, fill: col, role: "bar" });
            if (broken && (s.wfKind === "start" || s.wfKind === "total") && h > fs * 3) {
              // axis-break marker: white band across the bar near its foot
              const by = plotBottom - Math.min(h * 0.25, fs * 2.2);
              P.push({ type: "rect", x: x - 2, y: by, w: barW + 4, h: fs * 0.45, fill: "#FFFFFF" });
              P.push({ type: "line", x1: x - 2, y1: by, x2: x + barW + 2, y2: by, color: o.axisColor, w: 0.75 });
              P.push({ type: "line", x1: x - 2, y1: by + fs * 0.45, x2: x + barW + 2, y2: by + fs * 0.45, color: o.axisColor, w: 0.75 });
            }
            top = Math.min(top, y1);
            if (!o.showLabels) return;
            let val = s.v, lbl;
            if (type === "stacked100" && o.pct100Labels) lbl = formatNumber(s.share, o, { unit: "%", prefix: "" });
            else if (isWF && (s.wfKind === "up" || s.wfKind === "down")) lbl = formatNumber(val, o, { sign: true });
            else lbl = formatNumber(val, o);
            lbl = isWF ? labelText(s.i, c, val, wfDen, lbl, s.wfKind === "up" || s.wfKind === "down" ? { sign: true } : undefined)
                       : labelText(s.i, c, val, barDen[c], lbl);
            const negLbl = (type === "stacked100" ? s.share : val) < 0;
            const fitsInside = h >= fs * 1.35 && barW >= textW(lbl, fs) * 0.8;
            const pos = posFor(s.i);
            const single = barIdx.length === 1 && type !== "stacked100";
            if (isWF) {
              if (pos === "auto") {
                // waterfall: label above the bar (below for decreases)
                const ly = s.wfKind === "down" ? y2 + fs * 0.1 : y1 - fs * 1.5;
                P.push({ type: "text", x: cx(c) - slot / 2, y: ly, w: slot, h: fs * 1.4, text: lbl, size: fs, color: negLbl ? NEG_RED : o.textColor, align: "center" });
                top = Math.min(top, ly);
              } else {
                const ly = barLabel(pos, { x, w: barW, y1, y2, lbl, col, neg: negLbl, wideX: cx(c) - slot / 2, wideW: slot });
                if (ly !== null) top = Math.min(top, ly);
              }
            } else if (single) {
              // plain bar: auto/above handled with the total above the bar further down
              if (pos === "center" || pos === "below") barLabel(pos, { x, w: barW, y1, y2, lbl, col, neg: negLbl, wideX: cx(c) - slot / 2, wideW: slot });
            } else if (pos === "auto" || pos === "center") {
              if (fitsInside) P.push({
                type: "text", x, y: y1, w: barW, h, text: lbl, size: fs,
                color: negLbl ? NEG_RED : onColor(col), align: "center",
                fill: negLbl && luminance(col) < 0.4 ? "#FFFFFF" : undefined,
              });
            } else {
              barLabel(pos, { x, w: barW, y1, y2, lbl, col, neg: negLbl, wideX: x, wideW: barW, stackedTop: true });
            }
          });
          // totals above the stack
          if (!isWF && o.showTotals && (barIdx.length > 1 || type === "stacked100")) {
            const t = e.total;
            const lbl = formatNumber(t, o);
            const ly = t >= 0 || type === "stacked100" ? Yc(e.top) - fs * 1.55 : Yc(e.bottom) + fs * 0.15;
            P.push({ type: "text", x: cx(c) - slot / 2, y: ly, w: slot, h: fs * 1.45, text: lbl, size: fs, color: t < 0 ? NEG_RED : o.textColor, align: "center" });
            top = Math.min(top, ly);
          } else if (!isWF && o.showLabels && barIdx.length === 1 && type !== "stacked100" && ["auto", "above"].indexOf(posFor(barIdx[0])) >= 0) {
            // single-series stacked == plain bar: put value above instead of inside
            const t = e.total;
            const ly = t >= 0 ? Yc(e.top) - fs * 1.5 : Yc(e.bottom) + fs * 0.1;
            P.push({ type: "text", x: cx(c) - slot / 2, y: ly, w: slot, h: fs * 1.4, text: labelText(barIdx[0], c, t, barDen[c], formatNumber(t, o)), size: fs, color: t < 0 ? NEG_RED : o.textColor, align: "center" });
            top = Math.min(top, ly);
          }
        }
        labelTop[c] = top;
      }

      if (sideLegend) {
        const c = sideLegend === "start" ? 0 : n - 1;
        const items = cat[c].segs.map((sg) => ({ i: sg.i, y: (Yc(sg.hi) + Yc(sg.lo)) / 2 })).sort((a, b) => a.y - b.y);
        const gap = fs * 1.25;
        for (let k = 1; k < items.length; k++) if (items[k].y - items[k - 1].y < gap) items[k].y = items[k - 1].y + gap;
        items.forEach((it) => {
          const nm = seriesName(it.i), w = textW(nm, fs);
          const x = sideLegend === "start"
            ? (ticks.length ? px0 - axisW - fs * 0.2 - w : cx(0) - barW / 2 - fs * 0.4 - w)
            : cx(n - 1) + barW / 2 + fs * 0.4;
          P.push({ type: "text", x, y: it.y - fs * 0.65, w, h: fs * 1.3, text: nm, size: fs, color: o.textColor, align: sideLegend === "start" ? "right" : "left" });
        });
      }

      // waterfall connectors
      if (isWF && o.connectors) {
        for (let c = 0; c < n - 1; c++) {
          const level = Yc(cat[c].total);
          P.push({ type: "line", x1: cx(c) + barW / 2, y1: level, x2: cx(c + 1) - barW / 2, y2: level, color: o.axisColor, w: 0.75, dash: "Dash" });
        }
      }
    }

    // ---- lines
    const lineLabels = [];   // collected first so overlapping labels can be moved apart
    const polylines = [];    // every drawn line path, so labels can keep clear of them
    lineIdx.forEach((i) => {
      const s = spec.series[i];
      const col = seriesColor(spec, o, i);
      const yf = secondary.indexOf(i) >= 0 ? Y2 : Yc;
      const pts = [];
      for (let c = 0; c < n; c++) {
        const v = s.values[c];
        pts.push(typeof v === "number" ? { x: cx(c), y: yf(v), v, c } : null);
      }
      // split into runs of consecutive values (blank cells break the line)
      const runs = [];
      let run = [];
      pts.forEach((p) => { if (p) run.push(p); else { if (run.length) runs.push(run); run = []; } });
      if (run.length) runs.push(run);
      runs.forEach((r) => {
        const smooth = o.smooth && r.length > 2;
        const path = smooth ? smoothPath(r, 2.5) : r;
        polylines.push(path);
        const lw = s.lineWeight || o.lineWidth || 1.75;
        const dash = DASH_PPT[s.dash || o.lineDash || "solid"] || null;
        if (smooth) {
          // exact curve (drawn as one vector graphic where Office supports it) ...
          const xs = path.map((q) => q.x), ys = path.map((q) => q.y);
          P.push({ type: "curve", d: bezierPath(r), color: col, w: lw, dash,
            bbox: { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) } });
        }
        if (smooth && dash) {
          // ... fallback pieces for dashed curves: cut the pattern ourselves so it runs on smoothly
          dashPolyline(path, dash, lw).forEach(([x1, y1, x2, y2]) => P.push({ type: "line", x1, y1, x2, y2, color: col, w: lw, smoothPiece: true }));
        } else {
          // ... and short straight pieces as the fallback for Office versions without it
          for (let k = 0; k < path.length - 1; k++) {
            P.push({ type: "line", x1: path[k].x, y1: path[k].y, x2: path[k + 1].x, y2: path[k + 1].y, color: col, w: lw, join: !dash, smoothPiece: smooth, dash: smooth ? undefined : dash || undefined });
          }
        }
      });
      const lpos = posFor(i);
      pts.forEach((p) => {
        if (!p) return;
        if (ms > 0 && !(o.showLabels && lpos === "center")) P.push(Object.assign({ type: "dot", shape, x: p.x - ms / 2, y: p.y - ms / 2, w: ms, h: ms }, dotStyle(col, ms)));
        if (o.showLabels && lpos !== "none") {
          const ux = s.unit !== undefined ? { unit: s.unit } : {};
          const lbl = labelText(i, p.c, p.v, lineDen[p.c], formatNumber(p.v, o, ux), ux);
          const overBars = barIdx.length && secondary.indexOf(i) < 0;   // only then can labels collide with bars
          lineLabels.push({ p, lbl, col, overBars, series: i, pos: lpos });
        }
      });
    });
    // think-cell style placement, per category: the highest point's label goes above it,
    // the lowest point's below it. A label never sits on a line or another label: if its
    // preferred spot is crossed by a line it tries the other side, then diagonally beside
    // the dot, and is left out only if nothing is free. No boxes behind the numbers.
    const lh = fs * 1.3;
    const gapA = Math.max(fs * 0.25, ms / 2 + fs * 0.2);   // clearance between dot and label
    const lastLabel = {};
    const placed = [];   // all label rects so far
    const segHitsRect = (x1, y1, x2, y2, r) => {
      // Liang-Barsky clip of the segment against the rectangle
      let t0 = 0, t1 = 1;
      const dx = x2 - x1, dy = y2 - y1;
      const pq = [[-dx, x1 - r.x], [dx, r.x + r.w - x1], [-dy, y1 - r.y], [dy, r.y + r.h - y1]];
      for (const [p, q] of pq) {
        if (p === 0) { if (q < 0) return false; continue; }
        const t = q / p;
        if (p < 0) { if (t > t1) return false; if (t > t0) t0 = t; }
        else { if (t < t0) return false; if (t < t1) t1 = t; }
      }
      return true;
    };
    const lineHits = (r) => polylines.some((pl) => {
      for (let k = 0; k < pl.length - 1; k++) if (segHitsRect(pl[k].x, pl[k].y, pl[k + 1].x, pl[k + 1].y, r)) return true;
      return false;
    });
    const dotHits = (r) => lineLabels.some((L) => L.p.x + ms / 2 > r.x && L.p.x - ms / 2 < r.x + r.w && L.p.y + ms / 2 > r.y && L.p.y - ms / 2 < r.y + r.h);
    const labelHits = (r) => placed.some((q) => r.x < q.x + q.w && r.x + r.w > q.x && r.y < q.y + q.h && r.y + r.h > q.y);
    for (let c = 0; c < n; c++) {
      const here = lineLabels.filter((L) => L.p.c === c).sort((a, b) => a.p.y - b.p.y);
      // fixed positions first, so the automatic ones can avoid them
      here.filter((L) => L.pos !== "auto").forEach((L) => {
        const lw = textW(L.lbl, fs) - 2;
        let r;
        if (L.pos === "center") r = { x: L.p.x - lw / 2 - 2, y: L.p.y - lh / 2, w: lw + 4, h: lh };
        else if (L.pos === "below") r = { x: L.p.x - lw / 2, y: L.p.y + gapA, w: lw, h: lh };
        else r = { x: L.p.x - lw / 2, y: L.p.y - gapA - lh, w: lw, h: lh };
        placed.push(r);
        P.push({ type: "text", x: r.x, y: r.y, w: r.w, h: r.h, text: L.lbl, size: fs, color: L.p.v < 0 ? NEG_RED : o.textColor, align: "center", fill: L.pos === "center" ? "#FFFFFF" : undefined, stroke: L.pos === "center" ? L.col : undefined, round: L.pos === "center" });
        if (r.y < L.p.y) labelTop[c] = Math.min(labelTop[c] === undefined ? Infinity : labelTop[c], r.y);
        lastLabel[L.series + ":" + c] = lw;
      });
      let prevMid = null;   // labels keep the same top-to-bottom order as their points
      here.forEach((L, k) => {
        if (L.pos !== "auto") return;
        const lw = textW(L.lbl, fs) - 2;
        const lowest = k === here.length - 1 && here.length > 1;
        const up = { x: L.p.x - lw / 2, y: L.p.y - gapA - lh };
        const down = { x: L.p.x - lw / 2, y: L.p.y + gapA };
        const side = lw / 2 + ms / 2 + fs * 0.15;
        const cands = (lowest ? [down, up] : [up, down]).concat([
          { x: L.p.x + side - lw / 2, y: up.y + lh * 0.35 }, { x: L.p.x - side - lw / 2, y: up.y + lh * 0.35 },
          { x: L.p.x + side - lw / 2, y: down.y - lh * 0.35 }, { x: L.p.x - side - lw / 2, y: down.y - lh * 0.35 },
        ]).map((q) => ({ x: q.x, y: q.y, w: lw, h: lh }));
        // shrink the test box a little so a line merely grazing a corner doesn't count
        const inner = (r) => ({ x: r.x + 1.5, y: r.y + lh * 0.18, w: r.w - 3, h: r.h - lh * 0.36 });
        const inOrder = (q) => prevMid === null || q.y + q.h / 2 > prevMid + lh * 0.3;
        let r = cands.find((q) => inOrder(q) && !labelHits(q) && !lineHits(inner(q)) && !dotHits(inner(q)));
        if (!r) r = cands.slice(0, 2).find((q) => inOrder(q) && !labelHits(q));   // accept a line crossing over an overlap
        if (!r) return;                                                            // nowhere sensible: leave unlabelled
        prevMid = r.y + r.h / 2;
        placed.push(r);
        P.push({ type: "text", x: r.x, y: r.y, w: r.w, h: r.h, text: L.lbl, size: fs, color: L.p.v < 0 ? NEG_RED : o.textColor, align: "center", ref: { series: L.series, c, py: L.p.y } });
        if (r.y < L.p.y) labelTop[c] = Math.min(labelTop[c] === undefined ? Infinity : labelTop[c], r.y);
        lastLabel[L.series + ":" + c] = Math.abs(r.x + r.w / 2 - L.p.x) < 1 ? lw : 0;
      });
    }

    // series names at the end of each line instead of a legend box
    if (endLegend) {
      const ends = lineIdx.map((i) => {
        const vals = spec.series[i].values;
        let c = vals.length - 1;
        while (c >= 0 && typeof vals[c] !== "number") c--;
        if (c < 0) return null;
        const x = cx(c), y = Yc(vals[c]);
        const lw = lastLabel[i + ":" + c] || 0;
        return { i, x: x + Math.max(ms / 2, lw / 2) + fs * 0.5, y, name: spec.series[i].name || "Series " + (i + 1) };
      }).filter(Boolean).sort((a, b) => a.y - b.y);
      ends.forEach((e) => {
        // beside the last dot; nudged up/down (or further right) if a line or label is in the way
        const w = textW(e.name, fs) + 2;
        const tries = [0, -0.9, 0.9, -1.8, 1.8].map((k) => ({ x: e.x, y: e.y - lh / 2 + k * lh, w, h: lh }))
          .concat([0, -0.9, 0.9].map((k) => ({ x: e.x + fs * 1.5, y: e.y - lh / 2 + k * lh, w, h: lh })));
        const inner = (r) => ({ x: r.x + 1, y: r.y + lh * 0.18, w: r.w - 2, h: r.h - lh * 0.36 });
        const r = tries.find((q) => !labelHits(q) && !lineHits(inner(q)) && !dotHits(inner(q))) ||
                  tries.find((q) => !labelHits(q)) || tries[0];
        placed.push(r);
        P.push({ type: "text", x: r.x, y: r.y, w: r.w, h: r.h, text: e.name, size: fs, color: darken(seriesColor(spec, o, e.i)), align: "left" });
      });
    }

    // ---- category labels (line charts get a thin category axis like think-cell)
    const catTop = pureLine ? plotBottom + fs * 0.9 : plotBottom + bottomPad + fs * 0.3;
    if (pureLine) P.push({ type: "line", x1: px0, y1: plotBottom + fs * 0.6, x2: pxEnd, y2: plotBottom + fs * 0.6, color: o.axisColor, w: 0.75 });

    // ---- value axis
    if (ticks.length) {
      const yb = pureLine ? plotBottom + fs * 0.6 : plotBottom;
      P.push({ type: "line", x1: px0, y1: Math.min(Y(ticks[ticks.length - 1]), Y(vmax)), x2: px0, y2: yb, color: o.axisColor, w: 0.75 });
      ticks.forEach((t) => {
        const y = Y(t), lbl = tickLabel(t), w = textW(lbl, fs);
        P.push({ type: "line", x1: px0 - fs * 0.3, y1: y, x2: px0, y2: y, color: o.axisColor, w: 0.75 });
        P.push({ type: "text", x: px0 - fs * 0.4 - w, y: y - fs * 0.65, w, h: fs * 1.3, text: lbl, size: fs, color: t < 0 ? NEG_RED : o.textColor, align: "right" });
      });
    }
    for (let c = 0; c < n; c++) {
      P.push({ type: "text", x: cx(c) - slot / 2, y: catTop, w: slot, h: catH - fs * 0.9, text: String(cats[c]), size: fs, color: o.textColor, align: "center", wrap: true, valign: "top" });
    }

    // ---- difference / CAGR arrows
    const arrowsTop = box.y + legendH;
    valid.forEach((a, k) => {
      const rowY = arrowsTop + k * arrowRowH + arrowRowH * 0.55;
      const base = (c) => (isWF ? cat[c].total : barIdx.length ? cat[c].total : num(spec.series[lineIdx[0]].values[c]));
      const va = base(a.from), vb = base(a.to);
      let text, isNeg;
      if (a.mode === "cagr") {
        const periods = Math.abs(a.to - a.from);
        const g = va > 0 && vb > 0 ? (Math.pow(vb / va, 1 / periods) - 1) * 100 : NaN;
        text = isNaN(g) ? "CAGR n/a" : "CAGR " + formatNumber(g, o, { decimals: 1, unit: "%", prefix: "", sign: true });
        isNeg = g < 0;
      } else if (a.mode === "abs") {
        text = formatNumber(vb - va, o, { sign: true }); isNeg = vb - va < 0;
      } else {
        const g = va !== 0 ? ((vb - va) / Math.abs(va)) * 100 : NaN;
        text = isNaN(g) ? "n/a" : formatNumber(g, o, { decimals: 1, unit: "%", prefix: "", sign: true });
        isNeg = g < 0;
      }
      const xa = cx(a.from), xb = cx(a.to);
      const ta = (labelTop[a.from] !== undefined && isFinite(labelTop[a.from]) ? labelTop[a.from] : plotTop) - fs * 0.25;
      const tb = (labelTop[a.to] !== undefined && isFinite(labelTop[a.to]) ? labelTop[a.to] : plotTop) - fs * 0.25;
      const triH = fs * 0.55, triW = fs * 0.7;
      const col = "#595959";
      P.push({ type: "line", x1: xa, y1: rowY, x2: xa, y2: ta, color: col, w: 0.75 });
      P.push({ type: "line", x1: Math.min(xa, xb), y1: rowY, x2: Math.max(xa, xb), y2: rowY, color: col, w: 0.75 });
      P.push({ type: "line", x1: xb, y1: rowY, x2: xb, y2: tb - triH, color: col, w: 0.75 });
      P.push({ type: "tri", x: xb - triW / 2, y: tb - triH, w: triW, h: triH, dir: "down", fill: col });
      const lw = textW(text, fs) + fs * 0.6;
      const lh = fs * 1.6;
      P.push({ type: "text", x: (xa + xb) / 2 - lw / 2, y: rowY - lh / 2, w: lw, h: lh, text, size: fs, color: isNeg ? NEG_RED : o.textColor, align: "center", fill: "#FFFFFF", stroke: col, round: true });
    });

    return { prims: P, box };
  }

  // Monotone cubic (Fritsch-Carlson) through points with increasing x; returns sampled points.
  // Smooth like Excel's smoothed line, but never overshoots above/below a data point.
  // Label text for one data point. "auto" keeps each chart type's own default (dflt).
  function contentText(spec, o, i, c, v, den, dflt, extra) {
    const sr = spec.series[i] || {};
    const lc = sr.labelContent && sr.labelContent !== "auto" ? sr.labelContent : (o.labelContent || "auto");
    if (lc === "auto") return dflt;
    const val = formatNumber(v, o, extra || {});
    const pct = den ? formatNumber((v / den) * 100, o, { unit: "%", prefix: "" }) : "";
    const nm = sr.name || "Series " + (i + 1), cn = String((spec.categories || [])[c] === undefined ? "" : spec.categories[c]);
    switch (lc) {
      case "value": return val;
      case "pct": return pct || val;
      case "valuePct": return pct ? `${val} (${pct})` : val;
      case "pctValue": return pct ? `${pct} (${val})` : val;
      case "category": return cn;
      case "series": return nm;
      case "seriesValue": return `${nm}: ${val}`;
      case "seriesPct": return `${nm}: ${pct || val}`;
      default: return dflt;
    }
  }

  // ---------------------------------------------------------------- donut / pie
  // One ring per category; the series are the segments. Angles run clockwise from 12 o'clock.
  function layoutDonut(spec, box, o) {
    const fs = o.fontSize, P = [];
    const cats = spec.categories, n = cats.length, ser = spec.series;
    const isNum = (v) => typeof v === "number" && isFinite(v);
    const hole = spec.type === "pie" ? 0 : Math.min(0.9, Math.max(0.2, isNum(o.holeSize) ? o.holeSize : 0.55));
    const names = ser.map((q, i) => q.name || "Series " + (i + 1));
    const legendTop = o.showLegend && o.legendPos === "top";
    const legendRight = o.showLegend && !legendTop;
    const legW = legendRight ? Math.max(...names.map((nm) => textW(nm, fs))) + fs * 2.2 : 0;
    const legH = legendTop ? fs * 2 : 0;
    const showCat = n > 1 || (cats[0] !== undefined && String(cats[0]).trim() !== "");
    const catH = showCat ? fs * 1.9 : 0;
    const sideRoom = fs * 3.8, vRoom = fs * 1.8;
    const cw = (box.w - legW) / n;
    const availH = box.h - legH - catH;
    const D = Math.max(24, Math.min(cw - 2 * sideRoom, availH - 2 * vRoom));
    const r1 = D / 2, r0 = r1 * hole;
    const lh = fs * 1.3;
    const pt = (cx, cy, r, a) => ({ x: cx + r * Math.sin(a), y: cy - r * Math.cos(a) });
    const posFor = (i) => { const sp = ser[i].labelPos; return sp && sp !== "auto" ? sp : (o.labelPos || "auto"); };

    for (let c = 0; c < n; c++) {
      const cx = box.x + cw * (c + 0.5), cy = box.y + legH + availH / 2;
      const vals = ser.map((q) => Math.max(0, num(q.values[c])));
      const sum = vals.reduce((a, b) => a + b, 0);
      if (sum <= 0) {
        P.push({ type: "text", x: cx - cw / 2, y: cy - lh / 2, w: cw, h: lh, text: "No data", size: fs, color: o.textColor, align: "center" });
      } else {
        let a = 0;
        const outside = [];
        vals.forEach((v, i) => {
          if (v <= 0) return;
          const a0 = a, a1 = a + (v / sum) * Math.PI * 2;
          a = a1;
          const fill = seriesColor(spec, o, i);
          P.push({ type: "wedge", cx, cy, r0, r1, a0, a1, fill, stroke: "#FFFFFF", sw: Math.min(1.25, r1 * 0.02 + 0.5) });
          const pos = posFor(i);
          if (!o.showLabels || pos === "none") return;
          const share = (v / sum) * 100;
          const lbl = contentText(spec, o, i, c, v, sum, o.pct100Labels !== false ? formatNumber(share, o, { unit: "%", prefix: "" }) : formatNumber(v, o));
          const am = (a0 + a1) / 2;
          const rm = hole > 0 ? (r0 + r1) / 2 : r1 * 0.62;
          const tw = textW(lbl, fs) - 2;
          const room = hole > 0 ? r1 - r0 : r1 * 0.7;
          const fits = (a1 - a0) * rm >= tw * 0.95 && room >= lh * 0.95;
          if (pos === "center" || (pos === "auto" && fits)) {
            const q = pt(cx, cy, rm, am);
            P.push({ type: "text", x: q.x - tw / 2, y: q.y - lh / 2, w: tw, h: lh, text: lbl, size: fs, color: onColor(fill), align: "center" });
          } else {
            outside.push({ am, lbl, tw });
          }
        });
        // outside labels with leader lines, nudged apart on each side
        [1, -1].forEach((side) => {
          const items = outside.filter((q) => (Math.sin(q.am) >= 0 ? 1 : -1) === side)
            .map((q) => Object.assign(q, { y: pt(cx, cy, r1 + fs * 0.9, q.am).y })).sort((p1, p2) => p1.y - p2.y);
          for (let k = 1; k < items.length; k++) if (items[k].y - items[k - 1].y < lh) items[k].y = items[k - 1].y + lh;
          items.forEach((q) => {
            const from = pt(cx, cy, r1 + 0.5, q.am);
            const ex = cx + side * (Math.abs(Math.sin(q.am)) * (r1 + fs * 0.6) + fs * 0.25);
            P.push({ type: "line", x1: from.x, y1: from.y, x2: ex, y2: q.y, color: o.axisColor, w: 0.5 });
            const x = side > 0 ? ex + fs * 0.25 : ex - fs * 0.25 - q.tw;
            P.push({ type: "text", x, y: q.y - lh / 2, w: q.tw, h: lh, text: q.lbl, size: fs, color: o.textColor, align: side > 0 ? "left" : "right" });
          });
        });
        if (hole > 0 && o.showTotals) {
          const t = formatNumber(sum, o), tw = textW(t, fs * 1.25);
          if (tw < r0 * 1.8) P.push({ type: "text", x: cx - r0, y: cy - fs * 0.85, w: r0 * 2, h: fs * 1.7, text: t, size: fs * 1.25, color: o.textColor, align: "center" });
        }
      }
      if (showCat) P.push({ type: "text", x: cx - cw / 2, y: cy + r1 + vRoom * 0.75, w: cw, h: lh, text: String(cats[c]), size: fs, color: o.textColor, align: "center" });
    }

    // legend: list on the right, or a row on top
    if (legendRight) {
      const x0 = box.x + box.w - legW + fs * 0.6;
      let y = box.y + legH + availH / 2 - (ser.length * fs * 1.5) / 2;
      ser.forEach((q, i) => {
        P.push({ type: "rect", x: x0, y: y + fs * 0.2, w: fs * 0.9, h: fs * 0.9, fill: seriesColor(spec, o, i) });
        P.push({ type: "text", x: x0 + fs * 1.3, y: y - fs * 0.1, w: textW(names[i], fs), h: fs * 1.3, text: names[i], size: fs, color: o.textColor, align: "left" });
        y += fs * 1.5;
      });
    } else if (legendTop) {
      let lx = box.x;
      ser.forEach((q, i) => {
        P.push({ type: "rect", x: lx + fs * 0.2, y: box.y + fs * 0.4, w: fs, h: fs, fill: seriesColor(spec, o, i) });
        const tw = textW(names[i], fs);
        P.push({ type: "text", x: lx + fs * 1.6, y: box.y + fs * 0.1, w: tw, h: fs * 1.6, text: names[i], size: fs, color: o.textColor, align: "left" });
        lx += fs * 1.6 + tw + fs;
      });
    }
    return { prims: P, box };
  }

  // SVG path for a ring segment (or a pie slice when r0 = 0)
  function wedgePath(p) {
    const f = (v) => +v.toFixed(2);
    const P = (r, a) => `${f(p.cx + r * Math.sin(a))} ${f(p.cy - r * Math.cos(a))}`;
    let a1 = p.a1;
    if (a1 - p.a0 >= Math.PI * 2 - 1e-6) {
      // a full ring can't be one arc: draw two halves
      const m = p.a0 + Math.PI;
      const outer = `M${P(p.r1, p.a0)} A${f(p.r1)} ${f(p.r1)} 0 1 1 ${P(p.r1, m)} A${f(p.r1)} ${f(p.r1)} 0 1 1 ${P(p.r1, p.a0)} Z`;
      return p.r0 > 0 ? outer + ` M${P(p.r0, p.a0)} A${f(p.r0)} ${f(p.r0)} 0 1 0 ${P(p.r0, m)} A${f(p.r0)} ${f(p.r0)} 0 1 0 ${P(p.r0, p.a0)} Z` : outer;
    }
    const large = a1 - p.a0 > Math.PI ? 1 : 0;
    if (p.r0 > 0) {
      return `M${P(p.r1, p.a0)} A${f(p.r1)} ${f(p.r1)} 0 ${large} 1 ${P(p.r1, a1)} L${P(p.r0, a1)} A${f(p.r0)} ${f(p.r0)} 0 ${large} 0 ${P(p.r0, p.a0)} Z`;
    }
    return `M${f(p.cx)} ${f(p.cy)} L${P(p.r1, p.a0)} A${f(p.r1)} ${f(p.r1)} 0 ${large} 1 ${P(p.r1, a1)} Z`;
  }

  // One SVG element for a vector primitive (smooth curve or ring segment)
  function vectorSVG(p) {
    if (p.type === "wedge") {
      return `<path d="${wedgePath(p)}" fill="${p.fill}" fill-rule="evenodd" stroke="${p.stroke || "none"}" stroke-width="${p.sw || 0}" stroke-linejoin="round"/>`;
    }
    return `<path d="${p.d}" fill="none" stroke="${p.color}" stroke-width="${p.w}"${dashAttrs(p.dash, p.w, "round")} stroke-linejoin="round"/>`;
  }
  // Bounding box of a vector primitive
  function vectorBox(p) {
    if (p.type === "wedge") return { x0: p.cx - p.r1, y0: p.cy - p.r1, x1: p.cx + p.r1, y1: p.cy + p.r1 };
    return p.bbox;
  }

  function monotoneTangents(pts) {
    const n = pts.length, d = [], m = new Array(n);
    for (let i = 0; i < n - 1; i++) d.push((pts[i + 1].y - pts[i].y) / (pts[i + 1].x - pts[i].x));
    m[0] = d[0]; m[n - 1] = d[n - 2];
    for (let i = 1; i < n - 1; i++) m[i] = d[i - 1] * d[i] <= 0 ? 0 : (d[i - 1] + d[i]) / 2;
    for (let i = 0; i < n - 1; i++) {
      if (d[i] === 0) { m[i] = 0; m[i + 1] = 0; continue; }
      const a = m[i] / d[i], b = m[i + 1] / d[i], h = a * a + b * b;
      if (h > 9) { const t = 3 / Math.sqrt(h); m[i] = t * a * d[i]; m[i + 1] = t * b * d[i]; }
    }
    return m;
  }

  // The same curve as smoothPath, written as exact cubic Bezier segments (SVG path data).
  // Hermite -> Bezier: control points sit a third of the way along each tangent.
  function bezierPath(pts) {
    const m = monotoneTangents(pts), f = (v) => +v.toFixed(2);
    let d = `M${f(pts[0].x)} ${f(pts[0].y)}`;
    for (let i = 0; i < pts.length - 1; i++) {
      const p0 = pts[i], p1 = pts[i + 1], h = (p1.x - p0.x) / 3;
      d += ` C${f(p0.x + h)} ${f(p0.y + m[i] * h)} ${f(p1.x - h)} ${f(p1.y - m[i + 1] * h)} ${f(p1.x)} ${f(p1.y)}`;
    }
    return d;
  }

  function smoothPath(pts, steps) {
    const n = pts.length, m = monotoneTangents(pts);
    // dense samples (~every 0.75pt along x), data points always kept
    const dense = [{ x: pts[0].x, y: pts[0].y, keep: true }];
    for (let i = 0; i < n - 1; i++) {
      const x0 = pts[i].x, x1 = pts[i + 1].x, y0 = pts[i].y, y1 = pts[i + 1].y, hx = x1 - x0;
      const k = Math.max(8, Math.ceil(hx / 0.75));
      for (let j = 1; j <= k; j++) {
        const t = j / k, t2 = t * t, t3 = t2 * t;
        const y = (2 * t3 - 3 * t2 + 1) * y0 + (t3 - 2 * t2 + t) * hx * m[i] + (-2 * t3 + 3 * t2) * y1 + (t3 - t2) * hx * m[i + 1];
        dense.push({ x: x0 + hx * t, y, keep: j === k });
      }
    }
    // keep a sample when the curve has turned more than maxTurn since the last kept point,
    // so bends get many short pieces (no visible corners) and straight stretches few
    const maxTurn = (steps || 2.5) * Math.PI / 180;
    const out = [dense[0]];
    let last = dense[0], dir = null;
    for (let i = 1; i < dense.length; i++) {
      const p = dense[i];
      const a = Math.atan2(p.y - last.y, p.x - last.x);
      if (dir === null) dir = a;
      const next = dense[i + 1];
      const turn = next ? Math.abs(Math.atan2(next.y - p.y, next.x - p.x) - dir) : 0;
      if (p.keep || turn > maxTurn || Math.hypot(p.x - last.x, p.y - last.y) > 40) {
        out.push({ x: p.x, y: p.y }); last = p; dir = null;
      }
    }
    return out;
  }

  function darken(hex) {
    const h = hex.replace("#", "");
    const k = luminance(hex) > 0.45 ? 0.55 : 1;
    const c = [0, 2, 4].map((i) => Math.round(parseInt(h.substr(i, 2), 16) * k));
    return "#" + c.map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
  }

  // ---------------------------------------------------------------- SVG
  function esc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }

  function toSVG(layout, o) {
    o = o || {};
    const b = layout.box;
    const pad = 4;
    const out = [];
    const font = o.fontName || "Segoe UI, Arial, sans-serif";
    out.push(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="${b.x - pad} ${b.y - pad} ${b.w + pad * 2} ${b.h + pad * 2}" width="${o.width || b.w + pad * 2}" font-family="${esc(font)}">`);
    out.push(`<rect x="${b.x - pad}" y="${b.y - pad}" width="${b.w + pad * 2}" height="${b.h + pad * 2}" fill="#FFFFFF"/>`);
    layout.prims.forEach((p) => {
      if (p.type === "rect") out.push(`<rect x="${p.x}" y="${p.y}" width="${p.w}" height="${p.h}" fill="${p.fill}"${p.stroke ? ` stroke="${p.stroke}" stroke-width="${p.strokeW || 0.75}"` : ""}/>`);
      else if (p.type === "curve" || p.type === "wedge") out.push(vectorSVG(p));
      else if (p.type === "line" && p.smoothPiece) { /* the curve above replaces these */ }
      else if (p.type === "line") out.push(`<line x1="${p.x1}" y1="${p.y1}" x2="${p.x2}" y2="${p.y2}" stroke="${p.color}" stroke-width="${p.w}"${p.dash ? dashAttrs(p.dash, p.w) : ""}/>`);
      else if (p.type === "tri") {
        const pts = p.dir === "down" ? `${p.x},${p.y} ${p.x + p.w},${p.y} ${p.x + p.w / 2},${p.y + p.h}` : `${p.x},${p.y + p.h} ${p.x + p.w},${p.y + p.h} ${p.x + p.w / 2},${p.y}`;
        out.push(`<polygon points="${pts}" fill="${p.fill}"/>`);
      } else if (p.type === "dot" && p.shape && p.shape !== "circle") {
        const cx2 = p.x + p.w / 2, cy2 = p.y + p.h / 2;
        const pts = p.shape === "square" ? `${p.x},${p.y} ${p.x + p.w},${p.y} ${p.x + p.w},${p.y + p.h} ${p.x},${p.y + p.h}`
          : p.shape === "diamond" ? `${cx2},${p.y} ${p.x + p.w},${cy2} ${cx2},${p.y + p.h} ${p.x},${cy2}`
          : `${cx2},${p.y} ${p.x + p.w},${p.y + p.h} ${p.x},${p.y + p.h}`;
        out.push(`<polygon points="${pts}" fill="${p.fill || "none"}"${p.stroke ? ` stroke="${p.stroke}" stroke-width="${p.strokeW || 0.75}"` : ""}/>`);
      } else if (p.type === "dot") out.push(`<ellipse cx="${p.x + p.w / 2}" cy="${p.y + p.h / 2}" rx="${p.w / 2}" ry="${p.h / 2}" fill="${p.fill || "none"}"${p.stroke ? ` stroke="${p.stroke}" stroke-width="${p.strokeW || 0.75}"` : ""}/>`);
      else if (p.type === "text") {
        if (p.fill || p.stroke) out.push(`<rect x="${p.x}" y="${p.y}" width="${p.w}" height="${p.h}" rx="${p.round ? p.h / 2 : 0}" fill="${p.fill || "none"}"${p.stroke ? ` stroke="${p.stroke}" stroke-width="0.75"` : ""}/>`);
        const anchor = p.align === "left" ? "start" : p.align === "right" ? "end" : "middle";
        const tx = p.align === "left" ? p.x : p.align === "right" ? p.x + p.w : p.x + p.w / 2;
        const ty = p.valign === "top" ? p.y + p.size * 1.05 : p.y + p.h / 2 + p.size * 0.36;
        out.push(`<text x="${tx}" y="${ty}" font-size="${p.size}" fill="${p.color}" text-anchor="${anchor}"${p.bold ? ' font-weight="bold"' : ""}>${esc(p.text)}</text>`);
      }
    });
    out.push("</svg>");
    return out.join("");
  }

  const api = { layoutChart, toSVG, formatNumber, parseCell, DEFAULTS, DEFAULT_PALETTE, NEG_RED, mergeOptions, vectorSVG, vectorBox, DASH_PPT };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.ChartEngine = api;
})(typeof self !== "undefined" ? self : this);

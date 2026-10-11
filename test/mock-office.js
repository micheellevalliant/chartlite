// Minimal strict mock of the PowerPoint JS API surface ChartLite uses.
// Throws on unknown enum values / bad sizes so renderer mistakes surface in tests.
(function () {
  const GEOM = new Set(["Rectangle", "RoundRectangle", "Ellipse", "Triangle", "FlowChartMerge", "Diamond"]);
  const DASH = new Set(["Dash", "DashDot", "DashDotDot", "LongDash", "LongDashDot", "RoundDot", "Solid", "SquareDot", "LongDashDotDot", "SystemDash", "SystemDot", "SystemDashDot"]);
  const AUTOSIZE = new Set(["AutoSizeNone", "AutoSizeTextToFitShape", "AutoSizeShapeToFitText"]);
  const VALIGN = new Set(["Top", "Middle", "Bottom", "TopCentered", "MiddleCentered", "BottomCentered"]);
  const HALIGN = new Set(["Left", "Center", "Right", "Justify", "JustifyLow", "Distributed", "ThaiDistributed"]);
  let nextId = 100;
  const log = (window.__pptLog = []);
  function fail(msg) { log.push("ERROR " + msg); throw new Error("MockPPT: " + msg); }

  function strict(obj, allowed, checks) {
    return new Proxy(obj, {
      set(t, k, v) {
        if (!allowed.includes(k)) fail("unknown property set: " + k);
        if (checks && checks[k] && !checks[k](v)) fail("bad value for " + k + ": " + v);
        t[k] = v; return true;
      },
    });
  }

  function makeTags() {
    const m = {};
    return {
      add(k, v) { if (typeof v !== "string") fail("tag value must be string"); m[k.toUpperCase()] = v; },
      getItemOrNullObject(k) { const v = m[k.toUpperCase()]; return { value: v, isNullObject: v === undefined, load() { return this; } }; },
      _m: m,
    };
  }

  function makeShape(slide, kind, opts, text) {
    if (!opts) fail("missing options");
    ["left", "top", "width", "height"].forEach((k) => {
      if (typeof opts[k] !== "number" || !isFinite(opts[k])) fail(kind + " bad " + k + "=" + opts[k]);
      if ((k === "width" || k === "height") && opts[k] < 0) fail(kind + " negative " + k);
    });
    // Real PowerPoint treats a 0 width/height as "not given" and uses its 72pt default.
    ["width", "height"].forEach((k) => {
      if (opts[k] === 0) { log.push("ERROR " + kind + " zero " + k + " -> PowerPoint draws it 72pt"); opts = Object.assign({}, opts, { [k]: 72 }); }
    });
    const font = strict({}, ["size", "color", "bold", "name"]);
    const para = strict({}, ["horizontalAlignment"], { horizontalAlignment: (v) => HALIGN.has(v) });
    const sh = { left: opts.left, top: opts.top, width: opts.width, height: opts.height };
    let autoNone = false;
    const grow = () => {
      if (autoNone || !textRange.text) return;
      const avail = Math.max(1, sh.width - 14.4);
      const lines = Math.max(1, Math.ceil((String(textRange.text).length * 18 * 0.55) / avail));
      sh.height = Math.max(sh.height, lines * 21.6 + 7.2);   // grows downward, top stays
    };
    const textRange = new Proxy({ text: text || "", font, paragraphFormat: para }, {
      set(t, k, v) { if (k !== "text") fail("unknown property set: " + k); t.text = v; grow(); return true; },
    });
    const textFrame = strict({ textRange }, ["autoSizeSetting", "wordWrap", "leftMargin", "rightMargin", "topMargin", "bottomMargin", "verticalAlignment"],
      { autoSizeSetting: (v) => { if (AUTOSIZE.has(v)) { autoNone = v === "AutoSizeNone"; return true; } return false; }, verticalAlignment: (v) => VALIGN.has(v) });
    const s = sh;
    Object.assign(s, {
      id: String(nextId++), kind, slide, level: 0, parentGroup: null,
      left: opts.left, top: opts.top, width: opts.width, height: opts.height,
      name: "", rotation: 0, children: null,
      tags: makeTags(),
      fillColor: null,
      fill: { setSolidColor(c) { s.fillColor = c; }, clear() { s.fillColor = null; } },
      lineFormat: strict({ visible: true, color: "#000000", weight: 0.75 }, ["visible", "color", "weight", "dashStyle"], { dashStyle: (v) => DASH.has(v) }),
      textFrame,
      load() { return this; },
      delete() { slide._remove(s); },
      getParentSlide() { return slide; },
    });
    Object.defineProperty(s, "type", { get: () => kind });
    s._grow = grow;
    s._req = { left: opts.left, top: opts.top, width: opts.width, height: opts.height };
    return s;
  }

  function makeSlide(id) {
    const slide = { id, items: [] };
    slide._remove = (s) => { slide.items = slide.items.filter((x) => x !== s); };
    const find = (id) => slide.items.find((x) => x.id === id);
    slide.shapes = {
      addGeometricShape(t, o) { if (!GEOM.has(t)) fail("geom type " + t); const s = makeShape(slide, "Geom:" + t, o); slide.items.push(s); return s; },
      addLine(t, o) { if (t !== "Straight") fail("line type " + t); const s = makeShape(slide, "Line", o); slide.items.push(s); return s; },
      addTextBox(txt, o) { const s = makeShape(slide, "TextBox", o, txt); s._grow(); slide.items.push(s); return s; },
      addGroup(ids) {
        const kids = ids.map((id) => { const k = find(id); if (!k) fail("group: missing id " + id); return k; });
        const l = Math.min(...kids.map((k) => k.left)), t = Math.min(...kids.map((k) => k.top));
        const r = Math.max(...kids.map((k) => k.left + k.width)), b = Math.max(...kids.map((k) => k.top + k.height));
        const g = makeShape(slide, "Group", { left: l, top: t, width: r - l, height: b - t });
        g.children = kids;
        kids.forEach((k) => { k.level = 1; k.parentGroup = g; });
        slide.items = slide.items.filter((x) => !kids.includes(x));
        slide.items.push(g);
        return g;
      },
      getItemOrNullObject(id) { return find(id) || { isNullObject: true, load() { return this; } }; },
      getItem(id) { return find(id) || fail("no shape " + id); },
      load() { return this; },
      get items() { return slide.items.slice(); },
    };
    slide.load = () => slide;
    slide.setSelectedShapes = (ids) => { window.__sel = ids.map((id) => find(id)); };
    return slide;
  }

  const slide = makeSlide("256#");
  window.__slide = slide;
  window.__sel = [];

  const presentation = {
    getSelectedShapes() { const items = window.__sel.filter(Boolean); return { items, load() { return this; }, getCount() { return { value: items.length }; } }; },
    getSelectedSlides() { return { getItemAt() { return slide; } }; },
    setSelectedSlides() {},
    slides: { getItem: (id) => (id === slide.id ? slide : fail("slide")), getItemOrNullObject: (id) => (id === slide.id ? slide : { isNullObject: true }) },
  };

  const verNum = (v) => { const [a, b] = String(v).split('.'); return +a * 1000 + +(b || 0); };
  let selHandler = null;
  window.Office = {
    HostType: { PowerPoint: "PowerPoint" },
    CoercionType: { XmlSvg: "xmlSvg" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { DocumentSelectionChanged: "documentSelectionChanged" },
    context: {
      requirements: { isSetSupported: (n, v) => (n === "PowerPointApi" && verNum(v) <= verNum(String(window.__apiLevel || "1.10"))) || (n === "ImageCoercion" && !window.__noSvg) },
      document: {
        addHandlerAsync(e, h) { selHandler = h; },
        // PowerPoint inserts an SVG as a "Graphic" shape on the active slide
        setSelectedDataAsync(data, opts, cb) {
          if (opts.coercionType !== "xmlSvg") fail("unexpected coercion " + opts.coercionType);
          if (!/^<svg[\s\S]*<\/svg>$/.test(data)) fail("not an svg document");
          const g = makeShape(slide, "Graphic", { left: opts.imageLeft || 0, top: opts.imageTop || 0, width: 100, height: 100 });
          g.svg = data;
          slide.items.push(g);
          window.__svgInserts = (window.__svgInserts || 0) + 1;
          setTimeout(() => cb({ status: "succeeded" }), 0);
        },
      },
    },
    onReady(cb) { setTimeout(() => cb({ host: "PowerPoint" }), 0); },
  };
  window.__fireSelection = () => selHandler && selHandler();
  window.PowerPoint = { run: async (fn) => fn({ presentation, sync: async () => {} }) };

  // Render what is "on the slide" as SVG, honoring rotation, for visual checks.
  window.__slideSVG = function () {
    const out = ['<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 540" width="960"><rect width="960" height="540" fill="#fff"/>'];
    const draw = (s) => {
      if (s.children) return s.children.forEach(draw);
      if (s.kind === "Graphic") {
        const vb = (s.svg.match(/viewBox="([^"]+)"/) || [])[1];
        const inner = s.svg.replace(/^<svg[^>]*>/, "").replace(/<\/svg>$/, "");
        out.push(`<svg x="${s.left}" y="${s.top}" width="${s.width}" height="${s.height}" viewBox="${vb}" preserveAspectRatio="none" overflow="visible">${inner}</svg>`);
        return;
      }
      const cx = s.left + s.width / 2, cy = s.top + s.height / 2;
      const tf = s.rotation ? ` transform="rotate(${s.rotation} ${cx} ${cy})"` : "";
      const lf = s.lineFormat;
      const stroke = lf.visible === false ? 'stroke="none"' : `stroke="${lf.color}" stroke-width="${lf.weight}"`;
      if (s.kind === "Line") {
        out.push(`<line stroke-linecap="butt" x1="${s.left}" y1="${s.top}" x2="${s.left + s.width}" y2="${s.top + s.height}" ${stroke}${lf.dashStyle ? ' stroke-dasharray="3 2"' : ""}${tf}/>`);
      } else if (s.kind === "Geom:Ellipse") {
        out.push(`<ellipse cx="${cx}" cy="${cy}" rx="${s.width / 2}" ry="${s.height / 2}" fill="${s.fillColor || "none"}" ${stroke}/>`);
      } else if (s.kind === "Geom:FlowChartMerge" || s.kind === "Geom:Triangle") {
        const d = s.kind === "Geom:FlowChartMerge";
        const pts = d ? `${s.left},${s.top} ${s.left + s.width},${s.top} ${cx},${s.top + s.height}` : `${s.left},${s.top + s.height} ${s.left + s.width},${s.top + s.height} ${cx},${s.top}`;
        out.push(`<polygon points="${pts}" fill="${s.fillColor}"/>`);
      } else if (s.kind !== "TextBox") {
        out.push(`<rect x="${s.left}" y="${s.top}" width="${s.width}" height="${s.height}" rx="${s.kind === "Geom:RoundRectangle" ? s.height / 2 : 0}" fill="${s.fillColor || "none"}" ${stroke}/>`);
      }
      const tr = s.textFrame.textRange;
      if (tr.text) {
        const f = tr.font, a = tr.paragraphFormat.horizontalAlignment;
        const x = a === "Left" ? s.left : a === "Right" ? s.left + s.width : cx;
        const y = s.textFrame.verticalAlignment === "Top" ? s.top + f.size : cy + f.size * 0.36;
        out.push(`<text x="${x}" y="${y}" font-size="${f.size}" fill="${f.color}" font-family="Arial" text-anchor="${a === "Left" ? "start" : a === "Right" ? "end" : "middle"}"${f.bold ? ' font-weight="bold"' : ""}>${String(tr.text).replace(/&/g, "&amp;").replace(/</g, "&lt;")}</text>`);
      }
    };
    slide.items.forEach(draw);
    out.push("</svg>");
    return out.join("");
  };
})();

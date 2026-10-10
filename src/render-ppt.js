/*
 * ChartLite PowerPoint renderer — draws engine primitives as native shapes,
 * groups them into one object and stores the chart spec on the group (shape tags),
 * so the chart can be selected later, edited in the pane and redrawn in place.
 * Requires PowerPointApi 1.8 (groups). Diagonal lines use rotation (1.10) when available.
 */
(function (root) {
  "use strict";
  const TAG_SPEC = "CHARTLITE_SPEC";
  const TAG_BOX = "CHARTLITE_BOX";       // box the chart was laid out in
  const TAG_DRAWN = "CHARTLITE_DRAWN";   // group bounds right after drawing
  const DEFAULT_BOX = { x: 60, y: 120, w: 600, h: 330 };

  const supports = (v) => {
    try { return Office.context.requirements.isSetSupported("PowerPointApi", v); } catch (e) { return false; }
  };

  // PowerPoint treats a 0 width/height as "not given" and draws a 72pt default instead,
  // so perfectly flat or vertical lines get a tiny non-zero size.
  const EPS = 0.01;
  const nz = (v) => Math.max(EPS, v);

  function addSegment(shapes, p, canRotate) {
    const dx = p.x2 - p.x1, dy = p.y2 - p.y1;
    let s;
    if (Math.abs(dx) < 0.01 || Math.abs(dy) < 0.01 || dx * dy > 0) {
      // horizontal, vertical or "falling" line: start at top-left, positive size
      s = shapes.addLine("Straight", {
        left: Math.min(p.x1, p.x2), top: Math.min(p.y1, p.y2),
        width: nz(Math.abs(dx)), height: nz(Math.abs(dy)),
      });
    } else if (canRotate) {
      // "rising" line: draw it horizontally through its midpoint, then rotate into place
      const L = Math.hypot(dx, dy);
      const mx = (p.x1 + p.x2) / 2, my = (p.y1 + p.y2) / 2;
      s = shapes.addLine("Straight", { left: mx - L / 2, top: my - EPS / 2, width: nz(L), height: EPS });
      let ang = (Math.atan2(dy, dx) * 180) / Math.PI;   // screen coords: negative = counter-clockwise
      if (ang > 90) ang -= 180; else if (ang < -90) ang += 180;
      s.rotation = ang;
    } else {
      // Office builds older than 2601 can't rotate shapes: draw the rising segment as a dotted trail
      const L = Math.hypot(dx, dy), step = 3.5, d = Math.max(1.2, p.w || 1);
      const out = [];
      for (let t = 0; t <= L; t += step) {
        const x = p.x1 + (dx * t) / L, y = p.y1 + (dy * t) / L;
        const dot = shapes.addGeometricShape("Ellipse", { left: x - d / 2, top: y - d / 2, width: d, height: d });
        dot.fill.setSolidColor(p.color);
        dot.lineFormat.visible = false;
        out.push(dot);
      }
      return out;
    }
    s.lineFormat.color = p.color;
    s.lineFormat.weight = p.w || 0.75;
    if (p.dash) s.lineFormat.dashStyle = p.dash;
    return s;
  }

  function styleText(s, p, fontName) {
    const tf = s.textFrame;
    tf.autoSizeSetting = "AutoSizeNone";
    tf.wordWrap = !!p.wrap;
    tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0;
    tf.verticalAlignment = p.valign === "top" ? "Top" : "Middle";
    const tr = tf.textRange;
    tr.font.size = p.size;
    tr.font.color = p.color;
    tr.font.bold = !!p.bold;
    if (fontName) tr.font.name = fontName;
    tr.paragraphFormat.horizontalAlignment = p.align === "left" ? "Left" : p.align === "right" ? "Right" : "Center";
  }

  function drawPrimitive(shapes, p, ctx) {
    const pos = (x, y, w, h) => ({ left: x, top: y, width: Math.max(0.5, w), height: Math.max(0.5, h) });
    let s;
    switch (p.type) {
      case "rect":
        s = shapes.addGeometricShape("Rectangle", pos(p.x, p.y, p.w, p.h));
        s.fill.setSolidColor(p.fill);
        if (p.stroke) { s.lineFormat.color = p.stroke; s.lineFormat.weight = p.strokeW || 0.75; }
        else s.lineFormat.visible = false;
        break;
      case "line":
        s = addSegment(shapes, p, ctx.canRotate);
        break;
      case "tri":
        s = shapes.addGeometricShape(p.dir === "down" ? "FlowChartMerge" : "Triangle", pos(p.x, p.y, p.w, p.h));
        s.fill.setSolidColor(p.fill);
        s.lineFormat.visible = false;
        break;
      case "dot":
        s = shapes.addGeometricShape("Ellipse", pos(p.x, p.y, p.w, p.h));
        if (p.fill) s.fill.setSolidColor(p.fill); else s.fill.clear();
        if (p.stroke) { s.lineFormat.color = p.stroke; s.lineFormat.weight = p.strokeW || 0.75; }
        else s.lineFormat.visible = false;
        break;
      case "text":
        if (p.fill || p.stroke) {
          s = shapes.addGeometricShape(p.round ? "RoundRectangle" : "Rectangle", pos(p.x, p.y, p.w, p.h));
          if (p.fill) s.fill.setSolidColor(p.fill); else s.fill.clear();
          if (p.stroke) { s.lineFormat.color = p.stroke; s.lineFormat.weight = 0.75; }
          else s.lineFormat.visible = false;
          s.textFrame.textRange.text = p.text;
        } else {
          s = shapes.addTextBox(p.text, pos(p.x, p.y, p.w, p.h));
        }
        styleText(s, p, ctx.fontName);
        break;
    }
    if (Array.isArray(s)) return s;
    if (s) s.name = "ChartLite " + p.type;
    return s;
  }

  // Draws spec into `shapes` inside `box`; returns {group} after syncing.
  async function drawChart(context, shapes, spec, box) {
    const layout = ChartEngine.layoutChart(spec, box);
    const ctx = { canRotate: supports("1.10"), fontName: (spec.options || {}).fontName || "" };
    const created = [].concat(...layout.prims.map((p) => drawPrimitive(shapes, p, ctx))).filter(Boolean);
    created.forEach((s) => s.load("id"));
    await context.sync();

    const group = shapes.addGroup(created.map((s) => s.id));
    group.name = "ChartLite Chart";
    group.tags.add(TAG_SPEC, JSON.stringify(spec));
    group.tags.add(TAG_BOX, JSON.stringify(box));
    group.load("id,left,top,width,height");
    await context.sync();
    group.tags.add(TAG_DRAWN, JSON.stringify({ x: group.left, y: group.top, w: group.width, h: group.height }));
    await context.sync();
    return group;
  }

  function readTag(shape, key) {
    const t = shape.tags.getItemOrNullObject(key);
    t.load("value");
    return t;
  }

  // Finds the ChartLite group for the current selection (clicking a bar inside the group also works).
  async function getSelectedChart(context) {
    const sel = context.presentation.getSelectedShapes();
    sel.load("items/id,items/type,items/level");
    await context.sync();
    if (!sel.items.length) return { none: true };
    let shape = sel.items[0];
    if (shape.level > 0) {
      shape = shape.parentGroup;
      shape.load("id,type,level");
      await context.sync();
      // nested groups are not produced by ChartLite, but climb once more just in case
      if (shape.level > 0) { shape = shape.parentGroup; shape.load("id,type"); await context.sync(); }
    }
    const tag = readTag(shape, TAG_SPEC);
    shape.load("left,top,width,height");
    await context.sync();
    if (tag.isNullObject) return { other: shape };
    const slide = shape.getParentSlide();
    slide.load("id");
    const box = readTag(shape, TAG_BOX), drawn = readTag(shape, TAG_DRAWN);
    await context.sync();
    return {
      shape, slideId: slide.id, id: shape.id,
      spec: JSON.parse(tag.value),
      box: box.isNullObject ? null : JSON.parse(box.value),
      drawn: drawn.isNullObject ? null : JSON.parse(drawn.value),
    };
  }

  // If the user moved/resized the group, map the layout box the same way.
  function adjustedBox(info) {
    const g = { x: info.shape.left, y: info.shape.top, w: info.shape.width, h: info.shape.height };
    if (!info.box || !info.drawn) return g;
    const sx = info.drawn.w ? g.w / info.drawn.w : 1;
    const sy = info.drawn.h ? g.h / info.drawn.h : 1;
    return {
      x: g.x + (info.box.x - info.drawn.x) * sx,
      y: g.y + (info.box.y - info.drawn.y) * sy,
      w: info.box.w * sx,
      h: info.box.h * sy,
    };
  }

  async function insertChart(spec) {
    return PowerPoint.run(async (context) => {
      let box = DEFAULT_BOX, note = "";
      const sel = await getSelectedChart(context).catch(() => ({ none: true }));
      if (sel.other) {
        const s = sel.other;
        box = { x: s.left, y: s.top, w: Math.max(120, s.width), h: Math.max(80, s.height) };
        note = " in the selected shape's area";
      }
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const group = await drawChart(context, slide.shapes, spec, box);
      slide.load("id");
      await context.sync();
      selectShape(context, slide.id, group.id);
      await context.sync();
      return { id: group.id, slideId: slide.id, note };
    });
  }

  function selectShape(context, slideId, shapeId) {
    try {
      context.presentation.setSelectedSlides([slideId]);
      context.presentation.slides.getItem(slideId).setSelectedShapes([shapeId]);
    } catch (e) { /* selection is a convenience only */ }
  }

  // Redraws an existing chart (by ids) with a new spec, keeping its position/size.
  async function updateChart(target, spec) {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemOrNullObject(target.slideId);
      await context.sync();
      if (slide.isNullObject) throw new Error("The chart's slide no longer exists.");
      const shape = slide.shapes.getItemOrNullObject(target.id);
      shape.load("left,top,width,height");
      await context.sync();
      if (shape.isNullObject) throw new Error("The chart was deleted. Use Insert to create a new one.");
      const box = readTag(shape, TAG_BOX), drawn = readTag(shape, TAG_DRAWN);
      await context.sync();
      const info = {
        shape,
        box: box.isNullObject ? null : JSON.parse(box.value),
        drawn: drawn.isNullObject ? null : JSON.parse(drawn.value),
      };
      const newBox = adjustedBox(info);
      shape.delete();
      await context.sync();
      const group = await drawChart(context, slide.shapes, spec, newBox);
      if (target.select !== false) { selectShape(context, target.slideId, group.id); await context.sync(); }
      return { id: group.id, slideId: target.slideId };
    });
  }

  async function readSelected() {
    return PowerPoint.run(async (context) => {
      const r = await getSelectedChart(context);
      if (!r.spec) return null;
      return { id: r.id, slideId: r.slideId, spec: r.spec };
    });
  }

  root.ChartPPT = { insertChart, updateChart, readSelected, supports, TAG_SPEC };
})(typeof self !== "undefined" ? self : this);

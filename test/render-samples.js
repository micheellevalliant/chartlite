// Renders sample charts to SVG so the layout can be inspected without PowerPoint.
// Usage: node test/render-samples.js <outdir>
const fs = require("fs");
const path = require("path");
const E = require("../src/engine.js");

const out = process.argv[2] || "out";
fs.mkdirSync(out, { recursive: true });
const box = { x: 0, y: 0, w: 600, h: 320 };

const samples = {
  waterfall: {
    type: "waterfall",
    categories: ["Flights", "Hotel", "Food", "Transport", "Tours", "Early-bird discount", "Total"],
    series: [{ name: "Travel budget", values: [800, 650, 300, 120, 180, -150, "e"] }],
    options: { decimals: 1, unit: "" },
  },
  stacked: {
    type: "stacked",
    categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
    series: [
      { name: "New York", values: [100, 110, 105, 120, 125, 130] },
      { name: "London", values: [80, 85, 90, 95, 100, 104] },
      { name: "Paris", values: [60, 64, 62, 70, 72, 78] },
      { name: "Tokyo", values: [40, 45, 44, 50, 55, 58] },
    ],
    options: { cagr: { from: 0, to: 5 }, deltas: [{ from: 4, to: 5, mode: "pct" }] },
  },
  clustered: {
    type: "clustered",
    categories: ["Q1", "Q2", "Q3", "Q4"],
    series: [
      { name: "Target", values: [420, 440, 465, 480] },
      { name: "Actual", values: [431, 425, 470, -15] },
    ],
  },
  stacked100: {
    type: "stacked100",
    categories: ["North", "South", "East", "West", "Central"],
    series: [
      { name: "Brand A", values: [50, 45, 55, 60, 40] },
      { name: "Brand B", values: [30, 35, 25, 25, 35] },
      { name: "Brand C", values: [20, 20, 20, 15, 25] },
    ],
  },
  line: {
    type: "line",
    categories: ["W1", "W2", "W3", "W4", "W5", "W6", "W7", "W8"],
    series: [
      { name: "Product A", values: [55.0, 56.5, 55.8, 57.2, 58.0, 57.5, 59.1, 60.0] },
      { name: "Product B", values: [45.0, 43.5, 44.2, 42.8, 42.0, 42.5, 40.9, 40.0] },
    ],
    options: { decimals: 1, unit: "%" },
  },
  combo: {
    type: "combo",
    categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
    series: [
      { name: "Revenue (Bn)", values: [210, 198, 225, 231, 240, 236] },
      { name: "Ach %", kind: "line", axis: "secondary", unit: "%", values: [97.5, 92.1, 101.3, 102.0, 104.6, 99.2] },
    ],
    options: { deltas: [{ from: 0, to: 5, mode: "abs" }] },
  },
  negatives_id_locale: {
    type: "stacked",
    categories: ["Branch A", "Branch B", "Branch C", "Branch D"],
    series: [
      { name: "Growth", values: [1234.5, -856.25, 2400.1, -120] },
    ],
    options: { locale: "id", decimals: 1, negStyle: "paren" },
  },
};

// v1.3 options: side legend, value axis, label positions, marker shapes
Object.assign(samples, {
  stacked_side_axis: {
    type: "stacked", categories: ["2024", "2025", "2026"],
    series: [{ name: "Series 1", values: [4.4, 5.9, 6] }, { name: "Series 2", values: [6, 2.3, 2.4] }, { name: "Series 3", values: [12.7, 36.9, 21.6] }],
    options: { decimals: 1, showAxis: true },
  },
  pct100_side: {
    type: "stacked100", categories: ["2024", "2025", "2026"],
    series: [{ name: "Series 1", values: [4.4, 5.9, 6] }, { name: "Series 2", values: [6, 2.3, 2.4] }, { name: "Series 3", values: [12.7, 36.9, 21.6] }],
    options: { decimals: 1 },
  },
  line_axis_markers: {
    type: "line", categories: ["2024", "2025", "2026", "2027"],
    series: [{ name: "Series 1", values: [4.1, 5.3, 3.2, 3.9], labelPos: "above" }, { name: "Series 2", values: [5.1, 2.4, 2, 2.6], labelPos: "below" }],
    options: { decimals: 1, showAxis: true, yMin: 0, markerShape: "square", markerSize: 7 },
  },
  line_center_labels: {
    type: "line", categories: ["Q1", "Q2", "Q3", "Q4"],
    series: [{ name: "Plan", values: [40, 45, 52, 60] }, { name: "Actual", values: [38, 47, 50, 63] }],
    options: { labelPos: "center", markerShape: "diamond" },
  },
  clustered_center: {
    type: "clustered", categories: ["Q1", "Q2", "Q3", "Q4"],
    series: [{ name: "Target", values: [420, 440, 465, 480] }, { name: "Actual", values: [431, 425, 470, 455] }],
    options: { labelPos: "center", showAxis: true },
  },
  bar_below: {
    type: "stacked", categories: ["A", "B", "C", "D"],
    series: [{ name: "Value", values: [120, 95, 140, 110] }],
    options: { labelPos: "below" },
  },
});

for (const [name, spec] of Object.entries(samples)) {
  const L = E.layoutChart(spec, box);
  fs.writeFileSync(path.join(out, name + ".svg"), E.toSVG(L, { width: 900 }));
  console.log(name, L.prims.length, "primitives");
}

// number format checks
const o = E.mergeOptions({ options: {} });
const assert = require("assert");
assert.strictEqual(E.formatNumber(1234.5, Object.assign({}, o, { decimals: 1 })), "1,234.5");
assert.strictEqual(E.formatNumber(-1234.5, Object.assign({}, o, { decimals: 1, locale: "id" })), "-1.234,5");
assert.strictEqual(E.formatNumber(-12, Object.assign({}, o, { negStyle: "paren" })), "(12)");
assert.strictEqual(E.formatNumber(5, o, { sign: true }), "+5");
assert.strictEqual(E.parseCell("1.234,5", "id"), 1234.5);
assert.strictEqual(E.parseCell("(1,234)", "en"), -1234);
assert.strictEqual(E.parseCell("12.5%", "en"), 12.5);
assert.strictEqual(E.parseCell("e"), "e");
assert.strictEqual(E.parseCell(""), null);
console.log("format/parse checks passed");

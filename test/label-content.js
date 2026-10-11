// Label content (absolute / % / both / names) and number scale (10^0 .. 10^12).
const assert = require("assert");
const E = require("../src/engine.js");
const box = { x: 0, y: 0, w: 600, h: 320 };
const texts = (spec) => E.layoutChart(spec, box).prims.filter((p) => p.type === "text").map((p) => p.text);
const stacked = (opts, extra) => Object.assign({
  type: "stacked", categories: ["2024", "2025"],
  series: [{ name: "Series 1", values: [4, 5] }, { name: "Series 2", values: [5, 4] }, { name: "Series 3", values: [15, 22] }],
  options: opts,
}, extra || {});

// % of the column total, and the combinations
let t = texts(stacked({ labelContent: "pct" }));
assert(t.includes("17%") && t.includes("71%"), t);                 // 4/24, 22/31
t = texts(stacked({ labelContent: "valuePct" }));
assert(t.includes("22 (71%)"), t);
t = texts(stacked({ labelContent: "pctValue" }));
assert(t.includes("71% (22)"), t);
t = texts(stacked({ labelContent: "series" }));
assert(t.filter((x) => x === "Series 3").length >= 2, t);
t = texts(stacked({ labelContent: "seriesPct", barWidth: 0.9 }, {}));
assert(t.includes("Series 3: 71%"), t);
t = texts(stacked({ labelContent: "category" }));
assert(t.filter((x) => x === "2025").length >= 2, t);
// per-series override beats the chart setting
const sp = stacked({ labelContent: "value" });
sp.series[2].labelContent = "pct";
t = texts(sp);
assert(t.includes("71%") && t.includes("4"), t);

// number scale: values divided, % untouched, axis ticks scaled
t = texts({ type: "clustered", categories: ["A", "B"], series: [{ name: "Rev", values: [4370000, 5850000] }], options: { scalePow: 6, decimals: 2, unit: "M", showAxis: true } });
assert(t.includes("4.37M") && t.includes("5.85M"), t);
assert(t.some((x) => /^[0-9.]+M$/.test(x) && x !== "4.37M" && x !== "5.85M"), "axis ticks should be scaled too: " + t);
t = texts({ type: "stacked", categories: ["A", "B"], series: [{ name: "Rev", values: [2.5e12, 3e12] }], options: { scalePow: 12, decimals: 1, unit: " T", locale: "id", cagr: { from: 0, to: 1 } } });
assert(t.includes("2,5 T") && t.includes("3,0 T"), t);
assert(t.some((x) => x === "CAGR +20,0%"), "CAGR must not be scaled: " + t);
t = texts({ type: "line", categories: ["Q1", "Q2"], series: [{ name: "Margin", values: [0.215, 0.234] }], options: { scalePow: -2, decimals: 1, unit: "%" } });
assert(t.includes("21.5%") && t.includes("23.4%"), t);
// doughnut: value instead of share, scaled
t = texts({ type: "donut", categories: ["2025"], series: [{ name: "A", values: [3e9] }, { name: "B", values: [1e9] }], options: { labelContent: "valuePct", scalePow: 9, unit: "B" } });
assert(t.includes("3B (75%)") && t.includes("4B"), t);
// waterfall: % of the start bar
t = texts({ type: "waterfall", categories: ["Start", "Up", "End"], series: [{ name: "V", values: [200, 50, "e"] }], options: { labelContent: "pct" } });
assert(t.includes("25%") && t.includes("100%"), t);
console.log("label content + number scale: all checks passed");

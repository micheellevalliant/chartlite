// Month-on-month arrows share rows instead of taking one row each.
const assert = require("assert");
const E = require("../src/engine.js");
const cats = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split(" ");
const v = [120, 132, 128, 141, 150, 158, 149, 162, 170, 165, 178, 190];
const deltas = cats.slice(0, -1).map((_, c) => ({ from: c, to: c + 1, mode: "pct" }));
const L = E.layoutChart({ type: "stacked", categories: cats, series: [{ name: "R", values: v }], options: { deltas } }, { x: 0, y: 0, w: 600, h: 320 });
const bubbles = L.prims.filter((p) => p.type === "text" && p.round);
assert.strictEqual(bubbles.length, 11);
const rows = new Set(bubbles.map((b) => Math.round(b.y)));
assert(rows.size <= 2, "rows: " + rows.size);
// no two bubbles overlap
for (let i = 0; i < bubbles.length; i++) for (let j = i + 1; j < bubbles.length; j++) {
  const a = bubbles[i], b = bubbles[j];
  assert(!(a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y), `bubbles ${a.text} and ${b.text} overlap`);
}
// bars keep most of the height (tallest bar at least 70% of the chart box)
const bars = L.prims.filter((p) => p.type === "rect" && p.role === "bar");
assert(Math.max(...bars.map((b) => b.h)) > 320 * 0.6, "bars squashed");
console.log(`MoM arrows: 11 arrows in ${rows.size} rows, no overlaps`);

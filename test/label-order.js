// Line labels must keep the vertical order of their points (never look swapped).
const E = require("../src/engine.js");
const cats = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split(" ");
const data = [
  [[2.10, 2.05, 2.22, 2.15, 2.18, 2.30, 2.26, 2.28, 2.35, 2.31, 2.40, 2.52], [2.20, 2.25, 2.19, 2.33, 2.30, 2.41, 2.38, 2.31, 2.45, 2.32, null, null]],
  [[1.43, 1.42, 1.46, 1.40, 1.41, 1.46, 1.47, 1.46, 1.46, 1.49, 1.55, 1.62], [1.54, 1.53, 1.57, 1.53, 1.49, 1.53, 1.49, 1.47, 1.53, 1.50, null, null]],
];
let checked = 0, bad = [];
for (const [a, b] of data) for (const ms of [null, 4, 6, 8, 10, 12]) for (const smooth of [false, true]) for (const h of [220, 330, 420]) {
  const L = E.layoutChart({ type: "line", categories: cats, series: [{ name: "A", values: a }, { name: "B", values: b }], options: { decimals: 2, smooth, markerSize: ms } }, { x: 60, y: 120, w: 600, h });
  const byCat = {};
  L.prims.filter((p) => p.ref).forEach((p) => (byCat[p.ref.c] = byCat[p.ref.c] || []).push(p));
  for (const c in byCat) {
    const ls = byCat[c].sort((p, q) => p.ref.py - q.ref.py);
    for (let k = 1; k < ls.length; k++) {
      checked++;
      if (ls[k].ref.py - ls[k - 1].ref.py > 0.01 && ls[k].y + ls[k].h / 2 <= ls[k - 1].y + ls[k - 1].h / 2) bad.push({ c, ms, smooth, h, hi: ls[k - 1].text, lo: ls[k].text });
    }
  }
}
if (bad.length) { console.log("swapped labels:", bad.slice(0, 5)); process.exit(1); }
console.log(`label order: ${checked} label pairs checked, none swapped`);

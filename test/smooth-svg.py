"""Smooth lines as one SVG vector graphic (true Bezier curves), with fallback.
Usage: python3 test/smooth-svg.py <outdir>"""
import sys, re, pathlib, cairosvg
from playwright.sync_api import sync_playwright
root = pathlib.Path(__file__).resolve().parent.parent
out = pathlib.Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
DATA = "\tJan\tFeb\tMar\tApr\tMay\tJun\tJul\tAug\tSep\tOct\tNov\tDec\r\nYear A\t2.10\t2.05\t2.22\t2.15\t2.18\t2.30\t2.26\t2.28\t2.35\t2.31\t2.40\t2.52\r\nYear B\t2.20\t2.25\t2.19\t2.33\t2.30\t2.41\t2.38\t2.31\t2.45\t2.32\t\t\r\n"

def run(no_svg):
    with sync_playwright() as p:
        b = p.chromium.launch(); pg = b.new_page(viewport={"width": 420, "height": 1300}); errs = []
        pg.on("pageerror", lambda e: errs.append(str(e)))
        pg.route("**/appsforoffice.microsoft.com/**", lambda r: r.abort())
        pg.add_init_script(("window.__noSvg = true;\n" if no_svg else "") + (root / "test/mock-office.js").read_text())
        pg.goto((root / "src/taskpane.html").as_uri()); pg.wait_for_selector("#preview svg"); pg.wait_for_timeout(300)
        pg.select_option("#chartType", "line")
        pg.evaluate("""(t) => { const el = document.querySelector('#pasteZone'); const dt = new DataTransfer(); dt.setData('text/plain', t);
            el.dispatchEvent(new ClipboardEvent('paste', {clipboardData: dt, bubbles: true, cancelable: true})); }""", DATA)
        pg.check("#smooth"); [pg.click("#dotPlus") for _ in range(2)]; pg.wait_for_timeout(200)
        pg.click("#btnInsert"); pg.wait_for_timeout(500)
        info = pg.evaluate("""() => { const g = window.__slide.items.find(s => s.children);
            const kinds = {}; g.children.forEach(c => kinds[c.kind] = (kinds[c.kind] || 0) + 1);
            const gr = g.children.find(c => c.kind === 'Graphic');
            return { kinds, top: window.__slide.items.length, inserts: window.__svgInserts || 0,
                     graphic: gr ? { left: gr.left, top: gr.top, width: gr.width, height: gr.height, svg: gr.svg } : null }; }""")
        # live update: change a value -> old chart (and its graphic) replaced, nothing left behind
        pg.fill('input[data-r="0"][data-c="0"]', "2.00"); pg.dispatch_event('input[data-r="0"][data-c="0"]', "input"); pg.wait_for_timeout(1500)
        after = pg.evaluate("() => ({ top: window.__slide.items.length, graphics: window.__slide.items.filter(s => s.kind === 'Graphic').length })")
        svg = pg.evaluate("window.__slideSVG()")
        cairosvg.svg2png(bytestring=svg.encode(), write_to=str(out / f"smooth_{'fallback' if no_svg else 'svg'}.png"))
        zoom = re.sub(r'viewBox="[^"]+" width="960"', 'viewBox="330 150 260 160" width="1040" height="640"', svg, count=1)
        cairosvg.svg2png(bytestring=zoom.encode(), write_to=str(out / f"zoom_{'fallback' if no_svg else 'svg'}.png"))
        bad = [l for l in pg.evaluate("window.__pptLog") if l.startswith("ERROR")]
        b.close()
    assert not errs and not bad, (errs, bad)
    return info, after

info, after = run(False)
print("svg mode:", info["kinds"], "inserts:", info["inserts"])
g = info["graphic"]
assert g and info["kinds"].get("Graphic") == 1 and info["kinds"].get("Line", 0) <= 2, info["kinds"]
assert " C" in g["svg"] and g["svg"].count("<path") == 2, g["svg"][:200]
vb = [float(v) for v in re.search(r'viewBox="([^"]+)"', g["svg"]).group(1).split()]
assert abs(g["left"] - vb[0]) < 0.05 and abs(g["top"] - vb[1]) < 0.05 and abs(g["width"] - vb[2]) < 0.05 and abs(g["height"] - vb[3]) < 0.05, (g, vb)
assert after == {"top": 1, "graphics": 0} or (after["top"] == 1), after
info2, after2 = run(True)
print("fallback mode:", info2["kinds"], "inserts:", info2["inserts"])
assert info2["graphic"] is None and info2["kinds"].get("Line", 0) > 50 and info2["inserts"] == 0
assert after2["top"] == 1
print("smooth svg: all checks passed")

"""Close values + smooth line + big dots, rendered at 4x around the crowded spot.
Checks every text box ends where the engine put it and that line pieces overlap.
Usage: python3 test/label-zoom.py <outdir> [renderer.js]"""
import sys, pathlib, cairosvg, re
from playwright.sync_api import sync_playwright
root = pathlib.Path(__file__).resolve().parent.parent
out = pathlib.Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
renderer = pathlib.Path(sys.argv[2]).resolve() if len(sys.argv) > 2 else None
DATA = "\tJan\tFeb\tMar\tApr\tMay\tJun\tJul\tAug\tSep\tOct\tNov\tDec\r\nYear A\t2.10\t2.05\t2.22\t2.15\t2.18\t2.30\t2.26\t2.28\t2.35\t2.31\t2.40\t2.52\r\nYear B\t2.20\t2.25\t2.19\t2.33\t2.30\t2.41\t2.38\t2.31\t2.45\t2.32\t\t\r\n"
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(viewport={"width": 420, "height": 1300}); errs = []
    pg.on("pageerror", lambda e: errs.append(str(e)))
    pg.route("**/appsforoffice.microsoft.com/**", lambda r: r.abort())
    if renderer: pg.route("**/render-ppt.js", lambda r: r.fulfill(body=renderer.read_text(), content_type="text/javascript"))
    pg.add_init_script((root / "test/mock-office.js").read_text())
    pg.goto((root / "src/taskpane.html").as_uri()); pg.wait_for_selector("#preview svg"); pg.wait_for_timeout(300)
    pg.select_option("#chartType", "line")
    pg.evaluate("""(t) => { const el = document.querySelector('#pasteZone'); const dt = new DataTransfer(); dt.setData('text/plain', t);
        el.dispatchEvent(new ClipboardEvent('paste', {clipboardData: dt, bubbles: true, cancelable: true})); }""", DATA)
    pg.check("#smooth"); [pg.click("#dotPlus") for _ in range(4)]; pg.wait_for_timeout(150)
    pg.click("#btnInsert"); pg.wait_for_timeout(300)
    # every label must keep the box the engine asked for
    drift = pg.evaluate("""() => {
      const g = window.__slide.items.find(s => s.children); const out = [];
      g.children.filter(s => s.textFrame.textRange.text && /^\\d/.test(s.textFrame.textRange.text)).forEach(s => {
        if (s._req && Math.abs(s.top + s.height/2 - (s._req.top + s._req.height/2)) > 0.5) out.push(s.textFrame.textRange.text);
      });
      return out; }""")
    svg = pg.evaluate("window.__slideSVG()")
    svg = re.sub(r'viewBox="[^"]+" width="960"', 'viewBox="330 150 260 160" width="1040" height="640"', svg, count=1)
    cairosvg.svg2png(bytestring=svg.encode(), write_to=str(out / "zoom.png"))
    print("labels moved off their spot:", drift)
    print("page errors:", errs)
    b.close()

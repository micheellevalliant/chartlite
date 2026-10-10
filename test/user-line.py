"""Two-series line chart with gaps and close values: paste the Excel block, insert, render the slide."""
import sys, pathlib, cairosvg
from playwright.sync_api import sync_playwright
root = pathlib.Path(__file__).resolve().parent.parent
out = pathlib.Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
DATA = "\tJan\tFeb\tMar\tApr\tMay\tJun\tJul\tAug\tSep\tOct\tNov\tDec\r\nYear A\t2.10\t2.05\t2.22\t2.15\t2.18\t2.30\t2.26\t2.28\t2.35\t2.31\t2.40\t2.52\r\nYear B\t2.20\t2.25\t2.19\t2.33\t2.30\t2.41\t2.38\t2.31\t2.45\t2.37\t\t\r\n"
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(viewport={"width": 420, "height": 1000}); errs = []
    pg.on("pageerror", lambda e: errs.append(str(e)))
    pg.route("**/appsforoffice.microsoft.com/**", lambda r: r.abort())
    pg.add_init_script((root / "test/mock-office.js").read_text()); pg.goto((root / "src/taskpane.html").as_uri())
    pg.wait_for_selector("#preview svg")
    pg.select_option("#chartType", "line")
    pg.evaluate("""(t) => { const el = document.querySelector('#pasteZone'); const dt = new DataTransfer(); dt.setData('text/plain', t);
        el.dispatchEvent(new ClipboardEvent('paste', {clipboardData: dt, bubbles: true, cancelable: true})); }""", DATA)
    pg.wait_for_timeout(200); pg.click("#btnInsert"); pg.wait_for_timeout(300)
    cairosvg.svg2png(bytestring=pg.evaluate("window.__slideSVG()").encode(), write_to=str(out / "user_line.png"))
    bad = [l for l in pg.evaluate("window.__pptLog") if l.startswith("ERROR")]
    print("errors:", len(bad), bad[:3]); print("page errors:", errs)
    b.close()

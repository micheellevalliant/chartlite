"""Axis min, Fit scale, smooth lines, dot size and dot fill — through the real pane onto the mock slide.
Usage: python3 test/scale-dots.py <outdir>"""
import sys, pathlib, cairosvg
from playwright.sync_api import sync_playwright

root = pathlib.Path(__file__).resolve().parent.parent
out = pathlib.Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
DATA = "\tJan\tFeb\tMar\tApr\tMay\tJun\tJul\tAug\r\nSales\t5,120,000\t5,310,000\t5,050,000\t5,480,000\t5,620,000\t5,410,000\t5,780,000\t5,930,000\r\n"

def paste(pg, t):
    pg.evaluate("""(t) => { const el = document.querySelector('#pasteZone'); const dt = new DataTransfer(); dt.setData('text/plain', t);
        el.dispatchEvent(new ClipboardEvent('paste', {clipboardData: dt, bubbles: true, cancelable: true})); }""", t)

def snap(pg, name):
    pg.evaluate("() => { window.__slide.items = []; window.__sel = []; }")
    pg.click("#btnInsert"); pg.wait_for_timeout(300)
    cairosvg.svg2png(bytestring=pg.evaluate("window.__slideSVG()").encode(), write_to=str(out / f"{name}.png"))
    return pg.evaluate("() => JSON.parse(window.__slide.items[0].tags._m.CHARTLITE_SPEC).options")

with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(viewport={"width": 420, "height": 1400}); errs = []
    pg.on("pageerror", lambda e: errs.append(str(e)))
    pg.route("**/appsforoffice.microsoft.com/**", lambda r: r.abort())
    pg.add_init_script((root / "test/mock-office.js").read_text())
    pg.goto((root / "src/taskpane.html").as_uri())
    pg.wait_for_selector("#preview svg"); pg.wait_for_timeout(300)

    pg.select_option("#chartType", "stacked"); paste(pg, DATA); pg.wait_for_timeout(200)
    snap(pg, "1_bars_from_zero")

    pg.fill("#yMin", "4,000,000"); pg.dispatch_event("#yMin", "input"); pg.wait_for_timeout(150)
    o = snap(pg, "2_bars_axis_min_4mn")
    assert o["yMin"] == 4000000, o["yMin"]

    pg.select_option("#chartType", "line"); pg.click("#btnAutoScale"); pg.click("#btnFitScale"); pg.wait_for_timeout(150)
    print("fit scale min:", pg.input_value("#yMin"), "|", pg.inner_text("#status"))
    pg.check("#smooth")
    for _ in range(3): pg.click("#dotPlus")
    pg.select_option("#markerFill", "white"); pg.wait_for_timeout(150)
    o = snap(pg, "3_line_smooth_bigdots_blank")
    assert o["smooth"] is True and o["markerSize"] == 9 and o["markerFill"] == "white", o
    print("dot size:", pg.input_value("#markerSize"))

    pg.select_option("#markerFill", "none"); pg.uncheck("#smooth"); pg.click("#dotMinus"); pg.wait_for_timeout(150)
    o = snap(pg, "4_line_straight_nofill")
    assert o["markerFill"] == "none" and o["markerSize"] == 8 and o["smooth"] is False, o

    # chart re-selected later must restore the scale & dot settings into the panel
    pg.evaluate("() => { window.__sel = [window.__slide.items[0]]; window.__fireSelection(); }"); pg.wait_for_timeout(2000)
    assert pg.input_value("#markerSize") == "8" and pg.input_value("#markerFill") == "none" and pg.input_value("#yMin") != ""

    pg.evaluate("() => document.querySelectorAll('details').forEach(d => d.open = true)")
    pg.screenshot(path=str(out / "pane_scale.png"), full_page=True)
    bad = [l for l in pg.evaluate("window.__pptLog") if l.startswith("ERROR")]
    assert not bad, bad
    assert not errs, errs
    print("scale/smooth/dots: all checks passed")
    b.close()

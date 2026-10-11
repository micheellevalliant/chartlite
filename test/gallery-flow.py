"""Elements gallery -> chart on slide -> datasheet popup -> paste/type -> live chart update.
Usage: python3 test/gallery-flow.py <outdir>"""
import sys, json, pathlib, cairosvg
from playwright.sync_api import sync_playwright
root = pathlib.Path(__file__).resolve().parent.parent
out = pathlib.Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
mock = (root / "test/mock-office.js").read_text()
EXCEL = "\t2024\t2025\t2026\r\nSeries 1\t4,37\t5,85\t6\r\nSeries 2\t5,98\t2,25\t2,4\r\nSeries 3\t12,65\t36,9\t21,6\r\n"

def specs(pg):
    return pg.evaluate("() => window.__slide.items.filter(s => s.children).map(s => JSON.parse(s.tags._m.CHARTLITE_SPEC))")
def paste(page, sel, text):
    page.evaluate("""([sel, t]) => { const dt = new DataTransfer(); dt.setData('text/plain', t);
        document.querySelector(sel).dispatchEvent(new ClipboardEvent('paste', {clipboardData: dt, bubbles: true, cancelable: true})); }""", [sel, text])

with sync_playwright() as p:
    b = p.chromium.launch(); ctx = b.new_context(viewport={"width": 420, "height": 1400}); errs = []
    ctx.route("**/appsforoffice.microsoft.com/**", lambda r: r.abort())
    pg = ctx.new_page(); pg.on("pageerror", lambda e: errs.append("pane: " + str(e)))
    pg.add_init_script(mock)
    pg.goto((root / "src/taskpane.html").as_uri()); pg.wait_for_selector("#preview svg"); pg.wait_for_timeout(300)
    assert pg.inner_text("#ver") == "ChartLite v1.3.0"
    assert pg.locator(".tile").count() == 7

    # 1) click the Stacked tile: chart on slide + datasheet popup
    with pg.expect_popup() as pop: pg.click('.tile[data-preset="stacked"]')
    sheet = pop.value; sheet.on("pageerror", lambda e: errs.append("sheet: " + str(e)))
    sheet.wait_for_selector('td[data-r="1"][data-c="0"]'); pg.wait_for_timeout(500)
    sp = specs(pg); assert len(sp) == 1 and sp[0]["categories"] == ["2024", "2025", "2026"], sp
    assert sheet.inner_text('td[data-r="1"][data-c="0"]') == "Series 1"
    assert "Stacked" in sheet.inner_text("#title")

    # 2) paste Excel data (comma decimals) into the datasheet at A1 -> chart updates live
    sheet.click('td[data-r="0"][data-c="1"]'); sheet.keyboard.press("ArrowLeft")
    paste(sheet, "#wrap", EXCEL); pg.wait_for_timeout(1800)
    sp = specs(pg); s0 = sp[-1]
    assert s0["options"]["locale"] == "id" and s0["series"][2]["values"] == [12.65, 36.9, 21.6], s0
    assert s0["options"]["decimals"] == 2, s0["options"]["decimals"]

    # 3) type a new number in the datasheet
    sheet.click('td[data-r="1"][data-c="3"]'); sheet.keyboard.type("10"); sheet.keyboard.press("Enter"); pg.wait_for_timeout(1800)
    assert specs(pg)[-1]["series"][0]["values"][2] == 10

    # 4) panel controls: per-series label position, value axis, bullet shape
    pg.evaluate("() => document.querySelectorAll('details').forEach(d => d.open = true)")
    pg.select_option('select[data-lpos="0"]', "below")
    pg.select_option("#showAxis", "1"); pg.select_option("#markerShape", "square"); pg.wait_for_timeout(1600)
    s0 = specs(pg)[-1]
    assert s0["series"][0].get("labelPos") == "below" and s0["options"]["showAxis"] is True and s0["options"]["markerShape"] == "square", s0
    cairosvg.svg2png(bytestring=pg.evaluate("window.__slideSVG()").encode(), write_to=str(out / "slide_stacked.png"))

    # 5) editing in the panel refreshes the open datasheet
    pg.fill('input[data-sname="1"]', "Renamed"); pg.dispatch_event('input[data-sname="1"]', "input"); pg.wait_for_timeout(700)
    assert sheet.inner_text('td[data-r="2"][data-c="0"]') == "Renamed"

    # 6) Line tile: a second chart, the same datasheet now shows its data
    pg.click('.tile[data-preset="line"]'); pg.wait_for_timeout(900)
    sp = specs(pg); assert len(sp) == 2 and sp[-1]["type"] == "line", [s["type"] for s in sp]
    assert "Line" in sheet.inner_text("#title") and sheet.inner_text('td[data-r="2"][data-c="0"]') == "Series 2"
    sheet.screenshot(path=str(out / "sheet_line.png"))
    pg.screenshot(path=str(out / "pane.png"), full_page=True)

    bad = [l for l in pg.evaluate("window.__pptLog") if l.startswith("ERROR")]
    assert not bad, bad
    assert not errs, errs
    print("gallery flow: all checks passed")
    b.close()

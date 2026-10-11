"""Panel: Label content + Number scale controls flow into the chart on the slide.
Usage: python3 test/content-pane.py <outdir>"""
import sys, pathlib, cairosvg
from playwright.sync_api import sync_playwright
root = pathlib.Path(__file__).resolve().parent.parent
out = pathlib.Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(viewport={"width": 440, "height": 1700}); errs = []
    pg.on("pageerror", lambda e: errs.append(str(e)))
    pg.route("**/appsforoffice.microsoft.com/**", lambda r: r.abort())
    pg.add_init_script((root / "test/mock-office.js").read_text())
    pg.goto((root / "src/taskpane.html").as_uri()); pg.wait_for_selector("#preview svg"); pg.wait_for_timeout(300)
    pg.evaluate("() => { window.open = () => null; }")
    pg.evaluate("() => document.querySelectorAll('details').forEach(d => d.open = true)")
    pg.select_option("#chartType", "stacked")
    pg.evaluate("""(t) => { const el = document.querySelector('#pasteZone'); const dt = new DataTransfer(); dt.setData('text/plain', t);
        el.dispatchEvent(new ClipboardEvent('paste', {clipboardData: dt, bubbles: true, cancelable: true})); }""",
        "\t2024\t2025\t2026\r\nSeries 1\t4000000\t5000000\t4000000\r\nSeries 2\t5000000\t4000000\t7000000\r\nSeries 3\t15000000\t22000000\t35000000\r\n")
    pg.click("#btnInsert"); pg.wait_for_timeout(400)
    pg.select_option("#scalePow", "6"); pg.wait_for_timeout(200)
    assert pg.input_value("#unit") == "M", pg.input_value("#unit")
    pg.select_option("#labelContent", "valuePct")
    pg.select_option('select[data-lc="2"]', "pct"); pg.wait_for_timeout(1600)
    spec = pg.evaluate("() => JSON.parse(window.__slide.items.find(s => s.children).tags._m.CHARTLITE_SPEC)")
    assert spec["options"]["scalePow"] == 6 and spec["options"]["labelContent"] == "valuePct" and spec["series"][2]["labelContent"] == "pct", spec
    txt = pg.evaluate("() => window.__slide.items.find(s => s.children).children.map(c => c.textFrame.textRange.text).filter(Boolean)")
    assert "71%" in txt and "4M (17%)" in txt and "31M" in txt, txt
    # switching the scale back clears the suggested suffix; a typed unit is kept
    pg.select_option("#scalePow", "0"); pg.wait_for_timeout(100); assert pg.input_value("#unit") == ""
    pg.fill("#unit", " juta"); pg.select_option("#scalePow", "6"); pg.wait_for_timeout(100); assert pg.input_value("#unit") == " juta"
    pg.select_option("#scalePow", "6"); pg.fill("#unit", "M"); pg.dispatch_event("#unit", "input"); pg.wait_for_timeout(1500)
    cairosvg.svg2png(bytestring=pg.evaluate("window.__slideSVG()").encode(), write_to=str(out / "slide_content.png"))
    pg.screenshot(path=str(out / "pane.png"), full_page=True)
    bad = [l for l in pg.evaluate("window.__pptLog") if l.startswith("ERROR")]
    assert not errs and not bad, (errs, bad)
    print("label content panel: all checks passed")
    b.close()

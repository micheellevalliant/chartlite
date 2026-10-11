"""Doughnut/pie charts (vector segments) and per-series line weight / dash, end to end on the mock.
Usage: python3 test/donut-styles.py <outdir>"""
import sys, re, pathlib, cairosvg
from playwright.sync_api import sync_playwright
root = pathlib.Path(__file__).resolve().parent.parent
out = pathlib.Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
mock = (root / "test/mock-office.js").read_text()

def page(p, extra=""):
    ctx = p.chromium.launch().new_context(viewport={"width": 440, "height": 1500})
    ctx.route("**/appsforoffice.microsoft.com/**", lambda r: r.abort())
    pg = ctx.new_page(); errs = []
    pg.on("pageerror", lambda e: errs.append(str(e)))
    pg.add_init_script(extra + mock)
    pg.goto((root / "src/taskpane.html").as_uri()); pg.wait_for_selector("#preview svg"); pg.wait_for_timeout(300)
    return pg, errs

def group(pg):
    return pg.evaluate("""() => { const g = window.__slide.items.filter(s => s.children).pop(); if (!g) return null;
        return { spec: JSON.parse(g.tags._m.CHARTLITE_SPEC),
                 kids: g.children.map(c => ({ kind: c.kind, svg: c.svg || null, dash: c.lineFormat.dashStyle || null, w: c.lineFormat.weight, color: c.lineFormat.color, text: c.textFrame.textRange.text })) }; }""")

with sync_playwright() as p:
    pg, errs = page(p)
    # 1) Doughnut tile -> one vector graphic with 4 ring segments + labels
    with pg.expect_popup() as pop: pg.click('.tile[data-preset="donut"]')
    sheet = pop.value; sheet.wait_for_selector('td[data-r="1"][data-c="0"]'); pg.wait_for_timeout(600)
    g = group(pg); svgs = [k["svg"] for k in g["kids"] if k["kind"] == "Graphic"]
    assert g["spec"]["type"] == "donut" and len(svgs) == 1 and svgs[0].count("<path") == 4 and " A" in svgs[0], g["spec"]["type"]
    assert {"40%", "25%", "20%", "15%", "100"} <= {k["text"] for k in g["kids"]}, [k["text"] for k in g["kids"] if k["text"]]
    # 2) paste two columns in the datasheet -> two doughnuts, still one graphic (8 segments)
    sheet.click('td[data-r="0"][data-c="1"]'); sheet.keyboard.press("ArrowLeft")
    sheet.evaluate("""(t) => { const dt = new DataTransfer(); dt.setData('text/plain', t);
        document.getElementById('wrap').dispatchEvent(new ClipboardEvent('paste', {clipboardData: dt, bubbles: true, cancelable: true})); }""",
        "\tNorth\tSouth\r\nBrand A\t52\t61\r\nBrand B\t30\t22\r\nBrand C\t12\t12\r\nOther\t6\t5\r\n")
    pg.wait_for_timeout(1800)
    g = group(pg); svgs = [k["svg"] for k in g["kids"] if k["kind"] == "Graphic"]
    assert g["spec"]["categories"] == ["North", "South"] and len(svgs) == 1 and svgs[0].count("<path") == 8, (g["spec"]["categories"], len(svgs))
    cairosvg.svg2png(bytestring=pg.evaluate("window.__slideSVG()").encode(), write_to=str(out / "slide_donut.png"))

    # 3) Line tile, then per-series weight and dash in Series style
    pg.click('.tile[data-preset="line"]'); pg.wait_for_timeout(800)
    pg.select_option('select[data-lw="0"]', "3"); pg.select_option('select[data-dash="0"]', "dash")
    pg.select_option('select[data-dash="1"]', "roundDot"); pg.wait_for_timeout(1600)
    g = group(pg)
    assert g["spec"]["series"][0].get("lineWeight") == 3 and g["spec"]["series"][0].get("dash") == "dash" and g["spec"]["series"][1].get("dash") == "roundDot", g["spec"]["series"]
    lines = [k for k in g["kids"] if k["kind"] == "Line" and k["w"] in (3,)]
    assert lines and all(k["dash"] == "Dash" for k in lines), lines[:3]
    assert any(k["dash"] == "RoundDot" for k in g["kids"] if k["kind"] == "Line"), "dotted series missing"
    # 4) smooth + dashed -> the curve graphic carries the dash pattern
    pg.evaluate("() => document.querySelectorAll('details').forEach(d => d.open = true)")
    pg.check("#smooth"); pg.wait_for_timeout(1600)
    g = group(pg); svgs = [k["svg"] for k in g["kids"] if k["kind"] == "Graphic"]
    assert len(svgs) == 1 and 'stroke-dasharray="12 9"' in svgs[0] and 'stroke-linecap="round"' in svgs[0], svgs[0][:400]
    cairosvg.svg2png(bytestring=pg.evaluate("window.__slideSVG()").encode(), write_to=str(out / "slide_lines.png"))
    # 5) chart-wide defaults
    pg.select_option("#lineWidth", "1"); pg.select_option("#lineDash", "longDash"); pg.wait_for_timeout(1600)
    o = group(pg)["spec"]["options"]; assert o["lineWidth"] == 1 and o["lineDash"] == "longDash", o
    pg.screenshot(path=str(out / "pane.png"), full_page=True)
    bad = [l for l in pg.evaluate("window.__pptLog") if l.startswith("ERROR")]
    assert not errs and not bad, (errs, bad)

    # 6) Office without vector graphics: doughnut gives a clear message and leaves nothing behind
    pg2, errs2 = page(p, "window.__noSvg = true;\n")
    pg2.evaluate("() => { window.open = () => null; }")
    pg2.click('.tile[data-preset="donut"]'); pg2.wait_for_timeout(600)
    assert "newer Office" in pg2.inner_text("#status"), pg2.inner_text("#status")
    assert pg2.evaluate("window.__slide.items.length") == 0
    assert not errs2, errs2
    print("donut + line styles: all checks passed")

"""End-to-end test: task pane in Chromium on top of the mock PowerPoint API.
Usage: python3 test/e2e.py <outdir>"""
import sys, json, pathlib, cairosvg
from playwright.sync_api import sync_playwright

root = pathlib.Path(__file__).resolve().parent.parent
out = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "out"); out.mkdir(parents=True, exist_ok=True)
mock = (root / "test" / "mock-office.js").read_text()

# What Excel puts on the clipboard for a range on Indonesian-locale Windows:
# blank corner, years as headers, 1.234,5 numbers, a (negative) and a #N/A
EXCEL_ID = "\t2022\t2023\t2024\t2025\r\nNew York\t1.100,5\t1.200,0\t1.300,5\t1.400,0\r\nLondon\t900,5\t950,0\t1.000,5\t1.050,0\r\nParis\t700,0\t750,5\t#N/A\t800,0\r\nTokyo\t(10,5)\t40,0\t60,5\t70,0\r\n"
EXCEL_WF = "Flights\tHotel\tFood\tTransport\tEarly-bird discount\tTotal\r\n1,250.5\t650.0\t300.0\t120.0\t-150.0\te\r\n"

def paste(page, selector, text):
    page.evaluate("""([sel, text]) => {
        const el = document.querySelector(sel); el.focus();
        const dt = new DataTransfer(); dt.setData('text/plain', text);
        el.dispatchEvent(new ClipboardEvent('paste', {clipboardData: dt, bubbles: true, cancelable: true}));
    }""", [selector, text])

def run(api_level):
    errors = []
    with sync_playwright() as p:
        b = p.chromium.launch()
        page = b.new_page(viewport={"width": 420, "height": 1100})
        page.on("pageerror", lambda e: errors.append(str(e)))
        page.route("**/appsforoffice.microsoft.com/**", lambda r: r.abort())
        page.add_init_script(f'window.__apiLevel="{api_level}";\n' + mock)
        page.goto((root / "src" / "taskpane.html").as_uri())
        page.wait_for_selector("#preview svg")

        # 1) paste an Excel block into the paste zone
        paste(page, "#pasteZone", EXCEL_ID)
        page.wait_for_timeout(200)
        cats = page.eval_on_selector_all("#sheet th input", "els => els.map(e => e.value)")
        names = page.eval_on_selector_all("#sheet [data-sname]", "els => els.map(e => e.value)")
        assert cats == ["2022", "2023", "2024", "2025"], cats
        assert names == ["New York", "London", "Paris", "Tokyo"], names
        assert page.input_value("#locale") == "id"
        assert "neg" in page.get_attribute('[data-r="3"][data-c="0"]', "class")
        status = page.inner_text("#status"); print("  status:", status)

        # 2) insert into the (mock) slide
        page.click("#btnInsert"); page.wait_for_timeout(300)
        slide = page.evaluate("""() => window.__slide.items.map(s => ({kind: s.kind, n: s.children ? s.children.length : 0, tag: !!s.tags._m.CHARTLITE_SPEC, id: s.id}))""")
        assert len(slide) == 1 and slide[0]["kind"] == "Group" and slide[0]["tag"], slide
        first_id = slide[0]["id"]
        cairosvg.svg2png(bytestring=page.evaluate("window.__slideSVG()").encode(), write_to=str(out / f"slide_stacked_{api_level}.png"))

        # 3) user moves/resizes the chart on the slide, clicks a bar inside the group
        page.evaluate("""() => { const g = window.__slide.items[0]; g.left += 40; g.top -= 20;
                                 window.__sel = [g.children[3]]; window.__fireSelection(); }""")
        page.wait_for_timeout(2000)

        # 4) Excel-style paste straight into a datasheet cell, live update redraws in place
        paste(page, '[data-r="0"][data-c="2"]', "1.999,9\t2.100,0")
        page.wait_for_timeout(1600)
        slide2 = page.evaluate("""() => window.__slide.items.map(s => ({kind: s.kind, id: s.id, left: s.left, spec: JSON.parse(s.tags._m.CHARTLITE_SPEC || 'null')}))""")
        assert len(slide2) == 1 and slide2[0]["id"] != first_id, slide2
        assert slide2[0]["spec"]["series"][0]["values"][2:] == [1999.9, 2100.0], slide2[0]["spec"]["series"][0]["values"]

        # 5) new waterfall from an en-locale paste, then a combo; render each
        page.evaluate("() => { window.__sel = []; window.__fireSelection(); }"); page.wait_for_timeout(2000)
        page.select_option("#chartType", "waterfall")
        paste(page, "#pasteZone", EXCEL_WF); page.wait_for_timeout(200)
        assert page.input_value("#locale") == "en"
        assert page.input_value("#decimals") == "1", page.input_value("#decimals")
        page.screenshot(path=str(out / f"pane_waterfall_{api_level}.png"), full_page=True)
        page.evaluate("() => { window.__slide.items = []; }")
        page.click("#btnInsert"); page.wait_for_timeout(300)
        cairosvg.svg2png(bytestring=page.evaluate("window.__slideSVG()").encode(), write_to=str(out / f"slide_waterfall_{api_level}.png"))

        page.evaluate("() => { window.__slide.items = []; window.__sel = []; window.__fireSelection(); }"); page.wait_for_timeout(2000)
        page.select_option("#chartType", "line")
        paste(page, "#pasteZone", "\tW1\tW2\tW3\tW4\tW5\tW6\nProduct A\t55.0%\t56.5%\t55.8%\t57.2%\t58.0%\t57.5%\nProduct B\t45.0%\t43.5%\t44.2%\t42.8%\t42.0%\t42.5%\n")
        page.evaluate("() => document.querySelectorAll('details').forEach(d => d.open = true)")
        page.fill("#unit", "%"); page.fill("#decimals", "1"); page.dispatch_event("#decimals", "input")
        page.wait_for_timeout(200)
        page.click("#btnInsert"); page.wait_for_timeout(300)
        cairosvg.svg2png(bytestring=page.evaluate("window.__slideSVG()").encode(), write_to=str(out / f"slide_line_{api_level}.png"))

        log = page.evaluate("window.__pptLog")
        b.close()
    bad = [l for l in log if l.startswith("ERROR")]
    assert not errors, errors
    assert not bad, bad
    print(f"  API {api_level}: all checks passed")

for lvl in ("1.10", "1.8"):
    run(lvl)

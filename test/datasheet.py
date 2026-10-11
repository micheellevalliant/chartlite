"""Datasheet popup on its own: load, type, paste a block from Excel, undo, transpose.
Usage: python3 test/datasheet.py <outdir>"""
import sys, json, pathlib, urllib.parse
from playwright.sync_api import sync_playwright
root = pathlib.Path(__file__).resolve().parent.parent
out = pathlib.Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
payload = {"type": "load", "chartType": "stacked", "categories": ["2024", "2025", "2026"],
           "series": [{"name": "Series 1", "raw": ["10", "12", "14"]}, {"name": "Series 2", "raw": ["6", "7", "9"]}]}
url = (root / "src/datasheet.html").as_uri() + "#d=" + urllib.parse.quote(json.dumps(payload))
def last(pg):
    pg.wait_for_timeout(350)
    d = [m for m in pg.evaluate("window.__sent") if m["type"] == "data"]
    return d[-1] if d else None
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(viewport={"width": 760, "height": 460}); errs = []
    pg.on("pageerror", lambda e: errs.append(str(e)))
    pg.route("**/appsforoffice.microsoft.com/**", lambda r: r.abort())
    pg.goto(url); pg.wait_for_timeout(200)
    assert pg.evaluate("window.__sent[0].type") == "ready"
    # type over C3 (row 3 = Series 2, col C = 2025): click, type, Enter
    pg.click('td[data-r="2"][data-c="2"]'); pg.keyboard.type("-3,5"); pg.keyboard.press("Enter")
    m = last(pg); assert m["series"][1]["raw"] == ["6", "-3,5", "9"], m
    assert "neg" in pg.get_attribute('td[data-r="2"][data-c="2"]', "class")
    # add a new category by typing in E1 and E2
    pg.click('td[data-r="0"][data-c="4"]'); pg.keyboard.type("2027"); pg.keyboard.press("Enter"); pg.keyboard.type("16"); pg.keyboard.press("Enter")
    m = last(pg); assert m["categories"] == ["2024", "2025", "2026", "2027"] and m["series"][0]["raw"][-1] == "16", m
    # paste an Excel block at A1 (replaces everything it covers)
    pg.click('td[data-r="0"][data-c="1"]'); pg.keyboard.press("ArrowLeft")
    pg.evaluate("""(t) => { const dt = new DataTransfer(); dt.setData('text/plain', t);
        document.getElementById('wrap').dispatchEvent(new ClipboardEvent('paste', {clipboardData: dt, bubbles: true, cancelable: true})); }""",
        "\tQ1\tQ2\tQ3\r\nNorth\t1.204,5\t1.310,2\t1.402,8\r\nSouth\t980,1\t1.021,7\t(12,5)\r\nWest\t760,4\t801,0\t869,2\r\n")
    m = last(pg); assert m["pasted"] and m["categories"][:3] == ["Q1", "Q2", "Q3"] and len(m["series"]) == 3 and m["series"][1]["raw"][2] == "(12,5)", m
    pg.screenshot(path=str(out / "datasheet.png"))
    # undo the paste
    pg.keyboard.press("Control+z"); m = last(pg); assert m["categories"][0] == "2024", m
    # redo, then transpose
    pg.keyboard.press("Control+y"); pg.click("#btnTranspose"); m = last(pg)
    assert m["categories"][:3] == ["North", "South", "West"] and m["series"][0]["name"] == "Q1", m
    # delete a range with the keyboard
    pg.click('td[data-r="1"][data-c="1"]'); pg.keyboard.press("Shift+ArrowRight"); pg.keyboard.press("Delete")
    m = last(pg); assert m["series"][0]["raw"][:2] == ["", ""], m
    # a new load from the panel replaces the grid without echoing back
    n = len(pg.evaluate("window.__sent"))
    pg.evaluate("(p) => window.postMessage({source: 'chartlite-pane', msg: p}, '*')", payload); pg.wait_for_timeout(400)
    assert len(pg.evaluate("window.__sent")) == n and pg.inner_text('td[data-r="1"][data-c="0"]') == "Series 1"
    assert not errs, errs
    print("datasheet: all checks passed")
    b.close()

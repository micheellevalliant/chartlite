"""Panel must not offer Insert until PowerPoint's API is really loaded."""
import pathlib
from playwright.sync_api import sync_playwright
root = pathlib.Path(__file__).resolve().parent.parent
mock = (root / "test/mock-office.js").read_text()
cases = {
    "normal": "",
    "onReady never fires": "window.Office.onReady = () => {};",
    "PowerPoint API missing": "delete window.PowerPoint;",
}
with sync_playwright() as p:
    b = p.chromium.launch()
    for name, patch in cases.items():
        pg = b.new_page(); errs = []
        pg.on("pageerror", lambda e: errs.append(str(e)))
        pg.route("**/appsforoffice.microsoft.com/**", lambda r: r.abort())
        pg.add_init_script(mock + "\n" + patch)
        pg.goto((root / "src/taskpane.html").as_uri())
        pg.wait_for_timeout(500)
        early = pg.is_disabled("#btnInsert")
        pg.wait_for_timeout(13000 if name != "normal" else 300)
        print(f"{name:24} insert disabled at start={early} now={pg.is_disabled('#btnInsert')} reload shown={pg.is_visible('#btnReload')} | {pg.inner_text('#status')[:90]}")
        assert not errs, errs
        pg.close()
    b.close()

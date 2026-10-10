// Serves the add-in files from ./src over HTTPS on https://localhost:3000
// using the Office dev certificate (run "npm run certs" once first).
const https = require("https");
const fs = require("fs");
const path = require("path");
const { getHttpsServerOptions } = require("office-addin-dev-certs");

const PORT = 3000;
const ROOT = path.join(__dirname, "src");
const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".png": "image/png", ".svg": "image/svg+xml", ".json": "application/json" };

(async () => {
  const options = await getHttpsServerOptions();
  https.createServer(options, (req, res) => {
    const urlPath = decodeURIComponent(req.url.split("?")[0]);
    const file = path.normalize(path.join(ROOT, urlPath === "/" ? "taskpane.html" : urlPath));
    if (!file.startsWith(ROOT)) { res.writeHead(403); return res.end(); }
    fs.readFile(file, (err, data) => {
      if (err) { res.writeHead(404); return res.end("Not found"); }
      res.writeHead(200, { "Content-Type": TYPES[path.extname(file)] || "application/octet-stream", "Cache-Control": "no-cache" });
      res.end(data);
    });
  }).listen(PORT, () => console.log(`ChartLite running at https://localhost:${PORT}/taskpane.html  (keep this window open)`));
})().catch((e) => { console.error("Could not start: " + e.message + "\nRun: npm run certs"); process.exit(1); });

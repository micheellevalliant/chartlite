// Builds a no-Node installer for a hosted copy of ChartLite (e.g. GitHub Pages).
// Usage: node tools/make-installer.js https://yourname.github.io/chartlite
// Output: dist/ChartLite-Install/{manifest.xml, install.bat, uninstall.bat, README.txt}
const fs = require("fs");
const path = require("path");

const url = (process.argv[2] || "").replace(/\/+$/, "");
if (!/^https:\/\/[^/]+/.test(url)) { console.error("Give an https:// address"); process.exit(1); }
const origin = url.match(/^https:\/\/[^/]+/)[0];
const root = path.join(__dirname, "..");
let xml = fs.readFileSync(path.join(root, "manifest.xml"), "utf8")
  .replace("<AppDomain>https://localhost:3000</AppDomain>", `<AppDomain>${origin}</AppDomain>`)
  .replace(/https:\/\/localhost:3000/g, url);
const id = xml.match(/<Id>([^<]+)<\/Id>/)[1];

const out = path.join(root, "dist", "ChartLite-Install");
fs.mkdirSync(out, { recursive: true });
fs.writeFileSync(path.join(out, "manifest.xml"), xml);

const crlf = (s) => s.replace(/\r?\n/g, "\r\n");
// Registers the manifest as a developer add-in for the current Windows user (no admin rights needed).
fs.writeFileSync(path.join(out, "install.bat"), crlf(`@echo off
title Install ChartLite
set "DEST=%LOCALAPPDATA%\\ChartLite"
if not exist "%DEST%" mkdir "%DEST%"
copy /Y "%~dp0manifest.xml" "%DEST%\\manifest.xml" >nul
reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "${id}" /t REG_SZ /d "%DEST%\\manifest.xml" /f >nul
if errorlevel 1 (
  echo Could not register ChartLite. Your PC policy may block add-in sideloading.
  pause & exit /b 1
)
echo.
echo  ChartLite installed.
echo  Close PowerPoint completely, open it again, then: Home tab ^> ChartLite.
echo  (If the button is missing: Home ^> Add-ins ^> More Add-ins ^> Developer Add-ins ^> ChartLite.)
echo.
pause
`));
fs.writeFileSync(path.join(out, "uninstall.bat"), crlf(`@echo off
reg delete "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "${id}" /f >nul 2>nul
del "%LOCALAPPDATA%\\ChartLite\\manifest.xml" >nul 2>nul
echo ChartLite removed. Restart PowerPoint.
pause
`));
fs.writeFileSync(path.join(out, "README.txt"), crlf(`ChartLite - installer
Hosted at: ${url}/taskpane.html

1. Double-click install.bat
2. Close and reopen PowerPoint
3. Home tab > ChartLite

No Node.js or server needed. Updates published to ${url} reach you automatically.
To remove: uninstall.bat
`));
console.log("Installer written to " + out);

// Point the manifest at a permanent host instead of https://localhost:3000
// Usage: node tools/set-host.js https://yourname.github.io/chartlite
// Writes manifest-hosted.xml; upload the contents of /src to that address.
const fs = require("fs");
const path = require("path");
const url = (process.argv[2] || "").replace(/\/+$/, "");
if (!/^https:\/\/[^/]+/.test(url)) { console.error("Give an https:// address, e.g. https://yourname.github.io/chartlite"); process.exit(1); }
const origin = url.match(/^https:\/\/[^/]+/)[0];
let xml = fs.readFileSync(path.join(__dirname, "..", "manifest.xml"), "utf8");
xml = xml.replace(/<AppDomain>https:\/\/localhost:3000<\/AppDomain>/, `<AppDomain>${origin}</AppDomain>`)
         .replace(/https:\/\/localhost:3000/g, url);
fs.writeFileSync(path.join(__dirname, "..", "manifest-hosted.xml"), xml);
console.log("Wrote manifest-hosted.xml for " + url);

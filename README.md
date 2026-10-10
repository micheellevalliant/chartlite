# ChartLite — think-cell-style charts for PowerPoint

A PowerPoint add-in (Home tab → **ChartLite**) that turns a range copied from Excel into a clean, consultant-style chart made of native PowerPoint shapes — and lets you select that chart later to edit the data and redraw it in place.

**Charts:** stacked bar · clustered bar · 100% stacked · waterfall / bridge · line · combo (bars + line, optional second axis)
**Extras:** totals above stacks · CAGR arrow · difference arrows (% or absolute) · waterfall connectors and automatic axis break · negative numbers always in red · `1,234.5` or `1.234,5` number style · prefix/unit (e.g. `Rp`, `Bn`, `%`)

---

## Quick install (recommended — no Node.js, no server)

The add-in is hosted on GitHub Pages at https://micheellevalliant.github.io/chartlite/src/taskpane.html.

1. Get the `install` folder from this repo (green **Code** button → **Download ZIP**, then unzip).
2. Double-click **`install/install.bat`** — it registers ChartLite for your Windows user only (no admin rights needed).
3. Close PowerPoint completely and reopen it → **Home** tab → **ChartLite**.
   If the button doesn't show: Home → Add-ins → More Add-ins → **Developer Add-ins** → ChartLite.

Updates pushed to this repo reach you automatically the next time PowerPoint opens. To remove it, run `install/uninstall.bat`.

You need **Microsoft 365 PowerPoint, Version 2504 or newer** on Windows.

## 1. Alternative: run it locally (for development)

You need **Windows + Microsoft 365 PowerPoint, Version 2504 or newer** (File → Account → About PowerPoint). Version 2601+ gives solid diagonal lines in line charts; older builds draw them dotted.

1. Install **Node.js LTS** from https://nodejs.org (default options).
2. Unzip this folder somewhere permanent, e.g. `C:\Tools\chartlite`.
3. Double-click **`start.bat`**.
   - First run installs the tools and a local HTTPS certificate — click **Yes** on the Windows certificate prompt.
   - PowerPoint opens with ChartLite loaded. Find it on the **Home** tab, far right.

Each work session, double-click `start.bat` again (it starts the small local server ChartLite runs from) and keep its server window open. `stop.bat` unloads it.

> **If `start.bat` is blocked by IT** — use the shared-folder method instead: run `npm run serve` in this folder, put `manifest.xml` in a shared folder, then in PowerPoint go to File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs, add that folder's `\\computer\share` path, tick *Show in Menu*, restart PowerPoint, and open it from Home → Add-ins → Advanced → **Shared Folder**.

## 2. Make a chart

1. In Excel, select a range like this and **Ctrl+C**:

   |              | Jan   | Feb   | Mar   |
   |--------------|-------|-------|-------|
   | New York     | 100,0 | 110,5 | 105,2 |
   | London       | 80,4  | 85,1  | 90,3  |

   Category names across the top, series names down the left. A blank top-left cell is fine (and expected).
2. In ChartLite, click the dashed **Paste from Excel** box and press **Ctrl+V** (or just Ctrl+V anywhere in the panel).
3. Pick the chart type, check the preview, click **Insert chart**.

What paste understands automatically:
- Number style — `1.234,5` (comma-decimal Excel settings) vs `1,234.5` is detected from the data.
- Decimal places and `%` are copied from how Excel displays the cells.
- `(120)` and `−120` are negatives; blank, `-`, `#N/A`, `#DIV/0!` are treated as empty.
- Currency text such as `$1,500` or `€ 1.500` is cleaned.
- If your months run *down* the rows instead of across, click **⇄ Transpose**.

You can also paste a block directly into any datasheet cell — it fills right and down from there, like Excel. **Enter** moves down a row.

**Tip:** to place a chart in an exact spot, draw any rectangle on the slide, select it, then click Insert — the chart is laid out in that area.

## 3. Edit an existing chart

Click the chart on the slide. ChartLite loads its data into the panel. With **Live** ticked, every change redraws the chart on the slide within a second; untick it and use **Update selected** instead.

Move or resize the chart group freely — the next update re-lays it out to the new size. The chart's data is stored inside the PowerPoint file, so it stays editable after saving, emailing, or reopening the deck (the recipient sees a normal chart; they need ChartLite only to *edit* it).

Note: manual tweaks made directly to shapes inside a chart (recoloring one bar, editing a label) are replaced on the next update. Make changes in the panel instead, or ungroup the chart to finalize it by hand.

## 4. Scale & lines

- **Axis min / Axis max** — set where the value scale starts and ends. If your values sit around 5,000,000, set Axis min to 4,000,000 (type it the way your data is written) so the trend is visible instead of flat bars from zero. Leave blank for automatic.
- **Fit scale to data** — sets Axis min just below your lowest value in one click. **Auto scale** clears it again.
- **Axis 2 min / max** — appears for combo charts, for a line on the second axis.
- **Smooth lines** — on/off. The curve passes exactly through every point and never overshoots above or below your real values.
- **Dot size − / +** — bigger or smaller line markers; `0` hides them.
- **Dot fill** — *Solid*, *Blank (white)* (hollow look, line hidden inside the dot) or *Blank (no fill)* (fully transparent).

These settings are saved with the chart, so they come back when you select it again.

## 5. Chart-type notes

**Waterfall** — uses the first series. The first value is the start bar, following values are changes (+/−), and typing **`e`** in a cell creates a total bar at that point. Example — a travel budget estimate:

| Flights | Hotel | Food | Transport | Early-bird discount | Total |
|---------|-------|------|-----------|---------------------|-------|
| 800     | 650   | 300  | 120       | -150                | e     |

Increases are green, decreases red, start/total bars dark (change under *Waterfall*). When all the action is far above zero, the scale is cut and a break mark drawn on the total bars (toggle *Auto axis break*).

**Combo** — in the datasheet, set a series to *Line*. Choose *Axis 2* for things like achievement % so the line gets its own scale and sits above the bars.

**Growth arrows** — *CAGR from/to* draws a compound-growth arrow between two categories (based on the category totals). *+ Difference arrow* adds a % change or absolute difference arrow; add several and they stack.

## 6. Use it permanently without running a server (optional)

Host the contents of `src/` on any HTTPS address (GitHub Pages or an internal web server — it's static files, and no data leaves your PC: everything runs inside PowerPoint). Then:

```
node tools/set-host.js https://yourname.github.io/chartlite
```

This writes `manifest-hosted.xml`; sideload that with the shared-folder method above. No more `start.bat`.

## Files

| Path | What it is |
|------|------------|
| `src/engine.js` | Chart layout engine (data → exact shape positions). Pure JS; also drives the preview. |
| `src/render-ppt.js` | Draws the layout as native PowerPoint shapes, groups them, stores the data on the group. |
| `src/taskpane.*` | The ChartLite panel: paste box, datasheet, options, preview. |
| `manifest.xml` | Tells PowerPoint where the add-in lives and adds the ribbon button. |
| `serve.js`, `start.bat`, `stop.bat` | Local HTTPS server and one-click launchers. |
| `test/` | Sample renders and an end-to-end test against a mock PowerPoint API. |

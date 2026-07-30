# London Neonatal Cot State — Map Dashboard

Turns the daily **Cot State** Excel export into a single, self-contained interactive
map (`London Neonatal Cot State.html`) you can email or open on any Mac or Windows machine.

**🌐 Live tool:** https://drhammadkhan.github.io/london-neonatal-cot-state/ — open it, drag
in the day's spreadsheet, and download the dashboard. Everything runs in your browser; the
spreadsheet never leaves your computer, and no real cot data is stored in this repository.

The dashboard shows every unit as a teardrop pin (fill = cots available, outline = staffing
status), with a **Heat** overlay you can switch between *spare capacity* (green) and
*pressure* (red where units are full). Drop **several days at once** and it also gets a
**date slider** to time-travel the map plus **availability sparklines** (London-wide and per
unit).

There are two ways to build it. Most of the time you want the first one.

---

## 1. The Builder — no install, Mac or Windows  ⭐

**`Cot State Builder.html`** — double-click to open it in any browser, then **drag the new
`Cot State … .xlsx` onto it** (or click to choose; select **several** for a trend view). It
reads the spreadsheets on the page, shows a live preview, and gives you a **Download
dashboard** button. The download is the finished `.html` file to send.

- No Python, no Terminal, no install — works the same on Mac and Windows.
- The Builder itself needs **no internet**. (The finished dashboard needs internet only to
  draw the map tiles; its data is baked in.)
- The download is auto-named by report date, e.g. `London Neonatal Cot State — 29 July 2026.html`.

That's it. Everything below is only for changing how it works.

---

## 2. The Python generator — for scripting / automation

Same result from the command line:

```bash
python3 generate.py                       # uses ALL "Cot State*.xlsx" here → latest day + trend
python3 generate.py "Cot State - 05.08.2026.xlsx"          # one file → single day, no trend
python3 generate.py "Cot State - 05.08.2026.xlsx" "Out.html"
```

With no arguments it reads **every** `Cot State*.xlsx` in the folder, uses the latest report
date as the map's default day, and includes the others as the trend history. Pass a single
file to force a one-day dashboard.

Requires Python 3 with `openpyxl` (`pip install openpyxl`).

---

## Adding a new hospital

Every unit's map location lives in **one file: `geo.json`**. If a **new unit** appears,
the Builder and `generate.py` both place it at central London and name it in a note/warning.

To fix it permanently:

1. Add one line to **`geo.json`**, e.g. `"New Hospital": [51.5000, -0.1000],`
   (the key must match the unit's short name in column C of the sheet).
2. `generate.py` picks it up immediately — nothing else to do.
3. Rebuild the Builder so it bakes in the new location:

   ```bash
   python3 build_builder.py
   ```

## Files

| File | Purpose |
|------|---------|
| **`Cot State Builder.html`** | The drag-&-drop tool. This is what you use day to day. |
| `London Neonatal Cot State.html` | A generated dashboard — the file you send. |
| **`geo.json`** | The one place unit map locations live. Add new hospitals here. |
| `template.html` | The dashboard shell (design/layout). `/*DATA*/` is where data is injected. |
| `generate.py` | Command-line generator (reads Excel → writes dashboard). |
| `builder_src.html` | Source for the Builder (UI + in-browser extractor). |
| `build_builder.py` | Re-assembles `Cot State Builder.html` from the source files. |
| `vendor/xlsx.full.min.js` | SheetJS, inlined into the Builder so it works offline. |

## Rebuilding the Builder (after editing design or logic)

If you change `template.html` (design) or `builder_src.html` (Builder UI/logic):

```bash
python3 build_builder.py
```

This regenerates `Cot State Builder.html` with the latest template and SheetJS inlined.

## How the data is read

- Units come from the `Cotstate` sheet (rows whose level is NICU/LNU/SCBU).
- Networks (NCL/NEL/NWL/SEL/SWL) are detected from the subtotal rows.
- The report date/time is read from the sheet header.
- Marker fill = cots available (red→green), outer ring = staffing RAG status,
  size ∝ funded cots. A unit not updated that day shows hollow/grey ("stale").

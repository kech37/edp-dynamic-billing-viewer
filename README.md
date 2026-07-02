# Energy Consumption Viewer

A local web tool to visualize EDP energy consumption reports (XLSX files with a
`POMIE_*` sheet: Data / Período / POMIE (€/MWh) / Fator de Perdas (%) / Consumo (kWh),
in 15-minute intervals).

## Run it

```sh
./start.sh
```

This starts a local server and opens http://localhost:8123/app/ in your browser.
(Requires Node.js — no dependencies to install; charting/parsing libraries are bundled in `app/lib/`.)

## Usage

- **Add reports** — drag & drop XLSX files anywhere on the page (or click "＋ Add reports").
  Each month, just drop the new report in; duplicates and overlapping periods are merged
  automatically, so re-uploading a file is always safe.
- **History** — all readings are stored in your browser (IndexedDB), so the full history
  persists between sessions on this machine/browser.
- **Backup** — "Export backup" downloads all stored data as JSON; "Import backup" restores
  it (e.g. on another machine or after clearing browser data). Keeping the XLSX files in
  `reports/` is also a backup — you can always re-drop them all.
- **Filtering** — click a period chip (or a bar in the period chart) to focus the whole
  dashboard on that period. The toggle in the top right switches between **calendar
  months** and **EDP billing periods (25th → 24th)**, which match your invoices exactly.

## What it shows

- KPIs: total consumption, average per day, estimated energy cost, average price paid,
  peak hour, and **base load** (your always-on power draw, estimated from 02:00–06:00 usage)
- Consumption & cost per period across the full history
- Average day profile (kWh per hour), weekdays vs weekends
- Daily consumption trend
- **Consumption vs market price** by hour — shows whether your usage is skewed to cheap
  or expensive hours (consumption-weighted POMIE vs flat average) and the cheapest
  3-hour window of the day
- **Top consumption days** — your heaviest days with cost and deviation from average.
  Tag unusual days (e.g. EV charging) manually or with "Auto-tag spike days"
  (> 2.5× your median day); a toggle in the filter bar excludes tagged days from all
  averages so they don't distort your habit metrics. Tags are kept in backups.
- **Base load trend** — always-on power (W) per period, estimated from 02:00–06:00 usage
- **Consumption vs temperature** — enter your city and daily mean temperatures are
  fetched from the free Open-Meteo archive (needs internet, cached locally); shows a
  scatter plot and the correlation, i.e. how much heating/cooling drives your usage
- Weekday × hour heatmap of average consumption
- **Compare periods** — pick any two periods and compare totals, average/day, cost,
  average price, and overlaid day profiles

**Cost estimate** = Consumo × (1 + Fator de Perdas) × POMIE / 1000 per 15-minute interval.
This is the OMIE-indexed energy component only — it excludes grid access tariffs, fixed
charges, and taxes, so it will not match the invoice total.

## Layout

```
app/          the web app (index.html, app.js, styles.css, lib/)
reports/      your monthly XLSX reports
serve.mjs     tiny static file server used by start.sh
```

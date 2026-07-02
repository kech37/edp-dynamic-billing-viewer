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
- **Filtering** — click a month chip (or a bar in the monthly chart) to focus the whole
  dashboard on that month.

## What it shows

- KPIs: total consumption, average per day, estimated energy cost, average price paid, peak hour
- Monthly consumption & cost across the full history
- Average day profile (kWh per hour), weekdays vs weekends
- Daily consumption trend
- Weekday × hour heatmap of average consumption

**Cost estimate** = Consumo × (1 + Fator de Perdas) × POMIE / 1000 per 15-minute interval.
This is the OMIE-indexed energy component only — it excludes grid access tariffs, fixed
charges, and taxes, so it will not match the invoice total.

## Layout

```
app/          the web app (index.html, app.js, styles.css, lib/)
reports/      your monthly XLSX reports
serve.mjs     tiny static file server used by start.sh
```

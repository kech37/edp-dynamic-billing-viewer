# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A local, single-page web app that visualizes EDP (Portuguese electricity utility) energy
consumption reports. Users drag-and-drop XLSX reports (each with a `POMIE_*` sheet of
15-minute interval readings) onto the page; the app parses them client-side, stores all
readings in IndexedDB, and renders a dashboard (KPIs, charts, heatmap, comparisons).
There is no backend beyond a static file server, no build step, and no framework —
plain JS/HTML/CSS plus two vendored libraries.

## Running

```sh
./start.sh
```

Starts `serve.mjs` (a ~35-line dependency-free static file server, no `npm install`
needed) on `http://localhost:8123/app/` and opens it in the browser. Requires only Node.js.

There is no test suite, linter, or build step in this repo — verify changes by running
the app in a browser and exercising the UI directly (see the `run` skill).

## Architecture

Everything lives in three files under `app/`, loaded as plain `<script>` tags (no
modules, no bundler):

- **`app/index.html`** — static DOM shell for every dashboard section (KPI cards,
  chart canvases, tables). `app.js` fills these elements by ID; it doesn't generate
  page structure beyond individual widgets (heatmap, table rows).
- **`app/app.js`** (single file, ~1000 lines) — all logic, organized top-to-bottom in
  the sections below. `app/lib/xlsx.full.min.js` (SheetJS) and `app/lib/chart.umd.min.js`
  (Chart.js) are vendored, loaded globally as `XLSX` and `Chart`.
- **`app/styles.css`** — plain CSS, no preprocessor.

### Data flow

1. **Parsing** (`parseWorkbook` + helpers): an XLSX ArrayBuffer is read with SheetJS,
   the sheet whose name starts with `pomie` is located, then the header row is found
   by matching normalized column names (`Data`, `Período`, `POMIE`, `Perdas`, `Consumo`)
   rather than fixed positions — reports' column order/formatting varies. Each data row
   becomes a reading record: `{ ts, date, time, kwh, pomie, loss, cost }`, where `cost`
   is computed inline (see Cost formula below).
2. **Storage** (`openDB`/`dbPutMany`/`dbGetAll`/...): four IndexedDB object stores —
   `readings` (keyed by `ts`, so re-importing the same file is idempotent/merges safely),
   `files` (import metadata for the "Imported reports" table), `tags` (user-marked
   outlier days), `weather` (fetched daily mean temperatures, cached). No schema
   migrations exist beyond `DB_VERSION` bumps in `openupgradeneeded`.
3. **State** (module-level `let`s near the top of the Aggregation section): `allReadings`,
   `fileMetas`, `tagsMap`, `weatherMap` are the in-memory mirror of IndexedDB, refreshed
   wholesale by `reload()` after every mutation. `currentPeriod`, `periodMode`, `cmpSel`,
   `excludeTagged` are UI/view state (the latter two persisted to `localStorage`).
4. **Rendering**: `reload()` is the single entry point that re-fetches everything from
   IndexedDB and calls each `render*()` function in sequence. There's no incremental
   update / virtual DOM — every mutation (tag toggle, period filter click, mode switch,
   import) calls `reload()` and repaints the whole dashboard. Charts are recreated each
   time via `upsertChart()`, which destroys the previous Chart.js instance before
   creating a new one (required by Chart.js to avoid leaks/ghost charts).

### Two filtering layers, don't conflate them

- **Period filter** (`currentPeriod`, applied by `periodOnlyReadings()`): restricts to
  one calendar month or billing period, or `"all"`.
- **Tag exclusion** (`excludeTagged`, applied on top by `filteredReadings()`):
  drops user-tagged outlier days (e.g. one-off EV charging spikes) from all averages/KPIs
  so they don't distort "normal" usage stats.

Most render functions call `filteredReadings()` (both filters). A few — the top-days
table, the base-load trend — deliberately use `periodOnlyReadings()` only, because tagged
days must stay visible there (e.g. so they can be un-tagged, or shown in the trend);
read the comments in `renderTopDays`/`renderBaseTrend` before changing exclusion logic.

### Period modes

Two ways to bucket dates into periods, toggled in the UI and driving nearly every chart:
- `"month"` — calendar month (`YYYY-MM`).
- `"billing"` — EDP's actual billing cycle, 25th of one month to 24th of the next,
  keyed by the start date. This is the default users care about since it matches invoices.

`periodKeyOf(date)` / `periodLabelOf(key)` are the two functions that encode this; any
new period-bucketed feature should go through them rather than reimplementing the date math.

### Cost formula

`cost = kwh * (1 + loss) * (pomie / 1000)` per 15-minute reading — the OMIE-indexed
energy component only. It intentionally excludes grid access tariffs, fixed charges,
and taxes, so totals will not match the actual invoice. This is called out in the UI
(KPI subtitle) — preserve that framing if you touch cost calculations.

### Base load estimate

"Base load" (always-on power draw — fridge, router, standby devices) is estimated from
average consumption during 02:00–05:59 (`renderKpis`, `renderBaseTrend`), assumed to be
free of active/discretionary usage. This is a heuristic, not a measured value.

### Reports directory

`reports/*.xlsx` holds the user's real monthly EDP exports and is git-ignored (contains
personal data — name, contract, CPE) per the comment in `.gitignore`. Treat any file
found there as sensitive; don't add code that uploads or transmits its contents anywhere
other than the local IndexedDB and the Open-Meteo weather lookup (which only sends a
city name/date range, not consumption data).

### External network calls

The only outbound requests are to Open-Meteo (`fetchWeather`): geocoding a city name,
then fetching historical daily mean temperatures for the covered date range, used purely
for the consumption-vs-temperature correlation chart. Everything else is fully local/offline.

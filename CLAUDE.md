# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A local, single-page web app: an OMIE billing calculator & viewer for EDP's (Portuguese
electricity utility) dynamic indexed tariff. Users drag-and-drop EDP XLSX reports onto
the page; each has a `POMIE_*` sheet of 15-minute interval readings and an `ELE_DINAMICA`
sheet with the real invoiced amounts and tariff constants. The app parses both
client-side, stores everything in IndexedDB, reconstructs the pre-tax invoice with EDP's
price formula, and renders a dashboard (KPIs, invoice reconciliation, charts, heatmap,
comparisons).
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

1. **Parsing** (`parseWorkbook` + `parseBillingSheet` + helpers): an XLSX ArrayBuffer is
   read with SheetJS. The sheet whose name starts with `pomie` provides the readings;
   its header row is found by matching normalized column names (`Data`, `Período`,
   `POMIE`, `Perdas`, `Consumo`) rather than fixed positions — reports' column
   order/formatting varies. Each data row becomes a reading record:
   `{ ts, date, time, kwh, pomie, loss, cost }`, where `cost` is computed inline with
   the tariff constants (see Cost formula below). Beware: the POMIE sheet ends with a
   summary row ("Preço OMIE médio global") whose `Consumo` cell holds the period
   *total* kWh — it is skipped because it has no parseable date/time; any alternative
   parsing must also skip it or totals double. The sheet whose name contains `dinamica` is
   parsed by `parseBillingSheet` into a billing record: real invoiced amounts before
   taxes (energy lines + power/fixed line), billing-period dates/days, and the tariff
   constants (K1, K2, K3, TAR energia, TAR potência) — with `DEFAULT_TARIFF` as
   fallback when a report lacks the sheet.
2. **Storage** (`openDB`/`dbPutMany`/`dbGetAll`/...): five IndexedDB object stores —
   `readings` (keyed by `ts`, so re-importing the same file is idempotent/merges safely),
   `files` (import metadata for the "Imported reports" table), `tags` (user-marked
   outlier days), `weather` (fetched daily mean temperatures, cached), `billing`
   (parsed ELE_DINAMICA invoice summaries, keyed by period start date). The v3 upgrade
   clears `readings`/`files` from older versions because the cost formula changed;
   otherwise no schema migrations exist beyond `DB_VERSION` bumps.
3. **State** (module-level `let`s near the top of the Aggregation section): `allReadings`,
   `fileMetas`, `billingMetas`, `tagsMap`, `weatherMap` are the in-memory mirror of IndexedDB, refreshed
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

EDP's indexed-tariff price formula, as printed in the ELE_DINAMICA sheet:

```
invoice = Σi (POMIE_i × (1+Perdas_i) × K1 + K2 + TAR_energia) × Consumo_i
        + (K3 + TAR_potência) × nº days
```

Per 15-minute reading: `cost = kwh * ((pomie/1000) * (1+loss) * K1 + K2 + TAR_energia)`
(the energy component). The fixed daily component `(K3 + TAR_potência)` is added at
display time (`fixedDailyEur()`) in the bill-estimate KPI, monthly chart cost line,
compare view, and `renderBilling`. Constants come from the report's ELE_DINAMICA sheet
(per-file for reading costs, latest via `activeTariff()` for display), falling back to
`DEFAULT_TARIFF` (K1=1.08, K2=0.0185, TARe=0.0607, K3=0.1171, TARp=0.2291). This
reconstructs the invoice's "Fatura Total antes de Taxas e Impostos" to within cents —
verified in `renderBilling`, which shows real vs recomputed per invoice (and calc vs
billed kWh, which also match). Taxes, levies and IVA are still excluded; preserve that
framing in the UI.

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

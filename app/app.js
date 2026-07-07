/* EDP Dynamic Electricity — OMIE Billing Calculator & Viewer.
   Parses EDP XLSX reports (POMIE_* readings sheet + ELE_DINAMICA billing sheet),
   stores 15-min readings in IndexedDB, reconstructs the pre-tax invoice with
   EDP's indexed-price formula, and renders a dashboard. */

"use strict";

// ---------- IndexedDB ----------

const DB_NAME = "edp-viewer";
const DB_VERSION = 3;

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = req.result;
      if (!db.objectStoreNames.contains("readings")) {
        db.createObjectStore("readings", { keyPath: "ts" });
      }
      if (!db.objectStoreNames.contains("files")) {
        db.createObjectStore("files", { keyPath: "name" });
      }
      if (!db.objectStoreNames.contains("tags")) {
        db.createObjectStore("tags", { keyPath: "date" });
      }
      if (!db.objectStoreNames.contains("weather")) {
        db.createObjectStore("weather", { keyPath: "date" });
      }
      if (!db.objectStoreNames.contains("billing")) {
        db.createObjectStore("billing", { keyPath: "start" });
      }
      // v3 switched reading costs from the bare OMIE component to the full
      // invoice formula; readings stored by older versions carry stale costs,
      // so drop them and let the user re-import the XLSX reports.
      if (e.oldVersion > 0 && e.oldVersion < 3) {
        req.transaction.objectStore("readings").clear();
        req.transaction.objectStore("files").clear();
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function txPromise(tx) {
  return new Promise((resolve, reject) => {
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}

async function dbPutMany(storeName, items) {
  const db = await openDB();
  const tx = db.transaction(storeName, "readwrite");
  const store = tx.objectStore(storeName);
  for (const item of items) store.put(item);
  return txPromise(tx);
}

async function dbGetAll(storeName) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const req = db.transaction(storeName).objectStore(storeName).getAll();
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function dbDelete(storeName, key) {
  const db = await openDB();
  const tx = db.transaction(storeName, "readwrite");
  tx.objectStore(storeName).delete(key);
  return txPromise(tx);
}

async function dbClearAll() {
  const db = await openDB();
  const stores = ["readings", "files", "tags", "weather", "billing"];
  const tx = db.transaction(stores, "readwrite");
  for (const s of stores) tx.objectStore(s).clear();
  return txPromise(tx);
}

// ---------- Parsing ----------

function normalize(s) {
  return String(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim();
}

function parseNumber(v) {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = parseFloat(v.replace(/\s/g, "").replace(",", "."));
    return Number.isNaN(n) ? null : n;
  }
  return null;
}

// -> "YYYY-MM-DD" or null
function parseDateCell(v) {
  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m) return null;
    return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  }
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}-${String(v.getDate()).padStart(2, "0")}`;
  }
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
  }
  return null;
}

// -> "HH:MM" or null
function parseTimeCell(v) {
  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{1,2}):(\d{2})$/);
    if (!m) return null;
    return `${m[1].padStart(2, "0")}:${m[2]}`;
  }
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return `${String(v.getHours()).padStart(2, "0")}:${String(v.getMinutes()).padStart(2, "0")}`;
  }
  if (typeof v === "number" && v >= 0 && v < 1) {
    const mins = Math.round(v * 1440);
    return `${String(Math.floor(mins / 60)).padStart(2, "0")}:${String(mins % 60).padStart(2, "0")}`;
  }
  return null;
}

// "14%" -> 0.14; 0.14 -> 0.14; 14 -> 0.14
function parseLossCell(v) {
  if (typeof v === "string") {
    const n = parseNumber(v.replace("%", ""));
    return n === null ? 0 : n / 100;
  }
  if (typeof v === "number") return v <= 1 ? v : v / 100;
  return 0;
}

// EDP indexed-tariff price formula (from the ELE_DINAMICA sheet):
//   invoice = Σi (POMIE_i × (1+Perdas_i) × K1 + K2 + TAR_energia) × Consumo_i
//           + (K3 + TAR_potência) × nº days
// Defaults below match the constants published in the reports; per-file values
// parsed from ELE_DINAMICA override them.
const DEFAULT_TARIFF = { k1: 1.08, k2: 0.0185, tarE: 0.0607, k3: 0.1171, tarP: 0.2291 };

// Parse the ELE_DINAMICA billing-summary sheet: real invoiced amounts (before
// taxes) plus the tariff constants. Returns null if the sheet is absent.
function parseBillingSheet(wb, fileName) {
  const sheetName = wb.SheetNames.find((n) => normalize(n).includes("dinamica"));
  if (!sheetName) return null;
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null });

  let label = null;
  let headerIdx = -1, cols = null;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    row.forEach((cell, c) => {
      if (cell !== null && normalize(cell) === "periodo de faturacao" && row[c + 1]) label = String(row[c + 1]);
    });
    const idx = {};
    row.forEach((cell, c) => {
      if (cell === null) return;
      const t = normalize(cell);
      if (t === "data inicio") idx.start = c;
      else if (t === "data fim") idx.end = c;
      else if (t.includes("dias")) idx.days = c;
      else if (t.startsWith("consumo")) idx.kwh = c;
      else if (t.includes("erse")) idx.k1 = c; // "(1+ Perdas ERSE i) x K1"
      else if (t === "k2") idx.k2 = c;
      else if (t === "k3") idx.k3 = c;
      else if (t.includes("tarenergia")) idx.tarE = c;
      else if (t.includes("tarpotencia")) idx.tarP = c;
      else if (t.startsWith("fatura total")) idx.total = c;
    });
    if (idx.start !== undefined && idx.total !== undefined) { cols = idx; headerIdx = i; break; }
  }
  if (!cols) return null;

  const b = { file: fileName, label, start: null, end: null, days: 0, billedKwh: 0, energyEur: 0, powerEur: 0, totalEur: 0 };
  const grab = (row, key) => {
    if (cols[key] === undefined || b[key] !== undefined) return;
    const v = parseNumber(row[cols[key]]);
    if (v !== null) b[key] = v;
  };
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const total = parseNumber(row[cols.total]);
    if (total === null) continue;
    b.totalEur += total;
    const start = parseDateCell(row[cols.start]);
    const end = parseDateCell(row[cols.end]);
    if (start && (!b.start || start < b.start)) b.start = start;
    if (end && (!b.end || end > b.end)) b.end = end;
    const kwh = parseNumber(row[cols.kwh]);
    if (kwh !== null) { // energy line
      b.billedKwh += kwh;
      b.energyEur += total;
      grab(row, "k1"); grab(row, "k2"); grab(row, "tarE");
    } else { // power / fixed-charge line
      b.powerEur += total;
      const m = String(row[cols.days] ?? "").match(/(\d+)/);
      if (m) b.days = Number(m[1]);
      grab(row, "k3"); grab(row, "tarP");
    }
  }
  return b.start ? b : null;
}

function parseWorkbook(arrayBuffer, fileName) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
  const sheetName = wb.SheetNames.find((n) => normalize(n).startsWith("pomie"));
  if (!sheetName) throw new Error(`${fileName}: no POMIE_* sheet found`);

  const billing = parseBillingSheet(wb, fileName);
  const tariff = {
    k1: billing?.k1 ?? DEFAULT_TARIFF.k1,
    k2: billing?.k2 ?? DEFAULT_TARIFF.k2,
    tarE: billing?.tarE ?? DEFAULT_TARIFF.tarE,
  };

  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null });

  // Locate the header row and map column indices by header text.
  let cols = null;
  let headerIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const row = rows[i] || [];
    const idx = {};
    row.forEach((cell, c) => {
      if (cell === null) return;
      const t = normalize(cell);
      if (t === "data") idx.date = c;
      else if (t.startsWith("periodo")) idx.time = c;
      else if (t.includes("pomie")) idx.pomie = c;
      else if (t.includes("perdas")) idx.loss = c;
      else if (t.includes("consumo")) idx.kwh = c;
    });
    if (idx.date !== undefined && idx.time !== undefined && idx.kwh !== undefined) {
      cols = idx;
      headerIdx = i;
      break;
    }
  }
  if (!cols) throw new Error(`${fileName}: could not find the header row (Data / Período / Consumo)`);

  const records = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const date = parseDateCell(row[cols.date]);
    const time = parseTimeCell(row[cols.time]);
    if (!date || !time) continue; // skips blank and summary rows ("Preço OMIE médio global")
    const kwh = parseNumber(row[cols.kwh]);
    if (kwh === null) continue;
    const pomie = cols.pomie !== undefined ? parseNumber(row[cols.pomie]) : null;
    const loss = cols.loss !== undefined ? parseLossCell(row[cols.loss]) : 0;
    // Full energy component of the invoice (per kWh), before taxes.
    const cost = pomie !== null
      ? kwh * ((pomie / 1000) * (1 + loss) * tariff.k1 + tariff.k2 + tariff.tarE)
      : null;
    records.push({ ts: `${date} ${time}`, date, time, kwh, pomie, loss, cost });
  }
  if (records.length === 0) throw new Error(`${fileName}: no readings found in sheet ${sheetName}`);
  return { records, billing };
}

// ---------- State & formatting ----------

let allReadings = [];
let fileMetas = [];
let billingMetas = []; // parsed ELE_DINAMICA invoice summaries, sorted by start date
let currentPeriod = "all"; // "all" or a period key (see periodKeyOf)
let periodMode = "month"; // "month" (calendar) or "billing" (25th → 24th, matches EDP invoices)
let cmpSel = { a: null, b: null }; // compare-view period selection
let tagsMap = new Map(); // date -> tag record; tagged days can be excluded from all metrics
let weatherMap = new Map(); // date -> mean temperature (°C)
let excludeTagged = localStorage.getItem("excludeTagged") !== "0";
let weatherAutoTried = false;
let charts = {};

const fmtKwh = new Intl.NumberFormat("pt-PT", { maximumFractionDigits: 1 });
const fmtKwh2 = new Intl.NumberFormat("pt-PT", { maximumFractionDigits: 2 });
const fmtEur = new Intl.NumberFormat("pt-PT", { style: "currency", currency: "EUR" });
const fmtEur4 = new Intl.NumberFormat("pt-PT", { style: "currency", currency: "EUR", maximumFractionDigits: 4 });

const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const DOW_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

function monthLabel(ym) {
  const [y, m] = ym.split("-").map(Number);
  return `${MONTH_NAMES[m - 1]} ${y}`;
}

// Period key for a date: "YYYY-MM" in month mode, or the ISO date of the
// billing-period start (the 25th) in billing mode.
function periodKeyOf(date) {
  if (periodMode === "month") return date.slice(0, 7);
  let [y, m, d] = date.split("-").map(Number);
  if (d < 25) { m -= 1; if (m === 0) { m = 12; y -= 1; } }
  return `${y}-${String(m).padStart(2, "0")}-25`;
}

function periodLabelOf(key) {
  if (periodMode === "month") return monthLabel(key);
  const [y, m] = key.split("-").map(Number);
  let ey = y, em = m + 1;
  if (em === 13) { em = 1; ey += 1; }
  return `25 ${MONTH_NAMES[m - 1]} – 24 ${MONTH_NAMES[em - 1]} ${ey}`;
}

function dateLabel(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return `${String(d).padStart(2, "0")} ${MONTH_NAMES[m - 1]} ${y}`;
}

// Monday-first weekday index for an ISO date
function dowIndex(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return (new Date(y, m - 1, d).getDay() + 6) % 7;
}

function toast(msg, isError = false) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.className = isError ? "error" : "";
  el.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { el.hidden = true; }, 4000);
}

// ---------- Ingest ----------

async function ingestFiles(fileList) {
  const files = Array.from(fileList).filter((f) => /\.xlsx?$/i.test(f.name));
  if (files.length === 0) {
    toast("No XLSX files selected.", true);
    return;
  }
  let added = 0;
  const errors = [];
  for (const file of files) {
    try {
      const buf = await file.arrayBuffer();
      const { records, billing } = parseWorkbook(buf, file.name);
      await dbPutMany("readings", records);
      if (billing) await dbPutMany("billing", [billing]);
      const dates = records.map((r) => r.date);
      await dbPutMany("files", [{
        name: file.name,
        rows: records.length,
        from: dates.reduce((a, b) => (a < b ? a : b)),
        to: dates.reduce((a, b) => (a > b ? a : b)),
        uploadedAt: new Date().toISOString(),
      }]);
      added += records.length;
    } catch (err) {
      console.error(err);
      errors.push(err.message);
    }
  }
  await reload();
  if (errors.length) toast(errors.join(" · "), true);
  else toast(`Imported ${files.length} file(s) — ${fmtKwh.format(added)} readings merged into history.`);
}
window.ingestFiles = ingestFiles; // also used by tests

// ---------- Aggregation ----------

function isExcluded(r) {
  return excludeTagged && tagsMap.has(r.date);
}

// Period filter only (tagged days kept) — used where tagged days must stay visible.
function periodOnlyReadings() {
  if (currentPeriod === "all") return allReadings;
  return allReadings.filter((r) => periodKeyOf(r.date) === currentPeriod);
}

function filteredReadings() {
  return periodOnlyReadings().filter((r) => !isExcluded(r));
}

function groupBy(records, keyFn) {
  const map = new Map();
  for (const r of records) {
    const k = keyFn(r);
    let g = map.get(k);
    if (!g) { g = { kwh: 0, cost: 0, dates: new Set() }; map.set(k, g); }
    g.kwh += r.kwh;
    if (r.cost !== null) g.cost += r.cost;
    g.dates.add(r.date);
  }
  return map;
}

// ---------- Rendering ----------

async function reload() {
  allReadings = await dbGetAll("readings");
  allReadings.sort((a, b) => (a.ts < b.ts ? -1 : 1));
  fileMetas = await dbGetAll("files");
  fileMetas.sort((a, b) => (a.from < b.from ? -1 : 1));
  billingMetas = await dbGetAll("billing");
  billingMetas.sort((a, b) => (a.start < b.start ? -1 : 1));
  tagsMap = new Map((await dbGetAll("tags")).map((t) => [t.date, t]));
  weatherMap = new Map((await dbGetAll("weather")).map((w) => [w.date, w.tmean]));

  const periods = [...new Set(allReadings.map((r) => periodKeyOf(r.date)))].sort();
  if (currentPeriod !== "all" && !periods.includes(currentPeriod)) currentPeriod = "all";

  const hasData = allReadings.length > 0;
  document.getElementById("empty-state").hidden = hasData;
  document.getElementById("dashboard").hidden = !hasData;

  const coverage = document.getElementById("coverage");
  if (hasData) {
    const first = allReadings[0].date;
    const last = allReadings[allReadings.length - 1].date;
    const days = new Set(allReadings.map((r) => r.date)).size;
    coverage.textContent = `${dateLabel(first)} → ${dateLabel(last)} · ${days} days · ${fmtKwh.format(allReadings.length)} readings`;
  } else {
    coverage.textContent = "No data yet — upload your EDP reports to get started.";
    return;
  }

  // Exclude-tagged-days toggle (only shown once days are tagged)
  const excludeToggle = document.getElementById("exclude-toggle");
  excludeToggle.hidden = tagsMap.size === 0;
  document.getElementById("exclude-check").checked = excludeTagged;
  document.getElementById("exclude-label").textContent =
    `Exclude ${tagsMap.size} tagged day${tagsMap.size === 1 ? "" : "s"}`;

  renderPeriodChips(periods);
  renderKpis();
  renderBilling();
  renderMonthlyChart(periods);
  renderHourlyChart();
  renderDailyChart();
  renderPriceChart();
  renderTopDays();
  renderBaseTrend(periods);
  renderWeather();
  renderHeatmap();
  renderCompare(periods);
  renderFilesTable();

  // Auto-refresh weather once per session if a city was configured before
  if (!weatherAutoTried && localStorage.getItem("weatherCity") && allReadings.length) {
    weatherAutoTried = true;
    const lastReading = allReadings[allReadings.length - 1].date;
    const covered = [...weatherMap.keys()].sort();
    const lastCovered = covered[covered.length - 1] || "";
    if (lastCovered < lastReading) fetchWeather().catch(() => {});
  }
}

function renderPeriodChips(periods) {
  const wrap = document.getElementById("month-filter");
  wrap.innerHTML = "";
  const mk = (value, label) => {
    const b = document.createElement("button");
    b.className = "chip" + (currentPeriod === value ? " active" : "");
    b.textContent = label;
    b.onclick = () => { currentPeriod = value; reload(); };
    wrap.appendChild(b);
  };
  mk("all", "All history");
  periods.forEach((p) => mk(p, periodLabelOf(p)));
}

// Tariff constants for display/estimates: latest parsed invoice wins, defaults otherwise.
function activeTariff() {
  const b = billingMetas[billingMetas.length - 1];
  return {
    k1: b?.k1 ?? DEFAULT_TARIFF.k1,
    k2: b?.k2 ?? DEFAULT_TARIFF.k2,
    tarE: b?.tarE ?? DEFAULT_TARIFF.tarE,
    k3: b?.k3 ?? DEFAULT_TARIFF.k3,
    tarP: b?.tarP ?? DEFAULT_TARIFF.tarP,
  };
}

function fixedDailyEur() {
  const t = activeTariff();
  return t.k3 + t.tarP;
}

function renderKpis() {
  const recs = filteredReadings();
  const totalKwh = recs.reduce((s, r) => s + r.kwh, 0);
  const totalCost = recs.reduce((s, r) => s + (r.cost || 0), 0);
  const days = new Set(recs.map((r) => r.date)).size;

  // Day (08–22h) vs night split
  let dayKwh = 0;
  for (const r of recs) {
    const h = Number(r.time.slice(0, 2));
    if (h >= 8 && h < 22) dayKwh += r.kwh;
  }

  // Average kWh per hour-of-day -> peak hour
  const byHour = groupBy(recs, (r) => Number(r.time.slice(0, 2)));
  let peakHour = null, peakAvg = 0;
  for (const [h, g] of byHour) {
    const avg = g.kwh / g.dates.size;
    if (avg > peakAvg) { peakAvg = avg; peakHour = h; }
  }

  const pomieAvg = recs.length ? recs.reduce((s, r) => s + (r.pomie || 0), 0) / recs.length : 0;

  // Base load: average power drawn during deep night (02:00–05:59), when
  // most usage is always-on appliances (fridge, router, standby, ...).
  const night = recs.filter((r) => { const h = Number(r.time.slice(0, 2)); return h >= 2 && h < 6; });
  const baseKwhPerQuarter = night.length ? night.reduce((s, r) => s + r.kwh, 0) / night.length : 0;
  const baseW = baseKwhPerQuarter * 4 * 1000;
  const baseDailyKwh = (baseW * 24) / 1000;
  const basePct = totalKwh && days ? ((baseDailyKwh * days) / totalKwh) * 100 : 0;

  document.getElementById("kpi-total").textContent = `${fmtKwh.format(totalKwh)} kWh`;
  document.getElementById("kpi-days").textContent = `${days} days of data`;
  document.getElementById("kpi-avg-day").textContent = `${fmtKwh2.format(days ? totalKwh / days : 0)} kWh`;
  document.getElementById("kpi-avg-night").textContent =
    `${fmtKwh.format(totalKwh ? (dayKwh / totalKwh) * 100 : 0)}% between 08:00–22:00`;
  document.getElementById("kpi-cost").textContent = fmtEur.format(totalCost + days * fixedDailyEur());
  document.getElementById("kpi-price").innerHTML =
    totalKwh ? `${fmtEur4.format(totalCost / totalKwh)}<span class="unit">/kWh</span>` : "–";
  document.getElementById("kpi-pomie").textContent = `avg POMIE ${fmtKwh2.format(pomieAvg)} €/MWh`;
  document.getElementById("kpi-peak").textContent =
    peakHour === null ? "–" : `${String(peakHour).padStart(2, "0")}h–${String(peakHour + 1).padStart(2, "0")}h`;
  document.getElementById("kpi-peak-sub").textContent =
    peakHour === null ? "–" : `${fmtKwh2.format(peakAvg)} kWh on an average day`;
  document.getElementById("kpi-base").textContent = `${Math.round(baseW)} W`;
  document.getElementById("kpi-base-sub").textContent =
    `≈ ${fmtKwh2.format(baseDailyKwh)} kWh/day · ${fmtKwh.format(basePct)}% of usage`;
}

// Real invoiced amounts (ELE_DINAMICA) vs the amounts recomputed from the
// 15-min readings — a per-invoice reconciliation table.
function renderBilling() {
  const card = document.getElementById("billing-card");
  if (!billingMetas.length) { card.hidden = true; return; }
  card.hidden = false;

  const t = activeTariff();
  document.getElementById("billing-formula").innerHTML =
    `Price formula from your reports: <strong>(POMIE × (1 + losses) × ${t.k1} + ${t.k2} + ${t.tarE}) €/kWh</strong>` +
    ` for energy, plus <strong>${fmtEur4.format(t.k3 + t.tarP)}/day</strong> fixed (K3 + TAR potência).` +
    ` Taxes, levies and IVA are not included — same basis as the report's "Fatura Total antes de Taxas e Impostos".`;

  const tbody = document.querySelector("#billing-table tbody");
  tbody.innerHTML = "";
  for (const b of billingMetas) {
    const recs = allReadings.filter((r) => r.date >= b.start && r.date <= b.end);
    const kwh = recs.reduce((s, r) => s + r.kwh, 0);
    const energy = recs.reduce((s, r) => s + (r.cost || 0), 0);
    const fixed = ((b.k3 ?? t.k3) + (b.tarP ?? t.tarP)) * (b.days || 0);
    const total = energy + fixed;
    const delta = total - b.totalEur;
    const cls = Math.abs(delta) < 0.5 ? "" : delta > 0 ? "delta-up" : "delta-down";
    const tr = document.createElement("tr");
    tr.innerHTML =
      `<td>${b.label || `${dateLabel(b.start)} → ${dateLabel(b.end)}`}</td>` +
      `<td class="num">${b.days || "–"}</td>` +
      `<td class="num">${fmtKwh.format(kwh)} / ${fmtKwh.format(b.billedKwh)}</td>` +
      `<td class="num">${fmtEur.format(energy)}</td>` +
      `<td class="num">${fmtEur.format(fixed)}</td>` +
      `<td class="num">${fmtEur.format(total)}</td>` +
      `<td class="num"><strong>${fmtEur.format(b.totalEur)}</strong></td>` +
      `<td class="num ${cls}">${delta >= 0 ? "+" : ""}${fmtEur.format(delta)}</td>`;
    tbody.appendChild(tr);
  }
}

function upsertChart(id, config) {
  if (charts[id]) { charts[id].destroy(); }
  charts[id] = new Chart(document.getElementById(id), config);
}

const ACCENT = "#e6007e";
const BLUE = "#2f6fed";

function renderMonthlyChart(periods) {
  const byPeriod = groupBy(allReadings, (r) => periodKeyOf(r.date));
  const kwh = periods.map((p) => byPeriod.get(p)?.kwh ?? 0);
  const fd = fixedDailyEur();
  const cost = periods.map((p) => {
    const g = byPeriod.get(p);
    return g ? g.cost + g.dates.size * fd : 0;
  });

  upsertChart("chart-monthly", {
    data: {
      labels: periods.map(periodLabelOf),
      datasets: [
        {
          type: "bar", label: "kWh", data: kwh, yAxisID: "y",
          backgroundColor: periods.map((p) => (currentPeriod === p ? ACCENT : "#f3a8ce")),
          borderRadius: 6,
        },
        {
          type: "line", label: "Est. bill € (pre-tax)", data: cost, yAxisID: "y1",
          borderColor: BLUE, backgroundColor: BLUE, tension: 0.3, pointRadius: 4,
        },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      onClick: (evt, elements) => {
        if (!elements.length) return;
        const p = periods[elements[0].index];
        currentPeriod = currentPeriod === p ? "all" : p;
        reload();
      },
      scales: {
        y: { title: { display: true, text: "kWh" }, beginAtZero: true },
        y1: { position: "right", grid: { drawOnChartArea: false }, title: { display: true, text: "€" }, beginAtZero: true },
      },
      plugins: { legend: { position: "bottom" } },
    },
  });
}

function renderHourlyChart() {
  const recs = filteredReadings();
  const hours = [...Array(24).keys()];

  const wd = recs.filter((r) => dowIndex(r.date) < 5);
  const we = recs.filter((r) => dowIndex(r.date) >= 5);
  const avgPerHour = (subset) => {
    const byHour = groupBy(subset, (r) => Number(r.time.slice(0, 2)));
    return hours.map((h) => {
      const g = byHour.get(h);
      return g ? g.kwh / g.dates.size : 0;
    });
  };

  upsertChart("chart-hourly", {
    type: "line",
    data: {
      labels: hours.map((h) => `${String(h).padStart(2, "0")}h`),
      datasets: [
        { label: "Weekdays", data: avgPerHour(wd), borderColor: ACCENT, backgroundColor: "rgba(230,0,126,0.08)", fill: true, tension: 0.35, pointRadius: 0 },
        { label: "Weekends", data: avgPerHour(we), borderColor: BLUE, backgroundColor: "rgba(47,111,237,0.08)", fill: true, tension: 0.35, pointRadius: 0 },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: { y: { title: { display: true, text: "kWh / hour" }, beginAtZero: true } },
      plugins: { legend: { position: "bottom" } },
      interaction: { mode: "index", intersect: false },
    },
  });
}

function renderDailyChart() {
  const recs = filteredReadings();
  const byDay = groupBy(recs, (r) => r.date);
  const days = [...byDay.keys()].sort();
  const kwh = days.map((d) => byDay.get(d).kwh);
  const isMonth = currentPeriod !== "all";

  document.getElementById("daily-title").textContent =
    isMonth ? `Daily consumption — ${periodLabelOf(currentPeriod)}` : "Daily consumption — all history";

  upsertChart("chart-daily", {
    type: isMonth ? "bar" : "line",
    data: {
      labels: days.map((d) => (isMonth ? d.slice(8) : dateLabel(d))),
      datasets: [{
        label: "kWh", data: kwh,
        backgroundColor: isMonth ? ACCENT : "rgba(230,0,126,0.1)",
        borderColor: ACCENT, borderRadius: 4, fill: !isMonth, tension: 0.2, pointRadius: 0,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: {
        y: { title: { display: true, text: "kWh" }, beginAtZero: true },
        x: { ticks: { maxTicksLimit: 20 } },
      },
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { title: (items) => (isMonth ? dateLabel(days[items[0].dataIndex]) : items[0].label) } },
      },
    },
  });
}

function renderPriceChart() {
  const recs = filteredReadings();
  const hours = [...Array(24).keys()];

  const kwhByHour = Array(24).fill(0);
  const pomieSum = Array(24).fill(0);
  const pomieCount = Array(24).fill(0);
  const datesByHour = Array.from({ length: 24 }, () => new Set());
  let wSum = 0, wKwh = 0, simpleSum = 0, simpleCount = 0;
  for (const r of recs) {
    const h = Number(r.time.slice(0, 2));
    kwhByHour[h] += r.kwh;
    datesByHour[h].add(r.date);
    if (r.pomie !== null) {
      pomieSum[h] += r.pomie;
      pomieCount[h]++;
      wSum += r.kwh * r.pomie;
      wKwh += r.kwh;
      simpleSum += r.pomie;
      simpleCount++;
    }
  }
  const avgKwh = hours.map((h) => (datesByHour[h].size ? kwhByHour[h] / datesByHour[h].size : 0));
  const avgPomie = hours.map((h) => (pomieCount[h] ? pomieSum[h] / pomieCount[h] : 0));

  // Consumption-weighted price vs flat average → is usage skewed to pricey hours?
  const weighted = wKwh ? wSum / wKwh : 0;
  const flat = simpleCount ? simpleSum / simpleCount : 0;
  const diffPct = flat ? (weighted / flat - 1) * 100 : 0;

  // Cheapest 3-hour window of the average day
  let bestStart = 0, bestPrice = Infinity;
  for (let h = 0; h <= 21; h++) {
    const p = (avgPomie[h] + avgPomie[h + 1] + avgPomie[h + 2]) / 3;
    if (p < bestPrice) { bestPrice = p; bestStart = h; }
  }

  const el = document.getElementById("price-insight");
  el.innerHTML =
    `Your usage is skewed to <strong>${diffPct >= 0 ? "pricier" : "cheaper"}</strong> hours: ` +
    `weighted price <strong>${fmtKwh2.format(weighted)} €/MWh</strong> vs ${fmtKwh2.format(flat)} €/MWh flat average ` +
    `(<strong>${diffPct >= 0 ? "+" : ""}${fmtKwh.format(diffPct)}%</strong>). ` +
    `Cheapest window: <strong>${String(bestStart).padStart(2, "0")}h–${String(bestStart + 3).padStart(2, "0")}h</strong> ` +
    `(${fmtKwh2.format(bestPrice)} €/MWh).`;

  upsertChart("chart-price", {
    data: {
      labels: hours.map((h) => `${String(h).padStart(2, "0")}h`),
      datasets: [
        {
          type: "bar", label: "Your consumption (kWh/h)", data: avgKwh, yAxisID: "y",
          backgroundColor: "rgba(230,0,126,0.35)", borderRadius: 3,
        },
        {
          type: "line", label: "POMIE (€/MWh)", data: avgPomie, yAxisID: "y1",
          borderColor: BLUE, backgroundColor: BLUE, tension: 0.35, pointRadius: 0, borderWidth: 2,
        },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: {
        y: { title: { display: true, text: "kWh / hour" }, beginAtZero: true },
        y1: { position: "right", grid: { drawOnChartArea: false }, title: { display: true, text: "€/MWh" }, beginAtZero: true },
      },
      plugins: { legend: { position: "bottom" } },
      interaction: { mode: "index", intersect: false },
    },
  });
}

function renderTopDays() {
  // Tagged days must stay visible (else they could never be un-tagged),
  // so this table ignores the exclusion toggle. The reference average
  // uses only non-excluded days, matching the rest of the dashboard.
  const recs = periodOnlyReadings();
  const byDay = groupBy(recs, (r) => r.date);
  const days = [...byDay.entries()].map(([date, g]) => ({ date, kwh: g.kwh, cost: g.cost }));
  const included = days.filter((d) => !(excludeTagged && tagsMap.has(d.date)));
  const avg = included.length ? included.reduce((s, d) => s + d.kwh, 0) / included.length : 0;
  days.sort((a, b) => b.kwh - a.kwh);
  const top = days.slice(0, 8);
  const maxKwh = top.length ? top[0].kwh : 0;

  const tbody = document.querySelector("#top-days tbody");
  tbody.innerHTML = "";
  for (const d of top) {
    const tagged = tagsMap.has(d.date);
    const tr = document.createElement("tr");
    if (tagged && excludeTagged) tr.className = "tagged";
    const pct = avg ? ((d.kwh / avg - 1) * 100) : 0;
    tr.innerHTML =
      `<td>${dateLabel(d.date)} <span class="hint">(${DOW_NAMES[dowIndex(d.date)]})</span></td>` +
      `<td class="num bar-cell"><div class="bar" style="width:${maxKwh ? (d.kwh / maxKwh) * 100 : 0}%"></div>` +
      `<span>${fmtKwh2.format(d.kwh)}</span></td>` +
      `<td class="num">${fmtEur.format(d.cost)}</td>` +
      `<td class="num">${pct >= 0 ? "+" : ""}${fmtKwh.format(pct)}%</td>`;
    const td = document.createElement("td");
    const btn = document.createElement("button");
    btn.className = "btn mini";
    btn.textContent = tagged ? "Untag" : "Tag";
    btn.title = tagged ? "Remove tag (day counts in averages again)" : "Tag as unusual day (excluded from averages)";
    btn.onclick = async () => {
      if (tagged) await dbDelete("tags", d.date);
      else await dbPutMany("tags", [{ date: d.date, taggedAt: new Date().toISOString() }]);
      reload();
    };
    td.appendChild(btn);
    tr.appendChild(td);
    tbody.appendChild(tr);
  }

  const tagged = [...tagsMap.keys()];
  const summary = document.getElementById("tag-summary-text");
  if (tagged.length) {
    let kwh = 0, cost = 0;
    for (const r of allReadings) {
      if (tagsMap.has(r.date)) { kwh += r.kwh; cost += r.cost || 0; }
    }
    summary.textContent = `${tagged.length} tagged day(s): ${fmtKwh.format(kwh)} kWh · ${fmtEur.format(cost)}`;
  } else {
    summary.textContent = "No tagged days yet.";
  }
}

// Tag days that are clear outliers vs the typical (median) day.
async function autoTagSpikes() {
  const byDay = groupBy(allReadings, (r) => r.date);
  const days = [...byDay.entries()].map(([date, g]) => ({ date, kwh: g.kwh }));
  const sorted = days.map((d) => d.kwh).sort((a, b) => a - b);
  const median = sorted[Math.floor(sorted.length / 2)] || 0;
  const threshold = Math.max(median * 2.5, 15);
  const spikes = days.filter((d) => d.kwh > threshold && !tagsMap.has(d.date));
  if (!spikes.length) {
    toast(`No untagged days above ${fmtKwh.format(threshold)} kWh (2.5× your median day).`);
    return;
  }
  await dbPutMany("tags", spikes.map((d) => ({ date: d.date, auto: true, taggedAt: new Date().toISOString() })));
  await reload();
  toast(`Tagged ${spikes.length} day(s) above ${fmtKwh.format(threshold)} kWh (2.5× your median day).`);
}

function renderBaseTrend(periods) {
  const baseW = periods.map((p) => {
    const night = allReadings.filter((r) => {
      if (periodKeyOf(r.date) !== p || isExcluded(r)) return false;
      const h = Number(r.time.slice(0, 2));
      return h >= 2 && h < 6;
    });
    if (!night.length) return 0;
    return (night.reduce((s, r) => s + r.kwh, 0) / night.length) * 4 * 1000;
  });

  upsertChart("chart-baseload", {
    type: "line",
    data: {
      labels: periods.map(periodLabelOf),
      datasets: [{
        label: "Base load (W)", data: baseW.map((v) => Math.round(v)),
        borderColor: ACCENT, backgroundColor: "rgba(230,0,126,0.08)",
        fill: true, tension: 0.3, pointRadius: 4,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: { y: { title: { display: true, text: "W" }, beginAtZero: true } },
      plugins: { legend: { display: false } },
    },
  });
}

// ---------- Weather ----------

function pearson(xs, ys) {
  const n = xs.length;
  if (n < 3) return 0;
  const mx = xs.reduce((a, b) => a + b, 0) / n;
  const my = ys.reduce((a, b) => a + b, 0) / n;
  let cov = 0, sx = 0, sy = 0;
  for (let i = 0; i < n; i++) {
    cov += (xs[i] - mx) * (ys[i] - my);
    sx += (xs[i] - mx) ** 2;
    sy += (ys[i] - my) ** 2;
  }
  return sx && sy ? cov / Math.sqrt(sx * sy) : 0;
}

async function fetchWeather() {
  const status = document.getElementById("weather-status");
  const city = document.getElementById("weather-city").value.trim() || localStorage.getItem("weatherCity") || "";
  if (!city || !allReadings.length) {
    status.textContent = city ? "Upload reports first." : "Enter a city first.";
    return;
  }
  try {
    status.textContent = "Fetching…";
    const geo = await (await fetch(
      `https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(city)}&count=1&language=pt&format=json`
    )).json();
    if (!geo.results?.length) throw new Error(`city "${city}" not found`);
    const { latitude, longitude, name } = geo.results[0];
    const start = allReadings[0].date;
    const end = allReadings[allReadings.length - 1].date;
    const wx = await (await fetch(
      `https://archive-api.open-meteo.com/v1/archive?latitude=${latitude}&longitude=${longitude}` +
      `&start_date=${start}&end_date=${end}&daily=temperature_2m_mean&timezone=auto`
    )).json();
    if (!wx.daily?.time) throw new Error("no weather data returned");
    const records = wx.daily.time
      .map((d, i) => ({ date: d, tmean: wx.daily.temperature_2m_mean[i] }))
      .filter((r) => r.tmean !== null && r.tmean !== undefined);
    await dbPutMany("weather", records);
    localStorage.setItem("weatherCity", name);
    weatherMap = new Map(records.map((w) => [w.date, w.tmean]));
    renderWeather();
  } catch (err) {
    console.error(err);
    status.textContent = `Weather fetch failed: ${err.message}`;
  }
}

function renderWeather() {
  const cityInput = document.getElementById("weather-city");
  if (!cityInput.value) cityInput.value = localStorage.getItem("weatherCity") || "";

  const recs = filteredReadings();
  const byDay = groupBy(recs, (r) => r.date);
  const points = [];
  for (const [date, g] of byDay) {
    const t = weatherMap.get(date);
    if (t !== undefined) points.push({ x: t, y: g.kwh, date });
  }

  const status = document.getElementById("weather-status");
  const insight = document.getElementById("weather-insight");
  if (!points.length) {
    status.textContent = weatherMap.size ? "No overlap between weather and selected period." : "";
    insight.textContent = "Load historical temperatures (Open-Meteo, free) to see how weather drives your consumption.";
    upsertChart("chart-weather", { type: "scatter", data: { datasets: [] }, options: { responsive: true, maintainAspectRatio: false } });
    return;
  }

  status.textContent = `${localStorage.getItem("weatherCity")} · ${points.length}/${byDay.size} days matched`;
  const r = pearson(points.map((p) => p.x), points.map((p) => p.y));
  const strength = Math.abs(r) > 0.5 ? "strong" : Math.abs(r) > 0.3 ? "moderate" : "weak";
  const direction = r < 0 ? "colder days push your consumption up (heating)" : "warmer days push your consumption up (cooling)";
  insight.innerHTML = Math.abs(r) < 0.15
    ? `Correlation <strong>r = ${fmtKwh2.format(r)}</strong> — your consumption barely depends on outside temperature.`
    : `Correlation <strong>r = ${fmtKwh2.format(r)}</strong> — a ${strength} relationship: ${direction}.`;

  upsertChart("chart-weather", {
    type: "scatter",
    data: {
      datasets: [{
        label: "day", data: points,
        backgroundColor: "rgba(230,0,126,0.55)", pointRadius: 4,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: {
        x: { title: { display: true, text: "Mean temperature (°C)" } },
        y: { title: { display: true, text: "kWh / day" }, beginAtZero: true },
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (item) => `${dateLabel(item.raw.date)}: ${fmtKwh2.format(item.raw.y)} kWh at ${fmtKwh.format(item.raw.x)} °C`,
          },
        },
      },
    },
  });
}

function renderCompare(periods) {
  const card = document.getElementById("compare-card");
  if (periods.length < 2) { card.hidden = true; return; }
  card.hidden = false;

  if (!periods.includes(cmpSel.a)) cmpSel.a = periods[periods.length - 2];
  if (!periods.includes(cmpSel.b)) cmpSel.b = periods[periods.length - 1];

  for (const [id, key] of [["cmp-a", "a"], ["cmp-b", "b"]]) {
    const sel = document.getElementById(id);
    sel.innerHTML = "";
    for (const p of periods) {
      const opt = document.createElement("option");
      opt.value = p;
      opt.textContent = periodLabelOf(p);
      sel.appendChild(opt);
    }
    sel.value = cmpSel[key];
    sel.onchange = () => { cmpSel[key] = sel.value; renderCompare(periods); };
  }

  const side = (p) => {
    const recs = allReadings.filter((r) => periodKeyOf(r.date) === p && !isExcluded(r));
    const kwh = recs.reduce((s, r) => s + r.kwh, 0);
    const cost = recs.reduce((s, r) => s + (r.cost || 0), 0);
    const days = new Set(recs.map((r) => r.date)).size;
    const bill = cost + days * fixedDailyEur();
    return { recs, kwh, cost, bill, days, perDay: days ? kwh / days : 0, price: kwh ? cost / kwh : 0 };
  };
  const A = side(cmpSel.a), B = side(cmpSel.b);

  const stat = (label, va, vb, fmt) => {
    const pct = va ? (vb / va - 1) * 100 : 0;
    const cls = pct >= 0 ? "delta-up" : "delta-down";
    return `<div class="cmp-stat">${label}<b>${fmt(va)} → ${fmt(vb)} ` +
      `<span class="${cls}">(${pct >= 0 ? "+" : ""}${fmtKwh.format(pct)}%)</span></b></div>`;
  };
  document.getElementById("cmp-stats").innerHTML =
    stat("Total", A.kwh, B.kwh, (v) => `${fmtKwh.format(v)} kWh`) +
    stat("Average per day", A.perDay, B.perDay, (v) => `${fmtKwh2.format(v)} kWh`) +
    stat("Estimated bill (pre-tax)", A.bill, B.bill, (v) => fmtEur.format(v)) +
    stat("Average energy price", A.price, B.price, (v) => fmtEur4.format(v));

  const hours = [...Array(24).keys()];
  const profile = (recs) => {
    const byHour = groupBy(recs, (r) => Number(r.time.slice(0, 2)));
    return hours.map((h) => {
      const g = byHour.get(h);
      return g ? g.kwh / g.dates.size : 0;
    });
  };

  upsertChart("chart-compare", {
    type: "line",
    data: {
      labels: hours.map((h) => `${String(h).padStart(2, "0")}h`),
      datasets: [
        { label: periodLabelOf(cmpSel.a), data: profile(A.recs), borderColor: BLUE, backgroundColor: "rgba(47,111,237,0.08)", fill: true, tension: 0.35, pointRadius: 0 },
        { label: periodLabelOf(cmpSel.b), data: profile(B.recs), borderColor: ACCENT, backgroundColor: "rgba(230,0,126,0.08)", fill: true, tension: 0.35, pointRadius: 0 },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: { y: { title: { display: true, text: "kWh / hour (average day)" }, beginAtZero: true } },
      plugins: { legend: { position: "bottom" } },
      interaction: { mode: "index", intersect: false },
    },
  });
}

function renderHeatmap() {
  const recs = filteredReadings();
  // avg kWh per (weekday, hour): total kwh / number of distinct dates with that weekday
  const cells = Array.from({ length: 7 }, () => Array(24).fill(0));
  const dayCount = Array(7).fill(0);
  const seenDates = new Set();
  for (const r of recs) {
    const dow = dowIndex(r.date);
    cells[dow][Number(r.time.slice(0, 2))] += r.kwh;
    if (!seenDates.has(r.date)) { seenDates.add(r.date); dayCount[dow]++; }
  }
  let max = 0;
  const avg = cells.map((row, d) => row.map((v) => {
    const a = dayCount[d] ? v / dayCount[d] : 0;
    if (a > max) max = a;
    return a;
  }));

  const wrap = document.getElementById("heatmap");
  const grid = document.createElement("div");
  grid.className = "hm-grid";
  grid.appendChild(document.createElement("div"));
  for (let h = 0; h < 24; h++) {
    const el = document.createElement("div");
    el.className = "hm-hour";
    el.textContent = h % 3 === 0 ? `${h}h` : "";
    grid.appendChild(el);
  }
  for (let d = 0; d < 7; d++) {
    const lbl = document.createElement("div");
    lbl.className = "hm-label";
    lbl.textContent = DOW_NAMES[d];
    grid.appendChild(lbl);
    for (let h = 0; h < 24; h++) {
      const cell = document.createElement("div");
      cell.className = "hm-cell";
      const t = max ? avg[d][h] / max : 0;
      cell.style.background = heatColor(t);
      cell.title = `${DOW_NAMES[d]} ${String(h).padStart(2, "0")}:00 — ${fmtKwh2.format(avg[d][h])} kWh avg`;
      grid.appendChild(cell);
    }
  }
  const legend = document.createElement("div");
  legend.className = "hm-legend";
  legend.innerHTML = `<span>0 kWh</span><div class="hm-legend-bar"></div><span>${fmtKwh2.format(max)} kWh</span>`;
  wrap.replaceChildren(grid, legend);
}

function heatColor(t) {
  // #f4f5f8 -> #e6007e
  const from = [244, 245, 248], to = [230, 0, 126];
  const c = from.map((f, i) => Math.round(f + (to[i] - f) * t));
  return `rgb(${c[0]},${c[1]},${c[2]})`;
}

function renderFilesTable() {
  const tbody = document.querySelector("#files-table tbody");
  tbody.innerHTML = "";
  for (const f of fileMetas) {
    const tr = document.createElement("tr");
    const up = new Date(f.uploadedAt);
    tr.innerHTML = `<td>${f.name}</td><td>${dateLabel(f.from)} → ${dateLabel(f.to)}</td>` +
      `<td>${fmtKwh.format(f.rows)}</td><td>${up.toLocaleDateString("en-GB")}</td>`;
    tbody.appendChild(tr);
  }
}

// ---------- Backup ----------

function exportBackup() {
  const payload = {
    version: 3,
    exportedAt: new Date().toISOString(),
    files: fileMetas,
    readings: allReadings,
    tags: [...tagsMap.values()],
    billing: billingMetas,
  };
  const blob = new Blob([JSON.stringify(payload)], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `energy-backup-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
}

async function importBackup(file) {
  try {
    const payload = JSON.parse(await file.text());
    if (!Array.isArray(payload.readings)) throw new Error("Invalid backup file");
    if (!payload.version || payload.version < 3) {
      throw new Error("backup predates the v3 cost-formula change — re-import the XLSX reports instead");
    }
    await dbPutMany("readings", payload.readings);
    if (Array.isArray(payload.files)) await dbPutMany("files", payload.files);
    if (Array.isArray(payload.tags)) await dbPutMany("tags", payload.tags);
    if (Array.isArray(payload.billing)) await dbPutMany("billing", payload.billing);
    await reload();
    toast(`Backup restored — ${fmtKwh.format(payload.readings.length)} readings merged.`);
  } catch (err) {
    console.error(err);
    toast(`Could not import backup: ${err.message}`, true);
  }
}

// ---------- Events ----------

function setupEvents() {
  const fileInput = document.getElementById("file-input");
  const jsonInput = document.getElementById("json-input");

  document.getElementById("btn-upload").onclick = () => fileInput.click();
  document.getElementById("empty-state").onclick = () => fileInput.click();
  fileInput.onchange = () => { ingestFiles(fileInput.files); fileInput.value = ""; };

  document.getElementById("btn-export").onclick = exportBackup;
  document.getElementById("btn-import").onclick = () => jsonInput.click();
  jsonInput.onchange = () => { if (jsonInput.files[0]) importBackup(jsonInput.files[0]); jsonInput.value = ""; };

  document.getElementById("btn-clear").onclick = async () => {
    if (!confirm("Delete ALL stored readings and history? Consider exporting a backup first.")) return;
    await dbClearAll();
    currentPeriod = "all";
    await reload();
    toast("All data cleared.");
  };

  document.getElementById("exclude-check").onchange = (e) => {
    excludeTagged = e.target.checked;
    localStorage.setItem("excludeTagged", excludeTagged ? "1" : "0");
    reload();
  };
  document.getElementById("btn-autotag").onclick = autoTagSpikes;
  document.getElementById("btn-weather").onclick = fetchWeather;

  // Calendar months vs billing periods (25th → 24th)
  document.querySelectorAll("#period-mode button").forEach((btn) => {
    btn.onclick = () => {
      if (periodMode === btn.dataset.mode) return;
      periodMode = btn.dataset.mode;
      currentPeriod = "all";
      cmpSel = { a: null, b: null };
      document.querySelectorAll("#period-mode button").forEach((b) => b.classList.toggle("active", b === btn));
      reload();
    };
  });

  // Drag & drop anywhere on the page
  const overlay = document.getElementById("drop-overlay");
  let dragDepth = 0;
  window.addEventListener("dragenter", (e) => { e.preventDefault(); dragDepth++; overlay.classList.add("active"); });
  window.addEventListener("dragleave", (e) => { e.preventDefault(); if (--dragDepth <= 0) { dragDepth = 0; overlay.classList.remove("active"); } });
  window.addEventListener("dragover", (e) => e.preventDefault());
  window.addEventListener("drop", (e) => {
    e.preventDefault();
    dragDepth = 0;
    overlay.classList.remove("active");
    if (e.dataTransfer?.files?.length) ingestFiles(e.dataTransfer.files);
  });
}

Chart.defaults.font.family = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif";
Chart.defaults.color = "#6b7385";
Chart.defaults.locale = "pt-PT";

setupEvents();
reload();

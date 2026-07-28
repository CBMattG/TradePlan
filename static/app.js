'use strict';
/* 2026 Tradeplan app — replicates the Excel tradeplan calculation chain.
   All recalculation happens here, instantly, as order quantities are edited. */

const WEEKS = 53;
let M = null;            // master data (read-only reference data)
let YEAR = '2026';       // active plan year
let YEARS = ['2026'];    // available years
let ORDERS = {};         // sku id -> [53] order quantities  (the editable input)
let IMPORTED = {};       // snapshot of Excel-imported orders (for "changed" markers)
let PO_WEBSA = null;     // { pos:{PO#:{supplier,lines:[{code,ordered,delivered,outstanding,due}]}}, codes, rows }
let PO_CONTAINERS = null;// { dates:{PO#:[{deliveryCB,etaPort,etd,status,container,shipment,supplier}]}, unparsed, truncated, rows }
let CHANNEL_INDEX = null; // { skus:{code:[{c,r,p}]}, customers:{no:name}, importedAt } — Channel Sales import (global)
let CHANGELOG = [];       // change-log entries (newest first), fetched from /api/changelog
let CHANGE_LABEL = null;  // one-shot label for the next data-edit save (else "Manual edits")
function labelNextSave(lbl) { CHANGE_LABEL = lbl; }   // tag the next autosave with a semantic action name
let PO_UNMATCHED = null;  // snapshot list for the "no PO" review cycle
let poCycleIdx = -1;      // cursor into PO_UNMATCHED
let poReviewActive = false;
let PO_RETIME = null;     // last computed re-time plan (for the preview dialog)
let PO_APPLY_UNDO = null; // { snapshot:{id:[53]}, scope, label } to undo the last apply
let PO_WEEK_TIP = {};     // week -> rich hover-tooltip HTML (per-PO qty breakdown for the on-screen supplier)
let PO_SCHED_CUR = null;  // on-screen supplier: week -> PO arrival entries (raw, for per-product cell tooltips)
let PO_POS_CUR = null;    // on-screen supplier: Map PO# -> { lines, qty } (raw, for per-product cell tooltips)
let SETTINGS = {};
let RES = new Map();     // sku id -> computed results
let AGG = null;          // aggregates
let skuById = new Map();
let supByName = new Map();
let currentSupplier = null;
let currentView = 'plan';
let searchTerm = '';
// view prefs persisted across years and app restarts (localStorage)
function loadPref(key, fallback) {
  try { const v = localStorage.getItem(key); return v == null ? fallback : JSON.parse(v); }
  catch { return fallback; }
}
function savePref(key, val) { try { localStorage.setItem(key, JSON.stringify(val)); } catch {} }
let GHOST_ON = loadPref('tp_ghost', true);   // supplier cards: draw the faint last-year trend line

/* ---------------- colour theme (light / dark) ---------------- */
// The saved theme is applied to <html data-theme> by an inline script in index.html
// (before first paint) and persisted under the plain key 'tp-theme'.
function currentTheme() { return document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light'; }
function applyTheme(theme) {
  if (theme === 'dark') document.documentElement.setAttribute('data-theme', 'dark');
  else document.documentElement.removeAttribute('data-theme');
  try { localStorage.setItem('tp-theme', theme); } catch {}
  updateThemeButton();
}
function toggleTheme() { applyTheme(currentTheme() === 'dark' ? 'light' : 'dark'); }
function updateThemeButton() {
  const b = document.getElementById('btn-theme');
  if (!b) return;
  const dark = currentTheme() === 'dark';
  // show the icon for the mode you'd switch TO
  b.textContent = dark ? '☀️' : '🌙';
  b.title = dark ? 'Switch to light mode' : 'Switch to dark mode';
}
const SORT_MODES = ['default', 'value-desc', 'value-asc', 'name'];
let sortMode = SORT_MODES.includes(loadPref('tp_sortMode')) ? loadPref('tp_sortMode') : 'default';
const STATUS_FILTERS = ['all', 'live', 'notlive'];
let statusFilter = STATUS_FILTERS.includes(loadPref('tp_statusFilter')) ? loadPref('tp_statusFilter') : 'all';   // rebuy catalog status
const REBUY_SCOPES = ['sup', 'all'];   // run/clear/commit rebuy suggestions for the current supplier only, or the whole year
let rebuyScope = REBUY_SCOPES.includes(loadPref('tp_rebuyScope')) ? loadPref('tp_rebuyScope') : 'sup';
let MODELED = new Map();     // sku id -> [53] seasonality-modelled forecast (when enabled)
let WEATHER = null;          // { ok, weeks: { wk: {temp,normal,anomaly} } }
let CALIB = null;            // empirical seasonal curves calibrated from sales history
const PROFILE_CACHE = new Map();       // group (season|category) curves
const SKU_PROFILE_CACHE = new Map();   // final per-SKU curves (group blended with own shape)
const FLAT = new Array(WEEKS).fill(1);
const EMPTY53 = new Array(WEEKS).fill(0);
function SEASON() { return SETTINGS.seasonality || {}; }
let saveTimer = null;
const DEFAULT_ROWS = ['ly', 'forecast', 'orders', 'proposed', 'stock', 'cover'];
const _savedRows = loadPref('tp_visibleRows');
let visible = new Set(Array.isArray(_savedRows) && _savedRows.length ? _savedRows : DEFAULT_ROWS);
// Plan-page product ordering within a supplier + which product accordions are open
const SUP_SORTS = ['orig', 'value', 'az'];
let supSort = SUP_SORTS.includes(loadPref('tp_supSort')) ? loadPref('tp_supSort') : 'orig';
let openSkus = new Set(Array.isArray(loadPref('tp_openSkus')) ? loadPref('tp_openSkus') : []);
let accForceOpen = false;    // search active → matching products render expanded
let reopenDD = null;         // keep a toolbar dropdown open across a renderPlan()

const ROWDEFS = [
  { key: 'ly',       label: 'LY Sales',            fmt: fmtU },
  { key: 'actual',   label: 'Actual Sales',        fmt: fmtU },
  { key: 'forecast', label: 'Sales Forecast',      fmt: fmtU },
  { key: 'value',    label: 'Sales Value',         fmt: fmtGBP },
  { key: 'orders',   label: 'Committed Orders',    edit: true },
  { key: 'proposed', label: 'Proposed Rebuy',      edit: true, prop: true },
  { key: 'fobRow',   label: 'Order FOB',           fmt: fmtGBP },
  { key: 'cbmRow',   label: 'CBM',                 fmt: fmt1 },
  { key: 'stock',    label: 'Stock Closing',       fmt: fmtU },
  { key: 'shv',      label: 'Stock Holding Value', fmt: fmtGBP },
  { key: 'cover',    label: 'Weeks Cover',         fmt: fmt1 },
];

/* ---------------- formatting ---------------- */
function fmtU(n) { return Math.abs(n) < .5 ? '–' : Math.round(n).toLocaleString('en-GB'); }
function fmt1(n) { return Math.abs(n) < .05 ? '–' : (Math.round(n * 10) / 10).toLocaleString('en-GB', { minimumFractionDigits: 1, maximumFractionDigits: 1 }); }
function fmtGBP(n) { return Math.abs(n) < .5 ? '–' : '£' + Math.round(n).toLocaleString('en-GB'); }
function fmtGBPk(n) { return '£' + (n >= 1e6 ? (n / 1e6).toFixed(2) + 'm' : Math.round(n / 1e3).toLocaleString() + 'k'); }
function esc(s) { return String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function titleCase(s) { return String(s ?? '').toLowerCase().replace(/\b[a-z]/g, c => c.toUpperCase()); }

function isoWeek(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - day);
  const y0 = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil(((date - y0) / 864e5 + 1) / 7);
}
function weekDate(w) { // w/c date of week w for the loaded year (from its week1_start)
  const base = (typeof M !== 'undefined' && M && M.week1_start) ? M.week1_start : '2025-12-29';
  const [y, mo, dd] = base.split('-').map(Number);
  const d = new Date(y, mo - 1, dd);
  d.setDate(d.getDate() + (w - 1) * 7);
  return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' });
}
// The real current planning week — but only for the year that actually contains
// today (so past/future years get no "this week" highlight). 0 = no highlight.
function highlightWeek() {
  if (typeof M === 'undefined' || !M || !M.week1_start) return 0;
  const [y, mo, dd] = M.week1_start.split('-').map(Number);
  const w1 = Date.UTC(y, mo - 1, dd);                    // UTC midnights → whole-day diff, DST-safe
  const now = new Date();
  const today = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
  const wk = Math.floor((today - w1) / (7 * 86400000)) + 1;
  return (wk >= 1 && wk <= WEEKS) ? wk : 0;
}

/* ---------------- Weeks Cover colour bands ---------------- */
function defaultCoverBands() {
  // Sweet-spot model: ~4 weeks is the green target; too low risks a stock-out,
  // too high means overstock. Every threshold/colour is editable in Settings.
  return [
    { max: 0,    bg: '#b30000' },  // stock-out (no cover)
    { max: 2,    bg: '#e06666' },  // critically low
    { max: 3,    bg: '#f0a35e' },  // low
    { max: 4,    bg: '#93c47d' },  // on target (~4 weeks) — good
    { max: 6,    bg: '#ffe066' },  // a little high
    { max: 8,    bg: '#f0a35e' },  // high
    { max: 12,   bg: '#e06666' },  // overstocked
    { max: null, bg: '#c27ba0' },  // heavily overstocked (catch-all)
  ];
}
function coverBand(v) {
  const bands = SETTINGS.cover_bands || [];
  for (const b of bands) {
    if (b.max === null || b.max === undefined || b.max === '') return b;  // catch-all
    if (v <= Number(b.max)) return b;
  }
  return bands[bands.length - 1] || { bg: '#ffffff' };
}
function textOn(hex) {  // pick readable text colour for a background
  const c = String(hex).replace('#', '');
  if (c.length < 6) return '#1c2733';
  const r = parseInt(c.substr(0, 2), 16), g = parseInt(c.substr(2, 2), 16), b = parseInt(c.substr(4, 2), 16);
  return (0.299 * r + 0.587 * g + 0.114 * b) < 140 ? '#fff' : '#1c2733';
}
function parseColor(s) {            // '#rgb' / '#rrggbb' / 'rgb(r,g,b)' → [r,g,b]
  s = String(s).trim();
  if (s[0] === '#') { const c = s.slice(1); const f = c.length === 3 ? c.split('').map(x => x + x).join('') : c; return [parseInt(f.substr(0, 2), 16), parseInt(f.substr(2, 2), 16), parseInt(f.substr(4, 2), 16)]; }
  const m = s.match(/(\d+)[,\s]+(\d+)[,\s]+(\d+)/); return m ? [+m[1], +m[2], +m[3]] : [255, 255, 255];
}
function rgbStr(a) { return `rgb(${Math.round(a[0])},${Math.round(a[1])},${Math.round(a[2])})`; }
function mixColor(a, b, t) { const pa = parseColor(a), pb = parseColor(b); t = Math.max(0, Math.min(1, t)); return rgbStr([pa[0] + (pb[0] - pa[0]) * t, pa[1] + (pb[1] - pa[1]) * t, pa[2] + (pb[2] - pa[2]) * t]); }
function textOnColor(s) { const [r, g, b] = parseColor(s); return (0.299 * r + 0.587 * g + 0.114 * b) < 145 ? '#fff' : '#1c2733'; }
// Continuous weeks-cover colour: interpolate between the cover-band anchors (each band's
// `max` is the anchor value, its colour the anchor colour) so the scale is smooth, not blocky.
function coverColor(v) {
  const bands = (SETTINGS.cover_bands && SETTINGS.cover_bands.length) ? SETTINGS.cover_bands : defaultCoverBands();
  const fin = bands.filter(b => b.max !== null && b.max !== undefined && b.max !== '').map(b => ({ x: Number(b.max), c: b.bg })).sort((a, b) => a.x - b.x);
  const tail = bands.find(b => b.max === null || b.max === undefined || b.max === '');
  if (!fin.length) return rgbStr(parseColor(tail ? tail.bg : '#ffffff'));
  if (v <= fin[0].x) return rgbStr(parseColor(fin[0].c));
  for (let i = 0; i < fin.length - 1; i++) if (v <= fin[i + 1].x) return mixColor(fin[i].c, fin[i + 1].c, (v - fin[i].x) / (fin[i + 1].x - fin[i].x));
  const last = fin[fin.length - 1];
  if (tail) { const span = Math.max(2, last.x - (fin.length > 1 ? fin[fin.length - 2].x : 0)); return mixColor(last.c, tail.bg, Math.min(1, (v - last.x) / span)); }
  return rgbStr(parseColor(last.c));
}
// Gradient background for a cover cell at week w: blends from the midpoint with the previous
// cell to the midpoint with the next, so colours flow continuously across the row (boundaries
// match between neighbours). Returns { bg: linear-gradient, color: readable text }.
function coverCellStyle(arr, w) {
  const mid = coverColor(arr[w]);
  const left = w > 0 ? mixColor(coverColor(arr[w - 1]), mid, 0.5) : mid;
  const right = w < arr.length - 1 ? mixColor(mid, coverColor(arr[w + 1]), 0.5) : mid;
  return { bg: `linear-gradient(to right, ${left}, ${right})`, color: textOnColor(mid) };
}

/* ---------------- calculation engine (mirror of the Excel formulas) ---------------- */
function computeSku(sku, orderOverride) {
  const mult = SETTINGS.multiplier, cur = SETTINGS.current_week;
  const ord = orderOverride || ORDERS[sku.id] || new Array(WEEKS).fill(0);
  const fc = new Array(WEEKS), stock = new Array(WEEKS), value = new Array(WEEKS),
        cover = new Array(WEEKS), fobRow = new Array(WEEKS), cbmRow = new Array(WEEKS),
        shv = new Array(WEEKS);
  const mod = MODELED.get(sku.id);   // seasonality model (future weeks only), if enabled
  for (let w = 0; w < WEEKS; w++) {
    const bf = (mod && (w + 1) >= cur) ? mod[w] : sku.base_forecast[w];
    fc[w] = Math.round(bf * mult);
  }
  for (let w = 0; w < WEEKS; w++) {
    const wk = w + 1;
    if (wk === 1 && cur !== 1) stock[w] = sku.running_stock[w];
    else if (wk === cur) stock[w] = ord[w] + sku.stock_now;
    else if (wk < cur) stock[w] = sku.running_stock[w];
    else stock[w] = Math.max(0, stock[w - 1] + ord[w] - fc[w]);
  }
  for (let w = 0; w < WEEKS; w++) {
    value[w] = (w + 1) < cur ? sku.actual[w] * sku.asp : Math.min(fc[w], stock[w]) * sku.asp;
    fobRow[w] = ord[w] * sku.fob;
    cbmRow[w] = ord[w] * sku.cbm;
    shv[w] = stock[w] * sku.landed;
    const seg = fc.slice(w, w + 4);
    const avg = seg.reduce((a, b) => a + b, 0) / seg.length;
    cover[w] = avg ? stock[w] / avg : stock[w];
  }
  // "Sales Forecast" row AS DISPLAYED: realised actuals for the weeks that have them
  // (1..data_week-1), then the forecast for the rest of the year — so the row reads as
  // the live full-year outturn (matching the Sales Value row, which already actualises)
  // and equals next year's "LY Sales". The pure forecast `fc` above still drives every
  // calculation (stock projection, cover, stock-out flags, rebuy demand, coverage %).
  const dataBnd = ((M && +M.data_week) || cur) - 1;
  const forecastDisp = fc.map((v, w) => (w < dataBnd ? (sku.actual[w] || 0) : v));
  return { ly: sku.ly, actual: sku.actual, forecast: fc, forecastDisp, orders: ord, stock, cover, value, fobRow, cbmRow, shv };
}

function zeros() { return new Array(WEEKS).fill(0); }

/* ---------------- seasonality + weather forecast model ---------------- */
function circDist(a, b) { const d = Math.abs(a - b); return Math.min(d, 53 - d); }
function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }

function smoothCircular(arr, win) {
  const n = arr.length, half = Math.floor(win / 2), out = new Array(n);
  for (let i = 0; i < n; i++) { let s = 0; for (let k = -half; k <= half; k++) s += arr[(i + k + n) % n]; out[i] = s / (2 * half + 1); }
  return out;
}
function normMean1(arr) { const m = arr.reduce((a, b) => a + b, 0) / arr.length; return m > 0 ? arr.map(v => v / m) : arr.map(() => 1); }
function peakWeek(prof) { let mi = 0; for (let i = 1; i < prof.length; i++) if (prof[i] > prof[mi]) mi = i; return mi + 1; }
// Compress a mean-1 profile's deviation from its mean so its peak is at most `limit`×
// the mean, then renormalise. Used to stop a low-seasonality (continuity) product's
// own last-year shape — which may be sharply peaked from a one-off spike or an early
// stockout — from being imposed at full height and inflating the de-seasonalised rate.
function capProfilePeak(prof, limit) {
  let mx = 0; for (const v of prof) if (v > mx) mx = v;
  if (mx <= limit) return prof;
  const k = (limit - 1) / (mx - 1);
  return normMean1(prof.map(v => 1 + (v - 1) * k));
}
const CONT_PEAK_CAP = 2.0;   // max peak-to-mean for a continuity product's seasonal shape (backtest-tuned on 2025)

// Empirical weekly demand shape (mean ~1) aggregated from real sales across a set
// of SKUs: last-year full-year shape, refined by this year's actuals where we have
// them. Returns null if there isn't enough history to be trustworthy.
function buildEmpiricalProfile(skus, cur) {
  const ly = zeros(), ty = zeros();
  let lyTot = 0;
  for (const s of skus) {
    for (let w = 0; w < WEEKS; w++) { ly[w] += s.ly[w] || 0; if (w < cur - 1) ty[w] += s.actual[w] || 0; }
    lyTot += (s.ly || []).reduce((a, b) => a + b, 0);
  }
  if (lyTot < 40) return null;                         // too little history → caller falls back
  let raw = smoothCircular(ly, 5);
  const earlyEnd = Math.max(1, cur - 1);
  let lyEarly = 0, tyEarly = 0;
  for (let w = 0; w < earlyEnd; w++) { lyEarly += raw[w]; tyEarly += ty[w]; }
  if (tyEarly > 0 && lyEarly > 0) {                    // blend this year's actuals into the weeks we have
    const tys = smoothCircular(ty.map(v => v * (lyEarly / tyEarly)), 5);
    for (let w = 0; w < earlyEnd; w++) raw[w] = 0.6 * raw[w] + 0.4 * tys[w];
  }
  let prof = normMean1(raw);
  const floor = 0.12;
  prof = prof.map(v => v * (1 - floor) + floor);
  return normMean1(prof);
}

// Build calibrated curves per season-group and per category from sales history.
function calibrateSeasonalShapes(cur) {
  const bySeason = new Map(), byCat = new Map(), groupS = {}, groupC = {};
  for (const s of M.skus) {
    (groupS[s.season] = groupS[s.season] || []).push(s);
    (groupC[s.category] = groupC[s.category] || []).push(s);
  }
  for (const [k, arr] of Object.entries(groupS)) { const p = buildEmpiricalProfile(arr, cur); if (p) bySeason.set(k, p); }
  for (const [k, arr] of Object.entries(groupC)) { const p = buildEmpiricalProfile(arr, cur); if (p) byCat.set(k, p); }
  CALIB = { bySeason, byCat, cur };
  PROFILE_CACHE.clear();
  SKU_PROFILE_CACHE.clear();
}

// Hardcoded Gaussian fallback when there isn't enough history to calibrate.
function gaussianProfile(season, category) {
  const s = (season || '').toLowerCase(), cat = (category || '').toUpperCase();
  let peak = 28, width = 13;
  if (s.includes('winter')) { peak = 52; width = 11; }
  else if (!s.includes('summer')) {
    if (/HEATER|CHIMINEA|COVER|FLEECE|GRIT|SALT|SNOW/.test(cat)) { peak = 50; width = 12; }
    else if (/BBQ|COOK|PIZZA|GRILL/.test(cat)) { peak = 28; width = 9; }
  }
  let raw = new Array(WEEKS);
  for (let w = 0; w < WEEKS; w++) raw[w] = Math.exp(-0.5 * (circDist(w + 1, peak) / width) ** 2);
  let prof = normMean1(raw).map(v => v * 0.85 + 0.15);
  return normMean1(prof);
}

// Group (season + category) calibrated curve — shared by every SKU in the group:
// the calibrated season curve, nudged by the category curve when it peaks in the
// same part of the year. This is the broad average shape.
function groupProfile(sku) {
  const key = (sku.season || '') + '|' + (sku.category || '');
  if (PROFILE_CACHE.has(key)) return PROFILE_CACHE.get(key);
  let prof;
  const seasonP = CALIB && CALIB.bySeason.get(sku.season);
  const catP = CALIB && CALIB.byCat.get(sku.category);
  if (seasonP && catP && circDist(peakWeek(seasonP), peakWeek(catP)) <= 10) {
    prof = normMean1(seasonP.map((v, w) => 0.6 * v + 0.4 * catP[w]));   // blend, season-anchored
  } else {
    prof = seasonP || catP || gaussianProfile(sku.season, sku.category);
  }
  PROFILE_CACHE.set(key, prof);
  return prof;
}

// Final per-SKU profile: the baseline curve sharpened by blending in the product's OWN
// last-year sales shape, so a sharp single-product spike (e.g. a Christmas line peaking
// in week 50) isn't smeared flat by the broader season/category average. The own shape
// is only trusted in proportion to how much of its own history the product has (>=40
// units needed; full trust by ~200), and the blend weight is the Forecast setting
// "Own-shape weight" (default 60%).
//   - Seasonal lines start from the calibrated season+category GROUP curve.
//   - Continuity lines carry no shared season, so we don't impose a group curve on them;
//     they start FLAT but still pick up their own within-year uplifts/drops (continuity
//     products are less seasonal, but a storage line that consistently lifts in spring,
//     or one whose sales are simply growing, should not be flat-lined). Because the flat
//     baseline holds no shape, continuity leans harder on its own history.
function seasonProfile(sku) {
  const s = (sku.season || '').toLowerCase();
  if (SKU_PROFILE_CACHE.has(sku.id)) return SKU_PROFILE_CACHE.get(sku.id);
  const isCont = s.includes('continu');
  let prof = isCont ? FLAT.slice() : groupProfile(sku);
  const owBase = clamp(SEASON().ownShapeWeight != null ? +SEASON().ownShapeWeight : 0.6, 0, 1);
  const ow = isCont ? clamp(owBase * 1.4, 0, 0.85) : owBase;   // flat baseline → trust own shape more
  if (ow > 0) {
    const lyTot = (sku.ly || []).reduce((a, b) => a + (b || 0), 0);
    const owEff = ow * clamp(lyTot / 200, 0, 1);     // lean on the baseline when own history is thin
    const own = owEff > 0 ? buildEmpiricalProfile([sku], (CALIB && CALIB.cur) || SETTINGS.current_week) : null;
    if (own) prof = normMean1(prof.map((v, w) => (1 - owEff) * v + owEff * own[w]));
  }
  if (isCont) prof = capProfilePeak(prof, CONT_PEAK_CAP);   // continuity = low-seasonality: don't impose a sharp spike
  SKU_PROFILE_CACHE.set(sku.id, prof);
  return prof;
}

function ensureCalib() {
  if (!CALIB || CALIB.cur !== SETTINGS.current_week) calibrateSeasonalShapes(SETTINGS.current_week);
}

// Per-SKU model: returns the demand-rate signals, the applied profile, and the
// resulting (strength-blended) modelled weekly forecast. Used by the builder and
// the per-SKU explainer. Assumes calibration has run (call ensureCalib first).
function computeSkuModel(sku) {
  const se = SEASON(), cur = SETTINGS.current_week, strength = seasonStrengthFor(YEAR);
  const prof = seasonProfile(sku);
  const seasonal = /summer|winter/i.test(sku.season || '');
  const isCont = /continu/i.test(sku.season || '');
  // (a) this-year de-seasonalised rate, ignoring weeks that were out of stock
  let tyNum = 0, tyDen = 0;
  for (let w = 0; w < cur - 1 && w < WEEKS; w++) {
    const inStock = sku.running_stock[w] > 0 || sku.actual[w] > 0;
    if (inStock) { tyNum += sku.actual[w]; tyDen += prof[w]; }
  }
  const rateTy = tyDen > 0 ? tyNum / tyDen : null;
  // (b) last-year rate, excluding suspected peak-season stockouts (0 sales in a peak week)
  const peakThresh = 0.6 * Math.max(...prof);
  let lyNum = 0, lyDen = 0, suspectCount = 0;
  for (let w = 0; w < WEEKS; w++) {
    const v = sku.ly[w] || 0;
    const suspect = seasonal && v === 0 && prof[w] >= peakThresh &&
      ((sku.ly[w - 1] || 0) > 0 || (sku.ly[w + 1] || 0) > 0);
    lyNum += v;
    if (suspect) suspectCount++; else lyDen += prof[w];
  }
  // A product with NO last-year sales at all (a brand-new line) has *no* prior-year
  // signal — 53 zeros mean "no history", not "we sold nothing". Reading them as a
  // genuine rate of 0 gives that zero an ~80% vote and crushes the level towards
  // nothing, so a new line entered with a 70/yr forecast would model out at ~6/yr
  // and round to 0 in every week. Drop the LY term instead and let the planner's
  // own annual figure (rateOrig) carry the level.
  const rateLy = (lyDen > 0 && lyNum > 0) ? lyNum / lyDen : null;
  // (c) original planner forecast as a stabiliser
  const rateOrig = sku.base_forecast.reduce((a, b) => a + b, 0) / WEEKS;
  let wTy = rateTy != null ? Math.min(tyDen, 8) : 0;
  let wLy = rateLy != null ? Math.min(lyDen, 8) * 0.8 : 0;
  let wOrig = 2;
  if (isCont) { wTy *= 2.4; wLy *= 0.3; wOrig = 1; }   // continuity: track this year's actual run-rate
  // On a GENERATED forecast year the "original" base_forecast isn't an independent planner
  // forecast — it's just the previous model's rollover. Anchoring the level on it is circular
  // and biases low (any line that outsold its old plan gets dragged back down), so we keep it
  // only as a light stabiliser here and lean on realised sales (LY) instead.
  if (M && +M.data_week === 1) wOrig *= 0.25;
  // Down-weight this year's run-rate when only a small share of the product's annual
  // demand has actually been observed (e.g. a Christmas line seen only in its quiet
  // off-season): otherwise those near-zero off-season weeks drag the demand rate down
  // and the whole line gets under-forecast. Full weight only once ~half the year's
  // demand mass is in. Continuity lines (mild shape) accumulate mass ~linearly so are
  // barely affected once mid-year.
  if (rateTy != null) {
    const totProf = prof.reduce((a, b) => a + b, 0) || WEEKS;
    wTy *= clamp((tyDen / totProf) / 0.5, 0, 1);
  }
  let D = (wTy * (rateTy || 0) + wLy * (rateLy || 0) + wOrig * rateOrig) / (wTy + wLy + wOrig);
  if (!isFinite(D) || D < 0) D = rateOrig;
  const cap = 4 * Math.max(rateOrig, rateLy || 0, 0.5);   // guard runaway numbers; don't clip a strong realised rate by a low prior forecast
  // optional whole-forecast year-on-year growth uplift (Settings → Forecast). Applied to the
  // modelled output only, so Target Sales mode (which re-normalises to its £ target) is unaffected.
  const growth = 1 + clamp(growthPctFor(YEAR), -90, 500) / 100;
  const modeled = new Array(WEEKS);
  for (let w = 0; w < WEEKS; w++) {
    let m = Math.min(D, cap) * prof[w];
    if (se.useWeather && WEATHER && WEATHER.ok && WEATHER.weeks) {
      const an = WEATHER.weeks[w + 1];
      if (an && typeof an.anomaly === 'number' && seasonal) {
        const sens = (se.weatherStrength ?? 0.5) * 0.06;
        const dir = /summer/i.test(sku.season) ? 1 : -1;
        m *= 1 + dir * clamp(an.anomaly, -8, 8) * sens;
      }
    }
    modeled[w] = Math.max(0, (1 - strength) * sku.base_forecast[w] + strength * m) * growth;
  }
  return { prof, seasonal, isCont, rateTy, rateLy, rateOrig, wTy, wLy, wOrig, D: Math.min(D, cap), suspectCount, modeled, growth };
}

// ---- per-year forecast mode + parameters ----
// SETTINGS.seasonality.byYear[year] = { mode:'off'|'seasonality'|'target', strength, growth, target }.
// Each year's forecast is built independently; weather / own-shape / location are shared.
function seasonCfg(year) {
  const by = (SETTINGS.seasonality && SETTINGS.seasonality.byYear) || {};
  return by[String(year)] || null;
}
function seasonModeFor(year) { const c = seasonCfg(year); return (c && c.mode) || 'off'; }
function seasonActiveFor(year) { return seasonModeFor(year) !== 'off'; }   // this year's forecast is re-modelled
function seasonStrengthFor(year) { const c = seasonCfg(year); return (c && c.strength != null) ? +c.strength : 0.7; }
function targetValueFor(year) {   // the £ sales target set for a year in Target mode (0 = none)
  const c = seasonCfg(year);
  return (c && c.mode === 'target' && +c.target > 0) ? +c.target : 0;
}
function growthPctFor(year) {      // YoY growth % applied in Seasonality mode (0 = none / other modes)
  const c = seasonCfg(year);
  return (c && c.mode === 'seasonality' && isFinite(+c.growth)) ? +c.growth : 0;
}
function seasonYears() {           // forecast byYear map, creating it on first use
  const s = SETTINGS.seasonality || (SETTINGS.seasonality = {});
  return s.byYear || (s.byYear = {});
}
// Migrate the old global enabled/strength/targetMode + per-year target/growth maps into byYear.
function migrateSeasonality(s) {
  if (s && !s.byYear) {
    const by = {}, tby = s.targetByYear || {}, gby = s.growthByYear || {};
    const str = s.strength != null ? s.strength : 0.7;
    for (const y of new Set([...Object.keys(tby), ...Object.keys(gby)])) {
      const t = +tby[y] || 0, g = +gby[y] || 0;
      if (s.targetMode && t > 0) by[y] = { mode: 'target', target: t, strength: str, growth: g };
      else if (s.enabled) by[y] = { mode: 'seasonality', strength: str, growth: g, target: t };
    }
    s.byYear = by;
    delete s.enabled; delete s.strength; delete s.targetMode; delete s.targetByYear; delete s.growthByYear;
  }
  return s;
}
function buildModeledForecasts() {
  MODELED = new Map();
  const mode = seasonModeFor(YEAR);
  if (mode === 'off') return;
  calibrateSeasonalShapes(SETTINGS.current_week);   // calibrate peaks from actual sales history
  if (mode === 'target') { buildTargetForecasts(targetValueFor(YEAR)); return; }
  for (const sku of M.skus) MODELED.set(sku.id, computeSkuModel(sku).modeled);
}

// Target Sales mode: distribute an overall £ sales target across products by their
// recency-weighted historical sales weight (the model's de-seasonalised demand D,
// which favours the most recent year), shaped weekly by each product's seasonality
// + category profile, then normalised so the whole-sheet forecast value hits target.
function buildTargetForecasts(target) {
  const se = SEASON(), cur = SETTINGS.current_week;
  const info = [];
  let sumHistVal = 0;
  for (const sku of M.skus) {
    const m = computeSkuModel(sku);                 // gives prof + D (recency-weighted demand rate)
    const histUnits = Math.max(0, m.D * WEEKS);     // annual units implied by weighted history
    const histVal = histUnits * (sku.asp || 0);     // weight the target by £ contribution
    sumHistVal += histVal;
    info.push({ sku, prof: m.prof, histUnits });
  }
  if (sumHistVal <= 0) { for (const sku of M.skus) MODELED.set(sku.id, computeSkuModel(sku).modeled); return; }
  for (const it of info) {
    const U = target * it.histUnits / sumHistVal;   // annual target units for this product (Σ value = target)
    const profSum = it.prof.reduce((a, b) => a + b, 0) || WEEKS;
    const arr = new Array(WEEKS);
    for (let w = 0; w < WEEKS; w++) {
      let v = U * it.prof[w] / profSum;             // weekly split by seasonality (profile mean 1)
      if (se.useWeather && WEATHER && WEATHER.ok && WEATHER.weeks && /summer|winter/i.test(it.sku.season || '')) {
        const an = WEATHER.weeks[w + 1];
        if (an && typeof an.anomaly === 'number') {
          const sens = (se.weatherStrength ?? 0.5) * 0.06, dir = /summer/i.test(it.sku.season) ? 1 : -1;
          v *= 1 + dir * clamp(an.anomaly, -8, 8) * sens;
        }
      }
      arr[w] = Math.max(0, v);
    }
    MODELED.set(it.sku.id, arr);
  }
  // normalise the modelled FUTURE weeks so past actuals + future forecast value = target
  // (exact for forecast years where cur=1; corrects rounding / weather / banked actuals)
  const mult = SETTINGS.multiplier || 1;
  let pastVal = 0, futureVal = 0;
  for (const sku of M.skus) {
    const mod = MODELED.get(sku.id), asp = sku.asp || 0;
    for (let w = 0; w < WEEKS; w++) {
      if ((w + 1) < cur) pastVal += (sku.actual[w] || 0) * asp;
      else futureVal += (mod[w] || 0) * mult * asp;
    }
  }
  if (futureVal > 0) {
    const k = Math.max(0, (target - pastVal) / futureVal);
    for (const sku of M.skus) { const mod = MODELED.get(sku.id); for (let w = 0; w < WEEKS; w++) if ((w + 1) >= cur) mod[w] *= k; }
  }
}
// Uncapped whole-sheet forecast demand value (Σ forecast units × ASP) — used to
// show how close the modelled forecast lands to the sales target.
function forecastDemandTotal() {
  let v = 0;
  for (const sku of M.skus) { const r = RES.get(sku.id); if (r) for (let w = 0; w < WEEKS; w++) v += r.forecast[w] * (sku.asp || 0); }
  return v;
}

function computeAll() {
  RES = new Map();
  const bySup = new Map();
  const g = { sales: zeros(), units: zeros(), fob: zeros(), cbm: zeros(), shv: zeros(),
              fcUnits: zeros(), fcSales: zeros(), pallet: zeros(), stillage: zeros(), racking: zeros() };
  const cSales = zeros();                          // sales value with committed + proposed orders
  let propUnits = 0, propFob = 0, propCbm = 0;     // proposed-rebuy additions (purely additive)
  for (const sku of M.skus) {
    const r = computeSku(sku);
    RES.set(sku.id, r);
    let s = bySup.get(sku.supplier);
    if (!s) { s = { sales: zeros(), units: zeros(), fob: zeros(), cbm: zeros(), shv: zeros(), fcUnits: zeros(), fcSales: zeros(), fcValue: 0, covNum: 0, covDen: 0 }; bySup.set(sku.supplier, s); }
    // forecast value (uncapped) + stocked-in coverage (committed supply vs forecast demand)
    const fcUnitsSku = r.forecast.reduce((a, b) => a + b, 0);
    const supplySku = (sku.stock_now || 0) + (ORDERS[sku.id] || EMPTY53).reduce((a, b) => a + b, 0);
    const fcValueSku = fcUnitsSku * (sku.asp || 0);
    s.fcValue += fcValueSku;
    if (fcUnitsSku > 0) { const w8 = fcValueSku || fcUnitsSku; s.covNum += (supplySku / fcUnitsSku) * w8; s.covDen += w8; }
    const asp = sku.asp || 0;
    for (let w = 0; w < WEEKS; w++) {
      const fcv = r.forecast[w] * asp;             // uncapped forecast sales £ (the plan)
      s.sales[w] += r.value[w];  s.units[w] += r.orders[w]; s.fob[w] += r.fobRow[w];
      s.cbm[w] += r.cbmRow[w];   s.shv[w] += r.shv[w];      s.fcUnits[w] += r.forecast[w]; s.fcSales[w] += fcv;
      g.sales[w] += r.value[w];  g.units[w] += r.orders[w]; g.fob[w] += r.fobRow[w];
      g.cbm[w] += r.cbmRow[w];   g.shv[w] += r.shv[w];      g.fcUnits[w] += r.forecast[w]; g.fcSales[w] += fcv;
      if (sku.fpq > 0 && sku.pallet_type) {
        const sp = Math.ceil(r.stock[w] / sku.fpq);
        if (sku.pallet_type === 'Pallet') g.pallet[w] += sp;
        else if (sku.pallet_type === 'Stillage') g.stillage[w] += sp;
        else if (sku.pallet_type === 'Racking') g.racking[w] += sp;
      }
    }
    // combined (committed + proposed) sales is stock-dependent, so re-simulate
    // SKUs that have proposals; FOB/units/CBM are additive so summed directly
    const prop = PROPOSED && PROPOSED.get(sku.id);
    const hasP = prop && prop.some(v => v);
    const rc = hasP ? computeSku(sku, combinedOrder(sku.id)) : r;
    for (let w = 0; w < WEEKS; w++) cSales[w] += rc.value[w];
    if (hasP) for (let w = 0; w < WEEKS; w++) {
      const q = prop[w] || 0;
      if (q) { propUnits += q; propFob += q * (sku.fob || 0); propCbm += q * (sku.cbm || 0); }
    }
  }
  const cap = SETTINGS.capacities;
  const rackWebsa = cap.racking_websa + cap.racking_refit;
  const rackTotal = cap.racking_websa + cap.racking_express + cap.racking_refit;
  const stillTotal = cap.stillage_websa + cap.stillage_lough + cap.stillage_express + (cap.stillage_free || 0);
  g.containers = g.cbm.map(v => v / SETTINGS.container_cbm);
  g.rackUseWebsa = g.pallet.map((p, w) => (p + g.racking[w]) / rackWebsa);
  g.rackUseTotal = g.pallet.map((p, w) => (p + g.racking[w]) / rackTotal);
  g.stillUseWebsa = g.stillage.map(v => v / cap.stillage_websa);
  g.stillUseTotal = g.stillage.map(v => v / stillTotal);
  const sum = a => a.reduce((x, y) => x + y, 0);
  g.totSales = sum(g.sales); g.totFob = sum(g.fob); g.totUnits = sum(g.units);
  g.totCbm = sum(g.cbm); g.totContainers = g.totCbm / SETTINGS.container_cbm;
  // committed / proposed split for the whole-plan totals band
  g.totSalesProposed = sum(cSales) - g.totSales;          // stock-dependent sales uplift
  g.totFobProposed = propFob;
  g.totUnitsProposed = propUnits;
  g.totCbmProposed = propCbm;
  g.totContainersProposed = propCbm / SETTINGS.container_cbm;
  g.peakShv = Math.max(...g.shv); g.peakShvWk = g.shv.indexOf(g.peakShv) + 1;
  g.peakRack = Math.max(...g.rackUseTotal); g.peakRackWk = g.rackUseTotal.indexOf(g.peakRack) + 1;
  g.peakStill = Math.max(...g.stillUseTotal); g.peakStillWk = g.stillUseTotal.indexOf(g.peakStill) + 1;
  for (const s of bySup.values()) {
    s.forecastValue = s.fcValue;
    s.stockedPct = s.covDen > 0 ? (s.covNum / s.covDen) * 100 : null;
  }
  AGG = { g, bySup };
}

/* ---------------- persistence ---------------- */
function markDirty() {
  document.getElementById('save-status').textContent = 'Saving…';
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveNow, 1200);
}
async function saveNow() {
  clearTimeout(saveTimer); saveTimer = null;
  try {
    const label = CHANGE_LABEL || 'Manual edits'; CHANGE_LABEL = null;
    const r = await fetch('/api/save?year=' + encodeURIComponent(YEAR), { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ orders: ORDERS, settings: SETTINGS, proposed: serializeProposed(), changeLabel: label }) });
    const j = await r.json();
    document.getElementById('save-status').textContent = j.ok
      ? 'Saved ' + new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })
      : 'Save failed!';
  } catch { document.getElementById('save-status').textContent = 'Save failed!'; }
}
window.addEventListener('beforeunload', e => { if (saveTimer) { saveNow(); e.preventDefault(); } });
// Remember the prior-year-end stock a forecast year was last chained to, so we can tell
// when a year's rebuy suggestions have gone stale (not on every visit).
function persistStockbase(obj) {
  fetch('/api/save?year=' + encodeURIComponent(YEAR), { method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ stockbase: obj }) }).catch(() => {});
}

/* ---------------- whole-plan totals band ---------------- */
function totalsHtml() {
  const g = AGG.g;
  let live = 0, notlive = 0;
  for (const s of M.skus) { if (s.status === 'Live') live++; else if (s.status === 'Not Live') notlive++; }
  const num = n => Math.round(n).toLocaleString('en-GB');
  const simple = (v, l) => `<div class="tb-item"><div class="v">${v}</div><div class="l">${l}</div></div>`;
  // committed / proposed / total breakdown: big value = total (committed + proposed),
  // sub-line shows the committed base and the proposed-rebuy addition
  const broken = (com, prop, fmt, l) =>
    `<div class="tb-item tb-bd"><div class="v">${fmt(com + prop)}</div><div class="l">${l}</div>`
    + `<div class="tb-split"><span class="c-com">${fmt(com)} committed</span>`
    + `<span class="c-prop">${prop > 0 ? '+' : ''}${fmt(prop)} proposed</span></div></div>`;
  return '<div class="tb-title">Whole&#8209;plan totals</div>'
    + broken(g.totSales, g.totSalesProposed, fmtGBPk, 'Forecast sales ' + YEAR)
    + broken(g.totFob, g.totFobProposed, fmtGBPk, 'Order FOB spend')
    + broken(g.totUnits, g.totUnitsProposed, num, 'Order units')
    + broken(g.totCbm, g.totCbmProposed, v => num(v) + ' cbm', 'Order volume')
    + broken(g.totContainers, g.totContainersProposed, v => v.toFixed(1), 'Containers')
    + simple(`<span class="c-live">${live}</span> / <span class="c-notlive">${notlive}</span>`, 'SKUs Live / Not Live')
    + simple(fmtGBPk(g.peakShv), 'Peak stock £ · W' + g.peakShvWk)
    + simple((g.peakRack * 100).toFixed(0) + '%', 'Peak racking · W' + g.peakRackWk)
    + simple((g.peakStill * 100).toFixed(0) + '%', 'Peak stillage · W' + g.peakStillWk);
}
function fillTotals() {
  const el = document.getElementById('totals-band');
  if (el) el.innerHTML = totalsHtml();
}

/* ---------------- sidebar ---------------- */
function supplierMatches(sup, term) {
  if (sup.name.toLowerCase().includes(term)) return true;
  return M.skus.some(k => k.supplier === sup.name &&
    (k.code.toLowerCase().includes(term) || (k.name || '').toLowerCase().includes(term)));
}
function supTotalSales(name) {
  const t = AGG.bySup.get(name);
  return t ? t.sales.reduce((a, b) => a + b, 0) : 0;
}
function supForecastValue(name) { const t = AGG.bySup.get(name); return t ? (t.forecastValue || 0) : 0; }
function stockedClass(pct) {
  if (pct == null) return 'stk-na';
  if (pct < 50) return 'stk-lo';
  if (pct < 90) return 'stk-mid';
  if (pct <= 115) return 'stk-ok';
  return 'stk-over';
}
// Supplier names that have any uncommitted Proposed Rebuy stock (one pass over PROPOSED).
function suppliersWithProposed() {
  const set = new Set();
  if (!PROPOSED) return set;
  for (const [id, arr] of PROPOSED) if (arr.some(v => v)) { const s = skuById.get(id); if (s) set.add(s.supplier); }
  return set;
}
// Toggle one supplier's sidebar caution flag without re-rendering the whole list.
function updateSupWarnIcon(supName) {
  const item = [...document.querySelectorAll('.sup-item')].find(el => el.dataset.sup === supName);
  if (!item) return;
  let has = false;
  if (PROPOSED) for (const sku of M.skus) { if (sku.supplier !== supName) continue; const p = PROPOSED.get(sku.id); if (p && p.some(v => v)) { has = true; break; } }
  const warn = item.querySelector('.si-warn');
  if (has && !warn) {
    const el = document.createElement('span');
    el.className = 'si-warn'; el.textContent = '⚠';
    el.title = 'Has Proposed Rebuy stock not yet committed — needs attention';
    item.insertBefore(el, item.firstChild);
  } else if (!has && warn) warn.remove();
}
function renderSidebar() {
  const term = searchTerm.toLowerCase();
  let sups = M.suppliers.filter(s => !term || supplierMatches(s, term));
  const metric = supForecastValue;
  if (sortMode === 'value-desc') sups = sups.slice().sort((a, b) => metric(b.name) - metric(a.name));
  else if (sortMode === 'value-asc') sups = sups.slice().sort((a, b) => metric(a.name) - metric(b.name));
  else if (sortMode === 'name') sups = sups.slice().sort((a, b) => a.name.localeCompare(b.name));

  const sortLabel = { 'value-desc': `forecast value, high→low`, 'value-asc': `forecast value, low→high`,
                      'name': 'supplier name (A–Z)', 'default': 'workbook order' }[sortMode];
  document.getElementById('sidebar-legend').innerHTML =
    `Sorted by <b>${sortLabel}</b> · <b>£ = forecast sales value</b> (demand × price, uncapped) · <b>% = stocked-in</b> (committed supply vs forecast; &lt;100% needs orders, &gt;100% overstock)`;
  document.querySelectorAll('.sortbtn').forEach(b => b.classList.toggle('active', b.dataset.sort === sortMode));

  const list = document.getElementById('supplier-list');
  const withProp = suppliersWithProposed();
  list.innerHTML = sups.map(s => {
    const t = AGG.bySup.get(s.name) || {};
    const fc = t.forecastValue || 0, pct = t.stockedPct;
    const nm = titleCase(s.name);
    const pctTxt = pct == null ? '–' : (pct >= 1000 ? '999+%' : Math.round(pct) + '%');
    const warn = withProp.has(s.name) ? `<span class="si-warn" title="Has Proposed Rebuy stock not yet committed — needs attention">⚠</span>` : '';
    return `<div class="sup-item${s.name === currentSupplier ? ' active' : ''}" data-sup="${esc(s.name)}">
      ${warn}<span class="n" title="${esc(nm)}">${esc(nm)}</span>
      <span class="si-fc" title="Forecast sales value (uncapped)">${fmtGBPk(fc)}</span>
      <span class="si-stk ${stockedClass(pct)}" title="Stocked-in %: committed supply vs forecast demand">${pctTxt}</span></div>`;
  }).join('') || '<div class="empty">No match</div>';
}

/* ---------------- plan view ---------------- */
function headerRow(propWeeks) {
  const hw = highlightWeek();
  let h = '<tr><th class="lbl">Week</th>';
  for (let w = 1; w <= WEEKS; w++) {
    const cls = [w === hw ? 'curwk' : '', propWeeks && propWeeks.has(w) ? 'has-prop' : ''].filter(Boolean).join(' ');
    const tip = propWeeks && propWeeks.has(w) ? ' title="Proposed rebuy orders present in this week"' : '';
    h += `<th class="${cls}"${tip}>W${w}<span class="d">${weekDate(w)}</span></th>`;
  }
  return h + '<th>Total</th></tr>';
}

// Set of week numbers (1-based) that contain any proposed rebuy across the given SKUs.
function proposedWeekSet(skus) {
  const set = new Set();
  if (!PROPOSED) return set;
  for (const k of skus) {
    const arr = PROPOSED.get(k.id);
    if (!arr) continue;
    for (let w = 0; w < WEEKS; w++) if (arr[w]) set.add(w + 1);
  }
  return set;
}

// Rich hover-tooltip for one week of the PO row: each PO landing that week, its outstanding
// quantity, and the per-product breakdown — so when several containers share a week you can
// see what quantity is from which PO. `entries` = that week's schedule rows; `pos` = the
// supplier's PO→{lines} map (for the product lines).
function poWeekTipHtml(w, entries, pos) {
  let body = '', tot = 0;
  const seen = new Set();
  for (const e of entries) {
    if (seen.has(e.po)) continue; seen.add(e.po);
    const info = pos.get(e.po);
    const lines = info ? info.lines.slice().sort((a, b) => b.outstanding - a.outstanding) : [];
    const qty = info ? info.qty : e.qty;
    tot += qty;
    const tag = e.booked ? (e.container ? ` · ${esc(e.container)}` : '') : ' · not booked (PO due date)';
    body += `<div class="pt-po"><span class="pt-ponum${e.booked ? '' : ' pt-unbooked'}">${esc(e.po)}</span><b>${fmtU(qty)}</b></div>`
      + `<div class="pt-meta">${e.booked ? esc(e.basis) : 'awaiting container booking'}${e.status ? ' · ' + esc(e.status) : ''}${tag}</div>`;
    for (const l of lines.slice(0, 6)) body += `<div class="pt-line"><span>${esc(l.code)}</span><span>${fmtU(l.outstanding)}</span></div>`;
    if (lines.length > 6) body += `<div class="pt-line pt-more">…+${lines.length - 6} more product(s)</div>`;
  }
  const n = seen.size;
  return `<div class="st-wk">Week ${w} · w/c ${weekDate(w)}</div>${body}`
    + `<div class="pt-tot">${n} PO${n > 1 ? 's' : ''} · ${fmtU(tot)} units outstanding</div>`;
}
// Per-product hover tooltip for a Committed Orders cell: only the POs landing that week
// which carry THIS product, showing that product's outstanding units (not the whole week).
// Returns null if the hovered product isn't on any PO that week.
function poProductTipHtml(w, code) {
  if (!PO_SCHED_CUR || !PO_POS_CUR) return null;
  const entries = PO_SCHED_CUR[w];
  if (!entries || !entries.length) return null;
  let body = '', tot = 0, n = 0; const seen = new Set();
  for (const e of entries) {
    if (seen.has(e.po)) continue; seen.add(e.po);
    const info = PO_POS_CUR.get(e.po);
    const line = info && info.lines.find(l => l.code === code && l.outstanding > 0);
    if (!line) continue;                            // this PO doesn't carry the hovered product
    n++; tot += line.outstanding;
    const tag = e.booked ? (e.container ? ` · ${esc(e.container)}` : '') : ' · not booked (PO due date)';
    body += `<div class="pt-po"><span class="pt-ponum${e.booked ? '' : ' pt-unbooked'}">${esc(e.po)}</span><b>${fmtU(line.outstanding)}</b></div>`
      + `<div class="pt-meta">${e.booked ? esc(e.basis) : 'awaiting container booking'}${e.status ? ' · ' + esc(e.status) : ''}${tag}</div>`;
  }
  if (!n) return null;                              // hovered product not on any PO this week
  return `<div class="st-wk">Week ${w} · w/c ${weekDate(w)} · ${esc(code)}</div>${body}`
    + `<div class="pt-tot">${n} PO${n > 1 ? 's' : ''} · ${fmtU(tot)} units of this product</div>`;
}
// Sticky PO row (sits under the week header): which Purchase Order / container is bringing
// each week's stock for the supplier. Booked (from Qlik) vs not-yet-booked (PO due date);
// "No PO" where current/future committed stock has no matching booking. '' if no PO data.
function poRowHtml(supName) {
  PO_WEEK_TIP = {};
  if (!poDataReady()) return '';
  const sched = supplierPoSchedule(supName), com = supCommittedByWeek(supName), prop = supProposedByWeek(supName);
  const pos = supplierPOs(supName);                 // PO# -> { lines:[{code,outstanding,…}], qty }
  PO_SCHED_CUR = sched; PO_POS_CUR = pos;           // stash raw data for per-product cell tooltips
  const cur = SETTINGS.current_week, hw = highlightWeek();
  let cells = '', poCount = 0;
  for (let w = 1; w <= WEEKS; w++) {
    const entries = sched[w];
    let inner = '';
    if (entries && entries.length) {
      const seen = new Set();
      for (const e of entries) {
        if (seen.has(e.po)) continue; seen.add(e.po); poCount++;
        inner += `<span class="po-chip po-clk${e.booked ? '' : ' po-unbooked'}" data-po="${esc(e.po)}">${esc(e.po)}</span>`;
      }
      PO_WEEK_TIP[w] = poWeekTipHtml(w, entries, pos);   // rich per-PO breakdown for hover
    } else if (w >= cur && (com[w - 1] + prop[w - 1]) > 0.5) {
      inner = `<span class="po-chip po-none" title="Committed or proposed stock this week with no matching PO / container booking">No PO</span>`;
    }
    cells += `<td class="po-cell${w === hw ? ' curwk' : ''}" data-pw="${w}">${inner}</td>`;
  }
  return `<tr class="po-row"><td class="lbl">PO <span class="po-row-note">/ container</span></td>${cells}<td class="tot">${poCount || ''}</td></tr>`;
}

function statusBadge(status) {
  const s = status || 'Unknown';
  const cls = s === 'Live' ? 'live' : s === 'Not Live' ? 'notlive' : 'unknown';
  const tip = s === 'Live' ? 'Live — plan for rebuy'
            : s === 'Not Live' ? 'Not Live — discontinued; run down remaining stock, do not rebuy'
            : 'Not found in Buying Report — catalog status unknown';
  return `<span class="badge ${cls}" title="${tip}">${s}</span>`;
}

// In a generated forecast year the "LY Sales" row is the prior year's realised actuals
// (weeks 1..lyActualWeeks) chained with its forecast for the remaining weeks — true for
// week index >= lyActualWeeks. Used to subtly tint the forecast tail of the row.
function lyIsForecast(w) { return M && M.lyActualWeeks != null && w >= M.lyActualWeeks; }

function skuRowsHtml(sku, idx) {
  const r = computeSku(sku, combinedOrder(sku.id));   // projection includes proposed rebuys
  const cur = SETTINGS.current_week, hw = highlightWeek();
  const inf = [
    sku.category,
    sku.cbm ? sku.cbm.toFixed(3) + ' cbm' : null,
  ].filter(Boolean).join(' · ');
  const osp = sku.os_purchases;
  const open = accForceOpen || openSkus.has(sku.id);
  const proposedArr0 = (PROPOSED && PROPOSED.get(sku.id)) || EMPTY53;
  const propUnits = proposedArr0.reduce((a, b) => a + b, 0);
  const coverNow = r.cover[Math.max(0, (SETTINGS.current_week || 1) - 1)];
  const cb = coverBand(coverNow);
  const statsHtml = `<div class="skh-stats">`
    + `<span class="skh-stat"><span class="skh-stat-l">Stock now</span><b>${Math.round(sku.stock_now).toLocaleString()}</b></span>`
    + `<span class="skh-stat"><span class="skh-stat-l">On purchase</span><b>${osp != null ? Math.round(osp).toLocaleString() : '—'}</b></span>`
    + `<span class="skh-stat"><span class="skh-stat-l">Cover now</span><b class="skh-cover" style="background:${cb.bg};color:${textOnColor(cb.bg)}">${fmt1(coverNow)} wks</b></span>`
    + (propUnits > 0.5 ? `<span class="skh-prop" title="This product has uncommitted proposed rebuy stock">● rebuy ${fmtU(propUnits)}</span>` : '')
    + `</div>`;
  const img = sku.image ? `<img class="acc-hit" src="${esc(sku.image)}" loading="lazy" onerror="this.remove()" title="Click to expand / collapse this product">` : '';
  const headCls = (sku.status === 'Not Live' ? 'skuhead notlive' : 'skuhead') + (aspSrc(sku) === 'orig' ? ' asp-stale' : '') + (open ? ' acc-open' : '');
  // YTD-to-most-recent-actual-week (week before current): actual sales £ vs forecast £
  const aw = cur - 1;
  let ytdHtml = '';
  if (aw >= 1) {
    const actual = r.value.slice(0, aw).reduce((a, b) => a + b, 0);
    const fcast = r.forecast.slice(0, aw).reduce((a, b) => a + b, 0) * (sku.asp || 0);
    const vv = actual - fcast, pct = fcast ? (vv / fcast) * 100 : 0, sign = vv >= 0 ? '+' : '−', cls = vv >= 0 ? 'c-up' : 'c-down';
    ytdHtml = `<div class="sku-ytd"><span class="ssy-lbl">YTD to W${aw}:</span>`
      + `<span>actual <b>${fmtGBP(actual)}</b></span><span>forecast <b>${fmtGBP(fcast)}</b></span>`
      + `<span>variance <b class="${cls}">${sign}${fmtGBP(Math.abs(vv))} (${sign}${Math.abs(pct).toFixed(0)}%)</b></span></div>`;
  }
  let h = `<tr class="${headCls}" data-acc="${esc(sku.id)}"><td colspan="${WEEKS + 2}"><div class="skuhead-inner"><span class="acc-car" title="Click to expand / collapse this product">▶</span><div class="skh-main">${img}<div class="skh-body">`
    + `<div class="skh-left">`
    +   `<div class="skh-line1"><span class="code acc-hit" title="Click to expand / collapse this product">${esc(sku.code)}</span><span class="nm acc-hit" title="Click to expand / collapse this product"> ${esc(sku.name || '')}</span><button class="sku-explain" data-sku="${esc(sku.id)}" title="Explain this forecast">&#9432;</button><span class="inf">${inf}</span></div>`
    +   `<div class="skh-pills">${statusBadge(sku.status)}${aspChip(sku)}${wkAspChip(sku)}</div>`
    +   `<div class="skh-pills">${fobChip(sku)}${landedChip(sku)}${cbmChip(sku)}${estLandedChip(sku)}${chanChip(sku)}</div>`
    + `</div>`
    + `<div class="skh-right">${statsHtml}${ytdHtml}</div>`
    + `</div></div></div></td></tr>`;
  const committed = ORDERS[sku.id] || EMPTY53;
  const proposed = proposedArr0;
  const rowOpen = `<tr data-skurow="${esc(sku.id)}"${open ? '' : ' class="acc-hide"'}>`;
  for (const def of ROWDEFS) {
    if (!visible.has(def.key)) continue;
    const src = def.key === 'orders' ? committed : def.key === 'proposed' ? proposed : r[def.key];
    h += `${rowOpen}<td class="lbl">${def.label}</td>`;
    let tot = 0;
    for (let w = 0; w < WEEKS; w++) {
      const v = src[w]; tot += v;
      const wk = w + 1;
      if (def.edit) {
        const cls = def.prop ? 'rb-prop' : (IMPORTED[sku.id] && IMPORTED[sku.id][w] !== v ? 'changed' : '');
        h += `<td class="ocell${def.prop ? ' rb-propcell' : ''}${wk === hw ? ' curwk' : ''}"><input data-sku="${esc(sku.id)}" data-r="${idx}" data-w="${w}" data-layer="${def.key}"
              class="${cls}" value="${v || ''}" inputmode="numeric"></td>`;
      } else {
        let cls = wk === hw ? 'curwk ' : '';
        let style = '';
        if (def.key === 'stock' && wk >= cur && v === 0 && r.forecast[w] > 0) cls += 'stockout ';
        else if (def.key === 'cover') { const cc = coverCellStyle(src, w); style = ` style="background:${cc.bg};color:${cc.color}"`; }
        if (def.key === 'forecast' && wk >= cur && seasonActiveFor(YEAR)) cls += 'modeled ';
        if (def.key === 'ly' && lyIsForecast(w)) cls += 'ly-fc ';
        if (Math.abs(v) < .5 && def.key !== 'cover') cls += 'zero';
        const ttl = def.key === 'ly' && lyIsForecast(w) ? ' title="Forecast — prior-year actual sales not yet available for this week"' : '';
        h += `<td class="${cls.trim()}"${style}${ttl} data-sku="${esc(sku.id)}" data-k="${def.key}" data-w="${w}">${def.fmt(v)}</td>`;
      }
    }
    if (def.key === 'forecast') {
      // weekly cells show the original forecast (tot = plan); the Total shows the
      // full-year OUTTURN (actuals banked so far + forecast for the rest), with the
      // plan total kept as a sub-value for tracking how the year is tracking vs target
      const outturn = r.forecastDisp.reduce((a, b) => a + b, 0);
      h += `<td class="tot tot-fc" data-sku="${esc(sku.id)}" data-k="forecast" data-w="T" title="Outturn = actual sales so far + forecast for the rest of the year">${fmtU(outturn)}<span class="tot-sub" title="Full-year sales forecast (plan) — compare with your target">plan ${fmtU(tot)}</span></td></tr>`;
    } else {
      const totFmt = def.edit ? fmtU : def.fmt;
      h += `<td class="tot" data-sku="${esc(sku.id)}" data-k="${def.key}" data-w="T">${totFmt(tot)}</td></tr>`;
    }
  }
  return h;
}

/* ---------------- supplier summary panel (visual subtotals) ---------------- */
function sparkline(arr, color, cur, label, fmt, ghost, extraAttrs) {
  const w = 160, h = 32, n = arr.length, gap = 0.8;
  const bw = (w - (n - 1) * gap) / n;
  const hasGhost = ghost && ghost.some(v => Math.abs(v) > 0);
  // scale bars AND the ghost line to a shared max so their heights are comparable
  const max = Math.max(1, ...arr.map(v => Math.abs(v)), ...(hasGhost ? ghost.map(v => Math.abs(v)) : []));
  let bars = '', hits = '';
  for (let i = 0; i < n; i++) {
    const bh = Math.abs(arr[i]) / max * (h - 1);
    const x = i * (bw + gap), y = h - bh, isCur = (i + 1) === cur;
    bars += `<rect${isCur ? ' class="spk-cur"' : ''} x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${bw.toFixed(2)}" height="${Math.max(0, bh).toFixed(2)}" `
          + `${isCur ? '' : `fill="${color}" opacity="0.8"`}></rect>`;
    // transparent full-height hit target so the whole week column is hoverable;
    // extraAttrs(i) can add this-week/LY/cumulative data for the rich tooltip
    hits += `<rect class="spk-hit" x="${(x - gap / 2).toFixed(2)}" y="0" width="${(bw + gap).toFixed(2)}" height="${h}" `
          + `fill="transparent" data-w="${i + 1}" data-l="${esc(label)}" data-v="${esc(fmt(arr[i]))}"${extraAttrs ? extraAttrs(i) : ''}></rect>`;
  }
  // faint dashed last-year line behind the bars (stock-holding card only)
  let ghostEl = '';
  if (hasGhost) {
    let gp = '';
    for (let i = 0; i < n; i++) {
      const gx = i * (bw + gap) + bw / 2, gy = h - Math.abs(ghost[i]) / max * (h - 1);
      gp += (i ? 'L' : 'M') + gx.toFixed(2) + ' ' + gy.toFixed(2) + ' ';
    }
    ghostEl = `<path class="spk-ghost" d="${gp}"/>`;
  }
  return `<svg class="spark" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none">${ghostEl}${bars}${hits}</svg>`;
}
// vs-LY delta pill: arrow (actual direction) + signed value + %. Colour = green when
// higher than last year, red when lower; `invert` flips only the COLOUR (not the arrow)
// for metrics where less is better (stock holding). `pre` optionally labels the basis
// (e.g. 'avg'). Returns '' when there's no LY figure.
function fmtSigned(v, fmt) { return (v >= 0 ? '+' : '−') + fmt(Math.abs(v)); }
function deltaChip(cur, prev, fmt, lyYear, pre, invert, extra) {
  if (prev == null || !isFinite(prev)) return '';
  const d = cur - prev;
  const flatEps = Math.max(Math.abs(prev) * 0.005, 1e-9);
  const flat = Math.abs(d) <= flatEps;
  const dir = flat ? 'flat' : (((d > 0) !== !!invert) ? 'up' : 'down');   // COLOUR class; invert flips only the colour, not the arrow
  const arrow = flat ? '▬' : (d > 0 ? '▲' : '▼');
  const pct = prev !== 0 ? (d / prev) * 100 : (Math.abs(cur) > 0 ? null : 0);
  const pctTxt = pct == null ? 'new' : (d >= 0 ? '+' : '−') + Math.abs(pct).toFixed(0) + '%';
  const preEl = pre ? `<span class="scd-pre">${pre}</span>` : '';
  return `<span class="sc-delta ${dir}" title="${pre ? pre + ' ' : ''}vs ${lyYear}: ${fmtSigned(d, fmt)} (${pctTxt})">`
    + `${preEl}${arrow} <span class="scd-val">${fmtSigned(d, fmt)}</span> <span class="scd-pct">${pctTxt}</span>${extra || ''}</span>`;
}
function supCard(label, total, arr, color, sub, cur, fmt, delta, ghost, extraAttrs) {
  return `<div class="sup-card" style="--c:${color}">
    <div class="sc-top"><span class="sc-label">${label}</span><span class="sc-total">${total}</span></div>
    ${sparkline(arr, color, cur, label, fmt, ghost, extraAttrs)}
    <div class="sc-sub"><span class="scs-txt">${sub}</span>${delta || ''}</div></div>`;
}
// Weekly sales line chart for the Summary tab. `sales` = realised/stock-capped sales-value
// £ (actuals where banked, else min(forecast,stock)×asp); `forecast` = uncapped forecast £
// (the plan) — drawn together on a SHARED £ axis so the gap shows how far actual sales are
// tracking off plan. `stock` = stock-holding £ on its OWN scale (different magnitude).
// `cont` = container count (committed+proposed CBM / container size) → baseline markers.
// The hover tooltip shows the real £ values + the actual-vs-forecast drift for past weeks.
function salesChart(sales, cont, stock, forecast, opt) {
  opt = opt || {};
  const W = 1000, H = opt.h || 130, padT = 8, padB = 22, n = sales.length;
  const innerH = H - padT - padB, base = padT + innerH;
  const hasFc = forecast && forecast.some(v => v > 0);
  const maxS = Math.max(1, ...sales, ...(hasFc ? forecast : []));
  const x = i => (i / (n - 1)) * W;
  const y = v => padT + (1 - v / maxS) * innerH;
  let line = '', area = `M0 ${base.toFixed(1)}`;
  for (let i = 0; i < n; i++) {
    const px = x(i).toFixed(1), py = y(sales[i]).toFixed(1);
    line += (i ? 'L' : 'M') + px + ' ' + py + ' ';
    area += ` L${px} ${py}`;
  }
  area += ` L${W} ${base.toFixed(1)} Z`;
  // forecast (plan) line — same £ axis as the sales line
  let fcLine = '';
  if (hasFc) {
    let fp = '';
    for (let i = 0; i < n; i++) fp += (i ? 'L' : 'M') + x(i).toFixed(1) + ' ' + y(forecast[i]).toFixed(1) + ' ';
    fcLine = `<path class="sl-fc" d="${fp}"/>`;
  }
  // stock-holding-value line on its own scale
  let stockLine = '';
  const hasStock = stock && stock.some(v => v > 0);
  if (hasStock) {
    const maxK = Math.max(1, ...stock);
    let sp = '';
    for (let i = 0; i < n; i++) sp += (i ? 'L' : 'M') + x(i).toFixed(1) + ' ' + (padT + (1 - stock[i] / maxK) * innerH).toFixed(1) + ' ';
    stockLine = `<path class="sl-stock" d="${sp}"/>`;
  }
  // faint quarter dividers + labels
  let grid = '', axis = `<text class="sl-ax" x="2" y="${H - 4}" text-anchor="start">W1</text>`;
  [13, 26, 39].forEach(q => {
    const gx = x(q);
    grid += `<line class="sl-grid" x1="${gx.toFixed(1)}" y1="${padT}" x2="${gx.toFixed(1)}" y2="${base.toFixed(1)}"/>`;
    axis += `<text class="sl-ax" x="${gx.toFixed(1)}" y="${H - 4}" text-anchor="middle">W${q + 1}</text>`;
  });
  axis += `<text class="sl-ax" x="${(W - 2).toFixed(1)}" y="${H - 4}" text-anchor="end">W${n}</text>`;
  // current-week marker — always the live date-based week (0 if this year ≠ current)
  let curLine = '';
  const cw = highlightWeek();
  if (cw >= 1 && cw <= n) { const cx = x(cw - 1).toFixed(1); curLine = `<line class="sl-cur" x1="${cx}" y1="${padT}" x2="${cx}" y2="${base.toFixed(1)}"/>`; }
  // container-arrival markers along the baseline (size ∝ containers that week)
  let marks = '';
  for (let i = 0; i < n; i++) {
    const c = cont[i];
    if (!(c > 0.05)) continue;
    const cx = x(i), r = (2 + Math.sqrt(c) * 1.6).toFixed(1);
    const lbl = c >= 0.95 ? Math.round(c) : '';
    marks += `<g class="sl-cont">`
      + `<circle cx="${cx.toFixed(1)}" cy="${base.toFixed(1)}" r="${r}"/>`
      + (lbl ? `<text class="sl-cont-lbl" x="${cx.toFixed(1)}" y="${(base - parseFloat(r) - 2).toFixed(1)}" text-anchor="middle">${lbl}</text>` : '')
      + `</g>`;
  }
  // transparent full-height hit columns drive the hover tooltip (added last = on top)
  const hw = W / (n - 1);
  let hits = '', cSales = 0, cFc = 0;
  for (let i = 0; i < n; i++) {
    cSales += sales[i];
    cFc += hasFc ? forecast[i] : 0;
    hits += `<rect class="sl-hit" x="${(x(i) - hw / 2).toFixed(1)}" y="0" width="${hw.toFixed(1)}" height="${H}" fill="transparent"`
      + ` data-w="${i + 1}" data-sales="${Math.round(sales[i])}" data-fc="${Math.round(hasFc ? forecast[i] : 0)}" data-stock="${Math.round(hasStock ? stock[i] : 0)}" data-cont="${cont[i].toFixed(2)}"`
      + ` data-csales="${Math.round(cSales)}" data-cfc="${Math.round(cFc)}"></rect>`;
  }
  return `<svg class="saleschart" viewBox="0 0 ${W} ${H}">${grid}<path class="sl-area" d="${area}"/>`
    + `<path class="sl-line" d="${line}"/>${fcLine}${stockLine}${curLine}${marks}${axis}${hits}</svg>`;
}
function supplierPanelHtml(t, supName) {
  if (!t) return '';
  const cur = highlightWeek() || SETTINGS.current_week;   // spark current-week marker = live date
  const sum = a => a.reduce((x, y) => x + y, 0);
  const avg = a => a.length ? sum(a) / a.length : 0;
  const peak = a => { const m = Math.max(...a); return { v: m, w: a.indexOf(m) + 1 }; };
  const sales = sum(t.sales), fob = sum(t.fob), units = sum(t.units), cbm = sum(t.cbm);
  const orderWks = t.units.filter(v => v > 0.5).length;
  const pSales = peak(t.sales), pShv = peak(t.shv);
  // last-year comparison for the same products (matched by code, like the export)
  const codes = new Set(M.skus.filter(s => s.supplier === supName).map(s => s.code));
  const ly = lyAggForCodes(codes);
  const lyN = LY_CACHE && LY_CACHE.year;
  const salesD = ly ? deltaChip(sales, sum(ly.sales), fmtGBP, lyN) : '';
  const fobD   = ly ? deltaChip(fob, sum(ly.fob), fmtGBP, lyN) : '';
  const unitsD = ly ? deltaChip(units, sum(ly.units), fmtU, lyN) : '';
  const cc = SETTINGS.container_cbm || 68;
  const dCont = ly ? (cbm - sum(ly.cbm)) / cc : 0;   // estimated container variance vs LY
  const cbmExtra = ly ? `<span class="scd-extra" title="≈ container variance vs ${lyN}">≈ ${(dCont >= 0 ? '+' : '−') + Math.abs(dCont).toFixed(1)} cont</span>` : '';
  const cbmD   = ly ? deltaChip(cbm, sum(ly.cbm), fmt1, lyN, null, false, cbmExtra) : '';
  const shvD   = ly ? deltaChip(avg(t.shv), avg(ly.shv), fmtGBP, lyN, 'avg', true) : '';   // avg holding (invert: less stock = green)
  const g = (GHOST_ON && ly) ? ly : null;    // last-year trend line on every card when enabled
  // Rich per-week tooltip data on every card: this-week TY + LY-same-week, plus running
  // cumulatives of both through the hovered week. The Stock Holding card additionally
  // carries the SALES cumulatives so the tooltip can judge stock-vs-sales health
  // (more sales on less stock = good, more stock on less sales = bad).
  const cum = a => { let s = 0; return a.map(v => (s += v)); };
  const mkAttrs = (tyArr, lyArr, fk, extra) => {
    const cty = cum(tyArr), cly = lyArr ? cum(lyArr) : null;
    return i => ` data-fk="${fk}" data-ty="${tyArr[i]}" data-cty="${cty[i]}"`
      + (lyArr ? ` data-ly="${lyArr[i]}" data-cly="${cly[i]}"` : '')
      + (extra ? extra(i) : '');
  };
  const salesTyCum = cum(t.sales), salesLyCum = ly ? cum(ly.sales) : null;
  const shvExtra = i => ` data-inv="1" data-csty="${salesTyCum[i]}"` + (salesLyCum ? ` data-csly="${salesLyCum[i]}"` : '');
  return [
    supCard('Sales', fmtGBP(sales), t.sales, '#0b5fff', `peak ${fmtGBP(pSales.v)} · W${pSales.w}`, cur, fmtGBP, salesD, g && g.sales, mkAttrs(t.sales, ly && ly.sales, 'gbp')),
    supCard('Order FOB', fmtGBP(fob), t.fob, '#16a085', `${fmtU(units)} units to order`, cur, fmtGBP, fobD, g && g.fob, mkAttrs(t.fob, ly && ly.fob, 'gbp')),
    supCard('Order Units', fmtU(units), t.units, '#8e44ad', orderWks ? `ordered across ${orderWks} week${orderWks > 1 ? 's' : ''}` : 'no orders planned', cur, fmtU, unitsD, g && g.units, mkAttrs(t.units, ly && ly.units, 'u')),
    supCard('Order CBM', fmt1(cbm), t.cbm, '#e67e22', `≈ ${(cbm / SETTINGS.container_cbm).toFixed(1)} containers`, cur, fmt1, cbmD, g && g.cbm, mkAttrs(t.cbm, ly && ly.cbm, '1')),
    supCard('Stock Holding', fmtGBP(pShv.v), t.shv, '#c0392b', `peak value · W${pShv.w}`, cur, fmtGBP, shvD, g && g.shv, mkAttrs(t.shv, ly && ly.shv, 'gbp', shvExtra)),
  ].join('');
}
function combinedSupAgg(supName) {     // supplier panel reflects committed + proposed
  const a = { sales: zeros(), units: zeros(), fob: zeros(), cbm: zeros(), shv: zeros() };
  for (const sku of M.skus) {
    if (sku.supplier !== supName) continue;
    const r = computeSku(sku, combinedOrder(sku.id));
    for (let w = 0; w < WEEKS; w++) {
      a.sales[w] += r.value[w]; a.units[w] += r.orders[w]; a.fob[w] += r.fobRow[w];
      a.cbm[w] += r.cbmRow[w]; a.shv[w] += r.shv[w];
    }
  }
  return a;
}
function refreshSupplierPanel() {
  const el = document.getElementById('sup-panel');
  if (el && currentSupplier) el.innerHTML = supplierPanelHtml(combinedSupAgg(currentSupplier), currentSupplier);
}
// Per-week order CBM for a supplier, split into committed (Committed Orders) and
// proposed (rebuy suggestions) — feeds the grid's weekly container-fill footer.
function supCbmByWeek(supName) {   // supName null/falsy = every supplier (whole-plan)
  const com = zeros(), prop = zeros();
  for (const sku of M.skus) {
    if ((supName && sku.supplier !== supName) || !(sku.cbm > 0)) continue;
    const o = ORDERS[sku.id], p = PROPOSED && PROPOSED.get(sku.id);
    for (let w = 0; w < WEEKS; w++) {
      if (o && o[w]) com[w] += o[w] * sku.cbm;
      if (p && p[w]) prop[w] += p[w] * sku.cbm;
    }
  }
  return { com, prop };
}
// Colour a weekly CBM cell by how close it is to filling a full container.
function cbmFillClass(cbm, CC) {
  if (cbm < 0.05) return '';
  const f = cbm / CC;
  if (f < 0.45) return 'cbm-lo';        // far from a full container
  if (f < 0.85) return 'cbm-mid';       // building up
  if (f <= 1.12) return 'cbm-full';     // ≈ a full container
  return 'cbm-over';                    // over a container (more than one)
}

/* ================= PO ↔ container matching (Phase 1: labelling) =================
   Links a supplier's planned/committed stock to live Purchase Orders + container
   bookings. WEBSA gives PO→product+outstanding-qty; Qlik gives PO→arrival date.
   Chain: app SKU → WEBSA PO(s) for that code → container arrival week → label. */
function poDataReady() { return !!(PO_WEBSA && PO_WEBSA.pos && PO_CONTAINERS && PO_CONTAINERS.dates); }

// ISO 'YYYY-MM-DD' → this year's plan week (1..53), or 0 if outside the year.
function isoToWeek(iso) {
  if (!iso) return 0;
  const base = (M && M.week1_start) ? M.week1_start : '2025-12-29';
  const [by, bm, bd] = base.split('-').map(Number);
  const [y, mo, d] = iso.split('-').map(Number);
  const wk = Math.floor((Date.UTC(y, mo - 1, d) - Date.UTC(by, bm - 1, bd)) / (7 * 86400000)) + 1;
  return (wk >= 1 && wk <= WEEKS) ? wk : 0;
}

// A PO's container arrival week(s) for the loaded year: [{week, booked, date, basis, status, container}].
// Delivery-to-CB preferred, else ETA UK Port; if no container booking, fall back to the
// PO's WEBSA Due Date (booked=false). Split/consolidated containers → multiple weeks.
function poArrivalWeeks(po) {
  const out = [];
  const rows = PO_CONTAINERS && PO_CONTAINERS.dates && PO_CONTAINERS.dates[po];
  if (rows && rows.length) {
    for (const r of rows) {
      const date = r.deliveryCB || r.etaPort;
      const wk = isoToWeek(date);
      if (wk) out.push({ week: wk, booked: true, date, basis: r.deliveryCB ? 'CB delivery' : 'UK port ETA', status: r.status || '', container: r.container || '' });
    }
  }
  if (!out.length) {
    const info = PO_WEBSA && PO_WEBSA.pos && PO_WEBSA.pos[po];
    const due = info && info.lines.map(l => l.due).find(Boolean);
    const wk = isoToWeek(due);
    if (wk) out.push({ week: wk, booked: false, date: due, basis: 'PO due date — not yet booked', status: '', container: '' });
  }
  const seen = new Set(), uniq = [];
  for (const e of out) if (!seen.has(e.week)) { seen.add(e.week); uniq.push(e); }
  return uniq;
}

// Outstanding POs that include any of a supplier's app SKUs (matched by code).
// Returns Map PO# -> { lines:[websa line], qty }.
function supplierPOs(supName) {
  const map = new Map();
  if (!poDataReady()) return map;
  const codes = new Set(M.skus.filter(s => s.supplier === supName).map(s => s.code));
  for (const po in PO_WEBSA.pos) {
    const lines = PO_WEBSA.pos[po].lines.filter(l => codes.has(l.code) && l.outstanding > 0);
    if (lines.length) map.set(po, { lines, qty: lines.reduce((a, l) => a + l.outstanding, 0) });
  }
  return map;
}

// Committed (raised-order) units per week for a supplier, from the editable ORDERS layer.
function supCommittedByWeek(supName) {
  const com = zeros();
  for (const sku of M.skus) {
    if (sku.supplier !== supName) continue;
    const o = ORDERS[sku.id];
    if (!o) continue;
    for (let w = 0; w < WEEKS; w++) if (o[w]) com[w] += o[w];
  }
  return com;
}
// Proposed-rebuy units per week for a supplier (includes "No PO" stock demoted by re-time).
function supProposedByWeek(supName) {
  const prop = zeros();
  if (!PROPOSED) return prop;
  for (const sku of M.skus) {
    if (sku.supplier !== supName) continue;
    const p = PROPOSED.get(sku.id);
    if (!p) continue;
    for (let w = 0; w < WEEKS; w++) if (p[w]) prop[w] += p[w];
  }
  return prop;
}

// Per-week PO schedule for a supplier: week(1..WEEKS) -> [{po, booked, status, container, basis, qty}].
function supplierPoSchedule(supName) {
  const byWeek = {};
  for (const [po, info] of supplierPOs(supName)) {
    for (const a of poArrivalWeeks(po)) {
      (byWeek[a.week] = byWeek[a.week] || []).push({ po, booked: a.booked, status: a.status, container: a.container, basis: a.basis, qty: info.qty });
    }
  }
  return byWeek;
}

// Build the diagnostic report (current+future weeks only — past stock has already landed).
function poDiagnostics() {
  const d = { notBooked: [], notRaised: [], notFoundWeeks: [], codesNotInApp: [],
              unparsed: (PO_CONTAINERS && PO_CONTAINERS.unparsed) || [],
              truncated: (PO_CONTAINERS && PO_CONTAINERS.truncated) || [] };
  if (!poDataReady()) return d;
  const cur = SETTINGS.current_week, appCodes = new Set(M.skus.map(s => s.code));
  const seenCode = new Set();
  for (const po in PO_WEBSA.pos) {
    const info = PO_WEBSA.pos[po];
    const appLines = info.lines.filter(l => l.outstanding > 0 && appCodes.has(l.code));
    for (const l of info.lines) {
      if (l.outstanding > 0 && !appCodes.has(l.code) && !seenCode.has(l.code)) { seenCode.add(l.code); d.codesNotInApp.push(l.code); }
    }
    if (!appLines.length) continue;
    const booked = !!(PO_CONTAINERS.dates[po] && PO_CONTAINERS.dates[po].length);
    if (!booked) {
      const wk = poArrivalWeeks(po)[0];
      d.notBooked.push({ po, supplier: info.supplier, qty: appLines.reduce((a, l) => a + l.outstanding, 0), week: wk ? wk.week : 0 });
    }
  }
  for (const s of M.suppliers) {
    const sched = supplierPoSchedule(s.name), com = supCommittedByWeek(s.name), supPos = supplierPOs(s.name);
    const weeks = [];
    for (let w = cur - 1; w < WEEKS; w++) if (com[w] > 0 && !(sched[w + 1] && sched[w + 1].length)) weeks.push(w + 1);
    if (weeks.length) d.notFoundWeeks.push({ supplier: s.name, weeks, hasPo: supPos.size > 0 });
    const poCodes = new Set();
    for (const [, info] of supPos) for (const l of info.lines) poCodes.add(l.code);
    for (const sku of M.skus) {
      if (sku.supplier !== s.name) continue;
      const o = ORDERS[sku.id];
      if (o && o.slice(cur - 1).some(v => v > 0) && !poCodes.has(sku.code)) d.notRaised.push({ supplier: s.name, code: sku.code });
    }
  }
  return d;
}

// Supplier names in the current sidebar sort order (so the review cycle follows the
// same order the user sees in the list).
function sortedSupplierNames() {
  const sups = M.suppliers.slice();
  if (sortMode === 'value-desc') sups.sort((a, b) => supForecastValue(b.name) - supForecastValue(a.name));
  else if (sortMode === 'value-asc') sups.sort((a, b) => supForecastValue(a.name) - supForecastValue(b.name));
  else if (sortMode === 'name') sups.sort((a, b) => a.name.localeCompare(b.name));
  return sups.map(s => s.name);
}

// Flat, sidebar-ordered list of committed weeks (current week on) with no PO match:
// [{supplier, week, units, hasPo}] — ordered by supplier (sidebar order) then week.
function poUnmatchedList() {
  if (!poDataReady()) return [];
  const cur = SETTINGS.current_week, out = [];
  for (const name of sortedSupplierNames()) {
    const sched = supplierPoSchedule(name), com = supCommittedByWeek(name), hasPo = supplierPOs(name).size > 0;
    for (let w = cur - 1; w < WEEKS; w++)
      if (com[w] > 0.5 && !(sched[w + 1] && sched[w + 1].length))
        out.push({ supplier: name, week: w + 1, units: Math.round(com[w]), hasPo });
  }
  return out;
}

/* ============ Phase 2: re-time committed arrivals to PO arrival weeks ============
   For each SKU on an outstanding PO, ensure the PO's arrival week holds the PO's
   outstanding qty, MOVING the shortfall in from that SKU's nearest mismatched future
   weeks (≥ current week, not themselves PO weeks) — a pure move (never invents or drops
   units, never touches past/delivered weeks). Arrival week = booked Qlik container week
   if booked, else the WEBSA PO due-date week — so the plan aligns to the PO/Qlik exports
   regardless of booking status. Stock not on any PO is left alone.
   scope: 'sup' = current supplier only; 'all' = whole year. Returns the proposed orders
   plus a per-SKU move list for the preview (nothing is written here). */
function computeRetime(scope) {
  const cur = SETTINGS.current_week;
  const res = { newOrders: {}, moves: [], skus: 0, placedUnits: 0, clearedUnits: 0, injectedUnits: 0, weeks: new Set() };
  if (!poDataReady()) return res;
  const inScope = sku => scope === 'all' ? true : sku.supplier === currentSupplier;
  // SKU code → PO arrival weeks (booked container week, else PO due-date fallback),
  // current week on, carrying each PO's outstanding qty. `poCodes` is every product on
  // an outstanding PO regardless of which YEAR it now arrives — so a product whose PO
  // has moved into another year (no in-year arrival) still gets its stale in-year
  // committed cleared, rather than lingering in an old arrival week.
  const arrivalByCode = {};
  const poCodes = new Set();
  for (const po in PO_WEBSA.pos) {
    for (const l of PO_WEBSA.pos[po].lines) if (l.outstanding > 0) poCodes.add(l.code);
    const weeks = poArrivalWeeks(po).filter(a => a.week >= cur);
    if (!weeks.length) continue;
    for (const l of PO_WEBSA.pos[po].lines) {
      if (l.outstanding <= 0) continue;
      (arrivalByCode[l.code] = arrivalByCode[l.code] || []).push(...weeks.map(a => ({ po, qty: l.outstanding, week: a.week })));
    }
  }
  for (const sku of M.skus) {
    if (!inScope(sku)) continue;
    if (!poCodes.has(sku.code)) continue;            // not on any outstanding PO → leave this product untouched
    const lines = arrivalByCode[sku.code] || [];     // empty ⇒ its PO now arrives in another year ⇒ clear in-year committed
    const before = (ORDERS[sku.id] || EMPTY53).slice();
    const target = {};                               // PO arrival week → outstanding qty
    for (const ln of lines) target[ln.week] = (target[ln.week] || 0) + ln.qty;
    // Set the committed row from the current week on to EXACTLY the PO schedule: each
    // PO arrival week holds its outstanding qty; every other future week is cleared to
    // zero (so old, superseded arrivals don't linger and double up). Past/delivered
    // weeks (< current week) are left untouched.
    const vec = before.slice();
    const placeWeeks = [], clearWeeks = [];
    let placed = 0, cleared = 0, injected = 0;
    for (let w = cur; w <= WEEKS; w++) {
      const want = target[w] || 0;
      const had = vec[w - 1] || 0;
      if (want > 0) {
        vec[w - 1] = want;
        placeWeeks.push({ week: w, qty: Math.round(want) });
        placed += want;
        if (want > had) injected += want - had;      // units the plan didn't previously hold
      } else if (had > 0) {
        vec[w - 1] = 0;
        clearWeeks.push({ week: w, units: Math.round(had) });
        cleared += had;
      }
    }
    if (!vec.some((v, i) => v !== before[i])) continue;   // no net change → skip
    res.newOrders[sku.id] = vec;
    res.skus++;
    res.placedUnits += placed; res.clearedUnits += cleared; res.injectedUnits += injected;
    placeWeeks.forEach(p => res.weeks.add(p.week));
    clearWeeks.forEach(c => res.weeks.add(c.week));
    const pos = [...new Set(lines.map(l => l.po))];
    res.moves.push({ id: sku.id, code: sku.code, name: sku.name || '', supplier: sku.supplier, placeWeeks, clearWeeks, pos, before, after: vec.slice() });
  }
  return res;
}
// Sticky grid footer: total order CBM (committed + proposed) per week, colour-coded by
// container fill, so you can see at a glance how close each week is to a full container.
function supCbmFooterHtml(supName, CC, hw) {
  const { com, prop } = supCbmByWeek(supName);
  let cells = '', totC = 0, totP = 0;
  for (let w = 0; w < WEEKS; w++) {
    const c = com[w], p = prop[w], v = c + p; totC += c; totP += p;
    const wk = w + 1;
    const tip = `Week ${wk}: ${fmt1(v)} CBM ≈ ${(v / CC).toFixed(2)} of a ${CC} CBM container`
      + (p > 0.05 ? ` — committed ${fmt1(c)} + proposed ${fmt1(p)}` : '');
    cells += `<td class="cbmf ${cbmFillClass(v, CC)}${wk === hw ? ' curwk' : ''}" data-cbmf="${w}" title="${esc(tip)}">${v > 0.05 ? fmt1(v) : ''}</td>`;
  }
  const tot = totC + totP;
  return `<tr class="cbm-foot"><td class="lbl">Order CBM <span class="cbmf-note">/ ${CC} = container</span></td>${cells}`
    + `<td class="tot" title="${fmt1(tot)} CBM total ≈ ${(tot / CC).toFixed(1)} containers">${fmt1(tot)}</td></tr>`;
}
function refreshCbmFooter() {
  const foot = document.querySelector('.grid .cbm-foot');
  if (!foot || !currentSupplier) return;
  const tmp = document.createElement('tbody');
  tmp.innerHTML = supCbmFooterHtml(currentSupplier, SETTINGS.container_cbm || 68, highlightWeek());
  foot.replaceWith(tmp.firstElementChild);
}
// Live-rebuild the sticky PO / "No PO" row so the flags follow committed/proposed edits:
// clears "No PO" when stock leaves a week, raises it when orders are keyed into a PO-less week.
function refreshPoRow() {
  const row = document.querySelector('.grid thead .po-row');
  if (!row || !currentSupplier || !poDataReady()) return;
  const tmp = document.createElement('table');
  tmp.innerHTML = `<thead>${poRowHtml(currentSupplier)}</thead>`;
  const fresh = tmp.querySelector('.po-row');
  if (fresh) row.replaceWith(fresh);
}
// Live-update the week-header proposed-order highlight for the SKUs currently on screen.
function refreshPropWeekHeaders() {
  const ths = document.querySelectorAll('.grid thead th');
  if (!ths.length || !PROPOSED) return;
  const ids = new Set();
  document.querySelectorAll('.grid tbody input[data-layer="proposed"]').forEach(inp => ids.add(inp.dataset.sku));
  const set = new Set();
  for (const id of ids) {
    const arr = PROPOSED.get(id);
    if (!arr) continue;
    for (let w = 0; w < WEEKS; w++) if (arr[w]) set.add(w + 1);
  }
  // ths[0] is the "Week" label; ths[1..WEEKS] are W1..W53; last is Total.
  for (let w = 1; w <= WEEKS; w++) {
    const th = ths[w];
    if (!th) continue;
    const on = set.has(w);
    th.classList.toggle('has-prop', on);
    if (on && !th.title) th.title = 'Proposed rebuy orders present in this week';
    else if (!on && th.title) th.removeAttribute('title');
  }
}
function initSparkTooltip() {
  const tip = document.createElement('div');
  tip.id = 'spark-tip';
  tip.style.display = 'none';
  document.body.appendChild(tip);
  document.addEventListener('mousemove', e => {
    const t = e.target, cl = t.classList;
    const spk = cl && cl.contains('spk-hit') ? t : null;
    const sl = cl && cl.contains('sl-hit') ? t : null;
    // PO breakdown shows on the PO row cell AND on the matching Committed Orders cell (same week)
    let poTip = null;
    if (!spk && !sl && t.closest) {
      const poc = t.closest('.po-cell');
      if (poc) poTip = PO_WEEK_TIP[+poc.dataset.pw];   // PO row chip: full week breakdown (all POs/products)
      else {
        const oc = t.closest('td.ocell'), inp = oc && oc.querySelector('input');
        if (inp && inp.dataset.layer === 'orders') {   // committed cell: just this product's POs that week
          const sku = skuById.get(inp.dataset.sku);
          if (sku) poTip = poProductTipHtml(+inp.dataset.w + 1, sku.code);
        }
      }
    }
    if (!spk && !sl && !poTip) { if (tip.style.display !== 'none') tip.style.display = 'none'; return; }
    if (poTip) {
      tip.innerHTML = poTip;
    } else if (spk) {
      const w = +spk.dataset.w, d = spk.dataset;
      if (!d.fk) {   // plain sparkline (no rich data): label + value only
        tip.innerHTML = `<div class="st-wk">Week ${w} · w/c ${weekDate(w)}</div>`
                      + `<div class="st-val">${esc(d.l)}: <b>${esc(d.v)}</b></div>`;
      } else {
        // supplier KPI card: this week vs LY same week + cumulatives-to-week vs LY
        const fmt = { gbp: fmtGBP, u: fmtU, 1: fmt1 }[d.fk] || fmtGBP;
        const lyYr = (LY_CACHE && LY_CACHE.year) || 'LY';
        const ty = +d.ty, cty = +d.cty;
        const hasLy = d.ly != null;
        const ly = +d.ly, cly = +d.cly;
        const pct = (a, b) => b ? `${a - b >= 0 ? '+' : '−'}${Math.abs(((a - b) / b) * 100).toFixed(0)}%` : (a ? 'new' : '0%');
        const varTxt = (a, b) => `${fmtSigned(a - b, fmt)} (${pct(a, b)})`;
        let h = `<div class="st-wk">Week ${w} · w/c ${weekDate(w)} — ${esc(d.l)}</div>`
              + `<div class="st-val">This week: <b>${fmt(ty)}</b>`
              + (hasLy ? ` <span class="st-lywk">· ${lyYr} same week: ${fmt(ly)}</span>` : '') + `</div>`
              + `<div class="st-cum-head">Cumulative to W${w}</div>`
              + `<div class="st-cum">This year: <b>${fmt(cty)}</b></div>`;
        if (hasLy) {
          h += `<div class="st-cum">${lyYr}: <b>${fmt(cly)}</b></div>`;
          if (d.inv) {
            // Stock Holding: judge the variance against SALES over the same weeks —
            // more sales on less stock = green, more stock on less sales = red
            const csty = +d.csty, csly = +d.csly;
            const stockUp = cty > cly, salesUp = csty >= csly;
            const cls = (!stockUp && salesUp) ? 'st-var-up' : (stockUp && !salesUp) ? 'st-var-down' : 'st-var-flat';
            const note = (!stockUp && salesUp) ? 'less stock, more sales ✓'
                       : (stockUp && !salesUp) ? 'more stock, less sales ✗'
                       : stockUp ? 'more stock, more sales' : 'less stock, less sales';
            h += `<div class="st-cum st-cum-var">Stock vs ${lyYr}: <b class="${cls}">${varTxt(cty, cly)}</b></div>`
               + `<div class="st-cum">Sales vs ${lyYr}: <b>${varTxt(csty, csly)}</b></div>`
               + `<div class="st-cum st-note ${cls}">${note}</div>`;
          } else {
            const cls = cty >= cly ? 'st-var-up' : 'st-var-down';
            h += `<div class="st-cum st-cum-var">Variance: <b class="${cls}">${varTxt(cty, cly)}</b></div>`;
          }
        }
        tip.innerHTML = h;
      }
    } else if (sl.dataset.kind === 'cap') {
      const w = +sl.dataset.w, pallet = +sl.dataset.pallet, racking = +sl.dataset.racking, stillage = +sl.dataset.stillage;
      const ruse = +sl.dataset.ruse, suse = +sl.dataset.suse;
      const rcls = ruse > 100 ? 'st-var-down' : 'st-var-up', scls = suse > 100 ? 'st-var-down' : 'st-var-up';
      tip.innerHTML = `<div class="st-wk">Week ${w} · w/c ${weekDate(w)}</div>`
        + `<div class="st-val st-sales">Pallets &amp; racking: <b>${fmtU(pallet + racking)}</b> spaces · <b class="${rcls}">${ruse}%</b> of capacity</div>`
        + (racking ? `<div class="st-val"><span class="st-lywk">${fmtU(pallet)} pallet + ${fmtU(racking)} racking</span></div>` : '')
        + `<div class="st-val st-stock">Stillages: <b>${fmtU(stillage)}</b> spaces · <b class="${scls}">${suse}%</b> of capacity</div>`
        + (ruse > 100 || suse > 100 ? `<div class="st-cum st-note st-var-down">over capacity this week ✗</div>` : '');
    } else {
      const w = +sl.dataset.w, sales = +sl.dataset.sales, fc = +sl.dataset.fc, stock = +sl.dataset.stock, cont = +sl.dataset.cont;
      const csales = +sl.dataset.csales, cfc = +sl.dataset.cfc;
      const past = w < SETTINGS.current_week;     // weeks with banked actuals
      const variance = (cur, plan) => { const d = cur - plan, p = plan ? (d / plan) * 100 : 0, s = d >= 0 ? '+' : '−'; return `${s}${fmtGBP(Math.abs(d))} (${s}${Math.abs(p).toFixed(0)}%)`; };
      let drift = '';
      if (fc > 0 && past) drift = `<div class="st-val st-drift">vs forecast: <b>${variance(sales, fc)}</b></div>`;
      // cumulative (year-to-date through the hovered week)
      let cum = `<div class="st-cum-head">Cumulative to W${w}</div>`
        + `<div class="st-cum">${past ? 'Actual' : 'Sales (inc. projected)'}: <b>${fmtGBP(csales)}</b></div>`;
      if (cfc > 0) cum += `<div class="st-cum">Forecast: <b>${fmtGBP(cfc)}</b></div>`
        + `<div class="st-cum st-cum-var">Variance: <b>${variance(csales, cfc)}</b></div>`;
      tip.innerHTML = `<div class="st-wk">Week ${w} · w/c ${weekDate(w)}</div>`
        + `<div class="st-val st-sales">${past ? 'Actual' : 'Projected'} sales: <b>${fmtGBP(sales)}</b></div>`
        + (fc > 0 ? `<div class="st-val st-fc">Forecast sales: <b>${fmtGBP(fc)}</b></div>` : '')
        + drift
        + (stock > 0 ? `<div class="st-val st-stock">Stock holding: <b>${fmtGBP(stock)}</b></div>` : '')
        + (cont > 0.05 ? `<div class="st-val st-cont">Containers: <b>${cont.toFixed(1)}</b></div>` : '')
        + cum;
    }
    tip.style.display = 'block';
    const r = tip.getBoundingClientRect(), pad = 14;
    let x = e.clientX + pad, y = e.clientY + pad;
    if (x + r.width > window.innerWidth) x = e.clientX - r.width - pad;
    if (y + r.height > window.innerHeight) y = e.clientY - r.height - pad;
    tip.style.left = x + 'px';
    tip.style.top = y + 'px';
  });
}

function renderPlan() {
  const main = document.getElementById('main');
  const sup = supByName.get(currentSupplier);
  if (!sup) { main.innerHTML = '<div class="empty">Select a supplier</div>'; return; }
  const term = searchTerm.toLowerCase();
  const allSkus = M.skus.filter(k => k.supplier === sup.name);
  const liveN = allSkus.filter(k => k.status === 'Live').length;
  const notN = allSkus.filter(k => k.status === 'Not Live').length;
  const unkN = allSkus.length - liveN - notN;

  let skus = allSkus;
  if (statusFilter === 'live') skus = skus.filter(k => k.status === 'Live');
  else if (statusFilter === 'notlive') skus = skus.filter(k => k.status === 'Not Live');
  if (term) {
    const hit = skus.filter(k => k.code.toLowerCase().includes(term) || (k.name || '').toLowerCase().includes(term));
    if (hit.length) skus = hit;
  }
  const meta = [
    sup.number ? `No. <b>${esc(sup.number)}</b>` : null,
    sup.contact ? `Contact <b>${esc(sup.contact)}</b>` : null,
    sup.port ? `Port <b>${esc(sup.port)}</b>` : null,
    sup.origin ? `Group <b>${esc(sup.origin)}</b>` : null,
    `<b>${allSkus.length}</b> products · <b class="c-live">${liveN} Live</b> · <b class="c-notlive">${notN} Not Live</b>${unkN ? ` · ${unkN} unknown` : ''}`,
    sup.email ? `<b>${esc(sup.email)}</b>` : null,
  ].filter(Boolean).join(' &nbsp;·&nbsp; ');
  // product ordering within the supplier: original import order / sales value / A–Z
  if (supSort === 'value') {
    const val = k => { const rr = RES.get(k.id); return rr ? rr.value.reduce((a, b) => a + b, 0) : 0; };
    skus = skus.slice().sort((a, b) => val(b) - val(a));
  } else if (supSort === 'az') {
    skus = skus.slice().sort((a, b) => a.code.localeCompare(b.code));
  }
  accForceOpen = !!term;   // searching → matching products render expanded

  const sfOpts = { all: `All (${allSkus.length})`, live: `Live (${liveN})`, notlive: `Not Live (${notN})` };
  const allOn = ROWDEFS.every(d => visible.has(d.key));
  const statusDD = `<details class="dd" id="dd-status"><summary>Status: <b>${sfOpts[statusFilter] || sfOpts.all}</b> ▾</summary><div class="dd-pop">`
    + Object.keys(sfOpts).map(k => `<label><input type="radio" name="dd-sf" class="sfbtn" data-sf="${k}"${statusFilter === k ? ' checked' : ''}> ${sfOpts[k]}</label>`).join('')
    + `</div></details>`;
  const rowsDD = `<details class="dd" id="dd-rows"><summary>Rows: <b>${visible.size} of ${ROWDEFS.length}</b> ▾</summary><div class="dd-pop">`
    + `<label class="dd-all"><input type="checkbox" class="chip" data-row="__all__"${allOn ? ' checked' : ''}> <b>All rows</b></label><hr>`
    + ROWDEFS.map(d => `<label><input type="checkbox" class="chip" data-row="${d.key}"${visible.has(d.key) ? ' checked' : ''}> ${d.label}</label>`).join('')
    + `</div></details>`;
  const sortSeg = `<span class="seg" title="Product order within this supplier">`
    + [['orig', 'Original'], ['value', 'Value'], ['az', 'A–Z']].map(([k, l]) =>
      `<button class="srtbtn${supSort === k ? ' on' : ''}" data-srt="${k}">${l}</button>`).join('') + `</span>`;

  const t = combinedSupAgg(sup.name);
  const CC = SETTINGS.container_cbm || 68, rbMode = (SETTINGS.rebuy && SETTINGS.rebuy.mode) || 'full';
  const pt = proposedSupTotals(sup.name), gt = proposedGrandTotals();
  // scope governs the Run rebuy / Clear / Commit buttons: this supplier only, or whole year
  const scoped = rebuyScope === 'sup' ? pt : gt;
  const scopeWord = rebuyScope === 'sup' ? 'this supplier' : 'the whole year';
  const sumLbl = rebuyScope === 'sup' ? esc(titleCase(sup.name)) : 'whole plan';
  const supProd = M.skus.filter(s => s.supplier === sup.name && PROPOSED.get(s.id) && PROPOSED.get(s.id).some(v => v)).length;
  const scopeProd = rebuyScope === 'sup' ? supProd : gt.skuCount;
  let body = '';
  skus.forEach((k, i) => body += skuRowsHtml(k, i));
  if (!skus.length)
    body = `<tr class="skuhead"><td colspan="${WEEKS + 2}"><div class="skuhead-inner" style="color:var(--dim);font-weight:400">No ${statusFilter === 'live' ? 'Live' : statusFilter === 'notlive' ? 'Not Live' : 'matching'} products for this supplier.</div></td></tr>`;

  const se = SEASON();
  const mode = seasonModeFor(YEAR);
  let seasonBanner = '';
  if (mode === 'target') {
    const tgt = targetValueFor(YEAR);
    seasonBanner = `<div id="season-banner">🎯 <b>Target sales mode</b> — ${YEAR} forecast scaled to a <b>${fmtGBPk(tgt)}</b> sales target (modelled demand <b>${fmtGBPk(forecastDemandTotal())}</b>), distributed by recency-weighted history + seasonality. <a id="season-edit-link">adjust</a> · <a id="target-off-link">turn off</a></div>`;
  } else if (mode === 'seasonality') {
    const gp = growthPctFor(YEAR);
    const growthNote = gp ? ` · <b>${gp > 0 ? '+' : ''}${gp}%</b> YoY growth applied` : '';
    seasonBanner = `<div id="season-banner">🌦️ <b>Seasonality model active</b> (${YEAR}) — future Sales Forecast re-modelled at ${Math.round(seasonStrengthFor(YEAR) * 100)}% strength${growthNote}${se.useWeather ? ' · ' + esc(weatherSummary()) : ''}. <a id="season-revert-link">Revert to original</a> · <a id="season-edit-link">adjust</a></div>`;
  }
  const allOpen = skus.length > 0 && skus.every(k => accForceOpen || openSkus.has(k.id));
  const rebuyDD = `<details class="dd" id="dd-rebuy"><summary>⟳ Rebuy: <b>${rbMode === 'full' ? `Full ${CC}` : 'Partial'}</b> · ${rebuyScope === 'sup' ? 'supplier' : 'year'} ▾</summary><div class="dd-pop">`
    + `<div class="dd-sec">Container basket</div>`
    + `<label><input type="radio" name="dd-rm" class="rb-mode" data-mode="full"${rbMode === 'full' ? ' checked' : ''}> Full ${CC} CBM</label>`
    + `<label><input type="radio" name="dd-rm" class="rb-mode" data-mode="partial"${rbMode === 'partial' ? ' checked' : ''}> Allow partial</label>`
    + `<hr><div class="dd-sec">Run / clear / commit apply to</div>`
    + `<label><input type="radio" name="dd-rs" class="rb-scope" data-scope="sup"${rebuyScope === 'sup' ? ' checked' : ''}> This supplier</label>`
    + `<label><input type="radio" name="dd-rs" class="rb-scope" data-scope="all"${rebuyScope === 'all' ? ' checked' : ''}> Whole year</label>`
    + `</div></details>`;
  const moreDD = `<details class="dd dd-right" id="dd-more"><summary title="Export / import / container dates / revert">⋯</summary><div class="dd-pop dd-menu">`
    + `<button id="btn-export-sup" title="Download this supplier's order-planning form, including proposed rebuys (Excel)">&#x2913; Export to Excel</button>`
    + `<button id="btn-export-multi" title="Export all suppliers or a selection — one form each, delivered as a single .zip">&#x2913; Export multiple…</button>`
    + `<button id="btn-import-sup" title="Upload an edited form to confirm those orders">&#x2911; Import from Excel</button>`
    + (poDataReady() ? `<button id="btn-sup-retime" title="Re-time this supplier's committed stock arrivals to their PO arrival weeks (booked Qlik container date, else PO due date)">&#8635; Apply Qlik Container Dates</button>${PO_APPLY_UNDO ? `<button id="btn-sup-retime-undo" title="Undo the last container re-time (${esc(PO_APPLY_UNDO.label)})">&#8624; Undo re-time</button>` : ''}` : '')
    + `<hr><button id="btn-revert-orders" class="danger" title="Revert all ${YEAR} orders — choose the most recent saved configuration or the original imported file">&#8634; Revert orders</button>`
    + `</div></details>`;
  main.innerHTML = `
    ${seasonBanner}
    <div id="totals-band">${totalsHtml()}</div>
    <div id="plan-toolbar">
      <div class="ptb-row1"><h1>${esc(sup.name)}</h1><div class="meta">${meta}</div></div>
      <div class="ptb-row2">
        ${sortSeg}
        ${statusDD}
        ${rowsDD}
        ${rebuyDD}
        <button id="btn-acc-all" class="tbtn" title="${allOpen ? 'Collapse every product to its summary strip' : 'Expand every product to its full weekly grid'}">${allOpen ? '⌃ Collapse all' : '⌄ Expand all'}</button>
        <span class="rb-bar-sum">${sumLbl}: <b>${scopeProd}</b> products · <b>${(scoped.cbm / CC).toFixed(1)}</b> containers · <b>${fmtGBPk(scoped.fob)}</b> FOB proposed</span>
        <span class="spacer"></span>
        ${REBUY_STALE ? `<span class="rb-stale" title="Suggestions are only ever built when you click Run rebuy — nothing has been changed for you.">⚠ suggestions may be out of date</span>` : ''}
        <button id="btn-run-rebuy" class="tbtn${REBUY_STALE ? ' rb-nudge' : ''}" title="(Re)build the Proposed Rebuy row for ${scopeWord} from the current committed orders.">&#8635; Run rebuy</button>
        <button id="btn-clear-prop" class="tbtn"${scoped.units ? '' : ' disabled'} title="Remove proposed rebuy suggestions for ${scopeWord}">Clear</button>
        <button id="btn-commit-all" class="tbtn primary"${scoped.units ? '' : ' disabled'} title="Add ${scopeWord}'s proposed rebuys (teal row) to Committed Orders as confirmed orders">&#10003; Commit ${fmtU(scoped.units)}</button>
        ${moreDD}
        <input type="file" id="import-file" accept=".xlsx" hidden>
      </div>
    </div>
    ${LY_CACHE ? `<div id="sup-panel-head"><label class="ghost-toggle" title="Overlay each card with a faint dashed ${LY_CACHE.year} trend line"><input type="checkbox" id="ghost-chk"${GHOST_ON ? ' checked' : ''}><span>${LY_CACHE.year} trend</span></label></div>` : ''}
    <div id="sup-panel">${supplierPanelHtml(t, sup.name)}</div>
    <div class="gridwrap"><table class="grid"><thead>${headerRow(proposedWeekSet(skus))}${poRowHtml(sup.name)}</thead>
      <tbody>${body}</tbody>
      <tfoot>${supCbmFooterHtml(sup.name, CC, highlightWeek())}</tfoot></table></div>`;
  if (reopenDD) { const dd = document.getElementById(reopenDD); if (dd) dd.open = true; reopenDD = null; }

  const ghostChk = main.querySelector('#ghost-chk');
  if (ghostChk) ghostChk.addEventListener('change', e => {
    GHOST_ON = e.target.checked; savePref('tp_ghost', GHOST_ON); refreshSupplierPanel();
  });
  main.querySelectorAll('.chip').forEach(c => c.addEventListener('change', () => {
    const k = c.dataset.row;
    if (k === '__all__') {
      const allOn = ROWDEFS.every(d => visible.has(d.key));
      visible = new Set(allOn ? DEFAULT_ROWS : ROWDEFS.map(d => d.key));
    } else {
      visible.has(k) ? visible.delete(k) : visible.add(k);
    }
    savePref('tp_visibleRows', [...visible]);
    reopenDD = 'dd-rows';   // keep the Rows popover open while ticking
    renderPlan();
  }));
  main.querySelectorAll('.sfbtn').forEach(b => b.addEventListener('change', () => {
    statusFilter = b.dataset.sf;
    savePref('tp_statusFilter', statusFilter);
    renderPlan();
  }));
  main.querySelectorAll('.srtbtn').forEach(b => b.addEventListener('click', () => {
    supSort = b.dataset.srt;
    savePref('tp_supSort', supSort);
    renderPlan();
  }));
  // product accordion: only the image, code, name and caret toggle the grid rows —
  // the rest of the banner (chips, stats, YTD) is inert so nothing mis-clicks
  main.querySelectorAll('tr.skuhead[data-acc]').forEach(hr => hr.addEventListener('click', e => {
    if (!e.target.closest('.acc-hit, .acc-car')) return;
    const id = hr.dataset.acc;
    const nowOpen = !hr.classList.contains('acc-open');
    hr.classList.toggle('acc-open', nowOpen);
    if (nowOpen) openSkus.add(id); else openSkus.delete(id);
    savePref('tp_openSkus', [...openSkus]);
    main.querySelectorAll(`tr[data-skurow="${CSS.escape(id)}"]`).forEach(rr => rr.classList.toggle('acc-hide', !nowOpen));
  }));
  document.getElementById('btn-acc-all').addEventListener('click', () => {
    const shown = skus.map(k => k.id);
    const everyOpen = shown.length && shown.every(id => accForceOpen || openSkus.has(id));
    if (everyOpen) shown.forEach(id => openSkus.delete(id));
    else shown.forEach(id => openSkus.add(id));
    savePref('tp_openSkus', [...openSkus]);
    renderPlan();
  });
  document.getElementById('btn-export-sup').addEventListener('click', () => exportSupplier(sup.name));
  document.getElementById('btn-export-multi').addEventListener('click', openExportDialog);
  document.getElementById('btn-import-sup').addEventListener('click', () => document.getElementById('import-file').click());
  document.getElementById('btn-revert-orders').addEventListener('click', openRevertDialog);
  document.getElementById('import-file').addEventListener('change', e => importSupplier(e, sup.name));
  { const rb = document.getElementById('btn-sup-retime'); if (rb) rb.addEventListener('click', applySupplierRetime); }
  { const ub = document.getElementById('btn-sup-retime-undo'); if (ub) ub.addEventListener('click', undoRetime); }
  const scopeArg = () => rebuyScope === 'sup' ? sup.name : null;
  const scopeTag = () => rebuyScope === 'sup' ? titleCase(sup.name) : 'whole year';
  main.querySelectorAll('.rb-scope').forEach(b => b.addEventListener('change', () => {
    rebuyScope = b.dataset.scope; savePref('tp_rebuyScope', rebuyScope);
    reopenDD = 'dd-rebuy'; renderPlan();
  }));
  document.getElementById('btn-commit-all').addEventListener('click', () => { labelNextSave(`Commit rebuy · ${scopeTag()}`); commitRebuy(scopeArg()); });
  document.getElementById('btn-run-rebuy').addEventListener('click', () => {
    labelNextSave(`Run rebuy · ${scopeTag()}`);
    REBUY_STALE = false;
    resetProposed(scopeArg()); computeAll(); markDirty(); renderPlan();
    document.getElementById('save-status').textContent = `Rebuy re-run · ${scopeTag()}`;
  });
  document.getElementById('btn-clear-prop').addEventListener('click', () => {
    labelNextSave(`Clear suggestions · ${scopeTag()}`);
    clearProposed(scopeArg()); computeAll(); markDirty(); renderPlan();
    document.getElementById('save-status').textContent = `Suggestions cleared · ${scopeTag()}`;
  });
  main.querySelectorAll('.rb-mode').forEach(b => b.addEventListener('change', () => {
    // remember the container basket, but don't re-run anything: it applies on the next
    // "Run rebuy" click, so switching the setting can't wipe suggestions already curated
    SETTINGS.rebuy.mode = b.dataset.mode; markDirty();
    REBUY_STALE = true;
    reopenDD = 'dd-rebuy'; renderPlan();
  }));
  const rv = document.getElementById('season-revert-link');
  if (rv) rv.addEventListener('click', () => { delete seasonYears()[String(YEAR)]; markDirty(); applySeasonality(); });
  const toff = document.getElementById('target-off-link');
  if (toff) toff.addEventListener('click', () => { delete seasonYears()[String(YEAR)]; markDirty(); applySeasonality(); });
  const ed = document.getElementById('season-edit-link');
  if (ed) ed.addEventListener('click', openSettings);
  main.querySelectorAll('.sku-explain').forEach(b => b.addEventListener('click', () => explainSku(b.dataset.sku)));
  main.querySelectorAll('.asp-chip').forEach(b => b.addEventListener('click', () => editAspInline(b.dataset.aspSku)));
  main.querySelectorAll('.cost-chip:not(.cbm-chip)').forEach(b => b.addEventListener('click', () => editCostInline(b.dataset.costSku, b.dataset.costK)));
  main.querySelectorAll('.cbm-chip').forEach(b => b.addEventListener('click', () => openCartonDialog(b.dataset.cbmSku)));
  main.querySelectorAll('.chan-chip').forEach(b => b.addEventListener('click', () => openChannelDialog(b.dataset.chanSku)));
  bindOrderInputs(main);
  const cw = main.querySelector('thead th.curwk');
  if (cw) cw.scrollIntoView({ block: 'nearest', inline: 'center' });
}

/* ---- editing ---- */
// Highlight an order/rebuy input when the qty isn't a whole multiple of the product's
// pack size (non-blocking — just flags it, with the nearest valid quantities).
function flagPackSize(inp, id, v) {
  const sku = skuById.get(id);
  const ps = sku && Math.round(+sku.pack_size || 0);
  const bad = ps > 1 && v > 0 && (v % ps !== 0);
  inp.classList.toggle('pack-warn', bad);
  if (bad) {
    const lo = Math.floor(v / ps) * ps, hi = lo + ps;
    inp.title = `⚠ ${v} isn't a multiple of the pack size (${ps}) — nearest ${lo ? lo + ' or ' : ''}${hi}`;
  } else if (inp.title) {
    inp.title = '';
  }
}
function bindOrderInputs(root) {
  root.querySelectorAll('td.ocell input').forEach(inp => {
    inp.addEventListener('change', () => commitEdit(inp));
    inp.addEventListener('keydown', e => orderKeyNav(e, inp));
    inp.addEventListener('paste', e => orderPaste(e, inp));
    inp.addEventListener('focus', () => inp.select());
    flagPackSize(inp, inp.dataset.sku, +inp.value || 0);   // flag any pre-existing non-multiples on render
  });
}
function commitEdit(inp) {
  const id = inp.dataset.sku, w = +inp.dataset.w, layer = inp.dataset.layer || 'orders';
  let v = parseFloat(String(inp.value).replace(/[^0-9.\-]/g, ''));
  if (!isFinite(v) || v < 0) v = 0;
  v = Math.round(v);
  inp.value = v || '';
  flagPackSize(inp, id, v);          // warn if the qty isn't a whole multiple of the pack size
  if (layer === 'proposed') {
    if (!PROPOSED) PROPOSED = new Map();
    if (!PROPOSED.has(id)) PROPOSED.set(id, zeros());
    if (PROPOSED.get(id)[w] === v) return;
    PROPOSED.get(id)[w] = v;
    computeAll();               // refresh combined committed+proposed whole-plan totals
    refreshSkuCells(id);        // projection (committed + proposed) updates live
    refreshSupplierPanel();
    refreshCbmFooter();         // weekly container-fill footer updates live
    refreshPropWeekHeaders();   // week-number highlight follows the edit live
    refreshPoRow();             // "No PO" flags follow the edit live
    fillTotals();               // whole-plan band proposed breakdown updates live
    const sk = skuById.get(id); if (sk) updateSupWarnIcon(sk.supplier);   // sidebar caution flag
    markDirty();                // saves proposed alongside orders/settings
    return;
  }
  if (!ORDERS[id]) ORDERS[id] = zeros();
  if (ORDERS[id][w] === v) return;
  ORDERS[id][w] = v;
  inp.classList.toggle('changed', IMPORTED[id] && IMPORTED[id][w] !== v);
  computeAll();                 // committed totals + sidebar KPIs
  refreshSkuCells(id);
  refreshSupplierPanel();
  refreshCbmFooter();           // weekly container-fill footer updates live
  refreshPoRow();               // "No PO" flags follow the edit live
  fillTotals();
  updateSidebarKpi(skuById.get(id).supplier);
  markDirty();
}
function refreshSkuCells(id) {
  const sku = skuById.get(id);
  const r = computeSku(sku, combinedOrder(id)), cur = SETTINGS.current_week, hw = highlightWeek();
  const committed = ORDERS[id] || EMPTY53, proposed = (PROPOSED && PROPOSED.get(id)) || EMPTY53;
  document.querySelectorAll(`td[data-sku="${CSS.escape(id)}"]`).forEach(td => {
    const k = td.dataset.k; if (!k) return;
    const def = ROWDEFS.find(d => d.key === k);
    const src = k === 'orders' ? committed : k === 'proposed' ? proposed : r[k];
    if (td.dataset.w === 'T') {
      if (k === 'forecast') {   // Total = outturn (actuals + remaining forecast) + plan sub-value
        const outturn = r.forecastDisp.reduce((a, b) => a + b, 0);
        td.innerHTML = `${fmtU(outturn)}<span class="tot-sub" title="Full-year sales forecast (plan) — compare with your target">plan ${fmtU(src.reduce((a, b) => a + b, 0))}</span>`;
      } else {
        td.textContent = (def.edit ? fmtU : def.fmt)(src.reduce((a, b) => a + b, 0));
      }
      return;
    }
    if (def.edit) return;       // edit-row cells are inputs, not refreshed here
    const w = +td.dataset.w, v = src[w], wk = w + 1;
    td.textContent = def.fmt(v);
    td.className = '';
    td.style.background = ''; td.style.color = '';
    if (wk === hw) td.classList.add('curwk');
    if (k === 'stock' && wk >= cur && v === 0 && r.forecast[w] > 0) td.classList.add('stockout');
    else if (k === 'cover') { const cc = coverCellStyle(src, w); td.style.background = cc.bg; td.style.color = cc.color; }
    if (k === 'forecast' && wk >= cur && seasonActiveFor(YEAR)) td.classList.add('modeled');
    if (k === 'ly' && lyIsForecast(w)) td.classList.add('ly-fc');
    if (Math.abs(v) < .5 && k !== 'cover') td.classList.add('zero');
  });
}
function updateSidebarKpi(supName) {           // refresh one supplier's stocked-in % chip
  const t = AGG.bySup.get(supName); if (!t) return;
  const item = [...document.querySelectorAll('.sup-item')].find(el => el.dataset.sup === supName);
  if (!item) return;
  const el = item.querySelector('.si-stk'); if (!el) return;
  const pct = t.stockedPct;
  el.textContent = pct == null ? '–' : (pct >= 1000 ? '999+%' : Math.round(pct) + '%');
  el.className = 'si-stk ' + stockedClass(pct);
}
function orderKeyNav(e, inp) {
  const r = +inp.dataset.r, w = +inp.dataset.w;
  let tr = null, tw = null;
  if (e.key === 'Enter' || e.key === 'ArrowDown') { tr = r + 1; tw = w; }
  else if (e.key === 'ArrowUp') { tr = r - 1; tw = w; }
  else if (e.key === 'ArrowRight' && inp.selectionStart >= inp.value.length) { tr = r; tw = w + 1; }
  else if (e.key === 'ArrowLeft' && inp.selectionStart <= 0) { tr = r; tw = w - 1; }
  if (tr === null) return;
  const next = document.querySelector(`input[data-r="${tr}"][data-w="${tw}"]`);
  if (next) { e.preventDefault(); commitEdit(inp); next.focus(); }
}
function orderPaste(e, inp) {
  const text = (e.clipboardData || window.clipboardData).getData('text');
  if (!/[\t\n]/.test(text)) return;       // single value: default behaviour
  e.preventDefault();
  const rows = text.replace(/\r/g, '').split('\n').filter(s => s.length);
  const r0 = +inp.dataset.r, w0 = +inp.dataset.w;
  rows.forEach((line, dr) => {
    line.split('\t').forEach((cell, dw) => {
      const target = document.querySelector(`input[data-r="${r0 + dr}"][data-w="${w0 + dw}"]`);
      if (target) { target.value = cell.trim(); commitEdit(target); }
    });
  });
}

/* ---------------- summary view ---------------- */
function weeklyTable(rows) {
  const hw = highlightWeek();
  let h = '<div class="gridwrap"><table class="grid"><thead>' + headerRow() + '</thead><tbody>';
  for (const [label, arr, fmt] of rows) {
    let tot = 0;
    h += `<tr><td class="lbl">${label}</td>`;
    for (let w = 0; w < WEEKS; w++) {
      tot += arr[w];
      h += `<td class="${w + 1 === hw ? 'curwk' : ''}${Math.abs(arr[w]) < .005 ? ' zero' : ''}">${fmt(arr[w])}</td>`;
    }
    const isPct = fmt === fmtPct;
    h += `<td class="tot">${isPct ? '' : fmt(tot)}</td></tr>`;
  }
  return h + '</tbody></table></div>';
}
function fmtPct(v) { return Math.abs(v) < .0005 ? '–' : (v * 100).toFixed(0) + '%'; }

// Combined Summary tab: whole-plan totals + sales-shape chart at the top, then
// per-supplier breakdowns (value split + sales chart with container markers).
function renderSummary() {
  const g = AGG.g;
  const CC = SETTINGS.container_cbm || 68;
  const sum = (a, from, to) => a.slice(from, to).reduce((x, y) => x + y, 0);
  const sumAll = a => a.reduce((x, y) => x + y, 0);
  const q = [[0, 13], [13, 26], [26, 39], [39, 53]];
  const hist = (M.history && M.history.quarters) || {};

  // Year-to-date comparison through the most recent actual data week (= week before
  // the current week). `salesArr` holds banked actuals over those weeks; `fcArr` the plan.
  const aw = SETTINGS.current_week - 1;
  const ytd = (salesArr, fcArr) => {
    if (aw < 1) return null;
    const actual = sum(salesArr, 0, aw), fcast = sum(fcArr, 0, aw);
    const v = actual - fcast, pct = fcast ? (v / fcast) * 100 : 0;
    return { actual, fcast, v, pct, sign: v >= 0 ? '+' : '−', cls: v >= 0 ? 'c-up' : 'c-down' };
  };
  const gy = ytd(g.sales, g.fcSales);
  const gSub = gy ? `<div class="kpi-ytd">
      <div class="kpi-ytd-h">YTD to W${aw} (banked actuals)</div>
      <div class="kpi-ytd-r"><span>Actual sales</span><b>${fmtGBPk(gy.actual)}</b></div>
      <div class="kpi-ytd-r"><span>Forecast</span><b>${fmtGBPk(gy.fcast)}</b></div>
      <div class="kpi-ytd-r"><span>Variance</span><b class="${gy.cls}">${gy.sign}${fmtGBPk(Math.abs(gy.v))} (${gy.sign}${Math.abs(gy.pct).toFixed(0)}%)</b></div>
    </div>` : '';

  const cards = [
    [fmtGBPk(g.totSales), 'Forecast sales ' + YEAR, gSub],
    [fmtGBPk(g.totFob), 'Order FOB spend'],
    [Math.round(g.totUnits).toLocaleString(), 'Order units'],
    [Math.round(g.totCbm).toLocaleString() + ' cbm', 'Order volume'],
    [g.totContainers.toFixed(1), 'Containers (' + CC + ' cbm)'],
    [fmtGBPk(g.peakShv), 'Peak stock holding (W' + g.peakShvWk + ')'],
    [(g.peakRack * 100).toFixed(0) + '%', 'Peak racking use (W' + g.peakRackWk + ')'],
    [(g.peakStill * 100).toFixed(0) + '%', 'Peak stillage use (W' + g.peakStillWk + ')'],
  ].map(([v, l, s]) => `<div class="card${s ? ' card-wide' : ''}"><div class="v">${v}</div><div class="l">${l}</div>${s || ''}</div>`).join('');

  // container arrivals per week (committed + proposed) for the whole plan
  const gc = supCbmByWeek(null);
  const gCont = gc.com.map((c, w) => (c + gc.prop[w]) / CC);

  const priorYears = Object.keys(hist);
  let qt = `<table class="flat"><thead><tr><th></th>${priorYears.map(y => `<th>${y} actual</th>`).join('')}
            <th>${YEAR} forecast</th><th>${YEAR} import FOB</th></tr></thead><tbody>`;
  ['Q1', 'Q2', 'Q3', 'Q4'].forEach((qn, i) => {
    qt += `<tr><td>${qn}</td>${priorYears.map(y => `<td>${fmtGBP(hist[y][i])}</td>`).join('')}
           <td>${fmtGBP(sum(g.sales, ...q[i]))}</td><td>${fmtGBP(sum(g.fob, ...q[i]))}</td></tr>`;
  });
  qt += `<tr class="tot"><td>Year</td>${priorYears.map(y => `<td>${fmtGBP(hist[y].reduce((a, b) => a + b, 0))}</td>`).join('')}
         <td>${fmtGBP(g.totSales)}</td><td>${fmtGBP(g.totFob)}</td></tr></tbody></table>`;

  const chartLegend = `<div class="sum-legend">
    <span class="lg-line">Sales £ (actual / projected)</span>
    <span class="lg-fc">Forecast sales £ (plan)</span>
    <span class="lg-stock">Stock holding £ (own scale)</span>
    <span class="lg-cont">Containers arriving (committed + proposed)</span>
    ${highlightWeek() ? '<span class="lg-cur">Current week</span>' : ''}
    <span class="lg-hint">Hover a chart for weekly detail</span></div>`;

  // per-supplier breakdown rows (sorted by sales, suppliers with any sales)
  const totSales = g.totSales || 1;
  const supRows = M.suppliers.map(s => {
    const t = AGG.bySup.get(s.name);
    if (!t) return null;
    return { name: s.name, origin: s.origin || '', t,
             sales: sumAll(t.sales), units: sumAll(t.units), fob: sumAll(t.fob),
             cbm: sumAll(t.cbm), peakShv: Math.max(0, ...t.shv) };
  }).filter(r => r && r.sales > 0.5).sort((a, b) => b.sales - a.sales);

  const supBlocks = supRows.map((r, idx) => {
    const cb = supCbmByWeek(r.name);
    const cont = cb.com.map((c, w) => (c + cb.prop[w]) / CC);
    const metrics = [
      ['Sales', fmtGBP(r.sales)], ['Share', (100 * r.sales / totSales).toFixed(1) + '%'],
      ['Units', fmtU(r.units)], ['FOB', fmtGBP(r.fob)],
      ['Containers', fmt1(r.cbm / CC)], ['Peak stock', fmtGBP(r.peakShv)],
    ].map(([l, v]) => `<span>${l} <b>${v}</b></span>`).join('');
    const sy = ytd(r.t.sales, r.t.fcSales);
    const ytdHtml = sy ? `<div class="sum-sup-ytd"><span class="ssy-lbl">YTD to W${aw}:</span>`
      + `<span>actual <b>${fmtGBP(sy.actual)}</b></span>`
      + `<span>forecast <b>${fmtGBP(sy.fcast)}</b></span>`
      + `<span>variance <b class="${sy.cls}">${sy.sign}${fmtGBP(Math.abs(sy.v))} (${sy.sign}${Math.abs(sy.pct).toFixed(0)}%)</b></span></div>` : '';
    return `<div class="sum-sup">
      <div class="sum-sup-head">
        <h3><span class="sum-sup-rank">#${idx + 1}</span><a data-sup="${esc(r.name)}" title="Open in Plan view">${esc(r.name)}</a>${r.origin ? `<span class="ssh-grp">${esc(r.origin)}</span>` : ''}</h3>
        <div class="sum-sup-metrics">${metrics}</div>
      </div>
      ${ytdHtml}
      ${salesChart(r.t.sales, cont, r.t.shv, r.t.fcSales, { h: 84 })}</div>`;
  }).join('');

  const main = document.getElementById('main');
  main.innerHTML = `
    <div class="cards">${cards}</div>
    <div class="sum-grand">
      <div class="sum-sec-title">Whole-year sales shape — all suppliers</div>
      ${salesChart(g.sales, gCont, g.shv, g.fcSales, { h: 150 })}
      ${chartLegend}
    </div>
    ${chanSummaryHtml()}
    <h2 class="sect">Weekly totals — ${YEAR}</h2>
    ${weeklyTable([
      ['Sales £', g.sales, fmtGBP],
      ['Forecast units', g.fcUnits, fmtU],
      ['Order units', g.units, fmtU],
      ['Order FOB £', g.fob, fmtGBP],
      ['Arrival CBM', g.cbm, fmt1],
      ['Containers', g.containers, fmt1],
      ['Stock holding £', g.shv, fmtGBP],
      ['Pallet spaces', g.pallet, fmtU],
      ['Stillage spaces', g.stillage, fmtU],
      ['Racking spaces', g.racking, fmtU],
      ['Racking use (WEBSA)', g.rackUseWebsa, fmtPct],
      ['Racking use (total)', g.rackUseTotal, fmtPct],
      ['Stillage use (WEBSA)', g.stillUseWebsa, fmtPct],
      ['Stillage use (total)', g.stillUseTotal, fmtPct],
    ])}
    <h2 class="sect">Quarterly sales vs history</h2>${qt}
    <h2 class="sect">Supplier breakdown — ${supRows.length} suppliers by ${YEAR} sales</h2>
    ${supBlocks}
    <div style="height:30px"></div>`;
  main.querySelectorAll('a[data-sup]').forEach(a => a.addEventListener('click', () => {
    currentSupplier = a.dataset.sup; savePref('tp_supplier', currentSupplier); setView('plan');
  }));
}

/* ================= Warehouse Capacity view ================= */
let WH_WEEK = 0;   // remembers the leaderboard week across re-renders
// Capacity utilisation chart — same visual language as salesChart. Two lines:
// pallets+racking and stillage, each as a % of their capacity, with a dashed 100%
// reference. Hit columns drive the shared hover tooltip (kind='cap').
function capacityChart(d, opt) {
  opt = opt || {};
  const W = 1000, H = opt.h || 150, padT = 12, padB = 22, n = d.rackUse.length;
  const innerH = H - padT - padB, base = padT + innerH;
  const maxV = Math.max(1.05, ...d.rackUse, ...d.stillUse) * 1.05;
  const x = i => (i / (n - 1)) * W;
  const y = v => padT + (1 - v / maxV) * innerH;
  let rl = '', area = `M0 ${base.toFixed(1)}`;
  for (let i = 0; i < n; i++) { const px = x(i).toFixed(1), py = y(d.rackUse[i]).toFixed(1); rl += (i ? 'L' : 'M') + px + ' ' + py + ' '; area += ` L${px} ${py}`; }
  area += ` L${W} ${base.toFixed(1)} Z`;
  let sl2 = '';
  for (let i = 0; i < n; i++) sl2 += (i ? 'L' : 'M') + x(i).toFixed(1) + ' ' + y(d.stillUse[i]).toFixed(1) + ' ';
  const capY = y(1).toFixed(1);
  const capLine = `<line class="cap-100" x1="0" y1="${capY}" x2="${W}" y2="${capY}"/>`
    + `<text class="sl-ax cap-100-lbl" x="4" y="${(+capY - 3).toFixed(1)}" text-anchor="start">capacity 100%</text>`;
  let grid = '', axis = `<text class="sl-ax" x="2" y="${H - 4}" text-anchor="start">W1</text>`;
  [13, 26, 39].forEach(qk => {
    const gx = x(qk);
    grid += `<line class="sl-grid" x1="${gx.toFixed(1)}" y1="${padT}" x2="${gx.toFixed(1)}" y2="${base.toFixed(1)}"/>`;
    axis += `<text class="sl-ax" x="${gx.toFixed(1)}" y="${H - 4}" text-anchor="middle">W${qk + 1}</text>`;
  });
  axis += `<text class="sl-ax" x="${(W - 2).toFixed(1)}" y="${H - 4}" text-anchor="end">W${n}</text>`;
  let curLine = ''; const cw = highlightWeek();
  if (cw >= 1 && cw <= n) { const cx = x(cw - 1).toFixed(1); curLine = `<line class="sl-cur" x1="${cx}" y1="${padT}" x2="${cx}" y2="${base.toFixed(1)}"/>`; }
  const hw = W / (n - 1); let hits = '';
  for (let i = 0; i < n; i++) {
    hits += `<rect class="sl-hit" x="${(x(i) - hw / 2).toFixed(1)}" y="0" width="${hw.toFixed(1)}" height="${H}" fill="transparent"`
      + ` data-kind="cap" data-w="${i + 1}" data-pallet="${Math.round(d.pallet[i])}" data-racking="${Math.round(d.racking[i])}"`
      + ` data-stillage="${Math.round(d.stillage[i])}" data-ruse="${(d.rackUse[i] * 100).toFixed(0)}" data-suse="${(d.stillUse[i] * 100).toFixed(0)}"></rect>`;
  }
  return `<svg class="saleschart" viewBox="0 0 ${W} ${H}">${grid}${capLine}<path class="sl-area" d="${area}"/>`
    + `<path class="sl-line" d="${rl}"/><path class="sl-stock" d="${sl2}"/>${curLine}${axis}${hits}</svg>`;
}
// Ranked "most space-demanding products" leaderboards for one week — pallets+racking
// and stillages, each a horizontal bar list sorted by spaces used.
function whSpaceLists(week) {
  const w = week - 1;
  const pal = [], still = [];
  for (const sku of M.skus) {
    if (!(sku.fpq > 0) || !sku.pallet_type) continue;
    const r = RES.get(sku.id); if (!r) continue;
    const spaces = Math.ceil((r.stock[w] || 0) / sku.fpq);
    if (spaces <= 0) continue;
    const rec = { code: sku.code, name: sku.name || '', supplier: sku.supplier, spaces, type: sku.pallet_type };
    (sku.pallet_type === 'Stillage' ? still : pal).push(rec);
  }
  const list = (arr, title, cls) => {
    arr.sort((a, b) => b.spaces - a.spaces);
    if (!arr.length) return `<div class="wh-col"><div class="wh-col-h">${title}</div><div class="muted-note">No stock held in this week.</div></div>`;
    const top = arr.slice(0, 15), max = top[0].spaces || 1, tot = arr.reduce((a, b) => a + b.spaces, 0);
    const rows = top.map(r => `<div class="wh-row">`
      + `<div class="wh-bar-wrap"><div class="wh-bar ${cls}" style="width:${(r.spaces / max * 100).toFixed(1)}%"></div>`
      + `<span class="wh-code">${esc(r.code)}</span><span class="wh-nm" title="${esc(r.supplier)}">${esc(r.name)}</span></div>`
      + `<span class="wh-sp"><b>${fmtU(r.spaces)}</b> ${r.type === 'Racking' ? 'rack' : 'sp'}</span></div>`).join('');
    return `<div class="wh-col"><div class="wh-col-h">${title}<span class="wh-col-tot">${fmtU(tot)} spaces · ${arr.length} products</span></div>`
      + rows + (arr.length > top.length ? `<div class="muted-note">+${arr.length - top.length} more</div>` : '') + `</div>`;
  };
  return list(pal, 'Pallets &amp; racking', 'wh-bar-pal') + list(still, 'Stillages', 'wh-bar-still');
}
function renderWarehouse() {
  const g = AGG.g;
  const cap = SETTINGS.capacities || {};
  const rackTotal = (+cap.racking_websa || 0) + (+cap.racking_express || 0) + (+cap.racking_refit || 0);
  const stillTotal = (+cap.stillage_websa || 0) + (+cap.stillage_lough || 0) + (+cap.stillage_express || 0) + (+cap.stillage_free || 0);
  const chart = capacityChart({ rackUse: g.rackUseTotal, stillUse: g.stillUseTotal, pallet: g.pallet, racking: g.racking, stillage: g.stillage }, { h: 150 });
  const cards = [
    [(g.peakRack * 100).toFixed(0) + '%', 'Peak pallet/racking use (W' + g.peakRackWk + ')'],
    [(g.peakStill * 100).toFixed(0) + '%', 'Peak stillage use (W' + g.peakStillWk + ')'],
    [Math.round(Math.max(0, ...g.pallet)).toLocaleString(), 'Peak pallet spaces'],
    [Math.round(Math.max(0, ...g.stillage)).toLocaleString(), 'Peak stillage spaces'],
    [Math.round(rackTotal).toLocaleString(), 'Racking capacity'],
    [Math.round(stillTotal).toLocaleString(), 'Stillage capacity'],
  ].map(([v, l]) => `<div class="card"><div class="v">${v}</div><div class="l">${l}</div></div>`).join('');
  const legend = `<div class="sum-legend">
    <span class="lg-line">Pallets &amp; racking (% of capacity)</span>
    <span class="lg-stock">Stillages (% of capacity)</span>
    <span class="lg-cap">100% capacity</span>
    ${highlightWeek() ? '<span class="lg-cur">Current week</span>' : ''}
    <span class="lg-hint">Hover for weekly detail</span></div>`;
  const combo = g.pallet.map((p, w) => p + g.racking[w] + g.stillage[w]);
  const peakW = combo.indexOf(Math.max(...combo)) + 1;
  const defW = (WH_WEEK >= 1 && WH_WEEK <= WEEKS) ? WH_WEEK : (highlightWeek() || peakW);
  const weekOpts = Array.from({ length: WEEKS }, (_, i) =>
    `<option value="${i + 1}"${i + 1 === defW ? ' selected' : ''}>Week ${i + 1} · w/c ${weekDate(i + 1)}</option>`).join('');
  const main = document.getElementById('main');
  main.innerHTML = `
    <div class="cards">${cards}</div>
    <div class="sum-grand">
      <div class="sum-sec-title">Warehouse space utilisation — ${YEAR}</div>
      ${chart}
      ${legend}
    </div>
    <div class="wh-head">
      <h2 class="sect" style="margin:0">Most space-demanding products</h2>
      <label class="wh-wk">Week <select id="wh-week">${weekOpts}</select></label>
    </div>
    <div id="wh-lists" class="wh-lists"></div>
    <h2 class="sect">Weekly capacity totals — ${YEAR}</h2>
    ${weeklyTable([
      ['Pallet spaces', g.pallet, fmtU],
      ['Racking spaces', g.racking, fmtU],
      ['Stillage spaces', g.stillage, fmtU],
      ['Pallet/racking use', g.rackUseTotal, fmtPct],
      ['Stillage use', g.stillUseTotal, fmtPct],
    ])}
    <div style="height:30px"></div>`;
  const sel = document.getElementById('wh-week');
  const draw = () => { WH_WEEK = +sel.value; document.getElementById('wh-lists').innerHTML = whSpaceLists(WH_WEEK); };
  sel.addEventListener('change', draw);
  draw();
}

/* ---------------- seasonality apply / weather ---------------- */
async function loadWeather() {
  const se = SEASON();
  try {
    const r = await fetch(`/api/weather?lat=${se.lat || 52.77}&lon=${se.lon || -1.21}`);
    WEATHER = await r.json();
  } catch (e) { WEATHER = { ok: false, error: e.message }; }
  return WEATHER;
}
function weatherSummary() {
  if (!seasonActiveFor(YEAR) || !SEASON().useWeather) return '';
  if (!WEATHER) return 'weather: loading…';
  if (!WEATHER.ok) return 'weather: unavailable (using seasonal curve only)';
  const wks = Object.keys(WEATHER.weeks || {}).map(Number).sort((a, b) => a - b);
  if (!wks.length) return 'weather: no forecast weeks in range';
  return `weather: W${wks[0]}–W${wks[wks.length - 1]} live (${WEATHER.lat?.toFixed?.(2)}, ${WEATHER.lon?.toFixed?.(2)})`;
}
// Never touches the proposed-rebuy layer: changing the forecast makes the existing
// suggestions stale (flagged in the Plan toolbar), but only a "Run rebuy" click may
// rebuild them — otherwise applying a forecast setting would silently undo every
// supplier the user had already reviewed and cleared.
async function applySeasonality(opts = {}) {
  const se = SEASON();
  const status = document.getElementById('season-status');
  if (seasonActiveFor(YEAR) && se.useWeather) {
    if (status) status.textContent = 'Fetching live weather…';
    await loadWeather();
  }
  buildModeledForecasts();
  computeAll();
  renderSidebar(); setView(currentView);
  if (status) status.textContent = seasonActiveFor(YEAR) ? weatherSummary() : 'Model off — showing original figures.';
}

/* ---------------- settings tabs + seasonal-curve diagnostic ---------------- */
const SETTINGS_TABS = ['forecast', 'capacity', 'cover', 'rebuy', 'supplier', 'imports', 'data', 'changelog'];
// Defaults mirror supplier_form.FORM_DEFAULTS (server-side). Percentages are stored
// as fractions of order value (0.01 = 1%); the Settings inputs show them as whole %.
const SUPPLIER_FORM_DEFAULTS = { sailing_days: 50, grace_days: 7, inland_days: 7, marketing_pct: 0.01, deposit_pct: 0.15 };
function formCfg() { return Object.assign({}, SUPPLIER_FORM_DEFAULTS, SETTINGS.supplier_form || {}); }
let settingsTab = 'forecast';
function setSettingsTab(name) {
  settingsTab = name;
  document.querySelectorAll('.stab').forEach(b => b.classList.toggle('active', b.dataset.tab === name));
  document.querySelectorAll('.stab-panel').forEach(p => p.classList.toggle('hidden', p.dataset.panel !== name));
}

function diagCurveSvg(prof) {
  const w = 150, h = 38, n = prof.length, max = Math.max(...prof, 1.0001);
  const y = v => (h - 2 - (v / max) * (h - 6));
  const pts = prof.map((v, i) => `${(i / (n - 1) * w).toFixed(1)},${y(v).toFixed(1)}`);
  const area = `0,${h} ` + pts.join(' ') + ` ${w},${h}`;
  const pk = peakWeek(prof), pkx = ((pk - 1) / (n - 1) * w).toFixed(1);
  return `<svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" class="diag-curve">
    <line class="diag-mean" x1="0" y1="${y(1).toFixed(1)}" x2="${w}" y2="${y(1).toFixed(1)}"/>
    <polygon class="diag-line" points="${area}"/>
    <line class="diag-peak" x1="${pkx}" y1="2" x2="${pkx}" y2="${h - 2}"/>
  </svg>`;
}
function renderSeasonDiag() {
  const el = document.getElementById('season-diag');
  if (!el) return;
  calibrateSeasonalShapes(SETTINGS.current_week);
  const stat = {};   // key -> { n, lyVol }
  for (const s of M.skus) {
    for (const [g, k] of [['s', s.season], ['c', s.category]]) {
      const id = g + '|' + k; (stat[id] = stat[id] || { n: 0, vol: 0 });
      stat[id].n++; stat[id].vol += (s.ly || []).reduce((a, b) => a + b, 0);
    }
  }
  const card = (label, kind, prof, key) => {
    const st = stat[key] || { n: 0, vol: 0 };
    return `<div class="diag-card"><div class="dc-top"><span class="dc-name" title="${esc(label)}">${esc(label)}</span>
      <span class="dc-peak">peak W${peakWeek(prof)}</span></div>${diagCurveSvg(prof)}
      <div class="dc-sub">${kind} · ${st.n} SKUs · ${Math.round(st.vol).toLocaleString()} LY units</div></div>`;
  };
  let html = '<div class="diag-head">Calibrated weekly demand shape (mean = 1). Orange line = peak week; dashed = average.</div>';
  html += '<div class="diag-head">By season group</div>';
  for (const [k, p] of CALIB.bySeason) html += card(k, 'season', p, 's|' + k);
  html += '<div class="diag-head">By category (top by volume)</div>';
  const cats = [...CALIB.byCat.entries()]
    .map(([k, p]) => ({ k, p, vol: (stat['c|' + k] || { vol: 0 }).vol }))
    .sort((a, b) => b.vol - a.vol).slice(0, 12);
  for (const c of cats) html += card(c.k, 'category', c.p, 'c|' + c.k);
  el.innerHTML = html;
}

function explainSku(id) {
  const sku = skuById.get(id);
  if (!sku) return;
  ensureCalib();
  const d = computeSkuModel(sku);
  const cur = SETTINGS.current_week;
  const futOrig = sku.base_forecast.slice(cur - 1).reduce((a, b) => a + b, 0);
  const futMod = d.modeled.slice(cur - 1).reduce((a, b) => a + b, 0);
  const pct = futOrig > 0 ? (futMod / futOrig - 1) * 100 : 0;
  const sig = (val, weight, label) => `<div class="xp-sig"><div class="v">${val == null ? 'n/a' : fmt1(val)}</div>
    <div class="l">${label}${val == null ? '' : ` · weight ${weight.toFixed(1)}`}</div></div>`;
  // weather note
  let weatherNote = '';
  if (d.seasonal && SEASON().useWeather && WEATHER && WEATHER.ok && WEATHER.weeks) {
    const dir = /summer/i.test(sku.season) ? 1 : -1, sens = (SEASON().weatherStrength ?? 0.5) * 0.06;
    const parts = Object.keys(WEATHER.weeks).map(Number).sort((a, b) => a - b).map(wk => {
      const an = WEATHER.weeks[wk].anomaly, f = dir * clamp(an, -8, 8) * sens;
      return `W${wk} ${an >= 0 ? '+' : ''}${an}°C → ${f >= 0 ? '+' : ''}${(f * 100).toFixed(0)}%`;
    });
    weatherNote = `<div class="xp-note"><b>Weather</b> (${esc(SEASON().locationName || '')}): ${parts.join(' · ')} applied to these near weeks.</div>`;
  }
  const offNote = seasonActiveFor(YEAR) ? '' :
    `<div class="xp-off">The seasonality model is <b>off for ${YEAR}</b>, so the plan shows the original forecast. Figures below are what enabling it (at ${Math.round(seasonStrengthFor(YEAR) * 100)}% strength) would produce.</div>`;
  document.getElementById('explain-content').innerHTML = `
    <h3><span class="code">${esc(sku.code)}</span> ${esc(sku.name || '')}</h3>
    <div class="xp-meta">${statusBadge(sku.status)}<span class="badge" style="background:#eef;color:#446">${esc(sku.season)}</span>
      <span class="badge" style="background:#efe;color:#464">${esc(sku.category || '')}</span></div>
    ${offNote}
    <div class="xp-section">Baseline weekly demand — blended from</div>
    <div class="xp-signals">
      ${sig(d.rateTy, d.wTy, 'This-year actual run-rate (in-stock weeks)')}
      ${sig(d.rateLy, d.wLy, `Last-year rate${d.suspectCount ? ` (excl. ${d.suspectCount} stock-out wk)` : ''}`)}
      ${sig(d.rateOrig, d.wOrig, 'Original planner forecast')}
      <div class="xp-sig blend"><div class="v">${fmt1(d.D)}</div><div class="l">Blended baseline (units/wk)</div></div>
    </div>
    <div class="xp-section">Applied seasonal shape — ${esc(sku.season)}, peak week ${peakWeek(d.prof)}</div>
    ${diagCurveSvg(d.prof).replace('diag-curve', 'xp-curve')}
    <div class="xp-section">Resulting future forecast (W${cur}–53)</div>
    <div class="xp-result">
      <div><div class="big">${Math.round(futOrig).toLocaleString()}</div><div class="l">original</div></div>
      <div style="color:var(--dim)">→</div>
      <div><div class="big" style="color:var(--accent)">${Math.round(futMod).toLocaleString()}</div><div class="l">modelled (${pct >= 0 ? '+' : ''}${pct.toFixed(0)}%)</div></div>
    </div>
    ${weatherNote}
    <div class="xp-note">Baseline = how many units/week this line sells once season is stripped out; the shape redistributes that across the year; strength blends model vs original. Past weeks are never changed.</div>`;
  document.getElementById('explain-dialog').showModal();
}

/* ---------------- settings ---------------- */
let coverDraft = [];   // working copy of cover bands while the dialog is open

function renderCoverBands() {
  const wrap = document.getElementById('cover-bands');
  wrap.innerHTML = coverDraft.map((b, i) => {
    const isLast = (b.max === null || b.max === undefined || b.max === '');
    return `<div class="band-row" data-i="${i}">
      <input type="color" class="band-col" value="${b.bg}">
      ${isLast
        ? '<span class="band-lbl">everything higher</span>'
        : `<span class="band-lbl">up to <input type="number" step="0.5" min="0" class="band-max" value="${b.max}"> weeks</span>
           <button type="button" class="band-del" title="Remove band">&times;</button>`}
    </div>`;
  }).join('');
  wrap.querySelectorAll('.band-max').forEach(inp => inp.addEventListener('input', e => {
    coverDraft[+e.target.closest('.band-row').dataset.i].max = e.target.value === '' ? 0 : parseFloat(e.target.value);
  }));
  wrap.querySelectorAll('.band-col').forEach(inp => inp.addEventListener('input', e => {
    coverDraft[+e.target.closest('.band-row').dataset.i].bg = e.target.value;
  }));
  wrap.querySelectorAll('.band-del').forEach(btn => btn.addEventListener('click', e => {
    coverDraft.splice(+e.target.closest('.band-row').dataset.i, 1);
    renderCoverBands();
  }));
}

// ---- upload freshness (Settings › Data): last-updated date + weekly staleness badge ----
function fmtUpDate(iso) {
  const p = String(iso).slice(0, 10).split('-').map(Number);
  const mon = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][p[1] - 1] || '';
  return `${p[2]} ${mon} ${p[0]}`;
}
// Whole Mon–Sun calendar weeks between the upload date and today (0 = this week, 1 = last week…).
function uploadWeekDiff(iso) {
  const ep = ((M && M.week1_start) || '2025-12-29').split('-').map(Number);   // week1_start is a Monday
  const eMs = Date.UTC(ep[0], ep[1] - 1, ep[2]);
  const bucket = ms => Math.floor((ms - eMs) / (7 * 86400000));
  const u = String(iso).slice(0, 10).split('-').map(Number);
  const n = new Date();
  return bucket(Date.UTC(n.getFullYear(), n.getMonth(), n.getDate())) - bucket(Date.UTC(u[0], u[1] - 1, u[2]));
}
// ---- File Imports tab: freshness-coloured, drag-reorderable import tiles ----
const IMPORT_DEFS = [
  { id: 'websa',  label: 'WEBSA Open PO',  input: 'po-file',         when: () => (PO_WEBSA && PO_WEBSA.importedAt) || null },
  { id: 'qlik',   label: 'Qlik Container', input: 'containers-file', when: () => (PO_CONTAINERS && PO_CONTAINERS.importedAt) || null },
  { id: 'buying', label: 'Buying Report',  input: 'buying-file',     when: () => (SETTINGS && SETTINGS.buying_updated_at) || null },
  { id: 'wksales', label: 'Weekly Sales',  input: 'wksales-file',    when: () => (SETTINGS && SETTINGS.wksales_updated_at) || null },
  { id: 'asp',    label: 'Sales / ASP',    input: 'asp-file',        when: () => (SETTINGS && SETTINGS.asp_updated_at) || null },
  { id: 'landed', label: 'Landed Costs',   input: 'landed-file',     when: () => (SETTINGS && SETTINGS.landed_updated_at) || null },
  { id: 'duty',   label: 'Duty Rates',     input: 'duty-file',       when: () => (SETTINGS && SETTINGS.duty_updated_at) || null },
  { id: 'chanidx', label: 'Channel Index', input: 'chanidx-file',   when: () => (CHANNEL_INDEX && CHANNEL_INDEX.importedAt) || null },
];
const IMPORT_DEF = Object.fromEntries(IMPORT_DEFS.map(d => [d.id, d]));
const DEFAULT_IMPORT_GROUPS = [
  { name: 'Weekly',  cadence: 'weekly',  items: ['websa', 'qlik', 'buying', 'wksales'] },
  { name: 'Monthly', cadence: 'monthly', items: ['asp', 'landed', 'duty', 'chanidx'] },
];
// Persisted groups (SETTINGS.import_groups), validated so every import appears exactly once.
function importGroups() {
  const saved = SETTINGS.import_groups;
  const groups = (Array.isArray(saved) && saved.length ? saved : DEFAULT_IMPORT_GROUPS)
    .map(g => ({ name: g.name, cadence: g.cadence === 'monthly' ? 'monthly' : 'weekly',
                 items: (g.items || []).filter(id => IMPORT_DEF[id]) }));
  const seen = new Set();
  for (const g of groups) g.items = g.items.filter(id => !seen.has(id) && seen.add(id));
  const missing = IMPORT_DEFS.map(d => d.id).filter(id => !seen.has(id));
  if (missing.length) { if (!groups.length) groups.push({ name: 'Weekly', cadence: 'weekly', items: [] }); groups[0].items.push(...missing); }
  return groups;
}
function saveImportGroups(groups) {
  SETTINGS.import_groups = groups.map(g => ({ name: g.name, cadence: g.cadence, items: g.items.slice() }));
  markDirty();
}
// Whole calendar months between an upload date and today (0 = this month, 1 = last month…).
function monthDiff(iso) {
  const u = String(iso).slice(0, 10).split('-').map(Number);
  const n = new Date();
  return (n.getFullYear() * 12 + n.getMonth()) - (u[0] * 12 + (u[1] - 1));
}
function importFreshness(iso, cadence) {
  if (!iso) return { cls: 'if-old', t: 'Never uploaded' };
  const d = cadence === 'monthly' ? monthDiff(iso) : uploadWeekDiff(iso);
  const unit = cadence === 'monthly' ? 'month' : 'week';
  if (d <= 0) return { cls: 'if-live', t: cadence === 'monthly' ? 'This month' : 'This week' };
  if (d === 1) return { cls: 'if-aged', t: 'Last ' + unit };
  return { cls: 'if-old', t: 'Outdated' };
}
let impDragId = null;
function renderUploadAges() {              // renders the File Imports tab tiles (name kept = existing call sites)
  const wrap = document.getElementById('imports-groups');
  if (!wrap) return;
  wrap.innerHTML = importGroups().map((g, gi) => {
    const tiles = g.items.map(id => {
      const def = IMPORT_DEF[id]; if (!def) return '';
      const iso = def.when(), f = importFreshness(iso, g.cadence);
      const when = iso ? 'Updated ' + fmtUpDate(iso) : 'Not uploaded yet';
      return `<button type="button" class="imp-tile ${f.cls}" draggable="true" data-imp="${id}" title="Click to upload · drag to reorder / move group">`
        + `<span class="imp-name">${esc(def.label)}</span>`
        + `<span class="imp-badge">${esc(f.t)}</span>`
        + `<span class="imp-when">${esc(when)}</span></button>`;
    }).join('');
    return `<div class="imp-group"><div class="imp-group-head">${esc(g.name)}<span class="imp-cadence">${esc(g.cadence)}</span></div>`
      + `<div class="imp-tiles" data-gi="${gi}">${tiles}<span class="imp-drop-end"></span></div></div>`;
  }).join('');
  wireImportDnd();
}
function wireImportDnd() {
  const wrap = document.getElementById('imports-groups'); if (!wrap) return;
  wrap.querySelectorAll('.imp-tile').forEach(t => {
    t.addEventListener('click', () => { if (impDragId) return; const d = IMPORT_DEF[t.dataset.imp]; const inp = d && document.getElementById(d.input); if (inp) inp.click(); });
    t.addEventListener('dragstart', e => { impDragId = t.dataset.imp; t.classList.add('imp-dragging'); e.dataTransfer.effectAllowed = 'move'; try { e.dataTransfer.setData('text/plain', t.dataset.imp); } catch (_) {} });
    t.addEventListener('dragend', () => { impDragId = null; wrap.querySelectorAll('.imp-over,.imp-dragging').forEach(x => x.classList.remove('imp-over', 'imp-dragging')); });
    t.addEventListener('dragover', e => { e.preventDefault(); t.classList.add('imp-over'); });
    t.addEventListener('dragleave', () => t.classList.remove('imp-over'));
    t.addEventListener('drop', e => { e.preventDefault(); e.stopPropagation(); dropImport(t.dataset.imp, +t.closest('.imp-tiles').dataset.gi); });
  });
  wrap.querySelectorAll('.imp-tiles').forEach(z => {
    z.addEventListener('dragover', e => e.preventDefault());
    z.addEventListener('drop', e => { e.preventDefault(); dropImport(null, +z.dataset.gi); });
  });
}
function dropImport(beforeId, targetGi) {
  if (!impDragId || beforeId === impDragId) return;
  const dragged = impDragId, groups = importGroups();
  for (const g of groups) g.items = g.items.filter(id => id !== dragged);
  const tg = groups[targetGi]; if (!tg) return;
  const idx = beforeId ? tg.items.indexOf(beforeId) : -1;
  if (idx >= 0) tg.items.splice(idx, 0, dragged); else tg.items.push(dragged);
  saveImportGroups(groups);
  renderUploadAges();
}
function openSettings() {
  renderChanMap();
  fetchChangelog().then(renderChangelog);
  const dlg = document.getElementById('settings-dialog');
  document.getElementById('po-status').textContent = poStatusText();
  renderUploadAges();
  document.getElementById('btn-po-undo').hidden = !PO_APPLY_UNDO;
  renderUnmatchedTable();
  document.getElementById('set-multiplier').value = SETTINGS.multiplier;
  document.getElementById('set-week').value = SETTINGS.current_week;
  document.getElementById('set-container').value = SETTINGS.container_cbm;
  { const lc = landedCalc();
    document.getElementById('lc-container-rate').value = lc.container_rate;
    document.getElementById('lc-fx').value = lc.fx;
    document.getElementById('lc-full-cbm').value = lc.full_container_cbm;
    document.getElementById('lc-inland-rate').value = lc.inland_rate;
    document.getElementById('lc-duty-pct').value = lc.duty_pct; }
  document.getElementById('iso-week-now').textContent = isoWeek(new Date());
  for (const k of Object.keys(SETTINGS.capacities)) {
    const el = document.getElementById('cap-' + k);
    if (el) el.value = SETTINGS.capacities[k];
  }
  { const P = palletDims(); ['l', 'w', 'h', 'maxKg'].forEach(k => { const el = document.getElementById('pd-' + k); if (el) el.value = P[k]; }); }
  coverDraft = (SETTINGS.cover_bands || defaultCoverBands()).map(b => ({ max: b.max, bg: b.bg }));
  if (!coverDraft.length || coverDraft[coverDraft.length - 1].max !== null)
    coverDraft.push({ max: null, bg: '#c27ba0' });   // ensure a catch-all exists
  renderCoverBands();
  // shared (all-year) model settings
  const se = SEASON();
  document.getElementById('season-ownshape').value = Math.round((se.ownShapeWeight != null ? se.ownShapeWeight : 0.6) * 100);
  document.getElementById('season-ownshape-val').textContent = Math.round((se.ownShapeWeight != null ? se.ownShapeWeight : 0.6) * 100) + '%';
  document.getElementById('season-weather').checked = se.useWeather !== false;
  document.getElementById('season-weatherStrength').value = Math.round((se.weatherStrength ?? 0.5) * 100);
  document.getElementById('season-weatherStrength-val').textContent = Math.round((se.weatherStrength ?? 0.5) * 100) + '%';
  document.getElementById('season-locname').value = se.locationName || '';
  document.getElementById('season-lat').value = se.lat ?? 52.77;
  document.getElementById('season-lon').value = se.lon ?? -1.21;
  document.getElementById('season-status').textContent = seasonActiveFor(YEAR) ? weatherSummary() : '';
  // per-year forecast mode: build a working draft, default the loaded year selected
  FC_DRAFT = {};
  for (const y of YEARS) { const c = seasonCfg(y); FC_DRAFT[y] = c ? { mode: c.mode, strength: c.strength, growth: c.growth, target: c.target } : { mode: 'off' }; }
  fcSelectedYear = null;
  fcSelectYear(String(YEAR));
  renderRebuySettings();
  { const f = formCfg();
    document.getElementById('sf-sailing').value = f.sailing_days;
    document.getElementById('sf-grace').value = f.grace_days;
    document.getElementById('sf-inland').value = f.inland_days;
    document.getElementById('sf-marketing').value = +(f.marketing_pct * 100).toFixed(2);
    document.getElementById('sf-deposit').value = +(f.deposit_pct * 100).toFixed(2); }
  renderAspEditor();
  renderCbmEditor();
  const diag = document.getElementById('season-diag');
  diag.classList.add('hidden'); diag.innerHTML = '';
  setSettingsTab('forecast');
  dlg.showModal();
}
function updateTargetReadout() {
  const el = document.getElementById('target-readout'); if (!el) return;
  const tv = parseFloat(document.getElementById('target-value').value) || 0;
  if (fcSelectedYear !== String(YEAR)) {   // demand comparison only available for the loaded year
    el.textContent = tv > 0 ? `Target for ${fcSelectedYear}. Switch the app to ${fcSelectedYear} to see how it compares with that year's forecast.` : '';
    return;
  }
  const cur = forecastDemandTotal();
  let s = `Current modelled forecast demand (${YEAR}): ${fmtGBPk(cur)}.`;
  if (tv > 0 && cur > 0) { const d = tv / cur * 100 - 100; s += ` Target = ${(d >= 0 ? '+' : '') + d.toFixed(1)}% vs current.`; }
  el.textContent = s;
}
/* ---- per-year forecast-mode editor (Settings → Forecast) ---- */
let FC_DRAFT = {};            // { year -> { mode, strength, growth, target } } working copy while the dialog is open
let fcSelectedYear = null;    // which year's config the controls currently show
function fcRenderYearTabs() {
  const wrap = document.getElementById('fc-year-tabs'); if (!wrap) return;
  wrap.innerHTML = YEARS.slice().sort().map(y => {
    const mode = (FC_DRAFT[y] && FC_DRAFT[y].mode) || 'off';
    const tag = mode === 'off' ? 'Original' : mode === 'target' ? 'Target' : 'Seasonality';
    return `<button type="button" class="fc-yr-tab fc-${mode}${y === fcSelectedYear ? ' on' : ''}" data-year="${y}">`
      + `<span class="fc-yr-y">${y}</span><span class="fc-yr-mode">${tag}</span></button>`;
  }).join('');
}
function fcUpdateModeUI() {
  const mode = (document.querySelector('input[name="fc-mode"]:checked') || {}).value || 'off';
  document.getElementById('fc-seasonality-params').classList.toggle('hidden', mode !== 'seasonality');
  document.getElementById('fc-target-params').classList.toggle('hidden', mode !== 'target');
}
function fcWriteControls(cfg) {
  cfg = cfg || { mode: 'off' };
  const mode = cfg.mode || 'off';
  document.querySelectorAll('input[name="fc-mode"]').forEach(r => { r.checked = (r.value === mode); });
  const str = cfg.strength != null ? cfg.strength : 0.7;
  document.getElementById('season-strength').value = Math.round(str * 100);
  document.getElementById('season-strength-val').textContent = Math.round(str * 100) + '%';
  document.getElementById('season-growth').value = (cfg.growth != null && +cfg.growth !== 0) ? cfg.growth : '';
  // only ever put a value in the target box while in Target mode; a stray number left in
  // the (hidden) box outside Target mode is step/min-valid now but still confusing, and
  // historically blocked the form from submitting. Pre-fill the current demand as a
  // starting point when entering Target mode with nothing saved.
  const tv = +cfg.target || 0;
  document.getElementById('target-value').value = (mode === 'target')
    ? (tv || (fcSelectedYear === String(YEAR) ? Math.round(forecastDemandTotal()) : '')) : '';
  fcUpdateModeUI();
  updateTargetReadout();
}
function fcReadControls() {
  const mode = (document.querySelector('input[name="fc-mode"]:checked') || {}).value || 'off';
  const strength = (parseFloat(document.getElementById('season-strength').value) || 0) / 100;
  const growth = parseFloat(document.getElementById('season-growth').value);
  const target = parseFloat(document.getElementById('target-value').value) || 0;
  const cfg = { mode };
  if (mode === 'seasonality') { cfg.strength = strength; cfg.growth = isFinite(growth) ? growth : 0; }
  if (mode === 'target') { cfg.target = target; cfg.strength = strength; }
  return cfg;
}
function fcSelectYear(year) {
  if (fcSelectedYear) FC_DRAFT[fcSelectedYear] = fcReadControls();   // stash the year we're leaving
  fcSelectedYear = String(year);
  fcRenderYearTabs();
  fcWriteControls(FC_DRAFT[fcSelectedYear]);
}
function seasonsList() {                 // distinct product seasons, friendly order
  const order = ['Summer', 'Winter', 'Continuity', 'No Defined Season'];
  const set = new Set(M.skus.map(s => s.season).filter(Boolean));
  return [...set].sort((a, b) => { const ia = order.indexOf(a), ib = order.indexOf(b); return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || a.localeCompare(b); });
}
function renderRebuySettings() {
  const rb = SETTINGS.rebuy || {};
  const def = rb.coverTarget != null ? rb.coverTarget : REBUY_TARGET;
  document.getElementById('rb-cover').value = def;
  document.getElementById('rb-peakboost').value = rb.peakBoost || 0;
  document.getElementById('rb-seasoncutoff').value = Math.round((rb.seasonalCutoff != null ? rb.seasonalCutoff : 0.40) * 100);
  document.getElementById('rb-maxwait').value = rb.maxWait != null ? rb.maxWait : 4;
  document.getElementById('rb-partialcbm').value = rb.partialCbm != null ? rb.partialCbm : 28;
  document.querySelectorAll('.rb-cc').forEach(el => el.textContent = SETTINGS.container_cbm || 68);
  const cs = rb.coverBySeason || {};
  document.getElementById('rb-season-cover').innerHTML = seasonsList().map(s =>
    `<label class="rb-season-row"><span>${esc(s)}</span><input type="number" step="0.5" min="0" class="rb-cov" data-season="${esc(s)}" value="${cs[s] != null ? cs[s] : ''}" placeholder="${def}"></label>`).join('');
  const sel = new Set(rb.oneProductSuppliers || []);
  const list = document.getElementById('rb-op-list');
  list.innerHTML = M.suppliers.map(s =>
    `<label class="ex-item"><input type="checkbox" value="${esc(s.name)}"${sel.has(s.name) ? ' checked' : ''}><span>${esc(titleCase(s.name))}</span></label>`).join('');
  document.getElementById('rb-op-search').oninput = e => {
    const q = e.target.value.toLowerCase();
    list.querySelectorAll('.ex-item').forEach(it => { it.style.display = it.textContent.toLowerCase().includes(q) ? '' : 'none'; });
  };
}
function applySettings() {
  SETTINGS.multiplier = parseFloat(document.getElementById('set-multiplier').value) || 1;
  SETTINGS.current_week = Math.min(53, Math.max(1, parseInt(document.getElementById('set-week').value) || 1));
  SETTINGS.current_week_by_year = Object.assign({}, SETTINGS.current_week_by_year, { [YEAR]: SETTINGS.current_week });
  SETTINGS.container_cbm = parseFloat(document.getElementById('set-container').value) || 68;
  { const d = LC_DEFAULTS, num = (id, def) => { const v = parseFloat(document.getElementById(id).value); return isFinite(v) && v >= 0 ? v : def; };
    SETTINGS.landed_calc = {
      container_rate: num('lc-container-rate', d.container_rate),
      fx: (parseFloat(document.getElementById('lc-fx').value) > 0 ? parseFloat(document.getElementById('lc-fx').value) : d.fx),
      full_container_cbm: (parseFloat(document.getElementById('lc-full-cbm').value) > 0 ? parseFloat(document.getElementById('lc-full-cbm').value) : d.full_container_cbm),
      inland_rate: num('lc-inland-rate', d.inland_rate),
      duty_pct: num('lc-duty-pct', d.duty_pct),
    }; }
  for (const k of Object.keys(SETTINGS.capacities)) {
    const el = document.getElementById('cap-' + k);
    if (el) SETTINGS.capacities[k] = parseFloat(el.value) || 0;
  }
  { const pd = SETTINGS.pallet_dims || {}; ['l', 'w', 'h', 'maxKg'].forEach(k => { const el = document.getElementById('pd-' + k); if (el) pd[k] = parseFloat(el.value) || PALLET_DEFAULT[k]; }); SETTINGS.pallet_dims = pd; }
  // normalise cover bands: finite thresholds sorted ascending, single catch-all last
  const cat = coverDraft.find(b => b.max === null || b.max === undefined || b.max === '');
  const finite = coverDraft.filter(b => !(b.max === null || b.max === undefined || b.max === ''))
    .map(b => ({ max: Number(b.max), bg: b.bg })).sort((a, b) => a.max - b.max);
  SETTINGS.cover_bands = [...finite, { max: null, bg: cat ? cat.bg : '#c27ba0' }];
  // forecast mode is PER-YEAR: commit the working draft (after capturing the year on screen)
  FC_DRAFT[fcSelectedYear] = fcReadControls();
  const byYear = {};
  for (const y of Object.keys(FC_DRAFT)) {
    const c = FC_DRAFT[y];
    if (!c || c.mode === 'off') continue;
    if (c.mode === 'target' && !(+c.target > 0)) continue;   // target with no value = original
    const o = { mode: c.mode };
    if (c.strength != null) o.strength = c.strength;
    if (c.mode === 'seasonality') o.growth = +c.growth || 0;
    if (c.mode === 'target') o.target = +c.target || 0;
    byYear[y] = o;
  }
  SETTINGS.seasonality = {
    ownShapeWeight: Math.min(1, Math.max(0, (parseFloat(document.getElementById('season-ownshape').value) || 0) / 100)),
    useWeather: document.getElementById('season-weather').checked,
    weatherStrength: (parseFloat(document.getElementById('season-weatherStrength').value) || 0) / 100,
    locationName: document.getElementById('season-locname').value.trim(),
    lat: parseFloat(document.getElementById('season-lat').value) || 52.77,
    lon: parseFloat(document.getElementById('season-lon').value) || -1.21,
    byYear,
  };
  // rebuy algorithm parameters
  const coverBySeason = {};
  document.querySelectorAll('#rb-season-cover .rb-cov').forEach(inp => {
    const v = inp.value.trim();
    if (v !== '') coverBySeason[inp.dataset.season] = parseFloat(v);
  });
  const oneProductSuppliers = [...document.querySelectorAll('#rb-op-list input:checked')].map(i => i.value);
  SETTINGS.rebuy = Object.assign({}, SETTINGS.rebuy, {
    coverTarget: parseFloat(document.getElementById('rb-cover').value) || REBUY_TARGET,
    peakBoost: parseFloat(document.getElementById('rb-peakboost').value) || 0,
    seasonalCutoff: Math.min(1, Math.max(0, (parseFloat(document.getElementById('rb-seasoncutoff').value) || 0) / 100)),
    maxWait: Math.max(0, parseInt(document.getElementById('rb-maxwait').value) || 0),
    partialCbm: Math.max(1, parseFloat(document.getElementById('rb-partialcbm').value) || 28),
    coverBySeason, oneProductSuppliers,
  });
  // supplier order-planning form assumptions (global; used by the Excel export)
  { const d = SUPPLIER_FORM_DEFAULTS;
    const intv = (id, def) => { const v = parseInt(document.getElementById(id).value, 10); return isFinite(v) && v >= 0 ? v : def; };
    const pct = (id, def) => { const v = parseFloat(document.getElementById(id).value); return isFinite(v) && v >= 0 ? v / 100 : def; };
    SETTINGS.supplier_form = {
      sailing_days: intv('sf-sailing', d.sailing_days),
      grace_days: intv('sf-grace', d.grace_days),
      inland_days: intv('sf-inland', d.inland_days),
      marketing_pct: pct('sf-marketing', d.marketing_pct),
      deposit_pct: pct('sf-deposit', d.deposit_pct),
    }; }
  markDirty();
  // rebuilds the forecast model, recomputes and re-renders. Existing rebuy suggestions
  // are LEFT ALONE (just flagged stale) — only "Run rebuy" may rebuild them.
  if (PROPOSED && PROPOSED.size) REBUY_STALE = true;
  applySeasonality();
}
async function restoreImportedOrders() {
  if (!confirm(`Replace ALL ${YEAR} order quantities with the ones imported from the Excel file? Your edits will be lost.`)) return;
  const r = await fetch('/api/imported-orders?year=' + encodeURIComponent(YEAR));
  ORDERS = await r.json();
  IMPORTED = JSON.parse(JSON.stringify(ORDERS));
  document.getElementById('settings-dialog').close();
  computeAll(); renderSidebar(); setView(currentView); markDirty();
}

/* ---------------- update ASP from a "Product Sales by Account" sales file ---------- */
let ASP_PARSED = null;    // { basic:{code:asp}, top:{code:asp}, fileSkus, withSales, fname } awaiting a choice
let ASP_CHOICE = null;    // 'original' | 'basic' | 'top' — the comparison row the user picked
async function aspFileChosen(e) {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  const status = document.getElementById('asp-status');
  status.textContent = 'Reading ' + file.name + '…';
  try {
    const r = await fetch('/api/parse-asp', { method: 'POST', body: await file.arrayBuffer() });
    const j = await r.json();
    if (!j.ok) { status.textContent = ''; alert('Could not read the file: ' + (j.error || 'unknown')); return; }
    ASP_PARSED = { basic: j.basic || {}, top: j.top || {}, fileSkus: j.fileSkus, withSales: j.withSales, fname: file.name };
    ASP_CHOICE = null;
    status.textContent = '';
    openAspDialog();
  } catch (err) { status.textContent = ''; alert('Read error: ' + err.message); }
}

/* ---- PO / container uploads + diagnostics ---- */
function poStatusText() {
  const w = (PO_WEBSA && PO_WEBSA.pos) ? Object.keys(PO_WEBSA.pos).length : 0;
  const c = (PO_CONTAINERS && PO_CONTAINERS.dates) ? Object.keys(PO_CONTAINERS.dates).length : 0;
  if (!w && !c) return 'No PO data loaded yet — upload both files to label arrivals.';
  return `Loaded: WEBSA ${w} POs · containers ${c} POs.${w && c ? '' : ' Upload both files to match arrivals.'}`;
}
async function poFileChosen(e, kind) {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  const status = document.getElementById('po-status');
  status.textContent = 'Reading ' + file.name + '…';
  const ep = kind === 'po' ? '/api/parse-po' : '/api/parse-containers';
  try {
    const r = await fetch(ep + '?year=' + encodeURIComponent(YEAR), { method: 'POST', body: await file.arrayBuffer() });
    const j = await r.json();
    if (!j.ok) { status.textContent = ''; alert('Could not read the file: ' + (j.error || 'unknown')); return; }
    const data = await (await fetch('/api/data?year=' + encodeURIComponent(YEAR))).json();
    PO_WEBSA = data.poWebsa || null;
    PO_CONTAINERS = data.poContainers || null;
    status.textContent = kind === 'po'
      ? `WEBSA loaded: ${j.poCount} POs / ${j.rows} lines. ${poStatusText()}`
      : `Containers loaded: ${j.poCount} POs over ${j.rows} rows (${j.unparsed} non-PO, ${j.truncated} "etc"). ${poStatusText()}`;
    renderUploadAges();         // refresh the "last updated" / staleness badge
    if (currentView === 'plan') renderPlan();
  } catch (err) { status.textContent = ''; alert('Read error: ' + err.message); }
}
function openPoDiag() {
  const c = document.getElementById('po-diag-content');
  if (!poDataReady()) {
    c.innerHTML = '<p class="muted-note">Upload both the WEBSA Open PO export and the Qlik Container export first.</p>';
  } else {
    const d = poDiagnostics();
    const sec = (title, n, blurb, arr, fmt, ok) =>
      `<div class="po-diag-sec"><h3>${title} <span class="${arr.length ? 'pd-n' : 'pd-ok'}">${arr.length}</span></h3><small>${blurb}</small>`
      + (arr.length ? '<ul class="po-diag-list">' + arr.slice(0, 200).map(fmt).join('')
          + (arr.length > 200 ? `<li>…and ${arr.length - 200} more</li>` : '') + '</ul>'
        : `<p class="po-diag-ok">✓ ${ok}</p>`) + '</div>';
    c.innerHTML =
      sec('POs awaiting a container booking', d.notBooked.length,
        "Outstanding POs covering your products that aren't on a Qlik booking yet (shown on their PO due date).",
        d.notBooked, o => `<li><b>${esc(o.po)}</b> — ${esc(titleCase(o.supplier || ''))} · ${fmtU(o.qty)} units${o.week ? ` · due ~W${o.week}` : ''}</li>`, 'All covered POs have a booking.')
      + sec('Committed weeks with no PO match', d.notFoundWeeks.length,
        'Suppliers where the plan shows committed stock from the current week on, but no PO / container arrives that week.',
        d.notFoundWeeks, o => `<li><b>${esc(titleCase(o.supplier))}</b> — W${o.weeks.join(', W')}${o.hasPo ? '' : ' · no PO raised for this supplier'}</li>`, 'Every committed week maps to a PO.')
      + sec('Planned stock not yet raised', d.notRaised.length,
        "Products with committed orders from the current week on that aren't on any outstanding PO.",
        d.notRaised, o => `<li><b>${esc(o.code)}</b> — ${esc(titleCase(o.supplier))}</li>`, 'All planned stock is on a PO.')
      + sec('Unmatched container rows', d.unparsed.length,
        "Qlik 'Order' cells with no recognisable PO number (charter refs, samples, blanks).",
        d.unparsed, o => `<li>${esc(o)}</li>`, 'Every container row carried a PO.')
      + sec('Possibly-incomplete consolidated containers', d.truncated.length,
        'Qlik \'Order\' cells ending in "etc" / "and more" / "+" — may reference more POs than listed.',
        d.truncated, o => `<li>${esc(o)}</li>`, 'No truncated PO lists.')
      + sec('WEBSA product codes not in the app', d.codesNotInApp.length,
        "Outstanding PO lines whose product code doesn't match an app SKU (new lines, barcodes, variants).",
        d.codesNotInApp, o => `<li>${esc(o)}</li>`, 'Every PO line matched an app SKU.');
  }
  document.getElementById('po-diag-dialog').showModal();
}

// Full details for a single PO (clicked from the PO row): supplier, arrival/booking,
// container leg(s), and every order line with ordered/delivered/outstanding.
function openPoDetail(po) {
  const c = document.getElementById('po-detail-content');
  const ttl = document.getElementById('po-detail-title');
  ttl.textContent = po;
  const info = PO_WEBSA && PO_WEBSA.pos && PO_WEBSA.pos[po];
  if (!info) { c.innerHTML = `<p class="muted-note">${esc(po)} isn't in the loaded WEBSA Open PO export.</p>`; document.getElementById('po-detail-dialog').showModal(); return; }
  const ctn = (PO_CONTAINERS && PO_CONTAINERS.dates && PO_CONTAINERS.dates[po]) || [];
  const arr = poArrivalWeeks(po)[0];
  const booked = !!ctn.length;
  const nameOf = code => { const s = M.skus.find(x => x.code === code); return s ? (s.name || '') : ''; };
  const appCodes = new Set(M.skus.map(s => s.code));
  // header: supplier + arrival summary
  const arrLine = arr
    ? (booked
        ? `Arriving <b>week ${arr.week}</b> (w/c ${weekDate(arr.week)}) · <span class="pd-badge pd-booked">booked</span> <span class="muted-note">${esc(arr.basis)}</span>`
        : `On <b>week ${arr.week}</b> (w/c ${weekDate(arr.week)}) by PO due date · <span class="pd-badge pd-unbooked">not booked</span>`)
    : '<span class="pd-badge pd-unbooked">no arrival date</span>';
  let h = `<div class="pd-head"><div><b>${esc(titleCase(info.supplier || ''))}</b></div><div>${arrLine}</div></div>`;
  // booking legs
  if (booked) {
    h += `<div class="po-diag-sec"><h3>Container booking${ctn.length > 1 ? 's' : ''} <span class="pd-n">${ctn.length}</span></h3>`
      + `<table class="flat pd-table"><thead><tr><th>Departure</th><th>ETA UK port</th><th>Delivery to CB</th><th>Status</th><th>Container</th><th>Shipment</th></tr></thead><tbody>`
      + ctn.map(r => `<tr><td>${fmtDate(r.etd)}</td><td>${fmtDate(r.etaPort)}</td><td>${fmtDate(r.deliveryCB)}</td><td>${esc(r.status || '–')}</td><td>${esc(r.container || '–')}</td><td>${esc(r.shipment || '–')}</td></tr>`).join('')
      + `</tbody></table></div>`;
  } else {
    h += `<p class="rt-short">No container booking in the Qlik export yet — shown on its PO due date (${fmtDate(info.lines.map(l => l.due).find(Boolean))}). It won't be moved by "Apply Qlik Container Dates" until a container is booked.</p>`;
  }
  // order lines
  const lines = info.lines.slice().sort((a, b) => b.outstanding - a.outstanding);
  const tot = lines.reduce((a, l) => ({ o: a.o + l.ordered, d: a.d + l.delivered, u: a.u + l.outstanding }), { o: 0, d: 0, u: 0 });
  h += `<div class="po-diag-sec"><h3>Order lines <span class="pd-n">${lines.length}</span></h3>`
    + `<table class="flat pd-table pd-lines"><thead><tr><th>Code</th><th>Description</th><th>Ordered</th><th>Delivered</th><th>Outstanding</th><th>Due</th></tr></thead><tbody>`
    + lines.map(l => `<tr${appCodes.has(l.code) ? '' : ' class="pd-notapp"'}><td class="pd-code">${esc(l.code)}${appCodes.has(l.code) ? '' : ' <span class="pum-flag">not in app</span>'}</td><td class="pd-name">${esc(nameOf(l.code))}</td><td>${fmtU(l.ordered)}</td><td>${fmtU(l.delivered)}</td><td><b>${fmtU(l.outstanding)}</b></td><td>${fmtDate(l.due)}</td></tr>`).join('')
    + `<tr class="tot"><td>Total</td><td></td><td>${fmtU(tot.o)}</td><td>${fmtU(tot.d)}</td><td>${fmtU(tot.u)}</td><td></td></tr>`
    + `</tbody></table></div>`;
  c.innerHTML = h;
  document.getElementById('po-detail-dialog').showModal();
}
function fmtDate(iso) {
  if (!iso) return '–';
  const [y, m, d] = iso.split('-').map(Number);
  return `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${String(y).slice(2)}`;
}

/* ---- "No PO" review: settings table + guided cycle through unmatched weeks ---- */
// Renders the actionable table into #po-unmatched (called when Settings opens).
function renderUnmatchedTable() {
  const wrap = document.getElementById('po-unmatched');
  const cnt = document.getElementById('pum-count');
  const btn = document.getElementById('btn-pum-cycle');
  if (!wrap) return;
  if (!poDataReady()) { wrap.innerHTML = '<p class="muted-note">Upload the WEBSA + Qlik files above to populate this.</p>'; cnt.textContent = ''; if (btn) btn.disabled = true; return; }
  const list = poUnmatchedList();
  PO_UNMATCHED = list; poCycleIdx = -1;
  if (!list.length) { wrap.innerHTML = `<p class="po-diag-ok">✓ Every committed week from W${SETTINGS.current_week} on matches a PO.</p>`; cnt.textContent = ''; if (btn) btn.disabled = true; return; }
  if (btn) btn.disabled = false;
  const groups = [];
  for (const e of list) { let g = groups[groups.length - 1]; if (!g || g.supplier !== e.supplier) { g = { supplier: e.supplier, items: [] }; groups.push(g); } g.items.push(e); }
  let h = '<table class="flat pum-table"><thead><tr><th>Supplier</th><th>Weeks with no PO</th><th>Committed units</th></tr></thead><tbody>';
  for (const g of groups) {
    const chips = g.items.map(e => `<button type="button" class="pum-wk" data-sup="${esc(e.supplier)}" data-w="${e.week}" title="Go to ${esc(titleCase(e.supplier))} week ${e.week}">W${e.week} · ${fmtU(e.units)}</button>`).join(' ');
    h += `<tr><td>${esc(titleCase(g.supplier))}${g.items[0].hasPo ? '' : ' <span class="pum-flag">no PO raised</span>'}</td><td>${chips}</td><td>${fmtU(g.items.reduce((a, e) => a + e.units, 0))}</td></tr>`;
  }
  wrap.innerHTML = h + '</tbody></table>';
  cnt.textContent = `${list.length} week${list.length > 1 ? 's' : ''} across ${groups.length} supplier${groups.length > 1 ? 's' : ''}`;
  wrap.querySelectorAll('.pum-wk').forEach(b => b.addEventListener('click', () => {
    const i = PO_UNMATCHED.findIndex(e => e.supplier === b.dataset.sup && e.week === +b.dataset.w);
    if (i >= 0) startPoReview(i);
  }));
}
// Begin the guided review (close Settings, jump to the entry, show the floating bar).
function startPoReview(idx) {
  PO_UNMATCHED = (PO_UNMATCHED && PO_UNMATCHED.length) ? PO_UNMATCHED : poUnmatchedList();
  if (!PO_UNMATCHED.length) return;
  if (idx == null) {                       // start at the current supplier's first unmatched week
    idx = PO_UNMATCHED.findIndex(e => e.supplier === currentSupplier);
    if (idx < 0) idx = 0;
  }
  poReviewActive = true;
  const dlg = document.getElementById('settings-dialog'); if (dlg && dlg.open) dlg.close();
  gotoUnmatched(idx);
}
function poReviewStep(dir) {
  if (!PO_UNMATCHED || !PO_UNMATCHED.length) return;
  gotoUnmatched((poCycleIdx + dir + PO_UNMATCHED.length) % PO_UNMATCHED.length);
}
function endPoReview() {
  poReviewActive = false;
  const bar = document.getElementById('po-review-bar'); if (bar) bar.remove();
}
// Navigate to a given unmatched entry: select supplier, show Plan, scroll to + flash the week.
function gotoUnmatched(idx) {
  if (!PO_UNMATCHED || idx < 0 || idx >= PO_UNMATCHED.length) return;
  poCycleIdx = idx;
  const e = PO_UNMATCHED[idx];
  if (currentSupplier !== e.supplier) { currentSupplier = e.supplier; savePref('tp_supplier', currentSupplier); renderSidebar(); }
  setView('plan');
  renderPoReviewBar();
  flashWeek(e.week);   // grid DOM is ready synchronously after setView
}
// Scroll the grid so the week is centred and briefly flash its header + PO cell.
function flashWeek(w) {
  const wrap = document.querySelector('.gridwrap'); if (!wrap) return;
  const ths = document.querySelectorAll('.grid thead tr:first-child th');
  const th = ths[w];                       // ths[0] = the "Week" label, ths[w] = W{w}
  if (th) wrap.scrollLeft = Math.max(0, th.offsetLeft - wrap.clientWidth / 2);
  const poCell = document.querySelectorAll('.grid thead .po-row td.po-cell')[w - 1];
  [th, poCell].forEach(el => { if (el) { el.classList.remove('pum-flash'); void el.offsetWidth; el.classList.add('pum-flash'); } });
}
// The floating review toolbar on the Plan (persists across re-renders since it lives on body).
function renderPoReviewBar() {
  if (!poReviewActive) return;
  let bar = document.getElementById('po-review-bar');
  if (!bar) {
    bar = document.createElement('div');
    bar.id = 'po-review-bar';
    document.body.appendChild(bar);
    bar.addEventListener('click', ev => {
      const a = ev.target.dataset.act;
      if (a === 'prev') poReviewStep(-1);
      else if (a === 'next') poReviewStep(1);
      else if (a === 'done') endPoReview();
    });
  }
  const e = PO_UNMATCHED[poCycleIdx] || {};
  bar.innerHTML = `<span class="prv-ttl">⚠ No-PO review</span>`
    + `<span class="prv-pos">${poCycleIdx + 1}/${PO_UNMATCHED.length}</span>`
    + `<span class="prv-loc"><b>${esc(titleCase(e.supplier || ''))}</b> · week <b>${e.week}</b> · ${fmtU(e.units || 0)} committed units with no PO</span>`
    + `<button type="button" data-act="prev" title="Previous">← Prev</button>`
    + `<button type="button" data-act="next" title="Next">Next →</button>`
    + `<button type="button" data-act="done" class="prv-done" title="End review">✕ Done</button>`;
}

/* ---- Phase 2 apply: whole-sheet preview dialog (Settings) + per-supplier apply (Plan) + undo ---- */
// Settings dialog = whole sheet (all suppliers). Per-supplier re-timing lives on the Plan view.
function openRetimeDialog() {
  document.getElementById('settings-dialog').close();
  document.getElementById('po-apply-dialog').showModal();
  renderRetimePreview();
}
function renderRetimePreview() {
  const box = document.getElementById('po-apply-body');
  const applyBtn = document.getElementById('po-apply-go');
  const r = computeRetime('all');
  PO_RETIME = r;
  if (!poDataReady()) { box.innerHTML = '<p class="muted-note">Upload the WEBSA + Qlik files first.</p>'; applyBtn.disabled = true; return; }
  if (!r.skus) {
    box.innerHTML = `<p class="po-diag-ok">✓ Nothing to re-time — every product on an outstanding PO already sits on its PO arrival week.</p>`;
    applyBtn.disabled = true; return;
  }
  applyBtn.disabled = false;
  let rows = '';
  for (const m of r.moves.slice(0, 400)) {
    const parts = m.placeWeeks.map(p => `→ <b>W${p.week}</b> ${fmtU(p.qty)}`);
    (m.clearWeeks || []).forEach(c => parts.push(`<span class="rt-demote">cleared W${c.week} ${fmtU(c.units)}</span>`));
    rows += `<tr><td class="rt-code">${esc(m.code)}</td><td class="rt-sup">${esc(titleCase(m.supplier))}</td><td class="rt-po">${m.pos.map(esc).join(', ') || '—'}</td><td>${parts.join(', ') || '—'}</td></tr>`;
  }
  const clearedNote = r.clearedUnits
    ? `<b>${fmtU(r.clearedUnits)}</b> units cleared from non-PO weeks · ` : '';
  const injectNote = r.injectedUnits
    ? `<b>${fmtU(r.injectedUnits)}</b> units newly added · ` : '';
  box.innerHTML =
    `<div class="rt-summary"><b>${r.skus}</b> product(s) · <b>${fmtU(r.placedUnits)}</b> units placed on PO weeks · ${clearedNote}${injectNote}<b>${r.weeks.size}</b> week(s) affected · scope: <b>whole year (all suppliers)</b></div>`
    + `<p class="muted-note">Each product on an outstanding PO has its committed arrivals (current week on) set to <b>exactly its PO schedule</b> — the PO's outstanding units on its arrival week (booked container date, else PO due date), with <b>every other future week cleared</b> so superseded arrivals don't linger. Products with <b>no outstanding PO are left untouched</b>. Past/delivered weeks are left alone. Undo immediately after applying.</p>`
    + `<table class="flat rt-table"><thead><tr><th>Code</th><th>Supplier</th><th>PO</th><th>Change (set / cleared)</th></tr></thead><tbody>${rows}</tbody></table>`
    + (r.moves.length > 400 ? `<p class="muted-note">…and ${r.moves.length - 400} more.</p>` : '');
}
// Shared apply: snapshot affected committed rows for undo, then set them to the PO
// schedule + save + re-render. (Only the committed layer is touched now — proposed is
// left alone, so the snapshot only needs the committed vectors.)
function doApplyRetime(plan, label) {
  if (!plan || !plan.skus) return false;
  const oIds = Object.keys(plan.newOrders);
  const snapO = {}, snapP = {};
  for (const id of oIds) snapO[id] = (ORDERS[id] || EMPTY53).slice();
  const parts = [`${plan.skus} product(s)`, `${fmtU(plan.placedUnits)} on PO weeks`];
  if (plan.clearedUnits) parts.push(`${fmtU(plan.clearedUnits)} cleared`);
  if (plan.injectedUnits) parts.push(`${fmtU(plan.injectedUnits)} added`);
  const detail = parts.join(', ');
  PO_APPLY_UNDO = { snapO, snapP, label, detail };
  for (const id of oIds) ORDERS[id] = plan.newOrders[id].slice();
  computeAll(); markDirty();
  renderSidebar();
  if (currentView === 'plan') renderPlan(); else setView(currentView);
  document.getElementById('save-status').textContent = `Re-timed ${label} (${detail}) — Undo available`;
  showRetimeToast();
  return true;
}
function applyRetime() {                 // from the Settings whole-sheet dialog
  if (!PO_RETIME || !PO_RETIME.skus) return;
  document.getElementById('po-apply-dialog').close();
  doApplyRetime(PO_RETIME, 'whole year');
}
function applySupplierRetime() {         // from the Plan supplier-header button (current supplier only)
  const status = document.getElementById('save-status');
  if (!poDataReady()) { status.textContent = 'Upload PO + container files first (Settings → Data)'; return; }
  const plan = computeRetime('sup');
  if (!plan.skus) { status.textContent = `Nothing to re-time for ${titleCase(currentSupplier)} — every PO'd product already matches its PO weeks`; return; }
  doApplyRetime(plan, titleCase(currentSupplier));
}
function undoRetime() {
  if (!PO_APPLY_UNDO) return;
  if (!PROPOSED) PROPOSED = new Map();
  for (const id in PO_APPLY_UNDO.snapO) ORDERS[id] = PO_APPLY_UNDO.snapO[id].slice();
  for (const id in PO_APPLY_UNDO.snapP) PROPOSED.set(id, PO_APPLY_UNDO.snapP[id].slice());
  computeAll(); markDirty();
  PO_APPLY_UNDO = null;
  const t = document.getElementById('po-retime-toast'); if (t) t.remove();
  renderSidebar();
  if (currentView === 'plan') renderPlan(); else setView(currentView);
  document.getElementById('save-status').textContent = 'Re-time undone — committed + proposed restored';
}
// Small floating Undo affordance after applying (mirrors the review bar; lives on body).
function showRetimeToast() {
  let t = document.getElementById('po-retime-toast');
  if (!t) {
    t = document.createElement('div'); t.id = 'po-retime-toast'; document.body.appendChild(t);
    t.addEventListener('click', ev => {
      const act = ev.target.dataset.act;
      if (act === 'undo') undoRetime();
      else if (act === 'apply') commitRetime();
      else if (act === 'dismiss') t.remove();
    });
  }
  t.innerHTML = `<span class="rt-msg">✓ Re-timed ${PO_APPLY_UNDO ? esc(PO_APPLY_UNDO.detail || PO_APPLY_UNDO.label) : ''}</span>`
    + `<button type="button" data-act="apply" class="rt-apply" title="Accept these changes and save them now — the sheet will reopen with them applied">✓ Apply &amp; save</button>`
    + `<button type="button" data-act="undo">↶ Undo</button><button type="button" data-act="dismiss" class="prv-done" title="Hide this — changes stay and autosave">✕</button>`;
}
// Accept the re-time: persist immediately and finalise (drop the undo affordances).
function commitRetime() {
  saveNow();                                     // write orders + proposed to disk now (clears the pending autosave)
  PO_APPLY_UNDO = null;                           // accepted — no longer offer Undo
  document.getElementById('btn-sup-retime-undo')?.remove();
  const pu = document.getElementById('btn-po-undo'); if (pu) pu.hidden = true;
  const t = document.getElementById('po-retime-toast'); if (t) t.remove();
}
// Per-SKU price function for a pricing basis ('original' = current app price, else the
// uploaded map with a fall-back to the current price for products not in the file).
function aspMapFor(basis) {
  if (basis === 'basic') return s => (ASP_PARSED.basic[s.code] != null ? ASP_PARSED.basic[s.code] : (+s.asp || 0));
  if (basis === 'top')   return s => (ASP_PARSED.top[s.code]   != null ? ASP_PARSED.top[s.code]   : (+s.asp || 0));
  return s => (+s.asp || 0);
}
// Whole-loaded-year totals under a given price function: average realised price per sold
// unit, stock-capped sales value (what you'd actually sell given projected stock) and the
// uncapped forecast plan value. Sold units are price-independent, so this isolates ASP.
function aspScenarioTotals(aspOf) {
  const cur = SETTINGS.current_week;
  let sales = 0, plan = 0, units = 0;
  for (const s of M.skus) {
    const r = computeSku(s, combinedOrder(s.id));
    let sold = 0, pl = 0;
    for (let w = 0; w < WEEKS; w++) {
      sold += (w + 1) < cur ? (s.actual[w] || 0) : Math.min(r.forecast[w], r.stock[w]);
      pl += r.forecast[w];
    }
    const a = aspOf(s) || 0; sales += sold * a; plan += pl * a; units += sold;
  }
  return { avg: units ? sales / units : 0, sales, plan };
}
function openAspDialog() {
  const p = ASP_PARSED;
  const matched = M.skus.filter(s => p.basic[s.code] != null).length;
  document.getElementById('asp-summary').innerHTML =
    `<p>From <b>${esc(p.fname)}</b>: <b>${p.withSales}</b> products with sales (of ${p.fileSkus} in the file). `
    + `<b>${matched}</b> of ${YEAR}'s ${M.skus.length} products matched; the rest keep their current price.</p>`
    + `<p class="muted-note">Totals are for <b>${YEAR}</b> — "sales" is stock-capped (what you'd actually sell given projected stock). <b>Click a row</b> to choose that pricing basis.</p>`;
  const bases = [
    { key: 'original', label: 'Original — current prices in app' },
    { key: 'basic', label: 'Basic average — all channels (volume-weighted)' },
    { key: 'top', label: 'Highest-selling customer — price at the top channel' },
  ];
  const base = aspScenarioTotals(aspMapFor('original')).sales;
  const rows = bases.map(b => {
    const t = aspScenarioTotals(aspMapFor(b.key));
    const d = b.key === 'original' ? null : t.sales - base, dPct = (d != null && base) ? d / base * 100 : 0;
    const vs = d == null ? '<span class="muted-note">baseline</span>'
      : `<span class="${d > 0 ? 'c-up' : d < 0 ? 'c-down' : ''}">${(d >= 0 ? '+' : '') + fmtGBPk(d)} (${(dPct >= 0 ? '+' : '') + dPct.toFixed(1)}%)</span>`;
    return `<tr class="asp-cmp-row" data-basis="${b.key}" tabindex="0"><td class="acl">${b.label}</td>`
      + `<td>£${t.avg.toFixed(2)}</td><td><b>${fmtGBPk(t.sales)}</b></td><td>${fmtGBPk(t.plan)}</td><td>${vs}</td></tr>`;
  }).join('');
  document.getElementById('asp-compare').innerHTML =
    `<table class="asp-cmp"><thead><tr><th>Pricing basis</th><th>Avg £/unit</th><th>Sales (stock‑capped)</th><th>Forecast plan</th><th>vs current</th></tr></thead><tbody>${rows}</tbody></table>`;
  ASP_CHOICE = null;
  document.getElementById('asp-detail').classList.add('hidden');
  document.getElementById('asp-preview').innerHTML = '';
  document.getElementById('asp-apply').disabled = true;
  document.getElementById('asp-dialog').showModal();
}
// User clicked a comparison row → lock that basis and (for an actual update) reveal the
// year picker + the largest price movers, then enable Apply.
function aspChooseBasis(basis) {
  ASP_CHOICE = basis;
  document.querySelectorAll('#asp-compare .asp-cmp-row').forEach(tr => tr.classList.toggle('sel', tr.dataset.basis === basis));
  const detail = document.getElementById('asp-detail'), apply = document.getElementById('asp-apply');
  if (basis === 'original') {   // keep current prices → nothing to apply
    detail.classList.add('hidden');
    document.getElementById('asp-preview').innerHTML = '<div class="muted-note">“Original” keeps your current prices — nothing to apply. Pick Basic or Top‑channel to update.</div>';
    detail.classList.remove('hidden');
    apply.disabled = true;
    return;
  }
  const map = basis === 'top' ? ASP_PARSED.top : ASP_PARSED.basic;
  document.getElementById('asp-years').innerHTML = YEARS.slice().sort().map(y =>
    `<label class="asp-yr"><input type="checkbox" value="${y}"${y >= String(YEAR) ? ' checked' : ''}> ${y}</label>`).join('');
  const changes = M.skus.filter(s => map[s.code] != null).map(s => {
    const cur = +(s.asp || 0), neu = map[s.code]; return { code: s.code, cur, neu, pct: cur ? (neu - cur) / cur * 100 : 0 };
  }).sort((a, b) => Math.abs(b.pct) - Math.abs(a.pct));
  const rows = changes.slice(0, 12).map(c =>
    `<tr><td>${esc(c.code)}</td><td>£${c.cur.toFixed(2)}</td><td>£${c.neu.toFixed(2)}</td>`
    + `<td class="${c.pct >= 0 ? 'c-up' : 'c-down'}">${(c.pct >= 0 ? '+' : '') + c.pct.toFixed(1)}%</td></tr>`).join('');
  document.getElementById('asp-preview').innerHTML =
    `<div class="asp-prev-note">Largest price moves under <b>${basis === 'top' ? 'highest-selling customer' : 'basic average'}</b> (vs current):</div>`
    + `<table class="asp-prev"><thead><tr><th>SKU</th><th>Current</th><th>New</th><th>&Delta;</th></tr></thead><tbody>${rows}</tbody></table>`;
  detail.classList.remove('hidden');
  apply.disabled = false;
}
async function applyAspUpdates() {
  if (!ASP_PARSED || !ASP_CHOICE || ASP_CHOICE === 'original') return;
  const map = ASP_CHOICE === 'top' ? ASP_PARSED.top : ASP_PARSED.basic;
  const years = [...document.querySelectorAll('#asp-years input:checked')].map(i => i.value);
  if (!years.length) { alert('Pick at least one year to apply to.'); return; }
  const status = document.getElementById('save-status');
  status.textContent = 'Updating ASPs…';
  try {
    const j = await postAsp(map, years, 'upload');
    if (!j.ok) { status.textContent = ''; alert('Update failed: ' + (j.error || 'unknown')); return; }
    const n = applyAspToMemory(map, years, 'upload');   // reflect immediately if loaded year was included
    SETTINGS.asp_updated_at = new Date().toISOString(); markDirty();   // stamp the ASP upload time
    renderUploadAges();
    document.getElementById('asp-dialog').close();
    document.getElementById('settings-dialog').close();
    status.textContent = `ASP updated (${(j.years || []).join(', ')})`;
    alert(`Average Selling Prices updated — ${ASP_CHOICE === 'top' ? 'highest-selling customer' : 'basic average'} basis.\n\n`
      + `Applied to ${(j.years || []).join(', ')}${years.includes(String(YEAR)) ? ` — ${n} products in ${YEAR}` : ''}.`);
    ASP_PARSED = null; ASP_CHOICE = null;
  } catch (err) { status.textContent = ''; alert('Update error: ' + err.message); }
}
// POST a {code:asp} map to the server for the given years, tagging the source.
function postAsp(aspMap, years, src) {
  return fetch('/api/apply-asp', { method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ asp: aspMap, years, src }) }).then(r => r.json());
}
// Mirror an ASP change into the in-memory loaded year (so it shows without a reload).
function applyAspToMemory(aspMap, years, src) {
  if (!years.includes(String(YEAR))) return 0;
  let n = 0;
  for (const s of M.skus) { const a = aspMap[s.code]; if (a != null && a > 0) { s.asp_prev = (+s.asp || 0); s.asp = a; s.asp_src = src; n++; } }
  computeAll(); renderSidebar();
  if (currentView === 'plan') renderPlan(); else setView(currentView);
  return n;
}
/* ---------------- update FOB + landed costs from a "Landed Costs" export ---------- */
let LANDED_PARSED = null;   // { map:{code:{fob,landed,fobLast,landedLast}}, fileSkus, fname }
let LANDED_EDITS = {};      // code -> {fob?, landed?}  manual overrides made in the preview
let LANDED_MANUAL = new Set();
// Effective imported FOB for a product: current outstanding (col H), else last-receipted (col E).
function landedEffFob(m) { return m.fob != null ? m.fob : (m.fobLast != null ? m.fobLast : null); }
async function landedFileChosen(e) {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  const status = document.getElementById('landed-status');
  status.textContent = 'Reading ' + file.name + '…';
  try {
    const r = await fetch('/api/parse-landed', { method: 'POST', body: await file.arrayBuffer() });
    const j = await r.json();
    if (!j.ok) { status.textContent = ''; alert('Could not read the file: ' + (j.error || 'unknown')); return; }
    LANDED_PARSED = { map: j.landed || {}, fileSkus: j.fileSkus, fname: file.name };
    LANDED_EDITS = {}; LANDED_MANUAL = new Set();
    status.textContent = '';
    openLandedDialog();
  } catch (err) { status.textContent = ''; alert('Read error: ' + err.message); }
}
function openLandedDialog() {
  const p = LANDED_PARSED; if (!p) return;
  document.getElementById('landed-years').innerHTML = YEARS.slice().sort().map(y =>
    `<label class="asp-yr"><input type="checkbox" value="${y}"${y >= String(YEAR) ? ' checked' : ''}> ${y}</label>`).join('');
  const matched = M.skus.filter(s => p.map[s.code]);
  let fobN = 0, lndN = 0, fbN = 0;
  for (const s of matched) {
    const m = p.map[s.code], ef = landedEffFob(m);
    if (ef != null && Math.abs(ef - (+s.fob || 0)) > 0.005) fobN++;
    if (m.landed != null && Math.abs(m.landed - (+s.landed || 0)) > 0.005) lndN++;
    if (m.fob == null && m.fobLast != null) fbN++;
  }
  document.getElementById('landed-summary').innerHTML =
    `<p>From <b>${esc(p.fname)}</b>: <b>${matched.length}</b> of ${YEAR}'s ${M.skus.length} products matched by code (${p.fileSkus} in the file).</p>`
    + `<p class="muted-note">FOB (USD) &larr; current outstanding, or last-receipted where nothing is outstanding (<b>${fbN}</b>). Landed (&pound;) &larr; historical average. `
    + `<b>${fobN}</b> FOB and <b>${lndN}</b> landed values differ from the app. Edit any cell to override it (&ldquo;manual&rdquo;); untouched values are tagged &ldquo;from import&rdquo;.</p>`;
  renderLandedPreview();
  document.getElementById('landed-apply').disabled = matched.length === 0;
  document.getElementById('landed-dialog').showModal();
}
function renderLandedPreview() {
  const p = LANDED_PARSED; if (!p) return;
  const q = (document.getElementById('landed-search').value || '').trim().toLowerCase();
  const rows = M.skus.filter(s => p.map[s.code])
    .filter(s => !q || s.code.toLowerCase().includes(q) || (s.name || '').toLowerCase().includes(q));
  const pctCell = (neu, cur) => {
    if (neu == null || !cur) return '<span class="lc-d">—</span>';
    const pct = (neu - cur) / cur * 100;
    return `<span class="lc-d ${pct > 0.05 ? 'c-up' : pct < -0.05 ? 'c-down' : ''}">${(pct >= 0 ? '+' : '') + pct.toFixed(1)}%</span>`;
  };
  const body = rows.slice(0, 800).map(s => {
    const m = p.map[s.code], ed = LANDED_EDITS[s.code] || {};
    const impFob = landedEffFob(m), impLnd = m.landed != null ? m.landed : null;
    const newFob = ed.fob != null ? ed.fob : impFob, newLnd = ed.landed != null ? ed.landed : impLnd;
    const curF = +s.fob || 0, curL = +s.landed || 0;
    const fb = m.fob == null && m.fobLast != null;
    const man = LANDED_MANUAL.has(s.code);
    return `<div class="lc-row${man ? ' lc-manual' : ''}" data-code="${esc(s.code)}">`
      + `<span class="lc-code">${esc(s.code)}${fb ? ' <span class="lc-fb" title="No outstanding order — using last-receipted FOB (col E)">E</span>' : ''}</span>`
      + `<span class="lc-name" title="${esc(s.name || '')}">${esc(s.name || '')}</span>`
      + `<span class="lc-cur">$${curF.toFixed(2)}</span>`
      + `<input class="lc-in" data-k="fob" data-code="${esc(s.code)}" value="${newFob != null ? Number(newFob).toFixed(2) : ''}" inputmode="decimal">`
      + pctCell(newFob, curF)
      + `<span class="lc-cur">£${curL.toFixed(2)}</span>`
      + `<input class="lc-in" data-k="landed" data-code="${esc(s.code)}" value="${newLnd != null ? Number(newLnd).toFixed(2) : ''}" inputmode="decimal">`
      + pctCell(newLnd, curL)
      + `</div>`;
  }).join('');
  document.getElementById('landed-list').innerHTML = body || '<div class="muted-note">No matching products.</div>';
  document.getElementById('landed-count').textContent = `${rows.length} product(s)`;
}
// Record a manual override typed into the preview (persists across filtering).
function landedEditInput(inp) {
  const code = inp.dataset.code, k = inp.dataset.k;
  const v = parseFloat(String(inp.value).replace(/[^0-9.]/g, ''));
  LANDED_EDITS[code] = LANDED_EDITS[code] || {};
  if (isFinite(v) && v > 0) LANDED_EDITS[code][k] = +v.toFixed(2); else delete LANDED_EDITS[code][k];
  LANDED_MANUAL.add(code);
  const row = inp.closest('.lc-row'); if (row) row.classList.add('lc-manual');
}
async function applyLandedUpdates() {
  const p = LANDED_PARSED; if (!p) return;
  const years = [...document.querySelectorAll('#landed-years input:checked')].map(i => i.value);
  if (!years.length) { alert('Pick at least one year to apply to.'); return; }
  const fob = {}, landed = {};
  for (const s of M.skus) {
    const m = p.map[s.code]; if (!m) continue;
    const ed = LANDED_EDITS[s.code] || {};
    const ef = ed.fob != null ? ed.fob : landedEffFob(m);
    const ln = ed.landed != null ? ed.landed : (m.landed != null ? m.landed : null);
    if (ef != null && ef > 0) fob[s.code] = +Number(ef).toFixed(2);
    if (ln != null && ln > 0) landed[s.code] = +Number(ln).toFixed(2);
  }
  const manual = [...LANDED_MANUAL];
  const status = document.getElementById('save-status'); status.textContent = 'Updating costs…';
  try {
    const r = await fetch('/api/apply-landed', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fob, landed, manual, years }) });
    const j = await r.json();
    if (!j.ok) { status.textContent = ''; alert('Update failed: ' + (j.error || 'unknown')); return; }
    const n = applyLandedToMemory(fob, landed, new Set(manual), years);
    SETTINGS.landed_updated_at = new Date().toISOString(); markDirty(); renderUploadAges();
    document.getElementById('landed-dialog').close();
    document.getElementById('settings-dialog').close();
    status.textContent = `FOB & landed costs updated (${(j.years || []).join(', ')})`;
    alert(`FOB & landed costs updated.\n\nApplied to ${(j.years || []).join(', ')}${years.includes(String(YEAR)) ? ` — ${n} products in ${YEAR}` : ''}.`
      + (manual.length ? `\n${manual.length} product(s) manually overridden.` : ''));
    LANDED_PARSED = null; LANDED_EDITS = {}; LANDED_MANUAL = new Set();
  } catch (err) { status.textContent = ''; alert('Update error: ' + err.message); }
}
// Mirror a cost update into the in-memory loaded year (so it shows without a reload).
function applyLandedToMemory(fobMap, landedMap, manualSet, years) {
  if (!years.includes(String(YEAR))) return 0;
  let n = 0;
  for (const s of M.skus) {
    let touched = false;
    if (fobMap[s.code] != null) { s.fob_prev = (+s.fob || 0); s.fob = fobMap[s.code]; s.fob_src = manualSet.has(s.code) ? 'manual' : 'upload'; touched = true; }
    if (landedMap[s.code] != null) { s.landed_prev = (+s.landed || 0); s.landed = landedMap[s.code]; s.landed_src = manualSet.has(s.code) ? 'manual' : 'upload'; touched = true; }
    if (touched) n++;
  }
  computeAll(); renderSidebar();
  if (currentView === 'plan') renderPlan(); else setView(currentView);
  return n;
}
// Cost provenance for the FOB / Landed figures on the plan: 'upload' (from a cost file),
// 'manual' (overridden), or 'orig' (never updated from a cost file).
function costSrc(sku, which) { const v = sku[which + '_src']; return (v === 'upload' || v === 'manual') ? v : 'orig'; }

/* ------------- update catalogue status + outstanding purchases from Buying Report ------------ */
let BUYING_PARSED = null;   // { map:{code:{status,osPurchases}}, fileSkus, fname }
async function buyingFileChosen(e) {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  const status = document.getElementById('buying-status');
  status.textContent = 'Reading ' + file.name + '…';
  try {
    const r = await fetch('/api/parse-buying', { method: 'POST', body: await file.arrayBuffer() });
    const j = await r.json();
    if (!j.ok) { status.textContent = ''; alert('Could not read the file: ' + (j.error || 'unknown')); return; }
    BUYING_PARSED = { map: j.buying || {}, fileSkus: j.fileSkus, fname: file.name };
    status.textContent = '';
    openBuyingDialog();
  } catch (err) { status.textContent = ''; alert('Read error: ' + err.message); }
}
function openBuyingDialog() {
  const p = BUYING_PARSED; if (!p) return;
  const matched = M.skus.filter(s => p.map[s.code]);
  const changes = matched.filter(s => p.map[s.code].status && p.map[s.code].status !== s.status)
    .map(s => ({ code: s.code, name: s.name || '', from: s.status || 'Unknown', to: p.map[s.code].status }));
  let live = 0, notlive = 0, ospN = 0, stkN = 0;
  for (const s of matched) {
    const r = p.map[s.code];
    if (r.status === 'Live') live++; else if (r.status === 'Not Live') notlive++;
    if (r.osPurchases != null && r.osPurchases !== (s.os_purchases || 0)) ospN++;
    if (r.stock != null && r.stock !== (+s.stock_now || 0)) stkN++;
  }
  document.getElementById('buying-years').innerHTML = YEARS.slice().sort().map(y =>
    `<label class="asp-yr"><input type="checkbox" value="${y}"${y >= String(YEAR) ? ' checked' : ''}> ${y}</label>`).join('');
  document.getElementById('buying-summary').innerHTML =
    `<p>From <b>${esc(p.fname)}</b>: <b>${matched.length}</b> of ${YEAR}'s ${M.skus.length} products matched (WEBSA rows only; ${p.fileSkus} in the file).</p>`
    + `<p class="muted-note">Status in file: <b>${live}</b> Live · <b>${notlive}</b> Not Live. <b>${changes.length}</b> status change(s), <b>${ospN}</b> outstanding-purchase update(s), <b>${stkN}</b> live-stock update(s) (Stock now refreshes from this report).</p>`;
  const badgeCls = st => st === 'Live' ? 'live' : st === 'Not Live' ? 'notlive' : 'unknown';
  const rows = changes.slice(0, 400).map(c =>
    `<div class="bc-row"><span class="bc-code">${esc(c.code)}</span><span class="bc-name" title="${esc(c.name)}">${esc(c.name)}</span>`
    + `<span class="badge ${badgeCls(c.from)}">${esc(c.from)}</span><span class="bc-arrow">→</span>`
    + `<span class="badge ${badgeCls(c.to)}">${esc(c.to)}</span></div>`).join('');
  document.getElementById('buying-changes').innerHTML = changes.length
    ? `<div class="bc-headline">Catalogue status changes (${changes.length})</div><div class="bc-list">${rows}</div>`
    : '<p class="muted-note">No catalogue-status changes — matched products already match the report. Outstanding purchases will still update.</p>';
  document.getElementById('buying-apply').disabled = matched.length === 0;
  document.getElementById('buying-dialog').showModal();
}
async function applyBuyingUpdates() {
  const p = BUYING_PARSED; if (!p) return;
  const years = [...document.querySelectorAll('#buying-years input:checked')].map(i => i.value);
  if (!years.length) { alert('Pick at least one year to apply to.'); return; }
  const status = document.getElementById('save-status'); status.textContent = 'Updating catalogue…';
  try {
    const r = await fetch('/api/apply-buying', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ buying: p.map, years }) });
    const j = await r.json();
    if (!j.ok) { status.textContent = ''; alert('Update failed: ' + (j.error || 'unknown')); return; }
    const n = applyBuyingToMemory(p.map, years);
    SETTINGS.buying_updated_at = new Date().toISOString(); markDirty(); renderUploadAges();
    document.getElementById('buying-dialog').close();
    document.getElementById('settings-dialog').close();
    status.textContent = `Catalogue status & purchases updated (${(j.years || []).join(', ')})`;
    alert(`Catalogue status & outstanding purchases updated.\n\nApplied to ${(j.years || []).join(', ')}${years.includes(String(YEAR)) ? ` — ${n} products in ${YEAR}` : ''}.`);
    BUYING_PARSED = null;
  } catch (err) { status.textContent = ''; alert('Update error: ' + err.message); }
}
function applyBuyingToMemory(map, years) {
  if (!years.includes(String(YEAR))) return 0;
  let n = 0;
  for (const s of M.skus) {
    const rec = map[s.code]; if (!rec) continue;
    let touched = false;
    if (rec.status) { s.status = rec.status; touched = true; }
    if (rec.osPurchases != null) { s.os_purchases = rec.osPurchases; touched = true; }
    if (rec.stock != null) { s.stock_now = rec.stock; touched = true; }   // live warehouse stock
    if (touched) n++;
  }
  computeAll(); renderSidebar();
  if (currentView === 'plan') renderPlan(); else setView(currentView);
  return n;
}

/* ---------- update per-product duty rates from the Tradeplan "Landed Costs" sheet ---------- */
let DUTY_PARSED = null;   // { map:{code:pct}, fileSkus, fname }
async function dutyFileChosen(e) {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  const status = document.getElementById('duty-status');
  status.textContent = 'Reading ' + file.name + '…';
  try {
    const r = await fetch('/api/parse-duty', { method: 'POST', body: await file.arrayBuffer() });
    const j = await r.json();
    if (!j.ok) { status.textContent = ''; alert('Could not read the file: ' + (j.error || 'unknown')); return; }
    DUTY_PARSED = { map: j.duty || {}, fileSkus: j.fileSkus, fname: file.name };
    status.textContent = '';
    openDutyDialog();
  } catch (err) { status.textContent = ''; alert('Read error: ' + err.message); }
}
function openDutyDialog() {
  const p = DUTY_PARSED; if (!p) return;
  const matched = M.skus.filter(s => p.map[s.code] != null);
  const changes = matched.filter(s => Math.abs((+p.map[s.code]) - (+s.duty_rate || 0)) > 0.005)
    .map(s => ({ code: s.code, name: s.name || '', from: (+s.duty_rate || 0), to: +p.map[s.code] }));
  const nonzero = matched.filter(s => +p.map[s.code] > 0).length;
  document.getElementById('duty-years').innerHTML = YEARS.slice().sort().map(y =>
    `<label class="asp-yr"><input type="checkbox" value="${y}"${y >= String(YEAR) ? ' checked' : ''}> ${y}</label>`).join('');
  document.getElementById('duty-summary').innerHTML =
    `<p>From <b>${esc(p.fname)}</b>: <b>${matched.length}</b> of ${YEAR}'s ${M.skus.length} products matched (${p.fileSkus} in the file).</p>`
    + `<p class="muted-note"><b>${nonzero}</b> products carry a non-zero duty; <b>${changes.length}</b> rate change(s) vs the app.</p>`;
  const rows = changes.slice(0, 500).map(c =>
    `<div class="bc-row"><span class="bc-code">${esc(c.code)}</span><span class="bc-name" title="${esc(c.name)}">${esc(c.name)}</span>`
    + `<span class="bc-num">${c.from.toFixed(1)}%</span><span class="bc-arrow">→</span><span class="bc-num"><b>${c.to.toFixed(1)}%</b></span></div>`).join('');
  document.getElementById('duty-changes').innerHTML = changes.length
    ? `<div class="bc-headline">Duty rate changes (${changes.length})</div><div class="bc-list">${rows}</div>`
    : '<p class="muted-note">No changes — matched products already carry these duty rates.</p>';
  document.getElementById('duty-apply').disabled = matched.length === 0;
  document.getElementById('duty-dialog').showModal();
}
async function applyDutyUpdates() {
  const p = DUTY_PARSED; if (!p) return;
  const years = [...document.querySelectorAll('#duty-years input:checked')].map(i => i.value);
  if (!years.length) { alert('Pick at least one year to apply to.'); return; }
  const status = document.getElementById('save-status'); status.textContent = 'Updating duty rates…';
  try {
    const r = await fetch('/api/apply-duty', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ duty: p.map, years }) });
    const j = await r.json();
    if (!j.ok) { status.textContent = ''; alert('Update failed: ' + (j.error || 'unknown')); return; }
    const n = applyDutyToMemory(p.map, years);
    SETTINGS.duty_updated_at = new Date().toISOString(); markDirty(); renderUploadAges();
    document.getElementById('duty-dialog').close();
    document.getElementById('settings-dialog').close();
    status.textContent = `Duty rates updated (${(j.years || []).join(', ')})`;
    alert(`Duty rates updated.\n\nApplied to ${(j.years || []).join(', ')}${years.includes(String(YEAR)) ? ` — ${n} products in ${YEAR}` : ''}.`);
    DUTY_PARSED = null;
  } catch (err) { status.textContent = ''; alert('Update error: ' + err.message); }
}
function applyDutyToMemory(map, years) {
  if (!years.includes(String(YEAR))) return 0;
  let n = 0;
  for (const s of M.skus) { const r = map[s.code]; if (r != null) { s.duty_rate = +r; n++; } }
  if (currentView === 'plan') renderPlan(); else setView(currentView);   // refresh the est-landed chips
  return n;
}

/* ---------- weekly actual sales upload (the "WKnn Sales" export) ----------
   File = Product SKU + Sales TY (£) + Qty TY (units) for ONE week. Applies to the
   VIEWED year only: sets each matched SKU's actual[week] (units straight from the
   file's Qty TY), stamps its weekly ASP (Sales TY / Qty TY, shown as a chip beside
   the main ASP pill), closes the week's stock (running_stock = stock_now + arrivals
   − sold, so the app trusts its own imports without waiting for import_data), and
   advances data_week to week+1 — which is what flips that week from forecast to
   actuals in the value/YTD/outturn boundaries. */
let WKSALES_PARSED = null;   // { map:{code:{val,qty}}, fileRows, fname, week }
async function wksalesFileChosen(e) {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  const status = document.getElementById('wksales-status');
  status.textContent = 'Reading ' + file.name + '…';
  try {
    const r = await fetch('/api/parse-wksales', { method: 'POST', body: await file.arrayBuffer() });
    const j = await r.json();
    if (!j.ok) { status.textContent = ''; alert('Could not read the file: ' + (j.error || 'unknown')); return; }
    const mWk = /wk\s*0?(\d{1,2})/i.exec(file.name);   // "WK28 Sales.xlsx" → 28
    const week = mWk ? +mWk[1] : (+M.data_week || SETTINGS.current_week || 1);
    WKSALES_PARSED = { map: j.sales || {}, fileRows: j.fileRows, fname: file.name, week };
    status.textContent = '';
    openWksalesDialog();
  } catch (err) { status.textContent = ''; alert('Read error: ' + err.message); }
}
// Per matched SKU: units straight from the file's Qty TY, plus that week's realised
// ASP (Sales TY / Qty TY) for the chip beside the main ASP pill.
function wksalesUnits() {
  const p = WKSALES_PARSED, units = {}, wkasp = {};
  let matched = 0, totalVal = 0, totalQty = 0;
  for (const s of M.skus) {
    const r = p.map[s.code];
    if (r == null) continue;
    matched++;
    units[s.code] = r.qty;
    totalQty += r.qty;
    totalVal += r.val || 0;
    if (r.qty > 0 && r.val > 0) wkasp[s.code] = Math.round((r.val / r.qty) * 100) / 100;
  }
  return { units, wkasp, matched, totalVal, totalQty };
}
function openWksalesDialog() {
  const p = WKSALES_PARSED; if (!p) return;
  const { matched, totalVal, totalQty } = wksalesUnits();
  const appCodes = new Set(M.skus.map(s => s.code));
  const unmatched = Object.keys(p.map).filter(c => !appCodes.has(c));
  const wkSel = document.getElementById('wksales-week');
  wkSel.innerHTML = Array.from({ length: WEEKS }, (_, i) =>
    `<option value="${i + 1}"${i + 1 === p.week ? ' selected' : ''}>Week ${i + 1} · w/c ${weekDate(i + 1)}</option>`).join('');
  document.getElementById('wksales-summary').innerHTML =
    `<p>From <b>${esc(p.fname)}</b>: <b>${matched}</b> of ${YEAR}'s ${M.skus.length} products matched `
    + `(${p.fileRows} rows in the file) · <b>${Math.round(totalQty).toLocaleString('en-GB')}</b> units / <b>${fmtGBP(totalVal)}</b> to apply to <b>${esc(YEAR)}</b>.</p>`
    + `<p class="muted-note">Units come straight from the file's <b>Qty TY</b>; each seller's weekly ASP (Sales ÷ Qty) is stamped `
    + `on its ASP chip. Each product's closing stock for the week is chained as <b>previous week's closing + arrivals − sales</b> `
    + `(re-uploading a past week re-chains every later week too). Products not in the file keep their existing sales for the `
    + `chosen week (0 for a new week).</p>`
    + (unmatched.length ? `<details class="muted-note"><summary>${unmatched.length} file code(s) not in the ${esc(YEAR)} plan</summary>${esc(unmatched.join(', '))}</details>` : '');
  wksalesWeekNote();
  document.getElementById('wksales-apply').disabled = matched === 0;
  document.getElementById('wksales-dialog').showModal();
}
function wksalesWeekNote() {
  const wk = +document.getElementById('wksales-week').value;
  const dw = +M.data_week || 1;
  const el = document.getElementById('wksales-week-note');
  if (wk < dw) el.innerHTML = `<b>Week ${wk} already has actuals</b> — applying will overwrite them.`;
  else if (wk > dw) el.innerHTML = `Heads-up: the next week expecting actuals is <b>week ${dw}</b> — applying week ${wk} leaves ${wk - dw === 1 ? `week ${dw}` : `weeks ${dw}–${wk - 1}`} un-actualised.`;
  else el.innerHTML = `Week ${wk} is the next week expecting actuals — the actuals boundary will advance to week ${wk + 1}.`;
}
async function applyWksalesUpdates() {
  const p = WKSALES_PARSED; if (!p) return;
  const week = +document.getElementById('wksales-week').value;
  const { units, wkasp, matched } = wksalesUnits();
  if (!matched) return;
  const status = document.getElementById('save-status'); status.textContent = 'Saving weekly sales…';
  try {
    const r = await fetch('/api/apply-wksales?year=' + encodeURIComponent(YEAR),
      { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ week, units, wkasp }) });
    const j = await r.json();
    if (!j.ok) { status.textContent = ''; alert('Update failed: ' + (j.error || 'unknown')); return; }
    applyWksalesToMemory(units, wkasp, week, j.dataWeek, j.closedThrough);
    SETTINGS.wksales_updated_at = new Date().toISOString(); markDirty(); renderUploadAges();
    document.getElementById('wksales-dialog').close();
    document.getElementById('settings-dialog').close();
    status.textContent = `Week ${week} actual sales saved (${YEAR})`;
    alert(`Week ${week} actual sales saved for ${YEAR}.\n\n${j.applied} products updated · actuals boundary now week ${j.dataWeek}.`);
    WKSALES_PARSED = null;
  } catch (err) { status.textContent = ''; alert('Update error: ' + err.message); }
}
function applyWksalesToMemory(units, wkasp, week, dataWeek, closedThrough) {
  for (const s of M.skus) {
    const u = units[s.code];
    if (!s.actual) s.actual = zeros();
    if (u != null) s.actual[week - 1] = u;
    if (wkasp[s.code] != null) { s.asp_wk = wkasp[s.code]; s.asp_wk_week = week; }
    // chain the closing stock (every sku; unsold = 0 sales) from the uploaded week
    // through the last actualised week — mirrors the server's apply_wksales
    if (!s.running_stock) s.running_stock = zeros();
    const ord = ORDERS[s.id] || EMPTY53;
    for (let w = week; w <= (closedThrough || week); w++) {
      const prev = w >= 2 ? (+s.running_stock[w - 2] || 0) : (+s.stock_now || 0);
      s.running_stock[w - 1] = Math.max(0, prev + (ord[w - 1] || 0) - (s.actual[w - 1] || 0));
    }
  }
  M.data_week = dataWeek;
  // live year: let the date-derived current week advance now the boundary allows it
  const dateWk = highlightWeek();
  if (dateWk) SETTINGS.current_week = Math.min(dateWk, M.data_week || dateWk);
  buildModeledForecasts();   // models anchor on this-year actuals — refresh them
  computeAll(); renderSidebar();
  if (currentView === 'plan') renderPlan(); else setView(currentView);
}

// ASP provenance for a SKU: 'upload' (sales file), 'manual' (hand-edited), or 'orig'
// (still the originally-imported price — not yet updated, needs review).
function aspSrc(sku) { return (sku.asp_src === 'upload' || sku.asp_src === 'manual') ? sku.asp_src : 'orig'; }
function aspChip(sku) {
  const src = aspSrc(sku);
  const lbl = { orig: 'not yet updated — click to set', upload: 'from sales upload', manual: 'manually set' }[src];
  return `<button class="asp-chip asp-${src}" data-asp-sku="${esc(sku.id)}" title="Average Selling Price — ${lbl}. Click to edit by hand.">ASP £${(+sku.asp || 0).toFixed(2)}</button>`;
}
// Display-only chip: the realised ASP from the last weekly-sales upload (Sales TY ÷
// Qty TY), tinted vs the main ASP so weekly price drift is visible at a glance.
function wkAspChip(sku) {
  const wa = +sku.asp_wk;
  if (!(wa > 0)) return '';
  const base = +sku.asp || 0;
  const pct = base > 0 ? ((wa - base) / base) * 100 : 0;
  const cls = pct > 2 ? 'wa-up' : pct < -2 ? 'wa-down' : 'wa-flat';
  const vs = base > 0 ? ` — ${pct >= 0 ? '+' : ''}${pct.toFixed(1)}% vs the ASP in use (£${base.toFixed(2)})` : '';
  return `<span class="wkasp-chip ${cls}" title="Realised ASP in the week-${sku.asp_wk_week || '?'} sales file (Sales ÷ Qty)${vs}.">Wk £${wa.toFixed(2)}</span>`;
}
// Years a manual price edit applies to: the year being viewed and all later (forecast) years.
function manualAspYears() { return YEARS.filter(y => y >= String(YEAR)); }
async function editAspInline(id) {
  const sku = skuById.get(id); if (!sku) return;
  const cur = +(sku.asp || 0);
  const raw = prompt(`Average Selling Price for ${sku.code}\n(${sku.name || ''})\n\nApplies to ${manualAspYears().join(', ')}.`, cur.toFixed(2));
  if (raw == null) return;
  const n = parseFloat(String(raw).replace(/[^0-9.]/g, ''));
  if (!isFinite(n) || n <= 0) { alert('Enter a positive price.'); return; }
  const years = manualAspYears();
  const j = await postAsp({ [sku.code]: +n.toFixed(2) }, years, 'manual');
  if (!j.ok) { alert('Could not save: ' + (j.error || 'unknown')); return; }
  applyAspToMemory({ [sku.code]: +n.toFixed(2) }, years, 'manual');
  document.getElementById('save-status').textContent = `ASP set for ${sku.code}`;
}
// FOB / Landed Cost chips — styled + click-to-edit like the ASP chip, coloured by provenance
// (green = from cost-file import, orange = manually set, amber = not yet updated).
function fobChip(sku) {
  const src = costSrc(sku, 'fob');
  const lbl = { orig: 'not yet updated — click to set', upload: 'from cost-file import', manual: 'manually set' }[src];
  return `<button class="cost-chip cc-${src}" data-cost-sku="${esc(sku.id)}" data-cost-k="fob" title="FOB item cost (USD) — ${lbl}. Click to edit by hand.">FOB $${(+sku.fob || 0).toFixed(2)}</button>`;
}
function landedChip(sku) {
  const src = costSrc(sku, 'landed');
  const lbl = { orig: 'not yet updated — click to set', upload: 'from cost-file import', manual: 'manually set' }[src];
  return `<button class="cost-chip cc-${src}" data-cost-sku="${esc(sku.id)}" data-cost-k="landed" title="Landed cost (GBP) — ${lbl}. Click to edit by hand.">Landed Cost £${(+sku.landed || 0).toFixed(2)}</button>`;
}
async function editCostInline(id, which) {
  const sku = skuById.get(id); if (!sku) return;
  const isFob = which === 'fob';
  const cur = +(isFob ? sku.fob : sku.landed) || 0;
  const label = isFob ? 'FOB item cost ($ USD)' : 'Landed cost (£ GBP)';
  const years = manualAspYears();
  const raw = prompt(`${label} for ${sku.code}\n(${sku.name || ''})\n\nApplies to ${years.join(', ')}.`, cur.toFixed(2));
  if (raw == null) return;
  const n = parseFloat(String(raw).replace(/[^0-9.]/g, ''));
  if (!isFinite(n) || n <= 0) { alert('Enter a positive value.'); return; }
  const v = +n.toFixed(2);
  const payload = isFob ? { fob: { [sku.code]: v } } : { landed: { [sku.code]: v } };
  try {
    const r = await fetch('/api/apply-landed', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ...payload, manual: [sku.code], years }) });
    const j = await r.json();
    if (!j.ok) { alert('Could not save: ' + (j.error || 'unknown')); return; }
  } catch (err) { alert('Save error: ' + err.message); return; }
  applyLandedToMemory(isFob ? { [sku.code]: v } : {}, isFob ? {} : { [sku.code]: v }, new Set([sku.code]), years);
  document.getElementById('save-status').textContent = `${isFob ? 'FOB' : 'Landed cost'} set for ${sku.code}`;
}

/* ---- carton dimensions → item CBM + pallet/stillage loading ----
   Item CBM = (Σ carton L×W×H ÷ 1,000,000) ÷ pack size. Pallet loading = cartons that
   fit per layer (best footprint orientation) × layers under the max stack height,
   scaled to items. A carton that can't sit on the pallet ⇒ Stillage (manual qty). */
const PALLET_DEFAULT = { l: 120, w: 100, h: 180, maxKg: 1000 };   // cm + max load weight (kg)
function palletDims() { const p = SETTINGS.pallet_dims || {}; return { l: +p.l || PALLET_DEFAULT.l, w: +p.w || PALLET_DEFAULT.w, h: +p.h || PALLET_DEFAULT.h, maxKg: +p.maxKg || PALLET_DEFAULT.maxKg }; }
function cartonList(sku) { return Array.isArray(sku.cartons) ? sku.cartons : []; }
function cartonMetrics(cartons, packSize) {
  const P = palletDims();
  const pack = Math.max(1, Math.round(packSize || 1));
  const valid = (cartons || []).filter(c => +c.l > 0 && +c.w > 0 && +c.h > 0);
  const N = valid.length;
  const vol = valid.reduce((a, c) => a + (+c.l) * (+c.w) * (+c.h), 0);   // cm³ total across the pack's cartons
  const cbm = N ? +((vol / 1e6) / pack).toFixed(4) : null;              // m³ per single item
  const packKg = valid.reduce((a, c) => a + (+c.kg || 0), 0);          // kg per pack (Σ carton weights)
  const kgPerItem = packKg > 0 ? packKg / pack : 0;                    // kg per single item
  let rep = null, area = -1;                                            // largest-footprint carton drives the fit
  for (const c of valid) { const a = (+c.l) * (+c.w); if (a > area) { area = a; rep = c; } }
  let fits = false, perPallet = 0;
  if (rep) {
    const perLayer = Math.max(Math.floor(P.l / rep.l) * Math.floor(P.w / rep.w),
                              Math.floor(P.l / rep.w) * Math.floor(P.w / rep.l));
    const layers = Math.floor(P.h / rep.h);
    if (perLayer > 0 && layers > 0) { fits = true; perPallet = perLayer * layers; }
  }
  const volumeQty = fits ? Math.max(1, Math.floor(perPallet * (pack / Math.max(1, N)))) : 0;   // by dimensions
  const volWeight = volumeQty * kgPerItem;                             // weight of a volume-full pallet
  const weightQty = (P.maxKg > 0 && kgPerItem > 0) ? Math.floor(P.maxKg / kgPerItem) : null;   // items the weight allows
  const effWeightQty = weightQty != null ? Math.max(0, Math.min(volumeQty, weightQty)) : volumeQty;
  const overweightPct = (P.maxKg > 0 && volWeight > P.maxKg) ? (volWeight / P.maxKg - 1) * 100 : 0;
  return { cbm, fits, palletQty: volumeQty, volumeQty, weightQty, effWeightQty, volWeight, kgPerItem,
           overweightPct, maxKg: P.maxKg, hasWeight: kgPerItem > 0, needsStillage: N > 0 && !fits, cartons: N, pack };
}
// The loading qty to use for warehouse spacing, given the weight-limit toggle.
function cartonLoadQty(m, weightLimited) {
  if (m.needsStillage) return null;                                    // stillage → manual qty
  return (weightLimited && m.hasWeight) ? m.effWeightQty : m.volumeQty;
}
// CBM chip — clickable like the cost chips; opens the carton editor. Provenance:
// 'calc' (from carton sizes), 'manual' (hand-set), else 'import'.
function cbmSrcTag(sku) { return sku.cbm_src === 'calc' ? 'calc' : sku.cbm_src === 'manual' ? 'manual' : 'import'; }
function cbmChip(sku) {
  const src = cbmSrcTag(sku);
  const lbl = { import: 'from the last import', manual: 'manually set', calc: 'calculated from carton sizes' }[src];
  const cc = src === 'calc' ? 'cc-calc' : src === 'manual' ? 'cc-manual' : 'cc-orig';
  const nC = cartonList(sku).length;
  const extra = nC ? ` · ${nC} carton${nC > 1 ? 's' : ''}${(+sku.pack_size > 1) ? ` · pack ${sku.pack_size}` : ''}` : '';
  const still = sku.pallet_type === 'Stillage' ? ' · Stillage' : '';
  return `<button class="cost-chip cbm-chip ${cc}" data-cbm-sku="${esc(sku.id)}" title="Item CBM (m³) — ${lbl}${extra}. Click to edit carton sizes &amp; pack.">CBM ${sku.cbm ? (+sku.cbm).toFixed(3) : '—'}${still}</button>`;
}
// ---- carton editor dialog ----
let CARTON_EDIT = null;
function openCartonDialog(id) {
  const sku = skuById.get(id); if (!sku) return;
  const cartons = cartonList(sku).map(c => ({ l: +c.l || '', w: +c.w || '', h: +c.h || '', kg: +c.kg || '' }));
  if (!cartons.length) cartons.push({ l: '', w: '', h: '', kg: '' });
  CARTON_EDIT = { id, code: sku.code, pack: +sku.pack_size || 1, cartons };
  document.getElementById('carton-title').textContent = `${sku.code} — ${sku.name || ''}`;
  document.getElementById('carton-pack').value = CARTON_EDIT.pack;
  document.getElementById('carton-weightlimit').checked = (sku.load_basis !== 'volume');   // default: weight-limited
  const P = palletDims();
  document.getElementById('carton-pallet-note').textContent = `Pallet space: ${P.l}×${P.w}×${P.h} cm, max ${P.maxKg} kg (edit in Settings → Capacity)`;
  const qtyInp = document.getElementById('carton-qty');
  qtyInp.value = (sku.fpq_src === 'manual' && sku.fpq) ? Math.round(sku.fpq) : '';
  document.getElementById('carton-msg').textContent = '';
  cartonRenderRows();
  document.getElementById('carton-dialog').showModal();
}
function cartonRenderRows() {
  const box = document.getElementById('carton-rows');
  box.innerHTML = CARTON_EDIT.cartons.map((c, i) => `<div class="carton-row" data-i="${i}">`
    + `<span class="carton-n">#${i + 1}</span>`
    + `<input type="number" class="carton-l" min="0" step="0.1" value="${c.l}" placeholder="L">`
    + `<input type="number" class="carton-w" min="0" step="0.1" value="${c.w}" placeholder="W">`
    + `<input type="number" class="carton-h" min="0" step="0.1" value="${c.h}" placeholder="H">`
    + `<input type="number" class="carton-kg" min="0" step="0.01" value="${c.kg}" placeholder="kg">`
    + `<button type="button" class="carton-del" data-i="${i}" title="Remove carton"${CARTON_EDIT.cartons.length > 1 ? '' : ' disabled'}>✕</button></div>`).join('');
  box.querySelectorAll('.carton-row input').forEach(inp => inp.addEventListener('input', cartonReadState));
  box.querySelectorAll('.carton-del').forEach(b => b.addEventListener('click', () => {
    CARTON_EDIT.cartons.splice(+b.dataset.i, 1); cartonRenderRows();
  }));
  cartonRecalc();
}
function cartonReadState() {
  CARTON_EDIT.pack = Math.max(1, parseInt(document.getElementById('carton-pack').value, 10) || 1);
  document.querySelectorAll('#carton-rows .carton-row').forEach(row => {
    const i = +row.dataset.i;
    CARTON_EDIT.cartons[i] = { l: parseFloat(row.querySelector('.carton-l').value) || 0,
      w: parseFloat(row.querySelector('.carton-w').value) || 0, h: parseFloat(row.querySelector('.carton-h').value) || 0,
      kg: parseFloat(row.querySelector('.carton-kg').value) || 0 };
  });
  cartonRecalc();
}
function cartonRecalc() {
  const m = cartonMetrics(CARTON_EDIT.cartons, CARTON_EDIT.pack);
  const out = document.getElementById('carton-calc');
  const wrap = document.getElementById('carton-qty-wrap');
  const lbl = document.getElementById('carton-qty-lbl');
  const qty = document.getElementById('carton-qty');
  const wlWrap = document.getElementById('carton-wl-wrap');
  const weightLimited = document.getElementById('carton-weightlimit').checked;
  if (m.cbm == null) {
    out.innerHTML = `<div class="muted-note">Enter at least one carton's dimensions to calculate CBM and loading.</div>`;
    wrap.hidden = true; wlWrap.hidden = true; return;
  }
  const cbmLine = `<div class="carton-cbm">Item CBM: <b>${m.cbm.toFixed(4)}</b> m³${m.hasWeight ? ` · item weight <b>${m.kgPerItem.toFixed(2)}</b> kg` : ''}</div>`;
  if (m.needsStillage) {
    out.innerHTML = `<div class="carton-flag stillage">⚠ Too large for a pallet — flagged <b>Stillage</b> storage.</div>` + cbmLine;
    lbl.textContent = 'Stillage loading qty (units) *'; wrap.hidden = false; qty.placeholder = 'units per stillage'; wlWrap.hidden = true;
    return;
  }
  // both loading figures: by volume (dimensions) and by the weight cap
  let lines = cbmLine + `<div class="carton-cbm">By volume: <b>${m.volumeQty.toLocaleString()}</b> units/pallet`;
  if (m.hasWeight) {
    lines += ` — weighs <b>${Math.round(m.volWeight).toLocaleString()}</b> kg`;
    lines += m.overweightPct > 0.5 ? ` <span class="carton-flag">(${m.overweightPct.toFixed(0)}% over the ${m.maxKg} kg limit)</span>` : ` (within the ${m.maxKg} kg limit)`;
    lines += `</div><div class="carton-cbm">By weight limit: <b>${m.effWeightQty.toLocaleString()}</b> units/pallet</div>`;
  } else {
    lines += `</div><div class="muted-note">Add carton weights to also cap loading by the pallet weight limit.</div>`;
  }
  const used = cartonLoadQty(m, weightLimited);
  lines += `<div class="carton-used">Used for spacing: <b>${used.toLocaleString()}</b> units/pallet <span class="carton-basis">(${weightLimited && m.hasWeight ? 'weight-limited' : 'by volume'})</span></div>`;
  out.innerHTML = lines;
  wlWrap.hidden = !m.hasWeight;
  lbl.textContent = 'Override loading qty (optional)'; wrap.hidden = false; qty.placeholder = String(used);
}
async function submitCartons() {
  cartonReadState();
  const sku = skuById.get(CARTON_EDIT.id); if (!sku) return;
  const msg = document.getElementById('carton-msg');
  const cartons = CARTON_EDIT.cartons.filter(c => c.l > 0 && c.w > 0 && c.h > 0);
  if (!cartons.length) { msg.textContent = 'Enter at least one carton with L, W and H.'; return; }
  const m = cartonMetrics(cartons, CARTON_EDIT.pack);
  const weightLimited = document.getElementById('carton-weightlimit').checked;
  const raw = (document.getElementById('carton-qty').value || '').trim();
  const override = raw ? Math.max(1, Math.round(parseFloat(raw))) : null;
  const loadingQty = override != null ? override : cartonLoadQty(m, weightLimited);
  if (m.needsStillage && loadingQty == null) { msg.textContent = 'Enter the stillage loading qty.'; return; }
  const years = manualAspYears();
  const payload = { code: sku.code, cartons, packSize: CARTON_EDIT.pack, cbm: m.cbm,
    palletType: m.needsStillage ? 'Stillage' : (sku.pallet_type === 'Racking' ? 'Racking' : 'Pallet'),
    fpq: loadingQty, fpqSrc: override != null ? 'manual' : 'calc', loadBasis: weightLimited ? 'weight' : 'volume', years };
  msg.textContent = 'Saving…';
  try {
    const r = await fetch('/api/apply-cartons', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    const j = await r.json();
    if (!j.ok) { msg.textContent = j.error || 'Save failed.'; return; }
    applyCartonsToMemory(sku.code, payload);
    document.getElementById('carton-dialog').close();
    document.getElementById('save-status').textContent = `Carton details saved for ${sku.code} (${years.join(', ')})`;
  } catch (e) { msg.textContent = 'Error: ' + e.message; }
}
function applyCartonsToMemory(code, p) {
  for (const s of M.skus) if (s.code === code) {
    s.cartons = p.cartons; s.pack_size = p.packSize; s.load_basis = p.loadBasis;
    if (p.cbm != null) { s.cbm = p.cbm; s.cbm_src = 'calc'; }
    s.pallet_type = p.palletType; if (p.fpq != null) s.fpq = p.fpq; s.fpq_src = p.fpqSrc;
  }
  computeAll(); renderSidebar();
  if (currentView === 'plan') renderPlan(); else setView(currentView);
}

/* ------------- estimated future landed cost (mirrors the old Excel "Landed Costs" calc) ------- */
const LC_DEFAULTS = { container_rate: 3000, fx: 1.31, full_container_cbm: 67, inland_rate: 1100, duty_pct: 0 };
function landedCalc() { return Object.assign({}, LC_DEFAULTS, SETTINGS.landed_calc || {}); }
// Build a product's estimated landed cost (£/unit) from FOB + freight → £ → +duty → +inland haulage.
function estLanded(sku) {
  const cbm = +sku.cbm || 0;
  if (cbm <= 0) return null;                                     // needs the product's CBM
  const P = landedCalc();
  const qtyPer = P.full_container_cbm / cbm;                     // units per full container
  if (!(qtyPer > 0)) return null;
  const freightPerItem = P.container_rate / qtyPer;             // sea freight $/unit
  const fob = +sku.fob || 0;
  const cf = fob > 0 ? Math.round((fob + freightPerItem) * 100) / 100 : (+sku.landed || 0); // cost+freight ($, rounded like the Excel) — fall back to current landed £ if no FOB
  const gbp = cf / (P.fx > 0 ? P.fx : 1);                        // → £
  const inlandPerItem = P.inland_rate / qtyPer;                 // UK haulage £/unit
  const dutyPct = sku.duty_rate != null ? +sku.duty_rate : (+P.duty_pct || 0);  // per-product rate, else global default
  const est = gbp * (1 + dutyPct / 100) + inlandPerItem;
  const cur = +sku.landed || 0;
  return { est, qtyPer, freightPerItem, gbp, inlandPerItem, cur, dutyPct,
           diff: est - cur, diffPct: cur ? (est / cur - 1) * 100 : null };
}
function estLandedChip(sku) {
  const e = estLanded(sku);
  if (!e) return `<span class="est-chip est-na" title="Estimated landed cost needs the product's CBM">Est. landed n/a</span>`;
  const dir = e.diffPct == null ? '' : e.diffPct > 0.5 ? 'est-dearer' : e.diffPct < -0.5 ? 'est-cheaper' : '';
  const pct = e.diffPct == null ? '' : ` <span class="est-diff">(${e.diffPct >= 0 ? '+' : ''}${e.diffPct.toFixed(1)}%)</span>`;
  const tip = `Estimated future landed cost £${e.est.toFixed(2)} vs current £${e.cur.toFixed(2)} — `
    + `FOB $${(+sku.fob || 0).toFixed(2)} + freight $${e.freightPerItem.toFixed(2)}/unit ÷ FX = £${e.gbp.toFixed(2)}, `
    + `+${e.dutyPct}% duty, +£${e.inlandPerItem.toFixed(2)} inland (${Math.round(e.qtyPer)} units/container)`;
  return `<span class="est-chip ${dir}" title="${esc(tip)}">Est. landed £${e.est.toFixed(2)}${pct}</span>`;
}
const ASP_SRC_LABEL = { orig: 'not updated', upload: 'upload', manual: 'manual' };
// Bulk manual-price editor (Settings → Data): a filterable list of products with an
// editable ASP each, tagged by source so not-yet-updated products stand out.
function renderAspEditor() {
  const list = document.getElementById('asp-edit-list'); if (!list || !M) return;
  const q = (document.getElementById('asp-edit-search').value || '').toLowerCase().trim();
  const staleOnly = document.getElementById('asp-edit-stale-only').checked;
  const rows = M.skus.filter(s => {
    if (staleOnly && aspSrc(s) !== 'orig') return false;
    if (q && !((s.code || '').toLowerCase().includes(q) || (s.name || '').toLowerCase().includes(q))) return false;
    return true;
  }).sort((a, b) => (aspSrc(a) === 'orig' ? 0 : 1) - (aspSrc(b) === 'orig' ? 0 : 1) || a.code.localeCompare(b.code));
  const staleTotal = M.skus.filter(s => aspSrc(s) === 'orig').length;
  document.getElementById('asp-edit-count').textContent = `${staleTotal} not updated · ${rows.length} shown`;
  list.innerHTML = rows.map(s => {
    const src = aspSrc(s);
    const cur = +s.asp || 0;
    const prev = (s.asp_prev != null && +s.asp_prev > 0) ? +s.asp_prev : null;   // price before the last update
    const vpct = prev ? (cur - prev) / prev * 100 : null;
    return `<div class="asp-edit-row asp-${src}"><span class="aer-code">${esc(s.code)}</span>`
      + `<span class="aer-name">${esc(s.name || '')}</span>`
      + `<span class="aer-src" title="price source">${ASP_SRC_LABEL[src]}</span>`
      + `<span class="aer-prev" title="price before the last update">${prev != null ? '£' + prev.toFixed(2) : '—'}</span>`
      + `<span class="aer-var ${vpct == null ? '' : vpct >= 0 ? 'c-up' : 'c-down'}" data-prev="${prev != null ? prev : ''}" title="change from the previous price">${vpct == null ? '—' : (vpct >= 0 ? '+' : '') + vpct.toFixed(1) + '%'}</span>`
      + `<input class="aer-input" type="number" step="0.01" min="0" data-code="${esc(s.code)}" data-orig="${cur.toFixed(2)}" value="${cur.toFixed(2)}"></div>`;
  }).join('') || '<div class="muted-note">No products match.</div>';
  list.querySelectorAll('.aer-input').forEach(inp => inp.addEventListener('input', () => { updateManualAspButton(); updateAerVariance(inp); }));
  updateManualAspButton();
}
// Live-update one manual-editor row's variance cell as its price input changes.
function updateAerVariance(inp) {
  const row = inp.closest('.asp-edit-row'); if (!row) return;
  const cell = row.querySelector('.aer-var'); if (!cell) return;
  const prev = parseFloat(cell.dataset.prev), cur = parseFloat(inp.value);
  if (!isFinite(prev) || prev <= 0 || !isFinite(cur)) { cell.textContent = '—'; cell.className = 'aer-var'; return; }
  const v = (cur - prev) / prev * 100;
  cell.textContent = (v >= 0 ? '+' : '') + v.toFixed(1) + '%';
  cell.className = 'aer-var ' + (v >= 0 ? 'c-up' : 'c-down');
}
function collectManualAspChanges() {
  const out = {};
  document.querySelectorAll('#asp-edit-list .aer-input').forEach(inp => {
    const v = parseFloat(inp.value);
    if (isFinite(v) && v > 0 && Math.abs(v - parseFloat(inp.dataset.orig)) > 0.005) out[inp.dataset.code] = +v.toFixed(2);
  });
  return out;
}
function updateManualAspButton() {
  const n = Object.keys(collectManualAspChanges()).length;
  const btn = document.getElementById('btn-apply-manual-asp');
  if (!btn) return;
  btn.disabled = !n;
  btn.textContent = n ? `Apply ${n} manual price${n > 1 ? 's' : ''}` : 'Apply manual prices';
}
async function applyManualAsp() {
  const changes = collectManualAspChanges();
  const codes = Object.keys(changes);
  if (!codes.length) return;
  const years = manualAspYears();
  const status = document.getElementById('save-status');
  status.textContent = 'Saving prices…';
  try {
    const j = await postAsp(changes, years, 'manual');
    if (!j.ok) { status.textContent = ''; alert('Save failed: ' + (j.error || 'unknown')); return; }
    applyAspToMemory(changes, years, 'manual');
    renderAspEditor();   // refresh tags (now 'manual')
    status.textContent = `Set ${codes.length} manual price${codes.length > 1 ? 's' : ''} (${years.join(', ')})`;
  } catch (err) { status.textContent = ''; alert('Save error: ' + err.message); }
}

// Bulk manual-CBM editor (Settings → Data): mirrors the ASP editor. CBM has two
// provenance states — 'import' (from the workbook build) and 'manual' (hand-edited) —
// so you can see which m³/unit figures came straight from the last import.
const CBM_SRC_LABEL = { import: 'import', manual: 'manual' };
function cbmSrc(sku) { return (sku.cbm_src === 'manual' || sku.cbm_src === 'calc') ? 'manual' : 'import'; }
function postCbm(cbmMap, years) {
  return fetch('/api/apply-cbm', { method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ cbm: cbmMap, years }) }).then(r => r.json());
}
// Mirror a CBM change into the in-memory loaded year (so it shows without a reload).
function applyCbmToMemory(cbmMap, years) {
  if (!years.includes(String(YEAR))) return 0;
  let n = 0;
  for (const s of M.skus) { const v = cbmMap[s.code]; if (v != null && v > 0) { s.cbm_prev = (+s.cbm || 0); s.cbm = v; s.cbm_src = 'manual'; n++; } }
  computeAll(); renderSidebar();
  if (currentView === 'plan') renderPlan(); else setView(currentView);
  return n;
}
function renderCbmEditor() {
  const list = document.getElementById('cbm-edit-list'); if (!list || !M) return;
  const q = (document.getElementById('cbm-edit-search').value || '').toLowerCase().trim();
  const manualOnly = document.getElementById('cbm-edit-manual-only').checked;
  const rows = M.skus.filter(s => {
    if (manualOnly && cbmSrc(s) !== 'manual') return false;
    if (q && !((s.code || '').toLowerCase().includes(q) || (s.name || '').toLowerCase().includes(q))) return false;
    return true;
  }).sort((a, b) => a.code.localeCompare(b.code));
  const manualTotal = M.skus.filter(s => cbmSrc(s) === 'manual').length;
  document.getElementById('cbm-edit-count').textContent = `${manualTotal} manually edited · ${rows.length} shown`;
  list.innerHTML = rows.map(s => {
    const src = cbmSrc(s);
    const tint = src === 'manual' ? 'asp-manual' : 'asp-upload';   // reuse the ASP editor tints
    const cur = +s.cbm || 0;
    const prev = (s.cbm_prev != null && +s.cbm_prev > 0) ? +s.cbm_prev : null;
    const vpct = prev ? (cur - prev) / prev * 100 : null;
    return `<div class="asp-edit-row ${tint}"><span class="aer-code">${esc(s.code)}</span>`
      + `<span class="aer-name">${esc(s.name || '')}</span>`
      + `<span class="aer-src" title="CBM source">${CBM_SRC_LABEL[src]}</span>`
      + `<span class="aer-prev" title="CBM before the last manual edit">${prev != null ? prev.toFixed(3) : '—'}</span>`
      + `<span class="aer-var ${vpct == null ? '' : vpct >= 0 ? 'c-up' : 'c-down'}" data-prev="${prev != null ? prev : ''}" title="change from the previous CBM">${vpct == null ? '—' : (vpct >= 0 ? '+' : '') + vpct.toFixed(1) + '%'}</span>`
      + `<input class="aer-input" type="number" step="0.001" min="0" data-code="${esc(s.code)}" data-orig="${cur.toFixed(3)}" value="${cur.toFixed(3)}"></div>`;
  }).join('') || '<div class="muted-note">No products match.</div>';
  list.querySelectorAll('.aer-input').forEach(inp => inp.addEventListener('input', () => { updateManualCbmButton(); updateAerVariance(inp); }));
  updateManualCbmButton();
}
function collectManualCbmChanges() {
  const out = {};
  document.querySelectorAll('#cbm-edit-list .aer-input').forEach(inp => {
    const v = parseFloat(inp.value);
    if (isFinite(v) && v > 0 && Math.abs(v - parseFloat(inp.dataset.orig)) > 0.0005) out[inp.dataset.code] = +v.toFixed(3);
  });
  return out;
}
function updateManualCbmButton() {
  const n = Object.keys(collectManualCbmChanges()).length;
  const btn = document.getElementById('btn-apply-manual-cbm');
  if (!btn) return;
  btn.disabled = !n;
  btn.textContent = n ? `Apply ${n} manual CBM${n > 1 ? 's' : ''}` : 'Apply manual CBM';
}
async function applyManualCbm() {
  const changes = collectManualCbmChanges();
  const codes = Object.keys(changes);
  if (!codes.length) return;
  const years = manualAspYears();
  const status = document.getElementById('save-status');
  status.textContent = 'Saving CBM…';
  try {
    const j = await postCbm(changes, years);
    if (!j.ok) { status.textContent = ''; alert('Save failed: ' + (j.error || 'unknown')); return; }
    applyCbmToMemory(changes, years);
    renderCbmEditor();   // refresh tags (now 'manual')
    status.textContent = `Set ${codes.length} manual CBM${codes.length > 1 ? 's' : ''} (${years.join(', ')})`;
  } catch (err) { status.textContent = ''; alert('Save error: ' + err.message); }
}

/* ---------------- Add a new product (Settings → Data) ----------------
   Builds a full SKU (+ supplier if new) in master.json for the current year and all
   later years, then reloads so the supplier appears in the sidebar and its product in
   the Plan grid. No sales history: ly/actual seed to zero, base_forecast to the annual
   figure spread by the season curve (see newProductProfile); costs/CBM/ASP tag as 'manual'. */
function openAddProductDialog() {
  if (!M) return;
  const dlg = document.getElementById('addprod-dialog');
  const sups = [...M.suppliers].sort((a, b) => a.name.localeCompare(b.name));
  document.getElementById('ap-supplier').innerHTML =
    sups.map(s => `<option value="${esc(s.name)}">${esc(titleCase(s.name))}</option>`).join('')
    + `<option value="__new__">➕ New supplier…</option>`;
  document.getElementById('ap-supplier').value =
    (currentSupplier && sups.some(s => s.name === currentSupplier)) ? currentSupplier : (sups[0] && sups[0].name) || '__new__';
  const cats = [...new Set(M.skus.map(s => s.category).filter(Boolean))].sort();
  document.getElementById('ap-cat-list').innerHTML = cats.map(c => `<option value="${esc(c)}"></option>`).join('');
  ['ap-code', 'ap-name', 'ap-fob', 'ap-landed', 'ap-asp', 'ap-cbm', 'ap-stock', 'ap-forecast', 'ap-fpq',
   'ap-image', 'ap-category', 'ap-sup-name', 'ap-sup-origin', 'ap-sup-port', 'ap-sup-contact', 'ap-sup-email',
   'ap-sup-number'].forEach(id => { const e = document.getElementById(id); if (e) e.value = ''; });
  document.getElementById('ap-season').value = 'Continuity';
  document.getElementById('ap-status').value = 'Live';
  document.getElementById('ap-pallet').value = '';
  document.getElementById('ap-msg').textContent = '';
  apForecastNote();
  apCartonReset();
  apToggleNewSupplier();
  dlg.showModal();
}
function apToggleNewSupplier() {
  document.getElementById('ap-newsup').hidden = document.getElementById('ap-supplier').value !== '__new__';
}
// ---- carton mini-editor embedded in the Add-product form (auto-fills the CBM field) ----
let AP_CARTONS = [], AP_CBM_MANUAL = false;
function apCartonReset() {
  AP_CARTONS = [{ l: '', w: '', h: '', kg: '' }]; AP_CBM_MANUAL = false;
  document.getElementById('ap-pack').value = 1;
  document.getElementById('ap-weightlimit').checked = true;
  apCartonRender();
}
function apCartonRender() {
  const box = document.getElementById('ap-cartons');
  box.innerHTML = AP_CARTONS.map((c, i) => `<div class="carton-row" data-i="${i}">`
    + `<span class="carton-n">#${i + 1}</span>`
    + `<input type="number" class="apc-l" min="0" step="0.1" value="${c.l}" placeholder="L">`
    + `<input type="number" class="apc-w" min="0" step="0.1" value="${c.w}" placeholder="W">`
    + `<input type="number" class="apc-h" min="0" step="0.1" value="${c.h}" placeholder="H">`
    + `<input type="number" class="apc-kg" min="0" step="0.01" value="${c.kg}" placeholder="kg">`
    + `<button type="button" class="carton-del" data-i="${i}"${AP_CARTONS.length > 1 ? '' : ' disabled'}>✕</button></div>`).join('');
  box.querySelectorAll('input').forEach(inp => inp.addEventListener('input', apCartonRead));
  box.querySelectorAll('.carton-del').forEach(b => b.addEventListener('click', () => { AP_CARTONS.splice(+b.dataset.i, 1); apCartonRender(); }));
  apCartonRecalc();
}
function apCartonRead() {
  document.querySelectorAll('#ap-cartons .carton-row').forEach(row => {
    const i = +row.dataset.i;
    AP_CARTONS[i] = { l: parseFloat(row.querySelector('.apc-l').value) || 0, w: parseFloat(row.querySelector('.apc-w').value) || 0,
      h: parseFloat(row.querySelector('.apc-h').value) || 0, kg: parseFloat(row.querySelector('.apc-kg').value) || 0 };
  });
  AP_CBM_MANUAL = false;   // a carton change re-derives CBM
  apCartonRecalc();
}
function apCartonMetrics() { return cartonMetrics(AP_CARTONS, Math.max(1, parseInt(document.getElementById('ap-pack').value, 10) || 1)); }
function apCartonRecalc() {
  const m = apCartonMetrics();
  const out = document.getElementById('ap-carton-calc');
  const weightLimited = document.getElementById('ap-weightlimit').checked;
  if (m.cbm == null) { out.innerHTML = `<span class="muted-note">Optional — add carton sizes to auto-calculate CBM and pallet loading.</span>`; return; }
  if (!AP_CBM_MANUAL) document.getElementById('ap-cbm').value = m.cbm.toFixed(4);   // auto-fill CBM (still editable)
  if (m.needsStillage) {
    out.innerHTML = `<span class="carton-flag">⚠ Too big for a pallet → Stillage. Add it, then set the stillage qty via its CBM chip.</span> · CBM <b>${m.cbm.toFixed(4)}</b>`;
    return;
  }
  const used = cartonLoadQty(m, weightLimited);
  let s = `CBM <b>${m.cbm.toFixed(4)}</b> m³ · volume <b>${m.volumeQty}</b>/pallet`;
  if (m.hasWeight) s += ` (${Math.round(m.volWeight)} kg${m.overweightPct > 0.5 ? `, <span class="carton-flag">${m.overweightPct.toFixed(0)}% over</span>` : ''}) · weight-limit <b>${m.effWeightQty}</b>/pallet`;
  s += ` · using <b>${used}</b>/pallet`;
  out.innerHTML = s;
}
// Weekly shape to spread a new product's ANNUAL forecast over the 53 weeks. A new line
// has no sales history of its own, so we use the same baseline curve the forecast model
// would give it: the calibrated season+category group curve (Summer/Winter land their
// units in the right part of the year), and a flat line for continuity — which really
// is the same units every week. Returns a mean-1 profile, so the annual total is kept.
function newProductProfile(season, category) {
  if (!M || /continu/i.test(season || '')) return FLAT.slice();
  ensureCalib();
  return groupProfile({ season: season || '', category: category || null });
}
function apSpreadForecast(annual, season, category) {
  if (!(annual > 0)) return null;
  const prof = newProductProfile(season, category);
  return prof.map(v => +((annual / WEEKS) * v).toFixed(4));
}
// Live hint under the annual-forecast field: shows how the units will be spread.
function apForecastNote() {
  const el = document.getElementById('ap-fc-note');
  if (!el) return;
  const annual = parseFloat(document.getElementById('ap-forecast').value);
  const season = document.getElementById('ap-season').value;
  const cat = (document.getElementById('ap-category').value || '').trim();
  if (!(annual > 0)) { el.textContent = ''; return; }
  const bf = apSpreadForecast(annual, season, cat);
  if (/continu/i.test(season)) { el.textContent = `≈ ${(annual / WEEKS).toFixed(1)} units every week`; return; }
  const pk = peakWeek(bf);
  el.textContent = `spread by the ${season} curve — peaks wk ${pk} at ~${Math.round(bf[pk - 1])}/wk`;
}
async function submitAddProduct() {
  const val = id => (document.getElementById(id).value || '').trim();
  const numv = id => { const n = parseFloat(val(id).replace(/[^0-9.\-]/g, '')); return isFinite(n) ? n : null; };
  const msg = document.getElementById('ap-msg');
  const code = val('ap-code'), name = val('ap-name');
  if (!code || !name) { msg.textContent = 'Product code and name are required.'; return; }
  let supplier = document.getElementById('ap-supplier').value;
  const newSup = {};
  if (supplier === '__new__') {
    supplier = val('ap-sup-name');
    if (!supplier) { msg.textContent = 'Enter the new supplier name.'; return; }
    newSup.origin = val('ap-sup-origin'); newSup.port = val('ap-sup-port');
    newSup.contact = val('ap-sup-contact'); newSup.email = val('ap-sup-email');
    const sn = parseInt(val('ap-sup-number'), 10); if (isFinite(sn)) newSup.number = sn;
  }
  const years = manualAspYears();
  // carton data (optional): derives CBM + pallet/stillage loading, and the pack size
  const cartons = AP_CARTONS.filter(c => c.l > 0 && c.w > 0 && c.h > 0);
  const pack = Math.max(1, parseInt(val('ap-pack'), 10) || 1);
  const weightLimited = document.getElementById('ap-weightlimit').checked;
  const cm = cartons.length ? cartonMetrics(cartons, pack) : null;
  const cbmField = numv('ap-cbm');
  const payload = {
    code, name, supplier,
    season: val('ap-season') || 'No Defined Season', category: val('ap-category'),
    status: val('ap-status') || 'Live',
    fob: numv('ap-fob'), landed: numv('ap-landed'), asp: numv('ap-asp'),
    cbm: cbmField != null ? cbmField : (cm ? cm.cbm : null),
    stock_now: numv('ap-stock') || 0, annualForecast: numv('ap-forecast') || 0,
    // weekly split of the annual figure, shaped by the product's season (flat for continuity)
    baseForecast: apSpreadForecast(numv('ap-forecast') || 0, val('ap-season'), val('ap-category')),
    fpq: numv('ap-fpq') != null ? numv('ap-fpq') : (cm ? cartonLoadQty(cm, weightLimited) : null),
    palletType: cm && cm.needsStillage ? 'Stillage' : (val('ap-pallet') || (cm ? 'Pallet' : '')),
    image: val('ap-image'),
    cartons, packSize: pack, loadBasis: weightLimited ? 'weight' : 'volume',
    newSupplier: newSup, years,
  };
  msg.textContent = 'Adding…';
  try {
    const r = await fetch('/api/add-product', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    const j = await r.json();
    if (!j.ok) { msg.textContent = j.error || 'Could not add the product.'; return; }
    document.getElementById('addprod-dialog').close();
    document.getElementById('settings-dialog').close();
    currentSupplier = supplier;                       // focus the (possibly new) supplier
    try { localStorage.setItem('tp_supplier', supplier); } catch {}
    await loadYear(YEAR);                             // reload so the SKU + supplier show up
    setView('plan');
    document.getElementById('save-status').textContent = `Added ${code} (${supplier.slice(0, 20)}) to ${years.join(', ')}`;
  } catch (e) { msg.textContent = 'Error: ' + e.message; }
}

/* ---------------- per-supplier Excel form export / import ---------------- */
async function exportSupplier(name) {
  // Always flush first so the server has the latest orders AND the current
  // proposed-rebuy layer on disk (proposed.json) — the proposal layer is built
  // in memory at year-load without marking dirty, so a conditional flush would
  // skip it and the exported form would omit every suggested order.
  await saveNow();
  const a = document.createElement('a');
  a.href = '/api/export-supplier?year=' + encodeURIComponent(YEAR) + '&name=' + encodeURIComponent(name);
  document.body.appendChild(a);
  a.click();
  a.remove();
  document.getElementById('save-status').textContent = 'Exported ' + name.slice(0, 22);
}

// Export several suppliers at once: one form each, bundled into a single .zip so
// the browser does one download instead of one prompt per supplier.
function openExportDialog() {
  const dlg = document.getElementById('export-dialog');
  document.getElementById('ex-year').textContent = YEAR;
  const list = document.getElementById('ex-list');
  list.innerHTML = M.suppliers.map(s =>
    `<label class="ex-item"><input type="checkbox" value="${esc(s.name)}"${s.name === currentSupplier ? ' checked' : ''}><span>${esc(titleCase(s.name))}</span></label>`).join('');
  const go = document.getElementById('ex-go');
  const upd = () => {
    const n = list.querySelectorAll('input:checked').length;
    document.getElementById('ex-count').textContent = `${n} selected`;
    go.textContent = n ? `Export ${n} supplier${n > 1 ? 's' : ''}` : 'Export';
    go.disabled = !n;
  };
  list.oninput = upd;
  document.getElementById('ex-search').oninput = e => {
    const q = e.target.value.toLowerCase();
    list.querySelectorAll('.ex-item').forEach(it => { it.style.display = it.textContent.toLowerCase().includes(q) ? '' : 'none'; });
  };
  document.getElementById('ex-all').onclick = () =>
    { list.querySelectorAll('.ex-item').forEach(it => { if (it.style.display !== 'none') it.querySelector('input').checked = true; }); upd(); };
  document.getElementById('ex-none').onclick = () =>
    { list.querySelectorAll('input').forEach(i => { i.checked = false; }); upd(); };
  document.getElementById('ex-cancel').onclick = () => dlg.close();
  go.onclick = async () => {
    const names = [...list.querySelectorAll('input:checked')].map(i => i.value);
    dlg.close();
    await exportSuppliersZip(names);
  };
  upd();
  dlg.showModal();
}
// True if any of the supplier's SKUs has a planned order (committed + proposed) —
// i.e. the export form would actually contain a week of orders.
function supplierHasOrders(name) {
  for (const sku of M.skus) {
    if (sku.supplier !== name) continue;
    if (combinedOrder(sku.id).some(v => v)) return true;
  }
  return false;
}
function showNotice(title, html) {
  document.getElementById('notice-title').textContent = title;
  document.getElementById('notice-body').innerHTML = html;
  document.getElementById('notice-dialog').showModal();
}
function skippedNoticeHtml(skipped) {
  return `<p>${skipped.length} supplier${skipped.length > 1 ? 's were' : ' was'} skipped — no orders planned for ${YEAR}:</p>
    <ul class="notice-list">${skipped.map(n => `<li>${esc(titleCase(n))}</li>`).join('')}</ul>`;
}
async function exportSuppliersZip(names) {
  if (!names || !names.length) return;
  // skip suppliers with nothing planned, and tell the user which were left out
  const withOrders = [], skipped = [];
  for (const n of names) (supplierHasOrders(n) ? withOrders : skipped).push(n);
  if (!withOrders.length) {
    showNotice('Nothing to export',
      `None of the selected suppliers have any orders planned for ${YEAR}, so no forms were created.` + skippedNoticeHtml(skipped));
    document.getElementById('save-status').textContent = 'Export skipped — no planned orders';
    return;
  }
  const note = () => { if (skipped.length) showNotice('Some suppliers skipped', skippedNoticeHtml(skipped)); };
  if (withOrders.length === 1) { await exportSupplier(withOrders[0]); note(); return; }   // single → plain .xlsx
  const status = document.getElementById('save-status');
  status.textContent = `Exporting ${withOrders.length} forms…`;
  await saveNow();                       // flush orders + proposed so every form is current
  try {
    const r = await fetch('/api/export-suppliers?year=' + encodeURIComponent(YEAR),
      { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ names: withOrders }) });
    if (!r.ok) throw new Error('server ' + r.status);
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `${YEAR} Order Planning forms.zip`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    status.textContent = `Exported ${withOrders.length} forms (.zip)` + (skipped.length ? `, ${skipped.length} skipped` : '');
    note();
  } catch (err) {
    status.textContent = 'Export failed!';
    alert('Export failed: ' + err.message);
  }
}
async function importSupplier(e, name) {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  if (!confirm(`Import "${file.name}"?\n\nThis confirms ${name}'s orders from the form: each product's Order Forecast is set to match the file exactly (committed + any proposed rebuys you exported and edited), and any week 0/blank in the form becomes blank. ${titleCase(name)}'s proposed suggestions are cleared (they are now committed) — no new ones are generated unless you click Run rebuy. Other suppliers are untouched.`))
    return;
  const status = document.getElementById('save-status');
  status.textContent = 'Importing…';
  try {
    const buf = await file.arrayBuffer();
    const r = await fetch('/api/import-supplier?year=' + encodeURIComponent(YEAR) + '&name=' + encodeURIComponent(name), { method: 'POST', body: buf });
    const j = await r.json();
    if (!j.ok) { status.textContent = ''; alert('Import failed: ' + (j.error || 'unknown error')); return; }
    const data = await (await fetch('/api/data?year=' + encodeURIComponent(YEAR))).json();
    ORDERS = data.orders || {};
    for (const sku of M.skus) if (!ORDERS[sku.id]) ORDERS[sku.id] = zeros();
    // the form's quantities are now COMMITTED, so this supplier's old suggestions would
    // double-count — clear just those, and leave every other supplier exactly as-is
    clearProposed(name); computeAll(); markDirty(); renderSidebar(); setView('plan');
    status.textContent = 'Imported ' + new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
    let msg = `Imported from ${file.name}\n\nUpdated ${j.updated} product${j.updated !== 1 ? 's' : ''} across ${j.weeks.length} week${j.weeks.length !== 1 ? 's' : ''} (${(j.total_units || 0).toLocaleString()} units in the form).`;
    if (j.unmatched && j.unmatched.length)
      msg += `\n\n${j.unmatched.length} code(s) not found for this supplier and skipped:\n${j.unmatched.slice(0, 20).join(', ')}`;
    alert(msg);
  } catch (err) {
    status.textContent = '';
    alert('Import error: ' + err.message);
  }
}

/* ---------------- export ---------------- */
function exportCsv() {
  const head = ['SKU', 'Supplier', 'Product'];
  for (let w = 1; w <= WEEKS; w++) head.push('W' + w);
  head.push('Total');
  const lines = [head.join(',')];
  for (const sku of M.skus) {
    const ord = ORDERS[sku.id] || zeros();
    const tot = ord.reduce((a, b) => a + b, 0);
    if (!tot) continue;
    const q = s => '"' + String(s ?? '').replace(/"/g, '""') + '"';
    lines.push([q(sku.code), q(sku.supplier), q(sku.name), ...ord.map(v => v || ''), tot].join(','));
  }
  const blob = new Blob([lines.join('\r\n')], { type: 'text/csv' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `tradeplan-orders-${YEAR}.csv`;
  a.click();
}
// Export the weekly sales-unit forecast (SKU + W1..W53) to Excel. Uses the
// client-computed forecast (so it reflects the multiplier / seasonality / target mode).
async function exportForecast() {
  const rows = M.skus.map(sku => ({ code: sku.code, status: sku.status || '', forecast: (RES.get(sku.id) || {}).forecast || [] }));
  const status = document.getElementById('save-status');
  status.textContent = 'Building forecast workbook…';
  try {
    const r = await fetch('/api/export-forecast?year=' + encodeURIComponent(YEAR),
      { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ rows, week1: M.week1_start }) });
    if (!r.ok) throw new Error('server ' + r.status);
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `${YEAR} Sales Unit Forecast.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    status.textContent = 'Exported sales unit forecast';
  } catch (e) { status.textContent = 'Export failed!'; alert('Forecast export failed: ' + e.message); }
}

/* ---------------- 12-month auto-rebuy container scheduler ---------------- */
const REBUY_TARGET = 4;          // weeks cover to maintain
const REBUY_H = 52;              // rolling horizon (weeks ahead), crossing into next year
let PLAN = null;                 // proposed rebuy plan (separate layer)
let PROPOSED = null;             // editable Map sku.id -> [53] proposed additional orders (this year)
// True when something the suggestions were derived from has since moved (a forecast
// year's opening stock re-chained, the container basket / supplier filter changed).
// Purely advisory: it shows a hint next to "Run rebuy" — it never rebuilds anything.
let REBUY_STALE = false;

let CONFIGS = [];                  // saved named configurations (newest first)
// Refresh the saved-config list + the topbar "Load config" dropdown.
async function refreshConfigList() {
  try { CONFIGS = (await (await fetch('/api/configs')).json()).configs || []; } catch { CONFIGS = []; }
  const sel = document.getElementById('config-select');
  if (sel) sel.innerHTML = '<option value="">Load config…</option>' +
    CONFIGS.map(c => `<option value="${esc(c.file)}">${esc(c.name)}</option>`).join('');
}
// Build scope <option>s for a select: "All years" + each year in `years`.
function scopeOptions(sel, years, includeAll = true) {
  sel.innerHTML = (includeAll ? '<option value="all">All years</option>' : '')
    + years.map(y => `<option value="${y}">${y} only</option>`).join('');
}
// Save a configuration under a name (all years, or just the current year).
function openSaveConfigDialog() {
  const dlg = document.getElementById('saveconfig-dialog');
  const inp = document.getElementById('saveconfig-name');
  const exSel = document.getElementById('saveconfig-existing-sel');
  const ok = document.getElementById('saveconfig-ok');
  inp.value = '';
  const scope = document.getElementById('saveconfig-scope');
  scope.innerHTML = `<option value="all">All years</option><option value="${YEAR}">${YEAR} only</option>`;
  // dropdown of existing configs to overwrite; blank = brand-new config
  exSel.innerHTML = '<option value="">— New configuration —</option>'
    + CONFIGS.map(c => `<option value="${esc(c.name)}">${esc(c.name)}</option>`).join('');
  exSel.value = '';
  const syncOk = () => {                 // reflect whether the typed name overwrites an existing config
    const exists = CONFIGS.some(c => c.name === inp.value.trim());
    ok.textContent = exists ? 'Overwrite' : 'Save';
  };
  exSel.onchange = () => { if (exSel.value) inp.value = exSel.value; syncOk(); inp.focus(); };
  inp.oninput = () => {                   // keep the dropdown in step when the name is typed/edited
    const m = CONFIGS.find(c => c.name === inp.value.trim());
    exSel.value = m ? m.name : '';
    syncOk();
  };
  document.getElementById('saveconfig-existing').innerHTML = CONFIGS.length
    ? `${CONFIGS.length} saved configuration(s) — pick one above to replace it, or type a new name.`
    : 'No saved configurations yet — type a name to create your first.';
  syncOk();
  document.getElementById('saveconfig-cancel').onclick = () => dlg.close();
  ok.onclick = () => doSaveConfig(inp.value.trim(), scope.value);
  inp.onkeydown = e => { if (e.key === 'Enter') doSaveConfig(inp.value.trim(), scope.value); };
  dlg.showModal(); inp.focus();
}
async function doSaveConfig(name, scope) {
  if (!name) { document.getElementById('saveconfig-name').focus(); return; }
  document.getElementById('saveconfig-dialog').close();
  const status = document.getElementById('save-status');
  status.textContent = 'Saving configuration…';
  await saveNow();                 // flush the current year's working state into the json files first
  try {
    const r = await fetch(`/api/save-config?name=${encodeURIComponent(name)}&scope=${encodeURIComponent(scope || 'all')}`, { method: 'POST' });
    const j = await r.json();
    if (!j.ok) throw new Error(j.error || 'failed');
    await refreshConfigList();
    status.textContent = `Saved “${name}” (${scope === 'all' || !scope ? 'all years' : scope})`;
  } catch (e) { status.textContent = 'Save failed!'; alert('Save configuration failed: ' + e.message); }
}
// Restore a named configuration (scope = 'all' or a single year) and reload.
async function loadConfig(file, scope, label) {
  clearTimeout(saveTimer); saveTimer = null;   // discard pending working-state save
  const status = document.getElementById('save-status');
  status.textContent = 'Loading configuration…';
  try {
    const r = await fetch(`/api/load-config?file=${encodeURIComponent(file)}&scope=${encodeURIComponent(scope || 'all')}`, { method: 'POST' });
    const j = await r.json();
    if (!j.ok) throw new Error(j.error || 'failed');
    await loadYear(YEAR);          // reload from the restored orders / proposed / settings
    // the restored committed + proposed layers belong together, so nothing is stale
    if (REBUY_STALE) { REBUY_STALE = false; if (currentView === 'plan') renderPlan(); }
    status.textContent = `Loaded “${label || j.name || file}” (${scope === 'all' || !scope ? 'all years' : scope})`;
  } catch (e) { status.textContent = 'Load failed!'; alert('Load configuration failed: ' + e.message); }
}
// Picking a config from the dropdown → choose scope, then load.
function openLoadConfigDialog(file, defaultScope) {
  const c = CONFIGS.find(x => x.file === file);
  if (!c) return;
  const dlg = document.getElementById('loadconfig-dialog');
  document.getElementById('loadconfig-info').innerHTML =
    `<b>${esc(c.name)}</b> — saved ${new Date(c.saved_at).toLocaleString('en-GB', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' })}. Covers ${(c.years || []).join(', ') || '—'}.`;
  const scope = document.getElementById('loadconfig-scope');
  scopeOptions(scope, c.years || []);
  if (defaultScope && (c.years || []).includes(String(defaultScope))) scope.value = String(defaultScope);
  document.getElementById('loadconfig-cancel').onclick = () => dlg.close();
  document.getElementById('loadconfig-ok').onclick = () => { dlg.close(); loadConfig(c.file, scope.value, c.name); };
  dlg.showModal();
}
async function openRevertDialog() {
  const dlg = document.getElementById('revert-dialog');
  await refreshConfigList();
  scopeOptions(document.getElementById('revert-scope'), YEARS);
  document.getElementById('revert-scope').value = YEAR;   // default to the current year
  const saveBtn = document.getElementById('revert-to-save');
  const when = document.getElementById('revert-save-when');
  if (CONFIGS.length) {
    saveBtn.disabled = false;
    const c = CONFIGS[0];
    when.textContent = `“${c.name}” — saved ${new Date(c.saved_at).toLocaleString('en-GB', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' })}`;
    saveBtn.onclick = () => { dlg.close(); loadConfig(c.file, document.getElementById('revert-scope').value, c.name); };
  } else {
    saveBtn.disabled = true;
    when.textContent = 'No saved configuration yet — use “Save config” first';
    saveBtn.onclick = null;
  }
  document.getElementById('revert-to-original').onclick = () => revertToOriginal(document.getElementById('revert-scope').value);
  document.getElementById('revert-cancel').onclick = () => dlg.close();
  dlg.showModal();
}
async function revertToOriginal(scope) {
  document.getElementById('revert-dialog').close();
  const label = scope === 'all' ? 'ALL years' : `${scope || YEAR}`;
  if (!confirm(`Revert ${label} order quantities back to the originally imported Excel figures?\n\nThis removes committed rebuys, manual edits and proposed suggestions for ${label}. Use "Run rebuy" afterwards to regenerate suggestions.`)) return;
  clearTimeout(saveTimer); saveTimer = null;
  const status = document.getElementById('save-status');
  try {
    const r = await fetch(`/api/revert-original?year=${encodeURIComponent(YEAR)}&scope=${encodeURIComponent(scope || YEAR)}`, { method: 'POST' });
    const j = await r.json();
    if (!j.ok) throw new Error(j.error || 'failed');
    await loadYear(YEAR);
    status.textContent = `Reverted ${label} to original Excel`;
  } catch (e) { status.textContent = 'Revert failed!'; alert('Revert failed: ' + e.message); }
}
function leadWeeks(sup) {        // order placed -> stock arrives
  const s = ((sup && sup.origin) || '') + ' ' + ((sup && sup.port) || '');
  return /none|eur|europe|baltic|domestic|poland|turkey/i.test(s) ? 3 : 11;
}
// Next-year (rollover) annual weekly forecast for a SKU = the de-seasonalised demand
// rate (D, recency-weighted: this year's sell-through + stock-out-corrected last year
// + the original planner forecast as a stabiliser) × the calibrated seasonal profile.
// This is ALWAYS model-driven — a generated forecast year must be rebuilt from sales
// patterns (so fast movers / stock-out lines rise and slow movers decline), not copied
// verbatim from the prior year. The seasonality DISPLAY toggle only controls whether
// the currently-viewed year is re-modelled on screen; it does not turn this off, else
// a rolled-over year would just clone the year it came from.
function nextYearForecast(sku) {
  const mult = SETTINGS.multiplier;
  ensureCalib();
  const d = computeSkuModel(sku);
  const raw = d.prof.map(p => Math.max(0, d.D * p * mult));   // fractional weekly demand
  // Round to whole weekly units WITHOUT losing the annual total: per-week Math.round
  // would zero out a low-volume line (e.g. 12 units/yr spread so thin no single week
  // reaches 0.5 → every week rounds to 0). Largest-remainder keeps the year's total.
  const total = Math.round(raw.reduce((a, b) => a + b, 0));
  const out = raw.map(v => Math.floor(v));
  let rem = total - out.reduce((a, b) => a + b, 0);
  const order = raw.map((v, w) => ({ w, frac: v - Math.floor(v) })).sort((a, b) => b.frac - a.frac);
  for (let i = 0; i < order.length && rem > 0; i++, rem--) out[order[i].w]++;
  return out;
}

// Weeks-of-cover target for a SKU — per-season (Settings), plus an optional extra
// "in-season boost" so seasonal lines carry more buffer through their whole season
// (not just the single peak week). The boost scales with how far above the yearly
// average demand is over the weeks this order will actually STOCK (from arrival =
// decision week + lead), so the extra buffer is built ahead of and across the season.
function coverTargetFor(sku, woy, lead) {
  const rb = SETTINGS.rebuy || {};
  const cs = rb.coverBySeason && rb.coverBySeason[sku.season];
  let T = (cs != null && cs !== '') ? +cs : (rb.coverTarget != null ? +rb.coverTarget : REBUY_TARGET);
  const boost = +rb.peakBoost || 0;
  if (boost > 0 && /summer|winter/i.test(sku.season || '')) {
    const prof = seasonProfile(sku);                 // mean-1 profile: 1 = yearly average
    const span = Math.max(3, Math.round(T));
    let sum = 0;
    for (let k = 0; k < span; k++) { const wk = (((woy + (lead || 0) + k) % WEEKS) + WEEKS) % WEEKS; sum += prof[wk]; }
    const intensity = sum / span;                    // avg demand over the stocked weeks vs average
    const frac = clamp(intensity - 1, 0, 1);         // 0 off-season → full boost at ~2× average (broad in-season window)
    T += boost * frac;
  }
  return Math.max(0, T);
}

// Whether to rebuy a seasonal SKU for an order ARRIVING at time-index arrIdx: only
// when the demand around the arrival is "in season" (≥ cutoff × the line's peak-week
// demand). This stops summer/winter lines being restocked into their off-season (e.g.
// summer stock arriving in winter) — they run down instead — UNLESS the demand there
// is genuinely strong (then it passes). Non-seasonal lines always rebuy.
function seasonalRebuyOk(sku, arrIdx, demArr, cover, cutoff) {
  if (!/summer|winter/i.test(sku.season || '')) return true;
  let peak = 0;
  for (let k = 0; k < demArr.length; k++) if (demArr[k] > peak) peak = demArr[k];
  if (peak <= 0) return true;
  const span = Math.max(1, Math.round(cover));
  let sum = 0, n = 0;
  for (let k = arrIdx; k < arrIdx + span && k < demArr.length; k++) { sum += demArr[k] || 0; n++; }
  return (n ? sum / n : 0) >= peak * cutoff;
}

// Forward auto-rebuy scheduler over a rolling 12 months. Existing app orders are
// treated as committed incoming supply; the plan only proposes ADDITIONAL units
// where a deficit remains, consolidated into container shipments timed by lead.
// On a forecast year (cur=1) the simulation starts L weeks "before" week 1 so
// stock can arrive in weeks 1..L (those orders are placed in the prior year's tail).
// Per-supplier one-product-per-container mode ships each SKU alone in full-container lots.
function buildRebuyPlan() {
  const cur = SETTINGS.current_week, CC = SETTINGS.container_cbm || 68;
  const rb = SETTINGS.rebuy || {};
  const mode = rb.mode || 'full', maxWait = rb.maxWait || 4;
  // how full a shipment aims to be before it goes: a full container, or — in "Allow
  // partial" mode — a half-container target (default 28 CBM) so partial shipments build
  // up to a worthwhile load instead of going out as tiny amounts.
  const fillTarget = mode === 'partial' ? Math.min(CC, Math.max(1, +rb.partialCbm || 28)) : CC;
  const cutoff = rb.seasonalCutoff != null ? +rb.seasonalCutoff : 0.40;   // off-season rebuy threshold
  const supSel = (rb.suppliers && rb.suppliers.length) ? new Set(rb.suppliers) : null;  // null = all
  const oneProd = new Set(rb.oneProductSuppliers || []);
  if ((+rb.peakBoost || 0) > 0) ensureCalib();      // peak boost needs calibrated profiles
  const maxCover = Math.max(+rb.coverTarget || REBUY_TARGET, REBUY_TARGET,
    ...Object.values(rb.coverBySeason || {}).map(v => +v || 0)) + (+rb.peakBoost || 0);
  const Lmax = 11, H2 = REBUY_H + Lmax + Math.ceil(maxCover) + 4;
  const proposed = new Map(), containers = [];
  const bySup = new Map();
  for (const sku of M.skus) {
    if (sku.status !== 'Live' || !RES.get(sku.id)) continue;
    if (supSel && !supSel.has(sku.supplier)) continue;
    if (!bySup.has(sku.supplier)) bySup.set(sku.supplier, []);
    bySup.get(sku.supplier).push(sku);
  }
  const woyOf = t => (cur - 1) + t;                  // absolute week offset from this year's week 1
  for (const [supName, skus] of bySup) {
    const L = leadWeeks(supByName.get(supName));
    const single = oneProd.has(supName);
    const start = cur === 1 ? -L : 0;                // forecast year: allow arrivals in weeks 1..L
    const proj = {}, propArr = {}, dem = {};
    for (const sku of skus) {
      proj[sku.id] = sku.stock_now || 0;
      propArr[sku.id] = new Array(H2).fill(0);
      const ty = RES.get(sku.id).forecast, ny = nextYearForecast(sku), d = new Array(H2);
      for (let t = 0; t < H2; t++) { const woy = woyOf(t); d[t] = woy < WEEKS ? (ty[woy] || 0) : (woy < 2 * WEEKS ? (ny[woy - WEEKS] || 0) : 0); }
      dem[sku.id] = d;
    }
    const demAt = (id, k) => (k >= 0 && k < H2) ? dem[id][k] : 0;
    const propAt = (id, k) => (k >= 0 && k < H2) ? propArr[id][k] : 0;
    const committed = (id, t) => { const woy = woyOf(t); return (woy >= 0 && woy < WEEKS) ? (ORDERS[id] ? ORDERS[id][woy] || 0 : 0) : 0; };
    const ship = (t, lines, cbm) => {
      const arr = t + L; if (arr < 0 || arr >= REBUY_H) return;
      let fob = 0; const out = [];
      for (const ln of lines) {
        if (ln.qty <= 0) continue;
        propArr[ln.id][arr] += ln.qty;
        const woy = woyOf(arr);
        if (woy >= 0 && woy < WEEKS) { if (!proposed.has(ln.id)) proposed.set(ln.id, zeros()); proposed.get(ln.id)[woy] += ln.qty; }
        const sk = skuById.get(ln.id); fob += ln.qty * (sk.fob || 0); out.push({ code: sk.code, qty: ln.qty });
      }
      if (!out.length) return;
      const aw = woyOf(arr), ow = woyOf(t);
      const yrWk = (woy, base) => woy < 0 ? { y: base - 1, w: woy + WEEKS + 1 } : (woy < WEEKS ? { y: base, w: woy + 1 } : { y: base + 1, w: woy - WEEKS + 1 });
      const a = yrWk(aw, +YEAR), o = yrWk(ow, +YEAR);
      containers.push({ supplier: supName, lead: L, cbm: +cbm.toFixed(1), nContainers: +(cbm / CC).toFixed(2), fob, lines: out,
        arrYear: a.y, arrWeek: a.w, ordYear: o.y, ordWeek: o.w, arrAbs: aw, single });
    };

    if (single) {
      // one product per container: order each SKU alone in whole-container lots
      for (const sku of skus) {
        if (!(sku.cbm > 0)) continue;
        const perC = Math.max(1, Math.round(CC / sku.cbm));   // units that fill ~one container
        let p = sku.stock_now || 0;
        for (let t = start; t < REBUY_H; t++) {
          if (t >= 0) p = Math.max(0, p + committed(sku.id, t) + propAt(sku.id, t) - demAt(sku.id, t));
          const T = coverTargetFor(sku, woyOf(t + 1), L);
          const end = t + 1 + L + T;
          let D = 0, incoming = 0;
          for (let k = t + 1; k < end; k++) { D += demAt(sku.id, k); incoming += committed(sku.id, k) + propAt(sku.id, k); }
          const deficit = D - p - incoming;
          if (deficit > 0 && seasonalRebuyOk(sku, t + L, dem[sku.id], T, cutoff)) {
            const nC = Math.max(1, Math.ceil(deficit / perC)), qty = nC * perC;
            ship(t, [{ id: sku.id, qty }], qty * sku.cbm);
          }
        }
        proj[sku.id] = p;
      }
      continue;
    }

    // mixed basket (default): consolidate the supplier's SKUs into ~full containers
    // Will any line already in the basket run out within the lead window? If so we must
    // ship now so the container ARRIVES (t+L) by the stockout instead of after it.
    const stockoutWithinLead = (t) => {
      for (const sku of skus) {
        if (!basket[sku.id]) continue;            // only lines we're actively rebuying
        let p = proj[sku.id] || 0;
        for (let k = t + 1; k <= t + L && k < H2; k++) {
          p += committed(sku.id, k) + propAt(sku.id, k) - demAt(sku.id, k);
          if (p <= 0 && demAt(sku.id, k) > 0) return true;
        }
      }
      return false;
    };
    // FULL mode only: when we're forced to ship before the basket naturally fills (to
    // beat an imminent stockout), top the container up toward CC by pulling the
    // supplier's SOONEST upcoming demand forward — so a container we're paying for
    // anyway goes out close to full instead of half-empty. Only adds units genuinely
    // needed later (walks each line's projected stock forward from the arrival and adds
    // real shortfalls, soonest first), capped at the container, and respects the
    // seasonal off-season guard. Never invents demand, so thin suppliers still ship light.
    const topUpToContainer = (t, target) => {
      if (basketCbm >= target) return;
      const arr = t + L; if (arr < 0 || arr >= REBUY_H) return;
      const rp = {}, okLine = {};
      for (const sku of skus) {
        let p = proj[sku.id] || 0;
        for (let k = t + 1; k <= arr && k < H2; k++) p += committed(sku.id, k) + propAt(sku.id, k) - demAt(sku.id, k);
        rp[sku.id] = p + (basket[sku.id] || 0);   // stock right after this shipment lands
        okLine[sku.id] = (sku.cbm > 0) &&
          seasonalRebuyOk(sku, arr, dem[sku.id], coverTargetFor(sku, woyOf(t + 1), L), cutoff);
      }
      for (let k = arr + 1; k < H2 && basketCbm < target - 1e-6; k++) {
        for (const sku of skus) {
          if (!okLine[sku.id]) continue;
          rp[sku.id] += committed(sku.id, k) + propAt(sku.id, k) - demAt(sku.id, k);
          if (rp[sku.id] < 0) {
            const space = Math.floor((target - basketCbm) / sku.cbm);
            if (space <= 0) continue;             // no whole unit of this line fits
            const addU = Math.min(Math.ceil(-rp[sku.id]), space);
            basket[sku.id] = (basket[sku.id] || 0) + addU;
            basketCbm += addU * sku.cbm;
            rp[sku.id] += addU;
          }
        }
      }
    };
    let basketCbm = 0, basketStart = -1; const basket = {};
    for (let t = start; t < REBUY_H; t++) {
      for (const sku of skus) if (t >= 0) proj[sku.id] = Math.max(0, proj[sku.id] + committed(sku.id, t) + propAt(sku.id, t) - demAt(sku.id, t));
      for (const sku of skus) {
        if (!(sku.cbm > 0)) continue;
        const T = coverTargetFor(sku, woyOf(t + 1), L);
        if (!seasonalRebuyOk(sku, t + L, dem[sku.id], T, cutoff)) continue;   // off-season seasonal line → let it run down
        const end = t + 1 + L + T;
        let D = 0, incoming = 0;
        for (let k = t + 1; k < end; k++) { D += demAt(sku.id, k); incoming += committed(sku.id, k) + propAt(sku.id, k); }
        const already = basket[sku.id] || 0;
        const deficit = Math.round(D - (proj[sku.id] || 0) - incoming - already);
        if (deficit > 0) { if (basketStart < 0) basketStart = t; basket[sku.id] = already + deficit; basketCbm += deficit * sku.cbm; }
      }
      const waited = mode === 'partial' && basketStart >= 0 && (t - basketStart) >= maxWait && basketCbm > 0;
      // ship when the fill target is reached (full container / partial half-load), after
      // the partial max-wait, or to beat an imminent stock-out (avoids gaps)
      if (basketCbm >= fillTarget || waited || (basketCbm > 0 && stockoutWithinLead(t))) {
        topUpToContainer(t, fillTarget);   // forced/early ship → build up toward the target
        ship(t, Object.keys(basket).map(id => ({ id, qty: basket[id] })), basketCbm);
        basketCbm = 0; basketStart = -1; for (const k in basket) delete basket[k];
      }
    }
    // flush any leftover basket so demand isn't dropped (topped up toward the target)
    if (basketCbm > 0) { topUpToContainer(REBUY_H - 1 - L, fillTarget); ship(REBUY_H - 1 - L, Object.keys(basket).map(id => ({ id, qty: basket[id] })), basketCbm); }
  }
  let totFob = 0, totCbm = 0, totUnits = 0, totC = 0;
  for (const c of containers) { totFob += c.fob; totCbm += c.cbm; totC += c.nContainers; for (const l of c.lines) totUnits += l.qty; }
  PLAN = { proposed, containers, totFob, totCbm, totUnits, totContainers: totC,
           skuCount: proposed.size, supCount: new Set(containers.map(c => c.supplier)).size, cur, mode };
  return PLAN;
}

function supplierFilterSummary() {
  const sel = SETTINGS.rebuy.suppliers;
  if (!sel || !sel.length) return 'All suppliers';
  return `${sel.length} of ${M.suppliers.length} suppliers`;
}
// (re)build the editable proposal layer by running the algorithm. supName: rebuild only
// that supplier's suggestions, leaving every other supplier's proposals untouched;
// null/omitted = rebuild the whole year. (buildRebuyPlan derives proposals purely from
// committed ORDERS per supplier, so one supplier's slice of a full rebuild is the same
// as a supplier-only rebuild.)
function resetProposed(supName) {
  buildRebuyPlan();
  if (!supName) {
    PROPOSED = new Map();
    for (const [id, arr] of PLAN.proposed) PROPOSED.set(id, arr.slice());
    return;
  }
  if (!PROPOSED) PROPOSED = new Map();
  for (const sku of M.skus) if (sku.supplier === supName) PROPOSED.delete(sku.id);   // drop old suggestions for this supplier
  for (const [id, arr] of PLAN.proposed) {
    const s = skuById.get(id);
    if (s && s.supplier === supName) PROPOSED.set(id, arr.slice());
  }
}
// empty the proposal layer (teal Proposed Rebuy row → blank). supName: clear only that
// supplier's suggestions; null/omitted = clear the whole year.
function clearProposed(supName) {
  if (!PROPOSED || !supName) { PROPOSED = new Map(); return; }
  for (const sku of M.skus) if (sku.supplier === supName) PROPOSED.delete(sku.id);
}
function restoreProposed(obj) {       // rebuild the proposal layer from a saved {id:[53]} map
  PROPOSED = new Map();
  for (const id in obj) if (skuById.has(id) && Array.isArray(obj[id])) PROPOSED.set(id, obj[id].slice());
}
function serializeProposed() {        // non-empty proposed arrays, for server export merge
  const o = {};
  if (PROPOSED) for (const [id, arr] of PROPOSED) if (arr.some(v => v)) o[id] = arr;
  return o;
}
function proposedGrandTotals() {
  let units = 0, cbm = 0, fob = 0; const sups = new Set();
  for (const [id, arr] of PROPOSED) {
    const s = skuById.get(id); if (!s) continue;
    const q = arr.reduce((a, b) => a + b, 0);
    units += q; cbm += q * (s.cbm || 0); fob += q * (s.fob || 0); if (q > 0) sups.add(s.supplier);
  }
  return { units, cbm, fob, skuCount: PROPOSED.size, supCount: sups.size };
}
function proposedSupTotals(name) {    // proposed units/cbm/fob for one supplier
  let units = 0, cbm = 0, fob = 0;
  for (const sku of M.skus) {
    if (sku.supplier !== name) continue;
    const arr = PROPOSED.get(sku.id); if (!arr) continue;
    const q = arr.reduce((a, b) => a + b, 0);
    units += q; cbm += q * (sku.cbm || 0); fob += q * (sku.fob || 0);
  }
  return { units, cbm, fob };
}
function rebuySuppliersWithProposals() {
  const set = new Set();
  for (const id of PROPOSED.keys()) { const s = skuById.get(id); if (s) set.add(s.supplier); }
  return M.suppliers.filter(s => set.has(s.name)).map(s => s.name);
}

function renderRebuy() {
  if (!PROPOSED) PROPOSED = new Map();   // never auto-build: only "Run rebuy" may do that
  if (!PLAN) buildRebuyPlan();           // container roll-over stats only; leaves PROPOSED alone
  const cur = SETTINGS.current_week, CC = SETTINGS.container_cbm || 68;
  const mode = SETTINGS.rebuy.mode || 'full';
  const supChecks = M.suppliers.map(s => {
    const on = !SETTINGS.rebuy.suppliers || SETTINGS.rebuy.suppliers.includes(s.name);
    return `<label class="rb-sup-item"><input type="checkbox" class="rb-sup-cb" value="${esc(s.name)}" ${on ? 'checked' : ''}> ${esc(titleCase(s.name))}</label>`;
  }).join('');
  const gt = proposedGrandTotals();
  const rollover = PLAN.containers.filter(c => c.arrYear > +YEAR).length;
  const cards = [
    [(gt.cbm / CC).toFixed(1), `Containers to commit (${YEAR})`],
    [fmtGBPk(gt.fob), 'FOB to commit'],
    [Math.round(gt.units).toLocaleString(), 'Units to add'],
    [gt.skuCount.toLocaleString(), 'Products'],
    [gt.supCount.toLocaleString(), 'Suppliers'],
    [rollover.toLocaleString(), `Containers rolling to ${+YEAR + 1}`],
  ].map(([v, l]) => `<div class="card"><div class="v">${v}</div><div class="l">${l}</div></div>`).join('');
  const rolloverNote = rollover ? ` <b>${rollover}</b> further container${rollover !== 1 ? 's' : ''} of need roll into <b>${+YEAR + 1}</b> (beyond this year's weeks) — switch to the ${+YEAR + 1} year to plan those.` : '';
  document.getElementById('main').innerHTML = `
    <div class="cards">${cards}</div>
    <div id="rebuy-controls">
      <div class="rb-ctl"><span class="rb-ctl-lbl">Containers:</span>
        <button class="rb-mode${mode === 'full' ? ' on' : ''}" data-mode="full">Full ${CC} CBM only</button>
        <button class="rb-mode${mode === 'partial' ? ' on' : ''}" data-mode="partial">Allow partial</button></div>
      <details id="rb-sup"><summary>Predict for: <b>${supplierFilterSummary()}</b></summary>
        <div class="rb-sup-actions"><button type="button" id="rb-sup-all">Select all</button><button type="button" id="rb-sup-none">Select none</button></div>
        <div class="rb-sup-list">${supChecks}</div>
      </details>
      <span class="spacer"></span>
      <button id="btn-commit-all" class="primary">Commit all suppliers</button>
      <button id="btn-rebuy-reset" title="Rebuild suggestions, discarding edits">Reset suggestions</button>
    </div>
    <div id="rebuy-note">Editable proposed rebuys (<b>additional</b> orders on top of committed) to hold ~${REBUY_TARGET} weeks cover ·
      <b>${mode === 'full' ? `full ${CC} CBM only` : 'partial allowed'}</b> · lead China ~11 / non‑China ~3 wk · ${+YEAR + 1} uses the rollover forecast.
      Edit the blue <b>Proposed Rebuy</b> row, then commit per supplier or all.${rolloverNote}</div>
    <div id="rebuy-grid"></div>`;
  document.querySelectorAll('.rb-mode').forEach(b => b.addEventListener('click', () => {
    SETTINGS.rebuy.mode = b.dataset.mode; markDirty(); resetProposed(); renderRebuy();
  }));
  const readSupChecks = () => {
    const boxes = [...document.querySelectorAll('.rb-sup-cb')];
    const checked = boxes.filter(b => b.checked).map(b => b.value);
    SETTINGS.rebuy.suppliers = checked.length === boxes.length ? null : checked;
    markDirty(); resetProposed(); renderRebuy();
  };
  document.querySelectorAll('.rb-sup-cb').forEach(cb => cb.addEventListener('change', readSupChecks));
  document.getElementById('rb-sup-all').addEventListener('click', () => { document.querySelectorAll('.rb-sup-cb').forEach(c => c.checked = true); readSupChecks(); });
  document.getElementById('rb-sup-none').addEventListener('click', () => { document.querySelectorAll('.rb-sup-cb').forEach(c => c.checked = false); readSupChecks(); });
  document.getElementById('btn-commit-all').addEventListener('click', () => commitRebuy(null));
  document.getElementById('btn-rebuy-reset').addEventListener('click', () => { resetProposed(); renderRebuy(); });
  // pick a supplier that has proposals
  const withProp = rebuySuppliersWithProposals();
  if (!withProp.includes(currentSupplier)) currentSupplier = withProp[0] || (M.suppliers[0] && M.suppliers[0].name);
  renderSidebar();
  renderRebuyGrid();
}

const REBUY_ROWS = [
  { key: 'forecast', label: 'Sales Forecast', fmt: fmtU },
  { key: 'committed', label: 'Committed Order', fmt: fmtU },
  { key: 'proposed', label: 'Proposed Rebuy', edit: true },
  { key: 'stock', label: 'Projected Stock', fmt: fmtU },
  { key: 'cover', label: 'Weeks Cover', fmt: fmt1 },
];
function combinedOrder(id) {
  const c = ORDERS[id] || EMPTY53, p = (PROPOSED && PROPOSED.get(id)) || EMPTY53;
  return c.map((v, w) => (v || 0) + (p[w] || 0));
}
function renderRebuyGrid() {
  const wrap = document.getElementById('rebuy-grid'); if (!wrap) return;
  const sup = supByName.get(currentSupplier);
  if (!sup) { wrap.innerHTML = '<div class="empty">No supplier with proposed rebuys.</div>'; return; }
  const cur = SETTINGS.current_week;
  const skus = M.skus.filter(k => k.supplier === sup.name && PROPOSED.has(k.id));
  const t = proposedSupTotals(sup.name);
  let body = '';
  skus.forEach((sku, idx) => {
    const r = computeSku(sku, combinedOrder(sku.id));
    const inf = `${sku.season} · FOB £${sku.fob.toFixed(2)} · ${sku.cbm ? sku.cbm.toFixed(3) + ' cbm' : ''} · stock now ${Math.round(sku.stock_now).toLocaleString()}`;
    const img = sku.image ? `<img src="${esc(sku.image)}" loading="lazy" onerror="this.remove()">` : '';
    body += `<tr class="skuhead"><td colspan="${WEEKS + 2}"><div class="skuhead-inner">${img}<span class="code">${esc(sku.code)}</span> <span class="nm">${esc(sku.name || '')}</span> ${statusBadge(sku.status)}<span class="inf">${inf}</span></div></td></tr>`;
    for (const def of REBUY_ROWS) {
      body += `<tr><td class="lbl">${def.label}</td>`;
      let tot = 0;
      const arr = def.key === 'committed' ? (ORDERS[sku.id] || zeros())
        : def.key === 'proposed' ? (PROPOSED.get(sku.id) || zeros()) : r[def.key];
      for (let w = 0; w < WEEKS; w++) {
        const v = arr[w]; tot += v; const wk = w + 1;
        if (def.edit) {
          body += `<td class="ocell rb-prop${wk === cur ? ' curwk' : ''}"><input data-sku="${esc(sku.id)}" data-r="${idx}" data-w="${w}" value="${v || ''}" inputmode="numeric"></td>`;
        } else {
          let cls = wk === cur ? 'curwk ' : '', style = '';
          if (def.key === 'cover') { const cc = coverCellStyle(arr, w); style = ` style="background:${cc.bg};color:${cc.color}"`; }
          if (def.key === 'stock' && wk >= cur && v === 0 && r.forecast[w] > 0) cls += 'stockout ';
          if (Math.abs(v) < .5 && def.key !== 'cover') cls += 'zero';
          body += `<td class="${cls.trim()}"${style} data-sku="${esc(sku.id)}" data-k="${def.key}" data-w="${w}">${def.fmt(v)}</td>`;
        }
      }
      if (def.key === 'forecast') {
        const outturn = r.forecastDisp.reduce((a, b) => a + b, 0);
        body += `<td class="tot tot-fc" data-sku="${esc(sku.id)}" data-k="forecast" data-w="T" title="Outturn = actual sales so far + forecast for the rest of the year">${fmtU(outturn)}<span class="tot-sub" title="Full-year sales forecast (plan)">plan ${fmtU(tot)}</span></td></tr>`;
      } else {
        body += `<td class="tot" data-sku="${esc(sku.id)}" data-k="${def.key}" data-w="T">${(def.edit ? fmtU : def.fmt)(tot)}</td></tr>`;
      }
    }
  });
  if (!skus.length) body = `<tr class="skuhead"><td colspan="${WEEKS + 2}"><div class="skuhead-inner" style="font-weight:400;color:var(--dim)">No proposed rebuys for this supplier (committed orders already cover the target).</div></td></tr>`;
  wrap.innerHTML = `
    <div class="sup-head">
      <div class="sup-head-main"><h1>${esc(titleCase(sup.name))}</h1>
        <div class="meta">Proposed: <b>${fmtU(t.units)}</b> units · <b>${fmt1(t.cbm)}</b> cbm (≈${(t.cbm / (SETTINGS.container_cbm || 68)).toFixed(1)} containers) · <b>${fmtGBP(t.fob)}</b> FOB</div></div>
      <div class="sup-actions">
        <button id="btn-commit-sup" class="primary"${t.units ? '' : ' disabled'}>Commit ${esc(titleCase(sup.name)).slice(0, 22)}</button>
      </div>
    </div>
    <div class="gridwrap"><table class="grid"><thead>${headerRow()}</thead><tbody>${body}</tbody></table></div>`;
  document.getElementById('btn-commit-sup').addEventListener('click', () => commitRebuy(sup.name));
  wrap.querySelectorAll('td.rb-prop input').forEach(inp => {
    inp.addEventListener('change', () => commitProposedEdit(inp));
    inp.addEventListener('focus', () => inp.select());
    inp.addEventListener('keydown', e => orderKeyNav(e, inp));
  });
}
function commitProposedEdit(inp) {
  const id = inp.dataset.sku, w = +inp.dataset.w;
  let v = Math.round(parseFloat(String(inp.value).replace(/[^0-9.\-]/g, '')));
  if (!isFinite(v) || v < 0) v = 0;
  inp.value = v || '';
  if (!PROPOSED.has(id)) PROPOSED.set(id, zeros());
  PROPOSED.get(id)[w] = v;
  // recompute this SKU's projected stock + cover cells and supplier totals
  const sku = skuById.get(id), r = computeSku(sku, combinedOrder(id)), cur = SETTINGS.current_week;
  document.querySelectorAll(`td[data-sku="${CSS.escape(id)}"]`).forEach(td => {
    const k = td.dataset.k; if (k !== 'stock' && k !== 'cover') return;
    if (td.dataset.w === 'T') { td.textContent = (k === 'cover' ? fmt1 : fmtU)(r[k].reduce((a, b) => a + b, 0)); return; }
    const wk = +td.dataset.w, val = r[k][wk]; td.textContent = (k === 'cover' ? fmt1 : fmtU)(val);
    td.className = ''; td.style.background = ''; td.style.color = '';
    if (wk + 1 === cur) td.classList.add('curwk');
    if (k === 'cover') { const cc = coverCellStyle(r[k], wk); td.style.background = cc.bg; td.style.color = cc.color; }
    else if (k === 'stock' && wk + 1 >= cur && val === 0 && r.forecast[wk] > 0) td.classList.add('stockout');
    else if (Math.abs(val) < .5) td.classList.add('zero');
  });
  const t = proposedSupTotals(sku.supplier);
  const meta = document.querySelector('#rebuy-grid .sup-head .meta');
  if (meta) meta.innerHTML = `Proposed: <b>${fmtU(t.units)}</b> units · <b>${fmt1(t.cbm)}</b> cbm (≈${(t.cbm / (SETTINGS.container_cbm || 68)).toFixed(1)} containers) · <b>${fmtGBP(t.fob)}</b> FOB`;
}
function commitRebuy(supName) {
  const cur = SETTINGS.current_week;
  const ids = [...PROPOSED.keys()].filter(id => !supName || (skuById.get(id) && skuById.get(id).supplier === supName));
  const units = ids.reduce((a, id) => a + PROPOSED.get(id).reduce((x, y) => x + y, 0), 0);
  if (!units) { alert('Nothing to commit.'); return; }
  const who = supName ? titleCase(supName) : 'all suppliers';
  if (!confirm(`Commit proposed rebuys for ${who}?\n\nAdds ${Math.round(units).toLocaleString()} units across ${ids.length} product(s) on top of existing ${YEAR} orders. They become editable (blue) order quantities in the Plan view and flow into supplier exports.`)) return;
  for (const id of ids) {
    if (!ORDERS[id]) ORDERS[id] = zeros();
    const p = PROPOSED.get(id);
    for (let w = cur - 1; w < WEEKS; w++) ORDERS[id][w] += p[w];
    PROPOSED.delete(id);
  }
  computeAll(); markDirty();
  renderSidebar(); renderPlan();
  document.getElementById('save-status').textContent = `Committed rebuys: ${who}`;
}

/* ---------------- generate a forecast-only year (e.g. 2027) ---------------- */
// Builds a master for `newYear` from the currently-loaded year: forecasted sales
// = the rollover forecast (model baseline × seasonal shape), starting stock = this
// year's projected closing stock, so projected stock then depletes by forecast.
function buildForecastYear(newYear) {
  const bnd = (+M.data_week || SETTINGS.current_week || WEEKS) - 1;   // actuals run weeks 1..bnd
  const skus = M.skus.map(sku => {
    const r = RES.get(sku.id);
    const ny = nextYearForecast(sku);
    const ly = sku.actual.map((a, w) => (w < bnd ? a : (r ? r.forecast[w] : 0)));   // this year's realised+forecast sales
    const endStock = r ? Math.max(0, Math.round(r.stock[WEEKS - 1])) : (sku.stock_now || 0);
    return {
      id: sku.id, code: sku.code, supplier: sku.supplier, name: sku.name,
      category: sku.category, status: sku.status, season: sku.season,
      fob: sku.fob, landed: sku.landed, asp: sku.asp, cbm: sku.cbm,
      fpq: sku.fpq, pallet_type: sku.pallet_type, image: sku.image,
      stock_now: endStock, ly, actual: zeros(), base_forecast: ny.slice(), running_stock: zeros(),
    };
  });
  const orders = {}; for (const s of skus) orders[s.id] = zeros();
  const w1 = new Date(M.week1_start); w1.setDate(w1.getDate() + 364);
  const master = {
    generated: new Date().toISOString().slice(0, 19), source: `forecast generated from ${YEAR}`, year: +newYear,
    weeks: WEEKS, week1_start: w1.toISOString().slice(0, 10), data_week: 1,
    suppliers: M.suppliers.map(s => Object.assign({}, s)),
    skus, history: M.history, capacities: M.capacities, container_cbm: M.container_cbm || 68, multiplier: 1,
  };
  return { master, orders };
}

async function buildForecastYearAction() {
  const baseYear = +YEAR;                 // build/rebuild the year AFTER the one being viewed
  const newYear = baseYear + 1;
  const exists = YEARS.includes(String(newYear));
  // Never overwrite an imported (real) year — only build new years or rebuild years
  // that were themselves generated as forecasts. (Rebuilding lets a corrected base
  // year flow forward into an existing forecast year.)
  if (exists) {
    try {
      const mj = await (await fetch('/api/data?year=' + newYear)).json();
      const src = (mj.master && mj.master.source) || '';
      if (!/^forecast generated/i.test(src)) {
        alert(`${newYear} holds imported actuals, not a generated forecast — it won't be rebuilt. Switch to your latest real year to build the next forecast year from it.`);
        return;
      }
    } catch (e) { /* couldn't verify — fall through to the explicit confirm below */ }
  }
  if (!confirm(`${exists ? 'Rebuild' : 'Build'} a ${newYear} forecast year from ${baseYear}?\n\nIt pulls in the forecasted weekly sales (the ${SEASON().enabled ? 'seasonality model’s' : 'current'} rollover forecast) and projected stock for every product, starting from ${baseYear}’s projected year-end stock. ${exists ? `This REBUILDS the existing ${newYear} year — its proposed rebuys and any committed ${newYear} orders are regenerated from scratch.` : ''} You can then browse it per supplier and run its own rebuy plan.`)) return;
  const status = document.getElementById('save-status');
  status.textContent = `Building ${newYear}…`;
  try {
    if (String(YEAR) !== String(baseYear)) await loadYear(String(baseYear));
    buildModeledForecasts(); computeAll();
    const payload = buildForecastYear(newYear);
    const r = await fetch('/api/save-year?year=' + newYear, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    const j = await r.json();
    if (!j.ok) { status.textContent = ''; alert('Build failed: ' + (j.error || 'unknown')); return; }
    YEARS = j.years;
    const sel = document.getElementById('year-select');
    sel.innerHTML = YEARS.map(y => `<option value="${y}">${y}</option>`).join('');
    refreshBuildYearBtn();
    await loadYear(String(newYear));
    status.textContent = `${newYear} forecast built`;
  } catch (e) { status.textContent = ''; alert('Build error: ' + e.message); }
}
function refreshBuildYearBtn() {
  const btn = document.getElementById('btn-build-year');
  if (!btn) return;
  const next = (+YEAR || Math.max(...YEARS.map(Number))) + 1;   // year after the one being viewed
  const exists = YEARS.includes(String(next));
  btn.textContent = (exists ? '↻ ' : '+ ') + next;
  btn.title = `${exists ? 'Rebuild' : 'Build'} a ${next} forecast year (forecasted sales + projected stock) from ${next - 1}`;
}

/* ---------------- view switching / init ---------------- */
/* ================= Sales channels (Channel Sales import) =================
   The 'Channel Sales' export gives, per SKU, each customer's units (TY + LY) and
   selling price. The importer turns units into a historical unit-share weight per
   customer; customers roll up to four channels via a user-maintained mapping in
   Settings (SETTINGS.channel_map) — Marketplace / DSV / Direct / Ex-Works;
   anything unassigned reports as Unmapped. Split units = year forecast × share;
   split value = units × customer price (falling back to the SKU's ASP). */
const CHANNELS = ['Marketplace', 'DSV', 'Direct', 'Ex-Works'];
function chanReady() { return !!(CHANNEL_INDEX && CHANNEL_INDEX.skus); }
function chanMap() { if (!SETTINGS.channel_map) SETTINGS.channel_map = {}; return SETTINGS.channel_map; }
function custChannel(c) { return chanMap()[c] || 'Unmapped'; }
function custName(c) { return (CHANNEL_INDEX && CHANNEL_INDEX.customers && CHANNEL_INDEX.customers[c]) || ''; }
// normalised customer rows for one SKU: [{c, name, ratio, price}] (ratios sum to 1)
function skuChannelRows(sku) {
  const rows = chanReady() ? CHANNEL_INDEX.skus[sku.code] : null;
  if (!rows || !rows.length) return null;
  const tot = rows.reduce((a, r) => a + r.r, 0);
  if (!(tot > 0)) return null;
  return rows.map(r => ({ c: r.c, name: custName(r.c), ratio: r.r / tot, price: r.p })).sort((a, b) => b.ratio - a.ratio);
}
// aggregate a SKU's rows to channel level: {name: {ratio, price(weighted), custs:[...]}}
function skuChannelSplit(sku) {
  const rows = skuChannelRows(sku);
  if (!rows) return null;
  const ch = {};
  for (const r of rows) {
    const k = custChannel(r.c);
    const g = ch[k] || (ch[k] = { ratio: 0, pxr: 0, custs: [] });
    g.ratio += r.ratio;
    g.pxr += (r.price != null ? r.price : sku.asp || 0) * r.ratio;
    g.custs.push(r);
  }
  for (const g of Object.values(ch)) g.price = g.ratio ? g.pxr / g.ratio : 0;
  return ch;
}
function chanChip(sku) {
  if (!chanReady() || !CHANNEL_INDEX.skus[sku.code]) return '';
  return `<button class="chan-chip" data-chan-sku="${esc(sku.id)}" title="Channel split — forecast units & value allocated by each customer's historical share. Click for the breakdown.">⇄ Channels</button>`;
}
function openChannelDialog(id) {
  const sku = skuById.get(id);
  if (!sku) return;
  const split = skuChannelSplit(sku);
  const body = document.getElementById('channel-body');
  if (!split) { body.innerHTML = '<p class="muted-note">No channel index rows for this product.</p>'; }
  else {
    const r = RES.get(id) || {};
    const fcTot = (r.forecast || []).reduce((a, b) => a + b, 0);
    const names = [...CHANNELS.filter(c => split[c]), ...(split.Unmapped ? ['Unmapped'] : [])];
    const blended = Object.values(split).reduce((a, g) => a + g.pxr, 0);
    const rowsH = names.map(n => {
      const g = split[n];
      const u = fcTot * g.ratio;
      const custs = g.custs.map(c =>
        `<div class="chd-cust"><span>${esc(c.name || c.c)}</span><span>${(c.ratio * 100).toFixed(1)}%</span><span>${c.price != null ? '£' + c.price.toFixed(2) : '– (ASP)'}</span></div>`).join('');
      return `<div class="chd-ch${n === 'Unmapped' ? ' chd-unm' : ''}">
        <div class="chd-head"><b>${esc(n)}</b><span>${(g.ratio * 100).toFixed(1)}% of units</span>
          <span>${fmtU(u)} units</span><span>${fmtGBP(u * g.price)}</span><span>avg £${g.price.toFixed(2)}</span></div>
        <div class="chd-custs">${custs}</div></div>`;
    }).join('');
    const vsAsp = sku.asp > 0 ? ((blended - sku.asp) / sku.asp) * 100 : null;
    body.innerHTML = `<p class="muted-note">${esc(sku.code)} — full-year forecast <b>${fmtU(fcTot)}</b> units, allocated by
      each customer's historical unit share. Blended channel ASP <b>£${blended.toFixed(2)}</b>${vsAsp != null ? ` (${vsAsp >= 0 ? '+' : ''}${vsAsp.toFixed(1)}% vs the ASP in use £${(+sku.asp).toFixed(2)})` : ''}.
      ${split.Unmapped ? 'Assign customers to channels in Settings → Data → Sales channels.' : ''}</p>` + rowsH;
  }
  document.getElementById('channel-title').textContent = `${sku.code} — channel split`;
  document.getElementById('channel-dialog').showModal();
}
// whole-plan channel table for the Summary tab (forecast plan units × ratio × price)
function chanSummaryHtml() {
  if (!chanReady()) return '';
  const agg = {};
  let totU = 0, totV = 0, missing = 0;
  for (const sku of M.skus) {
    const rows = skuChannelRows(sku);
    const r = RES.get(sku.id);
    const fcTot = r ? r.forecast.reduce((a, b) => a + b, 0) : 0;
    if (!fcTot) continue;
    if (!rows) { missing++; continue; }
    for (const c of rows) {
      const k = custChannel(c.c);
      const g = agg[k] || (agg[k] = { u: 0, v: 0 });
      const u = fcTot * c.ratio;
      g.u += u;
      g.v += u * (c.price != null ? c.price : sku.asp || 0);
      totU += u;
      g.v && 0;
    }
  }
  totV = Object.values(agg).reduce((a, g) => a + g.v, 0);
  if (!totU) return '';
  const names = [...CHANNELS.filter(c => agg[c]), ...(agg.Unmapped ? ['Unmapped'] : [])];
  const rows = names.map(n => {
    const g = agg[n];
    return `<tr${n === 'Unmapped' ? ' class="chs-unm"' : ''}><td>${esc(n)}</td>
      <td class="r">${fmtU(g.u)}</td><td class="r">${(100 * g.u / totU).toFixed(1)}%</td>
      <td class="r">${fmtGBP(g.v)}</td><td class="r">${(100 * g.v / (totV || 1)).toFixed(1)}%</td>
      <td class="r">£${(g.v / (g.u || 1)).toFixed(2)}</td></tr>`;
  }).join('');
  return `<div class="sum-grand chan-sum">
    <div class="sum-sec-title">Channel split — full-year forecast (units × customer share × channel price)</div>
    <table class="chs-tbl"><thead><tr><th>Channel</th><th class="r">Units</th><th class="r">% units</th>
      <th class="r">Sales value</th><th class="r">% value</th><th class="r">Blended ASP</th></tr></thead>
    <tbody>${rows}</tbody></table>
    <div class="muted-note">${agg.Unmapped ? 'Unmapped customers — assign them to channels in Settings → Data → Sales channels. ' : ''}${missing ? missing + ' forecast products have no channel index rows (fall entirely outside the split).' : ''}</div></div>`;
}
// Settings → Data: customer → channel mapping editor
function renderChanMap() {
  const wrap = document.getElementById('chan-map-wrap');
  if (!wrap) return;
  if (!chanReady()) { wrap.innerHTML = '<p class="muted-note">Upload the Channel Sales export (Settings → File Imports → Channel Index) to list customers.</p>'; return; }
  const weight = {};
  for (const rows of Object.values(CHANNEL_INDEX.skus))
    for (const r of rows) weight[r.c] = (weight[r.c] || 0) + r.r;
  const custs = Object.keys(weight).sort((a, b) => weight[b] - weight[a]);
  wrap.innerHTML = `<table class="chm-tbl"><thead><tr><th>Customer</th><th>Name</th><th class="r">Weight</th><th>Channel</th></tr></thead><tbody>` +
    custs.map(c => `<tr><td>${esc(c)}</td><td>${esc(custName(c) || '—')}</td><td class="r">${weight[c].toFixed(1)}</td>
      <td><select class="chm-sel" data-cust="${esc(c)}">${['Unmapped', ...CHANNELS].map(n =>
        `<option${custChannel(c) === n ? ' selected' : ''}>${n}</option>`).join('')}</select></td></tr>`).join('') +
    `</tbody></table>`;
  wrap.querySelectorAll('.chm-sel').forEach(sel => sel.addEventListener('change', () => {
    if (sel.value === 'Unmapped') delete chanMap()[sel.dataset.cust];
    else chanMap()[sel.dataset.cust] = sel.value;
    markDirty();
  }));
}
async function chanFileChosen(e) {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  const status = document.getElementById('chanidx-status');
  status.textContent = 'Reading ' + file.name + '…';
  try {
    const r = await fetch('/api/parse-channelindex', { method: 'POST', body: await file.arrayBuffer() });
    const j = await r.json();
    if (!j.ok) { status.textContent = ''; alert('Could not read the file: ' + (j.error || 'unknown')); return; }
    const data = await (await fetch('/api/data?year=' + encodeURIComponent(YEAR))).json();
    CHANNEL_INDEX = data.channelIndex || null;
    const basisTxt = j.basis === 'ly' ? ' · weighted by last-year units (start-of-year fallback)'
      : j.basis === 'ty' ? ' · weighted by this-year units'
      : j.basis === 'online' ? ' · weighted by online (CB Online) units' : '';
    status.textContent = `Channel index loaded: ${j.skus} SKUs · ${j.rows} customer shares · ${j.customers} customer names${basisTxt}.`;
    renderUploadAges(); renderChanMap();
    if (currentView === 'plan') renderPlan();
  } catch (err) { status.textContent = ''; alert('Read error: ' + err.message); }
}

/* ================= Change log + revert =================
   The server logs every file upload ('upload') and data-edit save ('edit') with a
   timestamp and a before-snapshot (kept for the most recent 5 revertable changes).
   The topbar quick menu reverts the last few; the full history lives in
   Settings → Changelog. Reverting restores the affected data files and reloads. */
async function fetchChangelog() {
  try { const j = await (await fetch('/api/changelog')).json(); CHANGELOG = (j && j.entries) || []; }
  catch { CHANGELOG = []; }
}
function clTime(iso) {
  const d = new Date(iso);
  if (isNaN(d)) return iso || '';
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: '2-digit' })
    + ' ' + d.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
}
function clKindChip(k) {
  const m = { upload: ['cl-up', 'Upload'], edit: ['cl-ed', 'Edit'], revert: ['cl-rv', 'Revert'] };
  const [c, t] = m[k] || ['cl-ed', k];
  return `<span class="cl-kind ${c}">${esc(t)}</span>`;
}
function clRowHtml(e) {
  const cnt = e.count > 1 ? ` <span class="cl-count">×${e.count}</span>` : '';
  const rev = (e.revertable && !e.reverted)
    ? `<button class="cl-revert" data-cl="${esc(e.id)}" title="Undo this change — restores the data exactly as it was immediately before it">↺ Revert</button>` : '';
  const done = e.reverted ? '<span class="cl-done">reverted</span>' : '';
  return `<div class="cl-row"><span class="cl-time">${clTime(e.ts)}</span>${clKindChip(e.kind)}`
    + `<span class="cl-label">${esc(e.label)}${cnt}${e.detail ? `<span class="cl-detail">${esc(e.detail)}</span>` : ''}</span>${done}${rev}</div>`;
}
async function revertChange(id) {
  const e = CHANGELOG.find(x => x.id === id);
  if (!e) return;
  if (!confirm(`Revert this change?\n\n${e.label} — ${clTime(e.ts)}\n\nThis restores the data exactly as it was immediately before the change. Anything done since will be undone.`)) return;
  const status = document.getElementById('save-status'); status.textContent = 'Reverting…';
  try {
    const j = await (await fetch('/api/revert-change?id=' + encodeURIComponent(id), { method: 'POST' })).json();
    if (!j.ok) { status.textContent = ''; alert('Revert failed: ' + (j.error || 'unknown')); return; }
    await fetchChangelog();
    await loadYear(YEAR);          // reload the viewed year + globals from the restored files
    renderChangelog(); renderHistoryQuick();
    status.textContent = 'Change reverted';
    alert('Reverted: ' + e.label + '.');
  } catch (err) { status.textContent = ''; alert('Revert error: ' + err.message); }
}
// topbar quick menu — the 5 most recent revertable changes
function renderHistoryQuick() {
  const pop = document.getElementById('history-menu');
  if (!pop) return;
  const recent = CHANGELOG.filter(e => e.revertable && !e.reverted).slice(0, 5);
  const rows = recent.length ? recent.map(clRowHtml).join('') : '<div class="cl-empty">No revertable changes yet.</div>';
  pop.innerHTML = `<div class="cl-qhead">Recent changes</div><div class="cl-qlist">${rows}</div>`
    + `<div class="cl-qfoot"><a id="cl-fulllink">Full changelog →</a></div>`;
  const fl = document.getElementById('cl-fulllink');
  if (fl) fl.addEventListener('click', () => { closeHistoryMenu(); openSettings(); setSettingsTab('changelog'); renderChangelog(); });
  pop.querySelectorAll('.cl-revert').forEach(b => b.addEventListener('click', () => revertChange(b.dataset.cl)));
}
let CL_FILTER = 'all';
function renderChangelog() {
  const wrap = document.getElementById('changelog-list');
  if (!wrap) return;
  document.querySelectorAll('#changelog-panel .cl-fbtn').forEach(b => b.classList.toggle('on', b.dataset.f === CL_FILTER));
  const list = CHANGELOG.filter(e => CL_FILTER === 'all' || e.kind === CL_FILTER);
  wrap.innerHTML = list.length ? list.map(clRowHtml).join('') : '<div class="cl-empty">Nothing logged in this view yet.</div>';
  wrap.querySelectorAll('.cl-revert').forEach(b => b.addEventListener('click', () => revertChange(b.dataset.cl)));
}
function toggleHistoryMenu() {
  const pop = document.getElementById('history-menu');
  if (!pop) return;
  if (pop.classList.contains('open')) { closeHistoryMenu(); return; }
  fetchChangelog().then(() => {
    renderHistoryQuick();
    const r = document.getElementById('btn-history').getBoundingClientRect();
    pop.style.top = (r.bottom + 6) + 'px';
    pop.style.right = Math.max(8, window.innerWidth - r.right) + 'px';
    pop.classList.add('open');
    setTimeout(() => document.addEventListener('click', historyOutside), 0);
  });
}
function closeHistoryMenu() {
  const pop = document.getElementById('history-menu');
  if (pop) pop.classList.remove('open');
  document.removeEventListener('click', historyOutside);
}
function historyOutside(e) {
  if (!e.target.closest('#history-menu') && e.target.id !== 'btn-history') closeHistoryMenu();
}

/* ================= Arrivals view: upcoming container arrivals =================
   Joins the two global PO uploads (WEBSA Open PO lines <-> Qlik container bookings)
   with the loaded year's master (name / season / current stock by product code) —
   the live replacement for the manual "Container Arrivals Summary" workbook.
   Balance units = ordered − delivered (WEBSA outstanding); arrival date =
   delivery-to-CB else UK-port ETA (same convention as the Plan's PO row). */
let ARR_FILTER = '';
const ARR_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
// Booking lead time deducted from a PO's WEBSA due date to get its estimated
// booking date. User-adjustable in the Arrivals toolbar; persisted in SETTINGS.
const ARR_LEAD_DEFAULT = { sailing: 50, grace: 7, inland: 7 };
function arrLead() {
  const s = SETTINGS.arr_lead || {};
  const pick = k => Number.isFinite(+s[k]) ? Math.max(0, Math.round(+s[k])) : ARR_LEAD_DEFAULT[k];
  return { sailing: pick('sailing'), grace: pick('grace'), inland: pick('inland') };
}
function arrLeadDays() { const l = arrLead(); return l.sailing + l.grace + l.inland; }
// Estimated booking date for an outstanding PO = its WEBSA due date − the total lead.
function arrBookDate(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  dt.setUTCDate(dt.getUTCDate() - arrLeadDays());
  return `${dt.getUTCFullYear()}-${String(dt.getUTCMonth() + 1).padStart(2, '0')}-${String(dt.getUTCDate()).padStart(2, '0')}`;
}
function arrMonthLabel(key) { return key === 'none' ? 'No due date' : `${ARR_MONTHS[+key.slice(5) - 1]} ${key.slice(0, 4)}`; }
function arrTodayIso() { const n = new Date(); return `${n.getFullYear()}-${String(n.getMonth() + 1).padStart(2, '0')}-${String(n.getDate()).padStart(2, '0')}`; }
function arrTodayUk() { const n = new Date(); return `${String(n.getDate()).padStart(2, '0')}-${String(n.getMonth() + 1).padStart(2, '0')}-${n.getFullYear()}`; }   // DD-MM-YYYY for export note + filename
function arrDow(iso) { const [y, m, d] = iso.split('-').map(Number); return ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][new Date(Date.UTC(y, m - 1, d)).getUTCDay()]; }

// Booked = every future-dated Qlik container leg (one card per PO+leg; a split PO
// appears on each of its containers). Awaiting = outstanding WEBSA POs with no dated
// booking at all — shown by due date, flagged overdue when that's already past.
function buildArrivalEvents() {
  const today = arrTodayIso();
  const skuByCode = new Map(M.skus.map(s => [s.code, s]));
  const poLines = po => (((PO_WEBSA.pos[po] || {}).lines) || [])
    .filter(l => l.outstanding > 0 && l.code !== 'POMARKETING')
    .map(l => ({ code: l.code, qty: l.outstanding, due: l.due, sku: skuByCode.get(l.code) || null }));
  const booked = [], awaiting = [];
  for (const po in PO_CONTAINERS.dates) {
    const dated = (PO_CONTAINERS.dates[po] || []).filter(l => l.deliveryCB || l.etaPort);
    for (const leg of dated) {
      const date = leg.deliveryCB || leg.etaPort;
      if (date < today) continue;
      booked.push({ po, date, leg, split: dated.length > 1, lines: poLines(po),
        supplier: (PO_WEBSA.pos[po] && PO_WEBSA.pos[po].supplier) || leg.supplier || '' });
    }
  }
  for (const po in PO_WEBSA.pos) {
    const legs = PO_CONTAINERS.dates && PO_CONTAINERS.dates[po];
    if (legs && legs.some(l => l.deliveryCB || l.etaPort)) continue;
    const lines = poLines(po);
    if (!lines.length) continue;
    const due = lines.map(l => l.due).find(Boolean) || '';
    awaiting.push({ po, date: due, lines, overdue: !!due && due < today, supplier: PO_WEBSA.pos[po].supplier || '' });
  }
  booked.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : a.po < b.po ? -1 : 1);
  awaiting.sort((a, b) => (a.date || '9999') < (b.date || '9999') ? -1 : 1);
  return { booked, awaiting };
}
function arrMatch(ev, q) {
  if (!q) return true;
  if (ev.po.toLowerCase().includes(q) || (ev.supplier || '').toLowerCase().includes(q)) return true;
  if (ev.leg && (ev.leg.container || '').toLowerCase().includes(q)) return true;
  return ev.lines.some(l => l.code.toLowerCase().includes(q) || (l.sku && l.sku.name.toLowerCase().includes(q)));
}
function arrLinesHtml(lines) {
  if (!lines.length) return '<div class="arr-nolines">No outstanding product lines on the WEBSA report for this PO.</div>';
  const num = n => Math.round(n).toLocaleString('en-GB');
  const tot = lines.reduce((a, l) => a + l.qty, 0);
  return '<table class="arr-lines"><thead><tr><th>Product</th><th>Description</th><th>Season</th>'
    + '<th class="r">Balance units</th><th class="r">Current stock</th></tr></thead><tbody>'
    + lines.map(l => `<tr><td class="arr-code">${esc(l.code)}</td><td>${esc(l.sku ? l.sku.name : '—')}</td>`
      + `<td>${esc(l.sku ? (l.sku.season || '—') : '—')}</td><td class="r"><b>${num(l.qty)}</b></td>`
      + `<td class="r">${l.sku ? num(l.sku.stock_now || 0) : '—'}</td></tr>`).join('')
    + (lines.length > 1 ? `</tbody><tfoot><tr><td colspan="3">Total</td><td class="r"><b>${num(tot)}</b></td><td></td></tr></tfoot></table>` : '</tbody></table>');
}
function arrCardHtml(ev) {
  const lg = ev.leg || {};
  const st = (lg.status || '').toUpperCase();
  const badge = ev.leg
    ? `<span class="arr-badge ${st.includes('NOT') ? 'arr-notpaid' : 'arr-paid'}">${esc(lg.status || '—')}</span>`
    : `<span class="arr-badge ${ev.overdue ? 'arr-over' : 'arr-notpaid'}">${ev.overdue ? 'OVERDUE — NOT BOOKED' : 'NOT BOOKED'}</span>`;
  const dates = ev.leg
    ? `ETD ${fmtDate(lg.etd)} → UK port ${fmtDate(lg.etaPort)} → CB ${fmtDate(lg.deliveryCB)}`
    : `WEBSA due ${fmtDate(ev.date || null)}${ev.date ? ` · book by ${fmtDate(arrBookDate(ev.date))}` : ''}`;
  return `<div class="arr-card"><div class="arr-card-head">`
    + `<span class="po-chip po-clk" data-po="${esc(ev.po)}">${esc(ev.po)}</span>`
    + `<span class="arr-sup">${esc(ev.supplier || '—')}</span>${badge}`
    + (ev.split ? '<span class="arr-badge arr-splitb" title="This PO ships across more than one container — the balance shown is the whole PO\'s">SPLIT SHIPMENT</span>' : '')
    + (lg.container ? `<span class="arr-cno">${esc(lg.container)}</span>` : '')
    + `<span class="arr-dates">${dates}</span></div>`
    + arrLinesHtml(ev.lines) + '</div>';
}
function arrFilteredEvents() {
  const q = ARR_FILTER.trim().toLowerCase();
  const all = buildArrivalEvents();
  return { booked: all.booked.filter(ev => arrMatch(ev, q)), awaiting: all.awaiting.filter(ev => arrMatch(ev, q)) };
}
// Physical upcoming containers per 'YYYY-MM' (dedup by container no; PO as fallback key).
function arrMonthCounts(booked) {
  const byMonth = new Map();
  for (const ev of booked) {
    const k = ev.date.slice(0, 7);
    if (!byMonth.has(k)) byMonth.set(k, new Set());
    byMonth.get(k).add(ev.leg.container || ev.po);
  }
  return [...byMonth.keys()].sort().map(k => ({ key: k, label: arrMonthLabel(k), count: byMonth.get(k).size }));
}
// Outstanding POs to book per 'YYYY-MM', keyed by the month of their estimated
// booking date (WEBSA due − 64d) — how many containers must be booked each month.
function arrBookMonthCounts(awaiting) {
  const byMonth = new Map();
  for (const ev of awaiting) {
    const bd = arrBookDate(ev.date);
    const k = bd ? bd.slice(0, 7) : 'none';
    byMonth.set(k, (byMonth.get(k) || 0) + 1);
  }
  return [...byMonth.keys()].sort().map(k => ({ key: k, label: arrMonthLabel(k), count: byMonth.get(k) }));
}
// Wrap awaiting POs into alternately-shaded blocks by estimated booking month.
function arrBookMonthBlocks(awaiting, body) {
  const months = [];
  for (const ev of awaiting) {
    const bd = arrBookDate(ev.date);
    const key = bd ? bd.slice(0, 7) : 'none';
    if (!months.length || months[months.length - 1].key !== key) months.push({ key, label: arrMonthLabel(key), events: [] });
    months[months.length - 1].events.push(ev);
  }
  return months.map((mo, i) =>
    `<div class="arr-week${i % 2 ? ' wband' : ''}"><div class="arr-week-head">${esc(mo.label)}`
    + `<span class="awh-sub">est. booking month</span>`
    + `<span class="awh-n">${mo.events.length} to book</span></div>${body(mo.events)}</div>`).join('');
}
function arrBodyHtml() {
  const q = ARR_FILTER.trim().toLowerCase();
  const { booked, awaiting } = arrFilteredEvents();
  const landCards = arrMonthCounts(booked).map(m =>
    `<div class="arr-mcard"><b>${m.count}</b><span>${esc(m.label)} · landing UK</span></div>`).join('');
  const bookCards = arrBookMonthCounts(awaiting).map(m =>
    `<div class="arr-mcard arr-mbook"><b>${m.count}</b><span>${esc(m.label)} · to book</span></div>`).join('');
  const mcards = landCards + bookCards;
  // booked cards grouped into Mon–Sun week blocks (alternating band shading, like
  // the export), each holding its day sub-groups
  const days = arrWeekBlocks(booked, w => {
    const n = new Set(w.events.map(e => e.leg.container || e.po)).size;
    return `${n} container${n === 1 ? '' : 's'}`;
  }, evs => {
    let inner = '', cur = '';
    for (const ev of evs) {
      if (ev.date !== cur) { cur = ev.date; inner += `<div class="arr-day-head">${arrDow(ev.date)} ${fmtDate(ev.date)}</div>`; }
      inner += arrCardHtml(ev);
    }
    return inner;
  }) || `<div class="empty">${q ? 'No upcoming containers match the filter.' : 'No future-dated containers in the Qlik export.'}</div>`;
  const await_ = awaiting.length
    ? `<details class="arr-awaiting" open><summary>${awaiting.length} outstanding PO${awaiting.length === 1 ? '' : 's'} awaiting a container booking — grouped by estimated booking month (WEBSA due − ${arrLeadDays()} days)</summary>`
      + arrBookMonthBlocks(awaiting, evs => evs.map(arrCardHtml).join(''))
      + '</details>'
    : '';
  return `<div class="arr-mcards">${mcards}</div>${days}${await_}`;
}
// Monday of a date's calendar week + display label (plan week number when in-year,
// else just the w/c date — e.g. an overdue PO due back in a prior year).
function arrWeekInfo(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  dt.setUTCDate(dt.getUTCDate() - (dt.getUTCDay() + 6) % 7);
  const mon = `${dt.getUTCFullYear()}-${String(dt.getUTCMonth() + 1).padStart(2, '0')}-${String(dt.getUTCDate()).padStart(2, '0')}`;
  const wk = isoToWeek(iso);
  return { key: mon, label: wk ? `Week ${wk}` : `w/c ${fmtDate(mon)}`, sub: wk ? `w/c ${fmtDate(mon)}` : '' };
}
// Wrap sorted events into alternately-shaded week blocks. `counts(w)` renders the
// block's count chip, `body(events)` its cards. Returns '' when there are no events.
function arrWeekBlocks(events, counts, body) {
  const weeks = [];
  for (const ev of events) {
    const wi = ev.date ? arrWeekInfo(ev.date) : { key: 'none', label: 'No due date', sub: '' };
    if (!weeks.length || weeks[weeks.length - 1].key !== wi.key) weeks.push({ ...wi, events: [] });
    weeks[weeks.length - 1].events.push(ev);
  }
  return weeks.map((w, i) =>
    `<div class="arr-week${i % 2 ? ' wband' : ''}"><div class="arr-week-head">${esc(w.label)}`
    + (w.sub ? `<span class="awh-sub">${esc(w.sub)}</span>` : '')
    + `<span class="awh-n">${counts(w)}</span></div>${body(w.events)}</div>`).join('');
}
// Export the page (as filtered on screen) to a shareable .xlsx — the client sends
// its already-joined rows so the workbook always matches what the user is looking at.
async function exportArrivals() {
  const { booked, awaiting } = arrFilteredEvents();
  const line = l => ({ code: l.code, name: l.sku ? l.sku.name : '', season: l.sku ? (l.sku.season || '') : '',
    qty: l.qty, stock: l.sku ? Math.round(l.sku.stock_now || 0) : null });
  const flt = ARR_FILTER.trim();
  const payload = {
    generated: `${arrTodayUk()}${flt ? ` · filtered: "${flt}"` : ''}`,
    lead: arrLead(),
    months: arrMonthCounts(booked),
    booked: booked.map(ev => ({ date: ev.date, week: isoToWeek(ev.date) || null, po: ev.po, supplier: ev.supplier,
      container: ev.leg.container || '', status: ev.leg.status || '', split: ev.split,
      etd: ev.leg.etd || null, etaPort: ev.leg.etaPort || null, deliveryCB: ev.leg.deliveryCB || null,
      lines: ev.lines.map(line) })),
    awaiting: awaiting.map(ev => ({ date: ev.date || null, bookDate: arrBookDate(ev.date) || null, po: ev.po,
      supplier: ev.supplier, overdue: ev.overdue, lines: ev.lines.map(line) })),
  };
  const status = document.getElementById('save-status');
  status.textContent = 'Building arrivals workbook…';
  try {
    const r = await fetch('/api/export-arrivals', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    if (!r.ok) throw new Error('server ' + r.status);
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `Upcoming Containers ${arrTodayUk()}.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    status.textContent = 'Exported upcoming containers';
  } catch (e) { status.textContent = 'Export failed!'; alert('Arrivals export failed: ' + e.message); }
}
function renderArrivals() {
  const main = document.getElementById('main');
  if (!poDataReady()) {
    main.innerHTML = '<div class="arr-wrap"><div class="empty">Upload the WEBSA Open PO and Qlik Container exports (Settings → File Imports) to build this page.</div></div>';
    return;
  }
  const L = arrLead();
  const leadInput = (id, label, val) => `<label class="arr-lead-f">${label}<input type="number" id="${id}" min="0" max="365" step="1" value="${val}"> days</label>`;
  main.innerHTML = `<div class="arr-wrap"><div class="arr-top"><h2>Upcoming containers</h2>`
    + `<input id="arr-search" type="search" placeholder="Filter by PO, product, supplier, container…" value="${esc(ARR_FILTER)}">`
    + `<button id="btn-arr-export" title="Download this page as an Excel workbook to share with the team — respects the current filter">Export xlsx</button>`
    + `<span class="arr-note">Balance units = ordered − delivered (WEBSA Open PO) · arrival = delivery-to-CB, else UK-port ETA (Qlik) · current stock as of the last weekly data import · click a PO for full detail</span></div>`
    + `<div class="arr-leadbar"><span class="arr-lead-label" title="How far before a PO's due date the container must be booked. The est. booking date used for the ‘to book’ months is the due date minus this total.">Booking lead time</span>`
    + leadInput('arr-lead-sailing', 'Sailing', L.sailing)
    + leadInput('arr-lead-grace', 'Factory grace', L.grace)
    + leadInput('arr-lead-inland', 'UK inland', L.inland)
    + `<span class="arr-lead-total" id="arr-lead-total">= <b>${arrLeadDays()}</b> days before the PO due date</span></div>`
    + `<div id="arr-body">${arrBodyHtml()}</div></div>`;
  const inp = document.getElementById('arr-search');
  inp.addEventListener('input', () => { ARR_FILTER = inp.value; document.getElementById('arr-body').innerHTML = arrBodyHtml(); });
  document.getElementById('btn-arr-export').addEventListener('click', exportArrivals);
  ['arr-lead-sailing', 'arr-lead-grace', 'arr-lead-inland'].forEach(id =>
    document.getElementById(id).addEventListener('change', arrApplyLead));
}
// Read the three lead inputs, persist them, and re-render with the new booking dates.
function arrApplyLead() {
  const g = id => Math.max(0, Math.min(365, parseInt(document.getElementById(id).value, 10) || 0));
  const v = { sailing: g('arr-lead-sailing'), grace: g('arr-lead-grace'), inland: g('arr-lead-inland') };
  SETTINGS.arr_lead = v; markDirty();
  document.getElementById('arr-lead-sailing').value = v.sailing;   // reflect clamped values
  document.getElementById('arr-lead-grace').value = v.grace;
  document.getElementById('arr-lead-inland').value = v.inland;
  document.getElementById('arr-lead-total').innerHTML = `= <b>${arrLeadDays()}</b> days before the PO due date`;
  document.getElementById('arr-body').innerHTML = arrBodyHtml();
}

function setView(v) {
  currentView = v;
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.view === v));
  document.getElementById('sidebar').style.display = v === 'plan' ? '' : 'none';
  document.getElementById('main').classList.toggle('plan', v === 'plan');
  if (v === 'plan') renderPlan();
  else if (v === 'arrivals') renderArrivals();
  else if (v === 'warehouse') renderWarehouse();
  else renderSummary();
}

// Project a year's per-SKU year-end stock (committed orders PLUS proposed rebuys —
// matching what the Plan grid's Stock Closing shows) AND its rolled-forward "LY Sales"
// row (the prior year's realised actuals where they exist, then its forecast for the
// rest of the year), computed in a sandbox so it doesn't disturb the currently-loaded
// year. Chains a forecast year's starting stock + LY reference to the prior year's
// CURRENT data, so re-imported actuals and Order Forecast / Proposed Rebuy edits carry
// forward live (no rebuild needed). Returns Map id -> { stock, ly:[53] }.
async function computeYearEndStocks(year) {
  let data;
  try { data = await (await fetch('/api/data?year=' + encodeURIComponent(year))).json(); }
  catch { return null; }
  if (!data || data.error || !data.master) return null;
  const snap = { YEAR, M, ORDERS, SETTINGS, MODELED, CALIB, RES, AGG, PROPOSED, skuById, supByName };
  const profSnap = new Map(PROFILE_CACHE);
  try {
    YEAR = String(year);
    M = data.master;
    ORDERS = data.orders || {};
    SETTINGS = Object.assign({ multiplier: M.multiplier || 1, container_cbm: M.container_cbm || 68, capacities: M.capacities, cover_bands: defaultCoverBands() }, data.settings || {});
    if (!SETTINGS.capacities) SETTINGS.capacities = M.capacities;
    if (!SETTINGS.cover_bands || !SETTINGS.cover_bands.length) SETTINGS.cover_bands = defaultCoverBands();
    SETTINGS.seasonality = migrateSeasonality(Object.assign({ useWeather: true, weatherStrength: 0.5, lat: 52.77, lon: -1.21, locationName: 'Loughborough, UK', ownShapeWeight: 0.6, byYear: undefined }, SETTINGS.seasonality || {}));
    SETTINGS.rebuy = Object.assign({ mode: 'full', maxWait: 4, suppliers: null, coverTarget: REBUY_TARGET, coverBySeason: {}, peakBoost: 0, oneProductSuppliers: [], seasonalCutoff: 0.40, partialCbm: 28 }, SETTINGS.rebuy || {});
    const cwby = SETTINGS.current_week_by_year || {};
    SETTINGS.current_week = cwby[YEAR] || M.data_week || isoWeek(new Date());
    skuById = new Map(); supByName = new Map(); MODELED = new Map(); CALIB = null; PROPOSED = null; PROFILE_CACHE.clear();
    for (const sku of M.skus) { skuById.set(sku.id, sku); if (!ORDERS[sku.id]) ORDERS[sku.id] = zeros(); }
    for (const s of M.suppliers) supByName.set(s.name, s);
    restoreProposed(data.proposed || {});   // include the prior year's proposed rebuys in the projection
    buildModeledForecasts();
    // actuals exist for weeks 1..(data_week-1); the rest of the prior year is forecast
    const bnd = (+M.data_week || SETTINGS.current_week || WEEKS) - 1;
    const out = new Map();
    for (const sku of M.skus) {
      const r = computeSku(sku, combinedOrder(sku.id));   // committed + proposed → year-end stock
      const ly = sku.actual.map((a, w) => (w < bnd ? a : r.forecast[w]));   // realised + forecast
      out.set(sku.id, { stock: Math.max(0, Math.round(r.stock[WEEKS - 1])), ly });
    }
    out.lyActualWeeks = bnd;   // weeks 0..bnd-1 are prior-year actuals; bnd.. are forecast
    return out;
  } catch { return null; }
  finally {
    YEAR = snap.YEAR; M = snap.M; ORDERS = snap.ORDERS; SETTINGS = snap.SETTINGS; MODELED = snap.MODELED;
    CALIB = snap.CALIB; RES = snap.RES; AGG = snap.AGG; PROPOSED = snap.PROPOSED; skuById = snap.skuById; supByName = snap.supByName;
    PROFILE_CACHE.clear(); for (const [k, v] of profSnap) PROFILE_CACHE.set(k, v);
  }
}

/* ---- last-year comparison for the supplier cards (vs-LY deltas + stock ghost) ----
   Swaps the prior year's data into the globals, runs the same calc engine, and
   returns per-product-code weekly aggregates. Snapshot/restore mirrors
   computeYearEndStocks so the live year is untouched. Cached per year-load. */
let LY_CACHE = null;   // { year, byCode: Map(code -> {sales,fob,units,cbm,shv}) } or null
async function computeLyAgg(lyYear) {
  let data;
  try { data = await (await fetch('/api/data?year=' + encodeURIComponent(lyYear))).json(); }
  catch { return null; }
  if (!data || data.error || !data.master) return null;
  const snap = { YEAR, M, ORDERS, SETTINGS, MODELED, CALIB, RES, AGG, PROPOSED, skuById, supByName };
  const profSnap = new Map(PROFILE_CACHE);
  try {
    YEAR = String(lyYear);
    M = data.master;
    ORDERS = data.orders || {};
    SETTINGS = Object.assign({ multiplier: M.multiplier || 1, container_cbm: M.container_cbm || 68, capacities: M.capacities, cover_bands: defaultCoverBands() }, data.settings || {});
    if (!SETTINGS.capacities) SETTINGS.capacities = M.capacities;
    if (!SETTINGS.cover_bands || !SETTINGS.cover_bands.length) SETTINGS.cover_bands = defaultCoverBands();
    SETTINGS.seasonality = migrateSeasonality(Object.assign({ useWeather: true, weatherStrength: 0.5, lat: 52.77, lon: -1.21, locationName: 'Loughborough, UK', ownShapeWeight: 0.6, byYear: undefined }, SETTINGS.seasonality || {}));
    SETTINGS.rebuy = Object.assign({ mode: 'full', maxWait: 4, suppliers: null, coverTarget: REBUY_TARGET, coverBySeason: {}, peakBoost: 0, oneProductSuppliers: [], seasonalCutoff: 0.40, partialCbm: 28 }, SETTINGS.rebuy || {});
    // pin the LY year's OWN current week (its actuals/forecast + stock-projection
    // boundary) — without this the aggregate is computed with the live year's week,
    // which misplaces the whole stock curve and inflates the stock-holding ghost
    const cwby = SETTINGS.current_week_by_year || {};
    SETTINGS.current_week = cwby[YEAR] || M.data_week || isoWeek(new Date());
    skuById = new Map(); supByName = new Map(); MODELED = new Map(); CALIB = null; PROPOSED = null; PROFILE_CACHE.clear();
    for (const sku of M.skus) { skuById.set(sku.id, sku); if (!ORDERS[sku.id]) ORDERS[sku.id] = zeros(); }
    for (const s of M.suppliers) supByName.set(s.name, s);
    restoreProposed(data.proposed || {});      // committed + proposed, like the live cards
    buildModeledForecasts();
    const byCode = new Map();
    for (const sku of M.skus) {
      const r = computeSku(sku, combinedOrder(sku.id));
      let e = byCode.get(sku.code);
      if (!e) { e = { sales: zeros(), fob: zeros(), units: zeros(), cbm: zeros(), shv: zeros() }; byCode.set(sku.code, e); }
      for (let w = 0; w < WEEKS; w++) {
        e.sales[w] += r.value[w]; e.fob[w] += r.fobRow[w]; e.units[w] += r.orders[w];
        e.cbm[w] += r.cbmRow[w]; e.shv[w] += r.shv[w];
      }
    }
    return byCode;
  } catch { return null; }
  finally {
    YEAR = snap.YEAR; M = snap.M; ORDERS = snap.ORDERS; SETTINGS = snap.SETTINGS; MODELED = snap.MODELED;
    CALIB = snap.CALIB; RES = snap.RES; AGG = snap.AGG; PROPOSED = snap.PROPOSED; skuById = snap.skuById; supByName = snap.supByName;
    PROFILE_CACHE.clear(); for (const [k, v] of profSnap) PROFILE_CACHE.set(k, v);
  }
}
async function refreshLyCache() {
  const ly = String(+YEAR - 1);
  if (LY_CACHE && LY_CACHE.year === ly) return;      // already have it
  if (!YEARS.includes(ly)) { LY_CACHE = null; return; }   // earliest year → no LY
  const byCode = await computeLyAgg(ly);
  LY_CACHE = byCode ? { year: ly, byCode } : null;
}
function lyAggForCodes(codes) {   // sum the LY weekly arrays across a supplier's product codes
  if (!LY_CACHE) return null;
  const a = { sales: zeros(), fob: zeros(), units: zeros(), cbm: zeros(), shv: zeros() };
  let hit = false;
  for (const code of codes) {
    const e = LY_CACHE.byCode.get(code);
    if (!e) continue;
    hit = true;
    for (let w = 0; w < WEEKS; w++) { a.sales[w] += e.sales[w]; a.fob[w] += e.fob[w]; a.units[w] += e.units[w]; a.cbm[w] += e.cbm[w]; a.shv[w] += e.shv[w]; }
  }
  return hit ? a : null;
}

async function loadYear(year) {
  YEAR = String(year);
  const resp = await fetch('/api/data?year=' + encodeURIComponent(YEAR));
  const data = await resp.json();
  if (data.error) { document.getElementById('main').innerHTML = `<div class="empty">${esc(data.error)}</div>`; return; }
  M = data.master;
  ORDERS = data.orders || {};
  PO_WEBSA = data.poWebsa || null;          // global PO ↔ container linking (shared across years)
  CHANNEL_INDEX = data.channelIndex || null;
  PO_CONTAINERS = data.poContainers || null;
  SETTINGS = Object.assign({
    multiplier: M.multiplier || 1,
    container_cbm: M.container_cbm || 68,
    capacities: M.capacities,
    cover_bands: defaultCoverBands(),
  }, data.settings || {});
  if (!SETTINGS.capacities) SETTINGS.capacities = M.capacities;
  if (!SETTINGS.cover_bands || !SETTINGS.cover_bands.length) SETTINGS.cover_bands = defaultCoverBands();
  SETTINGS.seasonality = migrateSeasonality(Object.assign(
    { useWeather: true, weatherStrength: 0.5,
      lat: 52.77, lon: -1.21, locationName: 'Loughborough, UK', ownShapeWeight: 0.6 },
    SETTINGS.seasonality || {}));
  if (!SETTINGS.seasonality.byYear) SETTINGS.seasonality.byYear = {};
  SETTINGS.rebuy = Object.assign({ mode: 'full', maxWait: 4, suppliers: null,
    coverTarget: REBUY_TARGET, coverBySeason: {}, peakBoost: 0, oneProductSuppliers: [], seasonalCutoff: 0.40, partialCbm: 28 }, SETTINGS.rebuy || {});
  if (!SETTINGS.rebuy.coverBySeason) SETTINGS.rebuy.coverBySeason = {};
  if (!Array.isArray(SETTINGS.rebuy.oneProductSuppliers)) SETTINGS.rebuy.oneProductSuppliers = [];
  // Current week: for the year that actually contains today, always track the live
  // computer date (via highlightWeek, which reads M.week1_start) so the "current week"
  // stays accurate without re-saving — capped at data_week so the actuals/forecast
  // boundary never claims weeks that haven't been imported yet. Other (past/future)
  // years fall back to their saved override / import week.
  const cwby = SETTINGS.current_week_by_year || {};
  const dateWk = highlightWeek();
  SETTINGS.current_week = dateWk ? Math.min(dateWk, M.data_week || dateWk)
                                 : (cwby[YEAR] || M.data_week || isoWeek(new Date()));
  // reset per-year state
  skuById = new Map(); supByName = new Map(); MODELED = new Map(); CALIB = null; PROPOSED = null;
  for (const sku of M.skus) { skuById.set(sku.id, sku); if (!ORDERS[sku.id]) ORDERS[sku.id] = zeros(); }
  for (const s of M.suppliers) supByName.set(s.name, s);
  try { IMPORTED = await (await fetch('/api/imported-orders?year=' + encodeURIComponent(YEAR))).json(); } catch { IMPORTED = {}; }
  // keep the selected supplier across years when it exists in the new year,
  // else fall back to the last saved one, then the first supplier
  const wantSup = currentSupplier || loadPref('tp_supplier', null);
  currentSupplier = (wantSup && M.suppliers.some(s => s.name === wantSup))
    ? wantSup : (M.suppliers[0] && M.suppliers[0].name);
  searchTerm = '';   // statusFilter persists across years (saved in localStorage)
  const sb = document.getElementById('search'); if (sb) sb.value = '';
  const ys = document.getElementById('year-select'); if (ys) ys.value = YEAR;
  refreshBuildYearBtn();   // button targets the year AFTER the one now being viewed
  try { localStorage.setItem('tp_year', YEAR); } catch {}

  // forecast year (all-forecast, data_week=1): re-derive each product's starting
  // stock AND its "LY Sales" row from the prior year's CURRENT data, so re-imported
  // actuals and Order Forecast / Proposed Rebuy edits flow through live (no rebuild).
  let stockChainChanged = false;
  if (+M.data_week === 1 && YEARS.includes(String(+YEAR - 1))) {
    const ends = await computeYearEndStocks(String(+YEAR - 1));
    if (ends) {
      M.lyActualWeeks = ends.lyActualWeeks;   // boundary where "LY Sales" switches actuals→forecast
      const base = data.stockbase, newBase = {};
      for (const sku of M.skus) {
        const neu = ends.get(sku.id);
        if (neu == null) { newBase[sku.id] = sku.stock_now; continue; }
        const ref = (base && base[sku.id] != null) ? base[sku.id] : sku.stock_now;
        if (Math.abs(neu.stock - ref) > 0.5) stockChainChanged = true;
        sku.stock_now = neu.stock; newBase[sku.id] = neu.stock;
        if (neu.ly) sku.ly = neu.ly;   // live "LY Sales" = prior year's realised actuals + forecast
      }
      persistStockbase(newBase);   // remember the basis so future loads only rebuild rebuys on a real change
    }
  }

  buildModeledForecasts();
  computeAll();
  // Restore the saved proposed layer EXACTLY as the user left it — including an empty
  // one, and including a year that has never been run (no suggestions at all). The
  // rebuy algorithm only ever runs when the user clicks "Run rebuy", so re-opening
  // the app / restarting the server never re-fills suggestions already cleared or
  // committed. If a forecast year's starting stock moved because the prior year was
  // edited, we only FLAG that the suggestions are stale (see the toolbar hint).
  restoreProposed(data.proposed || {});
  REBUY_STALE = stockChainChanged && PROPOSED.size > 0;
  computeAll();                    // band reflects the proposed layer
  await refreshLyCache();          // prior-year aggregates for the supplier cards' vs-LY deltas
  renderSidebar(); setView(currentView);   // stay on the page the user was viewing (Plan or Summary)
  if (seasonActiveFor(YEAR) && SEASON().useWeather) applySeasonality();   // refine with live weather (never touches proposals)
  document.title = `${YEAR} Tradeplan — week ${SETTINGS.current_week}`;
}

async function init() {
  let years = ['2026'], def = '2026';
  try { const yj = await (await fetch('/api/years')).json(); years = yj.years || years; def = yj.default || years[years.length - 1]; } catch {}
  YEARS = years;
  const sel = document.getElementById('year-select');
  sel.innerHTML = years.map(y => `<option value="${y}">${y}</option>`).join('');
  sel.addEventListener('change', async () => { if (saveTimer) await saveNow(); loadYear(sel.value); });
  document.getElementById('btn-build-year').addEventListener('click', buildForecastYearAction);
  refreshBuildYearBtn();
  let active = null;
  try { active = localStorage.getItem('tp_year'); } catch {}
  if (!years.includes(active)) active = def;

  initSparkTooltip();
  document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => setView(t.dataset.view)));
  document.getElementById('supplier-list').addEventListener('click', e => {
    const item = e.target.closest('.sup-item');
    if (item) { currentSupplier = item.dataset.sup; savePref('tp_supplier', currentSupplier); renderSidebar(); renderPlan(); }
  });
  document.getElementById('search').addEventListener('input', e => {
    searchTerm = e.target.value.trim();
    renderSidebar();
    if (currentView === 'plan') renderPlan();
  });
  document.querySelectorAll('.sortbtn').forEach(b => b.addEventListener('click', () => {
    sortMode = b.dataset.sort;
    savePref('tp_sortMode', sortMode);
    renderSidebar();
  }));
  document.getElementById('btn-theme').addEventListener('click', toggleTheme);
  updateThemeButton();
  document.getElementById('btn-settings').addEventListener('click', openSettings);
  document.getElementById('btn-history').addEventListener('click', e => { e.stopPropagation(); toggleHistoryMenu(); });
  document.querySelectorAll('#changelog-panel .cl-fbtn').forEach(b =>
    b.addEventListener('click', () => { CL_FILTER = b.dataset.f; renderChangelog(); }));
  document.getElementById('btn-settings-cancel').addEventListener('click', () => document.getElementById('settings-dialog').close());
  document.getElementById('settings-form').addEventListener('submit', e => { e.preventDefault(); document.getElementById('settings-dialog').close(); applySettings(); });
  document.getElementById('btn-add-band').addEventListener('click', () => {
    coverDraft.splice(Math.max(0, coverDraft.length - 1), 0, { max: 5, bg: '#ffd966' });
    renderCoverBands();
  });
  document.getElementById('btn-restore-bands').addEventListener('click', () => {
    coverDraft = defaultCoverBands().map(b => ({ max: b.max, bg: b.bg }));
    renderCoverBands();
  });
  document.getElementById('btn-restore-orders').addEventListener('click', restoreImportedOrders);
  // upload is triggered by the File Imports tiles (wireImportDnd); only the file-input change handlers live here
  document.getElementById('asp-file').addEventListener('change', aspFileChosen);
  document.getElementById('po-file').addEventListener('change', e => poFileChosen(e, 'po'));
  document.getElementById('containers-file').addEventListener('change', e => poFileChosen(e, 'containers'));
  document.getElementById('btn-po-diag').addEventListener('click', openPoDiag);
  document.getElementById('po-diag-close').addEventListener('click', () => document.getElementById('po-diag-dialog').close());
  document.getElementById('po-detail-close').addEventListener('click', () => document.getElementById('po-detail-dialog').close());
  document.addEventListener('click', e => { const chip = e.target.closest && e.target.closest('.po-chip[data-po]'); if (chip) openPoDetail(chip.dataset.po); });
  document.getElementById('btn-pum-cycle').addEventListener('click', () => startPoReview());
  document.getElementById('btn-po-apply').addEventListener('click', openRetimeDialog);
  document.getElementById('po-apply-cancel').addEventListener('click', () => document.getElementById('po-apply-dialog').close());
  document.getElementById('po-apply-go').addEventListener('click', applyRetime);
  document.getElementById('btn-po-undo').addEventListener('click', undoRetime);
  document.getElementById('asp-cancel').addEventListener('click', () => document.getElementById('asp-dialog').close());
  document.getElementById('asp-apply').addEventListener('click', applyAspUpdates);
  document.getElementById('landed-file').addEventListener('change', landedFileChosen);
  document.getElementById('landed-cancel').addEventListener('click', () => document.getElementById('landed-dialog').close());
  document.getElementById('landed-apply').addEventListener('click', applyLandedUpdates);
  document.getElementById('landed-search').addEventListener('input', renderLandedPreview);
  document.getElementById('landed-list').addEventListener('input', e => { if (e.target.classList.contains('lc-in')) landedEditInput(e.target); });
  document.getElementById('buying-file').addEventListener('change', buyingFileChosen);
  document.getElementById('buying-cancel').addEventListener('click', () => document.getElementById('buying-dialog').close());
  document.getElementById('buying-apply').addEventListener('click', applyBuyingUpdates);
  document.getElementById('duty-file').addEventListener('change', dutyFileChosen);
  document.getElementById('duty-cancel').addEventListener('click', () => document.getElementById('duty-dialog').close());
  document.getElementById('duty-apply').addEventListener('click', applyDutyUpdates);
  document.getElementById('chanidx-file').addEventListener('change', chanFileChosen);
  document.getElementById('channel-close').addEventListener('click', () => document.getElementById('channel-dialog').close());
  document.getElementById('wksales-file').addEventListener('change', wksalesFileChosen);
  document.getElementById('wksales-cancel').addEventListener('click', () => document.getElementById('wksales-dialog').close());
  document.getElementById('wksales-apply').addEventListener('click', applyWksalesUpdates);
  document.getElementById('wksales-week').addEventListener('change', wksalesWeekNote);
  document.getElementById('asp-compare').addEventListener('click', e => {
    const tr = e.target.closest('.asp-cmp-row'); if (tr) aspChooseBasis(tr.dataset.basis);
  });
  document.getElementById('asp-edit-search').addEventListener('input', renderAspEditor);
  document.getElementById('asp-edit-stale-only').addEventListener('change', renderAspEditor);
  document.getElementById('btn-apply-manual-asp').addEventListener('click', applyManualAsp);
  document.getElementById('cbm-edit-search').addEventListener('input', renderCbmEditor);
  document.getElementById('cbm-edit-manual-only').addEventListener('change', renderCbmEditor);
  document.getElementById('btn-apply-manual-cbm').addEventListener('click', applyManualCbm);
  document.getElementById('btn-add-product').addEventListener('click', openAddProductDialog);
  document.getElementById('ap-supplier').addEventListener('change', apToggleNewSupplier);
  document.getElementById('ap-forecast').addEventListener('input', apForecastNote);
  document.getElementById('ap-season').addEventListener('change', apForecastNote);
  document.getElementById('ap-category').addEventListener('input', apForecastNote);
  document.getElementById('ap-cancel').addEventListener('click', () => document.getElementById('addprod-dialog').close());
  document.getElementById('ap-save').addEventListener('click', submitAddProduct);
  document.getElementById('ap-pack').addEventListener('input', apCartonRecalc);
  document.getElementById('ap-weightlimit').addEventListener('change', apCartonRecalc);
  document.getElementById('ap-carton-add').addEventListener('click', () => { AP_CARTONS.push({ l: '', w: '', h: '', kg: '' }); apCartonRender(); });
  document.getElementById('ap-cbm').addEventListener('input', () => { AP_CBM_MANUAL = true; });   // hand-edited CBM wins
  document.getElementById('carton-pack').addEventListener('input', cartonReadState);
  document.getElementById('carton-weightlimit').addEventListener('change', cartonRecalc);
  document.getElementById('carton-add').addEventListener('click', () => { CARTON_EDIT.cartons.push({ l: '', w: '', h: '', kg: '' }); cartonRenderRows(); });
  document.getElementById('carton-cancel').addEventListener('click', () => document.getElementById('carton-dialog').close());
  document.getElementById('carton-save').addEventListener('click', submitCartons);
  document.getElementById('btn-export').addEventListener('click', exportCsv);
  document.getElementById('btn-export-forecast').addEventListener('click', exportForecast);
  document.getElementById('btn-save-config').addEventListener('click', openSaveConfigDialog);
  document.getElementById('config-select').addEventListener('change', e => {
    const file = e.target.value; e.target.value = '';
    if (file) openLoadConfigDialog(file);
  });
  refreshConfigList();
  const ss = document.getElementById('season-strength');
  ss.addEventListener('input', () => document.getElementById('season-strength-val').textContent = ss.value + '%');
  const os = document.getElementById('season-ownshape');
  os.addEventListener('input', () => document.getElementById('season-ownshape-val').textContent = os.value + '%');
  const ws = document.getElementById('season-weatherStrength');
  ws.addEventListener('input', () => document.getElementById('season-weatherStrength-val').textContent = ws.value + '%');
  document.getElementById('target-value').addEventListener('input', updateTargetReadout);
  document.getElementById('fc-year-tabs').addEventListener('click', e => {
    const b = e.target.closest('.fc-yr-tab'); if (b) fcSelectYear(b.dataset.year);
  });
  document.querySelectorAll('input[name="fc-mode"]').forEach(r =>
    r.addEventListener('change', () => {
      const mode = (document.querySelector('input[name="fc-mode"]:checked') || {}).value;
      const box = document.getElementById('target-value');
      if (mode === 'target') { if (!box.value && fcSelectedYear === String(YEAR)) box.value = Math.round(forecastDemandTotal()); }
      else box.value = '';   // leaving Target mode clears the box so nothing stale lingers
      fcUpdateModeUI(); updateTargetReadout();
      if (fcSelectedYear) { FC_DRAFT[fcSelectedYear] = fcReadControls(); fcRenderYearTabs(); }
    }));
  document.getElementById('btn-season-revert').addEventListener('click', () => {
    fcWriteControls({ mode: 'off' });   // reset the selected year to Original
    applySettings();
  });
  document.querySelectorAll('.stab').forEach(b => b.addEventListener('click', () => setSettingsTab(b.dataset.tab)));
  document.getElementById('explain-close').addEventListener('click', () => document.getElementById('explain-dialog').close());
  document.getElementById('notice-close').addEventListener('click', () => document.getElementById('notice-dialog').close());
  document.getElementById('btn-season-diag').addEventListener('click', () => {
    const diag = document.getElementById('season-diag');
    if (diag.classList.contains('hidden')) { renderSeasonDiag(); diag.classList.remove('hidden'); }
    else diag.classList.add('hidden');
  });
  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); saveNow(); }
  });
  await loadYear(active);
}

init();

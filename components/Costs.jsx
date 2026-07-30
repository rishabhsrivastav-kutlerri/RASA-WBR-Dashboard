'use client';

import { useState, Fragment } from 'react';
import '@/lib/chartSetup';
import { Bar, Line } from 'react-chartjs-2';
import Table from './Table';
import { ExportCsvButton, ExportImageButton } from './ExportButtons';
import { fmtPct, fmtVarPCColored } from '@/lib/fmt';
import { weekInfoForLabel } from '@/lib/fiscalCalendar';

const LOCATION_OPTIONS = ['MVT', 'National Landing', 'Mosaic', 'Rockville', 'All Locations'];
const FILTER_LABEL_STYLE = { fontSize: 11, fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.06em' };
const TREND_COLORS = ['#9f7cef', '#f9a8d4', '#86efac', '#fcd34d', '#93c5fd', '#fb923c', '#5eead4', '#f87171'];
const ALL_CATEGORIES_KEY = '__all__';

// Format a cost variance as plain text (positive = over budget = bad).
function fmtV(v) {
  if (v == null || isNaN(v)) return '-';
  const n = Number(v);
  const abs = (Math.abs(n) * 100).toFixed(1) + '%';
  return n < 0 ? `(${abs})` : abs;
}
function varCls(v) {
  if (v == null || isNaN(v)) return 'neu';
  const n = Number(v);
  return n === 0 ? 'neu' : n > 0 ? 'neg' : 'pos';
}

// Location Compare (COGS/Labor) heat-map cell coloring — per-row normalized
// deviation from that row's own mean across locations (excludes the "All
// Locations" aggregate column, which is never heat-colored). Pink = higher
// than the row's average (worse/higher cost); teal = lower (better).
function heatStyle(v, vals) {
  const arr = vals.filter(x => x != null); if (v == null || arr.length < 2) return '';
  const mean = arr.reduce((a, b) => a + b, 0) / arr.length;
  const md = Math.max(...arr.map(x => Math.abs(x - mean))) || 1;
  const t = (v - mean) / md, a = (Math.abs(t) * 0.24).toFixed(3);
  return 'background:' + (t > 0 ? `rgba(226,16,116,${a})` : `rgba(0,167,165,${a})`);
}
// React's style prop needs an object, not a CSS string — thin wrapper over
// heatStyle above to extract just the background value.
function heatBg(v, vals) {
  const css = heatStyle(v, vals);
  return css ? css.slice('background:'.length) : undefined;
}
const REAL_LOCS = ['MVT', 'National Landing', 'Mosaic', 'Rockville'];

const SUB_TABS = [
  { id: 'single',  label: 'Single Restaurant' },
  { id: 'compare', label: 'Compare all stores' },
];
const BASE_VIEWS = [
  { id: 'weekly', label: 'Weekly' },
  { id: 'ptd',    label: 'Period to Date' },
];
// Trailing 4/8-week views — Period 7 Week 2 (Week of July 6) onwards, see
// showTrailing gate below. Same PCR sheet as weekly/PTD; budget is pulled the
// same way as PTD (Budget Import sets one target per period, not per week —
// see the trailing4/trailing8 budget logic in lib/xlsxParser.js).
const TRAILING_VIEWS = [
  { id: 'trailing4',  label: 'Trailing 4 Weeks' },
  { id: 'trailing8',  label: 'Trailing 8 Weeks' },
  { id: 'trailing12', label: 'Trailing 12 Weeks' },
];
// Chart Range (Single Restaurant only) — a separate filter from Reporting
// Period above, driving only the trendline charts' window of history. Full
// History / Trailing 26 Weeks rely on the extra, manually-uploaded
// historical PCR workbooks that extend each category's week-by-week history
// beyond the current file's own 12 weeks (see extraPcrFiles in
// lib/xlsxParser.js) — weeks without those extra files just show whatever
// history is available (gracefully capped at the current file's 12 weeks).
const CHART_RANGE_VIEWS = [
  { id: 'full',       label: 'Full History' },
  { id: 'trailing26', label: 'Trailing 26 Weeks' },
  { id: 'trailing12', label: 'Trailing 12 Weeks' },
];

function computeTotal(d) {
  const existing = d.find(r => /^totals?$/i.test(r.loc));
  if (existing) return existing;
  const rows = d.filter(r => !/^totals?$/i.test(r.loc));
  const n = rows.length || 1;
  const avg = k => rows.reduce((s, r) => s + (r[k] || 0), 0) / n;
  const laborAct = avg('laborAct'), laborBud = avg('laborBud');
  const cogsAct  = avg('cogsAct'),  cogsBud  = avg('cogsBud');
  const pcAct    = avg('pcAct'),    pcBud    = avg('pcBud');
  return { loc: 'Totals', laborAct, laborBud, cogsAct, cogsBud, pcAct, pcBud, varPC: pcAct - pcBud, primeMarginAct: 1 - pcAct };
}

// Cost rows store Labor/COGS/PC as percentages of sales, not dollars — a
// closed location can't just be dropped from a plain average of percentages.
// Convert each remaining location's % back to dollars against its own sales
// (actual % against actual sales, budget % against budget sales), sum those
// dollars across the remaining locations, then divide by their combined
// sales to re-derive the consolidated percentage.
function computeWeightedTotal(rows, salesRows) {
  const salesByLoc = {};
  salesRows.forEach(r => { salesByLoc[r.loc] = r; });

  let laborActD = 0, laborBudD = 0, cogsActD = 0, cogsBudD = 0;
  let salesActSum = 0, salesBudSum = 0;
  for (const r of rows) {
    const s = salesByLoc[r.loc] || {};
    const sAct = s.actual || 0, sBud = s.budget || 0;
    laborActD += (r.laborAct || 0) * sAct;
    laborBudD += (r.laborBud || 0) * sBud;
    cogsActD  += (r.cogsAct  || 0) * sAct;
    cogsBudD  += (r.cogsBud  || 0) * sBud;
    salesActSum += sAct;
    salesBudSum += sBud;
  }

  const laborAct = salesActSum !== 0 ? laborActD / salesActSum : 0;
  const laborBud = salesBudSum !== 0 ? laborBudD / salesBudSum : 0;
  const cogsAct  = salesActSum !== 0 ? cogsActD  / salesActSum : 0;
  const cogsBud  = salesBudSum !== 0 ? cogsBudD  / salesBudSum : 0;
  const pcAct = laborAct + cogsAct;
  const pcBud = laborBud + cogsBud;
  return { loc: 'Totals', laborAct, laborBud, cogsAct, cogsBud, pcAct, pcBud, varPC: pcAct - pcBud, primeMarginAct: 1 - pcAct };
}

export default function Costs({ data, prevData, userRole }) {
  const isAdmin = userRole === 'admin';
  const [subTab, setSubTab] = useState('single');
  const [view, setView] = useState('weekly');
  const [chartRange, setChartRange] = useState('trailing12');
  const [singleLoc, setSingleLoc] = useState('All Locations');
  const [expandedCogsCat, setExpandedCogsCat] = useState(null);
  const [expandedLaborCat, setExpandedLaborCat] = useState(null);
  const [compareSubCat, setCompareSubCat] = useState('food');
  const [trendCategory, setTrendCategory] = useState(ALL_CATEGORIES_KEY);

  // Period 7 Week 3 (Week of July 13) onwards is shown with the immediately
  // previous week's Costs data until its own PCR file has been uploaded — a
  // real per-location primeMarginAct on a non-Totals row is the signal that
  // PCR was actually applied for the current week, not just the old
  // Flash-Results-COSTS-sourced actuals.
  const curWeekInfo = weekInfoForLabel(data.label);
  const isP7W3Plus = !!curWeekInfo && (curWeekInfo.period > 7 || (curWeekInfo.period === 7 && curWeekInfo.weekInPeriod >= 3));
  const curHasPcr = (data.weekly?.costs || []).some(r => !/^totals?$/i.test(r.loc) && r.primeMarginAct != null);
  const effectiveData = (isP7W3Plus && !curHasPcr && prevData) ? prevData : data;

  // Trailing 4/8-week filters and the COGS/Labor category-breakdown features
  // (Location Compare tables, category trendlines) are only available from
  // Period 7 Week 2 (Week of July 6) onwards.
  const weekInfo = weekInfoForLabel(effectiveData.label);
  const showTrailing = !!weekInfo && (weekInfo.period > 7 || (weekInfo.period === 7 && weekInfo.weekInPeriod >= 2));
  const views = showTrailing ? [...BASE_VIEWS, ...TRAILING_VIEWS] : BASE_VIEWS;
  const activeView = views.find(v => v.id === view) ? view : 'weekly';

  const cogsCatByLoc  = effectiveData.costsByCategory?.cogs;
  const laborCatByLoc = effectiveData.costsByCategory?.labor;
  // Follows the same Weekly / Period to Date / Trailing 4 / Trailing 8 filter
  // as the rest of the tab — all four are present in the PCR sheet.
  const catValueField = ['ptd', 'trailing4', 'trailing8', 'trailing12'].includes(activeView) ? activeView : 'weekly';
  // How many weeks the current Reporting Period filter covers — used by the
  // Line Items table's "Moves vs Prior N Wk" column below to know how many
  // weeks back the "prior" comparison window should span.
  const periodWeeksCount = activeView === 'ptd' ? (weekInfo?.weekInPeriod || 1)
    : activeView === 'trailing4' ? 4
    : activeView === 'trailing8' ? 8
    : activeView === 'trailing12' ? 12
    : 1;
  const CAT_LOCS = ['MVT', 'National Landing', 'Mosaic', 'Rockville', 'All Locations'];
  const catValue = (catByLoc, loc, key) => {
    const cat = catByLoc?.[loc]?.find(c => c.key === key);
    // "Others" in the Location Compare table uses the sum of its 3 subcategories
    // rather than the sheet's own "Total Other Costs" row.
    if (key === 'others' && cat?.subRows?.length) {
      return cat.subRows.reduce((sum, sub) => sum + (sub[catValueField] || 0), 0);
    }
    return cat?.[catValueField];
  };
  const catSubValue = (catByLoc, loc, catKey, subKey) => {
    const cat = catByLoc?.[loc]?.find(c => c.key === catKey);
    return cat?.subRows?.find(s => s.key === subKey)?.[catValueField];
  };

  // Category trend charts (Single Restaurant tab) — window of weeks plotted
  // follows the separate Chart Range filter (Full History / Trailing 26 /
  // Trailing 12), independent of the Reporting Period filter that drives the
  // KPI cards/tables above. "Full History" shows every week available (which
  // may be fewer than 26/12 for weeks without the extra historical PCR
  // files — Infinity here just means "don't cap it").
  const weeksWindow = chartRange === 'trailing26' ? 26 : chartRange === 'trailing12' ? 12 : Infinity;
  // "Others" trendline matches the Location Compare table: sum of its 3
  // subcategories per week, not the sheet's own "Total Other Costs" row.
  const othersWeeksFromSubRows = cat => {
    const subRows = cat?.subRows || [];
    const weekCount = subRows[0]?.weeks?.length || 0;
    const out = [];
    for (let i = 0; i < weekCount; i++) {
      let sum = 0, any = false;
      for (const sub of subRows) {
        const v = sub.weeks?.[i]?.value;
        if (v != null) { sum += v; any = true; }
      }
      out.push({ label: subRows[0].weeks[i].label, value: any ? sum : null });
    }
    return out;
  };
  const trendPoints = src => {
    const weeks = (src?.key === 'others' && src.subRows?.length) ? othersWeeksFromSubRows(src) : (src?.weeks || []);
    return weeks.slice(-weeksWindow).map(w => w.value != null ? +(w.value * 100).toFixed(2) : null);
  };
  // Sum every top-level category's weekly values (per-week array, oldest→
  // newest, full sheet history) into one combined COGS or Labor % trend.
  const sumCategoryWeeks = categories => {
    const weekCount = categories[0]?.weeks?.length || 0;
    const out = [];
    for (let i = 0; i < weekCount; i++) {
      let sum = 0, any = false;
      for (const cat of categories) {
        const catWeeks = (cat.key === 'others' && cat.subRows?.length) ? othersWeeksFromSubRows(cat) : cat.weeks;
        const v = catWeeks?.[i]?.value;
        if (v != null) { sum += v; any = true; }
      }
      out.push(any ? sum : null);
    }
    return out;
  };

  const singleCogsCats  = cogsCatByLoc?.[singleLoc] || [];
  const singleLaborCats = laborCatByLoc?.[singleLoc] || [];
  const trendLabelsAll = (singleCogsCats[0]?.weeks || singleLaborCats[0]?.weeks || []).slice(-weeksWindow).map(w => w.label);

  // Prime Cost trendline (Single Restaurant) — combined COGS % + Labor % per week.
  const cogsWeeksSum  = sumCategoryWeeks(singleCogsCats);
  const laborWeeksSum = sumCategoryWeeks(singleLaborCats);
  const pcWeeksAll = cogsWeeksSum.map((c, i) => (c != null && laborWeeksSum[i] != null) ? c + laborWeeksSum[i] : null);
  const primeCostTrendData = pcWeeksAll.slice(-weeksWindow).map(v => v != null ? +(v * 100).toFixed(2) : null);
  const cogsTrendData  = cogsWeeksSum.slice(-weeksWindow).map(v => v != null ? +(v * 100).toFixed(2) : null);
  const laborTrendData = laborWeeksSum.slice(-weeksWindow).map(v => v != null ? +(v * 100).toFixed(2) : null);
  const primeCostChartData = {
    labels: trendLabelsAll,
    datasets: [
      {
        label: 'Prime Cost %',
        data: primeCostTrendData,
        borderColor: '#9f7cef',
        backgroundColor: 'transparent',
        borderWidth: 2.5,
        pointRadius: 0,
        pointHoverRadius: 4,
        tension: 0,
      },
      {
        label: 'COGS %',
        data: cogsTrendData,
        borderColor: '#93c5fd',
        backgroundColor: 'transparent',
        borderWidth: 2.5,
        pointRadius: 0,
        pointHoverRadius: 4,
        tension: 0,
      },
      {
        label: 'Labor %',
        data: laborTrendData,
        borderColor: '#fcd34d',
        backgroundColor: 'transparent',
        borderWidth: 2.5,
        pointRadius: 0,
        pointHoverRadius: 4,
        tension: 0,
      },
    ],
  };

  // COGS-by-category / Labor-by-category trendlines (Single Restaurant) — one
  // line per top-level category, for the selected location.
  const buildAllCatSeries = list => list.map(cat => ({ key: cat.key, label: cat.label, data: trendPoints(cat) }));
  const cogsByCatSeries  = buildAllCatSeries(singleCogsCats);
  const laborByCatSeries = buildAllCatSeries(singleLaborCats);
  const buildMultiLineChart = series => ({
    labels: trendLabelsAll,
    datasets: series.map((s, i) => ({
      label: s.label,
      data: s.data,
      borderColor: TREND_COLORS[i % TREND_COLORS.length],
      backgroundColor: 'transparent',
      borderWidth: 2.5,
      pointRadius: 0,
      pointHoverRadius: 4,
      tension: 0,
    })),
  });
  const cogsByCatChartData  = buildMultiLineChart(cogsByCatSeries);
  const laborByCatChartData = buildMultiLineChart(laborByCatSeries);

  // Category trend chart with a Food/Beverages/etc. filter (Single
  // Restaurant) — COGS categories only.
  const isAllCategories = trendCategory === ALL_CATEGORIES_KEY;
  const selectedTrendCat = isAllCategories ? null : (singleCogsCats.find(c => c.key === trendCategory) || singleCogsCats[0]);
  const trendSeries = isAllCategories
    ? cogsByCatSeries
    : selectedTrendCat
      ? (selectedTrendCat.subRows?.length
          ? selectedTrendCat.subRows.map(sub => ({
              key: sub.key,
              label: sub.label,
              data: trendPoints(singleCogsCats.find(c => c.key === selectedTrendCat.key)?.subRows?.find(s => s.key === sub.key)),
            }))
          : [{
              key: selectedTrendCat.key,
              label: selectedTrendCat.label,
              data: trendPoints(singleCogsCats.find(c => c.key === selectedTrendCat.key)),
            }])
      : [];
  const trendChartData = buildMultiLineChart(trendSeries);

  // Line Items table below the Category Trend chart — same set of line items
  // as the chart above (all COGS categories, or the selected category's
  // subcategories). The "% of Sales" column follows the Reporting Period
  // filter (Weekly/PTD/Trailing4/8/12), same as the KPI cards/tables above —
  // independent of the Chart Range filter driving the trend chart itself.
  // "Moves vs Prior N Wk" also follows Reporting Period: N = periodWeeksCount
  // (1 for Weekly, the elapsed weeks for PTD, 4/8/12 for the Trailing views).
  // The "prior" side is the average of the weeks immediately before the
  // current N-week window, taken from the line item's own full weekly
  // history (not limited to the Chart Range window) — requires a full prior
  // window of history to be available, else shows "-". For Weekly
  // specifically, the prior side is the average of the 4 weeks before the
  // current week (not just the single prior week).
  const priorWindowSize = activeView === 'weekly' ? 4 : periodWeeksCount;
  const lineItemWeeks = src => (src?.key === 'others' && src.subRows?.length) ? othersWeeksFromSubRows(src) : (src?.weeks || []);
  const lineItemValue = src => (src?.key === 'others' && src?.subRows?.length)
    ? src.subRows.reduce((sum, sub) => sum + (sub[catValueField] || 0), 0)
    : src?.[catValueField];
  const priorPeriodAverage = src => {
    const weeks = lineItemWeeks(src);
    const n = periodWeeksCount, p = priorWindowSize;
    const priorSlice = weeks.slice(-(n + p), -n);
    const priorVals = priorSlice.map(w => w.value).filter(v => v != null);
    if (priorSlice.length < p || !priorVals.length) return null;
    return priorVals.reduce((s, v) => s + v, 0) / priorVals.length;
  };
  const lineItemsSrc = isAllCategories
    ? singleCogsCats
    : selectedTrendCat
      ? (selectedTrendCat.subRows?.length
          ? (singleCogsCats.find(c => c.key === selectedTrendCat.key)?.subRows || [])
          : [singleCogsCats.find(c => c.key === selectedTrendCat.key)].filter(Boolean))
      : [];
  const lineItemsTable = lineItemsSrc.map(src => {
    const current = lineItemValue(src);
    const prior = priorPeriodAverage(src);
    const moves = (current != null && prior != null) ? current - prior : null;
    return { key: src.key, label: src.label, current, moves };
  });

  // Wins & Watch (Single Restaurant) — every trackable COGS/Labor
  // category+subcategory for the selected location, pooled into one flat
  // list, each compared cur vs. prev via the same period-comparison logic as
  // the Line Items table above (cur = lineItemValue, prev = priorPeriodAverage
  // — Weekly's prev is the avg of the prior 4 weeks, PTD/Trailing N's prev is
  // the avg of the N periods before the current N-period window). Lower cost
  // % is better here, so a negative delta (cur < prev) is a Win and a
  // positive delta is a Watch item. Small moves are discarded as noise.
  const WINS_WATCH_THRESHOLD = 0.001; // 0.1 percentage point of sales
  const WINS_WATCH_TOP_N = 7;
  const allTrackableItems = [
    ...singleCogsCats.flatMap(cat => [
      { key: 'cogs-' + cat.key, group: 'COGS', label: cat.label, src: cat },
      ...(cat.subRows || []).map(sub => ({ key: 'cogs-' + cat.key + '-' + sub.key, group: 'COGS', label: sub.label, src: sub })),
    ]),
    ...singleLaborCats.map(cat => ({ key: 'labor-' + cat.key, group: 'Labor', label: cat.label, src: cat })),
  ];
  const winsWatchRows = allTrackableItems
    .map(item => {
      const cur = lineItemValue(item.src);
      const prev = priorPeriodAverage(item.src);
      const delta = (cur != null && prev != null) ? cur - prev : null;
      return { ...item, cur, prev, delta };
    })
    .filter(r => r.delta != null && Math.abs(r.delta) > WINS_WATCH_THRESHOLD);
  const wins = winsWatchRows.filter(r => r.delta < 0).sort((a, b) => a.delta - b.delta).slice(0, WINS_WATCH_TOP_N);
  const watch = winsWatchRows.filter(r => r.delta > 0).sort((a, b) => b.delta - a.delta).slice(0, WINS_WATCH_TOP_N);

  // Operator Notes (Single Restaurant) — the PCR Summary sheet's own COGS/
  // Labor notes for the selected location, only when something's actually
  // written ("No notes" placeholders are already filtered out in the
  // parser). "All Locations" shows every location's non-empty notes.
  const operatorNotesByLoc = effectiveData.operatorNotes || {};
  const operatorNotesList = singleLoc === 'All Locations'
    ? REAL_LOCS.flatMap(loc => {
        const n = operatorNotesByLoc[loc] || {};
        const out = [];
        if (n.cogs)  out.push({ key: loc + '-cogs',  label: `${loc} · COGS`,  text: n.cogs });
        if (n.labor) out.push({ key: loc + '-labor', label: `${loc} · Labor`, text: n.labor });
        return out;
      })
    : (() => {
        const n = operatorNotesByLoc[singleLoc] || {};
        const out = [];
        if (n.cogs)  out.push({ key: 'cogs',  label: 'COGS',  text: n.cogs });
        if (n.labor) out.push({ key: 'labor', label: 'Labor', text: n.labor });
        return out;
      })();

  // No grid lines behind the trendlines, and the x-axis (week labels) caps
  // how many ticks it shows and lets Chart.js auto-skip the rest — with
  // Trailing 26 Weeks/Full History plotting up to 36 weeks, showing every
  // single date label makes the axis unreadable.
  const trendChartOpts = {
    responsive: true,
    maintainAspectRatio: false,
    interaction: { mode: 'index', intersect: false },
    plugins: { legend: { position: 'bottom' } },
    scales: {
      x: { grid: { display: false }, ticks: { autoSkip: true, maxTicksLimit: 8 } },
      y: { grid: { display: false }, ticks: { callback: v => v + '%' } },
    },
  };
  // Category Trend chart's y-axis (below) rounds tick labels to 1 decimal —
  // Chart.js auto-generates tick step values that can carry long floating-
  // point tails (e.g. picking "Others") which the plain `v + '%'` callback
  // above prints verbatim.
  const categoryTrendOpts = {
    ...trendChartOpts,
    scales: {
      x: { grid: { display: false }, ticks: { autoSkip: true, maxTicksLimit: 8 } },
      y: { grid: { display: false }, ticks: { callback: v => Number(v).toFixed(1) + '%' } },
    },
  };

  const d = (effectiveData[activeView] && effectiveData[activeView].costs) || [];
  // Trailing 4 has no sales rows of its own (the PCR sheet doesn't break sales
  // out by trailing window), so its Totals row would otherwise fall back to
  // computeTotal's flat, Ballpark-included PCR sheet total instead of the
  // weighted, Ballpark-excluded total every other view uses. Sales-dollar
  // weights are a period-level concept anyway (same as budget — see
  // lib/xlsxParser.js), so PTD's sales rows are the correct weights to borrow.
  // Trailing 8's Totals row is excluded here — its overall Labor/COGS budget
  // is already the Budget Import sheet's own authoritative figure (set
  // directly on the Totals row in lib/xlsxParser.js), so it must go through
  // computeTotal below rather than being recomputed by computeWeightedTotal.
  const salesRows = (effectiveData[activeView] && effectiveData[activeView].sales)
    || (activeView === 'trailing4' ? (effectiveData.ptd?.sales || []) : []);
  const allRows = d.filter(r => !/^totals?$/i.test(r.loc));
  const hasBudget = allRows.some(r => r.laborBud != null);

  // Ballpark is permanently closed starting Period 7 (Week of June 29) —
  // drop it from the Costs tab from then on, no toggle needed.
  const excludeBallpark = !!weekInfo && weekInfo.period >= 7;
  const rows = excludeBallpark ? allRows.filter(r => r.loc !== 'Ballpark') : allRows;
  const locs = rows.map(r => r.loc);
  const total = (excludeBallpark && salesRows.length > 0) ? computeWeightedTotal(rows, salesRows) : computeTotal(d);
  const displayRows = excludeBallpark ? [...rows, total] : d;

  const varLabor = hasBudget ? (total.laborAct || 0) - (total.laborBud || 0) : null;
  const varCogs  = hasBudget ? (total.cogsAct  || 0) - (total.cogsBud  || 0) : null;
  const varPC    = hasBudget ? (total.varPC != null ? total.varPC : ((total.pcAct || 0) - (total.pcBud || 0))) : null;

  const baseOpts = {
    responsive: true,
    plugins: { legend: { position: 'bottom' } },
    scales: { y: { ticks: { callback: v => v + '%' }, min: 0 } },
  };
  const actualSeries = k => ({ label: `Actual ${k.label} %`, data: rows.map(r => +((r[k.key + 'Act'] || 0) * 100).toFixed(1)), backgroundColor: '#9f7cef', borderRadius: 4 });
  const budgetSeries = k => ({ label: `Budget ${k.label} %`, data: rows.map(r => +((r[k.key + 'Bud'] || 0) * 100).toFixed(1)), backgroundColor: '#93c5fd', borderRadius: 4 });
  const buildChart = k => ({
    labels: locs,
    datasets: hasBudget ? [actualSeries(k), budgetSeries(k)] : [actualSeries(k)],
  });
  const laborChart = buildChart({ key: 'labor', label: 'Labor' });
  const cogsChart   = buildChart({ key: 'cogs',  label: 'COGS' });
  const pcChart     = buildChart({ key: 'pc',    label: 'PC' });

  // Location Comparison (Single Restaurant) — horizontal Actual bars for
  // every location, with a black tick mark on each bar showing its Budget
  // target. The tick is a 'line'-type point (Chart.js built-in point style,
  // a short line rotated 90° so it reads as a vertical target mark) plotted
  // at the budget value, with the connecting line itself turned off
  // (showLine: false) so each location only gets its own isolated tick.
  const locCompOpts = {
    indexAxis: 'y',
    responsive: true,
    plugins: { legend: { position: 'bottom' } },
    scales: { x: { ticks: { callback: v => v + '%' }, min: 0 } },
  };
  const buildLocCompChart = k => ({
    labels: locs,
    datasets: [
      {
        type: 'bar',
        label: `Actual ${k.label} %`,
        data: rows.map(r => +((r[k.key + 'Act'] || 0) * 100).toFixed(1)),
        backgroundColor: '#9f7cef',
        borderRadius: 4,
        order: 0,
      },
      ...(hasBudget ? [{
        type: 'line',
        label: `Budget ${k.label} %`,
        data: rows.map(r => +((r[k.key + 'Bud'] || 0) * 100).toFixed(1)),
        showLine: false,
        order: 1,
        pointStyle: 'line',
        pointRotation: 90,
        pointRadius: 16,
        pointBorderWidth: 3,
        pointBorderColor: '#000',
        pointBackgroundColor: '#000',
        borderColor: '#000',
      }] : []),
    ],
  });
  const locCompCogsChart  = buildLocCompChart({ key: 'cogs',  label: 'COGS' });
  const locCompLaborChart = buildLocCompChart({ key: 'labor', label: 'Labor' });

  // Single Restaurant KPI cards — selected location's own row (or the
  // Totals row for "All Locations").
  const singleRowLoc = singleLoc === 'All Locations' ? 'Totals' : singleLoc;
  const singleRow = displayRows.find(r => r.loc === singleRowLoc) || total;
  const svarLabor = hasBudget ? (singleRow.laborAct || 0) - (singleRow.laborBud || 0) : null;
  const svarCogs  = hasBudget ? (singleRow.cogsAct  || 0) - (singleRow.cogsBud  || 0) : null;
  const svarPC    = hasBudget ? (singleRow.varPC != null ? singleRow.varPC : ((singleRow.pcAct || 0) - (singleRow.pcBud || 0))) : null;

  const showCatFeatures = showTrailing && !!effectiveData.costsByCategory;
  const showSingleTrends = showCatFeatures && weeksWindow > 0;

  return (
    <>
      {!showTrailing ? (
        // Legacy layout for weeks before Period 7 Week 2 (Week of July 6) —
        // unchanged from before the Single Restaurant / Compare all stores
        // sub-tabs existed: just the Weekly/PTD toggle, 3 KPI cards, 3
        // Actual-vs-Budget bar charts, and the All Locations table.
        <>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
            <span style={{ fontSize: 11, fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.06em' }}>Costs</span>
            <div className="toggle-group">
              {views.map(v => (
                <button key={v.id} className={`toggle-btn${activeView === v.id ? ' active' : ''}`} onClick={() => setView(v.id)}>{v.label}</button>
              ))}
            </div>
          </div>

          <div className="kpi-row">
            <div className="kpi-card">
              <div className="kpi-label">Actual Labor %</div>
              <div className="kpi-value">{fmtPct(total.laborAct)}</div>
              <div className="kpi-change neu">Bud: {fmtPct(total.laborBud)}</div>
              <div className={`kpi-change ${varCls(varLabor)}`}>Var: {fmtV(varLabor)}</div>
            </div>
            <div className="kpi-card">
              <div className="kpi-label">Actual COGS %</div>
              <div className="kpi-value">{fmtPct(total.cogsAct)}</div>
              <div className="kpi-change neu">Bud: {fmtPct(total.cogsBud)}</div>
              <div className={`kpi-change ${varCls(varCogs)}`}>Var: {fmtV(varCogs)}</div>
            </div>
            <div className="kpi-card">
              <div className="kpi-label">Prime Cost %</div>
              <div className="kpi-value">{fmtPct(total.pcAct)}</div>
              <div className="kpi-change neu">Bud: {fmtPct(total.pcBud)}</div>
              <div className={`kpi-change ${varCls(varPC)}`}>Var: {fmtV(varPC)}</div>
            </div>
          </div>

          <div className="charts-row">
            <div className="chart-card">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                <div className="chart-title" style={{ marginBottom: 0 }}>Labor % — Actual vs Budget</div>
                <ExportImageButton filename="Labor Actual vs Budget.png" />
              </div>
              <Bar data={laborChart} options={baseOpts} />
            </div>
            <div className="chart-card">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                <div className="chart-title" style={{ marginBottom: 0 }}>COGS % — Actual vs Budget</div>
                <ExportImageButton filename="COGS Actual vs Budget.png" />
              </div>
              <Bar data={cogsChart} options={baseOpts} />
            </div>
            <div className="chart-card">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                <div className="chart-title" style={{ marginBottom: 0 }}>Prime Cost % — Actual vs Budget</div>
                <ExportImageButton filename="Prime Cost Actual vs Budget.png" />
              </div>
              <Bar data={pcChart} options={baseOpts} />
            </div>
          </div>

          <div className="table-card">
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
              <div className="table-title" style={{ marginBottom: 0 }}>All Locations</div>
              <ExportCsvButton filename="All Locations.csv" />
            </div>
            <Table
              headers={[
                { label: 'Location' },
                { label: 'Labor Act', cls: 'right' },
                { label: 'Labor Bud', cls: 'right' },
                { label: 'COGS Act',  cls: 'right' },
                { label: 'COGS Bud',  cls: 'right' },
                { label: 'PC Act',    cls: 'right' },
                { label: 'PC Bud',    cls: 'right' },
                { label: 'Var PC',    cls: 'right' },
                { label: 'Prime Margin Act', cls: 'right' },
              ]}
              rows={displayRows.map(r => ({
                _cls: /^totals?$/i.test(r.loc) ? 'total-row' : '',
                cells: [r.loc, fmtPct(r.laborAct), fmtPct(r.laborBud), fmtPct(r.cogsAct), fmtPct(r.cogsBud), fmtPct(r.pcAct), fmtPct(r.pcBud), fmtVarPCColored(r.varPC), fmtPct(r.pcAct != null ? 1 - r.pcAct : null)],
              }))}
            />
          </div>
        </>
      ) : (
        <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-start', marginBottom: 16, flexWrap: 'wrap', gap: 14 }}>
        <span style={{ fontSize: 11, fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.06em' }}>Costs</span>
        <div className="toggle-group">
          {SUB_TABS.map(t => (
            <button key={t.id} className={`toggle-btn${subTab === t.id ? ' active' : ''}`} onClick={() => setSubTab(t.id)}>{t.label}</button>
          ))}
        </div>
      </div>

      {subTab === 'single' && (
        <>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 12, flexWrap: 'wrap', gap: 10 }}>
            <select
              value={singleLoc}
              onChange={e => setSingleLoc(e.target.value)}
              style={{ background: '#f3f4f6', border: '1px solid #e5e7eb', color: '#1a1f2e', padding: '6px 12px', borderRadius: 6, fontSize: 12, cursor: 'pointer', fontFamily: "'Montserrat',sans-serif" }}
            >
              {LOCATION_OPTIONS.map(loc => <option key={loc} value={loc}>{loc}</option>)}
            </select>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
              <span style={FILTER_LABEL_STYLE}>Reporting Period</span>
              <div className="toggle-group">
                {views.map(v => (
                  <button key={v.id} className={`toggle-btn${activeView === v.id ? ' active' : ''}`} onClick={() => setView(v.id)}>{v.label}</button>
                ))}
              </div>
            </div>
          </div>

          {showCatFeatures && (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-end', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
              <span style={FILTER_LABEL_STYLE}>Chart Range</span>
              <div className="toggle-group">
                {CHART_RANGE_VIEWS.map(c => (
                  <button key={c.id} className={`toggle-btn${chartRange === c.id ? ' active' : ''}`} onClick={() => setChartRange(c.id)}>{c.label}</button>
                ))}
              </div>
            </div>
          )}

          <div className="kpi-row">
            <div className="kpi-card">
              <div className="kpi-label">Prime Cost %</div>
              <div className="kpi-value">{fmtPct(singleRow.pcAct)}</div>
              <div className="kpi-change neu">Bud: {fmtPct(singleRow.pcBud)}</div>
              <div className={`kpi-change ${varCls(svarPC)}`}>Var: {fmtV(svarPC)}</div>
            </div>
            <div className="kpi-card">
              <div className="kpi-label">COGS %</div>
              <div className="kpi-value">{fmtPct(singleRow.cogsAct)}</div>
              <div className="kpi-change neu">Bud: {fmtPct(singleRow.cogsBud)}</div>
              <div className={`kpi-change ${varCls(svarCogs)}`}>Var: {fmtV(svarCogs)}</div>
            </div>
            <div className="kpi-card">
              <div className="kpi-label">Labor %</div>
              <div className="kpi-value">{fmtPct(singleRow.laborAct)}</div>
              <div className="kpi-change neu">Bud: {fmtPct(singleRow.laborBud)}</div>
              <div className={`kpi-change ${varCls(svarLabor)}`}>Var: {fmtV(svarLabor)}</div>
            </div>
            <div className="kpi-card">
              <div className="kpi-label">Prime Margin %</div>
              <div className="kpi-value">{fmtPct(singleRow.primeMarginAct)}</div>
            </div>
          </div>

          {showSingleTrends && (
            <div className="chart-card" style={{ marginBottom: 16 }}>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
                <div className="chart-title" style={{ marginBottom: 0 }}>Prime Cost Trend ({singleLoc}) · {CHART_RANGE_VIEWS.find(c => c.id === chartRange)?.label}</div>
                <ExportImageButton filename="Prime Cost Trend.png" />
              </div>
              <div style={{ height: 280 }}>
                <Line data={primeCostChartData} options={trendChartOpts} />
              </div>
            </div>
          )}

          {showSingleTrends && (
            <div className="charts-row">
              <div className="chart-card">
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                  <div className="chart-title" style={{ marginBottom: 0 }}>COGS by Category Trend ({singleLoc})</div>
                  <ExportImageButton filename="COGS by Category Trend.png" />
                </div>
                <div style={{ height: 280 }}>
                  <Line data={cogsByCatChartData} options={trendChartOpts} />
                </div>
              </div>
              <div className="chart-card">
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                  <div className="chart-title" style={{ marginBottom: 0 }}>Labor by Category Trend ({singleLoc})</div>
                  <ExportImageButton filename="Labor by Category Trend.png" />
                </div>
                <div style={{ height: 280 }}>
                  <Line data={laborByCatChartData} options={trendChartOpts} />
                </div>
              </div>
            </div>
          )}

          {showSingleTrends && (
            <div className="chart-card" style={{ marginTop: 16 }}>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 12, flexWrap: 'wrap', gap: 10 }}>
                <div className="chart-title" style={{ marginBottom: 0 }}>
                  Category Trend ({singleLoc}) · {CHART_RANGE_VIEWS.find(c => c.id === chartRange)?.label}
                </div>
                <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
                  <select
                    value={trendCategory}
                    onChange={e => setTrendCategory(e.target.value)}
                    style={{ background: '#f3f4f6', border: '1px solid #e5e7eb', color: '#1a1f2e', padding: '6px 12px', borderRadius: 6, fontSize: 12, cursor: 'pointer', fontFamily: "'Montserrat',sans-serif" }}
                  >
                    <option value={ALL_CATEGORIES_KEY}>All Categories</option>
                    {singleCogsCats.map(cat => <option key={cat.key} value={cat.key}>{cat.label}</option>)}
                  </select>
                  <ExportImageButton filename="Category Trend.png" />
                </div>
              </div>
              <div style={{ height: 320 }}>
                <Line data={trendChartData} options={categoryTrendOpts} />
              </div>
            </div>
          )}

          {showSingleTrends && (
            <div className="table-card">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
                <div className="table-title" style={{ marginBottom: 0 }}>Line Items ({singleLoc})</div>
                <ExportCsvButton filename="Line Items.csv" />
              </div>
              <table>
                <thead>
                  <tr>
                    <th>Line Items</th>
                    <th className="right">{views.find(v => v.id === activeView)?.label} % of Sales</th>
                    <th className="right">{activeView === 'weekly' ? 'Moves vs Avg of Prior 4 Wks' : `Moves vs Prior ${periodWeeksCount > 1 ? periodWeeksCount + ' ' : ''}Wk`}</th>
                  </tr>
                </thead>
                <tbody>
                  {lineItemsTable.map(li => (
                    <tr key={li.key}>
                      <td>{li.label}</td>
                      <td className="right">{fmtPct(li.current)}</td>
                      <td className="right" style={{ color: li.moves == null ? undefined : li.moves > 0 ? 'var(--red)' : li.moves < 0 ? 'var(--green)' : undefined }}>{fmtV(li.moves)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {showSingleTrends && (
            <div className="charts-row">
              <div className="table-card" style={{ marginBottom: 0 }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
                  <div className="table-title" style={{ marginBottom: 0 }}>Wins ({singleLoc})</div>
                  <ExportCsvButton filename="Wins.csv" />
                </div>
                <table>
                  <thead>
                    <tr>
                      <th>Category</th>
                      <th className="right">Current</th>
                      <th className="right">Delta</th>
                    </tr>
                  </thead>
                  <tbody>
                    {wins.map(r => (
                      <tr key={r.key}>
                        <td>{r.group} · {r.label}</td>
                        <td className="right">{fmtPct(r.cur)}</td>
                        <td className="right" style={{ color: 'var(--green)' }}>{fmtV(r.delta)}</td>
                      </tr>
                    ))}
                    {!wins.length && (
                      <tr><td colSpan={3} style={{ color: 'var(--muted)' }}>No material wins this period.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
              <div className="table-card" style={{ marginBottom: 0 }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
                  <div className="table-title" style={{ marginBottom: 0 }}>Watch ({singleLoc})</div>
                  <ExportCsvButton filename="Watch.csv" />
                </div>
                <table>
                  <thead>
                    <tr>
                      <th>Category</th>
                      <th className="right">Current</th>
                      <th className="right">Delta</th>
                    </tr>
                  </thead>
                  <tbody>
                    {watch.map(r => (
                      <tr key={r.key}>
                        <td>{r.group} · {r.label}</td>
                        <td className="right">{fmtPct(r.cur)}</td>
                        <td className="right" style={{ color: 'var(--red)' }}>{fmtV(r.delta)}</td>
                      </tr>
                    ))}
                    {!watch.length && (
                      <tr><td colSpan={3} style={{ color: 'var(--muted)' }}>No material watch items this period.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          <div className="charts-row">
            <div className="chart-card">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                <div className="chart-title" style={{ marginBottom: 0 }}>Location Comparison — COGS %</div>
                <ExportImageButton filename="Location Comparison COGS.png" />
              </div>
              <Bar data={locCompCogsChart} options={locCompOpts} />
            </div>
            <div className="chart-card">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                <div className="chart-title" style={{ marginBottom: 0 }}>Location Comparison — Labor %</div>
                <ExportImageButton filename="Location Comparison Labor.png" />
              </div>
              <Bar data={locCompLaborChart} options={locCompOpts} />
            </div>
          </div>

          {!!operatorNotesList.length && (
            <div className="table-card">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
                <div className="table-title" style={{ marginBottom: 0 }}>Operator Notes ({singleLoc})</div>
                <ExportCsvButton filename="Operator Notes.csv" />
              </div>
              <table>
                <tbody>
                  {operatorNotesList.map(n => (
                    <tr key={n.key}>
                      <td style={{ fontWeight: 600, whiteSpace: 'nowrap', width: 1 }}>{n.label}</td>
                      <td>{n.text}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}

      {subTab === 'compare' && (
        <>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-end', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
            <div className="toggle-group">
              {views.map(v => (
                <button key={v.id} className={`toggle-btn${activeView === v.id ? ' active' : ''}`} onClick={() => setView(v.id)}>{v.label}</button>
              ))}
            </div>
          </div>

          <div className="table-card">
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
              <div className="table-title" style={{ marginBottom: 0 }}>All Locations</div>
              <ExportCsvButton filename="All Locations.csv" />
            </div>
            <Table
              headers={[
                { label: 'Location' },
                { label: 'Labor Act', cls: 'right' },
                { label: 'Labor Bud', cls: 'right' },
                { label: 'COGS Act',  cls: 'right' },
                { label: 'COGS Bud',  cls: 'right' },
                { label: 'PC Act',    cls: 'right' },
                { label: 'PC Bud',    cls: 'right' },
                { label: 'Var PC',    cls: 'right' },
                { label: 'Prime Margin Act', cls: 'right' },
              ]}
              rows={displayRows.map(r => ({
                _cls: /^totals?$/i.test(r.loc) ? 'total-row' : '',
                cells: [r.loc, fmtPct(r.laborAct), fmtPct(r.laborBud), fmtPct(r.cogsAct), fmtPct(r.cogsBud), fmtPct(r.pcAct), fmtPct(r.pcBud), fmtVarPCColored(r.varPC), fmtPct(r.primeMarginAct)],
              }))}
            />
          </div>

          <div className="charts-row">
            <div className="chart-card">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                <div className="chart-title" style={{ marginBottom: 0 }}>Labor % — Actual vs Budget</div>
                <ExportImageButton filename="Labor Actual vs Budget.png" />
              </div>
              <Bar data={laborChart} options={baseOpts} />
            </div>
            <div className="chart-card">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                <div className="chart-title" style={{ marginBottom: 0 }}>COGS % — Actual vs Budget</div>
                <ExportImageButton filename="COGS Actual vs Budget.png" />
              </div>
              <Bar data={cogsChart} options={baseOpts} />
            </div>
            <div className="chart-card">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                <div className="chart-title" style={{ marginBottom: 0 }}>Prime Cost % — Actual vs Budget</div>
                <ExportImageButton filename="Prime Cost Actual vs Budget.png" />
              </div>
              <Bar data={pcChart} options={baseOpts} />
            </div>
          </div>

          {showCatFeatures && (
            <div className="charts-row">
              <div className="table-card" style={{ marginBottom: 0 }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
                  <div className="table-title" style={{ marginBottom: 0 }}>Location Compare — COGS</div>
                  <ExportCsvButton filename="Location Compare COGS.csv" />
                </div>
                <table>
                  <thead>
                    <tr>
                      <th>Category</th>
                      {CAT_LOCS.map(loc => <th key={loc} className="right">{loc}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {(cogsCatByLoc?.MVT || []).map(cat => {
                      const catVals = REAL_LOCS.map(loc => catValue(cogsCatByLoc, loc, cat.key));
                      return (
                        <Fragment key={cat.key}>
                          <tr
                            style={cat.subRows ? { cursor: 'pointer' } : undefined}
                            onClick={cat.subRows ? () => setExpandedCogsCat(expandedCogsCat === cat.key ? null : cat.key) : undefined}
                          >
                            <td>{cat.subRows ? (expandedCogsCat === cat.key ? '▾ ' : '▸ ') : ''}{cat.label}</td>
                            {CAT_LOCS.map(loc => {
                              const v = catValue(cogsCatByLoc, loc, cat.key);
                              const bg = loc === 'All Locations' ? undefined : heatBg(v, catVals);
                              return <td key={loc} className="right" style={bg ? { background: bg } : undefined}>{fmtPct(v)}</td>;
                            })}
                          </tr>
                          {cat.subRows && expandedCogsCat === cat.key && cat.subRows.map(sub => {
                            const subVals = REAL_LOCS.map(loc => catSubValue(cogsCatByLoc, loc, cat.key, sub.key));
                            return (
                              <tr key={sub.key}>
                                <td style={{ paddingLeft: 28, color: 'var(--muted)' }}>{sub.label}</td>
                                {CAT_LOCS.map(loc => {
                                  const v = catSubValue(cogsCatByLoc, loc, cat.key, sub.key);
                                  const bg = loc === 'All Locations' ? undefined : heatBg(v, subVals);
                                  return <td key={loc} className="right" style={bg ? { background: bg } : undefined}>{fmtPct(v)}</td>;
                                })}
                              </tr>
                            );
                          })}
                        </Fragment>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <div className="table-card" style={{ marginBottom: 0 }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
                  <div className="table-title" style={{ marginBottom: 0 }}>Location Compare — Labor</div>
                  <ExportCsvButton filename="Location Compare Labor.csv" />
                </div>
                <table>
                  <thead>
                    <tr>
                      <th>Category</th>
                      {CAT_LOCS.map(loc => <th key={loc} className="right">{loc}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {(laborCatByLoc?.MVT || []).map(cat => {
                      const catVals = REAL_LOCS.map(loc => catValue(laborCatByLoc, loc, cat.key));
                      return (
                        <Fragment key={cat.key}>
                          <tr
                            style={cat.subRows ? { cursor: 'pointer' } : undefined}
                            onClick={cat.subRows ? () => setExpandedLaborCat(expandedLaborCat === cat.key ? null : cat.key) : undefined}
                          >
                            <td>{cat.subRows ? (expandedLaborCat === cat.key ? '▾ ' : '▸ ') : ''}{cat.label}</td>
                            {CAT_LOCS.map(loc => {
                              const v = catValue(laborCatByLoc, loc, cat.key);
                              const bg = loc === 'All Locations' ? undefined : heatBg(v, catVals);
                              return <td key={loc} className="right" style={bg ? { background: bg } : undefined}>{fmtPct(v)}</td>;
                            })}
                          </tr>
                          {cat.subRows && expandedLaborCat === cat.key && cat.subRows.map(sub => {
                            const subVals = REAL_LOCS.map(loc => catSubValue(laborCatByLoc, loc, cat.key, sub.key));
                            return (
                              <tr key={sub.key}>
                                <td style={{ paddingLeft: 28, color: 'var(--muted)' }}>{sub.label}</td>
                                {CAT_LOCS.map(loc => {
                                  const v = catSubValue(laborCatByLoc, loc, cat.key, sub.key);
                                  const bg = loc === 'All Locations' ? undefined : heatBg(v, subVals);
                                  return <td key={loc} className="right" style={bg ? { background: bg } : undefined}>{fmtPct(v)}</td>;
                                })}
                              </tr>
                            );
                          })}
                        </Fragment>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {showCatFeatures && (
            <div className="table-card">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
                <div className="table-title" style={{ marginBottom: 0 }}>Location Compare — COGS Subcategories</div>
                <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
                  <select
                    value={compareSubCat}
                    onChange={e => setCompareSubCat(e.target.value)}
                    style={{ background: '#f3f4f6', border: '1px solid #e5e7eb', color: '#1a1f2e', padding: '6px 12px', borderRadius: 6, fontSize: 12, cursor: 'pointer', fontFamily: "'Montserrat',sans-serif" }}
                  >
                    {(cogsCatByLoc?.MVT || []).map(cat => <option key={cat.key} value={cat.key}>{cat.label}</option>)}
                  </select>
                  <ExportCsvButton filename="Location Compare COGS Subcategories.csv" />
                </div>
              </div>
              <table>
                <thead>
                  <tr>
                    <th>Subcategory</th>
                    {CAT_LOCS.map(loc => <th key={loc} className="right">{loc}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {((cogsCatByLoc?.MVT || []).find(c => c.key === compareSubCat)?.subRows || []).map(sub => {
                    const subVals = REAL_LOCS.map(loc => catSubValue(cogsCatByLoc, loc, compareSubCat, sub.key));
                    return (
                      <tr key={sub.key}>
                        <td>{sub.label}</td>
                        {CAT_LOCS.map(loc => {
                          const v = catSubValue(cogsCatByLoc, loc, compareSubCat, sub.key);
                          const bg = loc === 'All Locations' ? undefined : heatBg(v, subVals);
                          return <td key={loc} className="right" style={bg ? { background: bg } : undefined}>{fmtPct(v)}</td>;
                        })}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}
        </>
      )}
    </>
  );
}

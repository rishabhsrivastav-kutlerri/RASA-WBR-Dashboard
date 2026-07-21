'use client';

import { useState, Fragment } from 'react';
import '@/lib/chartSetup';
import { Bar, Line } from 'react-chartjs-2';
import Table from './Table';
import { fmtPct, fmtVarPCColored } from '@/lib/fmt';
import { weekInfoForLabel } from '@/lib/fiscalCalendar';

const TREND_LOCS = ['MVT', 'National Landing', 'Mosaic', 'Rockville', 'All Locations'];
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

const BASE_VIEWS = [
  { id: 'weekly', label: 'Weekly' },
  { id: 'ptd',    label: 'Period to Date' },
];
// Trailing 4/8-week views — Period 7 Week 2 (Week of July 6) only, see
// showTrailing gate below. Same PCR sheet as weekly/PTD; no budget data.
const TRAILING_VIEWS = [
  { id: 'trailing4', label: 'Trailing 4 Weeks' },
  { id: 'trailing8', label: 'Trailing 8 Weeks' },
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

export default function Costs({ data }) {
  const [view, setView] = useState('weekly');
  const [catFilter, setCatFilter] = useState('cogs');
  const [expandedCat, setExpandedCat] = useState(null);
  const [trendLoc, setTrendLoc] = useState('All Locations');
  const [trendCategory, setTrendCategory] = useState(ALL_CATEGORIES_KEY);

  // Trailing 4/8-week filters and the COGS/Labor location-compare table are
  // only available for Period 7 Week 2 (Week of July 6).
  const weekInfo = weekInfoForLabel(data.label);
  const showTrailing = !!weekInfo && weekInfo.period === 7 && weekInfo.weekInPeriod === 2;
  const views = showTrailing ? [...BASE_VIEWS, ...TRAILING_VIEWS] : BASE_VIEWS;
  const activeView = views.find(v => v.id === view) ? view : 'weekly';

  const catByLoc = catFilter === 'cogs' ? data.costsByCategory?.cogs : data.costsByCategory?.labor;
  const catList = catByLoc?.MVT || [];
  // Follows the same Weekly / Period to Date / Trailing 4 / Trailing 8 filter
  // as the rest of the tab — all four are present in the PCR sheet.
  const catValueField = ['ptd', 'trailing4', 'trailing8'].includes(activeView) ? activeView : 'weekly';
  const CAT_LOCS = ['MVT', 'National Landing', 'Mosaic', 'Rockville', 'All Locations'];
  const catValue = (loc, key) => {
    const cat = catByLoc?.[loc]?.find(c => c.key === key);
    // "Others" in the Location Compare table uses the sum of its 3 subcategories
    // rather than the sheet's own "Total Other Costs" row.
    if (key === 'others' && cat?.subRows?.length) {
      return cat.subRows.reduce((sum, sub) => sum + (sub[catValueField] || 0), 0);
    }
    return cat?.[catValueField];
  };
  const catSubValue = (loc, catKey, subKey) => {
    const cat = catByLoc?.[loc]?.find(c => c.key === catKey);
    return cat?.subRows?.find(s => s.key === subKey)?.[catValueField];
  };

  // Category trend chart — follows the same COGS/Labor filter as the table
  // above (catFilter) plus its own location + category filters. The window
  // of weeks plotted follows the top-level Weekly/PTD/Trailing4/Trailing8
  // filter: PTD shows the weeks elapsed so far this period, Trailing 4/8 show
  // that many most-recent weeks. Weekly has no multi-week window, so the
  // chart is hidden for that view. When the chosen category has subcategories
  // (e.g. Food), one line is plotted per subcategory.
  const trendCatByLoc = catByLoc;
  const trendCatList = catList;
  const isAllCategories = catFilter === 'labor' || trendCategory === ALL_CATEGORIES_KEY;
  const selectedTrendCat = isAllCategories ? null : (trendCatList.find(c => c.key === trendCategory) || trendCatList[0]);
  const weeksWindow = activeView === 'ptd' ? (weekInfo?.weekInPeriod || 1)
    : activeView === 'trailing4' ? 4
    : activeView === 'trailing8' ? 8
    : 0;
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
  const anyCatForLoc = trendCatByLoc?.[trendLoc]?.[0];
  const trendLabels = (anyCatForLoc?.weeks || []).slice(-weeksWindow).map(w => w.label);
  const locCatList = trendCatByLoc?.[trendLoc] || [];
  const trendSeries = isAllCategories
    ? locCatList.map(cat => ({ key: cat.key, label: cat.label, data: trendPoints(cat) }))
    : selectedTrendCat
      ? (selectedTrendCat.subRows?.length
          ? selectedTrendCat.subRows.map(sub => ({
              key: sub.key,
              label: sub.label,
              data: trendPoints(locCatList.find(c => c.key === selectedTrendCat.key)?.subRows?.find(s => s.key === sub.key)),
            }))
          : [{
              key: selectedTrendCat.key,
              label: selectedTrendCat.label,
              data: trendPoints(locCatList.find(c => c.key === selectedTrendCat.key)),
            }])
      : [];
  const trendChartData = {
    labels: trendLabels,
    datasets: trendSeries.map((s, i) => ({
      label: s.label,
      data: s.data,
      borderColor: TREND_COLORS[i % TREND_COLORS.length],
      backgroundColor: 'transparent',
      borderWidth: 2.5,
      pointRadius: 4,
      tension: 0.3,
    })),
  };
  const trendChartOpts = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: { legend: { position: 'bottom' } },
    scales: { y: { ticks: { callback: v => v + '%' } } },
  };

  const d = (data[activeView] && data[activeView].costs) || [];
  const salesRows = (data[activeView] && data[activeView].sales) || [];
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

  return (
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
          <div className="chart-title">Labor % — Actual vs Budget</div>
          <Bar data={laborChart} options={baseOpts} />
        </div>
        <div className="chart-card">
          <div className="chart-title">COGS % — Actual vs Budget</div>
          <Bar data={cogsChart} options={baseOpts} />
        </div>
        <div className="chart-card">
          <div className="chart-title">Prime Cost % — Actual vs Budget</div>
          <Bar data={pcChart} options={baseOpts} />
        </div>
      </div>

      <div className="table-card">
        <div className="table-title">Costs Detail by Location</div>
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

      {showTrailing && data.costsByCategory && (
        <div className="table-card">
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
            <div className="table-title" style={{ marginBottom: 0 }}>Location Compare — {catFilter === 'cogs' ? 'COGS' : 'Labor'}</div>
            <div className="toggle-group">
              <button className={`toggle-btn${catFilter === 'cogs' ? ' active' : ''}`} onClick={() => { setCatFilter('cogs'); setExpandedCat(null); setTrendCategory(ALL_CATEGORIES_KEY); }}>COGS</button>
              <button className={`toggle-btn${catFilter === 'labor' ? ' active' : ''}`} onClick={() => { setCatFilter('labor'); setExpandedCat(null); setTrendCategory(ALL_CATEGORIES_KEY); }}>Labor</button>
            </div>
          </div>
          <table>
            <thead>
              <tr>
                <th>Category</th>
                {CAT_LOCS.map(loc => <th key={loc} className="right">{loc}</th>)}
              </tr>
            </thead>
            <tbody>
              {catList.map(cat => (
                <Fragment key={cat.key}>
                  <tr
                    key={cat.key}
                    style={cat.subRows ? { cursor: 'pointer' } : undefined}
                    onClick={cat.subRows ? () => setExpandedCat(expandedCat === cat.key ? null : cat.key) : undefined}
                  >
                    <td>{cat.subRows ? (expandedCat === cat.key ? '▾ ' : '▸ ') : ''}{cat.label}</td>
                    {CAT_LOCS.map(loc => <td key={loc} className="right">{fmtPct(catValue(loc, cat.key))}</td>)}
                  </tr>
                  {cat.subRows && expandedCat === cat.key && cat.subRows.map(sub => (
                    <tr key={sub.key}>
                      <td style={{ paddingLeft: 28, color: 'var(--muted)' }}>{sub.label}</td>
                      {CAT_LOCS.map(loc => <td key={loc} className="right">{fmtPct(catSubValue(loc, cat.key, sub.key))}</td>)}
                    </tr>
                  ))}
                </Fragment>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {showTrailing && data.costsByCategory && weeksWindow > 0 && (
        <div className="chart-card" style={{ marginTop: 16 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 12, flexWrap: 'wrap', gap: 10 }}>
            <div className="chart-title" style={{ marginBottom: 0 }}>
              Category Trend ({trendLoc}) · {views.find(v => v.id === activeView)?.label}
            </div>
            <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
              <select
                value={trendLoc}
                onChange={e => setTrendLoc(e.target.value)}
                style={{ background: '#f3f4f6', border: '1px solid #e5e7eb', color: '#1a1f2e', padding: '6px 12px', borderRadius: 6, fontSize: 12, cursor: 'pointer', fontFamily: "'Montserrat',sans-serif" }}
              >
                {TREND_LOCS.map(loc => <option key={loc} value={loc}>{loc}</option>)}
              </select>
              {catFilter !== 'labor' && (
                <select
                  value={trendCategory}
                  onChange={e => setTrendCategory(e.target.value)}
                  style={{ background: '#f3f4f6', border: '1px solid #e5e7eb', color: '#1a1f2e', padding: '6px 12px', borderRadius: 6, fontSize: 12, cursor: 'pointer', fontFamily: "'Montserrat',sans-serif" }}
                >
                  <option value={ALL_CATEGORIES_KEY}>All Categories</option>
                  {trendCatList.map(cat => <option key={cat.key} value={cat.key}>{cat.label}</option>)}
                </select>
              )}
            </div>
          </div>
          <div style={{ height: 320 }}>
            <Line data={trendChartData} options={trendChartOpts} />
          </div>
        </div>
      )}
    </>
  );
}

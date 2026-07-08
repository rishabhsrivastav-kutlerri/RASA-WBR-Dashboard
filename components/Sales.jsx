'use client';

import { useState, useEffect } from 'react';
import '@/lib/chartSetup';
import { Bar, Doughnut } from 'react-chartjs-2';
import Table from './Table';
import { fmt$, fmtVarColored } from '@/lib/fmt';

const VIEWS = [
  { id: 'weekly', label: 'Weekly' },
  { id: 'ptd',    label: 'Period to Date' },
  { id: 'qtd',    label: 'Quarter to Date' },
  { id: 'ytd',    label: 'Year to Date' },
];

const ALL_LOCS = ['Ballpark', 'Mosaic', 'MVT', 'National Landing', 'Rockville'];

// Sum revenue-centre rows across multiple locations for open-only consolidated view.
function sumRC(data, view, openLocs) {
  const combined = {};
  for (const l of openLocs) {
    const locRC = data.revCenterByLoc?.[l]?.[view] || data.revCenterByLoc?.[l]?.weekly || [];
    for (const row of locRC) {
      if (row._isTotal) continue;
      if (!combined[row.center]) combined[row.center] = { actual: 0, ly: 0, budget: 0, hasBudget: false };
      combined[row.center].actual += row.actual || 0;
      combined[row.center].ly     += row.ly     || 0;
      if (row.budget != null) { combined[row.center].budget += row.budget; combined[row.center].hasBudget = true; }
    }
  }
  if (!Object.keys(combined).length) return null;
  return Object.entries(combined).map(([center, v]) => ({
    center,
    actual: v.actual,
    ly: v.ly,
    budget: v.hasBudget ? v.budget : null,
    varD: v.actual - v.ly,
    varP: v.ly !== 0 ? (v.actual - v.ly) / Math.abs(v.ly) : 'NA',
  }));
}

// Sum sub-category rows across multiple locations for open-only consolidated view.
function sumSC(data, viewKey, sub, openLocs) {
  const combined = {};
  let labelKey = 'sub';
  for (const l of openLocs) {
    const rows = data.subCatsByLoc?.[l]?.[viewKey]?.[sub] || data.subCatsByLoc?.[l]?.weekly?.[sub] || [];
    if (rows.length && rows[0].cat != null && rows[0].sub == null) labelKey = 'cat';
    for (const row of rows) {
      const key = String(row.sub ?? row.cat ?? '');
      if (!combined[key]) combined[key] = { actual: 0, ly: 0, lk: labelKey };
      combined[key].actual += row.actual || 0;
      combined[key].ly     += row.ly     || 0;
    }
  }
  if (!Object.keys(combined).length) return null;
  return Object.entries(combined).map(([label, v]) => ({
    [v.lk]: label,
    actual: v.actual,
    ly: v.ly,
    varD: v.actual - v.ly,
    varP: v.ly !== 0 ? (v.actual - v.ly) / Math.abs(v.ly) : 'NA',
  }));
}

const SUB_OPTIONS = [
  { id: 'all',      label: 'All (Revenue by Centre)' },
  { id: 'delivery', label: 'Delivery' },
  { id: 'pickup',   label: 'Pickup' },
  { id: 'offsites', label: 'Offsites' },
  { id: 'catering', label: 'Catering' },
];

const CAT_LABEL = { delivery: 'Delivery', pickup: 'Pickup (Takeout)', catering: 'Catering', offsites: 'Offsites' };
const PIE_COLORS = ['#9f7cef', '#f9a8d4', '#86efac', '#fcd34d', '#93c5fd', '#f9a8a8'];

export default function Sales({ data, prevData, openOnly, setOpenOnly, openLocSet }) {
  const [view, setView] = useState('weekly');
  const [sub, setSub]   = useState('all');
  const [loc, setLoc]   = useState('all');

  const views = data.qtdAvailable ? VIEWS : VIEWS.filter(v => v.id !== 'qtd');
  const vl = view === 'weekly' ? 'Weekly' : view === 'ptd' ? 'PTD' : view === 'qtd' ? 'QTD' : 'YTD';

  // Filter available location options; reset to 'all' if current loc becomes unavailable.
  const locOptions = openOnly && openLocSet?.size ? ALL_LOCS.filter(l => openLocSet.has(l)) : ALL_LOCS;
  useEffect(() => {
    if (openOnly && loc !== 'all' && openLocSet && !openLocSet.has(loc)) {
      setLoc('all');
    }
  }, [openOnly, openLocSet, loc]);

  // Var to Last Week — weekly view only. Pull the previous week's actuals for
  // whatever the current location/sub-category selection resolves to, and
  // return a colored % vs the matching row (or '-' when no match / not weekly).
  const isWeekly = view === 'weekly';
  const prevRC = isWeekly
    ? (loc !== 'all' ? prevData?.revCenterByLoc?.[loc]?.weekly : prevData?.revCenter?.weekly)
    : null;
  const prevSCAll = isWeekly
    ? (loc !== 'all' ? prevData?.subCatsByLoc?.[loc]?.weekly : prevData?.subCats?.weekly)
    : null;
  const lwCellFrom = (prevArr, keyField, label, actual) => {
    if (!prevArr || !prevArr.length) return '-';
    const p = /^total/i.test(String(label))
      ? prevArr.reduce((s, r) => s + (r.actual || 0), 0)
      : (prevArr.find(r => String(r[keyField]) === String(label)) || {}).actual;
    return (p != null && p !== 0) ? fmtVarColored((actual - p) / p) : '-';
  };

  // When open-only is active and viewing all-locations consolidated, compute
  // from per-location data to exclude closed locations.
  const openLocs = openOnly && openLocSet?.size ? ALL_LOCS.filter(l => openLocSet.has(l)) : null;

  let rc;
  if (loc !== 'all' && data.revCenterByLoc && data.revCenterByLoc[loc]) {
    rc = data.revCenterByLoc[loc][view] || data.revCenterByLoc[loc].weekly || [];
  } else if (openLocs && data.revCenterByLoc) {
    rc = sumRC(data, view, openLocs) || (data.revCenter && (data.revCenter[view] || data.revCenter.weekly)) || [];
  } else {
    rc = (data.revCenter && (data.revCenter[view] || data.revCenter.weekly)) || [];
  }

  const scView = (() => {
    if (openLocs && loc === 'all') {
      const viewKey = view === 'ytd' ? 'ytd' : view === 'ptd' ? 'ptd' : view === 'qtd' ? 'qtd' : 'weekly';
      const result = {};
      for (const s of ['delivery', 'pickup', 'offsites', 'catering']) {
        result[s] = sumSC(data, viewKey, s, openLocs) || [];
      }
      return result;
    }
    return (data.subCats && (data.subCats[view] || data.subCats.weekly)) || {};
  })();

  let chartTitle, pieTitle, tableTitle;
  let chartLabels, chartActual, chartLY;
  let tableHeaders, tableRows;

  const showBudget = sub === 'all';

  if (sub === 'all') {
    pieTitle   = `Actual Revenue Mix — ${vl}${loc !== 'all' ? ' · ' + loc : ''}`;
    tableTitle = `Revenue by Centre — ${loc !== 'all' ? loc : 'Consolidated'}`;

    const displayRC = rc.filter(r => r.center !== 'Discounts/Refunds' && r.center !== 'Delivery Fee' && !r._isTotal);
    chartLabels = displayRC.map(r => r.center);
    chartActual = displayRC.map(r => r.actual);
    chartLY     = displayRC.map(r => r.ly);
    const hasBudgetData = displayRC.some(r => r.budget != null && r.budget !== 0);
    chartTitle = `Revenue by Center — ${vl}${loc !== 'all' ? ' · ' + loc : ''} vs LY${hasBudgetData ? ' & Budget' : ''}`;

    // QTD and YTD: the parser includes a Totals row — use it directly.
    // Weekly / PTD: compute total from the data rows.
    const totalsRow = rc.find(r => r._isTotal);
    const dataRows  = rc.filter(r => !r._isTotal);
    let rcWithTotal;
    if (totalsRow) {
      rcWithTotal = [...dataRows, { ...totalsRow, center: 'Total' }];
    } else {
      const totalActual  = rc.reduce((s, r) => s + (r.actual || 0), 0);
      const totalLY      = rc.reduce((s, r) => s + (r.ly || 0), 0);
      const totalVarD    = totalActual - totalLY;
      const totalVarP    = totalLY === 0 ? 'NA' : totalVarD / Math.abs(totalLY);
      const hasBudget    = rc.some(r => r.budget != null);
      const totalBudget  = hasBudget ? rc.reduce((s, r) => s + (r.budget || 0), 0) : null;
      const totalVarDBud = totalBudget != null ? totalActual - totalBudget : null;
      const totalVarPBud = totalBudget != null && totalBudget !== 0 ? totalVarDBud / Math.abs(totalBudget) : null;
      rcWithTotal = [...rc, { center: 'Total', actual: totalActual, ly: totalLY, varD: totalVarD, varP: totalVarP, budget: totalBudget, varDBud: totalVarDBud, varPBud: totalVarPBud, _isTotal: true }];
    }

    tableHeaders = [
      { label: 'Revenue Centre' },
      { label: 'Actual', cls: 'right' },
      ...(isWeekly ? [{ label: <>Var % <span style={{ textTransform: 'none', fontSize: '0.85em' }}>vs</span> LW</>, cls: 'right' }] : []),
      { label: 'LY', cls: 'right' },
      { label: <>Var $ <span style={{ textTransform: 'none', fontSize: '0.85em' }}>vs</span> LY</>, cls: 'right' },
      { label: <>Var % <span style={{ textTransform: 'none', fontSize: '0.85em' }}>vs</span> LY</>, cls: 'right' },
      ...(showBudget ? [
        { label: 'Budget', cls: 'right' },
        { label: <>Var $ <span style={{ textTransform: 'none', fontSize: '0.85em' }}>vs</span> Bud</>, cls: 'right' },
        { label: <>Var % <span style={{ textTransform: 'none', fontSize: '0.85em' }}>vs</span> Bud</>, cls: 'right' },
      ] : []),
    ];
    tableRows = rcWithTotal.map(r => ({
      _cls: r._isTotal ? 'total-row' : '',
      cells: [
        r.center,
        fmt$(r.actual),
        ...(isWeekly ? [lwCellFrom(prevRC, 'center', r.center, r.actual)] : []),
        fmt$(r.ly),
        fmt$(r.varD),
        fmtVarColored(r.varP),
        ...(showBudget ? [
          fmt$(r.budget),
          fmt$(r.varDBud),
          fmtVarColored(r.varPBud),
        ] : []),
      ],
    }));
  } else {
    const lbl = CAT_LABEL[sub];

    let subData = null;
    const locSC = loc !== 'all' && data.subCatsByLoc ? data.subCatsByLoc[loc] : null;
    if (locSC) {
      const viewKey = view === 'ytd' ? 'ytd' : view === 'ptd' ? 'ptd' : view === 'qtd' ? 'qtd' : 'weekly';
      const locViewData = locSC[viewKey] && locSC[viewKey][sub] && locSC[viewKey][sub].length ? locSC[viewKey][sub] : null;
      if (locViewData) {
        subData = locViewData;
        tableTitle = `${lbl} — ${loc} (${vl})`;
      } else {
        subData = scView[sub] || [];
        tableTitle = `${lbl} — Sub-Category Breakdown`;
      }
    } else {
      subData = scView[sub] || [];
      tableTitle = `${lbl} — Sub-Category Breakdown`;
    }

    chartTitle = `${lbl} — ${vl}${loc !== 'all' ? ' · ' + loc : ''} vs LY`;
    pieTitle   = `${lbl} Revenue Mix — ${vl}${loc !== 'all' ? ' · ' + loc : ''}`;

    // Ritual is excluded from the Pickup breakdown across all views. Also drop
    // any existing Total row so it's recomputed from the remaining rows (i.e.
    // without Ritual's contribution).
    if (sub === 'pickup') {
      subData = subData.filter(r => {
        const name = String(r.sub ?? r.cat ?? '').trim();
        return !/^ritual$/i.test(name) && !r.isTotal && !/^total/i.test(name);
      });
    }

    const labelKey = subData.length && subData[0].sub != null ? 'sub' : 'cat';
    const chartData = subData.filter(r => !r.isTotal && !/^total/i.test(r[labelKey] || ''));
    chartLabels = chartData.map(r => r[labelKey]);
    chartActual = chartData.map(r => r.actual);
    chartLY     = chartData.map(r => r.ly);

    // Ensure a Total row exists. If the source already has one (YTD all-stores
    // tables include "Total"), keep it; otherwise compute from the data rows.
    let tableData = subData;
    const hasTotal = subData.some(r => r.isTotal || /^total/i.test(r[labelKey] || ''));
    if (!hasTotal && subData.length) {
      const tActual = subData.reduce((s, r) => s + (r.actual || 0), 0);
      const tLY     = subData.reduce((s, r) => s + (r.ly     || 0), 0);
      const tVarD   = tActual - tLY;
      // Any divide-by-zero (including 0/0) is undefined — render NA.
      const tVarP   = tLY === 0 ? 'NA' : tVarD / Math.abs(tLY);
      tableData = [...subData, { [labelKey]: 'Total', actual: tActual, ly: tLY, varD: tVarD, varP: tVarP, _isTotal: true }];
    }

    tableHeaders = [
      { label: 'Sub-Category' },
      { label: 'Actual', cls: 'right' },
      ...(isWeekly ? [{ label: <>Var % <span style={{ textTransform: 'none', fontSize: '0.85em' }}>vs</span> LW</>, cls: 'right' }] : []),
      { label: 'LY', cls: 'right' },
      { label: <>Var $ <span style={{ textTransform: 'none', fontSize: '0.85em' }}>vs</span> LY</>, cls: 'right' },
      { label: <>Var % <span style={{ textTransform: 'none', fontSize: '0.85em' }}>vs</span> LY</>, cls: 'right' },
    ];
    const prevSub = prevSCAll ? prevSCAll[sub] : null;
    tableRows = tableData.map(r => ({
      _cls: (r.isTotal || r._isTotal || /^total/i.test(r[labelKey] || '')) ? 'total-row' : '',
      cells: [r[labelKey], fmt$(r.actual), ...(isWeekly ? [lwCellFrom(prevSub, labelKey, r[labelKey], r.actual)] : []), fmt$(r.ly), fmt$(r.varD), fmtVarColored(r.varP)],
    }));
  }

  const chartBudget = showBudget
    ? rc.filter(r => !r._isTotal && r.center !== 'Discounts/Refunds' && r.center !== 'Delivery Fee').map(r => r.budget ?? 0)
    : null;

  const barData = {
    labels: chartLabels,
    datasets: [
      { label: 'Actual 2026', data: chartActual, backgroundColor: '#9f7cef',                borderRadius: 4 },
      { label: 'LY 2025',     data: chartLY,     backgroundColor: 'rgba(209,213,219,0.7)', borderRadius: 4 },
      ...(chartBudget ? [{ label: 'Budget', data: chartBudget, backgroundColor: '#93c5fd', borderRadius: 4 }] : []),
    ],
  };

  const pieIdx = chartActual
    .map((v, i) => ({ v, i }))
    .filter(o => o.v > 0);
  const pieData = {
    labels: pieIdx.map(o => chartLabels[o.i]),
    datasets: [{
      data: pieIdx.map(o => o.v),
      backgroundColor: PIE_COLORS.slice(0, pieIdx.length),
      borderWidth: 2,
      borderColor: '#fff',
    }],
  };

  return (
    <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
        <span style={{ fontSize: 11, fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.06em' }}>Revenue</span>
        <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
          {openLocSet && (
            <div className="toggle-group">
              <button className={`toggle-btn${!openOnly ? ' active' : ''}`} onClick={() => setOpenOnly(false)}>All Locations</button>
              <button className={`toggle-btn${openOnly ? ' active' : ''}`} onClick={() => setOpenOnly(true)}>Open Only</button>
            </div>
          )}
          <div className="toggle-group">
            {views.map(v => (
              <button key={v.id} className={`toggle-btn${view === v.id ? ' active' : ''}`} onClick={() => setView(v.id)}>{v.label}</button>
            ))}
          </div>
          <select
            value={loc}
            onChange={e => setLoc(e.target.value)}
            style={{ background: '#f3f4f6', border: '1px solid #e5e7eb', color: '#1a1f2e', padding: '6px 12px', borderRadius: 6, fontSize: 12, cursor: 'pointer', fontFamily: "'Montserrat',sans-serif" }}
          >
            <option value="all">All Locations</option>
            {locOptions.map(l => <option key={l} value={l}>{l}</option>)}
          </select>
        </div>
      </div>

      <div className="charts-row">
        <div className="chart-card" style={{ gridColumn: 'span 1' }}>
          <div className="chart-title">{chartTitle}</div>
          <Bar data={barData} options={{
            responsive: true,
            plugins: {
              legend: { position: 'bottom' },
              tooltip: {
                callbacks: {
                  label: ctx => {
                    const val = `${ctx.dataset.label}: ${fmt$(ctx.parsed.y)}`;
                    if (ctx.dataset.label === 'Actual 2026') {
                      const i   = ctx.dataIndex;
                      const ly  = chartLY?.[i];
                      const bud = chartBudget?.[i];
                      const lines = [val];
                      if (ly != null && ly !== 0) {
                        const pct = ((ctx.parsed.y - ly) / Math.abs(ly) * 100).toFixed(1);
                        lines.push(`  vs LY: ${pct >= 0 ? '+' : ''}${pct}%`);
                      }
                      if (bud != null && bud !== 0) {
                        const pct = ((ctx.parsed.y - bud) / Math.abs(bud) * 100).toFixed(1);
                        lines.push(`  vs Budget: ${pct >= 0 ? '+' : ''}${pct}%`);
                      }
                      return lines;
                    }
                    return val;
                  },
                },
              },
            },
          }} />
        </div>
        <div className="chart-card" style={{ gridColumn: 'span 1' }}>
          <div className="chart-title">{pieTitle}</div>
          <Doughnut data={pieData} options={{
            responsive: true,
            cutout: '55%',
            plugins: {
              legend: { position: 'bottom', labels: { font: { family: 'Montserrat', size: 11 }, color: '#6b7280', padding: 12 } },
              datalabels: {
                display: true,
                color: '#fff',
                font: { family: 'Montserrat', size: 12, weight: '700' },
                formatter: (value, ctx) => {
                  const total = ctx.dataset.data.reduce((a, b) => a + b, 0) || 1;
                  const pct = (value / total) * 100;
                  return pct >= 5 ? pct.toFixed(1) + '%' : '';
                },
              },
            },
          }} />
        </div>
      </div>

      <div className="table-card">
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 14, flexWrap: 'wrap', gap: 8 }}>
          <div className="table-title" style={{ marginBottom: 0 }}>{tableTitle}</div>
          <select
            value={sub}
            onChange={e => setSub(e.target.value)}
            style={{ background: '#f3f4f6', border: '1.5px solid var(--border)', color: '#1a1f2e', padding: '6px 14px', borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: 'pointer', fontFamily: "'Montserrat',sans-serif" }}
          >
            {SUB_OPTIONS.map(o => <option key={o.id} value={o.id}>{o.label}</option>)}
          </select>
        </div>
        <Table headers={tableHeaders} rows={tableRows} />
      </div>
    </>
  );
}

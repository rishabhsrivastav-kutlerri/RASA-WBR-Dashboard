'use client';

import { useState } from 'react';
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

const LOC_OPTIONS = ['Ballpark', 'Mosaic', 'MVT', 'National Landing', 'Rockville'];

const SUB_OPTIONS = [
  { id: 'all',      label: 'All (Revenue by Centre)' },
  { id: 'delivery', label: 'Delivery' },
  { id: 'pickup',   label: 'Pickup' },
  { id: 'offsites', label: 'Offsites' },
  { id: 'catering', label: 'Catering' },
];

const CAT_LABEL = { delivery: 'Delivery', pickup: 'Pickup (Takeout)', catering: 'Catering', offsites: 'Offsites' };
const PIE_COLORS = ['#9f7cef', '#f9a8d4', '#86efac', '#fcd34d', '#93c5fd', '#f9a8a8'];

export default function Sales({ data, prevData }) {
  const [view, setView] = useState('weekly');
  const [sub, setSub]   = useState('all');
  const [loc, setLoc]   = useState('all');

  const views = data.qtdAvailable ? VIEWS : VIEWS.filter(v => v.id !== 'qtd');
  const vl = view === 'weekly' ? 'Weekly' : view === 'ptd' ? 'PTD' : view === 'qtd' ? 'QTD' : 'YTD';

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

  let rc;
  if (loc !== 'all' && data.revCenterByLoc && data.revCenterByLoc[loc]) {
    rc = data.revCenterByLoc[loc][view] || data.revCenterByLoc[loc].weekly || [];
  } else {
    rc = (data.revCenter && (data.revCenter[view] || data.revCenter.weekly)) || [];
  }

  const scView = (data.subCats && (data.subCats[view] || data.subCats.weekly)) || {};

  let chartTitle, pieTitle, tableTitle;
  let chartLabels, chartActual, chartLY;
  let tableHeaders, tableRows;

  if (sub === 'all') {
    chartTitle = `Revenue by Center — ${vl}${loc !== 'all' ? ' · ' + loc : ''} vs LY`;
    pieTitle   = `Actual Revenue Mix — ${vl}${loc !== 'all' ? ' · ' + loc : ''}`;
    tableTitle = `Revenue by Centre — ${loc !== 'all' ? loc : 'Consolidated'}`;

    const displayRC = rc.filter(r => r.center !== 'Discounts/Refunds' && r.center !== 'Delivery Fee');
    chartLabels = displayRC.map(r => r.center);
    chartActual = displayRC.map(r => r.actual);
    chartLY     = displayRC.map(r => r.ly);

    const totalActual = rc.reduce((s, r) => s + (r.actual || 0), 0);
    const totalLY     = rc.reduce((s, r) => s + (r.ly || 0), 0);
    const totalVarD   = totalActual - totalLY;
    // Any divide-by-zero (including 0/0) is undefined — render NA.
    const totalVarP   = totalLY === 0 ? 'NA' : totalVarD / Math.abs(totalLY);
    const rcWithTotal = [...rc, { center: 'Total', actual: totalActual, ly: totalLY, varD: totalVarD, varP: totalVarP, _isTotal: true }];

    tableHeaders = [
      { label: 'Revenue Centre' },
      { label: 'Actual', cls: 'right' },
      ...(isWeekly ? [{ label: 'Var vs LW', cls: 'right' }] : []),
      { label: 'LY', cls: 'right' },
      { label: 'Var $', cls: 'right' },
      { label: 'Var %', cls: 'right' },
    ];
    tableRows = rcWithTotal.map(r => ({
      _cls: r._isTotal ? 'total-row' : '',
      cells: [r.center, fmt$(r.actual), ...(isWeekly ? [lwCellFrom(prevRC, 'center', r.center, r.actual)] : []), fmt$(r.ly), fmt$(r.varD), fmtVarColored(r.varP)],
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
      ...(isWeekly ? [{ label: 'Var vs LW', cls: 'right' }] : []),
      { label: 'LY', cls: 'right' },
      { label: 'Var $', cls: 'right' },
      { label: 'Var %', cls: 'right' },
    ];
    const prevSub = prevSCAll ? prevSCAll[sub] : null;
    tableRows = tableData.map(r => ({
      _cls: (r.isTotal || r._isTotal || /^total/i.test(r[labelKey] || '')) ? 'total-row' : '',
      cells: [r[labelKey], fmt$(r.actual), ...(isWeekly ? [lwCellFrom(prevSub, labelKey, r[labelKey], r.actual)] : []), fmt$(r.ly), fmt$(r.varD), fmtVarColored(r.varP)],
    }));
  }

  const barData = {
    labels: chartLabels,
    datasets: [
      { label: 'Actual 2026', data: chartActual, backgroundColor: '#9f7cef', borderRadius: 4 },
      { label: 'LY 2025',     data: chartLY,     backgroundColor: 'rgba(209,213,219,0.7)', borderRadius: 4 },
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
            {LOC_OPTIONS.map(l => <option key={l} value={l}>{l}</option>)}
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
              tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${fmt$(ctx.parsed.y)}` } },
            },
          }} />
        </div>
        <div className="chart-card" style={{ gridColumn: 'span 1' }}>
          <div className="chart-title">{pieTitle}</div>
          <Doughnut data={pieData} options={{
            responsive: true,
            cutout: '55%',
            plugins: { legend: { position: 'bottom', labels: { font: { family: 'Montserrat', size: 11 }, color: '#6b7280', padding: 12 } } },
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

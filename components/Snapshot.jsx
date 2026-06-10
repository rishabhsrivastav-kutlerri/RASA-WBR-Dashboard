'use client';

import { useState } from 'react';
import '@/lib/chartSetup';
import { Bar } from 'react-chartjs-2';
import Table from './Table';
import { fmt$, fmtVar, fmtVarColored } from '@/lib/fmt';

const VIEWS = [
  { id: 'weekly', label: 'Weekly' },
  { id: 'ptd',    label: 'Period to Date' },
  { id: 'qtd',    label: 'Quarter to Date' },
  { id: 'ytd',    label: 'Year to Date' },
];

export default function Snapshot({ data, prevData }) {
  const [view, setView] = useState('weekly');
  // QTD only exists in workbooks that ship the QTD sheets (newer weeks).
  const views = data.qtdAvailable ? VIEWS : VIEWS.filter(v => v.id !== 'qtd');
  const d = (data[view] && data[view].sales) || [];
  const total = d.find(r => /^totals?$/i.test(r.loc)) || d[d.length - 1] || {};
  const rows = d.filter(r => !/^totals?$/i.test(r.loc));
  const vl = view === 'weekly' ? 'Weekly' : view === 'ptd' ? 'PTD' : view === 'qtd' ? 'QTD' : 'YTD';

  // Var to Last Week — only meaningful for the weekly view; compare each
  // location's actual to the previous week's actual.
  const prevByLoc = {};
  if (view === 'weekly') (prevData?.weekly?.sales || []).forEach(r => { prevByLoc[r.loc] = r.actual; });
  const varLW = (loc, actual) => {
    const p = prevByLoc[loc];
    return (p != null && p !== 0) ? (actual - p) / p : null;
  };
  const lwCell = r => {
    const v = varLW(r.loc, r.actual);
    return v == null ? '-' : fmtVarColored(v);
  };
  const totalLW = varLW(total.loc || 'Totals', total.actual || 0);

  const salesChart = {
    labels: rows.map(r => r.loc),
    datasets: [
      { label: 'Actual', data: rows.map(r => r.actual), backgroundColor: '#9f7cef',                borderRadius: 4 },
      { label: 'LY',     data: rows.map(r => r.ly),     backgroundColor: 'rgba(209,213,219,0.7)',  borderRadius: 4 },
      { label: 'Budget', data: rows.map(r => r.budget), backgroundColor: '#93c5fd',                borderRadius: 4 },
    ],
  };
  const varChart = {
    labels: rows.map(r => r.loc),
    datasets: [{
      label: 'Var % vs LY',
      data: rows.map(r => +(r.varLY * 100).toFixed(1)),
      backgroundColor: rows.map(r => r.varLY >= 0 ? '#b99af3' : 'rgba(220,38,38,0.75)'),
      borderRadius: 4,
    }],
  };

  return (
    <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
        <span style={{ fontSize: 11, fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.06em' }}>Results</span>
        <div className="toggle-group">
          {views.map(v => (
            <button key={v.id} className={`toggle-btn${view === v.id ? ' active' : ''}`} onClick={() => setView(v.id)}>{v.label}</button>
          ))}
        </div>
      </div>

      <div className="kpi-row">
        <div className="kpi-card">
          <div className="kpi-label">Total Sales ({vl})</div>
          <div className="kpi-value">{fmt$(total.actual)}</div>
          {totalLW != null && (
            <div className={`kpi-change ${totalLW >= 0 ? 'pos' : 'neg'}`}>{fmtVar(totalLW)} vs last week</div>
          )}
        </div>
        <div className="kpi-card">
          <div className="kpi-label">Last Year (LY)</div>
          <div className="kpi-value">{fmt$(total.ly)}</div>
          <div className="kpi-change">
            {fmt$((total.actual || 0) - (total.ly || 0))} var &nbsp;|&nbsp;
            <span dangerouslySetInnerHTML={{ __html: fmtVarColored(total.varLY) }} />
          </div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">Budget (BUD)</div>
          <div className="kpi-value">{fmt$(total.budget)}</div>
          <div className="kpi-change">
            {fmt$((total.actual || 0) - (total.budget || 0))} var &nbsp;|&nbsp;
            <span dangerouslySetInnerHTML={{ __html: fmtVarColored(total.varBud) }} />
          </div>
        </div>
      </div>

      <div className="charts-row">
        <div className="chart-card">
          <div className="chart-title">Sales by Location</div>
          <Bar data={salesChart} options={{
            responsive: true,
            plugins: {
              legend: { position: 'bottom' },
              tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${fmt$(ctx.parsed.y)}` } },
            },
          }} />
        </div>
        <div className="chart-card">
          <div className="chart-title">Variance % vs LY</div>
          <Bar data={varChart} options={{
            responsive: true,
            plugins: { legend: { display: false } },
            scales: { y: { ticks: { callback: v => v + '%' } } },
          }} />
        </div>
      </div>

      <div className="table-card">
        <div className="table-title">Location Results</div>
        <Table
          headers={[
            { label: 'Location' },
            { label: 'Actual', cls: 'right' },
            ...(view === 'weekly' ? [{ label: <>Var <span style={{ textTransform: 'none', fontSize: '0.85em' }}>vs</span> LW</>, cls: 'right' }] : []),
            { label: 'Last Year (LY)', cls: 'right' },
            { label: 'Budget (BUD)', cls: 'right' },
            { label: <>Var <span style={{ textTransform: 'none', fontSize: '0.85em' }}>vs</span> LY</>, cls: 'right' },
            { label: <>Var <span style={{ textTransform: 'none', fontSize: '0.85em' }}>vs</span> Bud</>, cls: 'right' },
          ]}
          rows={d.map(r => ({
            _cls: /^totals?$/i.test(r.loc) ? 'total-row' : '',
            cells: [r.loc, fmt$(r.actual), ...(view === 'weekly' ? [lwCell(r)] : []), fmt$(r.ly), fmt$(r.budget), fmtVarColored(r.varLY), fmtVarColored(r.varBud)],
          }))}
        />
      </div>
    </>
  );
}

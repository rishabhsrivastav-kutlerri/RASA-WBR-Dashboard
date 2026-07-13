'use client';

import { useState } from 'react';
import '@/lib/chartSetup';
import { Bar } from 'react-chartjs-2';
import Table from './Table';
import { fmtPct, fmtVarPCColored } from '@/lib/fmt';

// Format a cost variance as plain text (positive = over budget = bad).
function fmtV(v) {
  const n = Number(v) || 0;
  const abs = (Math.abs(n) * 100).toFixed(1) + '%';
  return n < 0 ? `(${abs})` : abs;
}
function varCls(v) {
  const n = Number(v) || 0;
  return n === 0 ? 'neu' : n > 0 ? 'neg' : 'pos';
}

const VIEWS = [
  { id: 'weekly', label: 'Weekly' },
  { id: 'ptd',    label: 'Period to Date' },
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
  return { loc: 'Totals', laborAct, laborBud, cogsAct, cogsBud, pcAct, pcBud, varPC: pcAct - pcBud };
}

// Cost rows store Labor/COGS/PC as percentages of sales, not dollars — a
// closed location can't just be dropped from a plain average of percentages.
// Convert each open location's % back to dollars against its own sales
// (actual % against actual sales, budget % against budget sales), sum those
// dollars across the open locations, then divide by their combined sales to
// re-derive the consolidated percentage.
function computeOpenTotal(rows, salesRows) {
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
  return { loc: 'Totals', laborAct, laborBud, cogsAct, cogsBud, pcAct, pcBud, varPC: pcAct - pcBud };
}

export default function Costs({ data, openOnly, setOpenOnly, openLocSet }) {
  const [view, setView] = useState('weekly');
  const d = (data[view] && data[view].costs) || [];
  const salesRows = (data[view] && data[view].sales) || [];
  const allRows = d.filter(r => !/^totals?$/i.test(r.loc));

  const canFilter = openOnly && openLocSet?.size > 0;
  const rows = canFilter ? allRows.filter(r => openLocSet.has(r.loc)) : allRows;
  const locs = rows.map(r => r.loc);
  const total = canFilter ? computeOpenTotal(rows, salesRows) : computeTotal(d);
  const displayRows = canFilter ? [...rows, total] : d;

  const varLabor = (total.laborAct || 0) - (total.laborBud || 0);
  const varCogs  = (total.cogsAct  || 0) - (total.cogsBud  || 0);
  const varPC    = total.varPC != null ? total.varPC : ((total.pcAct || 0) - (total.pcBud || 0));

  const baseOpts = {
    responsive: true,
    plugins: { legend: { position: 'bottom' } },
    scales: { y: { ticks: { callback: v => v + '%' }, min: 0 } },
  };
  const laborChart = {
    labels: locs,
    datasets: [
      { label: 'Actual Labor %', data: rows.map(r => +((r.laborAct || 0) * 100).toFixed(1)), backgroundColor: '#9f7cef', borderRadius: 4 },
      { label: 'Budget Labor %', data: rows.map(r => +((r.laborBud || 0) * 100).toFixed(1)), backgroundColor: '#93c5fd', borderRadius: 4 },
    ],
  };
  const cogsChart = {
    labels: locs,
    datasets: [
      { label: 'Actual COGS %', data: rows.map(r => +((r.cogsAct || 0) * 100).toFixed(1)), backgroundColor: '#9f7cef', borderRadius: 4 },
      { label: 'Budget COGS %', data: rows.map(r => +((r.cogsBud || 0) * 100).toFixed(1)), backgroundColor: '#93c5fd', borderRadius: 4 },
    ],
  };
  const pcChart = {
    labels: locs,
    datasets: [
      { label: 'Actual PC %', data: rows.map(r => +((r.pcAct || 0) * 100).toFixed(1)), backgroundColor: '#9f7cef', borderRadius: 4 },
      { label: 'Budget PC %', data: rows.map(r => +((r.pcBud || 0) * 100).toFixed(1)), backgroundColor: '#93c5fd', borderRadius: 4 },
    ],
  };

  return (
    <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
        <span style={{ fontSize: 11, fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.06em' }}>Costs</span>
        <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
          {openLocSet && (
            <div className="toggle-group">
              <button className={`toggle-btn${!openOnly ? ' active' : ''}`} onClick={() => setOpenOnly(false)}>All Locations</button>
              <button className={`toggle-btn${openOnly ? ' active' : ''}`} onClick={() => setOpenOnly(true)}>Open Locations</button>
            </div>
          )}
          <div className="toggle-group">
            {VIEWS.map(v => (
              <button key={v.id} className={`toggle-btn${view === v.id ? ' active' : ''}`} onClick={() => setView(v.id)}>{v.label}</button>
            ))}
          </div>
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
          ]}
          rows={displayRows.map(r => ({
            _cls: /^totals?$/i.test(r.loc) ? 'total-row' : '',
            cells: [r.loc, fmtPct(r.laborAct), fmtPct(r.laborBud), fmtPct(r.cogsAct), fmtPct(r.cogsBud), fmtPct(r.pcAct), fmtPct(r.pcBud), fmtVarPCColored(r.varPC || 0)],
          }))}
        />
      </div>
    </>
  );
}

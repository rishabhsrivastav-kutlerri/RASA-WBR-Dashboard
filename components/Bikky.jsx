'use client';

import { useState } from 'react';
import '@/lib/chartSetup';
import { Bar } from 'react-chartjs-2';
import Table from './Table';
import { ExportCsvButton, ExportImageButton } from './ExportButtons';
import { fmt$2, fmtN, fmtPct, fmtVar } from '@/lib/fmt';

const SECTIONS = [
  { id: 'loc', label: 'Locations' },
  { id: 'acq', label: 'Customer Acquisition' },
  { id: 'onb', label: 'Customer Onboarding & Engagement' },
];

const LOC_PERIODS = [
  { id: 'weekly',  label: '7 Days' },
  { id: 'monthly', label: '30 Days' },
  { id: 'ninety',  label: '90 Days' },
  { id: 'ytd',     label: 'YTD' },
];

const ONB_PERIODS = [
  { id: 'weekly',  label: 'Weekly' },
  { id: 'monthly', label: 'Monthly' },
];

const SELECT_STYLE = {
  background: '#f3f4f6',
  border: '1.5px solid var(--border)',
  color: '#1a1f2e',
  padding: '6px 14px',
  borderRadius: 8,
  fontSize: 12,
  fontWeight: 600,
  cursor: 'pointer',
  fontFamily: "'Montserrat', sans-serif",
};

// Growth value as a colored badge: positive → green, negative → red.
// Returned as an HTML string so the Table renders it (it detects '<').
function growthCell(v) {
  if (v == null || isNaN(v)) return fmtVar(v);
  const cls = Number(v) >= 0 ? 'green' : 'red';
  return `<span class="badge ${cls}">${fmtVar(v)}</span>`;
}

// Inline variance chip for KPI cards — compares curr vs the previous period row.
// `label`: "vs LM" or "vs LW". `kind`: 'n' (count) | '$' (dollar) for showBoth mode.
function PrevKpiChip({ curr, prevRow, prevKey, label, showBoth = false, kind = 'n' }) {
  if (!prevRow) return <span className="kpi-change neu">— {label}</span>;
  const prev = prevRow[prevKey];
  if (prev == null || prev === 0 || isNaN(prev)) return <span className="kpi-change neu">— {label}</span>;
  const diff = Number(curr) - prev;
  const v = diff / Math.abs(prev);
  const cls = v >= 0 ? 'pos' : 'neg';
  const pctAbs = (Math.abs(v) * 100).toFixed(1) + '%';
  const pctTxt = v >= 0 ? `${pctAbs}` : `(${pctAbs})`;
  if (showBoth) {
    let absTxt, absLabel;
    if (kind === '$') {
      const n = Math.abs(diff).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      absTxt = diff >= 0 ? `$${n}` : `($${n})`;
      absLabel = 'Var $:';
    } else {
      const n = Math.round(Math.abs(diff)).toLocaleString('en-US');
      absTxt = diff >= 0 ? `${n}` : `(${n})`;
      absLabel = 'Var:';
    }
    return <span className={`kpi-change ${cls}`}>{absLabel} {absTxt}<br/>Var%: {pctTxt} {label}</span>;
  }
  return <span className={`kpi-change ${cls}`}>Var%: {pctTxt} {label}</span>;
}

function buildLocTotal(rows) {
  const list = rows.filter(r => !/^total$/i.test(r.loc));
  if (!list.length) return null;
  const orders = list.reduce((a, r) => a + (r.orders || 0), 0);
  const guests = list.reduce((a, r) => a + (r.guests || 0), 0);
  const newGuests = list.reduce((a, r) => a + (r.newGuests || 0), 0);
  const wAov = list.reduce((a, r) => a + (r.aov || 0) * (r.orders || 0), 0);
  const aov = orders ? wAov / orders : (list.reduce((a, r) => a + (r.aov || 0), 0) / list.length);
  return { loc: 'Total', orders, aov, guests, newGuests, ordersGrowth: 0, guestsGrowth: 0, newGuestsGrowth: 0 };
}

function LocSection({ bikky, isAdmin }) {
  const [locPeriod, setLocPeriod] = useState('weekly');
  const locs = bikky.locations || {};
  const rawRows = locPeriod === 'weekly' ? (locs.weekly?.curr || [])
                : locPeriod === 'monthly' ? (locs.monthly || [])
                : locPeriod === 'ytd' ? (locs.ytd || [])
                : (locs.ninety || []);
  const lbl = locPeriod === 'weekly' ? '7 Days' : locPeriod === 'monthly' ? '30 Days' : locPeriod === 'ytd' ? 'YTD' : '90 Days';
  // YTD locations table only ships from Jun 2026 onward; hide the filter otherwise.
  const locPeriods = (locs.ytd && locs.ytd.length) ? LOC_PERIODS : LOC_PERIODS.filter(p => p.id !== 'ytd');

  const totalFromRows = rawRows.find(r => /^total$/i.test(r.loc));
  let total = totalFromRows;
  if (!total) total = buildLocTotal(rawRows) || { orders: 0, guests: 0, newGuests: 0, aov: 0, ordersGrowth: 0, guestsGrowth: 0, newGuestsGrowth: 0 };

  // YoY growth from prev block (weekly only carries prev in parser)
  const prevTotal = locs.weekly?.prev?.find(r => /^total$/i.test(r.loc));
  if (prevTotal && prevTotal.orders) {
    total = {
      ...total,
      ordersGrowth: total.ordersGrowth ?? ((total.orders - prevTotal.orders) / prevTotal.orders),
      guestsGrowth: total.guestsGrowth ?? ((total.guests - prevTotal.guests) / prevTotal.guests),
      newGuestsGrowth: total.newGuestsGrowth ?? ((total.newGuests - prevTotal.newGuests) / prevTotal.newGuests),
    };
  }

  const dataRows = rawRows.filter(r => !/^total$/i.test(r.loc));
  const tableRows = [...dataRows, total ? { ...total, loc: 'Total' } : null].filter(Boolean);

  const chartData = {
    labels: dataRows.map(r => r.loc),
    datasets: [
      { label: 'Guests',     data: dataRows.map(r => r.guests),    backgroundColor: '#9f7cef', borderRadius: 4 },
      { label: 'New Guests', data: dataRows.map(r => r.newGuests), backgroundColor: '#93c5fd', borderRadius: 4 },
    ],
  };

  return (
    <>
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 12 }}>
        <div className="toggle-group">
          {locPeriods.map(p => (
            <button key={p.id} className={`toggle-btn${locPeriod === p.id ? ' active' : ''}`} onClick={() => setLocPeriod(p.id)}>{p.label}</button>
          ))}
        </div>
      </div>

      <div className="kpi-row">
        <div className="kpi-card">
          <div className="kpi-label">Total Orders</div>
          <div className="kpi-value">{fmtN(total.orders)}</div>
          <div className={`kpi-change ${(total.ordersGrowth || 0) >= 0 ? 'pos' : 'neg'}`}>{fmtVar(total.ordersGrowth)} YoY</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">Total Guests</div>
          <div className="kpi-value">{fmtN(total.guests)}</div>
          <div className={`kpi-change ${(total.guestsGrowth || 0) >= 0 ? 'pos' : 'neg'}`}>{fmtVar(total.guestsGrowth)} YoY</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">New Guests</div>
          <div className="kpi-value">{fmtN(total.newGuests)}</div>
          <div className={`kpi-change ${(total.newGuestsGrowth || 0) >= 0 ? 'pos' : 'neg'}`}>{fmtVar(total.newGuestsGrowth)} YoY</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">Avg Order Value</div>
          <div className="kpi-value">{fmt$2(total.aov)}</div>
        </div>
      </div>

      <div className="table-card" style={{ marginBottom: 16 }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
          <div className="table-title" style={{ marginBottom: 0 }}>By Location — {lbl}</div>
          <ExportCsvButton filename="By Location.csv" />
        </div>
        <Table
          headers={[
            { label: 'Location' },
            { label: 'Orders', cls: 'right' },
            { label: 'Orders Growth (YoY)', cls: 'right' },
            { label: 'AOV', cls: 'right' },
            { label: 'Guests', cls: 'right' },
            { label: 'Guest Growth (YoY)', cls: 'right' },
            { label: 'New Guests', cls: 'right' },
            { label: 'New Guest Growth (YoY)', cls: 'right' },
          ]}
          rows={tableRows.map(r => ({
            _cls: /^total$/i.test(r.loc) ? 'total-row' : '',
            cells: [
              r.loc,
              fmtN(r.orders),
              fmtVar(r.ordersGrowth),
              fmt$2(r.aov),
              fmtN(r.guests),
              growthCell(r.guestsGrowth),
              fmtN(r.newGuests),
              growthCell(r.newGuestsGrowth),
            ],
          }))}
        />
      </div>

      <div className="chart-card">
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
          <div className="chart-title" style={{ marginBottom: 0 }}>Guests by Location — {lbl}</div>
          <ExportImageButton filename="Guests by Location.png" />
        </div>
        <div style={{ height: 280 }}>
          <Bar data={chartData} options={{ responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }} />
        </div>
      </div>
    </>
  );
}

function AcqSection({ bikky, isAdmin }) {
  const [view, setView]     = useState('monthwise');   // monthwise | weekwise
  const [retSel, setRetSel] = useState('90d');         // 90d | 30d

  const acq = bikky.acquisition || {};
  let acqPeriod;
  if (view === 'weekwise') acqPeriod = 'weekly30d';
  else acqPeriod = retSel === '90d' ? 'monthly90d' : 'monthly30d';

  const rows = acqPeriod === 'monthly90d' ? (acq.monthly90d || [])
             : acqPeriod === 'monthly30d' ? (acq.monthly30d || [])
             : (acq.weekly30d || []);

  const isMonthly90 = acqPeriod === 'monthly90d';
  const isMonthly   = acqPeriod === 'monthly90d' || acqPeriod === 'monthly30d';

  const retLabel   = isMonthly90 ? '90d Return Rate' : '30d Return Rate';
  const retKey     = isMonthly90 ? 'returnRate90' : 'returnRate30';
  const spendKey   = isMonthly90 ? 'spend90' : 'spend30';
  const spendLabel = isMonthly90 ? '90d Spend' : '30d Spend';
  const periodLbl  = isMonthly ? 'Latest Month' : 'Latest Week';
  const tableTitle = isMonthly90 ? 'Monthly — 90 Day Return Rate'
                   : acqPeriod === 'monthly30d' ? 'Monthly — 30 Day Return Rate'
                   : 'Weekly — 30 Day Return Rate';
  const chartTitle = isMonthly90 ? 'New Guests & 90d Return Rate' : 'New Guests & 30d Return Rate';

  const latest = rows[0] || {};
  const retRate = latest[retKey] || 0;

  const dataRows = rows.filter(r => !/^average$/i.test(r.period));
  const chartData = {
    labels: dataRows.map(r => r.period),
    datasets: [
      { type: 'bar',  label: 'New Guests', data: dataRows.map(r => r.newGuests), backgroundColor: '#9f7cef', borderRadius: 4, yAxisID: 'y' },
      { type: 'line', label: retLabel,    data: dataRows.map(r => +(((r[retKey]) || 0) * 100).toFixed(1)), borderColor: '#fbbf24', backgroundColor: 'transparent', pointRadius: 3, borderWidth: 2, yAxisID: 'y1' },
    ],
  };

  return (
    <>
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginBottom: 12, flexWrap: 'wrap', alignItems: 'center' }}>
        <select value={view} onChange={e => setView(e.target.value)} style={SELECT_STYLE}>
          <option value="monthwise">Monthly</option>
          <option value="weekwise">Weekly</option>
        </select>
        {view === 'monthwise' && (
          <select value={retSel} onChange={e => setRetSel(e.target.value)} style={SELECT_STYLE}>
            <option value="90d">90 Days Return</option>
            <option value="30d">30 Days Return</option>
          </select>
        )}
      </div>

      {(() => {
        const prevRow = rows[1] || null;
        const vsLabel = isMonthly ? 'vs LM' : 'vs LW';
        return (
          <div className="kpi-row">
            <div className="kpi-card">
              <div className="kpi-label">New Guests ({periodLbl})</div>
              <div className="kpi-value">{fmtN(latest.newGuests)}</div>
              <PrevKpiChip curr={latest.newGuests} prevRow={prevRow} prevKey="newGuests" label={vsLabel} showBoth kind="n" />
            </div>
            <div className="kpi-card">
              <div className="kpi-label">{retLabel} ({periodLbl})</div>
              <div className="kpi-value">{fmtPct(retRate)}</div>
              <PrevKpiChip curr={latest[retKey]} prevRow={prevRow} prevKey={retKey} label={vsLabel} />
            </div>
            <div className="kpi-card">
              <div className="kpi-label">Avg Orders ({periodLbl})</div>
              <div className="kpi-value">{(Number(latest.avgOrders) || 0).toFixed(1)}</div>
              <PrevKpiChip curr={latest.avgOrders} prevRow={prevRow} prevKey="avgOrders" label={vsLabel} />
            </div>
            <div className="kpi-card">
              <div className="kpi-label">AOV ({periodLbl})</div>
              <div className="kpi-value">{fmt$2(latest.aov)}</div>
              <PrevKpiChip curr={latest.aov} prevRow={prevRow} prevKey="aov" label={vsLabel} showBoth kind="$" />
            </div>
          </div>
        );
      })()}

      <div className="table-card" style={{ marginBottom: 16 }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
          <div className="table-title" style={{ marginBottom: 0 }}>{tableTitle}</div>
          <ExportCsvButton filename="Customer Acquisition.csv" />
        </div>
        <Table
          headers={[
            { label: 'Period' },
            { label: 'New Guests', cls: 'right' },
            { label: 'Per Location', cls: 'right' },
            { label: retLabel, cls: 'right' },
            { label: 'Avg Orders', cls: 'right' },
            { label: 'AOV', cls: 'right' },
            { label: spendLabel, cls: 'right' },
          ]}
          rows={rows.map(r => ({
            _cls: /^average$/i.test(r.period) ? 'total-row' : '',
            cells: [
              r.year ? `Week of ${r.period} ${r.year}` : r.period,
              fmtN(r.newGuests),
              fmtN(r.perLoc),
              fmtPct(r[retKey]),
              (Number(r.avgOrders) || 0).toFixed(1),
              fmt$2(r.aov),
              fmt$2(r[spendKey]),
            ],
          }))}
        />
      </div>

      <div className="chart-card">
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
          <div className="chart-title" style={{ marginBottom: 0 }}>{chartTitle}</div>
          <ExportImageButton filename="New Guests and Return Rate.png" />
        </div>
        <div style={{ height: 280 }}>
          <Bar
            data={chartData}
            options={{
              responsive: true,
              maintainAspectRatio: false,
              plugins: { legend: { position: 'bottom' } },
              scales: {
                x: { ticks: { color: '#6b7280' }, grid: { color: '#e5e7eb' } },
                y: { ticks: { color: '#6b7280' }, grid: { color: '#e5e7eb' } },
                y1: { position: 'right', ticks: { color: '#fbbf24', callback: v => v + '%' }, grid: { drawOnChartArea: false } },
              },
            }}
          />
        </div>
      </div>
    </>
  );
}

function OnbSection({ bikky, isAdmin }) {
  const [onbPeriod, setOnbPeriod] = useState('monthly');
  const onb = bikky.onboarding || {};
  const rows = onbPeriod === 'monthly' ? (onb.monthly || []) : (onb.weekly || []);
  const lbl = onbPeriod === 'monthly' ? 'Onboarding — Monthly' : 'Onboarding — Weekly';
  const onbPeriodLbl = onbPeriod === 'monthly' ? 'Latest Month' : 'Latest Week';
  const periodSuffix = onbPeriod === 'monthly' ? 'Monthly' : 'Weekly';

  const latest = rows[0] || {};

  const dataRows = rows.filter(r => !/^average$/i.test(r.period));
  const onboardedChart = {
    labels: dataRows.map(r => r.period),
    datasets: [
      { label: 'Onboarded Guests', data: dataRows.map(r => r.onboarded), backgroundColor: '#9f7cef', borderRadius: 4 },
    ],
  };
  const engagedChart = {
    labels: dataRows.map(r => r.period),
    datasets: [
      { label: 'Engaged Guests', data: dataRows.map(r => r.engaged), backgroundColor: '#93c5fd', borderRadius: 4 },
    ],
  };
  const barOpts = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: { legend: { display: false } },
    scales: {
      x: { ticks: { color: '#6b7280' }, grid: { color: '#e5e7eb' } },
      y: { ticks: { color: '#6b7280' }, grid: { color: '#e5e7eb' } },
    },
  };

  return (
    <>
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 12 }}>
        <div className="toggle-group">
          {ONB_PERIODS.map(p => (
            <button key={p.id} className={`toggle-btn${onbPeriod === p.id ? ' active' : ''}`} onClick={() => setOnbPeriod(p.id)}>{p.label}</button>
          ))}
        </div>
      </div>

      {(() => {
        const prevRow = rows[1] || null;
        const vsLabel = onbPeriod === 'monthly' ? 'vs LM' : 'vs LW';
        return (
          <div className="kpi-row">
            <div className="kpi-card">
              <div className="kpi-label">Onboarded Guests ({onbPeriodLbl})</div>
              <div className="kpi-value">{fmtN(latest.onboarded)}</div>
              <PrevKpiChip curr={latest.onboarded} prevRow={prevRow} prevKey="onboarded" label={vsLabel} />
            </div>
            <div className="kpi-card">
              <div className="kpi-label">Avg Onboarding Latency ({onbPeriodLbl})</div>
              <div className="kpi-value">{fmtN(latest.latency)}</div>
              <PrevKpiChip curr={latest.latency} prevRow={prevRow} prevKey="latency" label={vsLabel} />
            </div>
            <div className="kpi-card">
              <div className="kpi-label">Avg Onboarding Spend ({onbPeriodLbl})</div>
              <div className="kpi-value">{fmt$2(latest.spend)}</div>
              <PrevKpiChip curr={latest.spend} prevRow={prevRow} prevKey="spend" label={vsLabel} />
            </div>
            <div className="kpi-card">
              <div className="kpi-label">Engaged Guests ({onbPeriodLbl})</div>
              <div className="kpi-value">{fmtN(latest.engaged)}</div>
              <PrevKpiChip curr={latest.engaged} prevRow={prevRow} prevKey="engaged" label={vsLabel} />
            </div>
          </div>
        );
      })()}

      <div className="table-card" style={{ marginBottom: 16 }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
          <div className="table-title" style={{ marginBottom: 0 }}>{lbl}</div>
          <ExportCsvButton filename="Onboarding.csv" />
        </div>
        <Table
          headers={[
            { label: 'Period' },
            { label: 'Onboarded', cls: 'right' },
            { label: 'Per Location', cls: 'right' },
            { label: 'Avg Latency (days)', cls: 'right' },
            { label: 'AOV', cls: 'right' },
            { label: 'Avg Spend', cls: 'right' },
            { label: 'Engaged Guests', cls: 'right' },
          ]}
          rows={rows.map(r => ({
            _cls: /^average$/i.test(r.period) ? 'total-row' : '',
            cells: [
              r.period,
              fmtN(r.onboarded),
              fmtN(r.perLoc),
              fmtN(r.latency),
              fmt$2(r.aov),
              fmt$2(r.spend),
              fmtN(r.engaged),
            ],
          }))}
        />
      </div>

      <div className="charts-row">
        <div className="chart-card">
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
            <div className="chart-title" style={{ marginBottom: 0 }}>Total Onboarded Guests ({periodSuffix})</div>
            <ExportImageButton filename="Total Onboarded Guests.png" />
          </div>
          <div style={{ height: 280 }}>
            <Bar data={onboardedChart} options={barOpts} />
          </div>
        </div>
        <div className="chart-card">
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
            <div className="chart-title" style={{ marginBottom: 0 }}>Total Engaged Guests ({periodSuffix})</div>
            <ExportImageButton filename="Total Engaged Guests.png" />
          </div>
          <div style={{ height: 280 }}>
            <Bar data={engagedChart} options={barOpts} />
          </div>
        </div>
      </div>
    </>
  );
}

export default function Bikky({ data, userRole }) {
  const isAdmin = userRole === 'admin';
  const [section, setSection] = useState('loc');
  const bikky = data?.bikky || {};

  return (
    <>
      <div className="toggle-group" style={{ marginBottom: 20 }}>
        {SECTIONS.map(s => (
          <button key={s.id} className={`toggle-btn${section === s.id ? ' active' : ''}`} onClick={() => setSection(s.id)}>{s.label}</button>
        ))}
      </div>

      {section === 'loc' && <LocSection bikky={bikky} isAdmin={isAdmin} />}
      {section === 'acq' && <AcqSection bikky={bikky} isAdmin={isAdmin} />}
      {section === 'onb' && <OnbSection bikky={bikky} isAdmin={isAdmin} />}
    </>
  );
}

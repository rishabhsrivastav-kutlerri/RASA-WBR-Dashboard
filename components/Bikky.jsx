'use client';

import { useState } from 'react';
import '@/lib/chartSetup';
import { Bar } from 'react-chartjs-2';
import Table from './Table';
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
];

const ONB_PERIODS = [
  { id: 'weekly',  label: 'Weekwise' },
  { id: 'monthly', label: 'Monthwise' },
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

function LocSection({ bikky }) {
  const [locPeriod, setLocPeriod] = useState('weekly');
  const locs = bikky.locations || {};
  const rawRows = locPeriod === 'weekly' ? (locs.weekly?.curr || [])
                : locPeriod === 'monthly' ? (locs.monthly || [])
                : (locs.ninety || []);
  const lbl = locPeriod === 'weekly' ? '7 Days' : locPeriod === 'monthly' ? '30 Days' : '90 Days';

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
      { label: 'New Guests', data: dataRows.map(r => r.newGuests), backgroundColor: '#d6c3f8', borderRadius: 4 },
    ],
  };

  return (
    <>
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 12 }}>
        <div className="toggle-group">
          {LOC_PERIODS.map(p => (
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
        <div className="table-title">By Location — {lbl}</div>
        <Table
          headers={[
            { label: 'Location' },
            { label: 'Orders', cls: 'right' },
            { label: 'Orders Growth', cls: 'right' },
            { label: 'AOV', cls: 'right' },
            { label: 'Guests', cls: 'right' },
            { label: 'Guest Growth', cls: 'right' },
            { label: 'New Guests', cls: 'right' },
            { label: 'New Guest Growth', cls: 'right' },
          ]}
          rows={tableRows.map(r => ({
            _cls: /^total$/i.test(r.loc) ? 'total-row' : '',
            cells: [
              r.loc,
              fmtN(r.orders),
              fmtVar(r.ordersGrowth),
              fmt$2(r.aov),
              fmtN(r.guests),
              fmtVar(r.guestsGrowth),
              fmtN(r.newGuests),
              fmtVar(r.newGuestsGrowth),
            ],
          }))}
        />
      </div>

      <div className="chart-card">
        <div className="chart-title">Guests by Location — {lbl}</div>
        <div style={{ height: 280 }}>
          <Bar data={chartData} options={{ responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }} />
        </div>
      </div>
    </>
  );
}

function AcqSection({ bikky }) {
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
          <option value="monthwise">Monthwise</option>
          <option value="weekwise">Weekwise</option>
        </select>
        {view === 'monthwise' && (
          <select value={retSel} onChange={e => setRetSel(e.target.value)} style={SELECT_STYLE}>
            <option value="90d">90 Days Return</option>
            <option value="30d">30 Days Return</option>
          </select>
        )}
      </div>

      <div className="kpi-row">
        <div className="kpi-card">
          <div className="kpi-label">New Guests ({periodLbl})</div>
          <div className="kpi-value">{fmtN(latest.newGuests)}</div>
          <div className="kpi-change neu">{fmtN(latest.perLoc)} per location</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">{retLabel} ({periodLbl})</div>
          <div className="kpi-value">{fmtPct(retRate)}</div>
          <div className={`kpi-change ${retRate >= 0.15 ? 'pos' : retRate >= 0.10 ? 'neu' : 'neg'}`}>Returning guests</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">Avg Orders ({periodLbl})</div>
          <div className="kpi-value">{(Number(latest.avgOrders) || 0).toFixed(1)}</div>
          <div className="kpi-change neu">per new guest</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">AOV ({periodLbl})</div>
          <div className="kpi-value">{fmt$2(latest.aov)}</div>
          <div className="kpi-change neu">per new guest</div>
        </div>
      </div>

      <div className="table-card" style={{ marginBottom: 16 }}>
        <div className="table-title">{tableTitle}</div>
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
              r.period,
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
        <div className="chart-title">{chartTitle}</div>
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

function OnbSection({ bikky }) {
  const [onbPeriod, setOnbPeriod] = useState('monthly');
  const onb = bikky.onboarding || {};
  const rows = onbPeriod === 'monthly' ? (onb.monthly || []) : (onb.weekly || []);
  const lbl = onbPeriod === 'monthly' ? 'Onboarding — Monthly' : 'Onboarding — Weekly';
  const onbPeriodLbl = onbPeriod === 'monthly' ? 'Latest Month' : 'Latest Week';
  const chartTitle = onbPeriod === 'monthly' ? 'Onboarded & Engaged (Monthly)' : 'Onboarded & Engaged (Weekly)';

  const latest = rows[0] || {};

  const dataRows = rows.filter(r => !/^average$/i.test(r.period));
  const chartData = {
    labels: dataRows.map(r => r.period),
    datasets: [
      { type: 'bar',  label: 'Onboarded', data: dataRows.map(r => r.onboarded), backgroundColor: '#9f7cef', borderRadius: 4, yAxisID: 'y' },
      { type: 'line', label: 'Engaged',   data: dataRows.map(r => r.engaged),   borderColor: '#fbbf24', backgroundColor: 'transparent', pointRadius: 3, borderWidth: 2, yAxisID: 'y1' },
    ],
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

      <div className="kpi-row">
        <div className="kpi-card">
          <div className="kpi-label">Onboarded Guests ({onbPeriodLbl})</div>
          <div className="kpi-value">{fmtN(latest.onboarded)}</div>
          <div className="kpi-change neu">{fmtN(latest.perLoc)} per location</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">Avg Onboarding Latency ({onbPeriodLbl})</div>
          <div className="kpi-value">{fmtN(latest.latency)}</div>
          <div className="kpi-change neu">days to 4th order</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">Avg Onboarding Spend ({onbPeriodLbl})</div>
          <div className="kpi-value">{fmt$2(latest.spend)}</div>
          <div className="kpi-change neu">per guest</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">Engaged Guests ({onbPeriodLbl})</div>
          <div className="kpi-value">{fmtN(latest.engaged)}</div>
          <div className="kpi-change pos">Habitual Guests</div>
        </div>
      </div>

      <div className="table-card" style={{ marginBottom: 16 }}>
        <div className="table-title">{lbl}</div>
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

      <div className="chart-card">
        <div className="chart-title">{chartTitle}</div>
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
                y1: { position: 'right', ticks: { color: '#fbbf24' }, grid: { drawOnChartArea: false } },
              },
            }}
          />
        </div>
      </div>
    </>
  );
}

export default function Bikky({ data }) {
  const [section, setSection] = useState('loc');
  const bikky = data?.bikky || {};

  return (
    <>
      <div className="toggle-group" style={{ marginBottom: 20 }}>
        {SECTIONS.map(s => (
          <button key={s.id} className={`toggle-btn${section === s.id ? ' active' : ''}`} onClick={() => setSection(s.id)}>{s.label}</button>
        ))}
      </div>

      {section === 'loc' && <LocSection bikky={bikky} />}
      {section === 'acq' && <AcqSection bikky={bikky} />}
      {section === 'onb' && <OnbSection bikky={bikky} />}
    </>
  );
}

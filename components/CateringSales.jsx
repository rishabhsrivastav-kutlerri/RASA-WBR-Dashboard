'use client';

import { useState } from 'react';
import '@/lib/chartSetup';
import { Line } from 'react-chartjs-2';
import Table from './Table';
import { fmt$ } from '@/lib/fmt';

const VIEWS = [
  { id: 'outbound', label: 'Outbound' },
  { id: 'inbound',  label: 'Inbound' },
];

// Maps the per-cell color tag emitted by the parser to a .badge class.
// 'lpurp' has no equivalent badge style, so those cells render plain.
const BG_CLASS = {
  mint:   'green',
  lgreen: 'green',
  pink:   'red',
  amber:  'amber',
};

const ORDER_HEADERS = [
  { label: 'Cohort' },
  { label: 'Order Status' },
  { label: 'Customer Name' },
  { label: 'Customer Company' },
  { label: 'Email' },
  { label: 'Phone' },
  { label: 'Order Value', cls: 'right' },
];

function statusBadge(status) {
  const cls = status === 'CLOSED' ? 'green' : 'amber';
  return `<span class="badge ${cls}">${status}</span>`;
}

function OrdersTable({ rows }) {
  return (
    <Table
      headers={ORDER_HEADERS}
      rows={rows.map(r => r.isTotal
        ? { _cls: 'total-row', cells: ['Total', '', '', '', '', '', r.value] }
        : {
            cells: [
              r.cohort,
              statusBadge(r.status),
              r.name || r.customer,
              r.company,
              r.email,
              r.phone,
              r.value,
            ],
          })}
    />
  );
}

function MetricsTable({ metrics, cols }) {
  return (
    <table>
      <thead>
        <tr>
          <th>Metrics</th>
          <th className="right">{cols.c1}</th>
          <th className="right">{cols.c2}</th>
          <th className="right">{cols.c3}</th>
          <th className="right">{cols.c4}</th>
          <th className="right">{cols.c5}</th>
          <th className="right">Weekly Plan</th>
          <th className="right">Overall</th>
        </tr>
      </thead>
      <tbody>
        {metrics.map((r, i) => {
          const vals = [r[cols.c1], r[cols.c2], r[cols.c3], r[cols.c4], r[cols.c5], r.plan, r.overall];
          return (
            <tr key={i}>
              <td>{r.highlight ? <strong>{r.metric}</strong> : r.metric}</td>
              {vals.map((v, j) => {
                const tag = r.bgs && r.bgs[j];
                const cls = tag ? BG_CLASS[tag] : null;
                return (
                  <td key={j} className="right">
                    {cls && v != null && v !== '-' && v !== ''
                      ? <span className={`badge ${cls}`}>{v}</span>
                      : v}
                  </td>
                );
              })}
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}

function OutboundTrendGrid({ metrics }) {
  const cols = metrics[0]?._cols;
  if (!cols) return null;
  const labels = [cols.c1, cols.c2, cols.c3, cols.c4, cols.c5];

  return (
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 16, marginBottom: 20 }}>
      {metrics.map((r, i) => {
        const isDollar = /Order Value/i.test(r.metric);
        const isPct    = /^%/.test(r.metric);
        const data = {
          labels,
          datasets: [{
            label: r.metric,
            data: r.raw || [],
            borderColor: '#7c3aed',
            backgroundColor: 'rgba(124,58,237,0.08)',
            borderWidth: 2.5,
            pointBackgroundColor: '#7c3aed',
            pointRadius: 4,
            fill: true,
            tension: 0.3,
          }],
        };
        const opts = {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: (ctx) => {
                  const v = ctx.parsed.y;
                  if (isDollar) return '$' + v.toLocaleString();
                  if (isPct)    return v + '%';
                  return String(v);
                },
              },
            },
          },
          scales: {
            x: { ticks: { font: { size: 10 } }, grid: { display: false } },
            y: {
              ticks: {
                font: { size: 10 },
                callback: (v) => isDollar ? '$' + v.toLocaleString() : isPct ? v + '%' : v,
              },
              grid: { color: 'rgba(0,0,0,0.05)' },
            },
          },
        };
        return (
          <div className="chart-card" key={i}>
            <div className="chart-title" style={{ marginBottom: 8, fontSize: 13 }}>
              5-Week Trend — {r.metric}
            </div>
            <div style={{ height: 160 }}>
              <Line data={data} options={opts} />
            </div>
          </div>
        );
      })}
    </div>
  );
}

function SummaryByDirTable({ rows }) {
  return (
    <Table
      headers={[{ label: 'Category' }, { label: 'Value', cls: 'right' }]}
      rows={rows.map(r => ({
        _cls: r.isTotal ? 'total-row' : '',
        cells: [r.label, r.value],
      }))}
    />
  );
}

function CustomerBreakdownTable({ rows }) {
  const totalRow = rows.length > 1 ? {
    label:   'Total',
    total:   rows.reduce((s, r) => s + (r.total   ?? 0), 0),
    repeat:  rows.reduce((s, r) => s + (r.repeat  ?? 0), 0),
    newCust: rows.reduce((s, r) => s + (r.newCust ?? 0), 0),
  } : null;

  const allRows = totalRow ? [...rows, totalRow] : rows;

  return (
    <Table
      headers={[
        { label: 'Category' },
        { label: 'Total', cls: 'right' },
        { label: 'Repeat', cls: 'right' },
        { label: 'New', cls: 'right' },
      ]}
      rows={allRows.map((r, i) => ({
        _cls: i === allRows.length - 1 && totalRow ? 'total-row' : '',
        cells: [
          r.label,
          r.total   ?? '-',
          r.repeat  ?? '-',
          r.newCust ?? '-',
        ],
      }))}
    />
  );
}

export default function CateringSales({ data }) {
  const [view, setView] = useState('outbound');
  const cs = data.catSales || {};

  const summary          = cs.summary || [];
  const customerBreakdown = cs.customerBreakdown || [];
  const metrics = cs.outboundMetrics || [];
  const cols    = metrics[0]?._cols;

  const obOrders = cs.orders?.outbound || cs.outboundOrders || [];
  const ibOrders = cs.orders?.inbound  || cs.inboundOrders  || [];
  const obSum    = cs.summaryByDir?.outbound || cs.outboundSummary || [];
  const ibSum    = cs.summaryByDir?.inbound  || cs.inboundSummary  || [];

  const trend   = cs.trend   || [];
  const ibTrend = cs.ibTrend || [];
  const tVals   = trend.map(t => t.sales != null ? t.sales : t.val);
  const ibVals  = ibTrend.map(t => t.val);

  function linReg(vals) {
    const n = vals.length;
    if (n < 2) return [];
    const sumX  = vals.reduce((_, __, i) => _ + i, 0);
    const sumY  = vals.reduce((a, v) => a + (v || 0), 0);
    const sumXY = vals.reduce((a, v, i) => a + i * (v || 0), 0);
    const sumX2 = vals.reduce((a, _, i) => a + i * i, 0);
    const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
    const intercept = (sumY - slope * sumX) / n;
    return vals.map((_, i) => Math.round(intercept + slope * i));
  }

  // Linear regression for trendlines
  const trendData   = linReg(tVals);
  const ibTrendData = linReg(ibVals);

  return (
    <>
      <div className="table-card" style={{ marginBottom: 16 }}>
        <div className="table-title">Catering — Summary (Inbound &amp; Outbound)</div>
        <Table
          headers={[
            { label: '#' },
            { label: 'Category' },
            { label: 'Order Value', cls: 'right' },
            { label: 'Closed',      cls: 'right' },
            { label: 'Confirmed',   cls: 'right' },
          ]}
          rows={summary.map(r => ({
            _cls: r.rowStyle || '',
            cells: [r.num, r.label, r.orderVal, r.closed, r.confirmed],
          }))}
        />
      </div>
      {customerBreakdown.length > 0 && (
        <div className="table-card" style={{ marginBottom: 16 }}>
          <div className="table-title">Catering — New vs Repeat Customers</div>
          <CustomerBreakdownTable rows={customerBreakdown} />
        </div>
      )}

      {cols && (
        <div className="table-card" style={{ marginBottom: 20 }}>
          <div className="table-title">Outbound Catering Team — Input &amp; Output Metrics</div>
          <div style={{ fontSize: 11, color: 'var(--muted)', marginBottom: 8 }}>
            Overall: Starts from October 15, 2025
          </div>
          <MetricsTable metrics={metrics} cols={cols} />
        </div>
      )}

      {cols && <OutboundTrendGrid metrics={metrics} />}

      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-end', marginBottom: 12 }}>
        <div className="toggle-group">
          {VIEWS.map(v => (
            <button key={v.id} className={`toggle-btn${view === v.id ? ' active' : ''}`} onClick={() => setView(v.id)}>{v.label}</button>
          ))}
        </div>
      </div>

      {view === 'outbound' && (
        <>
          <div className="table-card" style={{ marginBottom: 16 }}>
            <div className="table-title">Outbound Catering — Last Week Orders</div>
            <OrdersTable rows={obOrders} />
          </div>
          <div className="table-card" style={{ marginBottom: 20 }}>
            <div className="table-title">Catering — Outbound Summary</div>
            <SummaryByDirTable rows={obSum} />
          </div>
        </>
      )}

      {view === 'inbound' && (
        <>
          <div className="table-card" style={{ marginBottom: 16 }}>
            <div className="table-title">Inbound Catering — Last Week Orders</div>
            <OrdersTable rows={ibOrders} />
          </div>
          <div className="table-card" style={{ marginBottom: 16 }}>
            <div className="table-title">Catering — Inbound Summary</div>
            <SummaryByDirTable rows={ibSum} />
          </div>
          {customerBreakdown.some(r => /inbound/i.test(r.label)) && (
            <div className="table-card" style={{ marginBottom: 20 }}>
              <div className="table-title">Inbound — New vs Repeat Customers</div>
              <CustomerBreakdownTable rows={customerBreakdown.filter(r => /inbound/i.test(r.label))} />
            </div>
          )}
        </>
      )}

      {view === 'outbound' && trend.length > 0 && (
        <div className="chart-card">
          <div className="chart-title">WoW — Catering Trend — Outbound</div>
          <Line
            data={{
              labels: trend.map(t => t.week),
              datasets: [
                {
                  label: 'Outbound Order Value',
                  data: tVals,
                  borderColor: '#9f7cef',
                  backgroundColor: 'rgba(159,124,239,0.06)',
                  borderWidth: 2,
                  pointRadius: 3,
                  pointBackgroundColor: '#9f7cef',
                  tension: 0.1,
                  fill: true,
                },
                {
                  label: 'Trend',
                  data: trendData,
                  borderColor: '#9ca3af',
                  backgroundColor: 'transparent',
                  borderWidth: 1.5,
                  pointRadius: 0,
                  tension: 0,
                  fill: false,
                  borderDash: [6, 4],
                },
              ],
            }}
            options={{
              responsive: true,
              plugins: { legend: { position: 'bottom' } },
              scales: {
                x: { ticks: { color: '#6b7280', maxRotation: 45, minRotation: 45, font: { size: 10 } }, grid: { color: '#e5e7eb' } },
                y: { ticks: { color: '#6b7280', callback: v => '$' + v.toLocaleString() }, grid: { color: '#e5e7eb' } },
              },
            }}
          />
        </div>
      )}

      {view === 'inbound' && ibTrend.length > 0 && (
        <div className="chart-card">
          <div className="chart-title">WoW — Catering Trend — Inbound</div>
          <Line
            data={{
              labels: ibTrend.map(t => t.week),
              datasets: [
                {
                  label: 'Inbound Order Value',
                  data: ibVals,
                  borderColor: '#9f7cef',
                  backgroundColor: 'rgba(159,124,239,0.06)',
                  borderWidth: 2,
                  pointRadius: 3,
                  pointBackgroundColor: '#9f7cef',
                  tension: 0.1,
                  fill: true,
                },
                {
                  label: 'Trend',
                  data: ibTrendData,
                  borderColor: '#9ca3af',
                  backgroundColor: 'transparent',
                  borderWidth: 1.5,
                  pointRadius: 0,
                  tension: 0,
                  fill: false,
                  borderDash: [6, 4],
                },
              ],
            }}
            options={{
              responsive: true,
              plugins: { legend: { position: 'bottom' } },
              scales: {
                x: { ticks: { color: '#6b7280', maxRotation: 45, minRotation: 45, font: { size: 10 } }, grid: { color: '#e5e7eb' } },
                y: { ticks: { color: '#6b7280', callback: v => '$' + v.toLocaleString() }, grid: { color: '#e5e7eb' } },
              },
            }}
          />
        </div>
      )}
    </>
  );
}

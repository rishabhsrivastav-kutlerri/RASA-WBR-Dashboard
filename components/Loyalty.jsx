'use client';

import { useState } from 'react';
import '@/lib/chartSetup';
import { Bar, Doughnut, Line } from 'react-chartjs-2';
import Table from './Table';
import { fmt$, fmtN, fmtPct, fmtVar, CHART_COLORS } from '@/lib/fmt';
import { weekInfoForLabel } from '@/lib/fiscalCalendar';

const SECTIONS = [
  { id: 'lifecycle', label: 'Customer Lifecycle' },
  { id: 'sales',     label: 'Loyalty Sales' },
  { id: 'appweb',    label: 'App & Website Breakdown' },
  { id: 'delpick',   label: 'Delivery & Pickup Breakdown' },
];

function VarChip({ curr, prev, kind = '$' }) {
  if (curr == null || prev == null || isNaN(curr) || isNaN(prev) || prev === 0) {
    return <span className="kpi-change neu">— vs LW</span>;
  }
  const diff = curr - prev;
  const pct  = (diff / Math.abs(prev)) * 100;
  const cls  = diff >= 0 ? 'pos' : 'neg';
  const pctTxt = pct >= 0 ? `${pct.toFixed(1)}%` : `(${Math.abs(pct).toFixed(1)}%)`;
  if (kind === '$') {
    const n = Math.abs(Math.round(diff)).toLocaleString('en-US');
    const absTxt = diff >= 0 ? `$${n}` : `($${n})`;
    return <span className={`kpi-change ${cls}`}>Var $: {absTxt}<br/>Var%:{pctTxt} vs LW</span>;
  }
  // orders / counts
  const n = Math.abs(Math.round(diff)).toLocaleString('en-US');
  const absTxt = diff >= 0 ? `${n}` : `(${n})`;
  return <span className={`kpi-change ${cls}`}>Var: {absTxt}<br/>Var%:{pctTxt} vs LW</span>;
}

const PURPLE_PAIR = ['#9f7cef', '#93c5fd'];
const DP_COLORS = ['#9f7cef', '#93c5fd', '#bfdbfe'];

const grpBarOpts = {
  responsive: true,
  maintainAspectRatio: false,
  plugins: { legend: { position: 'bottom' } },
  scales: {
    x: { ticks: { color: '#6b7280', font: { size: 11 } }, grid: { color: 'rgba(0,0,0,0.04)' } },
    y: { ticks: { color: '#6b7280' }, grid: { color: 'rgba(0,0,0,0.04)' } },
  },
};

const doughnutOpts = (kind) => ({
  responsive: true,
  cutout: '55%',
  plugins: {
    legend: { position: 'bottom', labels: { color: '#6b7280', padding: 16, font: { size: 12 } } },
    tooltip: {
      callbacks: {
        label: (ctx) => {
          const total = ctx.dataset.data.reduce((a, b) => a + b, 0) || 1;
          const pct = ((ctx.parsed / total) * 100).toFixed(1);
          const v = kind === '$'
            ? '$' + ctx.parsed.toLocaleString()
            : ctx.parsed.toLocaleString() + ' orders';
          return `${ctx.label}: ${v} (${pct}%)`;
        },
      },
    },
  },
});

function SectionToggle({ section, setSection }) {
  return (
    <div className="toggle-group">
      {SECTIONS.map(s => (
        <button key={s.id}
          className={`toggle-btn${section === s.id ? ' active' : ''}`}
          onClick={() => setSection(s.id)}>{s.label}</button>
      ))}
    </div>
  );
}

export default function Loyalty({ data, prevData }) {
  const [section, setSection]   = useState('lifecycle');
  const [lcPeriod, setLcPeriod] = useState('wow');
  const [awPeriod, setAwPeriod] = useState('weekly');
  const [dpPeriod, setDpPeriod] = useState('weekly');
  const [discPer, setDiscPer]   = useState('weekly');

  const L    = data.loyalty     || {};
  const prevL = prevData?.loyalty || {};

  // Loyalty Signups / App Downloads trendlines are new in the source workbook
  // starting Period 7 Week 1 (Week of June 29) — do not show them before that.
  const weekInfo = weekInfoForLabel(data.label);
  const showTrend = !!weekInfo && weekInfo.period >= 7;

  return (
    <>
      <div className="toggle-group" style={{ marginBottom: 20 }}>
        {SECTIONS.map(s => (
          <button key={s.id}
            className={`toggle-btn${section === s.id ? ' active' : ''}`}
            onClick={() => setSection(s.id)}>{s.label}</button>
        ))}
      </div>

      {section === 'lifecycle' && <Lifecycle L={L} period={lcPeriod} setPeriod={setLcPeriod} showTrend={showTrend} />}
      {section === 'sales'     && <Sales L={L} prevL={prevL} period={discPer} setPeriod={setDiscPer} />}
      {section === 'appweb'    && <AppWeb L={L} period={awPeriod} setPeriod={setAwPeriod} />}
      {section === 'delpick'   && <DelPick L={L} period={dpPeriod} setPeriod={setDpPeriod} />}
    </>
  );
}

// ─── Customer Lifecycle ─────────────────────────────────────────────────────
function Lifecycle({ L, period, setPeriod, showTrend }) {
  const wow = L.lifecycle?.wow || [];
  const mom = L.lifecycle?.mom || [];
  const signupsTrend = L.lifecycleTrend?.signups || { weeks: [], values: [] };
  const appDlTrend   = L.lifecycleTrend?.appDownloads || { weeks: [], values: [] };
  const wowH = L.lifecycle?.wowHeaders || { metric: 'Metric', curr: 'Current', prev: 'Previous', var: 'Var (%)', ytd: 'YTD' };
  const momH = L.lifecycle?.momHeaders || { metric: 'Metric', mar: 'Previous Month', apr: 'Latest Month', var: 'Var (%)' };

  const wowChart = {
    labels: wow.map(r => r.metric),
    datasets: [
      { label: wowH.curr, data: wow.map(r => r.curr), backgroundColor: '#9f7cef', borderRadius: 4 },
      { label: wowH.prev, data: wow.map(r => r.prev), backgroundColor: '#93c5fd', borderRadius: 4 },
    ],
  };
  const lineOpts = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: { legend: { display: false } },
    scales: {
      x: { ticks: { color: '#6b7280', font: { size: 10 } }, grid: { display: false } },
      y: { ticks: { color: '#6b7280' }, grid: { color: 'rgba(0,0,0,0.04)' } },
    },
  };
  const signupsChart = {
    labels: signupsTrend.weeks,
    datasets: [{
      label: 'Loyalty Signups', data: signupsTrend.values,
      borderColor: '#9f7cef', backgroundColor: 'rgba(159,124,239,0.08)',
      borderWidth: 2.5, pointBackgroundColor: '#9f7cef', pointRadius: 3,
      fill: true, tension: 0.3,
    }],
  };
  const appDlChart = {
    labels: appDlTrend.weeks,
    datasets: [{
      label: 'App Downloads', data: appDlTrend.values,
      borderColor: '#93c5fd', backgroundColor: 'rgba(147,197,253,0.08)',
      borderWidth: 2.5, pointBackgroundColor: '#93c5fd', pointRadius: 3,
      fill: true, tension: 0.3,
    }],
  };

  const momData = mom.filter(r => r.metric !== 'Total Members in Loyalty');
  const momChart = {
    labels: momData.map(r => r.metric),
    datasets: [
      { label: momH.mar, data: momData.map(r => r.mar), backgroundColor: '#9f7cef', borderRadius: 4 },
      { label: momH.apr, data: momData.map(r => r.apr), backgroundColor: '#93c5fd', borderRadius: 4 },
    ],
  };

  // KPI card labels mirror the metric column heading in the input sheet
  // (Lifecycle - Table, col A for WoW and col L for MoM).
  const wowKPIs = (
    <div className="kpi-row" style={{ marginBottom: 16 }}>
      {wow.slice(0, 5).map((row, i) => (
        <div key={i} className="kpi-card">
          <div className="kpi-label">{row?.metric || ''}</div>
          <div className="kpi-value">{fmtN(row?.curr)}</div>
          <div className={`kpi-change ${(row?.var || 0) >= 0 ? 'pos' : 'neg'}`}>{fmtPct(row?.var)} WoW</div>
        </div>
      ))}
    </div>
  );

  const momKPIs = (
    <div className="kpi-row" style={{ marginBottom: 16 }}>
      {mom.slice(0, 5).map((row, i) => (
        <div key={i} className="kpi-card">
          <div className="kpi-label">{row?.metric || ''}</div>
          <div className="kpi-value">{fmtN(row?.apr)}</div>
          <div className={`kpi-change ${(row?.var || 0) >= 0 ? 'pos' : 'neg'}`}>{fmtPct(row?.var)} MoM</div>
        </div>
      ))}
    </div>
  );

  return (
    <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-end', marginBottom: 12 }}>
        <div className="toggle-group">
          <button className={`toggle-btn${period === 'wow' ? ' active' : ''}`} onClick={() => setPeriod('wow')}>WoW</button>
          <button className={`toggle-btn${period === 'mom' ? ' active' : ''}`} onClick={() => setPeriod('mom')}>MoM</button>
        </div>
      </div>

      {period === 'wow' ? wowKPIs : momKPIs}

      {period === 'wow' ? (
        <>
          <div className="chart-card" style={{ marginBottom: 16 }}>
            <div className="chart-title">Customer Lifecycle - WoW</div>
            <div style={{ height: 300 }}>
              <Bar data={wowChart} options={grpBarOpts} />
            </div>
          </div>
          <div className="table-card">
            <div className="table-title">Customer Lifecycle - WoW</div>
            <Table
              headers={[
                { label: wowH.metric },
                { label: wowH.curr, cls: 'right' },
                { label: wowH.prev, cls: 'right' },
                { label: wowH.var,  cls: 'right' },
                { label: wowH.ytd,  cls: 'right' },
              ]}
              rows={wow.map(r => ({ cells: [r.metric, fmtN(r.curr), fmtN(r.prev), fmtVar(r.var), fmtN(r.ytd)] }))}
            />
          </div>
        </>
      ) : (
        <>
          <div className="chart-card" style={{ marginBottom: 16 }}>
            <div className="chart-title">Customer Lifecycle - MoM</div>
            <div style={{ height: 300 }}>
              <Bar data={momChart} options={grpBarOpts} />
            </div>
          </div>
          <div className="table-card">
            <div className="table-title">Customer Lifecycle - MoM</div>
            <Table
              headers={[
                { label: momH.metric },
                { label: momH.mar, cls: 'right' },
                { label: momH.apr, cls: 'right' },
                { label: momH.var, cls: 'right' },
              ]}
              rows={mom.map(r => ({ cells: [r.metric, fmtN(r.mar), fmtN(r.apr), fmtVar(r.var)] }))}
            />
          </div>
        </>
      )}

      {showTrend && (signupsTrend.values.length > 0 || appDlTrend.values.length > 0) && (
        <>
          <div className="chart-card" style={{ marginBottom: 16 }}>
            <div className="chart-title">Loyalty Signups — Trend</div>
            <div style={{ height: 220 }}>
              <Line data={signupsChart} options={lineOpts} />
            </div>
          </div>
          <div className="chart-card">
            <div className="chart-title">App Downloads — Trend</div>
            <div style={{ height: 220 }}>
              <Line data={appDlChart} options={lineOpts} />
            </div>
          </div>
        </>
      )}
    </>
  );
}

// ─── Loyalty Sales ──────────────────────────────────────────────────────────
function Sales({ L, prevL, period, setPeriod }) {
  const rows = L.salesByLoc || [];
  const totals = L.salesTotals
    || (() => {
      const gt = rows.find(r => /grand total/i.test(r.loc)) || {};
      return {
        totalSales:   gt.totalSales,
        instoreSales: gt.inStoreSales,
        digitalSales: gt.digitalSales,
        instoreOrders: gt.inStoreOrders,
        digitalOrders: gt.digitalOrders,
        totalOrders:   gt.totalOrders,
      };
    })();

  const prevRows = prevL?.salesByLoc || [];
  const prevTotals = prevL?.salesTotals
    || (() => {
      const gt = prevRows.find(r => /grand total/i.test(r.loc)) || {};
      return {
        totalSales:    gt.totalSales,
        instoreSales:  gt.inStoreSales,
        digitalSales:  gt.digitalSales,
        instoreOrders: gt.inStoreOrders,
        digitalOrders: gt.digitalOrders,
      };
    })();

  const salLocs = rows.filter(r => !/grand total/i.test(r.loc));
  const salesByLocChart = {
    labels: salLocs.map(r => r.loc),
    datasets: [
      { label: 'In-Store Sales', data: salLocs.map(r => r.inStoreSales), backgroundColor: '#9f7cef', borderRadius: 4 },
      { label: 'Digital Sales',  data: salLocs.map(r => r.digitalSales), backgroundColor: '#93c5fd', borderRadius: 4 },
    ],
  };

  const inStore = L.weeklyInStore || [];
  const disc = (period === 'weekly' ? L.weeklyDiscounted : L.twentyEightDayDiscounted) || [];
  const discRows = disc.filter(r => !/grand total/i.test(r.loc));
  const discSalesTotal = discRows.reduce((a, r) => a + (r.discSales || 0), 0);
  const ndSalesTotal   = discRows.reduce((a, r) => a + (r.nonDiscSales || 0), 0);
  const discOrdTotal   = discRows.reduce((a, r) => a + (r.discOrders || 0), 0);
  const ndOrdTotal     = discRows.reduce((a, r) => a + (r.nonDiscOrders || 0), 0);

  const discSalesChart = {
    labels: ['Discounted', 'Non-Discounted'],
    datasets: [{ data: [discSalesTotal, ndSalesTotal], backgroundColor: PURPLE_PAIR, borderWidth: 2, borderColor: '#fff' }],
  };
  const discOrdersChart = {
    labels: ['Discounted', 'Non-Discounted'],
    datasets: [{ data: [discOrdTotal, ndOrdTotal], backgroundColor: PURPLE_PAIR, borderWidth: 2, borderColor: '#fff' }],
  };

  const discCols = [
    { label: 'Location' },
    { label: 'Disc Orders',     cls: 'right' },
    { label: 'Disc Sales',      cls: 'right' },
    { label: 'Non-Disc Orders', cls: 'right' },
    { label: 'Non-Disc Sales',  cls: 'right' },
    { label: 'Total Orders',    cls: 'right' },
    { label: 'Total Sales',     cls: 'right' },
  ];

  return (
    <>
      <div className="kpi-row" style={{ marginBottom: 16 }}>
        <div className="kpi-card">
          <div className="kpi-label">Total Loyalty Sales</div>
          <div className="kpi-value">{fmt$(totals.totalSales)}</div>
          <VarChip curr={totals.totalSales} prev={prevTotals.totalSales} kind="$" />
          {(totals.totalOrders != null || (totals.instoreOrders != null && totals.digitalOrders != null)) && (
            <div className="kpi-change neu">{fmtN(totals.totalOrders ?? ((totals.instoreOrders || 0) + (totals.digitalOrders || 0)))} orders</div>
          )}
        </div>
        <div className="kpi-card">
          <div className="kpi-label">In-Store Sales</div>
          <div className="kpi-value">{fmt$(totals.instoreSales)}</div>
          <VarChip curr={totals.instoreSales} prev={prevTotals.instoreSales} kind="$" />
          {totals.instoreOrders != null && (
            <div className="kpi-change neu">{fmtN(totals.instoreOrders)} orders</div>
          )}
        </div>
        <div className="kpi-card">
          <div className="kpi-label">Digital Sales</div>
          <div className="kpi-value">{fmt$(totals.digitalSales)}</div>
          <VarChip curr={totals.digitalSales} prev={prevTotals.digitalSales} kind="$" />
          {totals.digitalOrders != null && (
            <div className="kpi-change neu">{fmtN(totals.digitalOrders)} orders</div>
          )}
        </div>
      </div>

      <div className="table-card">
        <div className="table-title">Total Loyalty Orders — 7 Days (In-Store / Digital Breakdown)</div>
        <Table
          headers={[
            { label: 'Location' },
            { label: 'In-Store Orders', cls: 'right' },
            { label: 'In-Store Sales',  cls: 'right' },
            { label: 'Digital Orders',  cls: 'right' },
            { label: 'Digital Sales',   cls: 'right' },
            { label: 'Total Orders',    cls: 'right' },
            { label: 'Total Sales',     cls: 'right' },
          ]}
          rows={rows.map(r => ({
            _cls: /grand total/i.test(r.loc) ? 'total-row' : '',
            cells: [r.loc, fmtN(r.inStoreOrders), fmt$(r.inStoreSales), fmtN(r.digitalOrders), fmt$(r.digitalSales), fmtN(r.totalOrders), fmt$(r.totalSales)],
          }))}
        />
      </div>

      <div className="chart-card" style={{ marginBottom: 16 }}>
        <div className="chart-title">Loyalty Sales by Location — 7 Days</div>
        <div style={{ height: 280 }}>
          <Bar data={salesByLocChart} options={{ responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }} />
        </div>
      </div>

      <div className="table-card" style={{ marginBottom: 16 }}>
        <div className="table-title">In-Store Loyalty Sales by Location — 7 Days</div>
        <Table
          headers={[{ label: 'Location' }, { label: 'Orders', cls: 'right' }, { label: 'In-Store Loyalty Sales', cls: 'right' }]}
          rows={inStore.map(r => ({
            _cls: /grand total/i.test(r.loc) ? 'total-row' : '',
            cells: [r.loc, fmtN(r.orders), fmt$(r.sales)],
          }))}
        />
      </div>

      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 12, flexWrap: 'wrap', gap: 10 }}>
        <span style={{ fontSize: 11, fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.06em' }}>Discounted vs Non-Discounted</span>
        <div className="toggle-group">
          <button className={`toggle-btn${period === 'weekly' ? ' active' : ''}`} onClick={() => setPeriod('weekly')}>7 Days</button>
          <button className={`toggle-btn${period === '28d' ? ' active' : ''}`} onClick={() => setPeriod('28d')}>28 Days</button>
        </div>
      </div>

      <div className="table-card">
        <div className="table-title">Discounted (Disc) vs Non-Discounted (Non-Disc) — {period === 'weekly' ? '7 Days' : '28 Days'}</div>
        <Table
          headers={discCols}
          rows={disc.map(r => ({
            _cls: /grand total/i.test(r.loc) ? 'total-row' : '',
            cells: [r.loc, fmtN(r.discOrders), fmt$(r.discSales), fmtN(r.nonDiscOrders), fmt$(r.nonDiscSales), fmtN(r.totalOrders), fmt$(r.totalSales)],
          }))}
        />
      </div>
      <div className="grid2" style={{ marginTop: 16 }}>
        <div className="chart-card">
          <div className="chart-title">{period === 'weekly' ? '7 Days' : '28 Days'} Digital Loyalty Sales — Breakdown by Discounted Sales</div>
          <Doughnut data={discSalesChart} options={doughnutOpts('$')} />
        </div>
        <div className="chart-card">
          <div className="chart-title">{period === 'weekly' ? '7 Days' : '28 Days'} Digital Loyalty Orders — Breakdown by Discounted Orders</div>
          <Doughnut data={discOrdersChart} options={doughnutOpts('n')} />
        </div>
      </div>
    </>
  );
}

// ─── App & Website Breakdown ────────────────────────────────────────────────
function AppWeb({ L, period, setPeriod }) {
  const platRows = (period === 'weekly' ? L.weeklyPlatform : L.twentyEightDayPlatform) || [];
  const awRows   = (period === 'weekly' ? L.weeklyAppWeb : L.twentyEightDayAppWeb) || [];
  const pd = platRows.filter(r => r.platform !== 'Grand Total');
  const awBars = awRows.filter(r => !/grand total/i.test(r.loc));

  const ordersChart = {
    labels: pd.map(r => r.platform),
    datasets: [{ data: pd.map(r => r.orders), backgroundColor: PURPLE_PAIR, borderWidth: 2, borderColor: '#fff' }],
  };
  const salesChart = {
    labels: pd.map(r => r.platform),
    datasets: [{ data: pd.map(r => r.sales), backgroundColor: PURPLE_PAIR, borderWidth: 2, borderColor: '#fff' }],
  };
  const awBarChart = {
    labels: awBars.map(r => r.loc),
    datasets: [
      { label: 'App Sales', data: awBars.map(r => r.appSales), backgroundColor: '#9f7cef', borderRadius: 4 },
      { label: 'Web Sales', data: awBars.map(r => r.webSales), backgroundColor: '#93c5fd', borderRadius: 4 },
    ],
  };
  const lbl = period === 'weekly' ? '7 Days' : '28 Days';

  return (
    <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-end', marginBottom: 12 }}>
        <div className="toggle-group">
          <button className={`toggle-btn${period === 'weekly' ? ' active' : ''}`} onClick={() => setPeriod('weekly')}>7 Days</button>
          <button className={`toggle-btn${period === '28d' ? ' active' : ''}`} onClick={() => setPeriod('28d')}>28 Days</button>
        </div>
      </div>

      <div className="table-card" style={{ marginBottom: 16 }}>
        <div className="table-title">Platform Breakdown — {lbl}</div>
        <Table
          headers={[{ label: 'Platform' }, { label: 'Orders', cls: 'right' }, { label: 'Digital Loyalty Sales', cls: 'right' }]}
          rows={platRows.map(r => ({
            _cls: /grand total/i.test(r.platform) ? 'total-row' : '',
            cells: [r.platform, fmtN(r.orders), fmt$(r.sales)],
          }))}
        />
      </div>

      <div className="charts-row" style={{ marginBottom: 16 }}>
        <div className="chart-card">
          <div className="chart-title">Orders Share — {lbl}</div>
          <Doughnut data={ordersChart} options={doughnutOpts('n')} />
        </div>
        <div className="chart-card">
          <div className="chart-title">Sales Share — {lbl}</div>
          <Doughnut data={salesChart} options={doughnutOpts('$')} />
        </div>
      </div>

      <div className="table-card" style={{ marginBottom: 16 }}>
        <div className="table-title">App vs Web by Location — {lbl}</div>
        <Table
          headers={[
            { label: 'Location' },
            { label: 'App Sales',    cls: 'right' },
            { label: 'App Orders',   cls: 'right' },
            { label: 'Web Sales',    cls: 'right' },
            { label: 'Web Orders',   cls: 'right' },
            { label: 'Total Orders', cls: 'right' },
            { label: 'Total Sales',  cls: 'right' },
          ]}
          rows={awRows.map(r => ({
            _cls: /grand total/i.test(r.loc) ? 'total-row' : '',
            cells: [r.loc, fmt$(r.appSales), fmtN(r.appOrders), fmt$(r.webSales), fmtN(r.webOrders), fmtN(r.totalOrders), fmt$(r.totalSales)],
          }))}
        />
      </div>

      <div className="chart-card">
        <div className="chart-title">App vs Web Sales{period === '28d' ? ' by Location — 28 Days' : ' — 7 Days'}</div>
        <div style={{ height: 280 }}>
          <Bar data={awBarChart} options={{ responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }} />
        </div>
      </div>
    </>
  );
}

// ─── Delivery & Pickup Breakdown ────────────────────────────────────────────
function DelPick({ L, period, setPeriod }) {
  const om = (period === 'weekly' ? L.weeklyOrderMethod : L.twentyEightDayOrderMethod) || [];
  const omRows = om.filter(r => r.method !== 'Grand Total');

  const ordersChart = {
    labels: omRows.map(r => r.method),
    datasets: [{ data: omRows.map(r => r.orders), backgroundColor: DP_COLORS.slice(0, omRows.length), borderWidth: 2, borderColor: '#fff' }],
  };
  const salesChart = {
    labels: omRows.map(r => r.method),
    datasets: [{ data: omRows.map(r => r.sales), backgroundColor: DP_COLORS.slice(0, omRows.length), borderWidth: 2, borderColor: '#fff' }],
  };
  const lbl = period === 'weekly' ? '7 Days' : '28 Days';

  return (
    <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-end', marginBottom: 12 }}>
        <div className="toggle-group">
          <button className={`toggle-btn${period === 'weekly' ? ' active' : ''}`} onClick={() => setPeriod('weekly')}>7 Days</button>
          <button className={`toggle-btn${period === '28d' ? ' active' : ''}`} onClick={() => setPeriod('28d')}>28 Days</button>
        </div>
      </div>

      <div className="table-card" style={{ marginBottom: 16 }}>
        <div className="table-title">Delivery vs Pickup — {lbl}</div>
        <Table
          headers={[{ label: 'Order Method' }, { label: 'Orders', cls: 'right' }, { label: 'Digital Loyalty Sales', cls: 'right' }]}
          rows={om.map(r => ({
            _cls: /grand total/i.test(r.method) ? 'total-row' : '',
            cells: [r.method, fmtN(r.orders), fmt$(r.sales)],
          }))}
        />
      </div>
      <div className="charts-row">
        <div className="chart-card">
          <div className="chart-title">Orders Share — {lbl}</div>
          <Doughnut data={ordersChart} options={doughnutOpts('n')} />
        </div>
        <div className="chart-card">
          <div className="chart-title">Sales Share — {lbl}</div>
          <Doughnut data={salesChart} options={doughnutOpts('$')} />
        </div>
      </div>
    </>
  );
}

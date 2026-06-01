'use client';

import { useState } from 'react';
import Table from './Table';
import { fmt$, fmtN } from '@/lib/fmt';

const PROVIDERS = [
  { id: 'ue', label: 'Uber Eats' },
  { id: 'dd', label: 'DoorDash' },
];

function fmtWait(v) {
  if (!v || v === '-') return '-';
  const parts = String(v).split(':');
  if (parts.length !== 2) return v;
  const m = parseInt(parts[0], 10), s = parseInt(parts[1], 10);
  if (isNaN(m) || isNaN(s)) return v;
  if (m === 0) return s + 'secs';
  if (s === 0) return m + 'mins';
  return m + 'mins ' + s + 'secs';
}
function waitCls(v) {
  if (!v || v === '-') return 'neutral';
  const s = String(v);
  if (s <= '0:59') return 'green';
  if (s <= '2:59') return 'amber';
  return 'red';
}
// All thresholds below mirror the Google-Sheets CF rules read from the source
// XLSX (sheet '3PD Reporting - UE & DD', dxfs 4=red / 5=green / 6=yellow).
function ratingBadge(v) {
  if (!v || v === '-' || typeof v !== 'number') return <span className="badge neutral">-</span>;
  // F4:F8 — UE rating: >=4.5 green, else red (no yellow tier)
  const cls = v >= 4.5 ? 'green' : 'red';
  return <span className={`badge ${cls}`}>{Number(v).toFixed(1)}</span>;
}
function errRateBadge(v) {
  // B15:B19 / O15:O19 — error rate: <1% green, 1-2% yellow, >=2% red
  const n = Number(v) || 0;
  const cls = n < 0.01 ? 'green' : n < 0.02 ? 'amber' : 'red';
  return <span className={`badge ${cls}`}>{(n * 100).toFixed(1)}%</span>;
}
function roasUEBadge(v) {
  // D26:D28 — UE ads ROAS: >=4.5 green, 3-4.5 yellow, <3 red
  const n = Number(v) || 0;
  const cls = n >= 4.5 ? 'green' : n >= 3 ? 'amber' : 'red';
  return <span className={`badge ${cls}`}>{n.toFixed(1)}×</span>;
}
function roasDDPromoBadge(v) {
  // R25:R29 — DD promo ROAS: >=3.5 green, 2-3.5 yellow, <2 red
  const n = Number(v) || 0;
  const cls = n >= 3.5 ? 'green' : n >= 2 ? 'amber' : 'red';
  return <span className={`badge ${cls}`}>{n.toFixed(2)}×</span>;
}
function roasDDSponsorBadge(v) {
  // Y25:Y29 — DD sponsor ROAS: >=4.5 green, 3-4.5 yellow, <3 red
  const n = Number(v) || 0;
  const cls = n >= 4.5 ? 'green' : n >= 3 ? 'amber' : 'red';
  return <span className={`badge ${cls}`}>{n.toFixed(1)}×</span>;
}
function ueCtrBadge(v) {
  // G26:G28 — UE ads CTR: >=5% green, 3-5% yellow, <3% red
  const n = Number(v) || 0;
  const cls = n >= 0.05 ? 'green' : n >= 0.03 ? 'amber' : 'red';
  return <span className={`badge ${cls}`}>{(n * 100).toFixed(1)}%</span>;
}
function ddCtrBadge(v) {
  // U25:U29 — DD CTR: >=4.5% green, 3-4.5% yellow, <3% red
  const n = Number(v) || 0;
  const cls = n >= 0.045 ? 'green' : n >= 0.03 ? 'amber' : 'red';
  return <span className={`badge ${cls}`}>{(n * 100).toFixed(1)}%</span>;
}
function cvrBadge(v) {
  // I26:I28 — UE ads CVR: >=20% green, 12-20% yellow, <12% red
  const n = Number(v) || 0;
  const cls = n >= 0.20 ? 'green' : n >= 0.12 ? 'amber' : 'red';
  return <span className={`badge ${cls}`}>{(n * 100).toFixed(1)}%</span>;
}
function pctBadge(v, thresholds) {
  const n = Number(v) || 0;
  const [hiGreen, hiAmber] = thresholds;
  const cls = n >= hiGreen ? 'green' : n >= hiAmber ? 'amber' : 'red';
  return <span className={`badge ${cls}`}>{(n * 100).toFixed(0)}%</span>;
}
function avoidCancelRateBadge(v) {
  // J15:J19 / W15:W19 — avoidable cancellation rate: =0% green, >0% red
  const n = Number(v) || 0;
  const cls = n === 0 ? 'green' : 'red';
  return <span className={`badge ${cls}`}>{(n * 100).toFixed(1)}%</span>;
}
function uptimeBadge(v) {
  // L15:L19 / X15:X19 — uptime: >=98% green, <98% red
  const n = Number(v) || 0;
  const cls = n >= 0.98 ? 'green' : 'red';
  return <span className={`badge ${cls}`}>{(n * 100).toFixed(0)}%</span>;
}
function complaintCell(v) {
  if (!v || v === '-') return <span className="badge neutral">—</span>;
  return <span style={{ fontSize: 12, color: '#b45309' }}>{v}</span>;
}
function loveBadge(v) {
  // V4:V8 — DD Love %: >=60% green, 40-60% yellow, <40% red
  const n = Number(v) || 0;
  const cls = n >= 0.6 ? 'green' : n >= 0.4 ? 'amber' : 'red';
  return <span className={`badge ${cls}`}>{(n * 100).toFixed(1)}%</span>;
}
function likeBadge(v) {
  // W4:W8 — DD Like %: >=50% green, 30-50% yellow, <30% red
  const n = Number(v) || 0;
  const cls = n >= 0.5 ? 'green' : n >= 0.3 ? 'amber' : 'red';
  return <span className={`badge ${cls}`}>{(n * 100).toFixed(1)}%</span>;
}
function dislikeBadge(v) {
  // X4:X8 — DD Dislike %: <=5.9% green, 6-12% yellow, >=12% red
  const n = Number(v) || 0;
  const cls = n <= 0.059 ? 'green' : n < 0.12 ? 'amber' : 'red';
  return <span className={`badge ${cls}`}>{(n * 100).toFixed(1)}%</span>;
}

function UESection({ ue }) {
  const tUE = ue.perf?.find(r => /all stores/i.test(r.loc)) || {};
  const tUEOps = ue.ops?.find(r => /all stores/i.test(r.loc)) || {};
  return (
    <>
      <div className="kpi-row">
        <div className="kpi-card"><div className="kpi-label">Total Sales</div><div className="kpi-value">{fmt$(tUE.sales)}</div></div>
        <div className="kpi-card"><div className="kpi-label">Number of Orders</div><div className="kpi-value">{fmtN(tUE.orders)}</div></div>
        <div className="kpi-card"><div className="kpi-label">Total Net Payout</div><div className="kpi-value">{fmt$(tUE.payout)}</div></div>
        <div className="kpi-card"><div className="kpi-label">Overall Error Rate</div><div className="kpi-value">{((tUEOps.errRate || 0) * 100).toFixed(1)}%</div></div>
      </div>

      <div className="table-card">
        <div className="table-title">UE Performance by Location</div>
        <Table
          headers={[
            { label: 'Location' },
            { label: 'Sales', cls: 'right' },
            { label: 'Net Payout', cls: 'right' },
            { label: 'Orders', cls: 'right' },
            { label: 'AOV', cls: 'right' },
            { label: 'Rating', cls: 'right' },
            { label: 'Top Complaint' },
          ]}
          rows={(ue.perf || []).map(r => ({
            _cls: /all stores/i.test(r.loc) ? 'total-row' : '',
            cells: [r.loc, fmt$(r.sales), fmt$(r.payout), fmtN(r.orders), fmt$(r.aov), ratingBadge(r.rating), complaintCell(r.complaint)],
          }))}
        />
      </div>

      <div className="table-card">
        <div className="table-title">UE Operations by Location</div>
        <Table
          headers={[
            { label: 'Location' },
            { label: 'Error Rate', cls: 'right' },
            { label: 'Missing Items', cls: 'right' },
            { label: 'Customization Error', cls: 'right' },
            { label: 'Wrong Order', cls: 'right' },
            { label: 'Quality Issues', cls: 'right' },
            { label: 'Avg Courier Wait', cls: 'right' },
            { label: 'Avoidable Wait', cls: 'right' },
            { label: 'Orders w/ Avoid Wait', cls: 'right' },
            { label: 'Avoid Cancel Rate', cls: 'right' },
            { label: 'Menu CVR', cls: 'right' },
            { label: 'Uptime', cls: 'right' },
          ]}
          rows={(ue.ops || []).map(r => ({
            _cls: /all stores/i.test(r.loc) ? 'total-row' : '',
            cells: [
              r.loc,
              errRateBadge(r.errRate),
              fmtN(r.missItems), fmtN(r.custErr), fmtN(r.wrongOrder), fmtN(r.qualIssues),
              fmtWait(r.avgWait),
              <span key="avw" className={`badge ${waitCls(r.avoidWait)}`}>{fmtWait(r.avoidWait)}</span>,
              fmtN(r.avoidOrders),
              avoidCancelRateBadge(r.avoidCancelRate),
              pctBadge(r.menuCvr, [0.21, 0.12]),
              uptimeBadge(r.uptime),
            ],
          }))}
        />
      </div>

      <div className="table-card">
        <div className="table-title">UE Ads &amp; Campaigns</div>
        <Table
          headers={[
            { label: 'Campaign' },
            { label: 'Sales', cls: 'right' },
            { label: 'Spend', cls: 'right' },
            { label: 'ROAS', cls: 'right' },
            { label: 'Impressions', cls: 'right' },
            { label: 'Clicks', cls: 'right' },
            { label: 'CTR', cls: 'right' },
            { label: 'Orders', cls: 'right' },
            { label: 'CVR', cls: 'right' },
            { label: 'Cost Per Order', cls: 'right' },
            { label: 'New Customers', cls: 'right' },
          ]}
          rows={(ue.ads || []).map(r => ({
            _cls: /all campaign/i.test(r.campaign) ? 'total-row' : '',
            cells: [
              r.campaign,
              fmt$(r.sales), fmt$(r.spend),
              roasUEBadge(r.roas),
              fmtN(r.impressions), fmtN(r.clicks),
              ueCtrBadge(r.ctr),
              fmtN(r.orders),
              cvrBadge(r.cvr),
              '$' + (Number(r.cpo) || 0).toFixed(2),
              fmtN(r.newCust),
            ],
          }))}
        />
      </div>
    </>
  );
}

function DDAdsTable({ rows }) {
  const divL = { borderLeft: '2px solid var(--border)' };
  const divR = { borderRight: '2px solid var(--border)' };
  return (
    <table>
      <thead>
        <tr>
          <th rowSpan={2}>Location</th>
          <th colSpan={4} style={{ textAlign: 'center', ...divL, ...divR }}>Promotions</th>
          <th colSpan={7} style={{ textAlign: 'center', ...divR }}>Sponsored</th>
          <th rowSpan={2} className="right">Overall Orders</th>
        </tr>
        <tr>
          <th className="right" style={divL}>Promo Sales</th>
          <th className="right">Promo Spend</th>
          <th className="right">Orders from Promo</th>
          <th className="right" style={divR}>Promo ROAS</th>
          <th className="right">Impressions</th>
          <th className="right">Clicks</th>
          <th className="right">CTR</th>
          <th className="right">Sponsor Sales</th>
          <th className="right">Sponsor Spend</th>
          <th className="right">Orders from Sponsor</th>
          <th className="right" style={divR}>Sponsor ROAS</th>
        </tr>
      </thead>
      <tbody>
        {rows.map((r, i) => (
          <tr key={i} className={/all stores/i.test(r.loc) ? 'total-row' : ''}>
            <td>{r.loc}</td>
            <td className="right" style={divL}>{fmt$(r.promoSales)}</td>
            <td className="right">{fmt$(r.promoSpend)}</td>
            <td className="right">{fmtN(r.promoOrders)}</td>
            <td className="right" style={divR}>{roasDDPromoBadge(r.promoROAS)}</td>
            <td className="right">{fmtN(r.impressions)}</td>
            <td className="right">{fmtN(r.clicks)}</td>
            <td className="right">{ddCtrBadge(r.ctr)}</td>
            <td className="right">{fmt$(r.sponsorSales)}</td>
            <td className="right">{fmt$(r.sponsorSpend)}</td>
            <td className="right">{fmtN(r.sponsorOrders)}</td>
            <td className="right" style={divR}>{roasDDSponsorBadge(r.sponsorROAS)}</td>
            <td className="right">{fmtN(r.overallOrders)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function DDSection({ dd }) {
  const tDD = dd.perf?.find(r => /all stores/i.test(r.loc)) || {};
  const tDDOps = dd.ops?.find(r => /all stores/i.test(r.loc)) || {};
  return (
    <>
      <div className="kpi-row">
        <div className="kpi-card"><div className="kpi-label">Total Sales</div><div className="kpi-value">{fmt$(tDD.sales)}</div></div>
        <div className="kpi-card"><div className="kpi-label">Number of Orders</div><div className="kpi-value">{fmtN(tDD.orders)}</div></div>
        <div className="kpi-card"><div className="kpi-label">Total Net Payout</div><div className="kpi-value">{fmt$(tDD.payout)}</div></div>
        <div className="kpi-card"><div className="kpi-label">Overall Error Rate</div><div className="kpi-value">{((tDDOps.errRate || 0) * 100).toFixed(1)}%</div></div>
      </div>

      <div className="table-card">
        <div className="table-title">DD Performance by Location</div>
        <Table
          headers={[
            { label: 'Location' },
            { label: 'Sales', cls: 'right' },
            { label: 'Net Payout', cls: 'right' },
            { label: 'Orders', cls: 'right' },
            { label: 'AOV', cls: 'right' },
            { label: 'Top Complaints' },
          ]}
          rows={(dd.perf || []).map(r => ({
            _cls: /all stores/i.test(r.loc) ? 'total-row' : '',
            cells: [r.loc, fmt$(r.sales), fmt$(r.payout), fmtN(r.orders), fmt$(r.aov), complaintCell(r.complaint)],
          }))}
        />
      </div>

      <div className="table-card">
        <div className="table-title">DD Rating Distribution by Location</div>
        <Table
          headers={[
            { label: 'Location' },
            { label: 'Love %', cls: 'right' },
            { label: 'Like %', cls: 'right' },
            { label: 'Dislike %', cls: 'right' },
            { label: 'Total Reviews', cls: 'right' },
          ]}
          rows={(dd.ratings || []).map(r => ({
            _cls: /all stores/i.test(r.loc) ? 'total-row' : '',
            cells: [
              r.loc,
              loveBadge(r.lovePct ?? r.love),
              likeBadge(r.likePct ?? r.like),
              dislikeBadge(r.dislikePct ?? r.dislike),
              fmtN(r.totalReviews ?? r.reviews),
            ],
          }))}
        />
      </div>

      <div className="table-card">
        <div className="table-title">DD Operations by Location</div>
        <Table
          headers={[
            { label: 'Location' },
            { label: 'Error Rate', cls: 'right' },
            { label: 'Missing Items', cls: 'right' },
            { label: 'Ingredient Error', cls: 'right' },
            { label: 'Missing Side', cls: 'right' },
            { label: 'Incorrect Size', cls: 'right' },
            { label: 'Avg Dasher Wait', cls: 'right' },
            { label: 'Avoidable Wait', cls: 'right' },
            { label: 'Avoid Cancels', cls: 'right' },
            { label: 'Avoid Cancel Rate', cls: 'right' },
            { label: 'Uptime', cls: 'right' },
          ]}
          rows={(dd.ops || []).map(r => ({
            _cls: /all stores/i.test(r.loc) ? 'total-row' : '',
            cells: [
              r.loc,
              errRateBadge(r.errRate),
              fmtN(r.missItems), fmtN(r.ingErr), fmtN(r.missSide), fmtN(r.incSize),
              fmtWait(r.dashWait),
              <span key="avw" className={`badge ${waitCls(r.avoidWait)}`}>{fmtWait(r.avoidWait)}</span>,
              fmtN(r.avoidCancel),
              avoidCancelRateBadge(r.avoidCancelRate),
              <span key="up" className={`badge ${(Number(r.uptime) || 0) >= 0.98 ? 'green' : 'red'}`}>{((Number(r.uptime) || 0) * 100).toFixed(1)}%</span>,
            ],
          }))}
        />
      </div>

      <div className="table-card">
        <div className="table-title">DD Promotions &amp; Sponsored Ads</div>
        <DDAdsTable rows={dd.ads || []} />
      </div>
    </>
  );
}

export default function ThirdParty({ data }) {
  const [prov, setProv] = useState('ue');
  const ue = data?.ue || {};
  const dd = data?.dd || {};

  return (
    <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
        <span style={{ fontSize: 11, fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.06em' }}>3rd Party Delivery</span>
        <div className="toggle-group">
          {PROVIDERS.map(p => (
            <button key={p.id} className={`toggle-btn${prov === p.id ? ' active' : ''}`} onClick={() => setProv(p.id)}>{p.label}</button>
          ))}
        </div>
      </div>

      {prov === 'ue' ? <UESection ue={ue} /> : <DDSection dd={dd} />}
    </>
  );
}

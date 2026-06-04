'use client';

import { useState } from 'react';
import Table from './Table';
import { fmt$, fmtN, fmtPct, fmtVar } from '@/lib/fmt';

const SUB = [
  { id: 'catering', label: 'Catering Marketing' },
  { id: 'loyalty',  label: 'Loyalty Marketing' },
];
const PERIODS = [
  { id: '30d', label: '30 Days' },
  { id: '90d', label: '90 Days' },
];

const EMAIL_HEADERS = [
  { label: 'Campaign' },
  { label: 'Sent',      cls: 'right' },
  { label: 'Delivered', cls: 'right' },
  { label: 'Bounced',   cls: 'right' },
  { label: 'Spam',      cls: 'right' },
  { label: 'Unsub',     cls: 'right' },
  { label: 'Opened',    cls: 'right' },
  { label: 'Clicked',   cls: 'right' },
  { label: 'Ordered',   cls: 'right' },
  { label: 'Revenue',   cls: 'right' },
];

function openedCell(opened, delivered) {
  const pct = delivered > 0 ? Math.round((opened / delivered) * 100) : 0;
  return `${fmtN(opened)} <span style="color:#9ca3af;font-size:11px">(${pct}%)</span>`;
}

function EmailTable({ rows }) {
  return (
    <Table
      headers={EMAIL_HEADERS}
      rows={rows.map(r => ({
        _cls: /^total$/i.test(r.campaign) ? 'total-row' : '',
        cells: [
          r.campaign,
          fmtN(r.sent),
          fmtN(r.delivered),
          fmtN(r.bounced),
          fmtN(r.spam),
          fmtN(r.unsub),
          openedCell(r.opened, r.delivered),
          fmtN(r.clicked),
          r.ordered != null ? fmtN(r.ordered) : '-',
          fmt$(r.revenue),
        ],
      }))}
    />
  );
}

function FlowTable({ rows }) {
  return (
    <Table
      headers={[
        { label: 'Flow' },
        { label: 'Delivered', cls: 'right' },
        { label: 'Opened',    cls: 'right' },
        { label: 'Clicked',   cls: 'right' },
        { label: 'Revenue',   cls: 'right' },
      ]}
      rows={rows.map(r => ({
        _cls: /^total$/i.test(r.flow) ? 'total-row' : '',
        cells: [
          r.flow,
          fmtN(r.delivered),
          openedCell(r.opened, r.delivered),
          fmtN(r.clicked),
          fmt$(r.revenue),
        ],
      }))}
    />
  );
}

// EzCater conditional-formatting rules pulled from the source XLSX
// (styles.xml dxfId 4/5/6 on T3:T8 and X3:X8 of InputsOutputs Catering).
// Rendered through the shared .badge class so the chip style matches every
// other CF table in the dashboard.
function cvrCell(v) {
  const n = Number(v) || 0;
  const cls = n > 0.20 ? 'green' : n >= 0.15 ? 'amber' : 'red';
  return `<span class="badge ${cls}">${(n * 100).toFixed(1)}%</span>`;
}

function roasCell(v) {
  const n = Number(v) || 0;
  const cls = n >= 4 ? 'green' : n > 3.5 ? 'amber' : 'red';
  return `<span class="badge ${cls}">${n.toFixed(1)}×</span>`;
}

// ─── Catering Marketing ─────────────────────────────────────────────────────
function CateringMarketing({ data, sub, setSub, period, setPeriod }) {
  const c = data.catering || data.marketing?.catering || {};
  const is30 = period === '30d';
  const emails = is30 ? (c.email30d || []) : (c.email90d || []);
  // Parser exposes `flows30d/flows90d`; task spec used `flow30d/flow90d` — accept both.
  const flows  = is30
    ? (c.flow30d || c.flows30d || [])
    : (c.flow90d || c.flows90d || []);
  const ezAds = c.ezcaterAds || [];

  const tEmail = emails.find(r => /^total$/i.test(r.campaign)) || {};
  const tFlow  = flows.find(r => /^total$/i.test(r.flow)) || {};

  const emailOpenPct = tEmail.delivered > 0
    ? ((tEmail.opened / tEmail.delivered) * 100).toFixed(1) + '%' : '-';
  const flowOpenPct  = tFlow.delivered > 0
    ? ((tFlow.opened / tFlow.delivered) * 100).toFixed(1) + '%' : '-';

  const ezRows = ezAds.filter(r => !/^total$/i.test(r.loc || ''));
  // ROAS shown in the KPI card mirrors the EzCater input table's Total row exactly.
  const ezTotal = ezAds.find(r => r.isTotal || /^total$/i.test(r.loc || ''));
  const totalRoas = Number(ezTotal?.roas) || 0;

  const lbl = is30 ? '30 Days' : '90 Days';

  return (
    <>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 20 }}>
        <div className="toggle-group">
          {SUB.map(s => (
            <button key={s.id} className={`toggle-btn${sub === s.id ? ' active' : ''}`} onClick={() => setSub(s.id)}>{s.label}</button>
          ))}
        </div>
      </div>

      <div className="kpi-row" style={{ marginBottom: 16 }}>
        <div className="kpi-card">
          <div className="kpi-label">Emails</div>
          <div className="kpi-value">{fmt$(tEmail.revenue)}</div>
          <div className="kpi-change neu">
            Sent: {fmtN(tEmail.sent)} · Delivered: {fmtN(tEmail.delivered)} · Open: {emailOpenPct}
          </div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">Flows</div>
          <div className="kpi-value">{fmt$(tFlow.revenue)}</div>
          <div className="kpi-change neu">
            Delivered: {fmtN(tFlow.delivered)} · Open: {flowOpenPct}
          </div>
        </div>
        {is30 && (
          <div className="kpi-card">
            <div className="kpi-label">EzCater Ad ROAS (30d)</div>
            <div className="kpi-value">{totalRoas.toFixed(1)}×</div>
            <div className="kpi-change pos">{ezRows.length} Locations</div>
          </div>
        )}
      </div>

      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
        <span style={{ fontSize: 11, fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.06em' }}>Email & Flows</span>
        <div className="toggle-group">
          {PERIODS.map(p => (
            <button key={p.id} className={`toggle-btn${period === p.id ? ' active' : ''}`} onClick={() => setPeriod(p.id)}>{p.label}</button>
          ))}
        </div>
      </div>

      <div className="table-card">
        <div className="table-title">Email Campaigns — Last {lbl} (Klaviyo)</div>
        <EmailTable rows={emails} />
      </div>

      <div className="table-card">
        <div className="table-title">Automated Flows — Last {lbl}</div>
        <FlowTable rows={flows} />
      </div>

      {is30 && (
        <div className="table-card">
          <div className="table-title">EzCater Paid Ads — Last 30 Days</div>
          <Table
            headers={[
              { label: 'Restaurant' },
              { label: 'Views',                cls: 'right' },
              { label: 'Clicks',               cls: 'right' },
              { label: 'Conversion Rate',      cls: 'right' },
              { label: 'Orders',               cls: 'right' },
              { label: 'Ad Spend',             cls: 'right' },
              { label: 'Sales from Ads',       cls: 'right' },
              { label: 'ROAS',                 cls: 'right' },
              { label: 'Customers (New)',      cls: 'right' },
              { label: 'Customers (Existing)', cls: 'right' },
              { label: 'Customers (Lapsed)',   cls: 'right' },
            ]}
            rows={ezAds.map(r => {
              const isTotal = r.isTotal || /^total$/i.test(r.loc || '');
              return {
                _cls: isTotal ? 'total-row' : '',
                cells: [
                  r.loc,
                  fmtN(r.views),
                  fmtN(r.clicks),
                  isTotal ? fmtPct(r.cvr) : cvrCell(r.cvr),
                  fmtN(r.orders),
                  fmt$(r.adSpend != null ? r.adSpend : r.spend),
                  fmt$(r.sales),
                  isTotal ? (Number(r.roas) || 0).toFixed(1) + '×' : roasCell(r.roas),
                  fmtN(r.custNew),
                  fmtN(r.custExisting),
                  fmtN(r.custLapsed),
                ],
              };
            })}
          />
        </div>
      )}
    </>
  );
}

// ─── Loyalty Marketing ──────────────────────────────────────────────────────
function LoyaltyMarketing({ data, sub, setSub }) {
  const lm = data.loyaltyMarketing || data.marketing?.loyalty || {};
  const sms = lm.smsWoW || [];
  const e7  = lm.email7d || [];
  const e30 = lm.email30d || [];

  const cols = lm.smsCols || {};
  const isMoney = m => /Sales|Revenue|Costs|Value/i.test(m || '');
  const findSms = name => sms.find(r => r.metric === name) || {};
  const campaignsSent = findSms('Campaigns Sent');
  const attrSales     = findSms('Attributed Sales');
  const smsRoas       = findSms('SMS ROAS');
  const total30       = e30.find(r => /^total$/i.test(r.campaign)) || {};

  return (
    <>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 20 }}>
        <div className="toggle-group">
          {SUB.map(s => (
            <button key={s.id} className={`toggle-btn${sub === s.id ? ' active' : ''}`} onClick={() => setSub(s.id)}>{s.label}</button>
          ))}
        </div>
      </div>

      <div className="kpi-row" style={{ marginBottom: 16 }}>
        <div className="kpi-card">
          <div className="kpi-label">SMS/Push Campaigns (Week)</div>
          <div className="kpi-value">{fmtN(campaignsSent.curr)}</div>
          <div className={`kpi-change ${campaignsSent.var >= 0 ? 'pos' : 'neg'}`}>{fmtPct(campaignsSent.var)} WoW</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">SMS Attributed Sales</div>
          <div className="kpi-value">{fmt$(attrSales.curr)}</div>
          <div className={`kpi-change ${attrSales.var >= 0 ? 'pos' : 'neg'}`}>{fmtPct(attrSales.var)} WoW</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">SMS ROAS</div>
          <div className="kpi-value">{smsRoas.curr != null ? `${smsRoas.curr}×` : '-'}</div>
          <div className={`kpi-change ${smsRoas.var >= 0 ? 'pos' : 'neg'}`}>{fmtPct(smsRoas.var)} WoW</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">Email Revenue (30d)</div>
          <div className="kpi-value">{fmt$(total30.revenue)}</div>
          <div className="kpi-change neu">{total30.sent != null ? `${fmtN(total30.sent)} sent` : ''}</div>
        </div>
      </div>

      <div className="table-card">
        <div className="table-title">SMS + Push Campaigns — WoW Metrics</div>
        <Table
          headers={[
            { label: 'Metric' },
            { label: cols.curr || 'This Week',  cls: 'right' },
            { label: cols.prev || 'Prior Week', cls: 'right' },
            { label: 'Var (%)',                 cls: 'right' },
          ]}
          rows={sms.map(r => ({
            cells: [
              r.metric,
              isMoney(r.metric) ? fmt$(r.curr) : fmtN(r.curr),
              isMoney(r.metric) ? fmt$(r.prev) : fmtN(r.prev),
              fmtVar(r.var),
            ],
          }))}
        />
      </div>

      <div className="table-card">
        <div className="table-title">Email Campaigns — Last 7 Days (Klaviyo)</div>
        <EmailTable rows={e7} />
      </div>

      <div className="table-card">
        <div className="table-title">Email Campaigns — Last 30 Days (Klaviyo)</div>
        <EmailTable rows={e30} />
      </div>
    </>
  );
}

export default function Marketing({ data }) {
  const [sub, setSub]       = useState('catering');
  const [period, setPeriod] = useState('30d');

  if (sub === 'catering') {
    return <CateringMarketing data={data} sub={sub} setSub={setSub} period={period} setPeriod={setPeriod} />;
  }
  return <LoyaltyMarketing data={data} sub={sub} setSub={setSub} />;
}

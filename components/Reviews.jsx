'use client';

import { useState, useMemo } from 'react';
import '@/lib/chartSetup';
import { Bar } from 'react-chartjs-2';
import Table from './Table';
import { ExportCsvButton, ExportImageButton } from './ExportButtons';
import { fmtN } from '@/lib/fmt';

const SOURCES = [
  { id: 'instore',    label: 'Google + Yelp' },
  { id: 'thirdparty', label: '3rd Party (UE / DD / GH)' },
];
const PERIODS = [
  { id: 'weekly', label: '7 Days' },
  { id: 'day30',  label: '30 Days' },
  { id: 'ninety', label: '90 Days' },
];

// Non-location rows to drop from the 3rd Party reviews tables (both periods).
const EXCLUDED_3PD = ['Trey', 'Rahul', 'Sahil', 'Didier', 'Olga'];

// Weighted-average platform ratings across all location rows.
function platformAvgs(locRows, source) {
  const n = v => typeof v === 'number' && !isNaN(v) ? v : 0;
  if (source === 'instore') {
    const gRows = locRows.filter(r => typeof r.google === 'number' && !isNaN(r.google));
    const yRows = locRows.filter(r => typeof r.yelp   === 'number' && !isNaN(r.yelp));
    const gTotal = gRows.reduce((a, r) => a + n(r.gNum),  0);
    const yTotal = yRows.reduce((a, r) => a + n(r.yelpN), 0);
    const google = gTotal > 0
      ? gRows.reduce((a, r) => a + r.google * n(r.gNum),  0) / gTotal
      : gRows.length ? gRows.reduce((a, r) => a + r.google, 0) / gRows.length : null;
    const yelp = yTotal > 0
      ? yRows.reduce((a, r) => a + r.yelp   * n(r.yelpN), 0) / yTotal
      : yRows.length ? yRows.reduce((a, r) => a + r.yelp, 0) / yRows.length : null;
    return { google, yelp };
  }
  const avg = (rows, key) => {
    const valid = rows.filter(r => typeof r[key] === 'number' && !isNaN(r[key]));
    return valid.length ? valid.reduce((a, r) => a + r[key], 0) / valid.length : null;
  };
  return { ue: avg(locRows, 'ue'), dd: avg(locRows, 'dd'), gh: avg(locRows, 'gh') };
}

// Variance chip. `inverse` = increases are bad (error rate).
// `isRating` = show absolute delta instead of %.
// `showBoth` = show absolute count change AND percentage (for Reviews / 5★).
function VarChip({ curr, prev, inverse = false, isRating = false, showBoth = false, labeled = false, isPct = false }) {
  if (curr == null || prev == null || isNaN(curr) || isNaN(prev)) {
    return <span className="kpi-change neu">— vs LW</span>;
  }
  const diff = curr - prev;
  if (isPct) {
    const pp = diff * 100;
    const cls = diff === 0 ? 'neu' : (inverse ? diff < 0 : diff > 0) ? 'pos' : 'neg';
    const txt = pp >= 0 ? `${pp.toFixed(2)}%` : `(${Math.abs(pp).toFixed(2)}%)`;
    return <span className={`kpi-change ${cls}`}>{labeled ? 'Var%: ' : ''}{txt} vs LW</span>;
  }
  if (isRating) {
    const cls = diff === 0 ? 'neu' : (inverse ? diff < 0 : diff > 0) ? 'pos' : 'neg';
    const txt = diff >= 0 ? `${diff.toFixed(2)}` : `(${Math.abs(diff).toFixed(2)})`;
    return <span className={`kpi-change ${cls}`}>{labeled ? 'Var: ' : ''}{txt} vs LW</span>;
  }
  if (prev === 0) {
    return <span className={`kpi-change ${diff > 0 ? 'pos' : 'neu'}`}>{diff > 0 ? 'New' : '—'} vs LW</span>;
  }
  const pct = (diff / Math.abs(prev)) * 100;
  const cls = diff === 0 ? 'neu' : (inverse ? diff < 0 : diff > 0) ? 'pos' : 'neg';
  if (showBoth) {
    const abs = Math.round(Math.abs(diff)).toLocaleString('en-US');
    const absTxt = diff >= 0 ? `${abs}` : `(${abs})`;
    const pctTxt = pct >= 0 ? `${pct.toFixed(1)}%` : `(${Math.abs(pct).toFixed(1)}%)`;
    return <span className={`kpi-change ${cls}`}>{labeled ? 'Var: ' : ''}{absTxt}<br/>{labeled ? 'Var%: ' : ''}{pctTxt} vs LW</span>;
  }
  const txt = pct >= 0 ? `${pct.toFixed(1)}%` : `(${Math.abs(pct).toFixed(1)}%)`;
  return <span className={`kpi-change ${cls}`}>{labeled ? 'Var%: ' : ''}{txt} vs LW</span>;
}

const STAR_LOCS = [
  { value: 'all', label: 'All Locations' },
  { value: 'Ballpark', label: 'Ballpark' },
  { value: 'MVT', label: 'MVT' },
  { value: 'NL', label: 'NL' },
  { value: 'Mosaic', label: 'Mosaic' },
  { value: 'Rockville', label: 'Rockville' },
];

// Thresholds mirror the conditional-formatting rules on the rating cells in the
// source XLSX (sheet 'Weekly in-store leadership metr', ranges D/J/L/R/X/Z):
//   >=4.7 → dxf7 green, 4.51-4.69 → dxf6 yellow, <=4.5 → dxf4 red.
// Applied to the Rating column as well as the per-platform rating columns.
function ratingBadge(v) {
  if (v == null || v === '-' || typeof v !== 'number' || isNaN(v)) {
    return <span className="badge neutral">NA</span>;
  }
  // Rating cells: >=4.7 green, 4.51-4.69 yellow, <=4.5 red
  const cls = v >= 4.7 ? 'green' : v > 4.5 ? 'amber' : 'red';
  return <span className={`badge ${cls}`}>{Number(v).toFixed(1)}</span>;
}
function errRateBadge(v) {
  // M10:M15 / AA10:AA15 — 3rd Party error rate:
  // <=1.54% green, 1.55-2.19% yellow, >=2.19% red
  if (v == null || (typeof v === 'string' && /^n\/?a$/i.test(v.trim()))) {
    return <span className="badge neutral">NA</span>;
  }
  const n = typeof v === 'number' ? v : Number(v);
  if (isNaN(n)) return <span className="badge neutral">NA</span>;
  const cls = n <= 0.0154 ? 'green' : n < 0.0219 ? 'amber' : 'red';
  return <span className={`badge ${cls}`}>{(n * 100).toFixed(1)}%</span>;
}

function computeTotal(rows, source) {
  const list = rows.filter(r => !/total/i.test(r.loc));
  if (!list.length) return null;
  const reviews = list.reduce((a, r) => a + (r.reviews || 0), 0);
  const s5 = list.reduce((a, r) => a + (typeof r.s5 === 'number' && !isNaN(r.s5) ? r.s5 : 0), 0);
  // Only include rows with a real numeric rating in the weighted average.
  const rRows   = list.filter(r => typeof r.rating === 'number' && !isNaN(r.rating));
  const rW      = rRows.reduce((a, r) => a + (r.reviews || 0), 0);
  const wRating = rRows.reduce((a, r) => a + r.rating * (r.reviews || 0), 0);
  const rating  = rW > 0 ? wRating / rW
                : rRows.length ? rRows.reduce((a, r) => a + r.rating, 0) / rRows.length
                : null;
  if (source === 'thirdparty') {
    const eRows   = list.filter(r => typeof r.errRate === 'number' && !isNaN(r.errRate));
    const eW      = eRows.reduce((a, r) => a + (r.reviews || 0), 0);
    const wErr    = eRows.reduce((a, r) => a + r.errRate * (r.reviews || 0), 0);
    const errRate = eW > 0 ? wErr / eW
                  : eRows.length ? eRows.reduce((a, r) => a + r.errRate, 0) / eRows.length
                  : null;
    return { loc: 'Total', reviews, rating, s5, errRate };
  }
  return { loc: 'Total', reviews, rating, s5 };
}

// One reviews table (Google+Yelp or 3PD). Reused for the single per-source view
// and for the 30-day view that stacks both kinds.
function ReviewTable({ rows, kind, periodLabel, isAdmin }) {
  const filtered = (rows || []).filter(r => kind !== 'thirdparty' || !EXCLUDED_3PD.includes(String(r.loc).trim()));
  const found = filtered.find(r => /^total$/i.test(r.loc));
  const total = found || computeTotal(filtered, kind) || { loc: 'Total', reviews: 0, rating: 0, s5: 0, errRate: 0 };
  const dataRows = filtered.filter(r => !/^total$/i.test(r.loc));
  const tableRows = [...dataRows, { ...total, loc: 'Total' }];
  const title = kind === 'instore'
    ? `In-Store Reviews — Google + Yelp (${periodLabel})`
    : `3rd Party Reviews — UE / DD / GH (${periodLabel})`;
  const exportFilename = kind === 'instore'
    ? 'In-Store Reviews — Google + Yelp.csv'
    : '3rd Party Reviews — UE DD GH.csv';
  const headers = [
    { label: 'Location' },
    { label: '# Reviews', cls: 'right' },
    { label: 'Rating', cls: 'right' },
    { label: '5★', cls: 'right' }, { label: '4★', cls: 'right' }, { label: '3★', cls: 'right' },
    { label: '2★', cls: 'right' }, { label: '1★', cls: 'right' },
    ...(kind === 'instore'
      ? [{ label: 'Yelp', cls: 'right' }, { label: 'Yelp #', cls: 'right' }, { label: 'Google', cls: 'right' }, { label: 'Google #', cls: 'right' }]
      : [{ label: 'UE', cls: 'right' }, { label: 'DD', cls: 'right' }, { label: 'GH', cls: 'right' }, { label: 'Error Rate', cls: 'right' }]),
  ];
  return (
    <div className="table-card">
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
        <div className="table-title" style={{ marginBottom: 0 }}>{title}</div>
        <ExportCsvButton filename={exportFilename} />
      </div>
      <Table
        headers={headers}
        rows={tableRows.map(r => ({
          _cls: /^total$/i.test(r.loc) ? 'total-row' : '',
          cells: [
            r.loc, fmtN(r.reviews), ratingBadge(r.rating),
            fmtN(r.s5), fmtN(r.s4), fmtN(r.s3), fmtN(r.s2), fmtN(r.s1),
            ...(kind === 'instore'
              ? [ratingBadge(r.yelp), fmtN(r.yelpN), ratingBadge(r.google), fmtN(r.gNum)]
              : [ratingBadge(r.ue), ratingBadge(r.dd), ratingBadge(r.gh), errRateBadge(r.errRate)]),
          ],
        }))}
      />
    </div>
  );
}

export default function Reviews({ data, prevData, userRole }) {
  const isAdmin = userRole === 'admin';
  const [source, setSource] = useState('instore');
  const [period, setPeriod] = useState('weekly');
  const [starLoc, setStarLoc] = useState('all');

  // The 30-day tables only ship from Jun 2026 onward; hide the filter otherwise.
  const has30 = !!(data?.reviews?.instore?.day30?.length || data?.reviews?.thirdparty?.day30?.length);
  const periods = has30 ? PERIODS : PERIODS.filter(p => p.id !== 'day30');

  const rowsRaw = ((data?.reviews?.[source]?.[period]) || [])
    .filter(r => source !== 'thirdparty' || !EXCLUDED_3PD.includes(String(r.loc).trim()));
  const periodLabel = period === 'weekly' ? '7 Days' : period === 'day30' ? '30 Days' : '90 Days';

  const total = useMemo(() => {
    const found = rowsRaw.find(r => /^total$/i.test(r.loc));
    return found || computeTotal(rowsRaw, source) || { reviews: 0, rating: 0, s5: 0, errRate: 0 };
  }, [rowsRaw, source]);

  const dataRows = rowsRaw.filter(r => !/^total$/i.test(r.loc));

  // Previous week — same source + period
  const prevRowsRaw = useMemo(() =>
    ((prevData?.reviews?.[source]?.[period]) || [])
      .filter(r => source !== 'thirdparty' || !EXCLUDED_3PD.includes(String(r.loc).trim())),
    [prevData, source, period]);

  const prevTotal = useMemo(() => {
    const found = prevRowsRaw.find(r => /^total$/i.test(r.loc));
    return found || computeTotal(prevRowsRaw, source) || { reviews: 0, rating: 0, s5: 0, errRate: 0 };
  }, [prevRowsRaw, source]);

  const prevDataRows = useMemo(() => prevRowsRaw.filter(r => !/^total$/i.test(r.loc)), [prevRowsRaw]);

  // Platform-level averages (Google/Yelp or UE/DD/GH)
  const currPlat = useMemo(() => platformAvgs(dataRows,     source), [dataRows,     source]);
  const prevPlat = useMemo(() => platformAvgs(prevDataRows, source), [prevDataRows, source]);

  const starData = useMemo(() => {
    let s5 = 0, s4 = 0, s3 = 0, s2 = 0, s1 = 0, label;
    if (starLoc === 'all') {
      dataRows.forEach(r => { s5 += r.s5||0; s4 += r.s4||0; s3 += r.s3||0; s2 += r.s2||0; s1 += r.s1||0; });
      label = 'All Locations';
    } else {
      const row = dataRows.find(r => r.loc === starLoc);
      if (row) { s5 = row.s5||0; s4 = row.s4||0; s3 = row.s3||0; s2 = row.s2||0; s1 = row.s1||0; }
      label = starLoc;
    }
    return {
      labels: ['5 Star', '4 Star', '3 Star', '2 Star', '1 Star'],
      datasets: [{
        label,
        data: [s5, s4, s3, s2, s1],
        backgroundColor: ['#9f7cef', '#c3a8f5', '#b99af3', '#d6c3f8', 'rgba(220,38,38,0.75)'],
        borderRadius: 4,
      }],
    };
  }, [dataRows, starLoc]);

  const ratingChart = useMemo(() => {
    if (source === 'instore') {
      return {
        labels: dataRows.map(r => r.loc),
        datasets: [
          { label: 'Google Rating', data: dataRows.map(r => typeof r.google === 'number' ? r.google : null), backgroundColor: '#9f7cef', borderRadius: 4 },
          { label: 'Yelp Rating',   data: dataRows.map(r => typeof r.yelp   === 'number' ? r.yelp   : null), backgroundColor: '#93c5fd', borderRadius: 4 },
        ],
      };
    }
    return {
      labels: dataRows.map(r => r.loc),
      datasets: [
        { label: 'Uber Eats', data: dataRows.map(r => typeof r.ue === 'number' ? r.ue : null), backgroundColor: '#9f7cef', borderRadius: 4 },
        { label: 'DoorDash',  data: dataRows.map(r => typeof r.dd === 'number' ? r.dd : null), backgroundColor: '#7c3aed', borderRadius: 4 },
        { label: 'Grubhub',   data: dataRows.map(r => typeof r.gh === 'number' ? r.gh : null), backgroundColor: '#93c5fd', borderRadius: 4 },
      ],
    };
  }, [dataRows, source]);

  const chart1Title = source === 'instore'
    ? `Avg Rating by Location (${periodLabel})`
    : `Avg Rating by Location — 3rd Party (${periodLabel})`;
  const chart2Title = source === 'instore'
    ? `Star Distribution — In-Store (${periodLabel})`
    : `Star Distribution — 3rd Party (${periodLabel})`;
  const chart1Filename = source === 'instore'
    ? 'Avg Rating by Location.png'
    : 'Avg Rating by Location — 3rd Party.png';
  const chart2Filename = source === 'instore'
    ? 'Star Distribution — In-Store.png'
    : 'Star Distribution — 3rd Party.png';

  return (
    <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
        <span style={{ fontSize: 11, fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.06em' }}>Reviews</span>
        <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
          <div className="toggle-group">
            {SOURCES.map(s => (
              <button key={s.id} className={`toggle-btn${source === s.id ? ' active' : ''}`} onClick={() => setSource(s.id)}>{s.label}</button>
            ))}
          </div>
          <div className="toggle-group">
            {periods.map(p => (
              <button key={p.id} className={`toggle-btn${period === p.id ? ' active' : ''}`} onClick={() => setPeriod(p.id)}>{p.label}</button>
            ))}
          </div>
        </div>
      </div>

      <div className="kpi-row">
        {source === 'instore' ? (
          <>
            <div className="kpi-card">
              <div className="kpi-label">Avg Rating</div>
              <div className="kpi-value">{typeof total.rating === 'number' && !isNaN(total.rating) ? total.rating.toFixed(1) : '—'}</div>
              <VarChip curr={total.rating} prev={prevTotal.rating} isRating labeled />
              <div className="kpi-change neu">{fmtN(total.reviews)} reviews</div>
            </div>
            <div className="kpi-card">
              <div className="kpi-label">Google Avg Rating</div>
              <div className="kpi-value">{typeof total.google === 'number' && !isNaN(total.google) ? total.google.toFixed(1) : '—'}</div>
              <VarChip curr={total.google} prev={prevTotal.google} isRating labeled />
            </div>
            <div className="kpi-card">
              <div className="kpi-label">Yelp Avg Rating</div>
              <div className="kpi-value">{typeof total.yelp === 'number' && !isNaN(total.yelp) ? total.yelp.toFixed(1) : '—'}</div>
              <VarChip curr={total.yelp} prev={prevTotal.yelp} isRating labeled />
            </div>
          </>
        ) : (
          <>
            <div className="kpi-card">
              <div className="kpi-label">Avg Rating</div>
              <div className="kpi-value">{typeof total.rating === 'number' && !isNaN(total.rating) ? total.rating.toFixed(1) : '—'}</div>
              <VarChip curr={total.rating} prev={prevTotal.rating} isRating labeled />
              <div className="kpi-change neu">{fmtN(total.reviews)} reviews</div>
            </div>
            <div className="kpi-card">
              <div className="kpi-label">Overall Error Rate</div>
              <div className="kpi-value">{typeof total.errRate === 'number' && !isNaN(total.errRate) ? (total.errRate * 100).toFixed(1) + '%' : '—'}</div>
              <VarChip curr={total.errRate} prev={prevTotal.errRate} inverse labeled isPct />
            </div>
            <div className="kpi-card">
              <div className="kpi-label">Uber Eats Rating</div>
              <div className="kpi-value">{typeof total.ue === 'number' && !isNaN(total.ue) ? total.ue.toFixed(1) : '—'}</div>
              <VarChip curr={total.ue} prev={prevTotal.ue} isRating labeled />
            </div>
            <div className="kpi-card">
              <div className="kpi-label">DoorDash Rating</div>
              <div className="kpi-value">{typeof total.dd === 'number' && !isNaN(total.dd) ? total.dd.toFixed(1) : '—'}</div>
              <VarChip curr={total.dd} prev={prevTotal.dd} isRating labeled />
            </div>
            <div className="kpi-card">
              <div className="kpi-label">Grubhub Rating</div>
              <div className="kpi-value">{typeof total.gh === 'number' && !isNaN(total.gh) ? total.gh.toFixed(1) : '—'}</div>
              <VarChip curr={total.gh} prev={prevTotal.gh} isRating labeled />
            </div>
          </>
        )}
      </div>

      <ReviewTable rows={rowsRaw} kind={source} periodLabel={periodLabel} isAdmin={isAdmin} />

      <div className="charts-row">
        <div className="chart-card">
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
            <div className="chart-title" style={{ marginBottom: 0 }}>{chart1Title}</div>
            <ExportImageButton filename={chart1Filename} />
          </div>
          <Bar
            data={ratingChart}
            options={{
              responsive: true,
              plugins: { legend: { position: 'bottom' } },
              scales: { y: { min: source === 'instore' ? 1.0 : 0, max: 5.0, ticks: { stepSize: 1 } } },
            }}
          />
        </div>
        <div className="chart-card">
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 14, flexWrap: 'wrap', gap: 8 }}>
            <div className="chart-title" style={{ marginBottom: 0 }}>{chart2Title}</div>
            <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
              <select
                value={starLoc}
                onChange={e => setStarLoc(e.target.value)}
                style={{ background: '#f3f4f6', border: '1.5px solid var(--border)', color: '#1a1f2e', padding: '4px 10px', borderRadius: 7, fontSize: 11, fontWeight: 600, cursor: 'pointer', fontFamily: "'Montserrat', sans-serif" }}
              >
                {STAR_LOCS.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
              </select>
              <ExportImageButton filename={chart2Filename} />
            </div>
          </div>
          <Bar
            data={starData}
            options={{
              responsive: true,
              plugins: { legend: { position: 'bottom' } },
              scales: { y: { ticks: { stepSize: 1, callback: v => Number.isInteger(v) ? v : null } } },
            }}
          />
        </div>
      </div>
    </>
  );
}

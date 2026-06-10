'use client';

import { useState, useEffect } from 'react';
import { fetchScorecardIndex, fetchScorecard } from '@/lib/api';

const GRANS = [
  { id: 'weekly',  label: 'Weekly' },
  { id: 'period',  label: 'Period' },
  { id: 'quarter', label: 'Quarterly' },
];

// Display the cell exactly as the sheet shows it: `w` is Excel's formatted text
// (keeps %, $ and the sheet's own negative-in-brackets formatting). Fall back to
// the raw value only when no formatted text exists.
function fmtCell(c) {
  if (c.w != null && c.w !== '') return c.w;
  if (c.v == null || c.v === '') return '-';
  return String(c.v);
}

// Soften a sheet fill (#RRGGBB) to a pastel tint so the tables match the rest of
// the dashboard's look instead of the harsh saturated source fills.
function softBg(hex) {
  if (!hex || !/^#[0-9a-f]{6}$/i.test(hex)) return null;
  const r = parseInt(hex.slice(1, 3), 16), g = parseInt(hex.slice(3, 5), 16), b = parseInt(hex.slice(5, 7), 16);
  return `rgba(${r}, ${g}, ${b}, 0.4)`;
}

// Composite Score and Contributor Band get their color from Excel conditional
// formatting (not a static fill), so SheetJS can't read it — derive it here.
const G = 'rgba(52,168,84,0.38)', LG = 'rgba(182,215,168,0.55)', A = 'rgba(255,229,153,0.6)', R = 'rgba(255,92,95,0.4)';
function deriveBg(header, v) {
  if (header === 'Composite Score' && typeof v === 'number') {
    return v >= 4 ? G : v >= 3 ? LG : v >= 2 ? A : R;
  }
  if (header === 'Contributor Band' && typeof v === 'string') {
    const t = v.toLowerCase();
    if (/star|high/.test(t)) return G;
    if (/contributor/.test(t)) return LG;
    if (/low|non/.test(t)) return R;
  }
  return null;
}

// Reproduces the source workbook's color coding (softened), filling in the
// composite-score / contributor-band colors the sheet applies via rules.
function ColorTable({ title, data }) {
  if (!data || !data.headers || !data.headers.length) return null;
  return (
    <div className="table-card" style={{ marginBottom: 16 }}>
      <div className="table-title">{title}</div>
      <table className="scorecard-table">
        <thead>
          <tr>{data.headers.map((h, i) => <th key={i} className={i === 0 ? '' : 'right'}>{h}</th>)}</tr>
        </thead>
        <tbody>
          {data.rows.map((row, ri) => (
            <tr key={ri}>
              {row.map((c, ci) => {
                const bg = softBg(c.bg) || deriveBg(data.headers[ci], c.v);
                return (
                  <td key={ci} className={ci === 0 ? '' : 'right'} style={bg ? { background: bg } : undefined}>
                    {fmtCell(c)}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default function Scorecard() {
  const [index, setIndex] = useState(null);
  const [gran, setGran] = useState('weekly');
  const [item, setItem] = useState('');
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  useEffect(() => {
    fetchScorecardIndex()
      .then(setIndex)
      .catch(e => { setError(e.message); setLoading(false); });
  }, []);

  // Default the item to the first available for the chosen granularity.
  useEffect(() => {
    if (!index) return;
    const list = index[gran] || [];
    if (list.length) {
      setItem(list[0].id);
    } else {
      setItem('');
      setData(null);
      setLoading(false);
    }
  }, [index, gran]);

  useEffect(() => {
    if (!index || !item) return;
    // Guard the granularity↔item race: right after the granularity changes, the
    // item is briefly still the previous granularity's selection. Wait until the
    // item actually belongs to the current granularity before fetching, so we
    // never request a (period, weekly-file) mismatch (which 404s).
    const list = index[gran] || [];
    if (!list.some(o => o.id === item)) return;
    let cancelled = false; // ignore a stale response if the user switches again
    setLoading(true);
    fetchScorecard(gran, item)
      .then(d => { if (!cancelled) { setData(d); setLoading(false); setError(''); } })
      .catch(e => { if (!cancelled) { setError(e.message); setLoading(false); } });
    return () => { cancelled = true; };
  }, [index, gran, item]);

  const list = (index && index[gran]) || [];
  const current = list.find(i => i.id === item);

  return (
    <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
        <span style={{ fontSize: 11, fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.06em' }}>Leadership Scorecard</span>
        <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
          <div className="toggle-group">
            {GRANS.map(g => (
              <button key={g.id} className={`toggle-btn${gran === g.id ? ' active' : ''}`} onClick={() => setGran(g.id)}>{g.label}</button>
            ))}
          </div>
          <select
            value={item}
            onChange={e => setItem(e.target.value)}
            disabled={!list.length}
            style={{ background: '#f3f4f6', border: '1.5px solid var(--border)', color: '#1a1f2e', padding: '6px 14px', borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: 'pointer', fontFamily: "'Montserrat',sans-serif" }}
          >
            {list.length === 0 && <option value="">No data</option>}
            {list.map(o => <option key={o.id} value={o.id}>{o.label}</option>)}
          </select>
        </div>
      </div>

      {error && <div className="loading-screen"><span style={{ color: '#dc2626' }}>{error}</span></div>}
      {!error && loading && (
        <div style={{ textAlign: 'center', padding: '40px 0', color: 'var(--muted)' }}>
          <div className="spinner" style={{ margin: '0 auto 10px' }} />Loading scorecard…
        </div>
      )}
      {!error && !loading && !list.length && (
        <div className="loading-screen"><span>No {GRANS.find(g => g.id === gran)?.label.toLowerCase()} scorecards available.</span></div>
      )}
      {!error && !loading && data && (
        <>
          <ColorTable title={`Area Leader Dashboard${current ? ' — ' + current.label : ''}`} data={data.dashboard} />
          <ColorTable title="Scoring Matrix" data={data.matrix} />
        </>
      )}
    </>
  );
}

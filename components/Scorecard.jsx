'use client';

import { useState, useEffect } from 'react';
import { fetchScorecardIndex, fetchScorecard } from '@/lib/api';

const hexRGB = h => [parseInt(h.slice(1, 3), 16), parseInt(h.slice(3, 5), 16), parseInt(h.slice(5, 7), 16)];

// Pill style: the given color as the background, with text color chosen for
// readability (white on dark/saturated fills like purple/green/red, dark on
// light pastel fills).
function pill(hex) {
  const [r, g, b] = hexRGB(hex);
  const lum = 0.299 * r + 0.587 * g + 0.114 * b;
  return { background: hex, color: lum < 150 ? '#fff' : '#1a1f2e' };
}

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

// The Composite Score is colored by the Performance Rating Key — a conditional-
// formatting rule the xlsx parser can't surface as a cell fill, so we apply it
// here. Same discrete bands for Weekly, Period and Quarterly:
//   STAR ≥4.7 #33A854 · HIGH ≥3.7 #B6D7A8 · CONTRIBUTOR ≥2.7 #FFE599
//   · LOW ≥1.7 #EA9999 · NON <1.7 #FF5C5F.
const COMPOSITE_BANDS = [[4.7, '#33A854'], [3.7, '#B6D7A8'], [2.7, '#FFE599'], [1.7, '#EA9999']];
function compositeColor(v) {
  if (typeof v !== 'number') return null;
  for (const [t, hex] of COMPOSITE_BANDS) if (v >= t) return hex;
  return '#FF5C5F';
}

// Pill color: the sheet's own cell fill, except columns the sheet colors via a
// conditional-formatting rule (Composite Score; Training % ≥87% green / <87%
// red) which we apply explicitly. Uncolored cells stay uncolored.
function cellPill(header, c) {
  if (header === 'Composite Score') {
    const hex = compositeColor(c.v);
    return hex ? pill(hex) : null;
  }
  // Contributor Band — color by the Performance Rating Key (same scale as the
  // composite score), so the band label and score agree.
  if (header === 'Contributor Band' && typeof c.v === 'string') {
    const t = c.v.toLowerCase();
    const hex = /star/.test(t) ? '#33A854' : /high/.test(t) ? '#B6D7A8' : /contributor/.test(t) ? '#FFE599'
              : /low/.test(t) ? '#EA9999' : /non/.test(t) ? '#FF5C5F' : null;
    return hex ? pill(hex) : null;
  }
  if (/training/i.test(header) && typeof c.v === 'number') {
    const n = c.v <= 1.5 ? c.v * 100 : c.v; // accept 0.87 or 87
    return pill(n >= 87 ? '#33A854' : '#FF5C5F');
  }
  return (c.bg && /^#[0-9a-f]{6}$/i.test(c.bg)) ? pill(c.bg) : null;
}

// Colors cells from the sheet's own fills (plus the Composite Score CF rule).
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
                // The first column is a label (location / category / "All
                // Stores"), so it's never colored; every other cell keeps its color.
                const style = ci === 0 ? null : cellPill(data.headers[ci], c);
                return (
                  <td key={ci} className={ci === 0 ? '' : 'right'}>
                    {style ? <span className="sc-badge" style={style}>{fmtCell(c)}</span> : fmtCell(c)}
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

  // Default to the latest available selection for the chosen granularity.
  // The index is sorted ascending (oldest → newest), so the last entry is the
  // most recent week / period / quarter.
  useEffect(() => {
    if (!index) return;
    const list = index[gran] || [];
    if (list.length) {
      setItem(list[list.length - 1].id);
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
        </>
      )}
    </>
  );
}

'use client';

import { useState, useEffect, Fragment } from 'react';
import { fetchScorecardIndex, fetchScorecard } from '@/lib/api';
import { ExportCsvButton } from './ExportButtons';

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

// Only the Composite Score column gets a colored pill; all other columns are plain text.
function cellPill(header, c) {
  if (header === 'Composite Score') {
    const hex = compositeColor(c.v);
    return hex ? pill(hex) : null;
  }
  return null;
}

// Colors cells from the sheet's own fills (plus the Composite Score CF rule).
function ColorTable({ title, data, isAdmin }) {
  if (!data || !data.headers || !data.headers.length) return null;
  return (
    <div className="table-card" style={{ marginBottom: 16 }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
        <div className="table-title" style={{ marginBottom: 0 }}>{title}</div>
        <ExportCsvButton filename={`${title}.csv`} />
      </div>
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

// ─────────────────────────────────────────────────────────────────────────
// Scoring Matrix / Performance Rating Key — static reference tables shown
// above "Area Leader Dashboard". Hardcoded from the source workbooks'
// conditional formatting; nothing here is computed or bound to fetched data.
// The bright purple used in the sheets for the title bar / header row /
// "Total Weight" footer is swapped for the dashboard's own theme accent
// (var(--accent), the same purple used for the top nav and the active
// Weekly/Period/Quarterly toggle) so these tables read as native to the
// dashboard rather than pasted from Excel — every other fill is reproduced
// exactly as specified.
const THEME_PURPLE = 'var(--accent)';
const INTERP_BG = '#ECECEC';
const CELL_BORDER = '1px solid #E0E0E0';

const INTERP_RANGE = ['—', '1.0–1.9', '2.0–2.9', '3.0–3.9', '4.0–4.9', '5.0'];

const PRIME_COST_ROWS = weight => [
  { cat: 'Prime Cost', catBg: '#FDEBD0', metric: 'Labor vs Budget Variance (pts)', weight, bands: ['≥+3.0', '2.0–2.99', '1.0–1.99', '0.01–0.99', '0.0 to -0.99', '≤-1.0'] },
  { cat: 'Prime Cost', catBg: '#FDEBD0', metric: 'COGS vs Budget Variance (pts)', weight, bands: ['≥+3.0', '2.0–2.99', '1.0–1.99', '0.01–0.99', '0.0 to -0.99', '≤-1.0'] },
];

const WEEKLY_ROWS = [
  { cat: 'Sales', catBg: '#D4E6F1', metric: 'Sales vs Budget (%)', weight: '6%', bands: ['<90%', '90–94.9%', '95–98.9%', '99–101.9%', '102–104.9%', '≥105%'] },
  { cat: 'Sales', catBg: '#D4E6F1', metric: 'Sales vs Prior Year (%)', weight: '6%', bands: ['<95%', '95–99.9%', '100–104.9%', '105–107.4%', '107.5–114.9%', '≥115%'] },
  ...PRIME_COST_ROWS('14.5%'),
  { cat: 'Operational', catBg: '#EAF2FF', metric: 'In-Store Rating', weight: '29%', bands: ['<3.8', '3.8–3.99', '4.0–4.29', '4.3–4.49', '4.5–4.69', '≥4.7'] },
  { cat: 'Operational', catBg: '#EAF2FF', metric: '3P Error Rate (%)', weight: '12%', bands: ['>6%', '4.01–6%', '2.01–4%', '1.51–2%', '1.01–1.5%', '≤1%'] },
  { cat: 'Operational', catBg: '#EAF2FF', metric: 'Avoidable Wait (mins)', weight: '9%', bands: ['>8', '>5–8', '>2–5', '>1.5–2', '>1–1.5', '≤1'] },
  { cat: 'Operational', catBg: '#EAF2FF', metric: '3P Rating', weight: '9%', bands: ['<3.8', '3.8–3.99', '4.0–4.29', '4.3–4.49', '4.5–4.69', '≥4.7'] },
];

const PERIOD_ROWS = [
  { cat: 'Sales', catBg: '#D4E6F1', metric: 'Sales vs Budget (%)', weight: '4%', bands: ['<90%', '90–94.9%', '95–98.9%', '99–101.9%', '102–104.9%', '≥105%'] },
  { cat: 'Sales', catBg: '#D4E6F1', metric: 'Sales vs Prior Year (%)', weight: '4%', bands: ['<95%', '95–99.9%', '100–104.9%', '105–107.4%', '107.5–114.9%', '≥115%'] },
  { cat: 'EBITDA', catBg: '#D5F5E3', metric: 'EBITDA vs Budget (%)', weight: '12.5%', bands: ['<75%', '75–84%', '85–94%', '95–100%', '101–105%', '>105%'] },
  ...PRIME_COST_ROWS('10.5%'),
  { cat: 'Operational', catBg: '#EAF2FF', metric: 'In-Store Rating', weight: '21%', bands: ['<3.8', '3.8–3.99', '4.0–4.29', '4.3–4.49', '4.5–4.69', '≥4.7'] },
  { cat: 'Operational', catBg: '#EAF2FF', metric: '3P Error Rate (%)', weight: '8.5%', bands: ['>6%', '4.01–6%', '2.01–4%', '1.51–2%', '1.01–1.5%', '≤1%'] },
  { cat: 'Operational', catBg: '#EAF2FF', metric: 'Avoidable Wait (mins)', weight: '6.5%', bands: ['>8', '>5–8', '>2–5', '>1.5–2', '>1–1.5', '≤1'] },
  { cat: 'Operational', catBg: '#EAF2FF', metric: '3P Rating', weight: '6.5%', bands: ['<3.8', '3.8–3.99', '4.0–4.29', '4.3–4.49', '4.5–4.69', '≥4.7'] },
  { cat: 'Operational', catBg: '#EAF2FF', metric: 'Full Audit Score', weight: '16%', bands: ['<70', '70–74', '75–82', '83–88', '89–93', '≥94'] },
];

// Quarterly = Period rows, plus a training-completion eligibility gate (no
// weight, no interpolation range — it's pass/fail, not a scored band).
const QUARTER_ROWS = [
  ...PERIOD_ROWS,
  {
    cat: 'Eligibility', catBg: '#F2F3F4', metric: 'Training Completion % (Gate)', weight: '—',
    gate: [
      { text: '<87% = Ineligible', bg: '#FAD7D3', span: 3 },
      { text: '≥87% = Eligible', bg: '#D5F5E3', span: 3 },
    ],
    blankInterp: true,
  },
];

const SCORING_MATRICES = {
  weekly:  { title: 'RASA · Store Scoring Matrix – Weekly', rows: WEEKLY_ROWS },
  period:  { title: 'RASA · Period Scoring Matrix',         rows: PERIOD_ROWS },
  quarter: { title: 'RASA · Quaterly Scoring Matrix',       rows: QUARTER_ROWS },
};

const BAND_HEADERS = ['0', '1', '2', '3', '4', '5'];
const BAND_BG = ['#FADBD8', '#FAE5D3', '#FEF9E7', '#D5F5E3', '#D6EAF8', '#E8DAEF'];

const RATING_KEY_ROWS = [
  { perf: 'STAR', score: '4.7 – 5.0', color: '#33A854' },
  { perf: 'HIGH', score: '3.7 – 4.7', color: '#B6D7A8' },
  { perf: 'CONTRIBUTOR', score: '2.7 – 3.7', color: '#FFE599' },
  { perf: 'LOW', score: '1.7 – 2.7', color: '#EA9999' },
  { perf: 'NON', score: '0.0 – 1.7', color: '#FF5C5F' },
];

function ScoringMatrixTable({ gran }) {
  const matrix = SCORING_MATRICES[gran] || SCORING_MATRICES.weekly;
  const thStyle = { border: CELL_BORDER, padding: '7px 10px', background: THEME_PURPLE, color: '#fff', fontWeight: 700, fontSize: 12, textAlign: 'center' };
  const tdBase = { border: CELL_BORDER, padding: '6px 10px', fontSize: 12 };
  return (
    <div style={{ border: '1px solid #E0E0E0', borderRadius: 8, overflow: 'hidden', flex: '2 1 620px', minWidth: 0 }}>
      <div style={{ background: THEME_PURPLE, color: '#fff', fontWeight: 700, fontSize: 13, padding: '10px 12px', textAlign: 'center' }}>
        {matrix.title}
      </div>
      <div style={{ overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', tableLayout: 'fixed' }}>
          <colgroup>
            <col style={{ width: '10%' }} />
            <col style={{ width: '27%' }} />
            <col style={{ width: '8%' }} />
            {BAND_HEADERS.map(h => <col key={h} style={{ width: '9.17%' }} />)}
          </colgroup>
          <thead>
            <tr>
              <th style={thStyle}>Category</th>
              <th style={thStyle}>Metric</th>
              <th style={thStyle}>Weight (%)</th>
              {BAND_HEADERS.map(h => <th key={h} style={thStyle}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {matrix.rows.map((r, i) => (
              <Fragment key={i}>
                <tr>
                  <td style={{ ...tdBase, background: r.catBg, fontWeight: 700 }}>{r.cat}</td>
                  <td style={{ ...tdBase, fontWeight: 700 }}>{r.metric}</td>
                  <td style={{ ...tdBase, fontWeight: 700, textAlign: 'center' }}>{r.weight}</td>
                  {r.gate
                    ? r.gate.map((g, gi) => (
                        <td key={gi} colSpan={g.span} style={{ ...tdBase, background: g.bg, fontWeight: 600, textAlign: 'center' }}>{g.text}</td>
                      ))
                    : r.bands.map((b, bi) => (
                        <td key={bi} style={{ ...tdBase, background: BAND_BG[bi], fontWeight: 600, textAlign: 'center' }}>{b}</td>
                      ))}
                </tr>
                <tr>
                  <td style={{ ...tdBase, background: INTERP_BG, color: '#999999', fontStyle: 'italic' }} colSpan={3}>Interpolated score range</td>
                  {(r.blankInterp ? ['', '', '', '', '', ''] : INTERP_RANGE).map((v, vi) => (
                    <td key={vi} style={{ ...tdBase, background: INTERP_BG, color: '#555555', textAlign: 'center' }}>{v}</td>
                  ))}
                </tr>
              </Fragment>
            ))}
            <tr>
              <td style={{ ...tdBase, background: THEME_PURPLE, color: '#fff', fontWeight: 700 }}>Total Weight (%)</td>
              <td style={{ ...tdBase, background: THEME_PURPLE }} />
              <td style={{ ...tdBase, background: THEME_PURPLE, color: '#fff', fontWeight: 700, textAlign: 'center' }}>100%</td>
              {BAND_HEADERS.map(h => <td key={h} style={{ ...tdBase, background: THEME_PURPLE }} />)}
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

function RatingKeyTable() {
  const thStyle = { border: CELL_BORDER, padding: '7px 10px', background: 'var(--bg3)', fontWeight: 700, fontSize: 12, textAlign: 'center', color: 'var(--muted)' };
  const tdBase = { border: CELL_BORDER, padding: '6px 10px', fontSize: 12, background: '#F3F3F3' };
  return (
    <div style={{ border: '1px solid #E0E0E0', borderRadius: 8, overflow: 'hidden', width: 320, maxWidth: '100%' }}>
      <div style={{ background: THEME_PURPLE, color: '#fff', fontWeight: 700, fontSize: 13, padding: '10px 12px', textAlign: 'center' }}>
        Performance Rating Key
      </div>
      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
        <thead>
          <tr>
            <th style={thStyle}>Performance</th>
            <th style={thStyle}>Score</th>
            <th style={thStyle}>Colour</th>
          </tr>
        </thead>
        <tbody>
          {RATING_KEY_ROWS.map(r => (
            <tr key={r.perf}>
              <td style={{ ...tdBase, fontWeight: 700 }}>{r.perf}</td>
              <td style={{ ...tdBase, textAlign: 'center' }}>{r.score}</td>
              <td style={{ ...tdBase, padding: 0 }}><div style={{ background: r.color, height: 22, width: '100%' }} /></td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// Collapsed by default — the matrix only renders once the dropdown-style
// button is clicked.
function ScoringMatrixSection({ gran }) {
  const [open, setOpen] = useState(false);
  const matrix = SCORING_MATRICES[gran] || SCORING_MATRICES.weekly;
  return (
    <div className="table-card" style={{ marginBottom: 16 }}>
      <button
        type="button"
        onClick={() => setOpen(o => !o)}
        style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', width: '100%', gap: 10, background: '#f3f4f6', border: '1.5px solid var(--border)', color: '#1a1f2e', padding: '10px 14px', borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: 'pointer', fontFamily: "'Montserrat',sans-serif" }}
      >
        <span>{matrix.title}</span>
        <span style={{ display: 'inline-block', transition: 'transform 0.15s', transform: open ? 'rotate(180deg)' : 'none' }}>▾</span>
      </button>
      {open && (
        <div style={{ marginTop: 12 }}>
          <ScoringMatrixTable gran={gran} />
        </div>
      )}
    </div>
  );
}

export default function Scorecard({ userRole }) {
  const isAdmin = userRole === 'admin';
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
          <ScoringMatrixSection gran={gran} />
          <ColorTable title={`Area Leader Dashboard${current ? ' — ' + current.label : ''}`} data={data.dashboard} isAdmin={isAdmin} />
          <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
            <RatingKeyTable />
          </div>
        </>
      )}
    </>
  );
}

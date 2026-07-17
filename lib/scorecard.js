// Leadership Scorecard data source.
//
// Lives in scorecard/{weekly,period,quarter} at the project root (separate from
// the weekly-review data/ folder). Each .xlsx is a full "RASA … Performance
// Score" workbook; for the dashboard we only surface two sheets per selection:
//   • the granularity's Area Leader Dashboard (composite scorecard), and
//   • the Scoring Matrix reference table (shown beneath it).
import fs from 'fs';
import path from 'path';
import * as XLSX from 'xlsx';
import { weekInfoForLabel } from './fiscalCalendar.js';

const SCORECARD_DIR = path.join(process.cwd(), 'scorecard');

const DASHBOARD_SHEET = {
  weekly:  'Weekly Area Leader Dashboard',
  period:  'Periodic Area Leader Dashboard',
  quarter: 'Quaterly Area Leader Dashboard',
};

const MONTHS = {
  jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3, apr: 4, april: 4,
  may: 5, jun: 6, june: 6, jul: 7, july: 7, aug: 8, august: 8, sep: 9, sept: 9,
  september: 9, oct: 10, october: 10, nov: 11, november: 11, dec: 12, december: 12,
};

function listDir(sub) {
  const dir = path.join(SCORECARD_DIR, sub);
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir).filter(f => /\.xlsx$/i.test(f) && !f.startsWith('~$'));
}

// "RASA Weekly Performance Score - (May 25 2026 - May 31 2026).xlsx"
// → { id, label, sort }. Label maps the start date to fiscal P#W# when possible.
function weeklyItem(filename) {
  const m = /\(([A-Za-z]+)\s+(\d+)\s+(\d{4})\s*-\s*([A-Za-z]+)\s+(\d+)\s+(\d{4})\)/.exec(filename);
  let label = filename.replace(/\.xlsx$/i, ''), sort = filename;
  if (m) {
    const [, mon1, d1, y1, , d2] = m;
    const fiscal = weekInfoForLabel(`${mon1} ${d1}`);
    const range = `${mon1} ${d1}–${d2}`;
    label = fiscal ? `P${fiscal.period} W${fiscal.weekInPeriod} · ${range}` : range;
    const mi = MONTHS[mon1.toLowerCase()] || 0;
    // YYYYMMDD so ascending sort is chronological (and the latest week sorts last,
    // even across a calendar-year boundary).
    sort = y1 + String(mi).padStart(2, '0') + String(parseInt(d1, 10)).padStart(2, '0');
  }
  return { id: filename, label, sort };
}

// "P3 RASA Leadership Performance Score.xlsx" → { id, label:"P3" } (prefix + number).
function tokenItem(filename, re, prefix) {
  const m = re.exec(filename);
  const label = m ? prefix + m[1] : filename.replace(/\.xlsx$/i, '');
  // Zero-pad the number so ascending sort is numeric (P3 < P10), not lexicographic.
  const sort = m ? String(parseInt(m[1], 10)).padStart(4, '0') : filename;
  return { id: filename, label, sort };
}

// Index of every available scorecard, grouped by granularity.
export function listScorecards() {
  const weekly = listDir('weekly').map(weeklyItem).sort((a, b) => a.sort.localeCompare(b.sort));
  const period = listDir('period').map(f => tokenItem(f, /\bP(\d+)\b/i, 'P')).sort((a, b) => a.sort.localeCompare(b.sort));
  const quarter = listDir('quarter').map(f => tokenItem(f, /\bQ(\d+)\b/i, 'Q')).sort((a, b) => a.sort.localeCompare(b.sort));
  return { weekly, period, quarter };
}

const isStr = v => typeof v === 'string' && v.trim() !== '';

// The fill color of a cell as "#RRGGBB", mirroring the sheet's conditional
// formatting. White / no-fill → null (rendered with the default background).
function bgFromCell(cell) {
  let rgb = cell && cell.s && cell.s.fgColor && cell.s.fgColor.rgb;
  if (!rgb) return null;
  if (rgb.length === 8) rgb = rgb.slice(2); // strip alpha (AARRGGBB → RRGGBB)
  if (/^F{6}$/i.test(rgb)) return null;     // white
  return '#' + rgb;
}

// Read a block from `firstHeader` (e.g. "Restaurant" / "Category") at or after
// `startRow`, returning each cell as { v, bg } so the renderer can reproduce the
// sheet's colors.
function parseBlock(ws, firstHeader, stopFn, startRow = 0) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true });
  let hi = -1;
  for (let i = startRow; i < rows.length; i++) {
    if (rows[i] && isStr(rows[i][0]) && rows[i][0].trim() === firstHeader) { hi = i; break; }
  }
  if (hi < 0) return { headers: [], rows: [] };
  const headers = [];
  for (let c = 0; c < (rows[hi] || []).length; c++) {
    if (rows[hi][c] == null || String(rows[hi][c]).trim() === '') break;
    headers.push(String(rows[hi][c]).trim());
  }
  const out = [];
  for (let i = hi + 1; i < rows.length; i++) {
    if (stopFn(rows[i])) break;
    out.push(headers.map((_, c) => {
      const cell = ws[XLSX.utils.encode_cell({ r: i, c })];
      // `w` is the cell's display text exactly as Excel shows it (keeps %, $,
      // and the sheet's own negative-in-brackets formatting); `v` is the raw
      // value (used for deriving composite-score / band colors).
      return {
        v: cell ? cell.v : null,
        w: cell && cell.w != null ? String(cell.w).trim() : null,
        bg: bgFromCell(cell),
      };
    }));
  }
  return { headers, rows: out };
}

// Area Leader Dashboard: restaurant rows up to the first blank/numeric row
// (excludes the second sub-table and the column-index helper row).
const parseDashboard = ws => parseBlock(ws, 'Restaurant', r => !r || !isStr(r[0]));

// The Scoring Matrix sheet holds three titled matrices; pick the one that
// matches the granularity, then read its block (stopping at the next title).
const MATRIX_TITLE = {
  weekly:  /Scoring Matrix.*Weekly/i,   // "RASA · Store Scoring Matrix - Weekly"
  period:  /Period Scoring Matrix/i,
  quarter: /Qua?terly Scoring Matrix/i, // sheet spells it "Quaterly"
};
function parseMatrix(ws, granularity) {
  const re = MATRIX_TITLE[granularity];
  let start = 0;
  if (re) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true });
    const ti = rows.findIndex(r => r && isStr(r[0]) && re.test(r[0]));
    if (ti >= 0) start = ti + 1;
  }
  return parseBlock(ws, 'Category', r => {
    const blank = !r || r.every(c => c == null || String(c).trim() === '');
    return blank || (isStr(r[0]) && /^RASA\s*·/i.test(r[0]));
  }, start);
}

// Load one scorecard selection → its dashboard + the scoring matrix.
// `id` is the exact filename within scorecard/<granularity>/.
export function loadScorecard(granularity, id) {
  const sheetName = DASHBOARD_SHEET[granularity];
  if (!sheetName) { const e = new Error('Unknown granularity'); e.code = 'BAD_GRANULARITY'; throw e; }
  if (!id || id.includes('/') || id.includes('\\') || id.includes('..')) {
    const e = new Error('Invalid scorecard id'); e.code = 'BAD_ID'; throw e;
  }
  const file = path.join(SCORECARD_DIR, granularity, id);
  if (!fs.existsSync(file)) { const e = new Error('Scorecard not found'); e.code = 'NOT_FOUND'; throw e; }

  const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: true, cellStyles: true });
  const dash = wb.Sheets[sheetName];
  const matrix = wb.Sheets['Scoring Matrix'];
  return {
    dashboard: dash ? parseDashboard(dash) : { headers: [], rows: [] },
    matrix:    matrix ? parseMatrix(matrix, granularity) : { headers: [], rows: [] },
  };
}

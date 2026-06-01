// Formatters mirroring the helpers in the original HTML dashboard.
// All of these are pure and safe to use in both server (route handlers) and client (components).

// Accounting display convention used everywhere in the dashboard:
//   • Positive values get no sign prefix.
//   • Negative values are wrapped in parentheses with no minus sign.
// e.g. fmt$(1234) -> "$1,234"; fmt$(-1234) -> "($1,234)".
function _acct(absStr, neg) { return neg ? `(${absStr})` : absStr; }

// NA passthrough: when the source XLSX cell literally contains "NA"/"N/A",
// the parser preserves it as a string so the formatters can render "NA".
function _isNA(v) {
  return typeof v === 'string' && /^n\/?a$/i.test(v.trim());
}

export function fmt$(v) {
  if (_isNA(v)) return 'NA';
  if (v == null || isNaN(v)) return '-';
  const n = Number(v);
  const r = Math.round(n);
  return _acct('$' + Math.abs(r).toLocaleString('en-US'), r < 0);
}

// Same as fmt$ but preserves two decimal places — for per-guest / per-order
// amounts (AOV, spend) where cents matter (e.g. $63.97, not $64).
export function fmt$2(v) {
  if (_isNA(v)) return 'NA';
  if (v == null || isNaN(v)) return '-';
  const n = Number(v);
  return _acct('$' + Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }), n < 0);
}

export function fmtN(v) {
  if (_isNA(v)) return 'NA';
  if (v == null || v === '-') return '-';
  if (typeof v === 'number') {
    const r = Math.round(v);
    return _acct(Math.abs(r).toLocaleString('en-US'), r < 0);
  }
  return String(v);
}

export function fmtPct(v) {
  if (_isNA(v)) return 'NA';
  if (v == null || isNaN(v)) return '-';
  const n = Number(v);
  // Values stored as fractions (0.05 = 5%) — multiply by 100
  return _acct(Math.abs(n * 100).toFixed(1) + '%', n < 0);
}

export function fmtVar(v) {
  if (_isNA(v)) return 'NA';
  if (v == null || isNaN(v)) return '-';
  const n = Number(v);
  return _acct(Math.abs(n * 100).toFixed(1) + '%', n < 0);
}

export function fmtVarPC(v) {
  if (_isNA(v)) return 'NA';
  if (v == null || isNaN(v)) return '-';
  const n = Number(v);
  return _acct(Math.abs(n * 100).toFixed(1) + '%', n < 0);
}

// Same shape as fmtVarPC but rendered as a colored badge:
//   negative → red, positive → green, zero → neutral.
// Used in the Cost tab's "Var PC" column where negative means over-budget.
export function fmtVarPCColored(v) {
  if (_isNA(v)) return 'NA';
  if (v == null || isNaN(v)) return '-';
  const n = Number(v);
  const cls = n < 0 ? 'red' : n > 0 ? 'green' : 'neutral';
  const txt = _acct(Math.abs(n * 100).toFixed(1) + '%', n < 0);
  return `<span class="badge ${cls}">${txt}</span>`;
}

// Conditional-formatting pill for variance % cells in the Sales/Overview tables.
// Rule: <=0% red, 0.01-3% yellow, >3% green. Rendered through the shared
// .badge class so colors stay consistent across the dashboard.
export function fmtVarColored(v) {
  if (_isNA(v)) return 'NA';
  if (v == null || isNaN(v)) return '-';
  const n = Number(v);
  const cls = n <= 0 ? 'red' : n <= 0.03 ? 'amber' : 'green';
  const txt = _acct(Math.abs(n * 100).toFixed(1) + '%', n < 0);
  return `<span class="badge ${cls}">${txt}</span>`;
}

// Chart palette
export const CHART_COLORS = {
  primary: '#9f7cef',
  primaryLight: '#d6c3f8',
  primaryLighter: '#e8ddfb',
  pink: '#f9a8d4',
  pinkLight: '#f9a8a8',
  green: '#86efac',
  yellow: '#fcd34d',
  blue: '#93c5fd',
  gray: 'rgba(209,213,219,0.7)',
  amber: '#f59e0b',
  red: '#dc2626',
};

export const PIE_COLORS = ['#9f7cef', '#f9a8d4', '#86efac', '#fcd34d', '#93c5fd', '#f9a8a8'];

// Fiscal calendar (4-4-5 retail style) for FY2026.
// Encodes each period's start date and week count; from these we derive every
// week-start date so a week-folder label ("Week of May 11") can be mapped to its
// fiscal { period, weekInPeriod }. Pure module — safe on server and client.
//
// Source: WBR fiscal calendar table (P1 starts 2025-12-29; 5-week periods are
// P3, P6, P9, P12; all others are 4 weeks; 52 weeks total).

const PERIODS = [
  { period: 1,  start: [2025, 12, 29], weeks: 4 },
  { period: 2,  start: [2026, 1, 26],  weeks: 4 },
  { period: 3,  start: [2026, 2, 23],  weeks: 5 },
  { period: 4,  start: [2026, 3, 30],  weeks: 4 },
  { period: 5,  start: [2026, 4, 27],  weeks: 4 },
  { period: 6,  start: [2026, 5, 25],  weeks: 5 },
  { period: 7,  start: [2026, 6, 29],  weeks: 4 },
  { period: 8,  start: [2026, 7, 27],  weeks: 4 },
  { period: 9,  start: [2026, 8, 24],  weeks: 5 },
  { period: 10, start: [2026, 9, 28],  weeks: 4 },
  { period: 11, start: [2026, 10, 26], weeks: 4 },
  { period: 12, start: [2026, 11, 23], weeks: 5 },
];

const DAY_MS = 24 * 60 * 60 * 1000;

const MONTHS = {
  jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3,
  apr: 4, april: 4, may: 5, jun: 6, june: 6, jul: 7, july: 7,
  aug: 8, august: 8, sep: 9, sept: 9, september: 9,
  oct: 10, october: 10, nov: 11, november: 11, dec: 12, december: 12,
};

// All 52 week starts, keyed by "month-day" → { period, weekInPeriod }.
// Within a single fiscal year each (month, day) week-start is unique.
const WEEK_INDEX = (() => {
  const idx = {};
  for (const p of PERIODS) {
    const base = Date.UTC(p.start[0], p.start[1] - 1, p.start[2]);
    for (let k = 0; k < p.weeks; k++) {
      const d = new Date(base + k * 7 * DAY_MS);
      const key = `${d.getUTCMonth() + 1}-${d.getUTCDate()}`;
      idx[key] = { period: p.period, weekInPeriod: k + 1 };
    }
  }
  return idx;
})();

// Return the number of weeks in a given fiscal period number (1–12).
export function weeksInPeriod(periodNum) {
  const p = PERIODS.find(x => x.period === periodNum);
  return p ? p.weeks : 4;
}

// Return the sequential fiscal week number (1–52) for a week label.
// "Week of June 15" → P6W4 → 4+4+5+4+4+4 = 21 prior weeks + 4 = 25.
export function weekNumForLabel(text) {
  const info = weekInfoForLabel(text);
  if (!info) return null;
  const { period, weekInPeriod } = info;
  let n = weekInPeriod;
  for (const p of PERIODS) {
    if (p.period < period) n += p.weeks;
  }
  return n;
}

// Map a week label/folder name (e.g. "Week of May 11") to { period, weekInPeriod }.
// Returns null when the text has no month+day or doesn't line up with a week start.
export function weekInfoForLabel(text) {
  if (!text) return null;
  const m = /([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?/.exec(String(text));
  if (!m) return null;
  const month = MONTHS[m[1].toLowerCase()];
  const day = parseInt(m[2], 10);
  if (!month || !day) return null;
  return WEEK_INDEX[`${month}-${day}`] || null;
}

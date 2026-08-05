// ─────────────────────────────────────────────────────────────────────────────
// Reads a worksheet's real Excel "Conditional Formatting" (cellIs) rules
// straight out of the raw .xlsx (a zip of XML) — the `xlsx` package we use
// elsewhere for values doesn't expose these at all. Used only by the
// Catering tab's Sheet1 metrics table (components/CateringSales.jsx via
// lib/xlsxParser.js's parseCateringWB), which needs the sheet's own
// Green/Yellow/Red rules rather than an approximated ratio-to-plan rule —
// this is what lets an updated rule for just the newest week's column
// (e.g. a single-cell sqref like "AW13") apply only to that column while
// older weeks keep whatever rule already covered them.
import JSZip from 'jszip';

// "AA" -> 26 (0-indexed column number).
function colLetterToNum(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
// 0 -> "A", 26 -> "AA".
function colNumToLetter(n) {
  let s = '';
  n += 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
// "AW13" -> { col: 48, row: 13 }. Row stays 1-indexed (Excel-native).
function parseCellRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return null;
  return { col: colLetterToNum(m[1]), row: parseInt(m[2], 10) };
}
// "D12:AD13" -> { c1,r1,c2,r2 }. A bare "AW13" is treated as a 1×1 range.
function parseRange(range) {
  const [a, b] = range.split(':');
  const p1 = parseCellRef(a);
  const p2 = b ? parseCellRef(b) : p1;
  if (!p1 || !p2) return null;
  return { c1: Math.min(p1.col, p2.col), c2: Math.max(p1.col, p2.col), r1: Math.min(p1.row, p2.row), r2: Math.max(p1.row, p2.row) };
}
function rangeContains(range, col, row) {
  return col >= range.c1 && col <= range.c2 && row >= range.r1 && row <= range.r2;
}
function rangeArea(range) {
  return (range.c2 - range.c1 + 1) * (range.r2 - range.r1 + 1);
}

// Known dxf fill colors → this app's color tokens (see COL_MINT/AMBER/PINK
// in lib/xlsxParser.js). Falls back to '' (no highlight) for anything else
// (e.g. this sheet also has an unrelated light-purple dxf used elsewhere).
const DXF_COLOR_TOKEN = {
  FFB7E1CD: 'mint',
  FFFFE599: 'amber',
  FFEA9999: 'pink',
};

function evalRule(rule, value) {
  if (typeof value !== 'number') return false;
  const [f1, f2] = rule.formulas;
  switch (rule.operator) {
    case 'greaterThanOrEqual': return value >= f1;
    case 'greaterThan':        return value > f1;
    case 'lessThanOrEqual':    return value <= f1;
    case 'lessThan':           return value < f1;
    case 'equal':              return value === f1;
    case 'between':            return value >= Math.min(f1, f2) && value <= Math.max(f1, f2);
    default: return false;
  }
}

// Parses every <conditionalFormatting> block in a worksheet XML into
// { ranges: [{c1,c2,r1,r2}], rules: [{dxfId, priority, operator, formulas}] }
// groups, alongside a dxfId -> color-token lookup built from styles.xml.
function parseWorksheetCF(sheetXml, stylesXml) {
  const dxfBlockMatch = /<dxfs[^>]*>([\s\S]*?)<\/dxfs>/.exec(stylesXml);
  const dxfs = dxfBlockMatch ? dxfBlockMatch[1].match(/<dxf>[\s\S]*?<\/dxf>/g) || [] : [];
  const dxfToken = dxfs.map(dxf => {
    const rgb = /<fgColor rgb="([0-9A-Fa-f]{8})"/.exec(dxf);
    return rgb ? (DXF_COLOR_TOKEN[rgb[1].toUpperCase()] || '') : '';
  });

  // A single logical rule-set (e.g. "green if >=5 / yellow if <=4 / red if
  // <3" for one row) is usually written as THREE separate
  // <conditionalFormatting> blocks that all repeat the identical sqref, one
  // cfRule each — not one block containing three rules. Group by the exact
  // sqref string so all of them get evaluated together per cell, rather than
  // only ever considering whichever single block happened to be picked.
  const blocks = sheetXml.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/g) || [];
  const groupBySqref = new Map();
  for (const block of blocks) {
    const sqrefMatch = /sqref="([^"]*)"/.exec(block);
    if (!sqrefMatch) continue;
    const sqref = sqrefMatch[1];
    const ruleMatches = block.match(/<cfRule[\s\S]*?<\/cfRule>/g) || [];
    for (const rm of ruleMatches) {
      const type = /type="([^"]*)"/.exec(rm)?.[1];
      if (type !== 'cellIs') continue; // only the plain value-comparison rules this table uses
      const dxfId = parseInt(/dxfId="(\d+)"/.exec(rm)?.[1], 10);
      const priority = parseInt(/priority="(\d+)"/.exec(rm)?.[1] ?? '0', 10);
      const operator = /operator="([^"]*)"/.exec(rm)?.[1];
      const formulas = (rm.match(/<formula>([^<]*)<\/formula>/g) || [])
        .map(f => parseFloat(/<formula>([^<]*)<\/formula>/.exec(f)[1]));
      if (!operator || formulas.some(isNaN)) continue;
      if (!groupBySqref.has(sqref)) {
        const ranges = sqref.split(/\s+/).map(parseRange).filter(Boolean);
        if (!ranges.length) continue;
        groupBySqref.set(sqref, { ranges, rules: [] });
      }
      groupBySqref.get(sqref).rules.push({ dxfId, priority, operator, formulas, color: dxfToken[dxfId] || '' });
    }
  }
  return [...groupBySqref.values()];
}

// Loads the given sheet's conditional-formatting rule groups from the .xlsx
// buffer. Returns null (caller falls back to the old ratio approximation) if
// anything about this can't be found — a differently-structured workbook
// shouldn't break catering parsing entirely.
export async function readConditionalFormatGroups(buffer, sheetName) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const workbookXml = await zip.file('xl/workbook.xml').async('string');
    const sheetTag = new RegExp(`<sheet [^>]*name="${sheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`).exec(workbookXml);
    const rId = sheetTag && /r:id="([^"]*)"/.exec(sheetTag[0])?.[1];
    if (!rId) return null;

    const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const relTag = new RegExp(`<Relationship Id="${rId}"[^>]*/>`).exec(relsXml);
    const target = relTag && /Target="([^"]*)"/.exec(relTag[0])?.[1];
    if (!target) return null;

    const sheetPath = 'xl/' + target.replace(/^\/?xl\//, '');
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) return null;
    const sheetXml = await sheetFile.async('string');
    const stylesXml = await zip.file('xl/styles.xml').async('string');
    return parseWorksheetCF(sheetXml, stylesXml);
  } catch {
    return null;
  }
}

// Resolves the color token for one cell. When multiple rule groups cover the
// same cell (e.g. a broad historical-weeks range AND a newer single-cell
// override for just the current week), the group with the SMALLEST area wins
// — a single-cell sqref is a more specific, deliberately-added override, not
// a coincidence. Within a group, when more than one rule matches the value,
// the highest-priority-number (last-applied) rule wins, matching how Excel
// itself resolves overlapping, non-stopIfTrue cellIs rules.
export function resolveCellColor(groups, colNum, rowNum, value) {
  if (!groups) return '';
  const matching = groups.filter(g => g.ranges.some(r => rangeContains(r, colNum, rowNum)));
  if (!matching.length) return '';
  matching.sort((a, b) => Math.min(...a.ranges.map(rangeArea)) - Math.min(...b.ranges.map(rangeArea)));
  const group = matching[0];
  const hits = group.rules.filter(r => evalRule(r, value));
  if (!hits.length) return '';
  hits.sort((a, b) => a.priority - b.priority);
  return hits[hits.length - 1].color;
}

export { colNumToLetter };

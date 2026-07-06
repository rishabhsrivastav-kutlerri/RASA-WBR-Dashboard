// ─────────────────────────────────────────────────────────────────────────────
// WBR XLSX Parser
// Reads the three workbooks for a given week (WBR, Loyalty, Catering) and
// returns a normalized data object the dashboard renders from.
//
// This is a port of the parseWBR / parseLoyalty / parseCateringWB functions
// embedded in the original HTML dashboard, plus per-location and YTD sections
// that were previously hardcoded.
// ─────────────────────────────────────────────────────────────────────────────
import * as XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';
import { weekNumForLabel } from './fiscalCalendar.js';

// ── Helpers ─────────────────────────────────────────────────────────────────
function sheetRows(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
}
function safeNum(v) { return typeof v === 'number' ? v : null; }
function safeStr(v) { return v != null ? String(v).trim() : '-'; }
function num(v, d = 0) {
  if (v == null) return d;
  if (typeof v === 'number') return v;
  const n = parseFloat(v);
  return isNaN(n) ? d : n;
}
// Percentage cell parser. The flash sheets store percentages as fractions
// (0.305 = 30.5%), but some Totals cells are authored as percent-formatted
// strings ("24.80%"). parseFloat("24.80%") yields 24.8, which the renderer's
// ×100 would turn into 2480.0% — so a "%" string is divided back to its fraction.
function pctNum(v, d = 0) {
  if (typeof v === 'string' && v.trim().endsWith('%')) {
    const n = parseFloat(v);
    return isNaN(n) ? d : n / 100;
  }
  return num(v, d);
}
function isNAToken(v) {
  return typeof v === 'string' && /^n\/?a$/i.test(v.trim());
}
// Returns 'NA' when the source cell literally contains "NA"/"N/A", else a number.
// Use in display-only fields (variance %, ratings) so the renderer can show "NA"
// instead of converting to 0 or '-'.
function numNA(v, d = 0) {
  if (isNAToken(v)) return 'NA';
  return num(v, d);
}
// Returns 'NA' when a percent variance is mathematically undefined — i.e.
// the source cell already says NA, or LY=0 (division by zero is undefined,
// including the 0/0 case). Otherwise passes the raw value through.
function naIfDivByZero(varP, ly /*, varD */) {
  if (isNAToken(varP)) return 'NA';
  if (ly === 0) return 'NA';
  return varP;
}

// Column layout of the Budget Import - 2026 secondary tables (cols 17-58).
// Each row is "W#" (fiscal week 1-52). Columns per section:
//   WEEKLY: 17=key, 18=Total, 19-25=RCs
//   YTD:    28=key, 29=Total, 30-36=RCs
//   PTD:    39=key, 40=Total, 41-47=RCs
//   QTD:    50=key, 51=Total, 52-58=RCs
const BUDGET_RC_COLS = ['In-Store', 'Takeout', 'Delivery', 'Catering', 'Offsites', 'Delivery Fee', 'Discounts/Refunds'];

// Parse a budget sheet (consolidated or per-location) into { weekNum: { weekly:{}, ptd:{} } }.
// All budget sheets share the same column layout: W# rows, WEEKLY cols 19-25, PTD cols 41-47.
function parseBudgetSheet(wb, sheetName) {
  const rows = sheetRows(wb, sheetName);
  if (!rows || rows.length < 12) return null;
  const byWeek = {};
  for (const r of rows) {
    if (!r || !r[17]) continue;
    const m = /^W(\d+)$/.exec(safeStr(r[17]));
    if (!m) continue;
    const weekNum = parseInt(m[1], 10);
    const weekly = {}, ptd = {};
    for (let j = 0; j < BUDGET_RC_COLS.length; j++) {
      const wv = r[19 + j], pv = r[41 + j];
      if (typeof wv === 'number') weekly[BUDGET_RC_COLS[j]] = wv;
      if (typeof pv === 'number') ptd[BUDGET_RC_COLS[j]] = pv;
    }
    byWeek[weekNum] = { weekly, ptd };
  }
  return Object.keys(byWeek).length ? byWeek : null;
}

const LOC_BUDGET_SHEETS = {
  'Ballpark':         'Budget - BP - 2026',
  'MVT':              'Budget - MVT - 2026',
  'National Landing': 'Budget - NL - 2026',
  'Mosaic':           'Budget - Mosaic - 2026',
  'Rockville':        'Budget - Rockville - 2026',
};

// Revenue-center name normalization (XLSX has variants like "In - Store" / "Off-sites")
const RC_NAME_MAP = {
  'In - Store': 'In-Store',
  'Takeout ': 'Takeout',
  'Delivery ': 'Delivery',
  'Off-sites': 'Offsites',
  'Off-Sites': 'Offsites',
  'Delivery Fee Income': 'Delivery Fee',
  'Discounts/Refunds + Sales Adjustment + Open Ticket': 'Discounts/Refunds',
  'Discounts / Refunds + Sales Adjustments + Open Tickets': 'Discounts/Refunds',
};
function mapRC(name) {
  const s = safeStr(name);
  return RC_NAME_MAP[s] || s;
}

// ── Generic extractors ─────────────────────────────────────────────────────
function extractFlashSales(ws) {
  return ws.slice(1).filter(r => r && r[0]).map(r => ({
    loc:    safeStr(r[0]),
    actual: num(r[1]),
    ly:     num(r[2]),
    budget: num(r[3]),
    varLY:  num(r[4]),
    varBud: num(r[5]),
  }));
}
function extractFlashCosts(ws) {
  return ws.slice(1).filter(r => r && r[0]).map(r => ({
    loc:      safeStr(r[0]),
    laborAct: pctNum(r[1]),
    laborBud: pctNum(r[2]),
    cogsAct:  pctNum(r[3]),
    cogsBud:  pctNum(r[4]),
    pcAct:    pctNum(r[5]),
    pcBud:    pctNum(r[6]),
    varPC:    pctNum(r[7]),
  }));
}
function extractRC(ws) {
  // Returns rows except the Total row (which is rebuilt by the UI when needed)
  const out = [];
  for (const r of ws.slice(1)) {
    if (!r || !r[0]) continue;
    const c = safeStr(r[0]);
    if (c.toLowerCase() === 'total' || c.toLowerCase() === 'totals') break;
    const ly   = num(r[2]);
    const varD = num(r[3]);
    out.push({
      center: mapRC(c),
      actual: num(r[1]),
      ly,
      varD,
      varP: naIfDivByZero(typeof r[4] === 'number' ? r[4] : 0, ly, varD),
    });
  }
  return out;
}
function extractRCYtd(ws) {
  // "Total Revenue Center P v A" style:
  // cols: Revenue Center | YTD Actual | YTD Plan | PY YTD | Var$ Plan | Var% Plan | Var$ PY | VAR% PY
  const out = [];
  for (const r of ws.slice(1)) {
    if (!r || !r[0]) continue;
    const c = safeStr(r[0]);
    const isTotal = c.toLowerCase() === 'total' || c.toLowerCase() === 'totals';
    const ly      = num(r[3]);
    const varD    = num(r[6]);
    const budget  = typeof r[2] === 'number' ? r[2] : null;
    const varDBud = typeof r[4] === 'number' ? r[4] : (budget != null ? num(r[1]) - budget : null);
    const varPBud = typeof r[5] === 'number' ? r[5] : null;
    out.push({
      center:  isTotal ? 'Totals' : mapRC(c),
      actual:  num(r[1]),
      ly,
      varD,
      varP:    naIfDivByZero(typeof r[7] === 'number' ? r[7] : 0, ly, varD),
      budget,
      varDBud,
      varPBud,
      _isTotal: isTotal,
    });
    if (isTotal) break;
  }
  return out;
}
function extractSubCat(ws) {
  const out = [];
  for (const r of ws.slice(1)) {
    if (!r || !r[0]) continue;
    const sub = safeStr(r[0]);
    const isTotal = sub.toLowerCase() === 'total' || sub.toLowerCase() === 'totals';
    const ly    = num(r[2]);
    const varD  = num(r[3]);
    const rawVP = typeof r[4] === 'number' ? r[4] : (isNAToken(r[4]) ? 'NA' : null);
    out.push({
      sub: isTotal ? 'Total' : sub,
      actual: num(r[1]),
      ly,
      varD,
      varP: naIfDivByZero(rawVP, ly, varD),
      ...(isTotal ? { isTotal: true } : {}),
    });
    if (isTotal) break;
  }
  return out;
}

// ── QTD (Quarter to Date) ───────────────────────────────────────────────────
// Parses one orange "QTD 2-<loc>" P&L sheet. These differ from the weekly/PTD
// flash sheets: one row per P&L line, with Actual in col B(1), Prior Year in
// col H(7) and Budget in the right-most col L(11). Returns the pieces the
// Overview (total sales) and Revenue & Channels (revenue centers + sub-cats)
// tabs need. Variance is computed vs Prior Year, except Overview's budget var.
const QTD_CENTERS = [
  ['Total Single Plate In-House Sales', 'In-Store'],
  ['Total Takeout Sales',               'Takeout'],
  ['Total Delivery Sales',              'Delivery'],
  ['Total Catering Sales',              'Catering'],
  ['Total Offsites',                    'Offsites'],
  ['Total Delivery Fee Income',         'Delivery Fee'],
  ['Total Discounts / Refunds',         'Discounts/Refunds'],
];
const QTD_SUBCATS = [
  ['delivery', 'Delivery Sales - '],
  ['pickup',   'Takeout Sales - '],
  ['catering', 'Catering Sales - '],
  ['offsites', 'Offsites - '],
];
// Group the raw P&L sub-lines into the same buckets the weekly/PTD views show
// (e.g. catering collapses to RASA Catering / EzCater / Other 3rd Parties; RASA
// website+app combine). Offsites stay per-vendor. Returns the bucket label.
const QTD_GROUP = {
  delivery: n => /rasa/i.test(n) ? 'RASA - Website & App' : /doordash/i.test(n) ? 'Doordash' : /grubhub/i.test(n) ? 'Grubhub' : /uber/i.test(n) ? 'Uber Eats' : n,
  pickup:   n => /rasa/i.test(n) ? 'RASA Website & App' : /doordash/i.test(n) ? 'DoorDash' : /grubhub/i.test(n) ? 'GrubHub' : /uber/i.test(n) ? 'Uber Eats' : /ritual/i.test(n) ? 'Ritual' : /too good/i.test(n) ? 'Too Good To Go' : n,
  catering: n => /rasa/i.test(n) ? 'RASA Catering' : /ez\s*cater/i.test(n) ? 'EzCater' : 'Other 3rd Parties',
  offsites: n => n,
};
const QTD_ORDER = {
  delivery: ['RASA - Website & App', 'Doordash', 'Grubhub', 'Uber Eats'],
  pickup:   ['RASA Website & App', 'DoorDash', 'GrubHub', 'Uber Eats', 'Ritual', 'Too Good To Go'],
  catering: ['RASA Catering', 'EzCater', 'Other 3rd Parties'],
  offsites: [],
};
const qtdVarP = (actual, ly) => (ly ? (actual - ly) / Math.abs(ly) : 'NA');

function extractQTDSheet(rows) {
  const B = 1, H = 7, L = 11;
  const norm = s => safeStr(s).replace(/\s+/g, ' ').trim();

  // Total Sales — several rows share that label; use the one whose Budget (L) is set.
  let totalSales = null;
  for (const r of rows) {
    if (r && /^total sales$/i.test(norm(r[0])) && typeof r[L] === 'number') {
      const actual = num(r[B]), ly = num(r[H]), budget = num(r[L]);
      totalSales = {
        actual, ly, budget,
        varLY:  ly ? (actual - ly) / ly : 0,
        varBud: budget ? (actual - budget) / budget : 0,
      };
      break;
    }
  }

  // Revenue centers from the section subtotal rows.
  // Col L (index 11) carries the QTD budget for each section total row.
  const centerByLabel = new Map(QTD_CENTERS.map(([lbl, c]) => [lbl.toLowerCase(), c]));
  const revCenters = [];
  for (const r of rows) {
    if (!r) continue;
    const center = centerByLabel.get(norm(r[0]).toLowerCase());
    if (!center) continue;
    const actual = num(r[B]), ly = num(r[H]);
    const budget  = typeof r[L] === 'number' ? r[L] : null;
    const varDBud = budget != null ? actual - budget : null;
    const varPBud = budget != null && budget !== 0 ? varDBud / Math.abs(budget) : null;
    revCenters.push({ center, actual, ly, varD: actual - ly, varP: qtdVarP(actual, ly), budget, varDBud, varPBud });
  }

  // The QTD P&L lists Sales Adjustments / Open App Sales Adjustments / Open
  // Tickets as standalone lines, whereas the weekly/PTD revenue-center sheets
  // bundle them into the discounts total. Fold them into Discounts/Refunds here
  // so the views stay consistent and the centers still sum to Total Sales.
  const ADJ = ['sales adjustments', 'open app sales adjustments', 'open tickets'];
  let adjA = 0, adjL = 0, adjB = 0;
  for (const r of rows) {
    if (r && ADJ.includes(norm(r[0]).toLowerCase())) {
      adjA += num(r[B]); adjL += num(r[H]);
      if (typeof r[L] === 'number') adjB += r[L];
    }
  }
  const disc = revCenters.find(rc => rc.center === 'Discounts/Refunds');
  if (disc && (adjA || adjL)) {
    disc.actual += adjA;
    disc.ly += adjL;
    disc.varD = disc.actual - disc.ly;
    disc.varP = qtdVarP(disc.actual, disc.ly);
    if (disc.budget != null) {
      disc.budget  += adjB;
      disc.varDBud  = disc.actual - disc.budget;
      disc.varPBud  = disc.budget !== 0 ? disc.varDBud / Math.abs(disc.budget) : null;
    }
  }

  // Sub-categories from the detail rows under each section, grouped into the
  // same buckets the weekly/PTD views use (skip DNU/total/sub-header lines).
  const acc = { delivery: new Map(), pickup: new Map(), catering: new Map(), offsites: new Map() };
  for (const r of rows) {
    if (!r) continue;
    const label = norm(r[0]);
    if (!label || /^dnu/i.test(label) || /^total/i.test(label)) continue;
    for (const [cat, prefix] of QTD_SUBCATS) {
      if (label.toLowerCase().startsWith(prefix.toLowerCase())) {
        const name = label.slice(prefix.length).trim();
        const bucket = QTD_GROUP[cat](name);
        if (bucket) {
          const cur = acc[cat].get(bucket) || { actual: 0, ly: 0 };
          cur.actual += num(r[B]);
          cur.ly += num(r[H]);
          acc[cat].set(bucket, cur);
        }
        break;
      }
    }
  }
  const buildCat = cat => {
    const order = QTD_ORDER[cat];
    return [...acc[cat].keys()]
      .sort((a, b) => {
        const ia = order.indexOf(a), ib = order.indexOf(b);
        return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib);
      })
      .map(sub => {
        const { actual, ly } = acc[cat].get(sub);
        return { sub, actual, ly, varD: actual - ly, varP: qtdVarP(actual, ly) };
      })
      .filter(r => r.actual !== 0 || r.ly !== 0);
  };
  const subCats = {
    delivery: buildCat('delivery'),
    pickup:   buildCat('pickup'),
    catering: buildCat('catering'),
    offsites: buildCat('offsites'),
  };

  // Append a Totals row derived from totalSales so Sales.jsx can render it.
  if (totalSales) {
    const tA = totalSales.actual, tL = totalSales.ly, tB = totalSales.budget;
    const tVarDBud = tB != null ? tA - tB : null;
    const tVarPBud = tB != null && tB !== 0 ? tVarDBud / Math.abs(tB) : null;
    revCenters.push({
      center: 'Totals', actual: tA, ly: tL,
      varD: tA - tL, varP: qtdVarP(tA, tL),
      budget: tB, varDBud: tVarDBud, varPBud: tVarPBud,
      _isTotal: true,
    });
  }

  return { totalSales, revCenters, subCats };
}

// Per-location sub-cat sheets carry the same Sub Category | Actual | LY | Var$ | Var%
// shape as the all-stores sheets. Use the same helper but normalize the "cat" field
// to match the renderer expectations.
function extractSubCatLoc(ws) {
  const out = [];
  for (const r of ws.slice(1)) {
    if (!r || !r[0]) continue;
    const cat = safeStr(r[0]);
    const isTotal = cat.toLowerCase() === 'total' || cat.toLowerCase() === 'totals';
    if (isTotal) break;
    const ly    = num(r[2]);
    const varD  = num(r[3]);
    const rawVP = typeof r[4] === 'number' ? r[4] : (isNAToken(r[4]) ? 'NA' : null);
    out.push({
      cat,
      actual: num(r[1]),
      ly,
      varD,
      varP: naIfDivByZero(rawVP, ly, varD),
    });
  }
  return out;
}

// ── parseWBR — main weekly workbook ─────────────────────────────────────────
// Locate a location's WEEKLY "Revenue Center Info" P&L sheet by content rather
// than by a brittle auto-generated tab name. Mosaic and Rockville have no stable
// sheet name — historically they landed on "Sheet52"/"Sheet53", but Excel renumbers
// these generic tabs whenever sheets are added or removed (e.g. the Jun-2026 QTD
// rollout pushed them to Sheet65/Sheet66), which silently zeroed those locations'
// weekly sub-category tables. We instead match the sheet's own header: row 0 is the
// weekly-P&L title and row 2 reads "Location: N - RASA - <Name>". Consolidated
// ("ALL"/"…_All") and condensed ("P&l_1_*", "…_Sas_*", QTD) variants are excluded;
// among the rest we take the fullest sheet that carries the standard section breakdown.
function findWeeklyPnlSheet(wb, locName) {
  const titleRe = /^Profit & Loss - Week vs Same Last Year/i;
  const locRe = new RegExp(
    'Location:\\s*\\S.*RASA\\s*-\\s*' +
      locName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*$',
    'i',
  );
  const excludeName = /^P&l?_1_|^ALL\b|consolidated|overall|_Sas_|QTD/i;
  let best = null, bestRows = -1;
  for (const name of wb.SheetNames) {
    if (excludeName.test(name)) continue;
    const rows = sheetRows(wb, name);
    if (rows.length < 500) continue; // skip condensed / empty tabs
    if (!titleRe.test(String((rows[0] || [])[0] || ''))) continue;
    if (!locRe.test(String((rows[2] || [])[0] || ''))) continue;
    // Must carry the standard P&L sections the sub-category reader walks downstream.
    if (!rows.some(r => r && String(r[0]).trim() === 'Delivery Sales')) continue;
    if (rows.length > bestRows) { best = name; bestRows = rows.length; }
  }
  return best;
}

function parseWBR(wb) {
  const W = {};
  const ws = name => sheetRows(wb, name);

  // Weekly / PTD sales + costs
  const flashW = ws('Weekly Flash Results Sales');
  if (flashW.length > 1) W.weekly = { sales: extractFlashSales(flashW) };
  const costsW = ws('Weekly Flash Results COSTS');
  if (costsW.length > 1) (W.weekly = W.weekly || {}).costs = extractFlashCosts(costsW);

  const flashP = ws('Period Flash Results Sales');
  if (flashP.length > 1) W.ptd = { sales: extractFlashSales(flashP) };
  const costsP = ws('Period Flash Results COSTS');
  if (costsP.length > 1) (W.ptd = W.ptd || {}).costs = extractFlashCosts(costsP);

  // YTD sales — per-location actual + budget from "Restaurants Revenue P vs A",
  //             LY (PY 2025) from the per-location YTD revenue-center sheet totals row.
  const rrpa = ws('Restaurants Revenue P vs A');
  if (rrpa.length > 1) {
    const nameMap = { 'NL': 'National Landing' };
    const ytd = {};
    for (const r of rrpa.slice(1)) {
      if (!r || !r[0]) continue;
      const loc = safeStr(r[0]);
      if (loc.toLowerCase() === 'total') continue;
      const mapped = nameMap[loc] || loc;
      ytd[mapped] = { actual: num(r[1]), budget: num(r[2]) };
    }
    // pull LY from each location's YTD sheet totals row
    const ytdSheets = {
      'Ballpark':         'Ballpark Revenue Center P v A',
      'MVT':              'MVT Revenue Center P v A',
      'National Landing': 'NL Revenue Center P v A',
      'Mosaic':           'Mosaic Revenue Center P v A',
      'Rockville':        'Rockville Revenue Center P v A',
    };
    for (const [loc, sh] of Object.entries(ytdSheets)) {
      const rows = ws(sh);
      for (const r of rows) {
        if (r && r[0] && /Total/i.test(safeStr(r[0]))) {
          ytd[loc] = { ...(ytd[loc] || {}), ly: num(r[3]) };
          break;
        }
      }
    }
    const order = ['Ballpark', 'Mosaic', 'MVT', 'National Landing', 'Rockville'];
    let tA = 0, tL = 0, tB = 0;
    const rows = order.map(loc => {
      const d = ytd[loc] || {};
      const a = d.actual || 0, ly = d.ly || 0, b = d.budget || 0;
      tA += a; tL += ly; tB += b;
      return {
        loc, actual: a, ly, budget: b,
        varLY: ly ? (a - ly) / ly : 0,
        varBud: b ? (a - b) / b : 0,
      };
    });
    rows.push({
      loc: 'Totals', actual: tA, ly: tL, budget: tB,
      varLY: tL ? (tA - tL) / tL : 0,
      varBud: tB ? (tA - tB) / tB : 0,
    });
    W.ytd = { sales: rows, costs: extractFlashCosts(ws('Period Flash Results COSTS')) };
  }

  // Revenue Center — weekly / PTD / YTD aggregate
  const rcW = ws('ALL - Weekly Revenue Center Cha');
  const rcP = ws('ALL - PTD  Revenue Center Chart');
  const rcY = ws('Total Revenue Center P v A');
  W.revCenter = {
    weekly: rcW.length > 1 ? extractRC(rcW) : [],
    ptd:    rcP.length > 1 ? extractRC(rcP) : [],
    ytd:    rcY.length > 1 ? extractRCYtd(rcY) : [],
  };

  // Per-location revenue center
  const locSheets = {
    'Ballpark':         ['BP- Weekly Revenue Center Chart', 'PTD BP- Weekly Revenue Center C', 'Ballpark Revenue Center P v A'],
    'MVT':              ['MVT- Weekly Revenue Center Char', 'PTD MVT- Weekly Revenue Center ', 'MVT Revenue Center P v A'],
    'National Landing': ['NL - Weekly Revenue Center Char', 'PTD NL - Weekly Revenue Center ', 'NL Revenue Center P v A'],
    'Mosaic':           ['Mosaic - Weekly Revenue Center ', 'PTD Mosaic - Weekly Revenue Cen', 'Mosaic Revenue Center P v A'],
    'Rockville':        ['Rockville - Weekly Revenue Cent', 'PTD Rockville - Weekly Revenue ', 'Rockville Revenue Center P v A'],
  };
  W.revCenterByLoc = {};
  for (const [loc, [wSh, pSh, ySh]] of Object.entries(locSheets)) {
    const wRows = ws(wSh), pRows = ws(pSh), yRows = ws(ySh);
    W.revCenterByLoc[loc] = {
      weekly: wRows.length > 1 ? extractRC(wRows) : [],
      ptd:    pRows.length > 1 ? extractRC(pRows) : [],
      ytd:    yRows.length > 1 ? extractRCYtd(yRows) : [],
    };
  }

  // Sub-categories (weekly + PTD)
  const subSheets = {
    weekly: {
      delivery: 'Weekly All Sub Categories - Del',
      pickup:   'Weekly All Sub Categories - Pic',
      catering: 'Weekly All Sub Categories - Cat',
      offsites: 'Weekly All Sub Categories - Off',
    },
    ptd: {
      delivery: 'Period All Sub Categories - Del',
      pickup:   'Period All Sub Categories - Pic',
      catering: 'Period All Sub Categories - Cat',
      offsites: 'Period All Sub Categories - Off',
    },
  };
  W.subCats = { weekly: {}, ptd: {}, ytd: {} };
  for (const period of ['weekly', 'ptd']) {
    for (const [cat, sh] of Object.entries(subSheets[period])) {
      const rows = ws(sh);
      W.subCats[period][cat] = rows.length > 1 ? extractSubCat(rows) : [];
    }
  }
  // YTD all-stores sub-cat tables live on dedicated "Total Revenue Sub Categories ..."
  // sheets (cols: Sub Category | 2026 Actual YTD | PY 2025 YTD | Var $ | Var % | Mix%
  // | Mix% PY). First 5 columns share the shape extractSubCat reads, so the extra
  // mix-% columns are simply ignored.
  const ytdSubSheets = {
    delivery: 'Total Revenue Sub Categories De',
    pickup:   'Total Revenue Sub Categories Pi',
    catering: 'Total Revenue Sub Catergories P',
    offsites: 'Total Revenue Sub Categories Of',
  };
  for (const [cat, sh] of Object.entries(ytdSubSheets)) {
    const rows = ws(sh);
    W.subCats.ytd[cat] = rows.length > 1 ? extractSubCat(rows) : [];
  }

  // Per-location sub-category breakdown (delivery/pickup/catering/offsites)
  // for weekly / ptd / ytd views. Renderer keys: DATA.subCatsByLoc[loc][viewKey][sub].
  // The XLSX exposes only YTD per-location sub-cat sheets directly; for weekly/ptd
  // we fall back to the all-stores sub-cat tables (or an empty list when the
  // sheet is missing).
  const locSubSheets = {
    'Ballpark': {
      delivery: 'Ballpark Total Revenue Sub Cate',
      pickup:   'BP Revenue Sub Categories Picku',
      catering: 'Ballpark Revenue Sub Catergorie',
      offsites: 'BP Revenue Sub Categories Offsi',
    },
    'MVT': {
      delivery: 'MVT Total Revenue Sub Categorie',
      pickup:   'MVT Revenue Sub Categories Pick',
      catering: 'MVT Revenue Sub Catergories P o',
      offsites: 'MVT Revenue Sub Categories Offs',
    },
    'National Landing': {
      delivery: 'NL Total Revenue Sub Categories',
      pickup:   'NL Revenue Sub Categories Picku',
      catering: 'NL Revenue Sub Catergories P o ',
      offsites: 'NL Revenue Sub Pop ups and Offs',
    },
    'Mosaic': {
      delivery: 'Mosaic Total Revenue Sub Catego',
      pickup:   'Mosaic Revenue Sub Categories P',
      catering: 'Mosaic Revenue Sub Catergories ',
      offsites: 'Mosaic Revenue Sub Categories O',
    },
    'Rockville': {
      delivery: 'Rockville Total Revenue Sub Cat',
      pickup:   'Rockville Revenue Sub Categorie',
      catering: 'Rockville Revenue Sub Catergori',
      offsites: 'Rockville Revenue Pop ups  P o ',
    },
  };
  // Per-location WEEKLY P&L sheets (same cols as PTD; section breakdown).
  //   col 0 = label, col 2 = Actual, col 5 = Prior Year, col 9 = Var$, col 10 = Var%
  // Mosaic and Rockville don't have explicit "<LOC> - Weekly Revenue Center Info" names —
  // they live on Sheet52 / Sheet53. Probed by inspecting the workbook.
  const locWeeklySheets = {
    'Ballpark':         'BP - Weekly Revenue Center Info',
    'MVT':              'MVT - Weekly Revenue Center Inf',
    'National Landing': 'NL - Weekly Revenue Center Info',
    'Mosaic':           'Sheet52',
    'Rockville':        'Sheet53',
  };
  const locPtdSheets = {
    'Ballpark':         'Weekly Data - BP',
    'MVT':              'Weekly Data - MVT',
    'National Landing': 'Weekly Data - NL',
    'Mosaic':           'Weekly Data - Mosaic',
    'Rockville':        'Weekly Data - Rockville',
  };

  // Section header label per sub-category in the raw P&L sheets.
  const SECTION = {
    delivery: 'Delivery Sales',
    pickup:   'Takeout Sales',
    catering: 'Catering Sales',
    offsites: 'Offsites',
  };
  // YTD-style headings → raw row-suffix(es) within that section.
  // After "<Section> - " is stripped from a raw label, the remainder must equal one of
  // these candidates (case-sensitive). The "&"-style heading sums multiple raw rows.
  const SUFFIX_MAP = {
    delivery: {
      'RASA Website & App': ['RASA Website', 'RASA App'],
      'DoorDash':            ['DoorDash'],
      'GrubHub':             ['GrubHub'],
      'Uber Eats':           ['Uber Eats'],
    },
    pickup: {
      'RASA Website & App': ['RASA Website', 'RASA App'],
      'DoorDash':            ['DoorDash'],
      'GrubHub':             ['GrubHub'],
      'Uber Eats':           ['Uber Eats'],
      'Ritual':              ['Ritual'],
      'Too Good To Go':      ['Too Good To Go'],
    },
    catering: {
      'RASA Catering':            ['RASA Website'],
      'EzCater':                  ['EZ Cater'],
      // 'Other 3rd Party Catering' is handled by the totalized row lookup below.
    },
    offsites: {
      // Match by suffix == YTD heading (e.g. "Fooda" → "Offsites - Fooda" row).
    },
  };

  // Walk a section in a per-loc P&L sheet starting at its header row.
  //   map         — { suffix → { actual, ly } } for every "<Section> - X" line
  //   totalOther  — the "Total Catering Sales - Other 3rd Party" sub-total
  //   totalSection— the "Total Offsites" (or section total) line
  function readSection(sheetRows, sectionLabel) {
    const map = {};
    let totalOther = null, totalSection = null;
    let in_ = false;
    const prefix = sectionLabel + ' - ';
    for (const r of sheetRows) {
      if (!r) continue;
      const label = safeStr(r[0] || '').trim();
      if (!label) continue;
      if (!in_) {
        if (label === sectionLabel) in_ = true;
        continue;
      }
      // Sub-total for Other 3rd Party Catering — keep walking after capturing it.
      if (/^total\s+catering\s+sales\s*-\s*other\s+3rd\s+party$/i.test(label)) {
        totalOther = { actual: num(r[2]), ly: num(r[5]) };
        continue;
      }
      // Section total — capture then stop.
      if (label === 'Total ' + sectionLabel || label.toLowerCase() === ('total ' + sectionLabel.toLowerCase())) {
        totalSection = { actual: num(r[2]), ly: num(r[5]) };
        break;
      }
      if (label.startsWith('DNU_')) continue;
      if (label === sectionLabel) continue;
      if (label.startsWith(prefix)) {
        const suffix = label.slice(prefix.length).trim();
        if (suffix) map[suffix] = { actual: num(r[2]), ly: num(r[5]) };
      }
    }
    return { map, totalOther, totalSection };
  }

  // Build the per-loc rows for a given location/sub from the section data.
  function buildSubCatLoc(sectionData, sub, ytdRows) {
    const out = [];
    for (const ytdRow of ytdRows) {
      const head = ytdRow.cat;
      let actual = 0, ly = 0, matched = false;
      if (sub === 'catering' && /other\s+3rd\s+party/i.test(head) && sectionData.totalOther) {
        // The "Total Catering Sales - Other 3rd Party" row in the raw P&L includes EzCater.
        // Subtract EzCater (= raw "Catering Sales - EZ Cater") to match the YTD breakdown convention.
        const ez = sectionData.map['EZ Cater'] || { actual: 0, ly: 0 };
        actual = sectionData.totalOther.actual - ez.actual;
        ly     = sectionData.totalOther.ly     - ez.ly;
        matched = true;
      } else if (sub === 'offsites' && /off-?site\s+pop-?ups?/i.test(head) && sectionData.totalSection) {
        actual = sectionData.totalSection.actual; ly = sectionData.totalSection.ly; matched = true;
      } else {
        const candidates = (SUFFIX_MAP[sub] && SUFFIX_MAP[sub][head]) || [head];
        for (const c of candidates) {
          const hit = sectionData.map[c];
          if (hit) { actual += hit.actual; ly += hit.ly; matched = true; }
        }
      }
      const varD = actual - ly;
      // Any divide-by-zero (including 0/0) is undefined — render NA.
      const varP = ly === 0 ? 'NA' : varD / Math.abs(ly);
      out.push({ cat: head, actual, ly, varD, varP, _missing: !matched });
    }
    return out;
  }

  W.subCatsByLoc = {};
  for (const [loc, sheets] of Object.entries(locSubSheets)) {
    const ytd = {};
    for (const [sub, sh] of Object.entries(sheets)) {
      const rows = ws(sh);
      ytd[sub] = rows.length > 1 ? extractSubCatLoc(rows) : [];
    }
    let weeklyRows = ws(locWeeklySheets[loc]) || [];
    // The hardcoded weekly tab can vanish when Excel renumbers generic "SheetNN"
    // tabs (notably Mosaic/Rockville, which have no stable name). Fall back to a
    // content-based lookup so those locations don't silently drop to $0.
    if (weeklyRows.length <= 1) {
      const resolved = findWeeklyPnlSheet(wb, loc);
      if (resolved) weeklyRows = ws(resolved) || [];
    }
    const ptdRows    = ws(locPtdSheets[loc])    || [];
    const weekly = { delivery: [], pickup: [], catering: [], offsites: [] };
    const ptd    = { delivery: [], pickup: [], catering: [], offsites: [] };
    for (const sub of ['delivery','pickup','catering','offsites']) {
      const ytdHeads = ytd[sub];
      if (!ytdHeads || !ytdHeads.length) continue;
      const wSect = readSection(weeklyRows, SECTION[sub]);
      const pSect = readSection(ptdRows,    SECTION[sub]);
      // National Landing & Rockville expose only a single "Off-Site Pop-Ups"
      // aggregate in their YTD offsites sheet, even though the weekly/PTD P&L
      // sections carry the full vendor breakdown (Fooda, Aramark, …). When that
      // happens, drive the weekly/PTD rows off the section's vendor list so these
      // locations show the same breakdown as Mosaic et al.
      let wHeads = ytdHeads, pHeads = ytdHeads;
      if (sub === 'offsites' && ytdHeads.length === 1
          && /off-?site\s+pop-?ups?/i.test(ytdHeads[0].cat || '')) {
        const wVendors = Object.keys(wSect.map);
        const pVendors = Object.keys(pSect.map);
        if (wVendors.length) wHeads = wVendors.map(cat => ({ cat }));
        if (pVendors.length) pHeads = pVendors.map(cat => ({ cat }));
      }
      weekly[sub] = buildSubCatLoc(wSect, sub, wHeads);
      ptd[sub]    = buildSubCatLoc(pSect, sub, pHeads);
    }
    W.subCatsByLoc[loc] = { weekly, ptd, ytd };
  }

  // ── QTD (Quarter to Date) ──────────────────────────────────────────────────
  // Only present in workbooks carrying the orange "QTD 2-<loc>" P&L sheets
  // (rolled out from the Jun 2026 week onward). Older weeks omit it, so the UI
  // hides the QTD toggle for them. Adds a `qtd` view to Overview (sales) and to
  // the Revenue & Channels revenue-center / sub-cat structures.
  if (wb.SheetNames.includes('QTD 2-Consolidated')) {
    const qtdSheets = {
      'Ballpark':         'QTD 2-BP',
      'Mosaic':           'QTD 2-MO',
      'MVT':              'QTD 2-MVT',
      'National Landing': 'QTD 2-NL',
      'Rockville':        'QTD 2-RV',
    };
    const order = ['Ballpark', 'Mosaic', 'MVT', 'National Landing', 'Rockville'];
    const consol = extractQTDSheet(ws('QTD 2-Consolidated'));

    const salesRows = [];
    for (const loc of order) {
      const rows = ws(qtdSheets[loc]);
      const q = rows.length ? extractQTDSheet(rows) : null;
      const t = q && q.totalSales;
      salesRows.push({
        loc,
        actual: t ? t.actual : 0,
        ly:     t ? t.ly     : 0,
        budget: t ? t.budget : 0,
        varLY:  t ? t.varLY  : 0,
        varBud: t ? t.varBud : 0,
      });
      if (W.revCenterByLoc[loc]) W.revCenterByLoc[loc].qtd = q ? q.revCenters : [];
      if (W.subCatsByLoc[loc])   W.subCatsByLoc[loc].qtd   = q ? q.subCats : { delivery: [], pickup: [], catering: [], offsites: [] };
    }
    const ct = consol.totalSales || { actual: 0, ly: 0, budget: 0, varLY: 0, varBud: 0 };
    salesRows.push({ loc: 'Totals', actual: ct.actual, ly: ct.ly, budget: ct.budget, varLY: ct.varLY, varBud: ct.varBud });

    W.qtd = { sales: salesRows };
    W.revCenter.qtd = consol.revCenters;
    W.subCats.qtd = consol.subCats;
    W.qtdAvailable = true;
  }

  // Reviews + 3rd-party reviews — two side-by-side tables, split when r[0]==='3rd Party Metrics'
  const revRows = ws('Weekly in-store leadership metr');
  if (revRows.length > 1) {
    const isW = [], is90 = [], tpW = [], tp90 = [];
    let in3p = false;
    for (let i = 1; i < revRows.length; i++) {
      const r = revRows[i]; if (!r) continue;
      if (r[0] === '3rd Party Metrics') { in3p = true; continue; }
      if (!in3p) {
        if (r[1] && typeof r[2] === 'number')
          isW.push({ loc: safeStr(r[1]), reviews: r[2]||0, rating: r[3]??null, s5: r[4]??null, s4: r[5]??null, s3: r[6]??null, s2: r[7]??null, s1: r[8]??null, yelp: r[9], yelpN: r[10]??null, google: r[11], gNum: r[12]??null });
        if (r[15] && typeof r[16] === 'number')
          is90.push({ loc: safeStr(r[15]), reviews: r[16]||0, rating: r[17]??null, s5: r[18]??null, s4: r[19]??null, s3: r[20]??null, s2: r[21]??null, s1: r[22]??null, yelp: r[23], yelpN: r[24]??null, google: r[25], gNum: r[26]??null });
      } else {
        if (r[1] && typeof r[2] === 'number')
          tpW.push({ loc: safeStr(r[1]), reviews: r[2]||0, rating: r[3]??null, s5: r[4]??null, s4: r[5]??null, s3: r[6]??null, s2: r[7]??null, s1: r[8]??null, ue: r[9], dd: r[10], gh: r[11], errRate: r[12]??null });
        if (r[15] && typeof r[16] === 'number')
          tp90.push({ loc: safeStr(r[15]), reviews: r[16]||0, rating: r[17]??null, s5: r[18]??null, s4: r[19]??null, s3: r[20]??null, s2: r[21]??null, s1: r[22]??null, ue: r[23], dd: r[24], gh: r[25], errRate: r[26]??null });
      }
    }
    // 30-day In-Store (Yelp+Google) and 3PD tables sit to the right (cols AC–AO,
    // idx 28–40), with "In-Store Metrics" / "3rd Party Metrics" labels in col 28.
    // Rolled out from Jun 2026; absent in older weeks → empty → UI hides the filter.
    const is30 = [], tp30 = [];
    let in3p30 = false;
    for (let i = 0; i < revRows.length; i++) {
      const r = revRows[i]; if (!r) continue;
      if (safeStr(r[28]).trim() === '3rd Party Metrics') { in3p30 = true; continue; }
      if (r[29] == null || typeof r[30] !== 'number') continue; // need a location + numeric #reviews
      if (!in3p30) {
        is30.push({ loc: safeStr(r[29]), reviews: num(r[30]), rating: r[31]??null, s5: r[32]??null, s4: r[33]??null, s3: r[34]??null, s2: r[35]??null, s1: r[36]??null, yelp: r[37], yelpN: r[38]??null, google: r[39], gNum: r[40]??null });
      } else {
        tp30.push({ loc: safeStr(r[29]), reviews: num(r[30]), rating: r[31]??null, s5: r[32]??null, s4: r[33]??null, s3: r[34]??null, s2: r[35]??null, s1: r[36]??null, ue: r[37], dd: r[38], gh: r[39], errRate: r[40]??null });
      }
    }
    W.reviews = {
      instore:    { weekly: isW, ninety: is90, day30: is30 },
      thirdparty: { weekly: tpW, ninety: tp90, day30: tp30 },
    };
  }

  // 3PD Reporting (UE in cols 0-11, DD in cols 13-24)
  // DD perf table has no error-rate column — col 18 carries the top-complaint string.
  // UE ops has 12 columns: errRate / missItems / custErr / wrongOrder / qualIssues /
  // avgWait / avoidWait / avoidOrders / avoidCancel / avoidCancelRate / menuCvr / uptime.
  // DD ads grouped header: 4 Promo cols (Sales / Spend / Orders / ROAS) at 14-17,
  // 3 Promo extras (Impressions / Clicks / CTR) at 18-20, 4 Sponsored cols at 21-24,
  // overall orders / cancel rate trailing.
  const tpd = ws('3PD Reporting - UE & DD');
  if (tpd.length > 19) {
    const uePerf=[], ueOps=[], ueAds=[], ddPerf=[], ddOps=[], ddAds=[], ddRatings=[];
    // Performance overview — rows 3-8.
    // UE block cols 0-6: loc | sales | payout | orders | aov | rating | top complaint.
    // DD block cols 13-19: loc | sales | payout | orders | aov | top complaints (no rating).
    // DD ratings block cols 21-25: loc | love% | like% | dislike% | total reviews.
    for (let i = 3; i <= 8; i++) {
      const r = tpd[i]; if (!r || !r[0]) continue;
      uePerf.push({ loc: safeStr(r[0]), sales: num(r[1]), payout: num(r[2]), orders: num(r[3]), aov: num(r[4]), rating: num(r[5]), complaint: safeStr(r[6]) });
      if (r[13]) ddPerf.push({
        loc: safeStr(r[13]),
        sales:  num(r[14]),
        payout: num(r[15]),
        orders: num(r[16]),
        aov:    num(r[17]),
        complaint: safeStr(r[18] != null ? r[18] : '-'),    // DD perf top complaint sits at col 18
      });
      if (r[20]) {
        const lv = num(r[21]), lk = num(r[22]), dl = num(r[23]);
        ddRatings.push({
          loc: safeStr(r[20]),
          lovePct: lv, likePct: lk, dislikePct: dl,
          totalReviews: num(r[24]),
          // Mirror HTML DATA shape (short keys) too.
          love: lv, like: lk, dislike: dl, reviews: num(r[24]),
        });
      }
    }
    // Operations overview — rows 14-19.
    // UE cols 0-11: loc | errRate | missItems | custErr | wrongOrder | qualIssues |
    //               avgWait | avoidWait | avoidOrders | avoidCancelRate | menuCvr | uptime.
    // DD cols 13-23: loc | errRate | missItems | ingErr | missSide | incSize |
    //                dashWait | avoidWait | avoidCancel | avoidCancelRate | uptime.
    for (let i = 14; i <= 19; i++) {
      const r = tpd[i]; if (!r || !r[0]) continue;
      ueOps.push({
        loc: safeStr(r[0]),
        errRate: num(r[1]), missItems: num(r[2]), custErr: num(r[3]),
        wrongOrder: num(r[4]), qualIssues: num(r[5]),
        avgWait: safeStr(r[6]), avoidWait: safeStr(r[7]),
        avoidOrders: num(r[8]),
        avoidCancelRate: num(r[9]),
        menuCvr: num(r[10]),
        uptime: num(r[11]),
      });
      if (r[13]) ddOps.push({
        loc: safeStr(r[13]),
        errRate: num(r[14]), missItems: num(r[15]),
        ingErr: num(r[16]), missSide: num(r[17]), incSize: num(r[18]),
        dashWait: safeStr(r[19]), avoidWait: safeStr(r[20]),
        avoidCancel: num(r[21]),
        avoidCancelRate: num(r[22]),
        uptime: num(r[23]),
      });
    }
    // UE ads — rows 25-29 (campaigns incl. "All Campaign" total).
    for (let i = 25; i <= 29; i++) {
      const r = tpd[i]; if (!r) continue;
      if (r[0]) ueAds.push({ campaign: safeStr(r[0]), sales: num(r[1]), spend: num(r[2]), roas: num(r[3]), impressions: num(r[4]), clicks: num(r[5]), ctr: num(r[6]), orders: num(r[7]), cvr: num(r[8]), cpo: num(r[9]), newCust: num(r[10]) });
    }
    // DD ads — rows 24-29 (Rockville first row, then Ballpark/MVT/NL/Mosaic/All Stores).
    // Header (rows 22-23, cols 14-25): Promo Sales / Promo Spend / Promo Orders / Promo ROAS /
    // Impressions / Clicks / CTR / Sponsored Sales / Sponsored Spend / Sponsored Orders /
    // Sponsored ROAS / Overall Orders.
    for (let i = 24; i <= 29; i++) {
      const r = tpd[i]; if (!r) continue;
      if (r[13]) ddAds.push({
        loc: safeStr(r[13]),
        promoSales:    num(r[14]),
        promoSpend:    num(r[15]),
        promoOrders:   num(r[16]),
        promoROAS:     r[17] ?? null,
        impressions:   num(r[18]),
        clicks:        num(r[19]),
        ctr:           r[20] ?? null,
        sponsorSales:  num(r[21]),
        sponsorSpend:  num(r[22]),
        sponsorOrders: num(r[23]),
        sponsorROAS:   r[24] ?? null,
        overallOrders: num(r[25]),
        cancelRate:    num(r[26]),
      });
    }
    // Verbatim header text straight from the sheet (newlines/extra spaces collapsed),
    // so the 3PD tables label their columns exactly as the source does.
    const clean = v => (v == null ? '' : String(v).replace(/\s+/g, ' ').trim());
    const r2 = tpd[2] || [], r13 = tpd[13] || [], r22 = tpd[22] || [], r23 = tpd[23] || [], r24 = tpd[24] || [];
    const ueHeaders = {
      perf: r2.slice(0, 7).map(clean),    // cols 0-6
      ops:  r13.slice(0, 12).map(clean),  // cols 0-11
      ads:  r24.slice(0, 11).map(clean),  // cols 0-10
    };
    const ddHeaders = {
      perf:    r2.slice(13, 19).map(clean),   // cols 13-18
      ratings: r2.slice(20, 25).map(clean),   // cols 20-24
      ops:     r13.slice(13, 24).map(clean),  // cols 13-23
      adsGroup: { loc: clean(r22[13]), promo: clean(r22[14]), sponsor: clean(r22[18]), overall: clean(r22[25]) },
      ads:     r23.slice(14, 25).map(clean),  // cols 14-24 (11 sub-headers)
    };
    W.ue = { perf: uePerf, ops: ueOps, ads: ueAds, headers: ueHeaders };
    W.dd = { perf: ddPerf, ops: ddOps, ads: ddAds, ratings: ddRatings, headers: ddHeaders };
  }

  // Customer Insights — Locations (Weekly/Monthly/90 days) + Acquisition + Onboarding
  const ciRows = ws('Customer Insights');
  if (ciRows.length > 2) {
    const extractLocSection = (startIdx, c0 = 0) => {
      const out = [];
      for (let i = startIdx; i < Math.min(startIdx + 8, ciRows.length); i++) {
        const r = ciRows[i]; if (!r || r[c0] == null) continue;
        const loc = safeStr(r[c0]);
        const skip = new Set(['weekly','monthly','monthly ','90 days','ytd','location','customer acquisition','customer onboarding','customer onboarding  & engagement','customer onboarding & engagement','first order month','first order week','onboarding completion month','onboarding completion week']);
        if (skip.has(loc.toLowerCase())) continue;
        if (typeof r[c0 + 1] !== 'number') continue;
        out.push({
          loc,
          orders: num(r[c0 + 1]),
          ordersGrowth: num(r[c0 + 2]),
          aov: num(r[c0 + 3]),
          guests: num(r[c0 + 4]),
          guestsGrowth: num(r[c0 + 5]),
          newGuests: num(r[c0 + 6]),
          newGuestsGrowth: num(r[c0 + 7]),
        });
      }
      return out;
    };
    // YTD locations table (cols P–W = idx 15–22), marked by a "YTD" cell in col P.
    // Rolled out from Jun 2026; older weeks lack it → empty → UI hides the filter.
    let ytdLoc = [];
    for (let i = 0; i < Math.min(30, ciRows.length); i++) {
      if (ciRows[i] && safeStr(ciRows[i][15]).trim().toLowerCase() === 'ytd') {
        ytdLoc = extractLocSection(i + 2, 15); // data starts two rows below the "YTD" label
        break;
      }
    }
    const locations = {
      weekly:  { curr: extractLocSection(2), prev: [] },
      monthly: extractLocSection(12),
      ninety:  extractLocSection(23),
      ytd:     ytdLoc,
    };

    // Acquisition — monthly 12-month 90-day (cols 0-6),
    //               weekly 3-month 30-day (cols 8-14),
    //               monthly 12-month 30-day (cols 16-22).
    const monthly90d = [];
    for (let i = 36; i < Math.min(50, ciRows.length); i++) {
      const r = ciRows[i]; if (!r) continue;
      if (r[0] instanceof Date) {
        monthly90d.push({
          period: r[0].toLocaleString('en-US', { month: 'short', year: 'numeric' }),
          newGuests: num(r[1]), perLoc: num(r[2]),
          returnRate90: num(r[3]), avgOrders: num(r[4]),
          aov: num(r[5]), spend90: num(r[6]),
        });
      } else if (safeStr(r[0]).toLowerCase() === 'average') {
        monthly90d.push({ period: 'Average', newGuests: num(r[1]), perLoc: num(r[2]), returnRate90: num(r[3]), avgOrders: num(r[4]), aov: num(r[5]), spend90: num(r[6]) });
      }
    }
    const weekly30d = [];
    for (let i = 36; i < Math.min(48, ciRows.length); i++) {
      const r = ciRows[i]; if (!r || r.length <= 14 || r[8] == null) continue;
      let period, year;
      if (r[8] instanceof Date) { period = r[8].toLocaleString('en-US', { month: 'short', day: 'numeric' }); year = r[8].getFullYear(); }
      else if (safeStr(r[8]).toLowerCase() === 'average') period = 'Average';
      else continue;
      weekly30d.push({
        period, year,
        newGuests: num(r[9]), perLoc: num(r[10]), returnRate30: num(r[11]),
        avgOrders: num(r[12]), aov: num(r[13]), spend30: num(r[14]),
      });
    }
    const monthly30d = [];
    for (let i = 36; i < Math.min(50, ciRows.length); i++) {
      const r = ciRows[i]; if (!r || r.length <= 22 || r[16] == null) continue;
      if (r[16] instanceof Date) {
        monthly30d.push({
          period: r[16].toLocaleString('en-US', { month: 'short', year: 'numeric' }),
          newGuests: num(r[17]), perLoc: num(r[18]),
          returnRate30: num(r[19]), avgOrders: num(r[20]),
          aov: num(r[21]), spend30: num(r[22]),
        });
      } else if (safeStr(r[16]).toLowerCase() === 'average') {
        monthly30d.push({ period: 'Average', newGuests: num(r[17]), perLoc: num(r[18]), returnRate30: num(r[19]), avgOrders: num(r[20]), aov: num(r[21]), spend30: num(r[22]) });
      }
    }

    // Onboarding — monthly 12-month (cols 0-6) + weekly 3-month (cols 8-11)
    const onbMonthly = [];
    for (let i = 55; i < Math.min(68, ciRows.length); i++) {
      const r = ciRows[i]; if (!r || !(r[0] instanceof Date)) continue;
      onbMonthly.push({
        period: r[0].toLocaleString('en-US', { month: 'short', year: 'numeric' }),
        onboarded: num(r[1]), perLoc: num(r[2]),
        latency: num(r[3]), aov: num(r[4]),
        spend: num(r[5]), engaged: num(r[6]),
      });
    }
    if (ciRows[68] && safeStr(ciRows[68][0]).toLowerCase() === 'average') {
      const r = ciRows[68];
      onbMonthly.push({ period: 'Average', onboarded: num(r[1]), perLoc: num(r[2]), latency: num(r[3]), aov: num(r[4]), spend: num(r[5]), engaged: num(r[6]) });
    }
    // Weekly onboarding table (Customer Insights cols I-O = idx 8-14):
    // Week | Onboarded | Per Loc | Latency | AOV | Spend | Engaged
    const onbWeekly = [];
    for (let i = 53; i < Math.min(66, ciRows.length); i++) {
      const r = ciRows[i]; if (!r || r.length <= 11 || r[8] == null) continue;
      let period;
      if (r[8] instanceof Date) period = 'Week of ' + r[8].toLocaleString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
      else if (safeStr(r[8]).toLowerCase() === 'average') period = 'Average';
      else continue;
      onbWeekly.push({
        period,
        onboarded: num(r[9]), perLoc: num(r[10]), latency: num(r[11]),
        aov: num(r[12]), spend: num(r[13]), engaged: num(r[14]),
      });
    }

    // Simple-shape weekly sales row used by some legacy KPI panels
    const weeklyCurrSimple = locations.weekly.curr.map(r => ({
      loc: r.loc, orders: r.orders, aov: r.aov, guests: r.guests, newGuests: r.newGuests,
      sales: Math.round(r.orders * r.aov),
    }));
    // Comparison block (cols 15-31, rows 3-7) carries the previous-period totals
    // used as the WoW baseline (orders / guests / new guests).
    const weeklyPrevSimple = [];
    for (let i = 3; i < Math.min(9, ciRows.length); i++) {
      const r = ciRows[i]; if (!r || !r[15]) continue;
      const loc = safeStr(r[15]);
      if (!loc || loc === 'Location' || loc.toLowerCase() === 'current') continue;
      if (typeof r[17] !== 'number') continue;
      weeklyPrevSimple.push({
        loc,
        orders: num(r[17]),                       // comparison total orders
        aov:    num(r[21]),
        guests: num(r[25]),
        newGuests: num(r[29]),
        sales:  Math.round(num(r[17]) * num(r[21])),
      });
    }
    locations.weekly.prev = weeklyPrevSimple;
    W.bikky = {
      locations,
      acquisition: { monthly90d, monthly30d, weekly30d },
      onboarding: { monthly: onbMonthly, weekly: onbWeekly },
      weekly: { curr: weeklyCurrSimple, prev: weeklyPrevSimple },
    };
  }

  // Catering email/flows + EzCater Paid Ads — InputsOutputs Catering sheet.
  // Layout shifts between weeks (Week 11 had emails starting at col F; Week 18
  // moved them to col A; the EzCater block likewise moved Q→L). We locate each
  // table by scanning for its title row, then read columns relative to that anchor.
  const catRows = ws('InputsOutputs Catering');
  if (catRows.length > 5) {
    function findAnchor(re) {
      for (let r = 0; r < catRows.length; r++) {
        const row = catRows[r]; if (!row) continue;
        for (let c = 0; c < row.length; c++) {
          if (typeof row[c] === 'string' && re.test(row[c])) return { r, c };
        }
      }
      return null;
    }
    function readEmailTable(anchor) {
      const out = [];
      if (!anchor) return out;
      for (let i = anchor.r + 2; i < catRows.length; i++) {
        const r = catRows[i]; if (!r) continue;
        const campaign = safeStr(r[anchor.c] || '');
        if (!campaign) break;
        if (typeof r[anchor.c + 1] !== 'number') continue;
        out.push({
          campaign,
          sent:      num(r[anchor.c + 1]),
          delivered: num(r[anchor.c + 2]),
          bounced:   num(r[anchor.c + 3]),
          spam:      num(r[anchor.c + 4]),
          unsub:     Math.round(num(r[anchor.c + 5])),
          opened:    num(r[anchor.c + 6]),
          clicked:   num(r[anchor.c + 7]),
          ordered:   num(r[anchor.c + 8]),
          revenue:   num(r[anchor.c + 9]),
        });
      }
      return out;
    }
    function readFlowTable(anchor) {
      const out = [];
      if (!anchor) return out;
      for (let i = anchor.r + 2; i < catRows.length; i++) {
        const r = catRows[i]; if (!r) continue;
        const flow = safeStr(r[anchor.c] || '');
        if (!flow) break;
        if (typeof r[anchor.c + 1] !== 'number') continue;
        out.push({
          flow,
          delivered: num(r[anchor.c + 1]),
          opened:    num(r[anchor.c + 2]),
          clicked:   num(r[anchor.c + 3]),
          revenue:   num(r[anchor.c + 4]),
        });
      }
      return out;
    }
    function readEzcaterTable(anchor) {
      const out = [];
      if (!anchor) return out;
      for (let i = anchor.r + 2; i < catRows.length; i++) {
        const r = catRows[i]; if (!r) continue;
        const loc = safeStr(r[anchor.c] || '');
        if (!loc || loc === 'Restaurant') break;
        if (typeof r[anchor.c + 1] !== 'number') continue;
        out.push({
          loc,
          views:        num(r[anchor.c + 1]),
          clicks:       num(r[anchor.c + 2]),
          cvr:          num(r[anchor.c + 3]),
          orders:       num(r[anchor.c + 4]),
          spend:        num(r[anchor.c + 5]),
          adSpend:      num(r[anchor.c + 5]),
          sales:        num(r[anchor.c + 6]),
          roas:         num(r[anchor.c + 7]),
          custNew:      num(r[anchor.c + 8]),
          custExisting: num(r[anchor.c + 9]),
          custLapsed:   num(r[anchor.c + 10]),
          isTotal:      loc.toLowerCase() === 'total',
        });
      }
      return out;
    }

    function findAnchorAfter(re, afterRow) {
      for (let r = afterRow; r < catRows.length; r++) {
        const row = catRows[r]; if (!row) continue;
        for (let c = 0; c < row.length; c++) {
          if (typeof row[c] === 'string' && re.test(row[c])) return { r, c };
        }
      }
      return null;
    }

    const em30 = readEmailTable(findAnchor(/^Email Campaigns - Last 30/i));
    const em90 = readEmailTable(findAnchor(/^Email Campaigns - Last 90/i));
    const fl30 = readFlowTable( findAnchor(/^Flows - Last 30/i));
    const fl90 = readFlowTable( findAnchor(/^Flows - Last 90/i));
    const ezc30Anchor = findAnchor(/^Paid Ads on EZCater/i);
    const ezc90Anchor = ezc30Anchor ? findAnchorAfter(/^Paid Ads on EZCater/i, ezc30Anchor.r + 1) : null;
    const ezcaterAds    = readEzcaterTable(ezc30Anchor);
    const ezcaterAds90d = readEzcaterTable(ezc90Anchor);
    W.catering = { email30d: em30, email90d: em90, flows30d: fl30, flows90d: fl90, ezcaterAds, ezcaterAds90d };
  }

  // Pre-calculated weekly/PTD budgets per RC from Budget Import sheets, keyed by fiscal week.
  const budgetByWeek = parseBudgetSheet(wb, 'Budget Import - 2026');
  if (budgetByWeek) W.rcBudgetByWeek = budgetByWeek;

  const budgetByWeekLoc = {};
  for (const [loc, sh] of Object.entries(LOC_BUDGET_SHEETS)) {
    const d = parseBudgetSheet(wb, sh);
    if (d) budgetByWeekLoc[loc] = d;
  }
  if (Object.keys(budgetByWeekLoc).length) W.rcBudgetByWeekLoc = budgetByWeekLoc;

  return W;
}

// ── parseLoyalty ────────────────────────────────────────────────────────────
function parseLoyalty(wb) {
  const W = {};
  const ws = name => sheetRows(wb, name);

  // Lifecycle WoW + MoM. Some week-folders put WoW at cols 0..4 and MoM at cols 11..14;
  // others put WoW at cols 2..6 and MoM at cols 13..16. Detect the offset from the
  // header row that contains the literal "Metric" cell.
  const lc = ws('Lifecycle - Table');
  if (lc.length > 2) {
    const wow = [], mom = [];
    // Find the metric/header row, defaulting to row index 3.
    let wowMetricCol = 0, momMetricCol = 11;
    let wowHeaderRow = 3, momHeaderRow = 3;
    for (let i = 0; i < Math.min(lc.length, 6); i++) {
      const r = lc[i]; if (!r) continue;
      for (let c = 0; c < r.length; c++) {
        if (safeStr(r[c]).toLowerCase() === 'metric') {
          if (c <= 6) { wowMetricCol = c; wowHeaderRow = i; }
          else        { momMetricCol = c; momHeaderRow = i; }
        }
      }
    }
    // Pull the column header labels from the same row that contains "Metric"
    // so the table title and per-column headers reflect the actual input
    // (e.g. "Week of May 18" / "Week of May 11" / "Var (%)" / "YTD").
    const wowHeadRow = lc[wowHeaderRow] || [];
    const momHeadRow = lc[momHeaderRow] || [];
    const wowHeaders = {
      metric: safeStr(wowHeadRow[wowMetricCol]     || 'Metric'),
      curr:   safeStr(wowHeadRow[wowMetricCol + 1] || 'Current'),
      prev:   safeStr(wowHeadRow[wowMetricCol + 2] || 'Previous'),
      var:    safeStr(wowHeadRow[wowMetricCol + 3] || 'Var (%)'),
      ytd:    safeStr(wowHeadRow[wowMetricCol + 4] || 'YTD'),
    };
    const momHeaders = {
      metric: safeStr(momHeadRow[momMetricCol]     || 'Metric'),
      mar:    safeStr(momHeadRow[momMetricCol + 1] || 'Previous Month'),
      apr:    safeStr(momHeadRow[momMetricCol + 2] || 'Latest Month'),
      var:    safeStr(momHeadRow[momMetricCol + 3] || 'Var (%)'),
    };
    for (let i = 2; i < lc.length; i++) {
      const r = lc[i]; if (!r) continue;
      const mw = safeStr(r[wowMetricCol] || '');
      // Accept either a numeric curr or a row whose curr/prev are present (capture
      // "WoW Change in Engaged Customers" which has all-numeric values).
      if (mw && mw !== 'Metric' && (typeof r[wowMetricCol + 1] === 'number' || typeof r[wowMetricCol + 2] === 'number')) {
        wow.push({
          metric: mw,
          curr: num(r[wowMetricCol + 1]),
          prev: num(r[wowMetricCol + 2]),
          var:  num(r[wowMetricCol + 3]),
          ytd:  r[wowMetricCol + 4] != null ? r[wowMetricCol + 4] : '-',
        });
      }
      const mm = safeStr(r[momMetricCol] || '');
      if (mm && mm !== 'Metric' && typeof r[momMetricCol + 1] === 'number') {
        mom.push({
          metric: mm,
          mar: num(r[momMetricCol + 1]),
          apr: num(r[momMetricCol + 2]),
          var: num(r[momMetricCol + 3]),
        });
      }
    }
    W.lifecycle = { wow, mom, wowHeaders, momHeaders };
  }

  // Trend (WoW 2026): rows 3 dates, 5 sales, 36 signups, 37 app dl, 38 new cust, 39 activated, 40 engaged
  const wow = ws('WoW 2026');
  if (wow.length > 40) {
    const dateRow = wow[3] || [], salesRow = wow[5] || [],
          sigRow = wow[36] || [], appRow = wow[37] || [], ncRow = wow[38] || [],
          actRow = wow[39] || [], engRow = wow[40] || [];
    const weeks=[], sales=[], signups=[], appDl=[], newCust=[], activated=[], engaged=[];
    for (let c = 1; c < salesRow.length; c++) {
      const sv = salesRow[c];
      if (typeof sv === 'number' && sv > 0) {
        const d = dateRow[c];
        weeks.push(d instanceof Date
          ? d.toLocaleString('en-US', { month: 'short', day: 'numeric' })
          : 'W' + c);
        sales.push(Math.round(sv));
        signups.push(num(sigRow[c]));
        appDl.push(num(appRow[c]));
        newCust.push(num(ncRow[c]));
        activated.push(num(actRow[c]));
        engaged.push(num(engRow[c]));
      }
    }
    const last8 = a => a.slice(-8);
    W.trend = {
      weeks: last8(weeks), totalSales: last8(sales),
      signups: last8(signups), appDownloads: last8(appDl),
      newCust: last8(newCust), activated: last8(activated), engaged: last8(engaged),
    };
  }

  // Sales by location — read the consolidated breakdown from "Instore Orders - Tables".
  // Layout in rows 69-74 of the sheet (JS indices 68-73):
  //   D=loc | E=inStoreOrders | F=inStoreSales | G=digitalOrders | H=digitalSales
  //   I=totalOrders | J=totalSales
  // Data row 74 is "Grand Total".
  const inSt = ws('Instore Orders - Tables');
  const dg   = ws('Digital Orders - Tables');
  const renameLoc = (loc) => loc === 'Mount Vernon Triangle (DC)' ? 'Mt Vernon Triangle' : loc;
  const salesRows = [];
  for (let i = 68; i <= 73; i++) {
    const r = inSt[i]; if (!r) continue;
    const loc = safeStr(r[3] || '');
    if (!loc) continue;
    salesRows.push({
      loc: renameLoc(loc),
      inStoreOrders: num(r[4]),
      inStoreSales:  Math.round(num(r[5])),
      digitalOrders: num(r[6]),
      digitalSales:  Math.round(num(r[7])),
      totalOrders:   num(r[8]),
      totalSales:    Math.round(num(r[9])),
    });
  }
  W.salesByLoc = salesRows;

  // Order method + platform + app/web breakdowns from Digital Orders
  if (dg.length > 28) {
    // Weekly order method / platform — rows 5-7, cols 62/63/64 (method) and 67/68/69 (platform).
    // 28-day order method / platform — same column layout but rows 26-28.
    const omW = [], plW = [], om28 = [], pl28 = [];
    const omLabels = { DELIVERY: 'Delivery', PICKUP: 'Pickup' };
    const readOM = (range, into) => {
      for (let i = range[0]; i <= range[1]; i++) {
        const r = dg[i]; if (!r) continue;
        const method = safeStr(r[62] || '');
        if (method && method !== 'Order Method' && method !== 'Source: Orders')
          into.om.push({ method: omLabels[method] || method, orders: num(r[63]), sales: Math.round(num(r[64])) });
        const platform = safeStr(r[67] || '');
        if (platform && platform !== 'Platform' && platform !== 'Source: Orders')
          into.pl.push({ platform, orders: num(r[68]), sales: Math.round(num(r[69])) });
      }
    };
    readOM([5, 7],   { om: omW,  pl: plW  });
    readOM([26, 28], { om: om28, pl: pl28 });
    // Weekly app/web by location (rows 6-11, cols 42-48). Last row is "Grand Total".
    const aw7 = [];
    for (let i = 6; i <= 11; i++) {
      const r = dg[i]; if (!r) continue;
      const loc = safeStr(r[42] || ''); if (!loc) continue;
      aw7.push({ loc: renameLoc(loc), appSales: Math.round(num(r[43])), appOrders: num(r[44]), webSales: Math.round(num(r[45])), webOrders: num(r[46]), totalOrders: num(r[47]), totalSales: Math.round(num(r[48])) });
    }
    // 28-day app/web by location — rows 6-11, cols 52-58 (10 cols right of the 7-day block).
    // Sheet header at row 3 col 52: "LAST 28 DAYS LOYALTY SALES BY LOCATION - APP & WEB BREAKDOWN".
    const aw28 = [];
    for (let i = 6; i <= 11; i++) {
      const r = dg[i]; if (!r) continue;
      const loc = safeStr(r[52] || ''); if (!loc) continue;
      aw28.push({ loc: renameLoc(loc), appSales: Math.round(num(r[53])), appOrders: num(r[54]), webSales: Math.round(num(r[55])), webOrders: num(r[56]), totalOrders: num(r[57]), totalSales: Math.round(num(r[58])) });
    }
    W.weeklyOrderMethod        = omW.length ? omW : om28;
    W.weeklyPlatform           = plW.length ? plW : pl28;
    W.weeklyAppWeb             = aw7;
    W.twentyEightDayOrderMethod = om28;
    W.twentyEightDayPlatform    = pl28;
    W.twentyEightDayAppWeb     = aw28;

    // Weekly Disc vs Non-Disc — Digital Orders - Tables, data rows 69-74 (JS 68-73).
    //   col 45 = location, 46/47 = disc orders/sales, 48/49 = non-disc orders/sales,
    //   50/51 = total orders/sales. Last row is "Grand Total".
    const wDisc = [];
    for (let i = 68; i <= 73; i++) {
      const r = dg[i]; if (!r) continue;
      const loc = safeStr(r[45] || ''); if (!loc) continue;
      wDisc.push({
        loc: renameLoc(loc),
        discOrders:    num(r[46]),
        discSales:     Math.round(num(r[47])),
        nonDiscOrders: num(r[48]),
        nonDiscSales:  Math.round(num(r[49])),
        totalOrders:   num(r[50]),
        totalSales:    Math.round(num(r[51])),
      });
    }
    W.weeklyDiscounted = wDisc;

    // 28-day Disc vs Non-Disc — same sheet, same row range, cols 53-59 (shifted 8 cols).
    const t28Disc = [];
    for (let i = 68; i <= 73; i++) {
      const r = dg[i]; if (!r) continue;
      const loc = safeStr(r[53] || ''); if (!loc) continue;
      t28Disc.push({
        loc: renameLoc(loc),
        discOrders:    num(r[54]),
        discSales:     Math.round(num(r[55])),
        nonDiscOrders: num(r[56]),
        nonDiscSales:  Math.round(num(r[57])),
        totalOrders:   num(r[58]),
        totalSales:    Math.round(num(r[59])),
      });
    }
    W.twentyEightDayDiscounted = t28Disc;
  }

  // Weekly In-Store Loyalty Orders (top block of Instore Orders - Tables — rows 4-9, cols 0-2).
  if (inSt.length > 4) {
    const wIS = [];
    for (let i = 4; i <= 9; i++) {
      const r = inSt[i]; if (!r) continue;
      const loc = safeStr(r[0] || ''); if (!loc) continue;
      const disp = loc === 'Mount Vernon Triangle (DC)' ? 'Mt Vernon Triangle' : loc;
      if (typeof r[1] !== 'number') continue;
      wIS.push({ loc: disp, orders: num(r[1]), sales: Math.round(num(r[2])) });
    }
    if (wIS.length) W.weeklyInStore = wIS;
  }

  // SMS WoW + Email 30d + Email 90d (from SMS - Table)
  const sms = ws('SMS - Table');
  let loyMarketing = null;
  if (sms.length > 18) {
    const smsWoW = [];
    let smsCols = null;
    for (let i = 19; i <= 29; i++) {
      const r = sms[i]; if (!r) continue;
      const metric = safeStr(r[1] || '');
      if (metric === 'Metric') {
        smsCols = { curr: safeStr(r[2] || ''), prev: safeStr(r[3] || '') };
        continue;
      }
      if (!metric || r[2] == null) continue;
      const curr = r[2], prev = r[3], v = r[4];
      smsWoW.push({
        metric,
        curr: typeof curr === 'number' ? Math.round(curr*100)/100 : num(curr),
        prev: typeof prev === 'number' ? Math.round(prev*100)/100 : num(prev),
        var: num(v),
      });
    }

    // Scan ALL cells for an email section header matching the pattern.
    // Returns { anchorRow, anchorCol } or null.
    function findEmailAnchor(pattern) {
      for (let i = 0; i < sms.length; i++) {
        const r = sms[i]; if (!r) continue;
        for (let c = 0; c < r.length; c++) {
          if (pattern.test(safeStr(r[c] || ''))) return { anchorRow: i, anchorCol: c };
        }
      }
      return null;
    }

    // Pull a whole email table starting from the given anchor cell.
    // Row anchor+1 = column headers; following rows = data until blank first-cell.
    function parseEmailSection(pattern) {
      const anchor = findEmailAnchor(pattern);
      if (!anchor) return null;
      const { anchorRow, anchorCol } = anchor;

      // Derive platform name from anchor cell text, e.g. "(Open)" → "Open"
      const anchorText = safeStr(sms[anchorRow]?.[anchorCol] || '');
      const platformMatch = anchorText.match(/\(([^)]+)\)\s*$/);
      const platform = platformMatch ? platformMatch[1] : 'Klaviyo';

      // Column-header row
      const hdrRow = sms[anchorRow + 1] || [];
      const headers = [];
      for (let c = anchorCol; c < hdrRow.length; c++) {
        const h = safeStr(hdrRow[c] || '');
        if (!h) break;
        headers.push(h);
      }
      if (!headers.length) return null;

      // Data rows
      const rows = [];
      for (let i = anchorRow + 2; i < Math.min(anchorRow + 60, sms.length); i++) {
        const r = sms[i]; if (!r) continue;
        const firstCell = safeStr(r[anchorCol] || '');
        if (!firstCell) continue;
        if (/email\s+campaigns?/i.test(firstCell)) break;
        const cells = headers.map((_, hi) => {
          const v = r[anchorCol + hi];
          return (v === '' || v === undefined) ? null : v;
        });
        rows.push(cells);
      }

      // Return null if there are no real campaign rows (only "-" placeholders or zeros)
      const hasRealData = rows.some(cells => {
        const name = String(cells[0] ?? '');
        return name !== '-' && !/^total$/i.test(name);
      });
      if (!hasRealData) return null;

      return { headers, rows, platform };
    }

    // Match both "Email Campaigns - Last 30 Days" and "Email Campaign Last 30 Days (Open)"
    const email30d = parseEmailSection(/email\s+campaigns?\s+.*30/i);
    const email90d = parseEmailSection(/email\s+campaigns?\s+.*90/i);

    loyMarketing = { smsWoW, smsCols, email30d, email90d };
  }

  return { loyalty: W, loyaltyMarketing: loyMarketing };
}

// ── parseCateringWB ─────────────────────────────────────────────────────────
function parseCateringWB(wb, weekLabel) {
  const ws = name => sheetRows(wb, name);
  const out = {};

  // Resolve the orders/metrics sheet name. The caller passes "Week of <Mon> <Day>"
  // (derived from deriveWeekLabel); some workbooks include a leading space on the
  // sheet name (" Week of May 4") — match either form.
  const target = weekLabel.replace(/^\s*Week of\s+/i, '').trim();
  const matchSheet = wb.SheetNames.find(n => n.trim().toLowerCase() === ('week of ' + target).toLowerCase());
  const sheetName = matchSheet || ('Week of ' + target);
  const monthDay = target;

  // Orders sheet
  const rows = ws(sheetName);
  if (rows && rows.length >= 5) {
    const obOrders=[], ibOrders=[];
    let obClosed=0, obConfirmed=0, ibClosed=0, ibConfirmed=0;
    let curCohort = null;
    for (const r of rows) {
      if (!r) continue;
      const cohort = safeStr(r[0]);
      if (cohort === 'Outbound' || cohort === 'Inbound') curCohort = cohort;
      const status = safeStr(r[1]);
      const name = safeStr(r[2]);
      const email = safeStr(r[3]);
      const company = safeStr(r[4]);
      const phone = r[5] != null ? String(r[5]) : '';
      const val = num(r[6]);
      if (curCohort && status && name && name !== '-' && val > 0) {
        const entry = { cohort: curCohort, status, name, company, email, phone, value: '$' + val.toLocaleString('en-US', { maximumFractionDigits: 0 }) };
        if (curCohort === 'Outbound') {
          obOrders.push(entry);
          if (status === 'CLOSED') obClosed += val; else obConfirmed += val;
        } else {
          ibOrders.push(entry);
          if (status === 'CLOSED') ibClosed += val; else ibConfirmed += val;
        }
      }
    }
    const fmt = v => '$' + v.toLocaleString('en-US', { maximumFractionDigits: 0 });
    // The source sheet ends each cohort block with an Order Value total row (only
    // the last column filled). Append it so the orders table shows its total.
    if (obOrders.length) obOrders.push({ isTotal: true, value: fmt(obClosed + obConfirmed) });
    if (ibOrders.length) ibOrders.push({ isTotal: true, value: fmt(ibClosed + ibConfirmed) });
    out.outboundOrders = obOrders;
    out.outboundSummary = [
      { label: 'Total Order Value - Closed', value: fmt(obClosed) },
      { label: 'Total Order Value - Confirmed', value: fmt(obConfirmed) },
      { label: 'Total', value: fmt(obClosed + obConfirmed), isTotal: true },
    ];
    out.inboundOrders = ibOrders;
    out.inboundSummary = [
      { label: 'Total Order Value - Closed', value: fmt(ibClosed) },
      { label: 'Total Order Value - Confirmed', value: fmt(ibConfirmed) },
      { label: 'Total', value: fmt(ibClosed + ibConfirmed), isTotal: true },
    ];
    // Save raw totals so the page can build the slide-26 summary
    out._totals = { obClosed, obConfirmed, ibClosed, ibConfirmed };
  }

  // Sheet1 — Outbound team metrics
  const s1 = ws('Sheet1');
  if (s1 && s1.length > 15) {
    const hdr = s1[1] || [];
    let newCol = -1;
    for (let c = 0; c < hdr.length; c++) {
      const h = hdr[c];
      if (h && safeStr(h).replace(/\n/g, ' ').trim() === 'Week of ' + monthDay) { newCol = c; break; }
    }
    if (newCol > -1) {
      let planCol = -1, ovrCol = -1;
      for (let c = 0; c < hdr.length; c++) {
        const h = hdr[c]; if (h == null) continue;
        const lbl = safeStr(h).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
        if (lbl === 'Weekly Plan') planCol = c;
        else if (lbl === 'Overall') ovrCol = c;
      }
      if (planCol < 0) planCol = 39;
      if (ovrCol < 0) ovrCol = 41;
      const metricRows = {
        '# Calls / Week':              { idx: 2,  fmt: 'num' },
        '# Calls Connected':           { idx: 3,  fmt: 'num' },
        '% Calls Connected':           { idx: 4,  fmt: 'pct' },
        '# Emails / Week':             { idx: 7,  fmt: 'num' },
        '# Opens':                     { idx: 8,  fmt: 'num' },
        'MQLs':                        { idx: 9,  fmt: 'num' },
        'MQLs with a FUP date':        { idx: 10, fmt: 'num' },
        'Tasting - Completed':         { idx: 11, fmt: 'num' },
        'Outbound - Orders Closed':    { idx: 12, fmt: 'num', highlight: true },
        'Outbound - Orders Confirmed': { idx: 13, fmt: 'num' },
        'Outbound - Order Value':      { idx: 14, fmt: 'val', highlight: true },
        'Outbound - ROI':              { idx: 15, fmt: 'num', highlight: true },
        'Avg Item Cost':               { idx: 16, fmt: 'val', optional: true },
      };
      const colLabels = [];
      for (let c = newCol - 4; c < newCol; c++) {
        const h = hdr[c];
        const lbl = h ? safeStr(h).replace(/\n/g, ' ').trim() : '';
        const parts = lbl.replace('Week of', '').trim().split(' ');
        colLabels.push('Wk ' + (parts[0] || '') + ' ' + (parts[1] || ''));
      }
      const newLbl = 'Wk ' + monthDay;
      const fmtV = (v, f, metric) => {
        if (v == null || v === '-' || v === '') return '-';
        if (typeof v === 'number') {
          if (f === 'pct') return Math.round(v * 100) + '%';
          if (f === 'val') return '$' + Math.round(v).toLocaleString();
          if (metric === 'Outbound - ROI') return (Math.round(v * 100) / 100).toFixed(2);
          return String(Math.round(v));
        }
        return String(v);
      };
      // Color-coding palette (matches HTML mint/pink/amber/lpurp/lgreen tokens). The
      // renderer feeds these as inline-style background colors via `cc(value, bg)`.
      const COL_MINT = 'mint', COL_PINK = 'pink', COL_AMBER = 'amber', COL_LGREEN = 'lgreen';
      out.outboundMetrics = Object.entries(metricRows).flatMap(([metric, cfg]) => {
        const row = s1[cfg.idx] || [];
        if (cfg.optional) {
          const hasData = row.slice(newCol - 4, newCol + 1).some(v => typeof v === 'number');
          if (!hasData) return [];
        }
        const rec = { metric };
        for (let i = 0; i < 4; i++) rec[colLabels[i]] = fmtV(row[newCol - 4 + i], cfg.fmt, metric);
        rec[newLbl] = fmtV(row[newCol], cfg.fmt, metric);
        rec.plan    = fmtV(row[planCol], cfg.fmt, metric);
        if (cfg.fmt === 'val' && typeof row[ovrCol] === 'number') rec.overall = '$' + Math.round(row[ovrCol]).toLocaleString();
        else rec.overall = fmtV(row[ovrCol], cfg.fmt, metric);
        rec.highlight = cfg.highlight || false;
        rec._cols = { c1: colLabels[0], c2: colLabels[1], c3: colLabels[2], c4: colLabels[3], c5: newLbl };
        // Raw numeric values for sparkline charts (12 metrics × 5 weeks).
        // For percentage rows multiply by 100 so the sparkline reads as a percent.
        const rawAt = c => {
          const v = row[c];
          if (typeof v !== 'number') return 0;
          if (cfg.fmt === 'pct') return Math.round(v * 100);
          return Math.round(v * 100) / 100;
        };
        rec.raw = [
          rawAt(newCol - 4), rawAt(newCol - 3), rawAt(newCol - 2),
          rawAt(newCol - 1), rawAt(newCol),
        ];
        // Per-cell background colors — mint for "on/above plan", pink/amber for misses,
        // lgreen for overall outperform. The plan column always sits in #5 (plan)
        // followed by overall (#6). We only color the 5 weekly cells + overall when
        // applicable; the renderer expects exactly 7 entries (5 weeks + plan + overall).
        const plan = typeof row[planCol] === 'number' ? row[planCol] : null;
        const cellBg = (v) => {
          if (plan == null || typeof v !== 'number' || plan === 0) return '';
          const ratio = v / plan;
          if (ratio >= 1) return COL_MINT;
          if (ratio >= 0.75) return COL_AMBER;
          return COL_PINK;
        };
        const ovrBg = (() => {
          if (!cfg.highlight) return '';
          const v = row[ovrCol];
          if (typeof v !== 'number' || plan == null) return '';
          return v >= plan ? COL_LGREEN : '';
        })();
        rec.bgs = [
          cellBg(row[newCol - 4]), cellBg(row[newCol - 3]), cellBg(row[newCol - 2]),
          cellBg(row[newCol - 1]), cellBg(row[newCol]),
          '',          // plan column is never highlighted
          ovrBg,
        ];
        return rec;
      });
    }
  }

  // New vs Repeat customer breakdown table from Sheet1.
  // The table appears to the right of the main metrics: header row has 'Repeat Customer'
  // in one cell; 'Total' is one column left, 'New' is one column right, category two left.
  if (s1 && s1.length > 20) {
    let repeatCol = -1, custHdrRow = -1;
    for (let r = 0; r < s1.length; r++) {
      const row = s1[r]; if (!row) continue;
      for (let c = 30; c < row.length; c++) {
        if (typeof row[c] === 'string' && /repeat\s*customer/i.test(row[c].trim())) {
          repeatCol = c; custHdrRow = r; break;
        }
      }
      if (custHdrRow > -1) break;
    }
    if (custHdrRow > -1 && repeatCol > 2) {
      const C_TOTAL = repeatCol - 1, C_CAT = repeatCol - 2, C_NEW = repeatCol + 1;
      const custRows = [];
      for (let r = custHdrRow + 1; r < s1.length; r++) {
        const row = s1[r]; if (!row) continue;
        const cat = safeStr(row[C_CAT] || '');
        if (!cat) break;
        const total   = typeof row[C_TOTAL]    === 'number' ? row[C_TOTAL]    : null;
        const repeat  = typeof row[repeatCol]  === 'number' ? row[repeatCol]  : null;
        const newCust = typeof row[C_NEW]      === 'number' ? row[C_NEW]      : null;
        if (total == null && repeat == null && newCust == null) break;
        custRows.push({ label: cat, total, repeat, newCust });
      }
      if (custRows.length) out.customerBreakdown = custRows;
    }
  }

  // Pre-built catering summary block in Sheet1 (typically rows 18-24).
  // "Total Catering" sits in the num column; EzCater/Other3P values come
  // directly from the workbook and may differ from WBR sub-cat actuals.
  if (s1) {
    for (let i = 10; i < Math.min(s1.length, 40); i++) {
      const r = s1[i]; if (!r) continue;
      const col = r.findIndex(v => /^Total Catering$/i.test(safeStr(v || '')));
      if (col < 0) continue;
      const block = [];
      for (let j = Math.max(0, i - 6); j <= i; j++) {
        const rj = s1[j]; if (!rj) continue;
        const numField  = safeStr(rj[col]     || '');
        const label     = safeStr(rj[col + 1] || '');
        if (numField === '#' || /^Week of/i.test(label)) continue;
        const orderVal  = rj[col + 2];
        const closed    = rj[col + 3];
        const confirmed = rj[col + 4];
        if (numField || label) block.push({ num: numField, label, orderVal, closed, confirmed });
      }
      if (block.length) out._s1Summary = block;
      break;
    }
  }

  // WoWComparision trend
  const wowC = ws('WoWComparision');
  if (wowC && wowC.length > 1) {
    function parseTrendRow(hdrRow, valRow) {
      const trend = [];
      for (let c = 1; c < hdrRow.length; c++) {
        const h = hdrRow[c], v = valRow[c];
        if (h == null && v == null) continue;
        let label = '';
        if (h instanceof Date) label = h.toLocaleString('en-US', { month: 'short', day: 'numeric' });
        else if (h) {
          const ss = safeStr(h).replace('Week of', '').trim().replace(/\s+/g, ' ');
          const parts = ss.split(' ');
          label = parts.length >= 2 ? (parts[0].substring(0,3) + ' ' + parts[1]) : ss;
        }
        const val = typeof v === 'number' ? Math.round(v) : 0;
        if (label) trend.push({ week: label, val });
      }
      return trend;
    }

    const obTrend = parseTrendRow(wowC[0] || [], wowC[1] || []);
    if (obTrend.length) out.trend = obTrend;

    // Inbound trend added in June 8: header at row 10 (starts col 9), values at row 11
    if (wowC.length > 11) {
      const ibTrend = parseTrendRow(wowC[10] || [], wowC[11] || []);
      if (ibTrend.length) out.ibTrend = ibTrend;
    }
  }

  return out;
}

// ── File helpers ────────────────────────────────────────────────────────────
function pickFile(files, ...needles) {
  // Find a file whose name matches any of the case-insensitive substrings.
  for (const n of needles) {
    const found = files.find(f => f.toLowerCase().includes(n.toLowerCase()));
    if (found) return found;
  }
  return null;
}

function readWB(filePath) {
  const buf = fs.readFileSync(filePath);
  return XLSX.read(buf, { type: 'buffer', cellDates: true });
}

// Inflate "11-17" → "Week of May 11" by peeking at the catering XLSX (cheap —
// only sheet names are read). Falls back to "Week of <folderName>" when the
// workbook doesn't expose a "Week of <Mon> <Day>" sheet.
function inferWeekLabel(folderName, fallbackLabel) {
  if (fallbackLabel) return fallbackLabel;
  return 'Week of ' + folderName;
}

// Folder names like "11-17" represent a week starting on day 11. Look in the
// catering workbook for a "Week of <Month> 11" sheet — that's the canonical
// label. The workbook accumulates weeks across years, so pick by start-day match.
const MONTHS_MAP = {
  jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3,
  apr: 4, april: 4, may: 5, jun: 6, june: 6, jul: 7, july: 7,
  aug: 8, august: 8, sep: 9, sept: 9, september: 9,
  oct: 10, october: 10, nov: 11, november: 11, dec: 12, december: 12,
};

export function deriveWeekLabel(folderPath) {
  const folderName = path.basename(folderPath);
  // If the folder is already named like "Week of May 11", that's the label.
  if (/^\s*Week of\s+[A-Za-z\.]+\s+\d+/i.test(folderName)) {
    return folderName.trim().replace(/\s+/g, ' ');
  }
  const startDay = parseInt(String(folderName).split('-')[0], 10);
  try {
    const files = fs.readdirSync(folderPath).filter(f => f.toLowerCase().endsWith('.xlsx'));
    const catFile = pickFile(files, 'catering', 'internal purpose');
    if (catFile) {
      const wb = readWB(path.join(folderPath, catFile));
      const re = /^\s*Week of\s+([A-Za-z\.]+)\s+(\d+)(?:st|nd|rd|th)?\s*$/;
      const candidates = wb.SheetNames
        .map(n => {
          const m = re.exec(n);
          if (!m) return null;
          const mon = m[1].replace('.', '').toLowerCase();
          const day = parseInt(m[2], 10);
          const monthIdx = MONTHS_MAP[mon];
          if (!monthIdx) return null;
          return { name: n.trim(), monthIdx, day };
        })
        .filter(Boolean);
      if (!Number.isNaN(startDay)) {
        const dayMatches = candidates.filter(c => c.day === startDay);
        if (dayMatches.length) {
          dayMatches.sort((a, b) => b.monthIdx - a.monthIdx);
          return dayMatches[0].name;
        }
      }
      if (candidates.length) {
        candidates.sort((a, b) => (b.monthIdx - a.monthIdx) || (b.day - a.day));
        return candidates[0].name;
      }
    }
  } catch {}
  return 'Week of ' + folderName;
}

// ── Public: parse an entire week folder ─────────────────────────────────────
export function parseWeekFolder(folderPath, weekLabel) {
  const stat = fs.statSync(folderPath);
  if (!stat.isDirectory()) throw new Error('Not a directory: ' + folderPath);
  const files = fs.readdirSync(folderPath).filter(f => f.toLowerCase().endsWith('.xlsx'));

  const wbrFile = pickFile(files, 'weekly review', 'wbr', 'powered by kutlerri');
  const loyFile = pickFile(files, 'loyalty');
  const catFile = pickFile(files, 'catering', 'internal purpose');

  const label = weekLabel || deriveWeekLabel(folderPath);

  let weekly={}, ptd={}, ytd={}, revCenter={}, revCenterByLoc={}, subCats={},
      subCatsByLoc={}, reviews={}, ue={}, dd={}, bikky={}, catering={};
  let loyalty={}, loyaltyMarketing=null;
  let catSales={};
  let qtd=null, qtdAvailable=false;

  if (wbrFile) {
    const wb = readWB(path.join(folderPath, wbrFile));
    const W = parseWBR(wb);
    weekly = W.weekly || {};
    ptd = W.ptd || {};
    ytd = W.ytd || {};
    revCenter = W.revCenter || {};
    revCenterByLoc = W.revCenterByLoc || {};
    subCats = W.subCats || {};
    subCatsByLoc = W.subCatsByLoc || {};
    reviews = W.reviews || {};
    ue = W.ue || {};
    dd = W.dd || {};
    bikky = W.bikky || {};
    catering = W.catering || {};
    qtd = W.qtd || null;
    qtdAvailable = !!W.qtdAvailable;

    // Inject weekly and PTD per-RC budget from Budget Import sheets.
    // The sheets pre-calculate exact weekly and PTD values by fiscal week number.
    const weekNum = weekNumForLabel(label);
    if (weekNum) {
      const injectBudget = (rows, budMap) => {
        if (!rows || !budMap) return;
        for (const r of rows) {
          const bud = budMap[r.center];
          if (bud == null) continue;
          r.budget  = bud;
          r.varDBud = r.actual - bud;
          r.varPBud = bud !== 0 ? r.varDBud / Math.abs(bud) : null;
        }
      };
      // Consolidated
      if (W.rcBudgetByWeek) {
        const wkData = W.rcBudgetByWeek[weekNum];
        if (wkData) {
          injectBudget(revCenter.weekly, wkData.weekly);
          injectBudget(revCenter.ptd,    wkData.ptd);
        }
      }
      // Per-location
      if (W.rcBudgetByWeekLoc) {
        for (const [loc, byWeek] of Object.entries(W.rcBudgetByWeekLoc)) {
          const wkData = byWeek[weekNum];
          if (!wkData || !revCenterByLoc[loc]) continue;
          injectBudget(revCenterByLoc[loc].weekly, wkData.weekly);
          injectBudget(revCenterByLoc[loc].ptd,    wkData.ptd);
        }
      }
    }
  }
  if (loyFile) {
    const wb = readWB(path.join(folderPath, loyFile));
    const L = parseLoyalty(wb);
    loyalty = L.loyalty || {};
    if (L.loyaltyMarketing) loyaltyMarketing = L.loyaltyMarketing;
  }
  if (catFile) {
    const wb = readWB(path.join(folderPath, catFile));
    catSales = parseCateringWB(wb, label);
  }

  // Build the slide-26 catering summary, preferring the pre-built block from
  // Sheet1 (which uses the catering workbook's own EzCater/total figures).
  if (catSales._totals) {
    const t   = catSales._totals;
    const s1s = catSales._s1Summary;
    const fmt = v => '$' + Math.round(v).toLocaleString();
    if (s1s && s1s.length >= 5) {
      catSales.summary = s1s.map(row => {
        const isTotalRow = /^Total Catering$/i.test(row.num);
        const isIBOB     = /^1\+2$/.test(row.num);
        return {
          num:       isTotalRow ? '' : row.num,
          label:     isTotalRow ? 'Total Catering' : (row.label || ''),
          orderVal:  typeof row.orderVal  === 'number' ? fmt(row.orderVal)  : '-',
          closed:    typeof row.closed    === 'number' ? fmt(row.closed)    : '-',
          confirmed: typeof row.confirmed === 'number' ? fmt(row.confirmed) : '-',
          rowStyle:  isTotalRow || isIBOB ? 'total-row' : '',
        };
      });
    } else {
      const ez      = (subCats.weekly?.catering || []).find(r => /ezcater/i.test(r.sub))?.actual || 0;
      const other3p = (subCats.weekly?.catering || []).find(r => /other 3rd parties/i.test(r.sub))?.actual || 0;
      catSales.summary = [
        { num: '1',   label: 'Catering Outbound', orderVal: fmt(t.obClosed + t.obConfirmed), closed: fmt(t.obClosed), confirmed: fmt(t.obConfirmed), rowStyle: '' },
        { num: '2',   label: 'Catering Inbound',  orderVal: fmt(t.ibClosed + t.ibConfirmed), closed: fmt(t.ibClosed), confirmed: fmt(t.ibConfirmed), rowStyle: '' },
        { num: '1+2', label: 'IB + OB Catering',  orderVal: fmt(t.obClosed + t.obConfirmed + t.ibClosed + t.ibConfirmed), closed: fmt(t.obClosed + t.ibClosed), confirmed: fmt(t.obConfirmed + t.ibConfirmed), rowStyle: 'total-row' },
        { num: '3',   label: 'EzCater',           orderVal: '-', closed: fmt(ez),      confirmed: '-', rowStyle: '' },
        { num: '4',   label: 'Other 3rd-Parties', orderVal: '-', closed: fmt(other3p), confirmed: '-', rowStyle: '' },
        { num: '',    label: 'Total Catering',    orderVal: '-', closed: fmt(t.obClosed + t.ibClosed + ez + other3p), confirmed: '-', rowStyle: 'total-row' },
      ];
    }
    delete catSales._totals;
    delete catSales._s1Summary;
  }

  return {
    label,
    weekly, ptd, ytd,
    revCenter, revCenterByLoc, subCats, subCatsByLoc,
    reviews, ue, dd, bikky,
    catering, loyalty, loyaltyMarketing,
    catSales,
    qtd, qtdAvailable,
  };
}

// ── Public: list all week folders under data/ ───────────────────────────────
export function listWeekFolders(dataDir) {
  if (!fs.existsSync(dataDir)) return [];
  return fs.readdirSync(dataDir, { withFileTypes: true })
    .filter(d => d.isDirectory())
    .map(d => d.name)
    .sort();
}

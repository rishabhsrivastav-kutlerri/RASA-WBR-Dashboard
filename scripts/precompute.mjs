// ─────────────────────────────────────────────────────────────────────────────
// Build-time precompute
//
// Parses every week folder (data/<week>/) and every scorecard workbook
// (scorecard/<granularity>/*.xlsx) ONCE, here at build time, and writes the
// finished results as JSON under generated/. The runtime API routes then just
// read that JSON instead of re-parsing Excel on every request — which is where
// ~all of the Fluid Active CPU was going.
//
// Reuses the app's own parser (parseWeekFolder / loadScorecard / listScorecards)
// so the output is byte-for-byte what the routes produced before.
// ─────────────────────────────────────────────────────────────────────────────
import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import { parseWeekFolder, mergeCostsByCategoryHistory } from '../lib/xlsxParser.js';
import { listScorecards, loadScorecard } from '../lib/scorecard.js';
import { weekInfoForLabel, weekNumForLabel } from '../lib/fiscalCalendar.js';

const ROOT = process.cwd();
const OUT = path.join(ROOT, 'generated');

// Vercel restores .next/cache from the previous successful deploy before
// running the build, so a fingerprint + cached-output pair written here
// survives across deploys. Every admin upload changes exactly one file, so
// without this every deploy re-parses all weeks/scorecards from scratch —
// this lets unchanged ones skip straight to their last computed result.
const CACHE_DIR = path.join(ROOT, '.next', 'cache', 'precompute-cache');

function writeJson(relPath, data) {
  const file = path.join(OUT, relPath);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(data));
  return fs.statSync(file).size;
}

// Every fingerprint also folds in a hash of the parser code itself — without
// this, fixing a parsing bug wouldn't actually change any cached week's
// fingerprint (its source .xlsx files haven't changed), so the cache would
// keep serving the pre-fix output until someone re-uploads that week's file.
const PARSER_SOURCE_FILES = [
  path.join(ROOT, 'lib', 'xlsxParser.js'),
  path.join(ROOT, 'lib', 'scorecard.js'),
  path.join(ROOT, 'lib', 'fiscalCalendar.js'),
];
const CODE_VERSION = crypto.createHash('sha1')
  .update(PARSER_SOURCE_FILES.map(p => { try { return fs.readFileSync(p); } catch { return 'MISSING'; } }).join(''))
  .digest('hex');

// Fingerprint = hash of every input file's name + content that could affect
// this item's parsed output, plus CODE_VERSION above. Content-based, not
// mtime-based — a fresh git clone resets every file's mtime to checkout
// time, so mtime can't tell "changed" from "unchanged" across deploys the
// way file content can.
function fingerprint(filePaths) {
  const hash = crypto.createHash('sha1');
  hash.update(CODE_VERSION);
  for (const p of filePaths.slice().sort()) {
    hash.update(p);
    try { hash.update(fs.readFileSync(p)); } catch { hash.update('MISSING'); }
  }
  return hash.digest('hex');
}

function readCache(cacheKey) {
  try {
    return JSON.parse(fs.readFileSync(path.join(CACHE_DIR, cacheKey + '.json'), 'utf8'));
  } catch {
    return null;
  }
}

function writeCache(cacheKey, entry) {
  const file = path.join(CACHE_DIR, cacheKey + '.json');
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(entry));
}

function listWeekDirs() {
  const dataDir = path.join(ROOT, 'data');
  if (!fs.existsSync(dataDir)) return [];
  const dirs = fs
    .readdirSync(dataDir)
    .filter((w) => {
      try {
        return fs.statSync(path.join(dataDir, w)).isDirectory();
      } catch {
        return false;
      }
    });
  // Chronological (fiscal week) order, not alphabetical — "Week of July 13"
  // sorts before "Week of June 1" alphabetically, which would be backwards
  // for the costsByCategory carry-forward below. Unparseable labels sort
  // last (Infinity) rather than breaking the run.
  return dirs.sort((a, b) => (weekNumForLabel(a) ?? Infinity) - (weekNumForLabel(b) ?? Infinity));
}

function human(bytes) {
  return bytes > 1e6 ? (bytes / 1e6).toFixed(1) + ' MB' : Math.round(bytes / 1e3) + ' KB';
}

async function main() {
  const t0 = Date.now();
  // Start clean so deleted weeks/scorecards don't linger as stale JSON.
  fs.rmSync(OUT, { recursive: true, force: true });
  fs.mkdirSync(OUT, { recursive: true });

  // ── Weeks ──────────────────────────────────────────────────────────────────
  // Processed in chronological order so each week's costsByCategory can carry
  // forward the previous week's already-known category history — Full
  // History/Trailing 26 Weeks then keep growing every week even when a given
  // week only ever uploads one 12-week PCR file (see mergeCostsByCategoryHistory).
  const weeks = listWeekDirs();
  const sheets = [];
  let okWeeks = 0;
  let cachedWeeks = 0;
  let prevCostsByCategory = null;
  for (const week of weeks) {
    try {
      const weekDir = path.join(ROOT, 'data', week);
      const pcrDir = path.join(ROOT, 'PCR', week);
      const depFiles = [
        ...fs.readdirSync(weekDir).map(f => path.join(weekDir, f)),
        ...(fs.existsSync(pcrDir) ? fs.readdirSync(pcrDir).map(f => path.join(pcrDir, f)) : []),
      ];
      const fp = fingerprint(depFiles);
      const cacheKey = 'weeks/' + week;
      const cached = readCache(cacheKey);

      let data;
      if (cached && cached.fingerprint === fp) {
        data = cached.data;
        prevCostsByCategory = data.costsByCategory || prevCostsByCategory;
        cachedWeeks++;
      } else {
        data = await parseWeekFolder(path.join('data', week));
        if (data.costsByCategory) {
          data.costsByCategory = mergeCostsByCategoryHistory(data.costsByCategory, prevCostsByCategory);
          prevCostsByCategory = data.costsByCategory;
        }
        writeCache(cacheKey, { fingerprint: fp, data });
      }

      const size = writeJson(path.join('weeks', week + '.json'), data);
      const info = weekInfoForLabel(week);
      sheets.push({
        week,
        label: week,
        period: info ? info.period : null,
        weekInPeriod: info ? info.weekInPeriod : null,
      });
      okWeeks++;
      console.log(`  week  ✓ ${week.padEnd(22)} ${human(size)}${cached && cached.fingerprint === fp ? '  (cached)' : ''}`);
    } catch (err) {
      console.error(`  week  ✗ ${week}: ${err.message}`);
    }
  }
  writeJson('sheets.json', sheets);

  // ── Scorecards ───────────────────────────────────────────────────────────────
  let okCards = 0;
  let cachedCards = 0;
  try {
    const index = listScorecards(); // { weekly:[{id,label,sort}], period:[...], quarter:[...] }
    writeJson(path.join('scorecard', 'index.json'), index);
    for (const granularity of Object.keys(index)) {
      const byId = {};
      for (const item of index[granularity]) {
        try {
          const file = path.join(ROOT, 'scorecard', granularity, item.id);
          const fp = fingerprint([file]);
          const cacheKey = 'scorecard/' + granularity + '/' + item.id;
          const cached = readCache(cacheKey);
          if (cached && cached.fingerprint === fp) {
            byId[item.id] = cached.data;
            cachedCards++;
          } else {
            const data = loadScorecard(granularity, item.id);
            writeCache(cacheKey, { fingerprint: fp, data });
            byId[item.id] = data;
          }
          okCards++;
        } catch (err) {
          console.error(`  card  ✗ ${granularity}/${item.id}: ${err.message}`);
        }
      }
      writeJson(path.join('scorecard', granularity + '.json'), byId);
    }
    console.log(`  scorecards ✓ ${okCards} across ${Object.keys(index).length} granularities (${cachedCards} cached)`);
  } catch (err) {
    console.error(`  scorecards ✗ ${err.message}`);
    // Still emit an empty index so the route has something to read.
    writeJson(path.join('scorecard', 'index.json'), { weekly: [], period: [], quarter: [] });
  }

  console.log(
    `\nprecompute: ${okWeeks}/${weeks.length} weeks (${cachedWeeks} cached), ${okCards} scorecards (${cachedCards} cached) → generated/  (${((Date.now() - t0) / 1000).toFixed(1)}s)`,
  );
}

main().catch((err) => {
  console.error('precompute failed:', err);
  process.exit(1);
});

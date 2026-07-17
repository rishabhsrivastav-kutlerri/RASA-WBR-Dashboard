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
import { parseWeekFolder } from '../lib/xlsxParser.js';
import { listScorecards, loadScorecard } from '../lib/scorecard.js';
import { weekInfoForLabel } from '../lib/fiscalCalendar.js';

const ROOT = process.cwd();
const OUT = path.join(ROOT, 'generated');

function writeJson(relPath, data) {
  const file = path.join(OUT, relPath);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(data));
  return fs.statSync(file).size;
}

function listWeekDirs() {
  const dataDir = path.join(ROOT, 'data');
  if (!fs.existsSync(dataDir)) return [];
  return fs
    .readdirSync(dataDir)
    .filter((w) => {
      try {
        return fs.statSync(path.join(dataDir, w)).isDirectory();
      } catch {
        return false;
      }
    })
    .sort();
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
  const weeks = listWeekDirs();
  const sheets = [];
  let okWeeks = 0;
  for (const week of weeks) {
    try {
      const data = parseWeekFolder(path.join('data', week));
      const size = writeJson(path.join('weeks', week + '.json'), data);
      const info = weekInfoForLabel(week);
      sheets.push({
        week,
        label: week,
        period: info ? info.period : null,
        weekInPeriod: info ? info.weekInPeriod : null,
      });
      okWeeks++;
      console.log(`  week  ✓ ${week.padEnd(22)} ${human(size)}`);
    } catch (err) {
      console.error(`  week  ✗ ${week}: ${err.message}`);
    }
  }
  writeJson('sheets.json', sheets);

  // ── Scorecards ───────────────────────────────────────────────────────────────
  let okCards = 0;
  try {
    const index = listScorecards(); // { weekly:[{id,label,sort}], period:[...], quarter:[...] }
    writeJson(path.join('scorecard', 'index.json'), index);
    for (const granularity of Object.keys(index)) {
      const byId = {};
      for (const item of index[granularity]) {
        try {
          byId[item.id] = loadScorecard(granularity, item.id);
          okCards++;
        } catch (err) {
          console.error(`  card  ✗ ${granularity}/${item.id}: ${err.message}`);
        }
      }
      writeJson(path.join('scorecard', granularity + '.json'), byId);
    }
    console.log(`  scorecards ✓ ${okCards} across ${Object.keys(index).length} granularities`);
  } catch (err) {
    console.error(`  scorecards ✗ ${err.message}`);
    // Still emit an empty index so the route has something to read.
    writeJson(path.join('scorecard', 'index.json'), { weekly: [], period: [], quarter: [] });
  }

  console.log(
    `\nprecompute: ${okWeeks}/${weeks.length} weeks, ${okCards} scorecards → generated/  (${((Date.now() - t0) / 1000).toFixed(1)}s)`,
  );
}

main().catch((err) => {
  console.error('precompute failed:', err);
  process.exit(1);
});

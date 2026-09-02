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
import os from 'os';
import crypto from 'crypto';
import { execFile } from 'child_process';
import { fileURLToPath } from 'url';
import { mergeCostsByCategoryHistory } from '../lib/xlsxParser.js';
import { listScorecards } from '../lib/scorecard.js';
import { weekInfoForLabel, weekNumForLabel } from '../lib/fiscalCalendar.js';

const ROOT = process.cwd();
const OUT = path.join(ROOT, 'generated');
const WORKER = path.join(path.dirname(fileURLToPath(import.meta.url)), 'precompute-worker.mjs');
// On Windows, node_modules/.bin/tsx is a shell shim with no extension (plus a
// .cmd wrapper alongside it) — execFile can't exec either one directly
// without a shell (ENOENT on the extensionless file, EINVAL on the .cmd),
// unlike on Linux/Mac where the extensionless shim is a real executable
// script. Vercel's build servers are Linux, so this only bit local Windows
// dev — use the .cmd + shell:true only on win32, keep the direct exec
// everywhere else.
const IS_WIN = process.platform === 'win32';
const TSX_BIN = path.join(ROOT, 'node_modules', '.bin', IS_WIN ? 'tsx.cmd' : 'tsx');

// How many weeks/scorecards to parse at once, each in its own process (real
// parallelism — see precompute-worker.mjs for why that's necessary). Capped
// well under the machine's core count: each worker loads full xlsx buffers
// into memory (a loyalty.xlsx alone can be 10MB+), so unbounded concurrency
// risks a build OOM more than it saves time. Override with an env var if a
// given Vercel plan's build machine warrants a different number.
const CONCURRENCY = Number(process.env.PRECOMPUTE_CONCURRENCY) || Math.min(os.cpus().length, 6);

// shell:true on Windows hands the joined argv straight to cmd.exe without
// quoting any argument (Node does NOT do this itself — see the DEP0190
// warning) — a week name like "Week of Aug 17" would otherwise be split into
// 4 separate argv entries by cmd.exe's own whitespace tokenizing. Quote each
// arg ourselves and pair with windowsVerbatimArguments so Node doesn't also
// try to re-escape our already-quoted string.
const winQuote = s => '"' + String(s).replace(/"/g, '\\"') + '"';

// Runs one item in its own process via precompute-worker.mjs. Never rejects —
// callers get { ok:false, error } for both a clean thrown error inside the
// worker and an unexpected crash (non-JSON / non-zero exit), so a single bad
// week/scorecard can't take down the whole precompute run.
function runWorker(mode, args) {
  return new Promise((resolve) => {
    const rawArgs = [WORKER, mode, ...args];
    const execArgs = IS_WIN ? rawArgs.map(winQuote) : rawArgs;
    const opts = { maxBuffer: 64 * 1024 * 1024, shell: IS_WIN, windowsVerbatimArguments: IS_WIN };
    execFile(TSX_BIN, execArgs, opts, (err, stdout) => {
      try {
        resolve(JSON.parse(stdout));
      } catch {
        resolve({ ok: false, error: err ? err.message : 'worker produced no parseable output' });
      }
    });
  });
}

// Runs `items` through `fn` with at most `limit` in flight at once.
async function mapLimit(items, limit, fn) {
  const results = new Array(items.length);
  let next = 0;
  async function worker() {
    while (next < items.length) {
      const i = next++;
      results[i] = await fn(items[i], i);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}

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
  // costsByCategory must still be merged in chronological order (each week
  // carries forward the previous week's already-known category history), so
  // that part stays a plain sequential loop below. But the actual parsing —
  // reading and crunching the xlsx files, the expensive part — has no such
  // ordering requirement between weeks, so cache-miss weeks are parsed in
  // parallel worker processes first, then merged in order afterward.
  const weeks = listWeekDirs();

  // Phase 1: figure out which weeks are cache hits vs misses. Cheap — just
  // hashing file contents, not parsing them — so this stays sequential.
  const weekMeta = weeks.map((week) => {
    const weekDir = path.join(ROOT, 'data', week);
    const pcrDir = path.join(ROOT, 'PCR', week);
    const depFiles = [
      ...fs.readdirSync(weekDir).map(f => path.join(weekDir, f)),
      ...(fs.existsSync(pcrDir) ? fs.readdirSync(pcrDir).map(f => path.join(pcrDir, f)) : []),
    ];
    const fp = fingerprint(depFiles);
    const cacheKey = 'weeks/' + week;
    const cached = readCache(cacheKey);
    return { week, fp, cacheKey, cached: cached && cached.fingerprint === fp ? cached : null };
  });

  // Phase 2: parse every cache-miss week concurrently, each in its own process.
  const misses = weekMeta.filter(m => !m.cached);
  const parsedByWeek = new Map();
  await mapLimit(misses, CONCURRENCY, async (m) => {
    const result = await runWorker('week', [m.week]);
    parsedByWeek.set(m.week, result);
  });

  // Phase 3: apply the chronological costsByCategory merge and write output,
  // in order, using whichever result (cache or freshly-parsed) each week has.
  const sheets = [];
  let okWeeks = 0;
  let cachedWeeks = 0;
  let prevCostsByCategory = null;
  for (const m of weekMeta) {
    try {
      let data;
      if (m.cached) {
        data = m.cached.data;
        prevCostsByCategory = data.costsByCategory || prevCostsByCategory;
        cachedWeeks++;
      } else {
        const result = parsedByWeek.get(m.week);
        if (!result.ok) throw new Error(result.error);
        data = result.data;
        if (data.costsByCategory) {
          data.costsByCategory = mergeCostsByCategoryHistory(data.costsByCategory, prevCostsByCategory);
          prevCostsByCategory = data.costsByCategory;
        }
        writeCache(m.cacheKey, { fingerprint: m.fp, data });
      }

      const size = writeJson(path.join('weeks', m.week + '.json'), data);
      const info = weekInfoForLabel(m.week);
      sheets.push({
        week: m.week,
        label: m.week,
        period: info ? info.period : null,
        weekInPeriod: info ? info.weekInPeriod : null,
      });
      okWeeks++;
      console.log(`  week  ✓ ${m.week.padEnd(22)} ${human(size)}${m.cached ? '  (cached)' : ''}`);
    } catch (err) {
      console.error(`  week  ✗ ${m.week}: ${err.message}`);
    }
  }
  writeJson('sheets.json', sheets);

  // ── Scorecards ───────────────────────────────────────────────────────────────
  // Unlike weeks, scorecards have no ordering dependency on each other at all,
  // so every cache-miss card — across all granularities — parses concurrently.
  let okCards = 0;
  let cachedCards = 0;
  try {
    const index = listScorecards(); // { weekly:[{id,label,sort}], period:[...], quarter:[...] }
    writeJson(path.join('scorecard', 'index.json'), index);

    const cardMeta = Object.keys(index).flatMap((granularity) =>
      index[granularity].map((item) => {
        const file = path.join(ROOT, 'scorecard', granularity, item.id);
        const fp = fingerprint([file]);
        const cacheKey = 'scorecard/' + granularity + '/' + item.id;
        const cached = readCache(cacheKey);
        return { granularity, id: item.id, fp, cacheKey, cached: cached && cached.fingerprint === fp ? cached : null };
      })
    );

    const cardMisses = cardMeta.filter(c => !c.cached);
    const parsedByCard = new Map();
    await mapLimit(cardMisses, CONCURRENCY, async (c) => {
      const result = await runWorker('scorecard', [c.granularity, c.id]);
      parsedByCard.set(c.cacheKey, result);
    });

    const byGranularity = {};
    for (const c of cardMeta) {
      try {
        let data;
        if (c.cached) {
          data = c.cached.data;
          cachedCards++;
        } else {
          const result = parsedByCard.get(c.cacheKey);
          if (!result.ok) throw new Error(result.error);
          data = result.data;
          writeCache(c.cacheKey, { fingerprint: c.fp, data });
        }
        (byGranularity[c.granularity] ??= {})[c.id] = data;
        okCards++;
      } catch (err) {
        console.error(`  card  ✗ ${c.granularity}/${c.id}: ${err.message}`);
      }
    }
    for (const granularity of Object.keys(index)) {
      writeJson(path.join('scorecard', granularity + '.json'), byGranularity[granularity] || {});
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

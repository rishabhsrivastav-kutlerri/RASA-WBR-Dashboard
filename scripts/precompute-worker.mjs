// ─────────────────────────────────────────────────────────────────────────────
// One-shot worker invoked as a child process by precompute.mjs.
//
// fs.readFileSync + SheetJS's XLSX.read are both fully synchronous/blocking —
// parseWeekFolder has no real async I/O inside it despite being `async`. That
// means calling it many times via Promise.all in a single process would NOT
// run those calls concurrently at all (Node is single-threaded; blocking work
// can't interleave) — it'd just be the same sequential work with extra
// bookkeeping. Running each item in its own OS process is what actually lets
// multiple weeks/scorecards parse on separate CPU cores at once.
//
// Usage: tsx precompute-worker.mjs week "Week of Aug 10"
//        tsx precompute-worker.mjs scorecard weekly some-id.xlsx
//
// Always exits 0 and reports success/failure as JSON on stdout — a non-zero
// exit or unparseable stdout (e.g. an uncaught crash before it could print)
// is treated as a failure by the parent, which already isolates failures per
// item exactly like the old inline try/catch did.
// ─────────────────────────────────────────────────────────────────────────────
import path from 'path';
import { parseWeekFolder } from '../lib/xlsxParser.js';
import { loadScorecard } from '../lib/scorecard.js';

const ROOT = process.cwd();
const [, , mode, ...rest] = process.argv;

async function run() {
  if (mode === 'week') {
    const [week] = rest;
    return parseWeekFolder(path.join(ROOT, 'data', week));
  }
  if (mode === 'scorecard') {
    const [granularity, id] = rest;
    return loadScorecard(granularity, id);
  }
  throw new Error('unknown worker mode: ' + mode);
}

run()
  .then((data) => { process.stdout.write(JSON.stringify({ ok: true, data })); })
  .catch((err) => { process.stdout.write(JSON.stringify({ ok: false, error: err.message || String(err) })); });

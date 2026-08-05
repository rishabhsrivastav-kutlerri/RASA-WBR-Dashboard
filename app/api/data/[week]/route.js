import { NextResponse } from 'next/server';
import path from 'path';
import fs from 'fs';
import { parseWeekFolder, mergeCostsByCategoryHistory } from '@/lib/xlsxParser';
import { verifyAuth } from '@/lib/auth';
import { getWeekStatus, getPcrFileStatus, downloadFileAtPath } from '@/lib/githubStorage';
import { readGenerated } from '@/lib/generated';
import { weekNumForLabel } from '@/lib/fiscalCalendar';

export const runtime = 'nodejs';
export const maxDuration = 60;

// In-memory cache keyed by week + file SHAs (fallback path only).
// Cache is valid as long as SHAs haven't changed — no rebuild needed.
const cache = new Map();

const TMP_BASE = '/tmp/wbr-data';

export async function GET(request, { params }) {
  if (!verifyAuth(request)) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }
  try {
    const { week } = await params;
    if (!week || /[\\/]/.test(week) || week.includes('..')) {
      return NextResponse.json({ error: 'Invalid week name' }, { status: 400 });
    }

    // Fast path: precomputed week JSON (parsed once at build time). This is the
    // normal case and does zero Excel parsing at request time.
    const pre = readGenerated(path.join('weeks', week + '.json'));
    if (pre) return NextResponse.json(pre);

    // ── Fallback: parse live from GitHub ──────────────────────────────────────
    // Only reached if this week has no precomputed JSON yet (e.g. a fresh upload
    // whose rebuild hasn't finished). Behaves exactly as the app did before.
    const { present, paths, shas } = await getWeekStatus(week);
    if (!present.wbr && !present.loyalty && !present.catering) {
      return NextResponse.json({ error: 'Week not found: ' + week }, { status: 404 });
    }
    // PCR workbook (Costs-tab Labor/COGS actuals) — optional, lives in a
    // separate GitHub root. Absent for weeks that haven't been given one yet.
    const pcrStatus = await getPcrFileStatus(week);

    // SHAs change only when a file actually changes in GitHub — perfect cache key.
    const fp = [shas.wbr, shas.loyalty, shas.catering, pcrStatus.sha].join('|');
    const cacheKey = `${week}:${fp}`;
    if (cache.has(cacheKey)) return NextResponse.json(cache.get(cacheKey));

    // Download files from GitHub into /tmp so xlsxParser can read them normally.
    const tmpDir = path.join(TMP_BASE, week);
    fs.mkdirSync(tmpDir, { recursive: true });
    // Clear stale files before writing fresh ones.
    for (const f of fs.readdirSync(tmpDir)) fs.unlinkSync(path.join(tmpDir, f));

    // Download all xlsx files in parallel instead of sequentially.
    const filePaths = Object.values(paths).filter(Boolean);
    if (pcrStatus.path) filePaths.push(pcrStatus.path);
    await Promise.all(
      filePaths.map(async filePath => {
        const buf = await downloadFileAtPath(filePath);
        if (buf) fs.writeFileSync(path.join(tmpDir, path.basename(filePath)), buf);
      })
    );

    const data = await parseWeekFolder(tmpDir);

    // Best-effort carry-forward, same as the build-time precompute chain
    // (see scripts/precompute.mjs): this branch only runs for a week whose
    // precomputed JSON doesn't exist yet (e.g. right after an admin upload,
    // before the next deploy's precompute finishes), so without this it
    // would show only its own file's category history until then. Uses the
    // last successful precompute's sheets index + generated JSON — a fast
    // local read, no extra GitHub calls.
    if (data.costsByCategory) {
      const sheetsIdx = readGenerated('sheets.json') || [];
      const curNum = weekNumForLabel(week);
      if (curNum != null) {
        const prevWeek = sheetsIdx
          .filter(s => { const n = weekNumForLabel(s.week); return n != null && n < curNum; })
          .sort((a, b) => weekNumForLabel(b.week) - weekNumForLabel(a.week))[0];
        const prevData = prevWeek ? readGenerated(path.join('weeks', prevWeek.week + '.json')) : null;
        if (prevData?.costsByCategory) {
          data.costsByCategory = mergeCostsByCategoryHistory(data.costsByCategory, prevData.costsByCategory);
        }
      }
    }

    cache.set(cacheKey, data);
    return NextResponse.json(data);
  } catch (err) {
    console.error('[api/data/[week]]', err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

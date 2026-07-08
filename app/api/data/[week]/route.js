import { NextResponse } from 'next/server';
import path from 'path';
import fs from 'fs';
import { parseWeekFolder } from '@/lib/xlsxParser';
import { verifyAuth } from '@/lib/auth';
import { getWeekStatus, downloadFileAtPath } from '@/lib/githubStorage';

export const runtime = 'nodejs';
export const maxDuration = 60;

// In-memory cache keyed by week + file SHAs.
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

    // Lightweight GitHub call: get file paths + SHAs for this week's folder.
    const { present, paths, shas } = await getWeekStatus(week);
    if (!present.wbr && !present.loyalty && !present.catering) {
      return NextResponse.json({ error: 'Week not found: ' + week }, { status: 404 });
    }

    // SHAs change only when a file actually changes in GitHub — perfect cache key.
    const fp = [shas.wbr, shas.loyalty, shas.catering].join('|');
    const cacheKey = `${week}:${fp}`;
    if (cache.has(cacheKey)) return NextResponse.json(cache.get(cacheKey));

    // Download files from GitHub into /tmp so xlsxParser can read them normally.
    const tmpDir = path.join(TMP_BASE, week);
    fs.mkdirSync(tmpDir, { recursive: true });
    // Clear stale files before writing fresh ones.
    for (const f of fs.readdirSync(tmpDir)) fs.unlinkSync(path.join(tmpDir, f));

    // Download all xlsx files in parallel instead of sequentially.
    await Promise.all(
      Object.values(paths)
        .filter(Boolean)
        .map(async filePath => {
          const buf = await downloadFileAtPath(filePath);
          if (buf) fs.writeFileSync(path.join(tmpDir, path.basename(filePath)), buf);
        })
    );

    const data = parseWeekFolder(tmpDir);
    cache.set(cacheKey, data);
    return NextResponse.json(data);
  } catch (err) {
    console.error('[api/data/[week]]', err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

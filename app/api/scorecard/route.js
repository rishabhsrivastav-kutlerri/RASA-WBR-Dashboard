import path from 'path';
import { NextResponse } from 'next/server';
import { verifyAuth } from '@/lib/auth';
import { listScorecards, loadScorecard } from '@/lib/scorecard';
import { readGenerated } from '@/lib/generated';

export const runtime = 'nodejs';

const cache = new Map();

export async function GET(request) {
  if (!verifyAuth(request)) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }
  const url = new URL(request.url);
  const granularity = url.searchParams.get('granularity');
  const item = url.searchParams.get('item');

  try {
    // ── Index ─────────────────────────────────────────────────────────────────
    if (!granularity) {
      const pre = readGenerated(path.join('scorecard', 'index.json'));
      return NextResponse.json(pre || listScorecards());
    }

    // ── Scorecard data ────────────────────────────────────────────────────────
    if (!item) return NextResponse.json({ error: 'Missing item' }, { status: 400 });

    // Fast path: precomputed granularity map (parsed once at build time).
    const map = readGenerated(path.join('scorecard', granularity + '.json'));
    if (map && Object.prototype.hasOwnProperty.call(map, item)) {
      return NextResponse.json(map[item]);
    }

    // Fallback: live parse (item not in the precomputed map yet).
    const cacheKey = `${granularity}:${item}`;
    if (cache.has(cacheKey)) return NextResponse.json(cache.get(cacheKey));

    const data = loadScorecard(granularity, item);
    cache.set(cacheKey, data);
    return NextResponse.json(data);
  } catch (err) {
    const status = err.code === 'NOT_FOUND' ? 404 : (err.code === 'BAD_ID' || err.code === 'BAD_GRANULARITY') ? 400 : 500;
    return NextResponse.json({ error: err.message }, { status });
  }
}

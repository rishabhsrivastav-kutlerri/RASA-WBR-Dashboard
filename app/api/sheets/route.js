import { NextResponse } from 'next/server';
import { listWeeks } from '@/lib/githubStorage';
import { weekInfoForLabel } from '@/lib/fiscalCalendar';
import { verifyAuth } from '@/lib/auth';
import { readGenerated } from '@/lib/generated';

export const runtime = 'nodejs';

// Fallback cache: only used when the precomputed list is unavailable.
let sheetsCache = null;
let sheetsCacheAt = 0;
const SHEETS_TTL = 2 * 60 * 1000;

export async function GET(request) {
  if (!verifyAuth(request)) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  // Fast path: precomputed sheet list (built once per deploy).
  const pre = readGenerated('sheets.json');
  if (pre) return NextResponse.json({ sheets: pre });

  // Fallback: live GitHub listing (only if precompute is missing).
  try {
    if (sheetsCache && Date.now() - sheetsCacheAt < SHEETS_TTL) {
      return NextResponse.json({ sheets: sheetsCache });
    }
    const weekNames = await listWeeks();
    const sheets = weekNames.map(name => {
      const info = weekInfoForLabel(name);
      return {
        week: name,
        label: name,
        period: info ? info.period : null,
        weekInPeriod: info ? info.weekInPeriod : null,
      };
    });
    sheetsCache = sheets;
    sheetsCacheAt = Date.now();
    return NextResponse.json({ sheets });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

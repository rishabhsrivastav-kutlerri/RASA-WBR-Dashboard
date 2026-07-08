import { NextResponse } from 'next/server';
import { verifyAuth } from '@/lib/auth';
import { getLocationsConfig } from '@/lib/githubStorage';

export const runtime = 'nodejs';

const DEFAULT_LOCATIONS = {
  Ballpark: { open: false },
  MVT: { open: true },
  'National Landing': { open: true },
  Mosaic: { open: true },
  Rockville: { open: true },
};

export async function GET(request) {
  if (!verifyAuth(request)) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }
  try {
    const config = await getLocationsConfig();
    return NextResponse.json({ locations: config || DEFAULT_LOCATIONS });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

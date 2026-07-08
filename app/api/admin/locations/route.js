import { NextResponse } from 'next/server';
import { verifyAdmin } from '@/lib/auth';
import { getLocationsConfig, saveLocationsConfig } from '@/lib/githubStorage';

export const runtime = 'nodejs';

const DEFAULT_LOCATIONS = {
  Ballpark: { open: false },
  MVT: { open: true },
  'National Landing': { open: true },
  Mosaic: { open: true },
  Rockville: { open: true },
};

async function currentLocations() {
  const config = await getLocationsConfig();
  return config || { ...DEFAULT_LOCATIONS };
}

export async function GET(request) {
  if (!verifyAdmin(request)) {
    return NextResponse.json({ error: 'Forbidden' }, { status: 403 });
  }
  try {
    return NextResponse.json({ locations: await currentLocations() });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

export async function POST(request) {
  if (!verifyAdmin(request)) {
    return NextResponse.json({ error: 'Forbidden' }, { status: 403 });
  }
  try {
    const { name, open } = await request.json();
    if (!name || typeof name !== 'string') {
      return NextResponse.json({ error: 'name is required' }, { status: 400 });
    }
    if (typeof open !== 'boolean') {
      return NextResponse.json({ error: 'open must be a boolean' }, { status: 400 });
    }
    const locations = await currentLocations();
    locations[name.trim()] = { open };
    await saveLocationsConfig(locations);
    return NextResponse.json({ ok: true, name: name.trim(), open });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

export async function DELETE(request) {
  if (!verifyAdmin(request)) {
    return NextResponse.json({ error: 'Forbidden' }, { status: 403 });
  }
  try {
    const { name } = await request.json();
    if (!name) return NextResponse.json({ error: 'name is required' }, { status: 400 });
    const locations = await currentLocations();
    if (!locations[name]) {
      return NextResponse.json({ error: 'Location not found' }, { status: 404 });
    }
    delete locations[name];
    await saveLocationsConfig(locations);
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

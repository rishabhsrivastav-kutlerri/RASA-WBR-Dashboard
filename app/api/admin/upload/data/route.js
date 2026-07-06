import { NextResponse } from 'next/server';
import { verifyAdmin } from '@/lib/auth';
import { uploadWeekFile } from '@/lib/githubStorage';

export const runtime = 'nodejs';
export const maxDuration = 60;

const VALID_TYPES = new Set(['wbr', 'loyalty', 'catering']);

export async function POST(request) {
  if (!verifyAdmin(request)) {
    return NextResponse.json({ error: 'Forbidden' }, { status: 403 });
  }
  try {
    const form = await request.formData();
    const weekName = form.get('weekName');
    const fileType = form.get('fileType');
    const file = form.get('file');

    if (!weekName || !fileType || !file) {
      return NextResponse.json({ error: 'weekName, fileType, and file are required' }, { status: 400 });
    }
    if (!VALID_TYPES.has(fileType)) {
      return NextResponse.json({ error: 'fileType must be wbr, loyalty, or catering' }, { status: 400 });
    }
    if (weekName.includes('..') || weekName.includes('/') || weekName.includes('\\')) {
      return NextResponse.json({ error: 'Invalid weekName' }, { status: 400 });
    }
    if (!file.name.match(/\.xlsx$/i)) {
      return NextResponse.json({ error: 'Only .xlsx files are accepted' }, { status: 400 });
    }

    const buffer = Buffer.from(await file.arrayBuffer());
    const path = await uploadWeekFile(weekName, fileType, buffer);
    return NextResponse.json({ ok: true, weekName, fileType, path });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

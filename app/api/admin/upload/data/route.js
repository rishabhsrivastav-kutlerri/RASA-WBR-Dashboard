import { NextResponse } from 'next/server';
import { verifyAdmin } from '@/lib/auth';
import { uploadWeekFile } from '@/lib/githubStorage';
import { del } from '@vercel/blob';

export const runtime = 'nodejs';
export const maxDuration = 60;

const VALID_TYPES = new Set(['wbr', 'loyalty', 'catering']);

export async function POST(request) {
  if (!verifyAdmin(request)) {
    return NextResponse.json({ error: 'Forbidden' }, { status: 403 });
  }
  try {
    const contentType = request.headers.get('content-type') || '';

    let weekName, fileType, buffer;

    if (contentType.includes('application/json')) {
      // Phase 2 of client-side blob upload: client sends { weekName, fileType, blobUrl }
      const body = await request.json();
      weekName = body.weekName;
      fileType = body.fileType;
      const blobUrl = body.blobUrl;

      if (!weekName || !fileType || !blobUrl) {
        return NextResponse.json({ error: 'weekName, fileType, and blobUrl are required' }, { status: 400 });
      }
      if (!VALID_TYPES.has(fileType)) {
        return NextResponse.json({ error: 'fileType must be wbr, loyalty, or catering' }, { status: 400 });
      }
      if (weekName.includes('..') || weekName.includes('/') || weekName.includes('\\')) {
        return NextResponse.json({ error: 'Invalid weekName' }, { status: 400 });
      }

      // Download from Vercel Blob and push to GitHub, then delete the temp blob.
      const res = await fetch(blobUrl);
      if (!res.ok) throw new Error(`Failed to fetch blob: ${res.status}`);
      buffer = Buffer.from(await res.arrayBuffer());
      const path = await uploadWeekFile(weekName, fileType, buffer);
      await del(blobUrl).catch(() => {}); // best-effort cleanup
      return NextResponse.json({ ok: true, weekName, fileType, path });

    } else {
      // Legacy: direct multipart form upload (works for files under ~4MB).
      const form = await request.formData();
      weekName = form.get('weekName');
      fileType = form.get('fileType');
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

      buffer = Buffer.from(await file.arrayBuffer());
      const path = await uploadWeekFile(weekName, fileType, buffer);
      return NextResponse.json({ ok: true, weekName, fileType, path });
    }
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

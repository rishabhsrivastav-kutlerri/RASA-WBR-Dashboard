import { NextResponse } from 'next/server';
import { verifyAdmin } from '@/lib/auth';
import { deleteWeekFile, deleteWeek, deletePcrFile } from '@/lib/githubStorage';

export const runtime = 'nodejs';

export async function DELETE(request) {
  if (!verifyAdmin(request)) {
    return NextResponse.json({ error: 'Forbidden' }, { status: 403 });
  }
  try {
    const { weekName, fileType } = await request.json();
    if (!weekName) return NextResponse.json({ error: 'weekName is required' }, { status: 400 });
    if (weekName.includes('..') || weekName.includes('/') || weekName.includes('\\')) {
      return NextResponse.json({ error: 'Invalid weekName' }, { status: 400 });
    }
    if (fileType === 'pcr') {
      await deletePcrFile(weekName);
    } else if (fileType) {
      await deleteWeekFile(weekName, fileType);
    } else {
      await deleteWeek(weekName);
    }
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

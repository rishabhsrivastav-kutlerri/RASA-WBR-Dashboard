import { NextResponse } from 'next/server';
import { verifyAdmin } from '@/lib/auth';
import { downloadWeekFile, downloadPcrFile } from '@/lib/githubStorage';

export const runtime = 'nodejs';

const VALID_TYPES = new Set(['wbr', 'loyalty', 'catering', 'pcr']);

export async function GET(request) {
  if (!verifyAdmin(request)) {
    return NextResponse.json({ error: 'Forbidden' }, { status: 403 });
  }
  const { searchParams } = new URL(request.url);
  const weekName = searchParams.get('weekName');
  const fileType = searchParams.get('fileType');

  if (!weekName || !fileType) {
    return NextResponse.json({ error: 'weekName and fileType are required' }, { status: 400 });
  }
  if (!VALID_TYPES.has(fileType)) {
    return NextResponse.json({ error: 'Invalid fileType' }, { status: 400 });
  }

  try {
    const result = fileType === 'pcr'
      ? await downloadPcrFile(weekName)
      : await downloadWeekFile(weekName, fileType);
    if (!result) return NextResponse.json({ error: 'File not found' }, { status: 404 });
    return new NextResponse(result.buffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${result.filename}"`,
      },
    });
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

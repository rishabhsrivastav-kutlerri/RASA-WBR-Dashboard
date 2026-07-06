import { NextResponse } from 'next/server';
import { verifyAdmin } from '@/lib/auth';

export const runtime = 'nodejs';

// Returns GitHub credentials so the browser can commit large files directly to GitHub,
// bypassing Vercel's 4.5MB function payload limit entirely.
// Only accessible to verified admins.
export async function GET(request) {
  if (!verifyAdmin(request)) {
    return NextResponse.json({ error: 'Forbidden' }, { status: 403 });
  }
  return NextResponse.json({
    ghToken: process.env.GITHUB_TOKEN,
    owner:   process.env.GITHUB_OWNER,
    repo:    process.env.GITHUB_REPO,
    branch:  process.env.GITHUB_BRANCH || 'main',
  });
}

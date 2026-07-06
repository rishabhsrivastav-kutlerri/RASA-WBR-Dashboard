import { handleUpload } from '@vercel/blob/client';
import { NextResponse } from 'next/server';
import jwt from 'jsonwebtoken';

export const runtime = 'nodejs';

export async function POST(request) {
  const body = await request.json();

  try {
    const jsonResponse = await handleUpload({
      body,
      request,
      onBeforeGenerateToken: async (pathname, clientPayload) => {
        // clientPayload carries the JWT token from the browser.
        const jwtToken = clientPayload || '';
        let payload = null;
        try { payload = jwt.verify(jwtToken, process.env.JWT_SECRET); } catch {}
        if (!payload || payload.role !== 'admin') throw new Error('Forbidden');

        return {
          allowedContentTypes: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
          tokenPayload: JSON.stringify({ pathname }),
        };
      },
      onUploadCompleted: async () => {},
    });
    return NextResponse.json(jsonResponse);
  } catch (err) {
    return NextResponse.json({ error: err.message }, { status: 400 });
  }
}

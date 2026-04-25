import { NextResponse } from 'next/server';
import { generateClientTokenFromReadWriteToken } from '@vercel/blob/client';

export const dynamic = 'force-dynamic';

// Generate a short-lived client token for direct browser → Blob uploads
export async function POST(req: Request): Promise<NextResponse> {
  try {
    const { pathname } = (await req.json()) as { pathname: string };

    if (!pathname) {
      return NextResponse.json({ error: 'Missing pathname' }, { status: 400 });
    }

    const token = process.env.BLOB_READ_WRITE_TOKEN;
    if (!token) {
      return NextResponse.json({ error: 'Blob token not configured' }, { status: 500 });
    }

    const clientToken = await generateClientTokenFromReadWriteToken({
      pathname,
      token,
    });

    return NextResponse.json({ clientToken });
  } catch (error) {
    console.error('Token generation error:', error);
    return NextResponse.json(
      { error: (error as Error).message },
      { status: 500 },
    );
  }
}

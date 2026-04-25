import { NextResponse } from 'next/server';
import { put } from '@vercel/blob';
import type { RawRow } from '@/lib/types';

export const dynamic = 'force-dynamic';

const MAX_CHUNK_ROWS = 8000; // ~2.5MB of JSON per chunk — safely under 4.5MB limit

interface StageRequest {
  sessionId: string;
  fileName: string;
  rows: RawRow[];
  chunkIndex: number;
  totalChunks: number;
}

// Client sends parsed rows in chunks — server stores each chunk in Blob
export async function POST(req: Request) {
  try {
    const body = (await req.json()) as StageRequest;
    const { sessionId, fileName, rows, chunkIndex, totalChunks } = body;

    if (!sessionId || !fileName || !rows) {
      return NextResponse.json({ error: 'Missing sessionId, fileName, or rows' }, { status: 400 });
    }

    const safeName = fileName.replace(/[^a-zA-Z0-9_.-]/g, '_');
    const blobKey = `staging/${sessionId}/${safeName}_chunk${chunkIndex}.json`;

    await put(blobKey, JSON.stringify(rows), {
      access: 'public',
      addRandomSuffix: false,
      contentType: 'application/json',
    });

    return NextResponse.json({
      ok: true,
      chunkIndex,
      totalChunks,
      rowsStaged: rows.length,
      blobKey,
    });
  } catch (e) {
    const msg = e instanceof Error ? e.message : 'Stage failed';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}

export { MAX_CHUNK_ROWS };

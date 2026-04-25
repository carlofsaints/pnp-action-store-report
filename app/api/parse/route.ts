import { NextResponse } from 'next/server';
import { parseVendorFile } from '@/lib/excel-parser';
import type { RawRow, FileInfo, ParseResponse } from '@/lib/types';

export const dynamic = 'force-dynamic';

export async function POST(req: Request) {
  try {
    const { blobUrl, fileName } = (await req.json()) as { blobUrl: string; fileName: string };

    if (!blobUrl || !fileName) {
      return NextResponse.json({ error: 'Missing blobUrl or fileName' }, { status: 400 });
    }

    // Fetch file from Blob storage (already uploaded by client)
    const fileRes = await fetch(blobUrl);
    if (!fileRes.ok) {
      return NextResponse.json({ error: `Failed to fetch from Blob: ${fileRes.status}` }, { status: 500 });
    }

    const buffer = Buffer.from(await fileRes.arrayBuffer());
    const { rows, info } = parseVendorFile(buffer, fileName);

    const allRows: RawRow[] = rows;
    const reportDate = info.reportDate || new Date().toISOString().split('T')[0];

    const response: ParseResponse = {
      files: [info],
      allRows,
      reportDate,
      totalRows: allRows.length,
      uniqueStores: new Set(allRows.map(r => r.siteCode)).size,
      uniqueVendors: new Set(allRows.map(r => r.vendorName)).size,
      uniqueProducts: new Set(allRows.map(r => r.articleNumber)).size,
    };

    return NextResponse.json(response);
  } catch (e) {
    const msg = e instanceof Error ? e.message : 'Parse failed';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}

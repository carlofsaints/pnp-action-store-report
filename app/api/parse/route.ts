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

    // Fetch file from Blob storage
    const fileRes = await fetch(blobUrl);
    if (!fileRes.ok) {
      return NextResponse.json({ error: `Failed to fetch file from storage: ${fileRes.status}` }, { status: 500 });
    }
    const buffer = Buffer.from(await fileRes.arrayBuffer());

    const { rows, info } = parseVendorFile(buffer, fileName);

    const allRows: RawRow[] = rows;
    const fileInfos: FileInfo[] = [info];
    const reportDate = info.reportDate || new Date().toISOString().split('T')[0];

    const uniqueStores = new Set(allRows.map(r => r.siteCode)).size;
    const uniqueVendors = new Set(allRows.map(r => r.vendorName)).size;
    const uniqueProducts = new Set(allRows.map(r => r.articleNumber)).size;

    const response: ParseResponse = {
      files: fileInfos,
      allRows,
      reportDate,
      totalRows: allRows.length,
      uniqueStores,
      uniqueVendors,
      uniqueProducts,
    };

    return NextResponse.json(response);
  } catch (e) {
    const msg = e instanceof Error ? e.message : 'Parse failed';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}

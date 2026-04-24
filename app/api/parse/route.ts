import { NextResponse } from 'next/server';
import { parseVendorFile } from '@/lib/excel-parser';
import type { RawRow, FileInfo, ParseResponse } from '@/lib/types';

export const dynamic = 'force-dynamic';

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const files = formData.getAll('files') as File[];

    if (files.length === 0) {
      return NextResponse.json({ error: 'No files uploaded' }, { status: 400 });
    }

    const allRows: RawRow[] = [];
    const fileInfos: FileInfo[] = [];
    let reportDate = '';

    for (const file of files) {
      const buffer = Buffer.from(await file.arrayBuffer());
      const { rows, info } = parseVendorFile(buffer, file.name);
      allRows.push(...rows);
      fileInfos.push(info);
      if (!reportDate && info.reportDate) reportDate = info.reportDate;
    }

    // Use the most common report date across files
    if (!reportDate) {
      reportDate = new Date().toISOString().split('T')[0];
    }

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

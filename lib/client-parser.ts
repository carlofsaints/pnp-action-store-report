/**
 * Client-side Excel parser — runs in the browser using SheetJS.
 * Mirrors the server-side parseVendorFile() but uses ArrayBuffer instead of Buffer.
 */
import * as XLSX from 'xlsx';
import type { RawRow, FileInfo } from './types';

const EXPECTED_HEADERS = [
  'week ending date', 'site code', 'site description', 'site profile',
  'article number', 'article description', 'vendor number',
  'site article status', 'listing status', 'rp (mrp) type',
  'soh qty', 'dros qty', 'days cover', 'source of supply',
  'date last received', 'date last sold', 'last ordered date',
];

function normalizeDate(val: unknown): string {
  if (val === null || val === undefined || val === '') return '';
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const ddmm = s.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/);
  if (ddmm) {
    return `${ddmm[3]}-${ddmm[2].padStart(2, '0')}-${ddmm[1].padStart(2, '0')}`;
  }
  const n = Number(val);
  if (!isNaN(n) && n > 1000 && n < 100000) {
    const ms = (n - 25569) * 86400 * 1000;
    const d = new Date(ms);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
  }
  if (val instanceof Date && !isNaN(val.getTime())) {
    return `${val.getUTCFullYear()}-${String(val.getUTCMonth() + 1).padStart(2, '0')}-${String(val.getUTCDate()).padStart(2, '0')}`;
  }
  return s;
}

function parseFilename(filename: string): { vendorName: string; reportDate: string } {
  const match = filename.match(/^(.+?)\s+SDC\s+(\d{4}-\d{2}-\d{2})\.xlsx$/i);
  if (match) return { vendorName: match[1].trim(), reportDate: match[2] };
  return { vendorName: filename.replace(/\.xlsx?$/i, ''), reportDate: new Date().toISOString().split('T')[0] };
}

function safeNum(v: unknown): number {
  if (v === null || v === undefined || v === '') return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function colGetter(row: Record<string, unknown>): (target: string) => unknown {
  const map = new Map<string, string>();
  for (const key of Object.keys(row)) map.set(key.trim().toLowerCase(), key);
  return (target: string) => {
    const k = map.get(target.toLowerCase());
    return k ? row[k] : undefined;
  };
}

/** Parse a vendor Excel file in the browser — accepts ArrayBuffer from FileReader */
export function parseVendorFileClient(arrayBuffer: ArrayBuffer, filename: string): { rows: RawRow[]; info: FileInfo } {
  const { vendorName, reportDate } = parseFilename(filename);
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) throw new Error(`No sheets found in ${filename}`);

  const sheet = wb.Sheets[sheetName];

  // Fix corrupt !ref
  const cellKeys = Object.keys(sheet).filter(k => !k.startsWith('!'));
  if (cellKeys.length > 0) {
    let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
    for (const k of cellKeys) {
      const d = XLSX.utils.decode_cell(k);
      if (d.r < minR) minR = d.r;
      if (d.r > maxR) maxR = d.r;
      if (d.c < minC) minC = d.c;
      if (d.c > maxC) maxC = d.c;
    }
    sheet['!ref'] = XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
  }

  const jsonData = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });

  if (jsonData.length === 0) {
    return {
      rows: [],
      info: { fileName: filename, vendorName, vendorNumber: '', rowCount: 0, storeCount: 0, articleCount: 0, reportDate, warning: 'Empty file — no data rows' },
    };
  }

  const actualHeaders = Object.keys(jsonData[0]).map(h => h.trim().toLowerCase());
  const missing = EXPECTED_HEADERS.filter(h => !actualHeaders.includes(h));
  if (missing.length > 0) throw new Error(`Missing columns in ${filename}: ${missing.join(', ')}`);

  const rows: RawRow[] = jsonData.map((r) => {
    const get = colGetter(r);
    return {
      weekEndingDate: normalizeDate(get('week ending date')),
      siteCode: String(get('site code') ?? '').trim(),
      siteDescription: String(get('site description') ?? '').trim(),
      siteProfile: String(get('site profile') ?? '').trim(),
      articleNumber: String(get('article number') ?? '').trim(),
      articleDescription: String(get('article description') ?? '').trim(),
      vendorNumber: String(get('vendor number') ?? '').trim(),
      siteArticleStatus: String(get('site article status') ?? '').trim(),
      listingStatus: String(get('listing status') ?? '').trim(),
      rpType: String(get('rp (mrp) type') ?? '').trim(),
      sohQty: safeNum(get('soh qty')),
      drosQty: safeNum(get('dros qty')),
      daysCover: safeNum(get('days cover')),
      sourceOfSupply: String(get('source of supply') ?? '').trim(),
      dateLastReceived: normalizeDate(get('date last received')),
      dateLastSold: normalizeDate(get('date last sold')),
      lastOrderedDate: normalizeDate(get('last ordered date')),
      vendorName,
    };
  });

  const info: FileInfo = {
    fileName: filename,
    vendorName,
    vendorNumber: [...new Set(rows.map(r => r.vendorNumber))][0] ?? '',
    rowCount: rows.length,
    storeCount: new Set(rows.map(r => r.siteCode)).size,
    articleCount: new Set(rows.map(r => r.articleNumber)).size,
    reportDate,
  };

  return { rows, info };
}

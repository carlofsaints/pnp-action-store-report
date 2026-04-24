/**
 * Parse PnP SDC vendor Excel files using SheetJS.
 * File naming: "VENDOR_NAME SDC YYYY-MM-DD.xlsx"
 * Columns: Week Ending Date (A), Site Code (B), Site Description (C),
 *   Site Profile (D), Article Number (E), Article Description (F),
 *   Vendor Number (G), Site Article Status (H), Listing Status (I),
 *   RP (MRP) Type (J), SOH Qty (K), DROS Qty (L), Days Cover (M),
 *   Source Of Supply (N), Date Last Received (O), Date Last Sold (P),
 *   Last Ordered Date (Q).
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

/** Convert date value to YYYY-MM-DD string.
 *  Handles: Excel serial numbers, DD/MM/YYYY strings, YYYY-MM-DD strings, Date objects */
export function normalizeDate(val: unknown): string {
  if (val === null || val === undefined || val === '') return '';

  const s = String(val).trim();

  // Already YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  // DD/MM/YYYY format (PnP standard)
  const ddmm = s.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/);
  if (ddmm) {
    const dd = ddmm[1].padStart(2, '0');
    const mm = ddmm[2].padStart(2, '0');
    return `${ddmm[3]}-${mm}-${dd}`;
  }

  // Excel serial number
  const n = Number(val);
  if (!isNaN(n) && n > 1000 && n < 100000) {
    const utcDays = n - 25569;
    const ms = utcDays * 86400 * 1000;
    const d = new Date(ms);
    const yyyy = d.getUTCFullYear();
    const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(d.getUTCDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  // Date object (SheetJS sometimes returns these)
  if (val instanceof Date && !isNaN(val.getTime())) {
    const yyyy = val.getUTCFullYear();
    const mm = String(val.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(val.getUTCDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  return s;
}

/** Extract vendor name and report date from filename */
export function parseFilename(filename: string): { vendorName: string; reportDate: string } {
  const match = filename.match(/^(.+?)\s+SDC\s+(\d{4}-\d{2}-\d{2})\.xlsx$/i);
  if (match) {
    return { vendorName: match[1].trim(), reportDate: match[2] };
  }
  const name = filename.replace(/\.xlsx?$/i, '');
  return { vendorName: name, reportDate: new Date().toISOString().split('T')[0] };
}

/** Safe number parse — returns 0 for blanks/non-numeric */
function safeNum(v: unknown): number {
  if (v === null || v === undefined || v === '') return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

/** Case-insensitive column getter from a row object */
function colGetter(row: Record<string, unknown>): (target: string) => unknown {
  // Build a lowercase→key map once per row
  const map = new Map<string, string>();
  for (const key of Object.keys(row)) {
    map.set(key.trim().toLowerCase(), key);
  }
  return (target: string) => {
    const actualKey = map.get(target.toLowerCase());
    return actualKey ? row[actualKey] : undefined;
  };
}

/** Parse a single vendor Excel file buffer into RawRow[] */
export function parseVendorFile(buffer: Buffer, filename: string): { rows: RawRow[]; info: FileInfo } {
  const { vendorName, reportDate } = parseFilename(filename);

  const wb = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) throw new Error(`No sheets found in ${filename}`);

  const sheet = wb.Sheets[sheetName];

  // Fix corrupt !ref — recalculate actual range from cell keys
  const cellKeys = Object.keys(sheet).filter(k => !k.startsWith('!'));
  if (cellKeys.length > 0) {
    const decoded = cellKeys.map(k => XLSX.utils.decode_cell(k));
    const minR = Math.min(...decoded.map(d => d.r));
    const maxR = Math.max(...decoded.map(d => d.r));
    const minC = Math.min(...decoded.map(d => d.c));
    const maxC = Math.max(...decoded.map(d => d.c));
    sheet['!ref'] = XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
  }

  const jsonData = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });

  if (jsonData.length === 0) throw new Error(`No data rows in ${filename}`);

  // Validate headers — case-insensitive
  const actualHeaders = Object.keys(jsonData[0]).map(h => h.trim().toLowerCase());
  const missing = EXPECTED_HEADERS.filter(h => !actualHeaders.includes(h));
  if (missing.length > 0) {
    throw new Error(`Missing columns in ${filename}: ${missing.join(', ')}`);
  }

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

  const uniqueStores = new Set(rows.map(r => r.siteCode));
  const uniqueArticles = new Set(rows.map(r => r.articleNumber));
  const vendorNumbers = new Set(rows.map(r => r.vendorNumber));

  const info: FileInfo = {
    fileName: filename,
    vendorName,
    vendorNumber: [...vendorNumbers][0] ?? '',
    rowCount: rows.length,
    storeCount: uniqueStores.size,
    articleCount: uniqueArticles.size,
    reportDate,
  };

  return { rows, info };
}

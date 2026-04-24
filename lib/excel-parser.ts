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
  'Week Ending Date', 'Site Code', 'Site Description', 'Site Profile',
  'Article Number', 'Article Description', 'Vendor Number',
  'Site Article Status', 'Listing Status', 'RP (MRP) Type',
  'SOH Qty', 'DROS Qty', 'Days Cover', 'Source Of Supply',
  'Date Last Received', 'Date Last Sold', 'Last Ordered Date',
];

/** Convert Excel serial number to YYYY-MM-DD string */
export function excelSerialToDate(serial: unknown): string {
  if (serial === null || serial === undefined || serial === '') return '';
  const n = Number(serial);
  if (isNaN(n) || n <= 0) return String(serial);
  // Excel epoch: 1900-01-01 is serial 1 (with the Lotus 1-2-3 leap year bug)
  const utcDays = n - 25569; // offset from Unix epoch
  const ms = utcDays * 86400 * 1000;
  const d = new Date(ms);
  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

/** Extract vendor name and report date from filename */
export function parseFilename(filename: string): { vendorName: string; reportDate: string } {
  // Pattern: "VENDOR_NAME SDC YYYY-MM-DD.xlsx"
  const match = filename.match(/^(.+?)\s+SDC\s+(\d{4}-\d{2}-\d{2})\.xlsx$/i);
  if (match) {
    return { vendorName: match[1].trim(), reportDate: match[2] };
  }
  // Fallback: use whole filename minus extension as vendor, today as date
  const name = filename.replace(/\.xlsx?$/i, '');
  return { vendorName: name, reportDate: new Date().toISOString().split('T')[0] };
}

/** Safe number parse — returns 0 for blanks/non-numeric */
function safeNum(v: unknown): number {
  if (v === null || v === undefined || v === '') return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

/** Parse a single vendor Excel file buffer into RawRow[] */
export function parseVendorFile(buffer: Buffer, filename: string): { rows: RawRow[]; info: FileInfo } {
  const { vendorName, reportDate } = parseFilename(filename);

  const wb = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) throw new Error(`No sheets found in ${filename}`);

  const sheet = wb.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });

  if (jsonData.length === 0) throw new Error(`No data rows in ${filename}`);

  // Validate headers
  const actualHeaders = Object.keys(jsonData[0]);
  const missing = EXPECTED_HEADERS.filter(h => !actualHeaders.some(a => a.trim() === h));
  if (missing.length > 0) {
    throw new Error(`Missing columns in ${filename}: ${missing.join(', ')}`);
  }

  const rows: RawRow[] = jsonData.map((r) => ({
    weekEndingDate: excelSerialToDate(r['Week Ending Date']),
    siteCode: String(r['Site Code'] ?? '').trim(),
    siteDescription: String(r['Site Description'] ?? '').trim(),
    siteProfile: String(r['Site Profile'] ?? '').trim(),
    articleNumber: String(r['Article Number'] ?? '').trim(),
    articleDescription: String(r['Article Description'] ?? '').trim(),
    vendorNumber: String(r['Vendor Number'] ?? '').trim(),
    siteArticleStatus: String(r['Site Article Status'] ?? '').trim(),
    listingStatus: String(r['Listing Status'] ?? '').trim(),
    rpType: String(r['RP (MRP) Type'] ?? '').trim(),
    sohQty: safeNum(r['SOH Qty']),
    drosQty: safeNum(r['DROS Qty']),
    daysCover: safeNum(r['Days Cover']),
    sourceOfSupply: String(r['Source Of Supply'] ?? '').trim(),
    dateLastReceived: excelSerialToDate(r['Date Last Received']),
    dateLastSold: excelSerialToDate(r['Date Last Sold']),
    lastOrderedDate: excelSerialToDate(r['Last Ordered Date']),
    vendorName,
  }));

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

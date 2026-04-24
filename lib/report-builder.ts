/**
 * Build a 4-sheet per-store Excel report using ExcelJS.
 * Sheet 1: Menu & Overview
 * Sheet 2: OOS (Out of Stock)
 * Sheet 3: Phantom Stock
 * Sheet 4: Missing SKUs
 */
import ExcelJS from 'exceljs';
import type { RawRow } from './types';

// ── Colours ──────────────────────────────────────────────────────────────────
const IRAM_GREEN = '7CC042';
const CHARCOAL = '32373C';
const RED_BG = 'FFC7CE';
const RED_FONT = '9C0006';
const LIGHT_GREY = 'F2F2F2';
const WHITE = 'FFFFFF';
const HEADER_FONT = { bold: true, color: { argb: WHITE }, size: 11 };
const HEADER_FILL = (color: string): ExcelJS.FillPattern => ({
  type: 'pattern', pattern: 'solid', fgColor: { argb: color },
});

// ── Helpers ──────────────────────────────────────────────────────────────────

function dateWeeksAgo(reportDate: string, weeks: number): Date {
  const d = new Date(reportDate + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() - weeks * 7);
  return d;
}

function parseDate(dateStr: string): Date | null {
  if (!dateStr) return null;
  const d = new Date(dateStr + 'T00:00:00Z');
  return isNaN(d.getTime()) ? null : d;
}

function styleHeaderRow(ws: ExcelJS.Worksheet, row: number, colCount: number, color = CHARCOAL) {
  const r = ws.getRow(row);
  for (let c = 1; c <= colCount; c++) {
    const cell = r.getCell(c);
    cell.font = HEADER_FONT;
    cell.fill = HEADER_FILL(color);
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = {
      top: { style: 'thin', color: { argb: '999999' } },
      bottom: { style: 'thin', color: { argb: '999999' } },
    };
  }
  r.height = 24;
}

function autoWidth(ws: ExcelJS.Worksheet) {
  ws.columns.forEach((col) => {
    let max = 10;
    col.eachCell?.({ includeEmpty: false }, (cell) => {
      const len = String(cell.value ?? '').length;
      if (len > max) max = len;
    });
    col.width = Math.min(max + 3, 40);
  });
}

function zebraRows(ws: ExcelJS.Worksheet, startRow: number, endRow: number, colCount: number) {
  for (let r = startRow; r <= endRow; r++) {
    if ((r - startRow) % 2 === 1) {
      const row = ws.getRow(r);
      for (let c = 1; c <= colCount; c++) {
        row.getCell(c).fill = HEADER_FILL(LIGHT_GREY);
      }
    }
  }
}

// ── Ranking helpers ──────────────────────────────────────────────────────────

export interface RankingData {
  /** Overall store rank: Map<siteCode, { rank, total, totalDros }> */
  overallRanks: Map<string, { rank: number; total: number; totalDros: number }>;
  /** Per-vendor store rank: Map<"vendorNum|siteCode", { rank, total }> */
  vendorRanks: Map<string, { rank: number; total: number }>;
  /** Master product list per vendor: Map<vendorNum, { articleNum, articleDesc, vendorName }[]> */
  masterProducts: Map<string, { articleNumber: string; articleDescription: string; vendorName: string }[]>;
}

export function computeRankings(allRows: RawRow[]): RankingData {
  // ── Overall rank by sum of DROS Qty per Site Code ──
  const storeDros = new Map<string, number>();
  for (const r of allRows) {
    storeDros.set(r.siteCode, (storeDros.get(r.siteCode) ?? 0) + r.drosQty);
  }
  const sortedStores = [...storeDros.entries()].sort((a, b) => b[1] - a[1]);
  const overallRanks = new Map<string, { rank: number; total: number; totalDros: number }>();
  const totalStores = sortedStores.length;
  sortedStores.forEach(([code, dros], idx) => {
    overallRanks.set(code, { rank: idx + 1, total: totalStores, totalDros: dros });
  });

  // ── Per-vendor rank ──
  const vendorStoreDros = new Map<string, Map<string, number>>(); // vendorNum → Map<siteCode, dros>
  for (const r of allRows) {
    if (!vendorStoreDros.has(r.vendorNumber)) vendorStoreDros.set(r.vendorNumber, new Map());
    const m = vendorStoreDros.get(r.vendorNumber)!;
    m.set(r.siteCode, (m.get(r.siteCode) ?? 0) + r.drosQty);
  }
  const vendorRanks = new Map<string, { rank: number; total: number }>();
  for (const [vendor, storeMap] of vendorStoreDros) {
    const sorted = [...storeMap.entries()].sort((a, b) => b[1] - a[1]);
    sorted.forEach(([code], idx) => {
      vendorRanks.set(`${vendor}|${code}`, { rank: idx + 1, total: sorted.length });
    });
  }

  // ── Master product list per vendor ──
  const masterProducts = new Map<string, { articleNumber: string; articleDescription: string; vendorName: string }[]>();
  const seen = new Set<string>();
  for (const r of allRows) {
    const key = `${r.vendorNumber}|${r.articleNumber}`;
    if (seen.has(key)) continue;
    seen.add(key);
    if (!masterProducts.has(r.vendorNumber)) masterProducts.set(r.vendorNumber, []);
    masterProducts.get(r.vendorNumber)!.push({
      articleNumber: r.articleNumber,
      articleDescription: r.articleDescription,
      vendorName: r.vendorName,
    });
  }

  return { overallRanks, vendorRanks, masterProducts };
}

// ── Main builder ─────────────────────────────────────────────────────────────

export interface BuildParams {
  storeRows: RawRow[];
  allRows: RawRow[];
  rankings: RankingData;
  reportDate: string;
  phantomWeeksReceived: number;
  phantomWeeksSold: number;
  siteCode: string;
  storeName: string;
}

export async function buildStoreReport(params: BuildParams): Promise<Buffer> {
  const {
    storeRows, allRows, rankings, reportDate,
    phantomWeeksReceived, phantomWeeksSold, siteCode, storeName,
  } = params;

  const wb = new ExcelJS.Workbook();
  wb.creator = 'PnP Action Store Report — OuterJoin';
  wb.created = new Date();

  // ── Sheet 1: Menu & Overview ──────────────────────────────────────────────
  const ws1 = wb.addWorksheet('Menu & Overview');

  // Store header
  const rank = rankings.overallRanks.get(siteCode);
  ws1.mergeCells('A1:F1');
  const titleCell = ws1.getCell('A1');
  titleCell.value = storeName;
  titleCell.font = { bold: true, size: 16, color: { argb: CHARCOAL } };
  titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
  ws1.getRow(1).height = 30;

  ws1.mergeCells('A2:F2');
  const rankCell = ws1.getCell('A2');
  rankCell.value = rank
    ? `Store Rank: ${rank.rank} / ${rank.total}  |  Report Date: ${reportDate}`
    : `Report Date: ${reportDate}`;
  rankCell.font = { size: 11, color: { argb: '666666' } };

  // KPI section
  const uniqueVendors = new Set(storeRows.map(r => r.vendorName)).size;
  const uniqueProducts = new Set(storeRows.map(r => r.articleNumber)).size;
  const sohPositive = storeRows.filter(r => r.sohQty > 0).length;
  const sohZeroOrNeg = storeRows.filter(r => r.sohQty <= 0).length;

  const kpiStartRow = 4;
  ws1.getCell(`A${kpiStartRow}`).value = 'Metric';
  ws1.getCell(`B${kpiStartRow}`).value = 'Value';
  styleHeaderRow(ws1, kpiStartRow, 2, IRAM_GREEN);

  const kpis = [
    ['Total Vendors', uniqueVendors],
    ['Total Products', uniqueProducts],
    ['Products SOH > 0', sohPositive],
    ['Products SOH <= 0', sohZeroOrNeg],
  ];
  kpis.forEach(([label, value], i) => {
    ws1.getCell(`A${kpiStartRow + 1 + i}`).value = label;
    const vCell = ws1.getCell(`B${kpiStartRow + 1 + i}`);
    vCell.value = value;
    vCell.font = { bold: true };
  });
  zebraRows(ws1, kpiStartRow + 1, kpiStartRow + kpis.length, 2);

  // Vendor summary table
  const vendorTableStart = kpiStartRow + kpis.length + 2;
  const vendorHeaders = ['Vendor Name', 'Vendor Number', 'Products', 'SOH > 0', 'SOH <= 0', 'Vendor Rank'];
  vendorHeaders.forEach((h, i) => { ws1.getCell(vendorTableStart, i + 1).value = h; });
  styleHeaderRow(ws1, vendorTableStart, vendorHeaders.length, CHARCOAL);

  // Build vendor summary
  const vendorMap = new Map<string, { vendorNum: string; articles: Set<string>; sohPos: number; sohNeg: number }>();
  for (const r of storeRows) {
    if (!vendorMap.has(r.vendorName)) {
      vendorMap.set(r.vendorName, { vendorNum: r.vendorNumber, articles: new Set(), sohPos: 0, sohNeg: 0 });
    }
    const v = vendorMap.get(r.vendorName)!;
    v.articles.add(r.articleNumber);
    if (r.sohQty > 0) v.sohPos++; else v.sohNeg++;
  }

  let vRow = vendorTableStart + 1;
  for (const [vName, v] of [...vendorMap.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
    const vRank = rankings.vendorRanks.get(`${v.vendorNum}|${siteCode}`);
    ws1.getCell(vRow, 1).value = vName;
    ws1.getCell(vRow, 2).value = v.vendorNum;
    ws1.getCell(vRow, 3).value = v.articles.size;
    ws1.getCell(vRow, 4).value = v.sohPos;
    ws1.getCell(vRow, 5).value = v.sohNeg;
    ws1.getCell(vRow, 6).value = vRank ? `${vRank.rank} / ${vRank.total}` : '—';
    vRow++;
  }
  zebraRows(ws1, vendorTableStart + 1, vRow - 1, vendorHeaders.length);
  autoWidth(ws1);

  // ── Sheet 2: OOS (Out of Stock) ──────────────────────────────────────────
  const ws2 = wb.addWorksheet('OOS');
  const oosHeaders = ['Vendor Name', 'Vendor Number', 'Article Number', 'Article Description',
    'Site Article Status', 'SOH Qty', 'DROS Qty', 'Days Cover'];
  oosHeaders.forEach((h, i) => { ws2.getCell(1, i + 1).value = h; });
  styleHeaderRow(ws2, 1, oosHeaders.length, CHARCOAL);

  const oosRows = storeRows
    .filter(r => r.sohQty <= 0)
    .sort((a, b) => a.vendorName.localeCompare(b.vendorName) || a.articleDescription.localeCompare(b.articleDescription));

  oosRows.forEach((r, i) => {
    const rowNum = i + 2;
    ws2.getCell(rowNum, 1).value = r.vendorName;
    ws2.getCell(rowNum, 2).value = r.vendorNumber;
    ws2.getCell(rowNum, 3).value = r.articleNumber;
    ws2.getCell(rowNum, 4).value = r.articleDescription;
    ws2.getCell(rowNum, 5).value = r.siteArticleStatus;
    ws2.getCell(rowNum, 6).value = r.sohQty;
    ws2.getCell(rowNum, 7).value = r.drosQty;
    ws2.getCell(rowNum, 8).value = r.daysCover;

    // Red highlight for negative SOH
    if (r.sohQty < 0) {
      for (let c = 1; c <= oosHeaders.length; c++) {
        const cell = ws2.getCell(rowNum, c);
        cell.fill = HEADER_FILL(RED_BG);
        cell.font = { color: { argb: RED_FONT } };
      }
    }
  });

  zebraRows(ws2, 2, oosRows.length + 1, oosHeaders.length);
  autoWidth(ws2);

  // ── Sheet 3: Phantom Stock ────────────────────────────────────────────────
  const ws3 = wb.addWorksheet('Phantom Stock');
  const phantomHeaders = ['Vendor Name', 'Vendor Number', 'Article Number', 'Article Description',
    'Site Article Status', 'SOH Qty', 'DROS Qty', 'Days Cover', 'Date Last Received', 'Date Last Sold'];
  phantomHeaders.forEach((h, i) => { ws3.getCell(1, i + 1).value = h; });
  styleHeaderRow(ws3, 1, phantomHeaders.length, CHARCOAL);

  const cutoffReceived = dateWeeksAgo(reportDate, phantomWeeksReceived);
  const cutoffSold = dateWeeksAgo(reportDate, phantomWeeksSold);

  const phantomRows = storeRows.filter(r => {
    if (r.sohQty <= 0) return false;
    const lastRecv = parseDate(r.dateLastReceived);
    const lastSold = parseDate(r.dateLastSold);
    // Must have both dates and both must be before cutoff
    if (!lastRecv || !lastSold) return false;
    return lastRecv < cutoffReceived && lastSold < cutoffSold;
  }).sort((a, b) => a.vendorName.localeCompare(b.vendorName) || a.articleDescription.localeCompare(b.articleDescription));

  phantomRows.forEach((r, i) => {
    const rowNum = i + 2;
    ws3.getCell(rowNum, 1).value = r.vendorName;
    ws3.getCell(rowNum, 2).value = r.vendorNumber;
    ws3.getCell(rowNum, 3).value = r.articleNumber;
    ws3.getCell(rowNum, 4).value = r.articleDescription;
    ws3.getCell(rowNum, 5).value = r.siteArticleStatus;
    ws3.getCell(rowNum, 6).value = r.sohQty;
    ws3.getCell(rowNum, 7).value = r.drosQty;
    ws3.getCell(rowNum, 8).value = r.daysCover;
    ws3.getCell(rowNum, 9).value = r.dateLastReceived;
    ws3.getCell(rowNum, 10).value = r.dateLastSold;
  });

  zebraRows(ws3, 2, phantomRows.length + 1, phantomHeaders.length);
  autoWidth(ws3);

  // ── Sheet 4: Missing SKUs ─────────────────────────────────────────────────
  const ws4 = wb.addWorksheet('Missing SKUs');
  const missHeaders = ['Vendor Name', 'Vendor Number', 'Article Number', 'Article Description',
    'Common Status', 'DROS Avg (Active Stores)'];
  missHeaders.forEach((h, i) => { ws4.getCell(1, i + 1).value = h; });
  styleHeaderRow(ws4, 1, missHeaders.length, CHARCOAL);

  // Products this store HAS
  const storeArticles = new Set(storeRows.map(r => `${r.vendorNumber}|${r.articleNumber}`));

  // Build missing products across all vendors in this store's data
  type MissProduct = {
    vendorName: string; vendorNumber: string; articleNumber: string;
    articleDescription: string; commonStatus: string; drosAvg: number;
  };
  const missingProducts: MissProduct[] = [];

  // Get the unique vendors this store SHOULD have (all vendors in the dataset)
  const allVendorNums = new Set(storeRows.map(r => r.vendorNumber));

  for (const vendorNum of allVendorNums) {
    const masterList = rankings.masterProducts.get(vendorNum) ?? [];
    for (const mp of masterList) {
      const key = `${vendorNum}|${mp.articleNumber}`;
      if (storeArticles.has(key)) continue; // store already has it

      // Find the most common Site Article Status across other stores
      const otherRows = allRows.filter(r => r.vendorNumber === vendorNum && r.articleNumber === mp.articleNumber);
      const statusCounts = new Map<string, number>();
      let drosSum = 0;
      let drosCount = 0;
      for (const r of otherRows) {
        statusCounts.set(r.siteArticleStatus, (statusCounts.get(r.siteArticleStatus) ?? 0) + 1);
        if (r.sohQty > 0) {
          drosSum += r.drosQty;
          drosCount++;
        }
      }
      const commonStatus = [...statusCounts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? '';
      const drosAvg = drosCount > 0 ? Math.round((drosSum / drosCount) * 100) / 100 : 0;

      missingProducts.push({
        vendorName: mp.vendorName,
        vendorNumber: vendorNum,
        articleNumber: mp.articleNumber,
        articleDescription: mp.articleDescription,
        commonStatus,
        drosAvg,
      });
    }
  }

  // Sort by vendor, then DROS avg descending
  missingProducts.sort((a, b) =>
    a.vendorName.localeCompare(b.vendorName) || b.drosAvg - a.drosAvg
  );

  missingProducts.forEach((mp, i) => {
    const rowNum = i + 2;
    ws4.getCell(rowNum, 1).value = mp.vendorName;
    ws4.getCell(rowNum, 2).value = mp.vendorNumber;
    ws4.getCell(rowNum, 3).value = mp.articleNumber;
    ws4.getCell(rowNum, 4).value = mp.articleDescription;
    ws4.getCell(rowNum, 5).value = mp.commonStatus;
    ws4.getCell(rowNum, 6).value = mp.drosAvg;
  });

  zebraRows(ws4, 2, missingProducts.length + 1, missHeaders.length);
  autoWidth(ws4);

  // ── Write buffer ──────────────────────────────────────────────────────────
  const arrayBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

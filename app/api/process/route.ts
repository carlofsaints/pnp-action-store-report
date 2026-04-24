import { NextResponse } from 'next/server';
import { Resend } from 'resend';
import * as XLSX from 'xlsx';
import { parseVendorFile } from '@/lib/excel-parser';
import { buildStoreReport, computeRankings } from '@/lib/report-builder';
import { buildStoreEmail } from '@/lib/email-builder';
import { getDriveContext, uploadReport, downloadFile } from '@/lib/graph-iram';
import type { RawRow, ControlEntry, ProcessSummary, StoreResult } from '@/lib/types';

export const dynamic = 'force-dynamic';
export const maxDuration = 300; // 5 min for large batches

const CONTROL_FILE_PATH = process.env.PNP_CONTROL_FILE_PATH ?? 'PNP ACTION STORE REPORTS (MULTI VENDOR)/CONTROL FILES/iRam PNP REP STORE ALLOCATION.xlsx';

// ── Parse control file from SP ──────────────────────────────────────────────

async function loadControlFile(): Promise<Map<string, ControlEntry>> {
  const map = new Map<string, ControlEntry>();
  try {
    const ctx = await getDriveContext();
    const buffer = await downloadFile(CONTROL_FILE_PATH, ctx);
    const wb = XLSX.read(buffer, { type: 'buffer' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    if (!sheet) return map;

    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });

    for (const r of rows) {
      const get = (target: string): string => {
        const key = Object.keys(r).find(k => k.trim().toLowerCase() === target.toLowerCase());
        return key ? String(r[key] ?? '').trim() : '';
      };

      const siteCode = get('Site Code');
      const siteName = get('Site Name');
      const channel = get('Channel');
      const repName = get('Rep Name');
      const repEmail = get('Rep Email');

      if (siteCode) {
        map.set(siteCode, { siteCode, siteName, channel, repName, repEmail });
      }
    }
  } catch (e) {
    console.error('Control file load error:', e);
  }
  return map;
}

// ── Count phantom rows for a store ──────────────────────────────────────────

function countPhantom(rows: RawRow[], reportDate: string, weeksReceived: number, weeksSold: number): number {
  const cutRecv = new Date(reportDate + 'T00:00:00Z');
  cutRecv.setUTCDate(cutRecv.getUTCDate() - weeksReceived * 7);
  const cutSold = new Date(reportDate + 'T00:00:00Z');
  cutSold.setUTCDate(cutSold.getUTCDate() - weeksSold * 7);

  return rows.filter(r => {
    if (r.sohQty <= 0) return false;
    const lr = r.dateLastReceived ? new Date(r.dateLastReceived + 'T00:00:00Z') : null;
    const ls = r.dateLastSold ? new Date(r.dateLastSold + 'T00:00:00Z') : null;
    if (!lr || !ls) return false;
    return lr < cutRecv && ls < cutSold;
  }).length;
}

// ── Count missing SKUs for a store ──────────────────────────────────────────

function countMissing(storeRows: RawRow[], allRows: RawRow[]): number {
  const storeArticles = new Set(storeRows.map(r => `${r.vendorNumber}|${r.articleNumber}`));
  const allVendorNums = new Set(storeRows.map(r => r.vendorNumber));
  const masterProducts = new Map<string, Set<string>>();

  for (const r of allRows) {
    if (!allVendorNums.has(r.vendorNumber)) continue;
    if (!masterProducts.has(r.vendorNumber)) masterProducts.set(r.vendorNumber, new Set());
    masterProducts.get(r.vendorNumber)!.add(r.articleNumber);
  }

  let count = 0;
  for (const [vendor, articles] of masterProducts) {
    for (const art of articles) {
      if (!storeArticles.has(`${vendor}|${art}`)) count++;
    }
  }
  return count;
}

// ── Main handler — receives FormData with raw files ─────────────────────────

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const files = formData.getAll('files') as File[];
    const reportDate = formData.get('reportDate') as string || new Date().toISOString().split('T')[0];
    const phantomWeeksReceived = Number(formData.get('phantomWeeksReceived')) || 4;
    const phantomWeeksSold = Number(formData.get('phantomWeeksSold')) || 4;
    const actionMode = (formData.get('actionMode') as string) || 'both';

    if (files.length === 0) {
      return NextResponse.json({ error: 'No files provided' }, { status: 400 });
    }

    // Re-parse all files server-side
    const allRows: RawRow[] = [];
    const parseErrors: string[] = [];

    for (const file of files) {
      try {
        const buffer = Buffer.from(await file.arrayBuffer());
        const { rows } = parseVendorFile(buffer, file.name);
        allRows.push(...rows);
      } catch (e) {
        parseErrors.push(`${file.name}: ${e instanceof Error ? e.message : String(e)}`);
      }
    }

    if (allRows.length === 0) {
      return NextResponse.json({ error: `No data rows from any file. Errors: ${parseErrors.join('; ')}` }, { status: 400 });
    }

    // Compute global rankings
    const rankings = computeRankings(allRows);

    const storeResults: StoreResult[] = [];
    const errors: string[] = [...parseErrors];

    // Load control file (for email mode)
    const controlMap = (actionMode === 'sharepoint')
      ? new Map<string, ControlEntry>()
      : await loadControlFile();

    // Get SP context once
    let spCtx: Awaited<ReturnType<typeof getDriveContext>> | null = null;
    if (actionMode !== 'email') {
      try {
        spCtx = await getDriveContext();
      } catch (e) {
        const msg = `SharePoint connection failed: ${e instanceof Error ? e.message : String(e)}`;
        console.error(msg);
        errors.push(msg);
      }
    }

    // Group rows by Site Code
    const storeMap = new Map<string, RawRow[]>();
    for (const r of allRows) {
      if (!storeMap.has(r.siteCode)) storeMap.set(r.siteCode, []);
      storeMap.get(r.siteCode)!.push(r);
    }

    const resend = process.env.RESEND_API_KEY ? new Resend(process.env.RESEND_API_KEY) : null;
    let emailsSent = 0;
    let spUploaded = 0;
    let reportsGenerated = 0;

    // Process each store
    for (const [siteCode, storeRows] of storeMap) {
      const storeName = storeRows[0]?.siteDescription ?? siteCode;
      const result: StoreResult = {
        siteCode,
        storeName,
        oosCount: storeRows.filter(r => r.sohQty <= 0).length,
        phantomCount: countPhantom(storeRows, reportDate, phantomWeeksReceived, phantomWeeksSold),
        missingCount: countMissing(storeRows, allRows),
        repEmail: null,
        emailed: false,
        uploaded: false,
        error: null,
      };

      try {
        // Build report
        const reportBuffer = await buildStoreReport({
          storeRows,
          allRows,
          rankings,
          reportDate,
          phantomWeeksReceived,
          phantomWeeksSold,
          siteCode,
          storeName,
        });
        reportsGenerated++;

        const safeStore = storeName.replace(/[/\\?%*:|"<>]/g, '_');
        const fileName = `PNP - ${siteCode} - ${safeStore} - ${reportDate}.xlsx`;

        // Upload to SharePoint
        if (actionMode !== 'email' && spCtx) {
          try {
            await uploadReport(reportBuffer, reportDate, fileName, spCtx);
            result.uploaded = true;
            spUploaded++;
          } catch (e) {
            const msg = `SP upload failed for ${siteCode}: ${e instanceof Error ? e.message : String(e)}`;
            errors.push(msg);
            result.error = msg;
          }
        }

        // Send email
        if (actionMode !== 'sharepoint' && resend) {
          const ctrl = controlMap.get(siteCode);
          if (ctrl && ctrl.repEmail) {
            result.repEmail = ctrl.repEmail;
            const rank = rankings.overallRanks.get(siteCode);
            const repFullName = ctrl.repName;

            try {
              const html = buildStoreEmail({
                storeName,
                siteCode,
                rank: rank ? `${rank.rank} / ${rank.total}` : '—',
                reportDate,
                repName: repFullName || 'Team',
                oosCount: result.oosCount,
                phantomCount: result.phantomCount,
                missingCount: result.missingCount,
                totalProducts: new Set(storeRows.map(r => r.articleNumber)).size,
                sohPositive: storeRows.filter(r => r.sohQty > 0).length,
                sohZeroOrNeg: storeRows.filter(r => r.sohQty <= 0).length,
              });

              await resend.emails.send({
                from: 'PnP Action Store Report <noreply@outerjoin.co.za>',
                to: ctrl.repEmail,
                subject: `PnP Action Store Report – ${storeName} – ${reportDate}`,
                html,
                attachments: [
                  {
                    filename: fileName,
                    content: reportBuffer.toString('base64'),
                  },
                ],
              });
              result.emailed = true;
              emailsSent++;
            } catch (e) {
              const msg = `Email failed for ${siteCode} (${ctrl.repEmail}): ${e instanceof Error ? e.message : String(e)}`;
              errors.push(msg);
              if (!result.error) result.error = msg;
            }
          }
        }
      } catch (e) {
        const msg = `Report build failed for ${siteCode}: ${e instanceof Error ? e.message : String(e)}`;
        errors.push(msg);
        result.error = msg;
      }

      storeResults.push(result);
    }

    const summary: ProcessSummary = {
      storesProcessed: storeMap.size,
      reportsGenerated,
      emailsSent,
      spUploaded,
      errors,
      storeResults,
    };

    return NextResponse.json(summary, {
      headers: { 'Cache-Control': 'no-store' },
    });
  } catch (e) {
    const msg = e instanceof Error ? e.message : 'Processing failed';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}

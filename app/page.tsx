'use client';

import { useState, useRef, useCallback } from 'react';
import { useAuth, authFetch } from '@/lib/useAuth';
import { parseVendorFileClient } from '@/lib/client-parser';
import type { FileInfo, ProcessSummary, RawRow, StoreResult } from '@/lib/types';

type Stage = 'idle' | 'parsed' | 'processing' | 'done' | 'error';

/** One parsed file's data, kept separately so we can remove individual files */
interface LoadedFile {
  info: FileInfo;
  rows: RawRow[];
  blobUrl: string; // Vercel Blob URL — bypasses 4.5MB serverless limit
}

// ── Tiny UI helpers ──────────────────────────────────────────────────────────

function Badge({ label, value, color = 'accent' }: { label: string; value: string | number; color?: string }) {
  const colorClass =
    color === 'accent' ? 'text-accent' : color === 'success' ? 'text-success' : 'text-warning';
  return (
    <div className="bg-card border border-border rounded-lg p-4 text-center">
      <div className={`text-2xl font-bold ${colorClass}`}>{value}</div>
      <div className="text-muted text-sm mt-1">{label}</div>
    </div>
  );
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="bg-card border border-border rounded-xl p-6 mb-6 shadow-sm">
      <h2 className="text-lg font-bold mb-4 border-b border-border pb-3" style={{ color: '#7CC042' }}>{title}</h2>
      {children}
    </div>
  );
}

// ── Main Page ────────────────────────────────────────────────────────────────

export default function Home() {
  const { session, loading: authLoading, logout } = useAuth();
  const [stage, setStage] = useState<Stage>('idle');
  const [loadedFiles, setLoadedFiles] = useState<LoadedFile[]>([]);
  const [processSummary, setProcessSummary] = useState<ProcessSummary | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState<string | null>(null);

  // Settings
  const [phantomWeeksReceived, setPhantomWeeksReceived] = useState(4);
  const [phantomWeeksSold, setPhantomWeeksSold] = useState(4);
  const [actionMode, setActionMode] = useState<'email' | 'sharepoint' | 'both'>('both');
  const [reportDateOverride, setReportDateOverride] = useState('');

  const fileInputRef = useRef<HTMLInputElement>(null);

  // ── Derived from accumulated files ────────────────────────────────────────

  const allRows = loadedFiles.flatMap(f => f.rows);
  const allFileInfos = loadedFiles.map(f => f.info);
  const hasFiles = loadedFiles.length > 0;

  const totalRows = allRows.length;
  const uniqueStores = new Set(allRows.map(r => r.siteCode)).size;
  const uniqueVendors = new Set(allRows.map(r => r.vendorName)).size;
  const uniqueProducts = new Set(allRows.map(r => r.articleNumber)).size;
  const reportDate = reportDateOverride || allFileInfos[0]?.reportDate || '';

  const vendorBreakdown = hasFiles
    ? (() => {
        const map = new Map<string, { vendorNum: string; products: Set<string>; rows: number }>();
        for (const r of allRows) {
          if (!map.has(r.vendorName)) map.set(r.vendorName, { vendorNum: r.vendorNumber, products: new Set(), rows: 0 });
          const v = map.get(r.vendorName)!;
          v.products.add(r.articleNumber);
          v.rows++;
        }
        return [...map.entries()].map(([name, v]) => ({
          vendorName: name, vendorNumber: v.vendorNum, products: v.products.size, rows: v.rows,
        }));
      })()
    : [];

  // ── File upload + parse (APPENDS to existing files) ───────────────────────

  const handleFiles = useCallback(async (files: FileList | File[]) => {
    const fileArr = Array.from(files).filter(
      (f) => f.name.endsWith('.xlsx') || f.name.endsWith('.xls')
    );
    if (fileArr.length === 0) return;

    setIsUploading(true);
    setUploadProgress(null);
    setErrorMsg(null);

    const fileErrors: string[] = [];
    let anySuccess = false;

    // Parse files entirely in the browser — no server call needed
    for (let i = 0; i < fileArr.length; i++) {
      const file = fileArr[i];
      setUploadProgress(`Parsing file ${i + 1} of ${fileArr.length}: ${file.name}`);

      try {
        // Read file as ArrayBuffer and parse client-side with SheetJS
        const arrayBuffer = await file.arrayBuffer();
        const { rows, info } = parseVendorFileClient(arrayBuffer, file.name);

        if (rows.length === 0) {
          fileErrors.push(`${file.name}: No data rows found`);
          continue;
        }

        // Append to state — check for duplicate filenames
        setLoadedFiles(prev => {
          const existing = prev.map(f => f.info.fileName);
          if (existing.includes(info.fileName)) {
            fileErrors.push(`${file.name}: Duplicate — already loaded`);
            return prev;
          }
          return [...prev, { info, rows, blobUrl: '' }];
        });

        anySuccess = true;

        // Set report date from first successful file if not already set
        if (!reportDateOverride && info.reportDate) {
          setReportDateOverride(info.reportDate);
        }
      } catch (e) {
        fileErrors.push(`${file.name}: ${e instanceof Error ? e.message : 'Unknown error'}`);
      }
    }

    if (anySuccess) {
      setStage('parsed');
    }

    if (fileErrors.length > 0) {
      setErrorMsg(`${fileErrors.length} file(s) skipped:\n${fileErrors.join('\n')}`);
    }

    setIsUploading(false);
    setUploadProgress(null);
  }, [reportDateOverride]);

  // ── Remove a single file ──────────────────────────────────────────────────

  const removeFile = (fileName: string) => {
    setLoadedFiles(prev => {
      const updated = prev.filter(f => f.info.fileName !== fileName);
      if (updated.length === 0) {
        setStage('idle');
        setReportDateOverride('');
      }
      return updated;
    });
  };

  const clearAll = () => {
    setLoadedFiles([]);
    setStage('idle');
    setProcessSummary(null);
    setErrorMsg(null);
    setReportDateOverride('');
  };

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      void handleFiles(e.dataTransfer.files);
    },
    [handleFiles]
  );

  const onDragOver = (e: React.DragEvent) => { e.preventDefault(); setIsDragging(true); };
  const onDragLeave = () => setIsDragging(false);

  // ── Process ──────────────────────────────────────────────────────────────

  const handleProcess = async () => {
    if (loadedFiles.length === 0) return;

    setStage('processing');
    setErrorMsg(null);

    try {
      // Step 1: Stage all rows in Blob via chunked uploads (each chunk <4MB)
      const sessionId = `s-${Date.now().toString(36)}`;
      const CHUNK_SIZE = 8000; // rows per chunk — ~2.5MB JSON each

      setUploadProgress('Staging data for processing...');
      for (const lf of loadedFiles) {
        const totalChunks = Math.ceil(lf.rows.length / CHUNK_SIZE);
        for (let c = 0; c < totalChunks; c++) {
          const chunk = lf.rows.slice(c * CHUNK_SIZE, (c + 1) * CHUNK_SIZE);
          const stageRes = await authFetch('/api/stage', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              sessionId,
              fileName: lf.info.fileName,
              rows: chunk,
              chunkIndex: c,
              totalChunks,
            }),
          });
          if (!stageRes.ok) {
            const text = await stageRes.text();
            throw new Error(`Staging failed for ${lf.info.fileName}: ${text.slice(0, 200)}`);
          }
        }
      }
      setUploadProgress(null);

      // Step 2: Trigger processing — tiny JSON payload with just session ID + settings
      const res = await authFetch('/api/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sessionId,
          filesMeta: loadedFiles.map(lf => ({
            fileName: lf.info.fileName,
            vendorName: lf.info.vendorName,
            rowCount: lf.info.rowCount,
          })),
          reportDate,
          phantomWeeksReceived,
          phantomWeeksSold,
          actionMode,
        }),
      });

      if (!res.ok) {
        const text = await res.text();
        throw new Error(`Server error ${res.status}: ${text.slice(0, 300)}`);
      }

      const data = await res.json() as ProcessSummary;
      setProcessSummary(data);
      setStage('done');
    } catch (e) {
      setErrorMsg(e instanceof Error ? e.message : 'Processing failed');
      setStage('error');
    }
  };

  // ── Render ────────────────────────────────────────────────────────────────

  if (authLoading) {
    return <div className="min-h-screen bg-background flex items-center justify-center text-muted">Loading...</div>;
  }

  return (
    <div className="min-h-screen bg-background text-foreground">
      {/* Header */}
      <header className="border-b border-border px-6 py-3 flex items-center justify-between sticky top-0 bg-card z-10 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="w-1.5 h-8 rounded" style={{ background: '#7CC042' }} />
          <div>
            <h1 className="text-xl font-bold text-foreground">PnP Action Store Report</h1>
            <p className="text-muted text-xs">OOS, phantom stock &amp; missing SKU analysis &mdash; iRam</p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          {hasFiles && stage === 'parsed' && (
            <span className="bg-accent/10 text-accent border border-accent/30 text-xs px-3 py-1 rounded-full">
              {loadedFiles.length} file{loadedFiles.length !== 1 ? 's' : ''} loaded
            </span>
          )}
          {stage === 'processing' && (
            <span className="bg-warning/10 text-warning border border-warning/30 text-xs px-3 py-1 rounded-full animate-pulse">
              Processing...
            </span>
          )}
          {stage === 'done' && (
            <span className="bg-success/10 text-success border border-success/30 text-xs px-3 py-1 rounded-full">
              Complete
            </span>
          )}
          {session && (
            <div className="flex items-center gap-3">
              {session.role === 'admin' && (
                <a href="/admin" className="text-accent text-xs hover:underline">Admin</a>
              )}
              <span className="text-muted text-xs">{session.name}</span>
              <button onClick={logout} className="text-muted text-xs hover:text-danger underline">Logout</button>
            </div>
          )}
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img src="/iram-logo.png" alt="iRam" className="h-9 w-auto object-contain" />
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-6 py-8">

        {/* ── Section 1: Upload Files ── */}
        <Section title="1 — Upload Vendor Files">
          <p className="text-muted text-xs mb-3">
            Upload PnP Portal SDC export files one at a time or in batches. Each upload adds to the dataset.
            File naming: <code>VENDOR_NAME SDC YYYY-MM-DD.xlsx</code>
          </p>
          <div
            onDrop={onDrop}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onClick={() => fileInputRef.current?.click()}
            className={`border-2 border-dashed rounded-lg p-10 text-center cursor-pointer transition-colors ${
              isDragging ? 'border-accent bg-accent/5' : 'border-border hover:border-accent/50'
            }`}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              multiple
              className="hidden"
              onChange={(e) => { if (e.target.files) void handleFiles(e.target.files); e.target.value = ''; }}
            />
            {isUploading ? (
              <p className="text-accent animate-pulse">{uploadProgress ?? 'Parsing files...'}</p>
            ) : (
              <>
                <p className="text-foreground font-medium">
                  {hasFiles ? 'Drop more files here or click to add' : 'Drop Excel files here or click to browse'}
                </p>
                <p className="text-muted text-sm mt-1">Accepts .xlsx files &mdash; one per vendor</p>
              </>
            )}
          </div>

          {/* File list — with remove buttons */}
          {loadedFiles.length > 0 && (
            <div className="mt-4 space-y-2">
              {loadedFiles.map((lf, i) => (
                <div key={i} className="bg-background border border-border rounded-lg px-4 py-3 flex items-center justify-between">
                  <div className="min-w-0">
                    <span className="text-foreground font-medium text-sm block truncate">{lf.info.fileName}</span>
                    <span className="text-accent text-xs font-mono">{lf.info.vendorName}</span>
                  </div>
                  <div className="flex items-center gap-3">
                    <div className="text-right text-xs text-muted">
                      <div>{lf.info.rowCount.toLocaleString()} rows &middot; {lf.info.storeCount} stores &middot; {lf.info.articleCount} articles</div>
                    </div>
                    <button
                      onClick={(e) => { e.stopPropagation(); removeFile(lf.info.fileName); }}
                      className="text-muted hover:text-danger text-lg leading-none px-1"
                      title="Remove this file"
                    >
                      &times;
                    </button>
                  </div>
                </div>
              ))}
              <div className="flex justify-end">
                <button onClick={clearAll} className="text-muted text-xs hover:text-danger underline">
                  Clear all files
                </button>
              </div>
            </div>
          )}
        </Section>

        {/* ── Section 2: Data Preview ── */}
        {hasFiles && (
          <Section title="2 — Data Preview (Combined)">
            <div className="grid grid-cols-2 sm:grid-cols-5 gap-4 mb-6">
              <Badge label="Total Rows" value={totalRows.toLocaleString()} />
              <Badge label="Unique Stores" value={uniqueStores} />
              <Badge label="Unique Vendors" value={uniqueVendors} />
              <Badge label="Unique Products" value={uniqueProducts} />
              <Badge label="Report Date" value={reportDate || '—'} />
            </div>

            {vendorBreakdown.length > 0 && (
              <div className="overflow-x-auto">
                <table className="w-full text-sm border-collapse">
                  <thead>
                    <tr className="border-b border-border">
                      <th className="text-left py-2 px-3 text-muted font-medium">Vendor Name</th>
                      <th className="text-left py-2 px-3 text-muted font-medium">Vendor Number</th>
                      <th className="text-right py-2 px-3 text-muted font-medium">Products</th>
                      <th className="text-right py-2 px-3 text-muted font-medium">Rows</th>
                    </tr>
                  </thead>
                  <tbody>
                    {vendorBreakdown.map((v, i) => (
                      <tr key={i} className="border-b border-border/50">
                        <td className="py-2 px-3 text-foreground font-medium">{v.vendorName}</td>
                        <td className="py-2 px-3 text-muted">{v.vendorNumber}</td>
                        <td className="py-2 px-3 text-right text-foreground">{v.products}</td>
                        <td className="py-2 px-3 text-right text-foreground">{v.rows.toLocaleString()}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </Section>
        )}

        {/* ── Section 3: Settings ── */}
        {hasFiles && (
          <Section title="3 — Settings">
            <div className="space-y-5">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="text-foreground text-sm font-medium block mb-1">
                    Phantom: Weeks since last received
                  </label>
                  <input
                    type="number"
                    min={1}
                    max={52}
                    value={phantomWeeksReceived}
                    onChange={(e) => setPhantomWeeksReceived(Number(e.target.value) || 4)}
                    className="w-full bg-background border border-border rounded-lg px-3 py-2 text-foreground text-sm"
                  />
                </div>
                <div>
                  <label className="text-foreground text-sm font-medium block mb-1">
                    Phantom: Weeks since last sold
                  </label>
                  <input
                    type="number"
                    min={1}
                    max={52}
                    value={phantomWeeksSold}
                    onChange={(e) => setPhantomWeeksSold(Number(e.target.value) || 4)}
                    className="w-full bg-background border border-border rounded-lg px-3 py-2 text-foreground text-sm"
                  />
                </div>
              </div>

              <div>
                <label className="text-foreground text-sm font-medium block mb-1">Action Mode</label>
                <select
                  value={actionMode}
                  onChange={(e) => setActionMode(e.target.value as 'email' | 'sharepoint' | 'both')}
                  className="bg-background border border-border rounded-lg px-3 py-2 text-foreground text-sm"
                >
                  <option value="both">Email &amp; SharePoint</option>
                  <option value="sharepoint">SharePoint Only</option>
                  <option value="email">Email Only</option>
                </select>
              </div>

              <div>
                <label className="text-foreground text-sm font-medium block mb-1">
                  Report Date (auto-filled from filenames)
                </label>
                <input
                  type="date"
                  value={reportDateOverride}
                  onChange={(e) => setReportDateOverride(e.target.value)}
                  className="bg-background border border-border rounded-lg px-3 py-2 text-foreground text-sm"
                />
              </div>
            </div>
          </Section>
        )}

        {/* ── Section 4: Control File Preview ── */}
        {hasFiles && stage !== 'done' && (
          <Section title="4 — Control File">
            <p className="text-muted text-xs mb-2">
              The control file is auto-fetched from iRam SharePoint when processing begins.
              It maps Site Codes to rep names/emails for email distribution.
            </p>
            <p className="text-muted text-xs">
              Expected path: <code>{`PNP ACTION STORE REPORTS (MULTI VENDOR)/CONTROL FILES/iRam PNP REP STORE ALLOCATION.xlsx`}</code>
            </p>
            <p className="text-muted text-xs mt-1">
              Columns: <code>Site Code</code>, <code>Site Name</code>, <code>Channel</code>, <code>Rep Email</code>, <code>Rep Name</code>
            </p>
          </Section>
        )}

        {/* ── Section 5: Process & Send ── */}
        {hasFiles && (
          <Section title="5 — Process &amp; Send">
            {stage !== 'done' && (
              <button
                onClick={() => { void handleProcess(); }}
                disabled={stage === 'processing' || allRows.length === 0}
                className="w-full font-bold py-3 px-6 rounded-lg transition-colors text-sm text-white disabled:opacity-50 disabled:cursor-not-allowed"
                style={{ background: stage === 'processing' || allRows.length === 0 ? undefined : '#7CC042' }}
                onMouseEnter={(e) => { if (!e.currentTarget.disabled) e.currentTarget.style.background = '#6aad36'; }}
                onMouseLeave={(e) => { if (!e.currentTarget.disabled) e.currentTarget.style.background = '#7CC042'; }}
              >
                {stage === 'processing'
                  ? 'Processing...'
                  : actionMode === 'sharepoint'
                  ? `Generate & Upload ${uniqueStores} Store Reports to SharePoint`
                  : actionMode === 'email'
                  ? `Generate & Email ${uniqueStores} Store Reports`
                  : `Generate, Upload & Email ${uniqueStores} Store Reports`}
              </button>
            )}

            {stage === 'processing' && (
              <div className="mt-4 text-center">
                <p className="text-muted text-sm animate-pulse">
                  Building per-store XLSX reports, uploading to SharePoint and sending emails...
                </p>
                <p className="text-muted text-xs mt-1">This may take a few minutes for large batches.</p>
              </div>
            )}

            {errorMsg && (
              <div className="mt-4 bg-warning/10 border border-warning/30 text-warning rounded-lg px-4 py-3 text-sm whitespace-pre-line">
                {errorMsg}
              </div>
            )}

            {stage === 'done' && processSummary && (
              <div className="mt-4 space-y-4">
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                  <Badge label="Stores Processed" value={processSummary.storesProcessed} color="success" />
                  <Badge label="Reports Generated" value={processSummary.reportsGenerated} color="success" />
                  <Badge label="Emails Sent" value={processSummary.emailsSent} color={processSummary.emailsSent > 0 ? 'success' : 'warning'} />
                  <Badge label="SP Uploaded" value={processSummary.spUploaded} color={processSummary.spUploaded > 0 ? 'success' : 'warning'} />
                </div>

                {processSummary.errors.length > 0 && (
                  <div className="bg-warning/10 border border-warning/30 rounded-lg px-4 py-3">
                    <p className="text-warning text-sm font-medium mb-2">Errors ({processSummary.errors.length}):</p>
                    <ul className="text-warning/80 text-xs space-y-1 max-h-40 overflow-y-auto">
                      {processSummary.errors.map((err: string, i: number) => (
                        <li key={i}>&bull; {err}</li>
                      ))}
                    </ul>
                  </div>
                )}

                {processSummary.storeResults.length > 0 && (
                  <div className="overflow-x-auto max-h-96 overflow-y-auto">
                    <table className="w-full text-xs border-collapse">
                      <thead className="sticky top-0 bg-card">
                        <tr className="border-b border-border">
                          <th className="text-left py-2 px-2 text-muted font-medium">Site Code</th>
                          <th className="text-left py-2 px-2 text-muted font-medium">Store Name</th>
                          <th className="text-right py-2 px-2 text-muted font-medium">OOS</th>
                          <th className="text-right py-2 px-2 text-muted font-medium">Phantom</th>
                          <th className="text-right py-2 px-2 text-muted font-medium">Missing</th>
                          <th className="text-left py-2 px-2 text-muted font-medium">Rep</th>
                          <th className="text-center py-2 px-2 text-muted font-medium">Email</th>
                          <th className="text-center py-2 px-2 text-muted font-medium">SP</th>
                        </tr>
                      </thead>
                      <tbody>
                        {processSummary.storeResults.map((sr: StoreResult, i: number) => (
                          <tr key={i} className="border-b border-border/50">
                            <td className="py-1.5 px-2 text-muted font-mono">{sr.siteCode}</td>
                            <td className="py-1.5 px-2 text-foreground">{sr.storeName}</td>
                            <td className="py-1.5 px-2 text-right text-foreground">{sr.oosCount}</td>
                            <td className="py-1.5 px-2 text-right text-foreground">{sr.phantomCount}</td>
                            <td className="py-1.5 px-2 text-right text-foreground">{sr.missingCount}</td>
                            <td className="py-1.5 px-2 text-muted">{sr.repEmail ?? '—'}</td>
                            <td className="py-1.5 px-2 text-center">{sr.emailed ? '✓' : '—'}</td>
                            <td className="py-1.5 px-2 text-center">{sr.uploaded ? '✓' : '—'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}

                <button
                  onClick={clearAll}
                  className="text-muted text-sm hover:text-accent underline"
                >
                  Start a new batch
                </button>
              </div>
            )}
          </Section>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-border px-6 py-4 flex justify-end items-center gap-3">
        <span className="text-muted text-xs">Powered by</span>
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img src="/oj-logo.png" alt="OuterJoin" className="h-5 w-auto object-contain opacity-75" />
      </footer>
    </div>
  );
}

import { list, put, getDownloadUrl } from '@vercel/blob';

export interface AuditEntry {
  id: string;
  userId: string;
  userName: string;
  userEmail: string;
  timestamp: string; // ISO
  filesUploaded: { fileName: string; vendorName: string; rowCount: number }[];
  totalRows: number;
  reportDate: string;
  actionMode: 'email' | 'sharepoint' | 'both';
  phantomWeeksReceived: number;
  phantomWeeksSold: number;
  storesProcessed: number;
  reportsGenerated: number;
  emailsSent: number;
  spUploaded: number;
  errors: string[];
  durationMs: number;
}

const BLOB_KEY = 'audit-log.json';

export async function loadAuditLog(): Promise<AuditEntry[]> {
  try {
    const { blobs } = await list({ prefix: BLOB_KEY });
    const match = blobs.find(b => b.pathname === BLOB_KEY);
    if (!match) return [];
    const downloadUrl = await getDownloadUrl(match.url);
    const res = await fetch(downloadUrl);
    const text = await res.text();
    return JSON.parse(text) as AuditEntry[];
  } catch {
    return [];
  }
}

export async function appendAuditEntry(entry: AuditEntry): Promise<void> {
  const log = await loadAuditLog();
  log.push(entry);
  await put(BLOB_KEY, JSON.stringify(log, null, 2), {
    access: 'public',
    addRandomSuffix: false,
    contentType: 'application/json',
  });
}

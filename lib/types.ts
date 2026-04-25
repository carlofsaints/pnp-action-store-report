/** A single row from a parsed PnP SDC vendor file */
export interface RawRow {
  weekEndingDate: string;       // A — Week Ending Date
  siteCode: string;             // B — Site Code
  siteDescription: string;      // C — Site Description
  siteProfile: string;          // D — Site Profile
  articleNumber: string;        // E — Article Number
  articleDescription: string;   // F — Article Description
  vendorNumber: string;         // G — Vendor Number
  siteArticleStatus: string;    // H — Site Article Status
  listingStatus: string;        // I — Listing Status
  rpType: string;               // J — RP (MRP) Type
  sohQty: number;               // K — SOH Qty
  drosQty: number;              // L — DROS Qty
  daysCover: number;            // M — Days Cover
  sourceOfSupply: string;       // N — Source Of Supply
  dateLastReceived: string;     // O — Date Last Received (YYYY-MM-DD)
  dateLastSold: string;         // P — Date Last Sold (YYYY-MM-DD)
  lastOrderedDate: string;      // Q — Last Ordered Date (YYYY-MM-DD)
  vendorName: string;           // Extracted from filename
}

/** Metadata about a parsed vendor file */
export interface FileInfo {
  fileName: string;
  vendorName: string;
  vendorNumber: string;
  rowCount: number;
  storeCount: number;
  articleCount: number;
  reportDate: string;           // YYYY-MM-DD from filename
  warning?: string;             // e.g. "Empty file — no data rows"
}

/** Response from POST /api/parse */
export interface ParseResponse {
  files: FileInfo[];
  allRows: RawRow[];
  reportDate: string;
  totalRows: number;
  uniqueStores: number;
  uniqueVendors: number;
  uniqueProducts: number;
}

/** A rep entry from the control file */
export interface ControlEntry {
  siteCode: string;
  siteName: string;
  channel: string;
  repName: string;
  repEmail: string;
}

/** Request body for POST /api/process */
export interface ProcessRequest {
  allRows: RawRow[];
  reportDate: string;
  phantomWeeksReceived: number;
  phantomWeeksSold: number;
  actionMode: 'email' | 'sharepoint' | 'both';
}

/** Per-store result from processing */
export interface StoreResult {
  siteCode: string;
  storeName: string;
  oosCount: number;
  phantomCount: number;
  missingCount: number;
  repEmail: string | null;
  emailed: boolean;
  uploaded: boolean;
  error: string | null;
}

/** Summary from POST /api/process */
export interface ProcessSummary {
  storesProcessed: number;
  reportsGenerated: number;
  emailsSent: number;
  spUploaded: number;
  errors: string[];
  storeResults: StoreResult[];
}

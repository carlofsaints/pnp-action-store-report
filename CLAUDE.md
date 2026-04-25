# PnP Action Store Report

## Overview
Consolidates PnP Portal SDC export files from multiple vendors into per-store Excel reports with OOS analysis, phantom stock analysis, and missing SKU analysis. Reports are uploaded to iRam SharePoint and/or emailed to reps.

## Tech Stack
- Next.js 16 (App Router, TypeScript, Tailwind v4)
- xlsx (SheetJS) for input parsing
- exceljs for output report generation
- Microsoft Graph API (iRam tenant) for SharePoint
- Resend for email
- bcryptjs for password hashing
- @vercel/blob (v0.27) for user + audit log storage

## Key Patterns
- iRam green: #7CC042, dark charcoal: #32373C
- Dates from PnP files are Excel serial numbers — convert with `excelSerialToDate()`
- Vendor name extracted from filename: `VENDOR_NAME SDC YYYY-MM-DD.xlsx`
- graph-iram.ts reused from Phantom Consolidator pattern
- All API GET routes returning mutable data need `Cache-Control: no-store`

## Auth
- bcryptjs + Blob pattern (same as iRam RVL CRM)
- Session: `pnp_session` in localStorage
- `authFetch()` attaches `x-user-id` header
- Seed: POST `/api/seed` `{"secret":"pnp-seed-2026"}` → Carl + Johann admin, pw `pnp2026`
- Admin panel: `/admin` (users + audit log, admin role only)

## Env Vars
IRAM_TENANT_ID, IRAM_CLIENT_ID, IRAM_CLIENT_SECRET, IRAM_SP_HOST, IRAM_SP_LIBRARY, PNP_BASE_FOLDER, RESEND_API_KEY, NEXT_PUBLIC_SITE_URL, BLOB_READ_WRITE_TOKEN, SEED_SECRET

/**
 * iRAM SharePoint Graph client
 * Used for: uploading per-store XLSX reports and reading the control file.
 * Adapted from Phantom Consolidator.
 */

const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));

const TENANT_ID = process.env.IRAM_TENANT_ID!;
const CLIENT_ID = process.env.IRAM_CLIENT_ID!;
const CLIENT_SECRET = process.env.IRAM_CLIENT_SECRET!;
const SP_HOST = process.env.IRAM_SP_HOST ?? 'iramsa.sharepoint.com';
const LIBRARY_NAME = process.env.IRAM_SP_LIBRARY ?? 'In-store';
const BASE_FOLDER = process.env.PNP_BASE_FOLDER ?? 'PNP ACTION STORE REPORTS (MULTI VENDOR)';

// ── Auth ─────────────────────────────────────────────────────────────────────

async function getToken(): Promise<string> {
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
      }),
    }
  );
  const data = await res.json();
  if (!data.access_token) {
    throw new Error(`iRAM auth failed: ${data.error_description ?? JSON.stringify(data)}`);
  }
  return data.access_token as string;
}

function encodePath(path: string): string {
  return path.split('/').map((seg) => encodeURIComponent(seg)).join('/');
}

export type DriveContext = { token: string; driveId: string };

export async function getDriveContext(): Promise<DriveContext> {
  const token = await getToken();

  const siteRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SP_HOST}:/`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!siteRes.ok) throw new Error(`iRAM: could not get site: ${await siteRes.text()}`);
  const site = await siteRes.json();

  const drivesRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${site.id}/drives`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const drives = await drivesRes.json();
  const drive = drives.value?.find((d: { name: string }) => d.name.toLowerCase() === LIBRARY_NAME.toLowerCase());
  if (!drive) {
    const names = drives.value?.map((d: { name: string }) => d.name).join(', ');
    throw new Error(`iRAM: library "${LIBRARY_NAME}" not found. Available: ${names}`);
  }
  return { token, driveId: drive.id as string };
}

// ── Ensure folder path exists ────────────────────────────────────────────────

async function ensureFolderExists(
  token: string,
  driveId: string,
  folderPath: string
): Promise<void> {
  const segments = folderPath.split('/');
  let currentPath = '';

  for (const segment of segments) {
    const parentPath = currentPath ? encodePath(currentPath) : undefined;

    const url = parentPath
      ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${parentPath}:/children`
      : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`;

    await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        name: segment,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'replace',
      }),
    });
    // 409 = already exists, which is fine

    currentPath = currentPath ? `${currentPath}/${segment}` : segment;
  }
}

// ── Upload file ──────────────────────────────────────────────────────────────

export interface UploadResult {
  webUrl: string;
  fileId: string;
}

export async function uploadReport(
  buffer: Buffer,
  reportDate: string,
  fileName: string,
  ctx?: DriveContext
): Promise<UploadResult> {
  const { token, driveId } = ctx ?? await getDriveContext();

  const folderPath = `${BASE_FOLDER}/${reportDate}`;
  await ensureFolderExists(token, driveId, folderPath);

  const filePath = encodePath(`${folderPath}/${fileName}`);

  for (let attempt = 0; attempt < 4; attempt++) {
    const uploadRes = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
      {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        },
        body: new Uint8Array(buffer),
      }
    );

    if (uploadRes.status === 429) {
      if (attempt === 3) {
        throw new Error(`iRAM: upload failed (429): throttled after 3 retries`);
      }
      const retryAfterSec = parseInt(uploadRes.headers.get('Retry-After') ?? '15', 10);
      await sleep(Math.min(retryAfterSec, 30) * 1000);
      continue;
    }

    if (!uploadRes.ok) {
      throw new Error(`iRAM: upload failed (${uploadRes.status}): ${await uploadRes.text()}`);
    }

    const uploaded = await uploadRes.json();
    return {
      webUrl: uploaded.webUrl as string,
      fileId: uploaded.id as string,
    };
  }

  throw new Error('iRAM: upload failed — max retries exceeded');
}

// ── Download file from SP (for control file) ────────────────────────────────

export async function downloadFile(
  filePath: string,
  ctx?: DriveContext
): Promise<Buffer> {
  const { token, driveId } = ctx ?? await getDriveContext();

  const encoded = encodePath(filePath);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encoded}:/content`,
    { headers: { Authorization: `Bearer ${token}` }, redirect: 'follow' }
  );

  if (!res.ok) {
    throw new Error(`iRAM: download failed (${res.status}): ${await res.text()}`);
  }

  return Buffer.from(await res.arrayBuffer());
}

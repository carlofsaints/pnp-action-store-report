import { NextResponse } from 'next/server';
import { loadAuditLog } from '@/lib/auditData';
import { findUserById } from '@/lib/userData';

export const dynamic = 'force-dynamic';

export async function GET(req: Request) {
  try {
    const userId = req.headers.get('x-user-id');
    if (!userId) {
      return NextResponse.json({ error: 'Not authenticated' }, { status: 401 });
    }

    const user = await findUserById(userId);
    if (!user || user.role !== 'admin') {
      return NextResponse.json({ error: 'Admin access required' }, { status: 403 });
    }

    const log = await loadAuditLog();
    // Return most recent first
    log.reverse();

    return NextResponse.json(log, {
      headers: { 'Cache-Control': 'no-store' },
    });
  } catch (e) {
    return NextResponse.json(
      { error: e instanceof Error ? e.message : 'Failed to load audit log' },
      { status: 500 },
    );
  }
}

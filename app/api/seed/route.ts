import { NextResponse } from 'next/server';
import bcrypt from 'bcryptjs';
import { loadUsers, saveUsers, type User } from '@/lib/userData';

export const dynamic = 'force-dynamic';

const SEED_SECRET = process.env.SEED_SECRET || 'pnp-seed-2026';

export async function POST(req: Request) {
  try {
    const { secret } = await req.json();
    if (secret !== SEED_SECRET) {
      return NextResponse.json({ error: 'Invalid secret' }, { status: 403 });
    }

    const existing = await loadUsers();

    const seedUsers: Omit<User, 'password'>[] = [
      { id: 'carl', name: 'Carl dos Santos', email: 'carl@outerjoin.co.za', role: 'admin', createdAt: new Date().toISOString() },
      { id: 'johann', name: 'Johann Venter', email: 'johann@iram.co.za', role: 'admin', createdAt: new Date().toISOString() },
    ];

    const defaultPassword = await bcrypt.hash('pnp2026', 10);
    let added = 0;

    for (const su of seedUsers) {
      const exists = existing.some(u => u.email.toLowerCase() === su.email.toLowerCase());
      if (!exists) {
        existing.push({ ...su, password: defaultPassword });
        added++;
      }
    }

    await saveUsers(existing);

    return NextResponse.json({ ok: true, added, total: existing.length });
  } catch (e) {
    return NextResponse.json(
      { error: e instanceof Error ? e.message : 'Seed failed' },
      { status: 500 },
    );
  }
}

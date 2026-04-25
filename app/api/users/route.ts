import { NextResponse } from 'next/server';
import bcrypt from 'bcryptjs';
import { loadUsers, saveUsers, findUserById } from '@/lib/userData';
import type { User } from '@/lib/userData';

export const dynamic = 'force-dynamic';

async function requireAdmin(req: Request): Promise<User | NextResponse> {
  const userId = req.headers.get('x-user-id');
  if (!userId) return NextResponse.json({ error: 'Not authenticated' }, { status: 401 });
  const user = await findUserById(userId);
  if (!user || user.role !== 'admin') return NextResponse.json({ error: 'Admin access required' }, { status: 403 });
  return user;
}

export async function GET(req: Request) {
  const check = await requireAdmin(req);
  if (check instanceof NextResponse) return check;

  const users = await loadUsers();
  // Strip passwords before returning
  const safe = users.map(({ password: _, ...u }) => u);

  return NextResponse.json(safe, {
    headers: { 'Cache-Control': 'no-store' },
  });
}

export async function POST(req: Request) {
  const check = await requireAdmin(req);
  if (check instanceof NextResponse) return check;

  try {
    const { name, email, password, role } = await req.json();
    if (!name || !email || !password) {
      return NextResponse.json({ error: 'Name, email and password required' }, { status: 400 });
    }

    const users = await loadUsers();
    const exists = users.some(u => u.email.toLowerCase() === email.toLowerCase());
    if (exists) {
      return NextResponse.json({ error: 'Email already exists' }, { status: 409 });
    }

    const hash = await bcrypt.hash(password, 10);
    const newUser: User = {
      id: email.split('@')[0].toLowerCase() + '-' + Date.now().toString(36),
      name,
      email,
      password: hash,
      role: role === 'admin' ? 'admin' : 'user',
      createdAt: new Date().toISOString(),
    };

    users.push(newUser);
    await saveUsers(users);

    const { password: _, ...safe } = newUser;
    return NextResponse.json(safe, { status: 201 });
  } catch (e) {
    return NextResponse.json(
      { error: e instanceof Error ? e.message : 'Failed to create user' },
      { status: 500 },
    );
  }
}

export async function DELETE(req: Request) {
  const check = await requireAdmin(req);
  if (check instanceof NextResponse) return check;

  try {
    const { id } = await req.json();
    if (!id) return NextResponse.json({ error: 'User id required' }, { status: 400 });

    const users = await loadUsers();
    const filtered = users.filter(u => u.id !== id);
    if (filtered.length === users.length) {
      return NextResponse.json({ error: 'User not found' }, { status: 404 });
    }

    await saveUsers(filtered);
    return NextResponse.json({ ok: true });
  } catch (e) {
    return NextResponse.json(
      { error: e instanceof Error ? e.message : 'Failed to delete user' },
      { status: 500 },
    );
  }
}

import { list, put, getDownloadUrl } from '@vercel/blob';

export interface User {
  id: string;
  name: string;
  email: string;
  password: string; // bcrypt hash
  role: 'admin' | 'user';
  createdAt: string; // ISO
}

const BLOB_KEY = 'users.json';

export async function loadUsers(): Promise<User[]> {
  try {
    const { blobs } = await list({ prefix: BLOB_KEY });
    const match = blobs.find(b => b.pathname === BLOB_KEY);
    if (!match) return [];
    const downloadUrl = await getDownloadUrl(match.url);
    const res = await fetch(downloadUrl);
    const text = await res.text();
    return JSON.parse(text) as User[];
  } catch {
    return [];
  }
}

export async function saveUsers(users: User[]): Promise<void> {
  await put(BLOB_KEY, JSON.stringify(users, null, 2), {
    access: 'public',
    addRandomSuffix: false,
    contentType: 'application/json',
  });
}

export async function findUserById(id: string): Promise<User | null> {
  const users = await loadUsers();
  return users.find(u => u.id === id) ?? null;
}

export async function findUserByEmail(email: string): Promise<User | null> {
  const users = await loadUsers();
  return users.find(u => u.email.toLowerCase() === email.toLowerCase()) ?? null;
}

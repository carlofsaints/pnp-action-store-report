'use client';

import { useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';

export interface Session {
  id: string;
  name: string;
  email: string;
  role: 'admin' | 'user';
}

const SESSION_KEY = 'pnp_session';

export function useAuth(): { session: Session | null; loading: boolean; logout: () => void } {
  const [session, setSession] = useState<Session | null>(null);
  const [loading, setLoading] = useState(true);
  const router = useRouter();

  useEffect(() => {
    const raw = localStorage.getItem(SESSION_KEY);
    if (raw) {
      try {
        setSession(JSON.parse(raw) as Session);
      } catch {
        localStorage.removeItem(SESSION_KEY);
        router.push('/login');
      }
    } else {
      router.push('/login');
    }
    setLoading(false);
  }, [router]);

  const logout = () => {
    localStorage.removeItem(SESSION_KEY);
    setSession(null);
    router.push('/login');
  };

  return { session, loading, logout };
}

export function authFetch(url: string, opts: RequestInit = {}): Promise<Response> {
  const raw = typeof window !== 'undefined' ? localStorage.getItem(SESSION_KEY) : null;
  let userId = '';
  if (raw) {
    try { userId = (JSON.parse(raw) as Session).id; } catch { /* skip */ }
  }
  const headers = new Headers(opts.headers);
  if (userId) headers.set('x-user-id', userId);
  return fetch(url, { ...opts, headers });
}

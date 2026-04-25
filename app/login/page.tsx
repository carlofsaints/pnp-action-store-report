'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';

export default function LoginPage() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const router = useRouter();

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    setLoading(true);

    try {
      const res = await fetch('/api/auth', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password }),
      });

      if (!res.ok) {
        const data = await res.json();
        setError(data.error || 'Login failed');
        return;
      }

      const session = await res.json();
      localStorage.setItem('pnp_session', JSON.stringify(session));
      router.push('/');
    } catch {
      setError('Network error — try again');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-background flex items-center justify-center px-4">
      <div className="bg-card border border-border rounded-xl p-8 w-full max-w-sm shadow-sm">
        <div className="flex items-center gap-3 mb-6">
          <div className="w-1.5 h-8 rounded" style={{ background: '#7CC042' }} />
          <div>
            <h1 className="text-xl font-bold text-foreground">PnP Action Store Report</h1>
            <p className="text-muted text-xs">Sign in to continue</p>
          </div>
        </div>

        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="text-foreground text-sm font-medium block mb-1">Email</label>
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              required
              autoFocus
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-foreground text-sm"
              placeholder="you@example.com"
            />
          </div>
          <div>
            <label className="text-foreground text-sm font-medium block mb-1">Password</label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              required
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-foreground text-sm"
            />
          </div>

          {error && (
            <div className="bg-danger/10 border border-danger/30 text-danger rounded-lg px-3 py-2 text-sm">
              {error}
            </div>
          )}

          <button
            type="submit"
            disabled={loading}
            className="w-full font-bold py-2.5 px-4 rounded-lg text-white text-sm transition-colors disabled:opacity-50"
            style={{ background: '#7CC042' }}
            onMouseEnter={(e) => { if (!e.currentTarget.disabled) e.currentTarget.style.background = '#6aad36'; }}
            onMouseLeave={(e) => { if (!e.currentTarget.disabled) e.currentTarget.style.background = '#7CC042'; }}
          >
            {loading ? 'Signing in...' : 'Sign In'}
          </button>
        </form>

        <div className="mt-6 flex justify-center">
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img src="/oj-logo.png" alt="OuterJoin" className="h-5 w-auto object-contain opacity-50" />
        </div>
      </div>
    </div>
  );
}

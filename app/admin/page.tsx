'use client';

import { useState, useEffect } from 'react';
import { useAuth, authFetch } from '@/lib/useAuth';
import type { AuditEntry } from '@/lib/auditData';

interface SafeUser {
  id: string;
  name: string;
  email: string;
  role: 'admin' | 'user';
  createdAt: string;
}

export default function AdminPage() {
  const { session, loading, logout } = useAuth();
  const [users, setUsers] = useState<SafeUser[]>([]);
  const [auditLog, setAuditLog] = useState<AuditEntry[]>([]);
  const [tab, setTab] = useState<'users' | 'audit'>('users');

  // Add user form
  const [newName, setNewName] = useState('');
  const [newEmail, setNewEmail] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [newRole, setNewRole] = useState<'admin' | 'user'>('user');
  const [formError, setFormError] = useState('');
  const [formSuccess, setFormSuccess] = useState('');

  useEffect(() => {
    if (!session || session.role !== 'admin') return;
    void loadData();
  }, [session]);

  const loadData = async () => {
    const [usersRes, auditRes] = await Promise.all([
      authFetch('/api/users'),
      authFetch('/api/audit'),
    ]);
    if (usersRes.ok) setUsers(await usersRes.json());
    if (auditRes.ok) setAuditLog(await auditRes.json());
  };

  const handleAddUser = async (e: React.FormEvent) => {
    e.preventDefault();
    setFormError('');
    setFormSuccess('');

    const res = await authFetch('/api/users', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: newName, email: newEmail, password: newPassword, role: newRole }),
    });

    if (!res.ok) {
      const data = await res.json();
      setFormError(data.error || 'Failed to add user');
      return;
    }

    setFormSuccess(`${newName} added`);
    setNewName('');
    setNewEmail('');
    setNewPassword('');
    setNewRole('user');
    void loadData();
  };

  const handleDeleteUser = async (id: string, name: string) => {
    if (!confirm(`Delete ${name}?`)) return;
    await authFetch('/api/users', {
      method: 'DELETE',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id }),
    });
    void loadData();
  };

  if (loading) {
    return <div className="min-h-screen bg-background flex items-center justify-center text-muted">Loading...</div>;
  }

  if (!session || session.role !== 'admin') {
    return (
      <div className="min-h-screen bg-background flex items-center justify-center">
        <div className="bg-card border border-border rounded-xl p-8 text-center">
          <p className="text-foreground font-medium mb-2">Admin access required</p>
          <a href="/" className="text-accent text-sm underline">Back to dashboard</a>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-background text-foreground">
      <header className="border-b border-border px-6 py-3 flex items-center justify-between sticky top-0 bg-card z-10 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="w-1.5 h-8 rounded" style={{ background: '#7CC042' }} />
          <div>
            <h1 className="text-xl font-bold text-foreground">Admin Panel</h1>
            <p className="text-muted text-xs">Users &amp; Audit Log</p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <a href="/" className="text-accent text-sm hover:underline">Dashboard</a>
          <span className="text-muted text-xs">{session.name}</span>
          <button onClick={logout} className="text-muted text-xs hover:text-danger underline">Logout</button>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-8">
        {/* Tabs */}
        <div className="flex gap-4 mb-6">
          <button
            onClick={() => setTab('users')}
            className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
              tab === 'users' ? 'bg-accent text-white' : 'bg-card border border-border text-muted hover:text-foreground'
            }`}
          >
            Users ({users.length})
          </button>
          <button
            onClick={() => setTab('audit')}
            className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
              tab === 'audit' ? 'bg-accent text-white' : 'bg-card border border-border text-muted hover:text-foreground'
            }`}
          >
            Audit Log ({auditLog.length})
          </button>
        </div>

        {/* Users Tab */}
        {tab === 'users' && (
          <div className="space-y-6">
            {/* User List */}
            <div className="bg-card border border-border rounded-xl p-6 shadow-sm">
              <h2 className="text-lg font-bold mb-4" style={{ color: '#7CC042' }}>Users</h2>
              <div className="overflow-x-auto">
                <table className="w-full text-sm border-collapse">
                  <thead>
                    <tr className="border-b border-border">
                      <th className="text-left py-2 px-3 text-muted font-medium">Name</th>
                      <th className="text-left py-2 px-3 text-muted font-medium">Email</th>
                      <th className="text-left py-2 px-3 text-muted font-medium">Role</th>
                      <th className="text-left py-2 px-3 text-muted font-medium">Created</th>
                      <th className="text-right py-2 px-3 text-muted font-medium"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {users.map(u => (
                      <tr key={u.id} className="border-b border-border/50">
                        <td className="py-2 px-3 text-foreground font-medium">{u.name}</td>
                        <td className="py-2 px-3 text-muted">{u.email}</td>
                        <td className="py-2 px-3">
                          <span className={`text-xs px-2 py-0.5 rounded-full ${
                            u.role === 'admin' ? 'bg-accent/10 text-accent' : 'bg-muted/10 text-muted'
                          }`}>
                            {u.role}
                          </span>
                        </td>
                        <td className="py-2 px-3 text-muted text-xs">{new Date(u.createdAt).toLocaleDateString('en-GB')}</td>
                        <td className="py-2 px-3 text-right">
                          <button
                            onClick={() => handleDeleteUser(u.id, u.name)}
                            className="text-muted hover:text-danger text-xs underline"
                          >
                            Delete
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Add User Form */}
            <div className="bg-card border border-border rounded-xl p-6 shadow-sm">
              <h2 className="text-lg font-bold mb-4" style={{ color: '#7CC042' }}>Add User</h2>
              <form onSubmit={handleAddUser} className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="text-foreground text-sm font-medium block mb-1">Name</label>
                  <input
                    type="text"
                    value={newName}
                    onChange={(e) => setNewName(e.target.value)}
                    required
                    className="w-full bg-background border border-border rounded-lg px-3 py-2 text-foreground text-sm"
                  />
                </div>
                <div>
                  <label className="text-foreground text-sm font-medium block mb-1">Email</label>
                  <input
                    type="email"
                    value={newEmail}
                    onChange={(e) => setNewEmail(e.target.value)}
                    required
                    className="w-full bg-background border border-border rounded-lg px-3 py-2 text-foreground text-sm"
                  />
                </div>
                <div>
                  <label className="text-foreground text-sm font-medium block mb-1">Password</label>
                  <input
                    type="password"
                    value={newPassword}
                    onChange={(e) => setNewPassword(e.target.value)}
                    required
                    className="w-full bg-background border border-border rounded-lg px-3 py-2 text-foreground text-sm"
                  />
                </div>
                <div>
                  <label className="text-foreground text-sm font-medium block mb-1">Role</label>
                  <select
                    value={newRole}
                    onChange={(e) => setNewRole(e.target.value as 'admin' | 'user')}
                    className="w-full bg-background border border-border rounded-lg px-3 py-2 text-foreground text-sm"
                  >
                    <option value="user">User</option>
                    <option value="admin">Admin</option>
                  </select>
                </div>
                <div className="sm:col-span-2 flex items-center gap-4">
                  <button
                    type="submit"
                    className="font-bold py-2 px-6 rounded-lg text-white text-sm"
                    style={{ background: '#7CC042' }}
                  >
                    Add User
                  </button>
                  {formError && <span className="text-danger text-sm">{formError}</span>}
                  {formSuccess && <span className="text-success text-sm">{formSuccess}</span>}
                </div>
              </form>
            </div>
          </div>
        )}

        {/* Audit Tab */}
        {tab === 'audit' && (
          <div className="bg-card border border-border rounded-xl p-6 shadow-sm">
            <h2 className="text-lg font-bold mb-4" style={{ color: '#7CC042' }}>Audit Log</h2>
            {auditLog.length === 0 ? (
              <p className="text-muted text-sm">No processing runs recorded yet.</p>
            ) : (
              <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
                <table className="w-full text-xs border-collapse">
                  <thead className="sticky top-0 bg-card">
                    <tr className="border-b border-border">
                      <th className="text-left py-2 px-2 text-muted font-medium">Date/Time</th>
                      <th className="text-left py-2 px-2 text-muted font-medium">User</th>
                      <th className="text-left py-2 px-2 text-muted font-medium">Files</th>
                      <th className="text-right py-2 px-2 text-muted font-medium">Rows</th>
                      <th className="text-right py-2 px-2 text-muted font-medium">Stores</th>
                      <th className="text-right py-2 px-2 text-muted font-medium">Reports</th>
                      <th className="text-right py-2 px-2 text-muted font-medium">Emails</th>
                      <th className="text-right py-2 px-2 text-muted font-medium">SP</th>
                      <th className="text-left py-2 px-2 text-muted font-medium">Mode</th>
                      <th className="text-right py-2 px-2 text-muted font-medium">Duration</th>
                      <th className="text-right py-2 px-2 text-muted font-medium">Errors</th>
                    </tr>
                  </thead>
                  <tbody>
                    {auditLog.map(entry => (
                      <tr key={entry.id} className="border-b border-border/50">
                        <td className="py-1.5 px-2 text-muted whitespace-nowrap">
                          {new Date(entry.timestamp).toLocaleString('en-GB', { day: '2-digit', month: '2-digit', year: '2-digit', hour: '2-digit', minute: '2-digit' })}
                        </td>
                        <td className="py-1.5 px-2 text-foreground">{entry.userName}</td>
                        <td className="py-1.5 px-2 text-muted">{entry.filesUploaded.length}</td>
                        <td className="py-1.5 px-2 text-right text-foreground">{entry.totalRows.toLocaleString()}</td>
                        <td className="py-1.5 px-2 text-right text-foreground">{entry.storesProcessed}</td>
                        <td className="py-1.5 px-2 text-right text-foreground">{entry.reportsGenerated}</td>
                        <td className="py-1.5 px-2 text-right text-foreground">{entry.emailsSent}</td>
                        <td className="py-1.5 px-2 text-right text-foreground">{entry.spUploaded}</td>
                        <td className="py-1.5 px-2 text-muted">{entry.actionMode}</td>
                        <td className="py-1.5 px-2 text-right text-muted">{(entry.durationMs / 1000).toFixed(1)}s</td>
                        <td className="py-1.5 px-2 text-right">
                          {entry.errors.length > 0 ? (
                            <span className="text-danger" title={entry.errors.join('\n')}>{entry.errors.length}</span>
                          ) : (
                            <span className="text-success">0</span>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}

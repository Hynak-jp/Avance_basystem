'use client';

import { useState } from 'react';
import { signIn } from 'next-auth/react';
import { useRouter } from 'next/navigation';

export default function StaffLoginClient() {
  const router = useRouter();
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [pending, setPending] = useState(false);
  const [error, setError] = useState('');

  const onSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (pending) return;
    setPending(true);
    setError('');
    try {
      const normalizedEmail = email.trim().toLowerCase();
      const res = await signIn('credentials', {
        redirect: false,
        email: normalizedEmail,
        password,
        callbackUrl: '/form',
      });
      if (res?.ok) {
        router.replace(res.url || '/form');
        return;
      }
      setError('ログインに失敗しました。');
    } catch (err) {
      void err;
      setError('ログインに失敗しました。');
    } finally {
      setPending(false);
    }
  };

  return (
    <div className="max-w-sm mx-auto mt-10">
      <h1 className="text-2xl font-bold mb-4">スタッフログイン</h1>
      <form onSubmit={onSubmit} className="space-y-4">
        <div>
          <label className="block text-sm mb-1">メールアドレス</label>
          <input
            type="email"
            className="w-full border rounded px-3 py-2"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            autoComplete="username"
            required
          />
        </div>
        <div>
          <label className="block text-sm mb-1">パスワード</label>
          <input
            type="password"
            className="w-full border rounded px-3 py-2"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            autoComplete="current-password"
            required
          />
        </div>
        {error ? <p className="text-sm text-red-600">{error}</p> : null}
        <button
          type="submit"
          className="w-full bg-black text-white rounded px-3 py-2 disabled:opacity-60"
          disabled={pending}
        >
          {pending ? 'ログイン中…' : 'ログイン'}
        </button>
      </form>
    </div>
  );
}

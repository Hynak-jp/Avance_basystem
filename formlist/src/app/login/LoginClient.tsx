'use client';

import { useState, useTransition } from 'react';
import { signIn } from 'next-auth/react';

export default function LoginClient() {
  const [agreed, setAgreed] = useState(false);
  const [pending, startTransition] = useTransition();

  const handleLogin = () => {
    startTransition(() => {
      signIn('line', { callbackUrl: '/form' });
    });
  };

  return (
    <main className="container mx-auto p-6 max-w-xl">
      <h1 className="text-2xl font-bold mb-3">LINEログイン & 利用同意</h1>
      <p className="text-sm text-gray-700 mb-4">
        このフォームは破産申立のための必要情報を収集するものです。あなたのLINEアカウントと連携することで、
        各フォームへの入力情報と紐づけて管理します。
      </p>

      <label className="flex items-start gap-2 mb-6">
        <input
          type="checkbox"
          className="mt-1"
          checked={agreed}
          onChange={(e) => setAgreed(e.target.checked)}
        />
        <span>利用目的を理解し、同意します（チェックでボタンがアクティブになる）</span>
      </label>

      <button
        onClick={handleLogin}
        disabled={!agreed || pending}
        className={`px-4 py-2 rounded ${
          (!agreed || pending)
            ? 'bg-gray-300 text-gray-600 cursor-not-allowed'
            : 'bg-green-600 text-white hover:bg-green-700'
        }`}
      >
        {pending ? 'リダイレクト中…' : 'LINEでログイン'}
      </button>
    </main>
  );
}

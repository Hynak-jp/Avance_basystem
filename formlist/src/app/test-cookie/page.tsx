// app/test-cookie/page.tsx
'use client';

import { useState } from 'react';

export default function TestCookiePage() {
  const [value, setValue] = useState<string | null>(null);

  const setCookie = async () => {
    const res = await fetch('/api/set-cookie');
    const data = await res.json();
    alert(data.message);
  };

  const readCookie = async () => {
    const res = await fetch('/api/read-cookie');
    const data = await res.json();
    setValue(data.value);
  };

  return (
    <div className="p-6">
      <h1 className="text-xl font-bold mb-4">🍪 Cookie テストページ</h1>

      <button onClick={setCookie} className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded mr-4">
        Cookieをセット
      </button>

      <button onClick={readCookie} className="bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded">
        Cookieを読む
      </button>

      {value && (
        <p className="mt-6">
          Cookieの値：<strong>{value}</strong>
        </p>
      )}
    </div>
  );
}

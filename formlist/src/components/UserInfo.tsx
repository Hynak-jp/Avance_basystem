'use client';

import Image from 'next/image';
import { useSession, signOut } from 'next-auth/react';
import { useMemo, useState } from 'react';

export default function UserInfo() {
  const { data: session, status } = useSession();
  const [imgErr, setImgErr] = useState(false);

  const raw = session?.user?.image ? String(session.user.image) : '';
  const proxied = useMemo(() => (raw ? `/api/avatar?url=${encodeURIComponent(raw)}` : ''), [raw]);

  const showAvatar = status === 'authenticated' && !!proxied && !imgErr;

  return (
    <div className="flex items-center justify-between p-4 mb-6 rounded bg-gray-50">
      <div className="flex items-center gap-3">
        {showAvatar ? (
          <Image
            key={proxied}
            src={proxied}
            alt="avatar"
            width={40}
            height={40}
            className="rounded-full"
            onError={() => setImgErr(true)}
          />
        ) : (
          // セッション確定前やエラー時は簡易プレースホルダ（※フォールバック画像は出さない）
          <div className="w-10 h-10 rounded-full bg-gray-200" />
        )}
        <div>
          <p className="font-bold">{session?.user?.name ?? 'ログイン中'} さんでログイン中</p>
          {session?.lineId && <p className="text-sm text-gray-600">LINE ID: {session.lineId}</p>}
        </div>
      </div>

      {status === 'authenticated' && (
        <button
          onClick={() => signOut({ callbackUrl: '/login' })}
          className="px-4 py-2 bg-red-500 text-white rounded hover:bg-red-600"
        >
          ログアウト
        </button>
      )}
    </div>
  );
}

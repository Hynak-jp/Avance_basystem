'use client';

import { useEffect, useState } from 'react';

type Info = { version: string; commit: string; branch?: string; env: string };

export default function VersionBadge() {
  const [info, setInfo] = useState<Info | null>(null);
  const [updateAvailable, setUpdateAvailable] = useState(false);

  useEffect(() => {
    const key = 'formlist_commit';

    fetch('/api/version', { cache: 'no-store' })
      .then((response) => response.json())
      .then((payload: Info) => {
        setInfo(payload);
        const previous = localStorage.getItem(key);

        if (previous && previous !== payload.commit) {
          setUpdateAvailable(true);
        }

        localStorage.setItem(key, payload.commit);
      })
      .catch(() => {
        // 通信エラー時はバッジを非表示のままにする
      });
  }, []);

  if (!info) {
    return null;
  }

  return (
    <div className="fixed right-2 bottom-2 z-50 text-[11px] sm:text-xs">
      <div className="rounded bg-black/70 text-white px-2 py-1 shadow hover:bg-black/80">
        v{info.version} · {info.commit} · {info.env}
        {updateAvailable && (
          <button
            className="ml-2 underline"
            onClick={() => location.reload()}
            title="新しいバージョンがあります。再読み込み"
          >
            更新あり→再読み込み
          </button>
        )}
      </div>
    </div>
  );
}

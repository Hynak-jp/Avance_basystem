'use client';

import { useEffect } from 'react';
import { useRouter, useSearchParams } from 'next/navigation';
import { makeProgressStore } from '@/lib/progressStore';

export default function DoneClient({ lineId }: { lineId: string }) {
  const search = useSearchParams();
  const router = useRouter();
  const formId = search.get('formId') || 'unknown';
  const formKeyParam = search.get('formKey') || '';
  const form = search.get('form') || '';
  const storeKey = formKeyParam || formId || 'unknown';

  const store = makeProgressStore(lineId)();

  useEffect(() => {
    (async () => {
      const intakeFormId = process.env.NEXT_PUBLIC_INTAKE_FORM_ID;
      const isIntake = form === 'intake' || (intakeFormId && formId === intakeFormId);

      if (isIntake && lineId) {
        try {
          const r = await fetch('/api/intake/complete', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ lineId }),
            cache: 'no-store',
          });
          let data: { ok?: boolean } | null = null;
          try { data = (await r.json()) as { ok?: boolean }; } catch {}
          console.log('intake_complete:', r.status, data);
          if (r.ok && (data?.ok ?? true)) {
            store.setStatus(storeKey, 'done');
            router.replace('/form');
            return;
          }
        } catch (e) {
          console.error('intake_complete error', e);
        }
      }
      // フォールバック：受付以外 or 失敗時も一覧へ戻す
      store.setStatus(storeKey, 'done');
      router.replace('/form');
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [form, formId, formKeyParam, lineId, storeKey]);

  return <p>送信ありがとうございました。処理が完了すると一覧に戻ります…</p>;
}

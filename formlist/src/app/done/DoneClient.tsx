'use client';

import { useEffect } from 'react';
import { useRouter, useSearchParams } from 'next/navigation';
import { makeProgressStore } from '@/lib/progressStore';

export default function DoneClient({ lineId }: { lineId: string }) {
  const search = useSearchParams();
  const router = useRouter();
  const formId = search.get('formId') || 'unknown';
  const formKeyParam = search.get('formKey') || '';
  const storeKeyParam = search.get('storeKey') || '';
  const form = search.get('form') || '';
  const storeKey = storeKeyParam || formKeyParam || formId || 'unknown';
  const pParam = search.get('p') || '';
  const tsParam = search.get('ts') || '';
  const sigParam = search.get('sig') || '';
  const caseIdParam = search.get('caseId') || '';

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
            // 追加: 保存直後に収集を明示的に起動（V2 署名は /api/status 側で生成）
            try {
              const statusParams = new URLSearchParams();
              statusParams.set('action', 'status');
              statusParams.set('bust', '1');
              if (pParam) statusParams.set('p', pParam);
              if (tsParam) statusParams.set('ts', tsParam);
              if (sigParam) statusParams.set('sig', sigParam);
              if (caseIdParam) statusParams.set('caseId', caseIdParam);
              statusParams.set('lineId', lineId);
              fetch(`/api/status?${statusParams.toString()}`, {
                method: 'GET',
                headers: { 'x-line-id': lineId },
                cache: 'no-store',
              }).catch(() => {});
            } catch {}
            store.setStatus(storeKey, 'done');
            if (formId && formId !== storeKey) store.setStatus(formId, 'done');
            if (formKeyParam && formKeyParam !== storeKey) store.setStatus(formKeyParam, 'done');
            router.replace('/form');
            return;
          }
        } catch (e) {
          console.error('intake_complete error', e);
        }
      }
      // フォールバック：受付以外 or 失敗時も一覧へ戻す
      if (storeKey && storeKey !== 'unknown') {
        const ackParams = new URLSearchParams();
        ackParams.set('action', 'form_ack');
        if (pParam) ackParams.set('p', pParam);
        if (tsParam) ackParams.set('ts', tsParam);
        if (sigParam) ackParams.set('sig', sigParam);
        if (caseIdParam) ackParams.set('caseId', caseIdParam);
        ackParams.set('formKey', storeKey);
        ackParams.set('bust', '1');
        if (lineId) ackParams.set('lineId', lineId);
        if (formId && !ackParams.has('formId')) ackParams.set('formId', formId);
        try {
          fetch(`/api/status?${ackParams.toString()}`, { method: 'GET', cache: 'no-store' }).catch(() => {});
        } catch (e) {
          console.error('form_ack error', e);
        }
      }
      store.setStatus(storeKey, 'done');
      if (formId && formId !== storeKey) store.setStatus(formId, 'done');
      if (formKeyParam && formKeyParam !== storeKey) store.setStatus(formKeyParam, 'done');
      router.replace('/form');
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [form, formId, formKeyParam, lineId, storeKey]);

  return <p>送信ありがとうございました。処理が完了すると一覧に戻ります…</p>;
}

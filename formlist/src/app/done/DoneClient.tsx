'use client';

import { useEffect } from 'react';
import { useRouter, useSearchParams } from 'next/navigation';
import { makeProgressStore } from '@/lib/progressStore';

const PENDING_STORAGE_KEY = 'formlist:pendingForm';
const PENDING_TTL_MS = 10 * 60 * 1000;

type PendingFormInfo = {
  formKey?: string;
  storeKey?: string;
  formId?: string;
  lineId?: string;
  caseId?: string;
  savedAt?: number;
};

export default function DoneClient({ lineId }: { lineId: string }) {
  const search = useSearchParams();
  const router = useRouter();
  const formId = search.get('formId') || 'unknown';
  const formKeyParam = search.get('formKey') || '';
  const storeKeyParam = search.get('storeKey') || '';
  const form = search.get('form') || '';
  const storeKey = storeKeyParam || formKeyParam || 'unknown';
  const pParam = search.get('p') || '';
  const tsParam = search.get('ts') || '';
  const sigParam = search.get('sig') || '';
  const caseIdParam = search.get('caseId') || '';
  const queryLineId = search.get('lineId') || '';

  const store = makeProgressStore(lineId)();
  const normalizeCaseId = (value: string | null | undefined) => {
    if (!value) return '';
    const digits = String(value).replace(/\D/g, '');
    if (!digits) return '';
    return digits.slice(-4).padStart(4, '0');
  };
  const normalizedCaseIdParam = normalizeCaseId(caseIdParam);

  useEffect(() => {
    if (!lineId) return;
    let cancelled = false;

    const readPending = (): { pending: PendingFormInfo | null; expired: boolean } => {
      if (typeof window === 'undefined') return { pending: null, expired: false };
      try {
        const raw = window.localStorage.getItem(PENDING_STORAGE_KEY);
        if (!raw) return { pending: null, expired: false };
        const parsed = JSON.parse(raw) as PendingFormInfo;
        if (!parsed || typeof parsed.savedAt !== 'number') return { pending: null, expired: false };
        if (Date.now() - parsed.savedAt > PENDING_TTL_MS) {
          window.localStorage.removeItem(PENDING_STORAGE_KEY);
          return { pending: parsed, expired: true };
        }
        if (parsed.lineId && parsed.lineId !== lineId) return { pending: null, expired: false };
        return { pending: parsed, expired: false };
      } catch (e) {
        console.error('read pendingForm failed', e);
        return { pending: null, expired: false };
      }
    };

    const clearPending = () => {
      if (typeof window === 'undefined') return;
      try {
        window.localStorage.removeItem(PENDING_STORAGE_KEY);
      } catch (e) {
        console.error('clear pendingForm failed', e);
      }
    };

    const markDone = (key: string | null | undefined) => {
      if (!key || key === 'unknown') return;
      if (cancelled) return;
      store.setStatus(key, 'done');
    };

    const runIntakeComplete = async (localKey: string | null | undefined) => {
      const intakeFormId = process.env.NEXT_PUBLIC_INTAKE_FORM_ID;
      const isIntake = form === 'intake' || (intakeFormId && formId === intakeFormId);
      if (!isIntake) return false;
      try {
        const r = await fetch('/api/intake/complete', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ lineId }),
          cache: 'no-store',
        });
        let data: { ok?: boolean } | null = null;
        try {
          data = (await r.json()) as { ok?: boolean };
        } catch {}
        console.log('intake_complete:', r.status, data);
        const ok = r.ok && (data?.ok ?? true);
        if (!ok) return false;
        try {
          const statusParams = new URLSearchParams();
          statusParams.set('action', 'status');
          statusParams.set('bust', '1');
          statusParams.set('lineId', lineId);
          const normalizedCaseId = normalizedCaseIdParam;
          if (pParam) statusParams.set('p', pParam);
          if (tsParam) statusParams.set('ts', tsParam);
          if (sigParam) statusParams.set('sig', sigParam);
          if (normalizedCaseId) statusParams.set('caseId', normalizedCaseId);
          await fetch(`/api/status?${statusParams.toString()}`, {
            method: 'GET',
            headers: { 'x-line-id': lineId },
            cache: 'no-store',
          });
        } catch (e) {
          console.error('status GET error (intake)', e);
        }
        if (localKey) markDone(localKey);
        clearPending();
        if (!cancelled) router.replace('/form');
        return true;
      } catch (e) {
        console.error('intake_complete error', e);
        return false;
      }
    };

    const callGetFlow = async () => {
      const ackFormKey = storeKeyParam || formKeyParam || '';
      if (!ackFormKey || !pParam || !tsParam || !sigParam) return false;
      const normalizedCaseId = normalizeCaseId(caseIdParam);
      const ackParams = new URLSearchParams();
      ackParams.set('action', 'form_ack');
      ackParams.set('formKey', ackFormKey);
      ackParams.set('p', pParam);
      ackParams.set('ts', tsParam);
      ackParams.set('sig', sigParam);
      ackParams.set('lineId', lineId);
      if (normalizedCaseId) ackParams.set('caseId', normalizedCaseId);
      let ackOk = false;
      try {
        const r = await fetch(`/api/status?${ackParams.toString()}`, {
          method: 'GET',
          headers: { 'x-line-id': lineId },
          cache: 'no-store',
        });
        ackOk = r.ok;
      } catch (e) {
        console.error('form_ack GET error', e);
      }

      const statusParams = new URLSearchParams();
      statusParams.set('action', 'status');
      statusParams.set('bust', '1');
      statusParams.set('p', pParam);
      statusParams.set('ts', tsParam);
      statusParams.set('sig', sigParam);
      statusParams.set('lineId', lineId);
      if (normalizedCaseId) statusParams.set('caseId', normalizedCaseId);
      let statusOk = false;
      try {
        const r = await fetch(`/api/status?${statusParams.toString()}`, {
          method: 'GET',
          headers: { 'x-line-id': lineId },
          cache: 'no-store',
        });
        statusOk = r.ok;
      } catch (e) {
        console.error('status GET error', e);
      }
      return ackOk || statusOk;
    };

    const callPostFlow = async (pending: PendingFormInfo) => {
      const ackFormKey = pending.formKey || storeKeyParam || formKeyParam || '';
      if (!ackFormKey) {
        console.warn('form_ack POST skipped: missing formKey');
        return false;
      }
      const caseId = normalizeCaseId(pending.caseId || caseIdParam || '');
      if (!caseId) {
        console.warn('form_ack POST skipped: missing caseId', { ackFormKey, pending });
        return false;
      }
      const baseBody: Record<string, string | number> = {
        lineId,
        caseId,
        mode: 'server',
      };
      let ackOk = false;
      try {
        const r = await fetch('/api/status', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'x-line-id': lineId },
          body: JSON.stringify({ ...baseBody, action: 'form_ack', formKey: ackFormKey }),
          cache: 'no-store',
        });
        ackOk = r.ok;
      } catch (e) {
        console.error('form_ack POST error', e);
      }
      let statusOk = false;
      try {
        const r = await fetch('/api/status', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'x-line-id': lineId },
          body: JSON.stringify({ ...baseBody, action: 'status', bust: 1 }),
          cache: 'no-store',
        });
        statusOk = r.ok;
      } catch (e) {
        console.error('status POST error', e);
      }
      return ackOk || statusOk;
    };

    const callStatusPostOnly = async (caseId: string) => {
      const normalizedCaseId = normalizeCaseId(caseId);
      if (!normalizedCaseId) {
        console.warn('status POST skipped: missing caseId', { lineId, caseIdParam, caseId });
        return false;
      }
      const body = {
        action: 'status',
        bust: 1,
        lineId,
        caseId: normalizedCaseId,
        mode: 'server' as const,
      };
      try {
        const r = await fetch('/api/status', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'x-line-id': lineId },
          body: JSON.stringify(body),
          cache: 'no-store',
        });
        return r.ok;
      } catch (e) {
        console.error('status POST error', e);
        return false;
      }
    };

    (async () => {
      const hasSignature = Boolean(pParam && tsParam && sigParam && (storeKeyParam || formKeyParam));
      const pendingResult = hasSignature ? { pending: null, expired: false } : readPending();
      const pending = pendingResult && !hasSignature ? pendingResult.pending : null;
      const expired = pendingResult && !hasSignature ? pendingResult.expired : false;
      const localKey =
        storeKey !== 'unknown'
          ? storeKey
          : pending?.formKey || pending?.storeKey || formKeyParam || storeKeyParam || '';
      const intakeHandled = await runIntakeComplete(localKey);
      if (cancelled) return;
      if (intakeHandled) return;

      let completed = false;

      if (hasSignature) {
        completed = (await callGetFlow()) ?? false;
        clearPending();
      } else if (pending && !expired) {
        completed = (await callPostFlow(pending)) ?? false;
        clearPending();
      } else if (pending && expired) {
        completed = (await callStatusPostOnly(pending.caseId || caseIdParam || '')) ?? false;
        clearPending();
      } else if (!pending && normalizedCaseIdParam) {
        completed = (await callStatusPostOnly(normalizedCaseIdParam)) ?? false;
      } else {
        console.warn('form_ack skipped: context not found', {
          hasSignature,
          queryLineId,
          storeKeyParam,
        });
      }

      if (completed) {
        markDone(localKey);
      }
      if (!cancelled) router.replace('/form');
    })();

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    form,
    formId,
    formKeyParam,
    storeKey,
    storeKeyParam,
    pParam,
    tsParam,
    sigParam,
    caseIdParam,
    queryLineId,
    lineId,
    router,
  ]);

  return <p>送信ありがとうございました。処理が完了すると一覧に戻ります…</p>;
}

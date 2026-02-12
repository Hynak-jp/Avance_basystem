'use client';

import React from 'react';
import ProgressBar from './ProgressBar';
import ResetProgressButton from './ResetProgressButton';
import { makeProgressStore } from '@/lib/progressStore';
import FormCard from './FormCard';
import { useCaseFormsStatus } from '@/hooks/useCaseFormsStatus';

type FormProgressItem = {
  formId: string;
  formKey?: string;
  storeKey?: string;
  title: string;
  description: string;
  baseUrl: string;
  signedHref?: string;
  disabled?: boolean;
  disabledReason?: string;
  completed?: boolean;
};

type Props = {
  lineId: string;
  displayName?: string;
  caseId?: string | null;
  forms: FormProgressItem[];
};

type DraftStatusResponse = {
  ok?: boolean;
  status?: 'READY' | 'GENERATING' | 'ERROR';
  viewUrl?: string;
  message?: string;
};

const normalizeKey = (f: { formKey?: string; storeKey?: string; formId: string }) =>
  String(f.formKey || f.storeKey || f.formId || '')
    .trim()
    .toLowerCase();

const DRAFT_SUPPORTED_FORM_KEYS = new Set([
  's2002_userform',
  's2010_p1_career',
  's2010_p2_cause',
  's2011_income_m1',
  's2011_income_m2',
]);

const isDraftSupportedForm = (f: { formKey?: string; storeKey?: string; formId: string }) =>
  DRAFT_SUPPORTED_FORM_KEYS.has(normalizeKey(f));

export default function FormProgressClient({ lineId, displayName, caseId, forms }: Props) {
  const store = makeProgressStore(lineId)();
  const { formsMap } = useCaseFormsStatus(caseId ?? undefined, lineId);
  const [draftStateMap, setDraftStateMap] = React.useState<Record<string, DraftStatusResponse>>({});

  const doneCount = forms.filter((form) => {
    const isRepeatableDoc =
      String(form.formKey || form.storeKey || form.formId).toLowerCase() === 'doc_payslip';
    if (isRepeatableDoc) return false;
    const key = form.storeKey || form.formKey || form.formId;
    const serverRow = form.formKey ? formsMap.get(form.formKey) : undefined;
    if (serverRow) {
      if (!serverRow.canEdit && (serverRow.status === 'submitted' || serverRow.status === 'closed')) {
        return true;
      }
      if (serverRow.status === 'reopened') {
        return false;
      }
    }
    if (form.completed) return true;
    return store.statusByForm[key] === 'done';
  }).length;

  // Bootstrap on first render (client), redundant with server prefetch but safe
  // Stores nothing globally yet; ensures endpoint works per 2-2.
  React.useEffect(() => {
    (async () => {
      try {
        await fetch('/api/bootstrap', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ lineId, displayName }),
        });
      } catch {}
    })();
  }, [lineId, displayName]);

  React.useEffect(() => {
    forms.forEach((form) => {
      const key = form.storeKey || form.formKey || form.formId;
      const isRepeatableDoc = String(key).toLowerCase() === 'doc_payslip';
      if (isRepeatableDoc) return;
      if (form.completed && store.statusByForm[key] !== 'done') {
        store.setStatus(key, 'done');
      }
    });
  }, [forms, store]);

  React.useEffect(() => {
    forms.forEach((form) => {
      const key = form.storeKey || form.formKey || form.formId;
      const isRepeatableDoc = String(key).toLowerCase() === 'doc_payslip';
      if (isRepeatableDoc) return;
      const serverRow = form.formKey ? formsMap.get(form.formKey) : undefined;
      if (!serverRow) return;
      if (!serverRow.canEdit && (serverRow.status === 'submitted' || serverRow.status === 'closed')) {
        if (store.statusByForm[key] !== 'done') {
          store.setStatus(key, 'done');
        }
      } else if (serverRow.status === 'reopened') {
        if (store.statusByForm[key] === 'done') {
          store.setStatus(key, 'in_progress');
        }
      }
    });
  }, [forms, formsMap, store]);

  React.useEffect(() => {
    let cancelled = false;
    let timer: ReturnType<typeof setTimeout> | null = null;
    let tries = 0;
    const maxPolls = 4;
    const targetKeys = Array.from(
      new Set(forms.filter((f) => isDraftSupportedForm(f)).map((f) => normalizeKey(f)))
    );
    if (!targetKeys.length) {
      return () => {};
    }
    if (!caseId) {
      setDraftStateMap((prev) => {
        const next = { ...prev };
        targetKeys.forEach((k) => {
          next[k] = { status: 'GENERATING', message: 'case_not_ready' };
        });
        return next;
      });
      return () => {};
    }

    const fetchDraftAll = async () => {
      try {
        const rows = await Promise.all(
          targetKeys.map(async (formKey) => {
            const params = new URLSearchParams({ formKey, caseId: String(caseId) });
            const res = await fetch(`/api/draft/status?${params.toString()}`, { cache: 'no-store' });
            const json = (await res.json()) as DraftStatusResponse;
            if (!res.ok) {
              return [formKey, { status: 'GENERATING', message: json.message || `http_${res.status}` } as DraftStatusResponse] as const;
            }
            return [formKey, { status: json.status || 'GENERATING', viewUrl: json.viewUrl, message: json.message } as DraftStatusResponse] as const;
          })
        );
        if (cancelled) return;
        const nextMap: Record<string, DraftStatusResponse> = {};
        rows.forEach(([k, v]) => {
          nextMap[k] = v;
        });
        setDraftStateMap((prev) => ({ ...prev, ...nextMap }));
        const hasGenerating = rows.some(([, v]) => (v.status || 'GENERATING') === 'GENERATING');
        if (hasGenerating && tries < maxPolls) {
          tries += 1;
          timer = setTimeout(fetchDraftAll, 15000);
        }
      } catch (e) {
        if (cancelled) return;
        console.error('[form] draft status fetch failed', e);
        setDraftStateMap((prev) => {
          const next = { ...prev };
          targetKeys.forEach((k) => {
            next[k] = { status: 'GENERATING', message: 'status_fetch_failed' };
          });
          return next;
        });
      }
    };

    void fetchDraftAll();
    return () => {
      cancelled = true;
      if (timer) clearTimeout(timer);
    };
  }, [forms, caseId]);

  return (
    <>
      <ProgressBar total={forms.length} done={doneCount} />
      <div className="grid gap-6 md:grid-cols-2">
        {forms.map((form) => {
          const normalized = normalizeKey(form);
          const draftState = draftStateMap[normalized];
          const draftEnabled = isDraftSupportedForm(form);
          return (
            <FormCard
              key={form.storeKey || form.formId}
              formId={form.formId}
              title={form.title}
              description={form.description}
              baseUrl={form.baseUrl}
              signedHref={form.signedHref}
              lineId={lineId}
              caseId={caseId}
              disabled={form.disabled}
              disabledReason={form.disabledReason}
              completed={form.completed}
              formKey={form.formKey}
              storeKey={form.storeKey || form.formKey || form.formId}
              serverStatus={form.formKey ? formsMap.get(form.formKey) : undefined}
              draftStatus={draftEnabled ? draftState?.status ?? null : null}
              draftViewUrl={draftEnabled ? draftState?.viewUrl : undefined}
              draftMessage={draftEnabled ? draftState?.message : undefined}
            />
          );
        })}
      </div>

      <ResetProgressButton onReset={store.resetAll} />
    </>
  );
}

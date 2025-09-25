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

export default function FormProgressClient({ lineId, displayName, caseId, forms }: Props) {
  const store = makeProgressStore(lineId)();
  const { formsMap } = useCaseFormsStatus(caseId ?? undefined, lineId);

  const doneCount = forms.filter((form) => {
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
      if (form.completed && store.statusByForm[key] !== 'done') {
        store.setStatus(key, 'done');
      }
    });
  }, [forms, store]);

  React.useEffect(() => {
    forms.forEach((form) => {
      const key = form.storeKey || form.formKey || form.formId;
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

  return (
    <>
      <ProgressBar total={forms.length} done={doneCount} />
      <div className="grid gap-6 md:grid-cols-2">
        {forms.map((form) => (
          <FormCard
            key={form.storeKey || form.formId}
            formId={form.formId}
            title={form.title}
            description={form.description}
            baseUrl={form.baseUrl}
            signedHref={form.signedHref}
            lineId={lineId}
            disabled={form.disabled}
            disabledReason={form.disabledReason}
            completed={form.completed}
            formKey={form.formKey}
            storeKey={form.storeKey || form.formKey || form.formId}
            serverStatus={form.formKey ? formsMap.get(form.formKey) : undefined}
          />
        ))}
      </div>

      <ResetProgressButton onReset={store.resetAll} />
    </>
  );
}

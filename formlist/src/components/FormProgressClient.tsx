'use client';

import React from 'react';
import ProgressBar from './ProgressBar';
import ResetProgressButton from './ResetProgressButton';
import { makeProgressStore } from '@/lib/progressStore';
import FormCard from './FormCard';

type Props = {
  lineId: string;
  displayName?: string;
  forms: Array<{
    formId: string;
    title: string;
    description: string;
    baseUrl: string;
    href?: string;
    disabled?: boolean;
    disabledReason?: string;
  }>;
};

export default function FormProgressClient({ lineId, displayName, forms }: Props) {
  const store = makeProgressStore(lineId)();
  const doneCount = forms.filter((f) => store.statusByForm[f.formId] === 'done').length;

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

  return (
    <>
      <ProgressBar total={forms.length} done={doneCount} />
      <div className="grid gap-6 md:grid-cols-2">
        {forms.map((form) => (
          <FormCard
            key={form.formId}
            formId={form.formId}
            title={form.title}
            description={form.description}
            baseUrl={form.baseUrl}
            href={form.href}
            lineId={lineId}
            disabled={form.disabled}
            disabledReason={form.disabledReason}
          />
        ))}
      </div>

      <ResetProgressButton onReset={store.resetAll} />
    </>
  );
}

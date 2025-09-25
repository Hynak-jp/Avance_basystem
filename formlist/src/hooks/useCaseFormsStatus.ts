'use client';

import { useCallback, useEffect, useMemo, useState } from 'react';

export type CaseFormStatus = {
  caseId: string;
  form_key: string;
  status?: 'submitted' | 'reopened' | 'closed' | '';
  canEdit: boolean;
  reopened_at?: string | null;
  locked_reason?: string | null;
  reopen_until?: string | null;
  last_seq?: number;
};

export type CaseFormsStatusResponse = {
  ok: boolean;
  caseId: string;
  forms: CaseFormStatus[];
};

type HookState = {
  data: CaseFormsStatusResponse | null;
  isLoading: boolean;
  error: unknown;
};

export function useCaseFormsStatus(caseId?: string | null, lineId?: string | null) {
  const [state, setState] = useState<HookState>({ data: null, isLoading: false, error: null });

  const fetchStatus = useCallback(async () => {
    if (!caseId && !lineId) {
      setState((prev) => ({ ...prev, data: null }));
      return null;
    }
    setState((prev) => ({ ...prev, isLoading: true, error: null }));
    try {
      const params = new URLSearchParams();
      if (caseId) params.set('caseId', caseId);
      if (lineId) params.set('lineId', lineId);
      const res = await fetch(`/api/bas-status?${params.toString()}`, { cache: 'no-store' });
      const json = (await res.json()) as CaseFormsStatusResponse;
      setState({ data: json, isLoading: false, error: null });
      return json;
    } catch (error) {
      setState({ data: null, isLoading: false, error });
      return null;
    }
  }, [caseId, lineId]);

  useEffect(() => {
    void fetchStatus();
  }, [fetchStatus]);

  const getByKey = useCallback(
    (formKey: string) => state.data?.forms?.find((f) => f.form_key === formKey),
    [state.data]
  );

  const formsMap = useMemo(() => {
    if (!state.data?.forms) return new Map<string, CaseFormStatus>();
    return new Map(state.data.forms.map((row) => [row.form_key, row]));
  }, [state.data]);

  return {
    data: state.data,
    isLoading: state.isLoading,
    error: state.error,
    refresh: fetchStatus,
    getByKey,
    formsMap,
  } as const;
}

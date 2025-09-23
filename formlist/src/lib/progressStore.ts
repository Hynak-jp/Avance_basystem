'use client';

import { create } from 'zustand';
import { persist } from 'zustand/middleware';

export type FormStatus = 'not_started' | 'in_progress' | 'done';

type Store = {
  statusByForm: Record<string, FormStatus>;
  setStatus: (formId: string, status: FormStatus) => void;
  resetAll: () => void;
};

// lineIdごとに別ストアを生成
export const makeProgressStore = (lineId: string) =>
  create<Store>()(
    persist<Store>(
      (set, get) => ({
        statusByForm: {},
        setStatus: (formId: string, status: FormStatus) =>
          set({ statusByForm: { ...get().statusByForm, [formId]: status } }),
        resetAll: () => set({ statusByForm: {} }),
      }),
      {
        name: `progress_${lineId}`,
      }
    )
  );

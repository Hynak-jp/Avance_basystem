'use client';

import { create } from 'zustand';
import { persist } from 'zustand/middleware';

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export const isValidEmail = (value: string) => EMAIL_REGEX.test(value.trim());

type Store = {
  email: string | null;
  ownerLineId: string | null;
  setEmail: (email: string, ownerLineId: string) => void;
  setEmailOnce: (email: string, ownerLineId: string) => void;
  setEmailFromIntake: (email: string, ownerLineId: string) => void;
  clear: () => void;
};

type PersistedState = {
  email: string | null;
  ownerLineId: string | null;
};

const extractPersistedState = (persisted: unknown): PersistedState | null => {
  if (!persisted || typeof persisted !== 'object') return null;
  if ('state' in persisted) {
    const state = (persisted as { state?: unknown }).state;
    if (state && typeof state === 'object') {
      return state as PersistedState;
    }
    return null;
  }
  return persisted as PersistedState;
};

export const useUserEmailStore = create<Store>()(
  persist(
    (set, get) => ({
      email: null,
      ownerLineId: null,
      setEmail: (email, ownerLineId) => set({ email, ownerLineId }),
      setEmailOnce: (email, ownerLineId) => {
        const curOwner = get().ownerLineId;
        if (curOwner && curOwner !== ownerLineId) return;
        const current = get().email;
        if (current && current.trim().length > 0) return;
        set({ email, ownerLineId });
      },
      setEmailFromIntake: (email, ownerLineId) => {
        const curOwner = get().ownerLineId;
        if (curOwner && curOwner !== ownerLineId) return;
        set({ email, ownerLineId });
      },
      clear: () => set({ email: null, ownerLineId: null }),
    }),
    {
      name: 'formlist:userEmail',
      version: 2,
      partialize: (state) => ({ email: state.email, ownerLineId: state.ownerLineId }),
      migrate: (persisted, version) => {
        if (typeof version !== 'number' || version < 2) {
          return { email: null, ownerLineId: null };
        }
        const data = extractPersistedState(persisted);
        return {
          email: data?.email ?? null,
          ownerLineId: data?.ownerLineId ?? null,
        };
      },
    }
  )
);

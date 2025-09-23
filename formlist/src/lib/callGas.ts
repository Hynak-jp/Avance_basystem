// src/lib/callGas.ts
export async function callGas(p: { lineId: string; idemKey?: string; requestId?: string }) {
  const body = {
    ...p,
    secret: process.env.GAS_SHARED_SECRET,
    idemKey: p.idemKey ?? new Date().toISOString() + '#' + crypto.randomUUID(),
    requestId: p.requestId ?? crypto.randomUUID(),
  };

  const r = await fetch(process.env.NEXT_PUBLIC_GAS_URL!, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });

  const j = await r.json();
  if (!j.ok) throw new Error(j.error || `gas_status_${j.status}`);
  return j as { ok: true; folderId: string; requestId?: string };
}

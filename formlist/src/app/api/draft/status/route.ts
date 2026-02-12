import crypto from 'node:crypto';
import { NextRequest, NextResponse } from 'next/server';
import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import { b64UrlFromString, payloadV2, signV2, tsSec } from '@/lib/sig';

export const runtime = 'nodejs';

type GasJson = Record<string, unknown>;

function normalizeCaseId(value: string | null | undefined) {
  if (!value) return '';
  const digits = String(value).replace(/\D/g, '');
  if (!digits) return '';
  return digits.slice(-4).padStart(4, '0');
}

function b64Url(input: Buffer) {
  return input.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function signViewToken(lineId: string, caseId: string, formKey: string, exp: number) {
  const secret = process.env.DRAFT_VIEW_SECRET || process.env.BOOTSTRAP_SECRET || process.env.BAS_API_HMAC_SECRET || '';
  if (!secret) throw new Error('missing_draft_view_secret');
  const payload = [lineId, caseId, formKey, String(exp)].join('|');
  return b64Url(crypto.createHmac('sha256', secret).update(payload, 'utf8').digest());
}

async function callGasGet(action: string, lineId: string, caseId: string, extra?: Record<string, string>) {
  const endpoint = process.env.GAS_ENDPOINT;
  if (!endpoint) throw new Error('missing_gas_endpoint');
  const ts = tsSec();
  const pld = payloadV2(lineId, caseId, ts);
  const url = new URL(endpoint);
  url.searchParams.set('action', action);
  url.searchParams.set('ts', ts);
  url.searchParams.set('p', b64UrlFromString(pld));
  url.searchParams.set('sig', signV2(pld));
  Object.entries(extra || {}).forEach(([k, v]) => {
    if (v !== undefined && v !== null) url.searchParams.set(k, String(v));
  });
  const res = await fetch(url.toString(), { method: 'GET', cache: 'no-store' });
  const text = await res.text();
  let json: GasJson = {};
  try {
    json = JSON.parse(text) as GasJson;
  } catch {
    json = { ok: false, error: 'invalid_gas_response', text };
  }
  return { status: res.status, json };
}

export async function GET(req: NextRequest) {
  try {
    const session = await getServerSession(authOptions);
    const lineId = typeof session?.lineId === 'string' ? session.lineId : '';
    if (!lineId) return NextResponse.json({ ok: false, error: 'unauthorized' }, { status: 401 });

    const { searchParams } = new URL(req.url);
    const formKey = String(searchParams.get('formKey') || 's2002_userform').trim();
    let caseId = normalizeCaseId(searchParams.get('caseId'));

    if (!caseId) {
      const statusRes = await callGasGet('status', lineId, '', { bust: '1' });
      const statusJson = statusRes.json;
      caseId = normalizeCaseId(
        String((statusJson.caseId as string) || (statusJson.activeCaseId as string) || '')
      );
    }
    if (!caseId) {
      return NextResponse.json(
        { ok: true, formKey, status: 'GENERATING', message: 'case_not_ready' },
        { status: 200 }
      );
    }

    const draftRes = await callGasGet('draft_status', lineId, caseId, { formKey });
    const draft = draftRes.json;
    const rawStatus = String((draft.status as string) || '').trim().toUpperCase();
    const status =
      rawStatus === 'READY' || rawStatus === 'ERROR' || rawStatus === 'GENERATING'
        ? rawStatus
        : String((draft.error as string) || '').trim() === 'draft_error'
        ? 'ERROR'
        : 'GENERATING';
    const message =
      String((draft.message as string) || '').trim() ||
      String((draft.error as string) || '').trim() ||
      (draftRes.status >= 400 ? `http_${draftRes.status}` : undefined);
    const payload: Record<string, unknown> = {
      ok: true,
      formKey,
      caseId,
      status,
      updatedAt: (draft.updatedAt as string) || undefined,
      message: message || undefined,
    };

    if (status === 'READY') {
      const now = Math.floor(Date.now() / 1000);
      const exp = now + 30 * 60;
      const sig = signViewToken(lineId, caseId, formKey, exp);
      payload.viewUrl = `/api/draft/view?formKey=${encodeURIComponent(formKey)}&caseId=${encodeURIComponent(caseId)}&exp=${exp}&sig=${encodeURIComponent(sig)}`;
    }

    return NextResponse.json(payload, { status: 200 });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error('[draft/status] failed', message);
    return NextResponse.json({ ok: false, error: message }, { status: 500 });
  }
}

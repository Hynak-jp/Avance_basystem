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

function timingSafeEqualB64Url(a: string, b: string) {
  const aa = Buffer.from(String(a || ''), 'utf8');
  const bb = Buffer.from(String(b || ''), 'utf8');
  if (aa.length !== bb.length) return false;
  return crypto.timingSafeEqual(aa, bb);
}

function expectedViewSig(lineId: string, caseId: string, formKey: string, exp: number) {
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
    const formKey = String(searchParams.get('formKey') || '').trim();
    const caseId = normalizeCaseId(searchParams.get('caseId'));
    const exp = Number(searchParams.get('exp') || '0');
    const sig = String(searchParams.get('sig') || '').trim();

    if (!formKey || !caseId || !Number.isFinite(exp) || exp <= 0 || !sig) {
      return NextResponse.json({ ok: false, error: 'invalid_params' }, { status: 400 });
    }

    const now = Math.floor(Date.now() / 1000);
    if (exp < now) {
      return NextResponse.json({ ok: false, error: 'expired' }, { status: 410 });
    }

    // URL共有されても、現在のログインユーザー(lineId)で再検証する。
    const expectedSig = expectedViewSig(lineId, caseId, formKey, exp);
    if (!timingSafeEqualB64Url(expectedSig, sig)) {
      return NextResponse.json({ ok: false, error: 'forbidden' }, { status: 403 });
    }

    const activeRes = await callGasGet('status', lineId, '', { bust: '1' });
    const activeJson = activeRes.json;
    const activeCaseId = normalizeCaseId(
      String((activeJson.caseId as string) || (activeJson.activeCaseId as string) || '')
    );
    if (!activeCaseId || activeCaseId !== caseId) {
      return NextResponse.json({ ok: false, error: 'case_mismatch' }, { status: 403 });
    }

    const statusRes = await callGasGet('draft_status', lineId, caseId, { formKey });
    const statusJson = statusRes.json;
    const draftStatus = String((statusJson.status as string) || 'GENERATING');
    if (draftStatus === 'ERROR') {
      return NextResponse.json({ ok: false, error: 'draft_error' }, { status: 409 });
    }
    if (draftStatus !== 'READY') {
      return NextResponse.json({ ok: false, error: 'draft_not_ready' }, { status: 404 });
    }

    const pdfRes = await callGasGet('draft_pdf', lineId, caseId, { formKey });
    const pdfJson = pdfRes.json;
    if (pdfRes.status === 413 || String(pdfJson.error || '') === 'too_large') {
      return NextResponse.json(
        {
          ok: false,
          error: 'pdf_too_large',
          message: 'PDFサイズが大きすぎたため、表示できませんでした。',
        },
        { status: 413 }
      );
    }
    if (!pdfRes.status || pdfRes.status >= 400 || !pdfJson.ok) {
      return NextResponse.json({ ok: false, error: 'pdf_fetch_failed' }, { status: 404 });
    }
    const b64 = String((pdfJson.data as string) || '');
    if (!b64) {
      return NextResponse.json({ ok: false, error: 'empty_pdf' }, { status: 404 });
    }
    const fileName = String((pdfJson.fileName as string) || 'draft.pdf').replace(/[^\w.-]+/g, '_');
    const pdfBuffer = Buffer.from(b64, 'base64');

    return new NextResponse(pdfBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `inline; filename="${fileName}"`,
        'Cache-Control': 'no-store',
        'X-Content-Type-Options': 'nosniff',
      },
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error('[draft/view] failed', message);
    return NextResponse.json({ ok: false, error: message }, { status: 500 });
  }
}

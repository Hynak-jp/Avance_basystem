import crypto from 'crypto';
import { NextRequest, NextResponse } from 'next/server';
import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';

export const runtime = 'nodejs';

const GAS = process.env.GAS_ENDPOINT!;
const SECRET = process.env.BOOTSTRAP_SECRET ?? process.env.TOKEN_SECRET;

const allowedActions = new Set(['status', 'intake_ack', 'form_ack', 'markReopen']);
const allowedParams = new Set([
  'action',
  'p',
  'ts',
  'sig',
  'lineId',
  'line_id',
  'caseId',
  'case_id',
  'formKey',
  'form_key',
  'submission_id',
  'submissionId',
  'bust',
  'formId',
  'mode',
]);
const MAX_PARAM_LENGTH = 4096;

const b64u = (s: string) => s.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

function normalizeCaseId(value: string | null | undefined) {
  if (!value) return '';
  const digits = String(value).replace(/\D/g, '');
  if (!digits) return '';
  return digits.slice(-4).padStart(4, '0');
}

function computeSignature(lineId: string, caseId: string) {
  const ts = Math.floor(Date.now() / 1000).toString();
  const payload = `${lineId}|${caseId || ''}|${ts}`;
  const p = b64u(Buffer.from(payload, 'utf8').toString('base64'));
  const sig = b64u(crypto.createHmac('sha256', SECRET!).update(payload, 'utf8').digest('base64'));
  return { ts, p, sig };
}

function sanitizeParams(req: NextRequest, searchParams: URLSearchParams, forcedAction?: string) {
  const sanitized = new URLSearchParams();
  for (const [key, value] of searchParams.entries()) {
    if (!allowedParams.has(key)) continue;
    if (value.length > MAX_PARAM_LENGTH) {
      return { error: NextResponse.json({ ok: false, error: 'param_too_long', key }, { status: 400 }) };
    }
    if (key === 'p' && value && !/^[A-Za-z0-9_-]*$/.test(value)) {
      return { error: NextResponse.json({ ok: false, error: 'invalid_p' }, { status: 400 }) };
    }
    if (key === 'sig' && value && !/^[A-Za-z0-9_-]{16,}$/.test(value)) {
      return { error: NextResponse.json({ ok: false, error: 'invalid_sig' }, { status: 400 }) };
    }
    sanitized.append(key, value);
  }

  if (forcedAction) sanitized.set('action', forcedAction);
  const action = sanitized.get('action') || forcedAction || searchParams.get('action') || 'status';

  const headerLineId = req.headers.get('x-line-id') || '';
  const headerCaseId = req.headers.get('x-case-id') || '';

  const lineId = sanitized.get('lineId') || sanitized.get('line_id') || headerLineId;
  if (!lineId) {
    return { error: NextResponse.json({ ok: false, error: 'missing lineId' }, { status: 400 }) };
  }
  sanitized.set('lineId', lineId);
  sanitized.delete('line_id');

  const rawCaseId = sanitized.get('caseId') || sanitized.get('case_id') || headerCaseId || '';
  const normalizedCaseId = normalizeCaseId(rawCaseId);
  if (normalizedCaseId) sanitized.set('caseId', normalizedCaseId);
  sanitized.delete('case_id');

  if (!sanitized.has('ts') || !sanitized.has('p') || !sanitized.has('sig')) {
    const { ts, p, sig } = computeSignature(lineId, normalizedCaseId);
    sanitized.set('ts', ts);
    sanitized.set('p', p);
    sanitized.set('sig', sig);
  }

  return { params: sanitized, action };
}

async function forwardToGas(params: URLSearchParams) {
  const url = new URL(GAS);
  params.forEach((value, key) => {
    url.searchParams.append(key, value);
  });
  const r = await fetch(url.toString(), { method: 'GET', redirect: 'follow', cache: 'no-store' });
  const text = await r.text();
  try {
    const json = JSON.parse(text);
    return NextResponse.json(json, { status: r.status });
  } catch {
    return NextResponse.json({ ok: r.ok, text }, { status: r.status });
  }
}

export async function GET(req: NextRequest) {
  try {
    if (!GAS || !SECRET) {
      return NextResponse.json(
        { ok: false, error: 'missing_env', details: { GAS: !!GAS, SECRET: !!SECRET } },
        { status: 500 }
      );
    }

    const { searchParams } = new URL(req.url);
    const requestedAction = searchParams.get('action') || undefined;

    if (requestedAction === 'markReopen') {
      return NextResponse.json({ ok: false, error: 'use_post' }, { status: 405 });
    }

    const { params, action, error } = sanitizeParams(req, searchParams, requestedAction || 'status');
    if (error) return error;
    if (!params || !action) {
      return NextResponse.json({ ok: false, error: 'invalid_request' }, { status: 400 });
    }

    if (!allowedActions.has(action)) {
      return NextResponse.json({ ok: false, error: 'unsupported_action' }, { status: 400 });
    }

    return await forwardToGas(params);
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : String(e);
    return NextResponse.json({ ok: false, error: message }, { status: 500 });
  }
}

export async function POST(req: NextRequest) {
  try {
    if (!GAS || !SECRET) {
      return NextResponse.json(
        { ok: false, error: 'missing_env', details: { GAS: !!GAS, SECRET: !!SECRET } },
        { status: 500 }
      );
    }

    let body: Record<string, unknown>;
    try {
      body = (await req.json()) as Record<string, unknown>;
    } catch {
      return NextResponse.json({ ok: false, error: 'invalid_json' }, { status: 400 });
    }

    const requestedAction = typeof body.action === 'string' ? body.action : undefined;
    if (!requestedAction) {
      return NextResponse.json({ ok: false, error: 'missing_action' }, { status: 400 });
    }

    const paramsFromBody = new URLSearchParams();
    for (const [key, rawValue] of Object.entries(body)) {
      if (rawValue === undefined || rawValue === null) continue;
      if (Array.isArray(rawValue)) {
        for (const value of rawValue) {
          if (value === undefined || value === null) continue;
          paramsFromBody.append(key, String(value));
        }
      } else {
        paramsFromBody.set(key, String(rawValue));
      }
    }
    paramsFromBody.delete('p');
    paramsFromBody.delete('ts');
    paramsFromBody.delete('sig');

    const { params, action, error } = sanitizeParams(req, paramsFromBody, requestedAction);
    if (error) return error;
    if (!params || !action) {
      return NextResponse.json({ ok: false, error: 'invalid_request' }, { status: 400 });
    }

    if (!allowedActions.has(action)) {
      return NextResponse.json({ ok: false, error: 'unsupported_action' }, { status: 400 });
    }

    let sessionLineId: string | undefined;
    if (params.get('mode') === 'server') {
      const caseId = params.get('caseId');
      if (!caseId) {
        return NextResponse.json({ ok: false, error: 'missing_caseId' }, { status: 400 });
      }
      if (!/^\d{4}$/.test(caseId)) {
        return NextResponse.json({ ok: false, error: 'invalid_caseId' }, { status: 400 });
      }
      const session = await getServerSession(authOptions);
      sessionLineId = typeof session?.lineId === 'string' ? session.lineId : undefined;
      if (!sessionLineId) {
        return NextResponse.json({ ok: false, error: 'unauthorized' }, { status: 401 });
      }
      const requestedLineId = params.get('lineId');
      if (requestedLineId && requestedLineId !== sessionLineId) {
        console.warn('lineId_mismatch', { requestedLineId, sessionLineId });
      }
      params.set('lineId', sessionLineId);
      const { ts, p, sig } = computeSignature(sessionLineId, caseId);
      params.set('ts', ts);
      params.set('p', p);
      params.set('sig', sig);
    }

    return await forwardToGas(params);
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : String(e);
    return NextResponse.json({ ok: false, error: message }, { status: 500 });
  }
}

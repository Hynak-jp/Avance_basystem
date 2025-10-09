import { NextRequest, NextResponse } from 'next/server';
import crypto from 'crypto';

export const runtime = 'nodejs';

const GAS = process.env.GAS_ENDPOINT!;
const SECRET = process.env.BOOTSTRAP_SECRET ?? process.env.TOKEN_SECRET;

const b64u = (s: string) => s.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

export async function GET(req: NextRequest) {
  try {
    if (!GAS || !SECRET) {
      return NextResponse.json(
        { ok: false, error: 'missing_env', details: { GAS: !!GAS, SECRET: !!SECRET } },
        { status: 500 }
      );
    }
    const { searchParams } = new URL(req.url);
    const passthroughAction = searchParams.get('action');
    const allowedActions = new Set(['status', 'intake_ack', 'form_ack', 'markReopen']);
    if (passthroughAction && !allowedActions.has(passthroughAction)) {
      return NextResponse.json({ ok: false, error: 'unsupported_action' }, { status: 400 });
    }

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
    ]);
    const MAX_PARAM_LENGTH = 4096;

    const buildSanitizedParams = () => {
      const sanitized = new URLSearchParams();
      for (const [key, value] of searchParams.entries()) {
        if (!allowedParams.has(key)) continue;
        if (value.length > MAX_PARAM_LENGTH) {
          return { error: NextResponse.json({ ok: false, error: 'param_too_long', key }, { status: 400 }) };
        }
        if (key === 'p' && !/^[A-Za-z0-9_-]*$/.test(value)) {
          return { error: NextResponse.json({ ok: false, error: 'invalid_p' }, { status: 400 }) };
        }
        if (key === 'sig' && value && !/^[A-Za-z0-9_-]{16,}$/.test(value)) {
          return { error: NextResponse.json({ ok: false, error: 'invalid_sig' }, { status: 400 }) };
        }
        sanitized.append(key, value);
      }
      if (passthroughAction && !sanitized.has('action')) sanitized.set('action', passthroughAction);
      return { params: sanitized };
    };

    const forwardToGas = async (sanitized: URLSearchParams) => {
      const url = new URL(GAS);
      sanitized.forEach((value, key) => {
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
    };

    const hasSignedParams = searchParams.has('p') && searchParams.has('ts') && searchParams.has('sig');
    if (passthroughAction === 'status' && hasSignedParams) {
      const { params, error } = buildSanitizedParams();
      if (error) return error;
      return await forwardToGas(params!);
    }

    if (passthroughAction && passthroughAction !== 'status') {
      if (passthroughAction === 'markReopen') {
        return NextResponse.json({ ok: false, error: 'use_post' }, { status: 405 });
      }
      const { params, error } = buildSanitizedParams();
      if (error) return error;
      return await forwardToGas(params!);
    }

    const lineId = req.headers.get('x-line-id') ?? '';
    const caseId = req.headers.get('x-case-id') ?? '';
    if (!lineId) return NextResponse.json({ ok: false, error: 'missing_lineId' }, { status: 400 });

    const ts = Math.floor(Date.now() / 1000).toString(); // UNIX 秒
    // V2 payload: lineId|caseId|ts
    const payload = `${lineId}|${caseId}|${ts}`;
    const sig = b64u(crypto.createHmac('sha256', SECRET!).update(payload, 'utf8').digest('base64'));
    const p = b64u(Buffer.from(payload, 'utf8').toString('base64'));

    const url = new URL(GAS);
    url.searchParams.set('action', 'status');
    url.searchParams.set('ts', ts);
    url.searchParams.set('sig', sig);
    url.searchParams.set('p', p);

    const r = await fetch(url.toString(), { method: 'GET', redirect: 'follow', cache: 'no-store' });
    const text = await r.text();
    try {
      const json = JSON.parse(text);
      return NextResponse.json(json, { status: r.status });
    } catch {
      return NextResponse.json({ ok: r.ok, text }, { status: r.status });
    }
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : String(e);
    return NextResponse.json({ ok: false, error: message }, { status: 500 });
  }
}

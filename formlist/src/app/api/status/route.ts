import { NextRequest, NextResponse } from 'next/server';
import crypto from 'crypto';

export const runtime = 'nodejs';

const GAS = process.env.GAS_ENDPOINT!;
const SECRET = process.env.BOOTSTRAP_SECRET ?? process.env.TOKEN_SECRET;

const b64u = (s: string) => s.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

// V1: lineId|ts（status/intake_complete 用）
function signV1(lineId: string, ts: string) {
  const base = `${lineId}|${ts}`;
  const b64 = crypto.createHmac('sha256', SECRET!).update(base, 'utf8').digest('base64');
  return b64u(b64);
}

export async function GET(req: NextRequest) {
  try {
    if (!GAS || !SECRET) {
      return NextResponse.json(
        { ok: false, error: 'missing_env', details: { GAS: !!GAS, SECRET: !!SECRET } },
        { status: 500 }
      );
    }
    const lineId = req.headers.get('x-line-id') || '';
    if (!lineId) return NextResponse.json({ ok: false, error: 'missing_lineId' }, { status: 400 });

    const ts = Date.now().toString();
    const sig = signV1(lineId, ts);

    const url = new URL(GAS);
    url.searchParams.set('action', 'status');
    url.searchParams.set('lineId', lineId);
    url.searchParams.set('ts', ts);
    url.searchParams.set('sig', sig);

    const r = await fetch(url.toString(), { method: 'POST', cache: 'no-store' });
    const text = await r.text();
    try {
      const json = JSON.parse(text);
      return NextResponse.json(json, { status: r.status });
    } catch {
      return NextResponse.json({ ok: r.ok, text }, { status: r.status });
    }
  } catch (e: any) {
    return NextResponse.json({ ok: false, error: String(e?.message ?? e) }, { status: 500 });
  }
}


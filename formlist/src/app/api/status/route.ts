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

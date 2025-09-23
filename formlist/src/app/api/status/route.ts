import { NextRequest, NextResponse } from 'next/server';
import { createHmac } from 'crypto';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

export async function GET(req: NextRequest) {
  try {
    const lineId = req.headers.get('x-line-id') || '';
    if (!lineId) return NextResponse.json({ ok: false, error: 'missing_lineId' }, { status: 400 });

    const SECRET = process.env.BOOTSTRAP_SECRET;
    const GAS_ENDPOINT = process.env.GAS_ENDPOINT;
    if (!SECRET || !GAS_ENDPOINT)
      return NextResponse.json(
        { ok: false, error: 'missing_env', details: { SECRET: !!SECRET, GAS_ENDPOINT: !!GAS_ENDPOINT } },
        { status: 500 }
      );

    const ts = Math.floor(Date.now() / 1000);
    const sig = createHmac('sha256', SECRET).update(`${lineId}|${ts}`, 'utf8').digest('hex');

    const r = await fetch(GAS_ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'cache-control': 'no-store' },
      body: JSON.stringify({ action: 'status', lineId, ts, sig }),
      cache: 'no-store',
    });
    const text = await r.text();
    try {
      const json = JSON.parse(text);
      return NextResponse.json(json, { status: r.ok ? 200 : r.status });
    } catch {
      return NextResponse.json({ ok: r.ok, text }, { status: r.status });
    }
  } catch (e) {
    return NextResponse.json({ ok: false, error: String(e) }, { status: 500 });
  }
}


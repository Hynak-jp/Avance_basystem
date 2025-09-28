// formlist/src/app/api/bootstrap/route.ts
import { NextRequest, NextResponse } from 'next/server';
import crypto from 'crypto';

// Node の crypto を使うので念のため Node ランタイムを明示
export const runtime = 'nodejs';

const GAS_ENDPOINT = process.env.GAS_ENDPOINT!;
const SECRET = process.env.BOOTSTRAP_SECRET ?? process.env.TOKEN_SECRET;

function makePayload(lineId: string, caseId: string | undefined, tsMs: string) {
  return [lineId, caseId ?? '', tsMs].join('|');
}
function toB64Url(b64: string) {
  return b64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
function hmacB64url(secret: string, s: string) {
  // Node の digest('base64url') が使えない環境でも動くフォールバック
  const b64 = crypto.createHmac('sha256', secret).update(s, 'utf8').digest('base64');
  return toB64Url(b64);
}
function b64urlFromString(s: string) {
  return toB64Url(Buffer.from(s, 'utf8').toString('base64'));
}

export async function POST(req: NextRequest) {
  try {
    if (!GAS_ENDPOINT || !SECRET) {
      return NextResponse.json(
        {
          ok: false,
          error: 'missing_env',
          details: { GAS_ENDPOINT: !!GAS_ENDPOINT, BOOTSTRAP_SECRET: !!SECRET },
        },
        { status: 500 }
      );
    }

    // intake（初回受付）から呼ばれる前提
    const { lineId, caseId } = await req.json();
    if (!lineId) return NextResponse.json({ ok: false, error: 'missing_lineId' }, { status: 400 });

    // ts はミリ秒（文字列）
    const ts = Date.now().toString();
    // payload = lineId|caseId|ts（初回は caseId を空でOK）
    const payload = makePayload(lineId, caseId, ts);
    const sig = hmacB64url(SECRET!, payload);
    const p = b64urlFromString(payload);

    // GET /exec?action=bootstrap に統一
    const url = new URL(GAS_ENDPOINT);
    url.searchParams.set('action', 'bootstrap');
    url.searchParams.set('ts', ts);
    url.searchParams.set('sig', sig);
    url.searchParams.set('p', p);

    // デバッグ用の軽いログ（指紋のみ）
    try {
      // biome-ignore lint/suspicious/noConsole: intentional runtime log
      console.log('BAS:formlist:auth', {
        lineId_present: !!lineId,
        caseId_present: !!caseId,
        ts_len: ts.length,
        sig_len: sig.length,
        secret_fp: crypto
          .createHmac('sha256', SECRET!)
          .update('fingerprint')
          .digest('base64url')
          .slice(0, 12),
      });
    } catch {}

    const r = await fetch(url.toString(), {
      method: 'GET',
      headers: { 'cache-control': 'no-store' },
      cache: 'no-store',
    });

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

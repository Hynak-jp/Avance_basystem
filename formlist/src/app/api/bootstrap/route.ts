// formlist/src/app/api/bootstrap/route.ts
import { NextRequest, NextResponse } from 'next/server';
import { createHmac } from 'crypto';

// Node の crypto を使うので念のため Node ランタイムを明示
export const runtime = 'nodejs';

const GAS_ENDPOINT = process.env.GAS_ENDPOINT;
const SECRET = process.env.BOOTSTRAP_SECRET;

function hmacBase64Url(secret: string, lineId: string, ts: number) {
  const base = `${lineId}|${ts}`; // ★ 署名対象は lineId|ts
  return createHmac('sha256', secret).update(base, 'utf8').digest('base64url'); // ★ URL-safe
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

    const { lineId, displayName = '', email = '' } = await req.json();
    if (!lineId) return NextResponse.json({ ok: false, error: 'missing_lineId' }, { status: 400 });

    const ts = Math.floor(Date.now() / 1000); // UNIX秒
    const userKey = String(lineId).slice(0, 6).toLowerCase();

    // 送信する本文（署名対象ではない）
    const payload = { userKey, lineId, displayName, email, ts };
    const body = JSON.stringify(payload);

    // ★ Base64URL 署名を作成
    const sig = hmacBase64Url(SECRET!, lineId, ts);

    // ★ URL に sig, ts を付与（必要なら debug=1 を伝搬）
    const url = new URL(GAS_ENDPOINT!);
    url.searchParams.set('sig', sig);
    url.searchParams.set('ts', String(ts));
    // 両対応（qs/bodyどちらでも取得できる）に寄せるため冪等的に付与
    url.searchParams.set('lineId', lineId);
    url.searchParams.set('userKey', userKey);
    const wantDebug = req.nextUrl.searchParams.get('debug') === '1';
    if (wantDebug) url.searchParams.set('debug', '1');

    const r = await fetch(url.toString(), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'cache-control': 'no-store' },
      body,
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

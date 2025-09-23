// src/app/api/avatar/route.ts
export const runtime = 'edge'; // 速い＆外部fetchに強い

export async function GET(req: Request) {
  const { searchParams } = new URL(req.url);
  const url = searchParams.get('url');
  if (!url) return new Response('Missing url', { status: 400 });

  // 余計なヘッダを送らず取得（referrerなし）
  const upstream = await fetch(url, { redirect: 'follow' });
  if (!upstream.ok) {
    return new Response(`Upstream ${upstream.status}`, { status: 502 });
  }

  const contentType = upstream.headers.get('content-type') ?? 'image/jpeg';
  const body = await upstream.arrayBuffer();

  // 同一オリジンで配る（CSP: img-src 'self' でもOK）
  return new Response(body, {
    headers: {
      'content-type': contentType,
      // 5分キャッシュ（必要に応じて調整）
      'cache-control': 'public, max-age=300, s-maxage=300',
    },
  });
}

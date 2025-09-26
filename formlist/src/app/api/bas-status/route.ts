import { NextResponse } from 'next/server';
import { buildSignedQuery } from '@/lib/bas-sign';

export const dynamic = 'force-dynamic';
export const runtime = 'nodejs';

export async function GET(req: Request) {
  const endpoint = process.env.BAS_API_ENDPOINT;
  if (!endpoint) {
    return NextResponse.json({ ok: false, error: 'BAS_API_ENDPOINT is not configured' }, { status: 500 });
  }

  const { searchParams } = new URL(req.url);
  const lineId = searchParams.get('lineId') || '';
  const caseId = searchParams.get('caseId') || '';

  let query: string;
  try {
    query = buildSignedQuery({ lineId, caseId });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }

  const separator = endpoint.includes('?') ? '&' : '?';
  const url = `${endpoint}${separator}${query}`;

  try {
    const res = await fetch(url, { cache: 'no-store' });
    const json = await res.json();
    return NextResponse.json(json, { status: res.status });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 502 });
  }
}

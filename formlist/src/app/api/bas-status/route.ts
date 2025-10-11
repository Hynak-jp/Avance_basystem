import { NextResponse } from 'next/server';

export const dynamic = 'force-dynamic';
export const runtime = 'nodejs';

export async function GET() {
  return NextResponse.json(
    {
      ok: false,
      error: 'bas-status_endpoint_removed',
      message: 'Use /api/status?action=status with V2 signature instead.',
    },
    { status: 410 }
  );
}

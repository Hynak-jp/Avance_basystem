import { NextResponse } from 'next/server';
import pkg from '../../../../package.json';

export const revalidate = 0;

export async function GET() {
  const commit = process.env.VERCEL_GIT_COMMIT_SHA ?? '';
  const branch = process.env.VERCEL_GIT_COMMIT_REF ?? '';
  const env = process.env.VERCEL_ENV ?? process.env.NODE_ENV ?? 'development';

  return NextResponse.json(
    {
      version: pkg.version,
      commit: commit ? commit.slice(0, 7) : 'dev',
      branch,
      env,
    },
    {
      headers: {
        'Cache-Control': 'no-store',
      },
    },
  );
}

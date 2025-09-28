// formlist/src/lib/formUrl.ts
'use server';
import 'server-only';
import crypto from 'crypto';

const SECRET =
  process.env.BOOTSTRAP_SECRET ??
  process.env.TOKEN_SECRET ??
  (() => {
    throw new Error('Missing BOOTSTRAP_SECRET/TOKEN_SECRET');
  })();

const b64url = (s: string) => s.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
const makePayload = (lineId: string, caseId: string, ts: string) => `${lineId}|${caseId}|${ts}`;
const signV2 = (payload: string) => b64url(crypto.createHmac('sha256', SECRET).update(payload, 'utf8').digest('base64'));

function getOriginSafe(): string {
  return process.env.NEXT_PUBLIC_BASE_URL || '';
}

type MakeFormUrlOptions = {
  redirectUrl?: string; // 例: https://app.example/done?formId=309542
  formId?: string; // 表示用（署名には使わない）
  extra?: Record<string, string>;
};

/**
 * 送信後に戻る /done 側へ V2 署名（p/ts/sig）と lineId/caseId を付与し、
 * その redirect_url を FormMailer 側の baseUrl にクエリで渡す。
 */
export function makeFormUrl(baseUrl: string, lineId: string, caseId: string, opts: MakeFormUrlOptions = {}): string {
  const ts = Date.now().toString();
  const payload = makePayload(lineId, caseId, ts);
  const sig = signV2(payload);
  const p = b64url(Buffer.from(payload, 'utf8').toString('base64'));

  // 1) ユーザーが戻ってくるURL（自サイト）
  const redirect = new URL(opts?.redirectUrl ?? '/', getOriginSafe());
  redirect.searchParams.set('lineId', lineId);
  redirect.searchParams.set('caseId', caseId);
  redirect.searchParams.set('ts', ts);
  redirect.searchParams.set('sig', sig);
  redirect.searchParams.set('p', p);
  if (opts?.formId) redirect.searchParams.set('formId', opts.formId);
  if (opts?.extra) for (const [k, v] of Object.entries(opts.extra)) redirect.searchParams.set(k, v);

  // 2) FormMailer 側のURLに、redirect_url を渡す
  const url = new URL(baseUrl);
  url.searchParams.set('redirect_url', redirect.toString());
  url.searchParams.set('referrer', getOriginSafe());
  return url.toString();
}

/**
 * intake 用。caseId は空文字で署名。
 * intakeBase: intake の FormMailer URL
 * intakeRedirect: intake 完了後に戻す /done のURL
 */
export function makeIntakeUrl(intakeBase: string, intakeRedirect: string): string {
  const ts = Date.now().toString();
  const lineId = '';
  const caseId = '';
  const payload = makePayload(lineId, caseId, ts);
  const sig = signV2(payload);
  const p = b64url(Buffer.from(payload, 'utf8').toString('base64'));

  const redirect = new URL(intakeRedirect);
  redirect.searchParams.set('lineId', lineId);
  redirect.searchParams.set('caseId', caseId);
  redirect.searchParams.set('ts', ts);
  redirect.searchParams.set('sig', sig);
  redirect.searchParams.set('p', p);

  const url = new URL(intakeBase);
  url.searchParams.set('redirect_url', redirect.toString());
  url.searchParams.set('referrer', getOriginSafe());
  return url.toString();
}

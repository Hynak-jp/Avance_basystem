// formlist/src/lib/formUrl.ts
import 'server-only';
import crypto from 'crypto';

const SECRET =
  process.env.BOOTSTRAP_SECRET ??
  process.env.TOKEN_SECRET ??
  (() => {
    throw new Error('Missing BOOTSTRAP_SECRET/TOKEN_SECRET');
  })();

const b64url = (s: string) => s.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
const b64urlFromString = (s: string) => b64url(Buffer.from(s, 'utf8').toString('base64'));
const makePayload = (lineId: string, caseId: string, ts: string) => `${lineId}|${caseId}|${ts}`;
const signV2 = (payload: string) => b64url(crypto.createHmac('sha256', SECRET).update(payload, 'utf8').digest('base64'));
const normalizeCaseId = (raw: string) => {
  if (!raw) return '';
  const onlyDigits = raw.replace(/\D/g, '');
  if (!onlyDigits) return '';
  return onlyDigits.padStart(4, '0');
};

// 追加: originの安全取得（SSR/Edge/Browser全部OK）
export function getOriginSafe(): string | undefined {
  if (typeof window !== 'undefined') return window.location.origin; // client
  const env =
    process.env.NEXT_PUBLIC_SITE_ORIGIN ||
    (process.env.VERCEL_URL ? `https://${process.env.VERCEL_URL}` : undefined) ||
    process.env.NEXT_PUBLIC_BASE_URL ||
    undefined;
  return env; // undefined のままでもOK（後で判定）
}

// 追加: base が空/未定義なら 1引数版で呼ぶ
export function safeURL(input: string, base?: string) {
  if (base && base.trim().length > 0) return new URL(input, base);
  return new URL(input);
}

// ts を常に「UNIX秒」で生成（呼び出し側の曖昧さを排除）
export const tsSec = () => Math.floor(Date.now() / 1000).toString();

type MakeFormUrlOptions = {
  redirectUrl?: string; // 例: https://app.example/done?formId=309542
  formId?: string; // 表示用（署名には使わない）
  extra?: Record<string, string>;
  lineIdQueryKeys?: string[]; // フォームが参照できるクエリ名（デフォルト: line_id）
  caseIdQueryKeys?: string[]; // フォームが参照できるクエリ名（デフォルト: case_id）
};

/**
 * 送信後に戻る /done 側へ V2 署名（p/ts/sig）と lineId/caseId を付与し、
 * その redirect_url を FormMailer 側の baseUrl にクエリで渡す。
 */
export function makeFormUrl(baseUrl: string, lineId: string, caseId: string, opts: MakeFormUrlOptions = {}): string {
  const normalizedCaseId = normalizeCaseId(caseId);
  const ts = tsSec(); // UNIX秒
  const payload = makePayload(lineId, normalizedCaseId, ts);
  const sig = signV2(payload);
  const p = b64urlFromString(payload);

  // 1) ユーザーが戻ってくるURL（自サイト）
  const origin = getOriginSafe();
  const redirect = safeURL(opts?.redirectUrl ?? '/', origin);
  redirect.searchParams.set('lineId', lineId);
  redirect.searchParams.set('caseId', normalizedCaseId);
  redirect.searchParams.set('ts', ts);
  redirect.searchParams.set('sig', sig);
  redirect.searchParams.set('p', p);
  if (opts?.formId) redirect.searchParams.set('formId', opts.formId);
  if (opts?.extra) for (const [k, v] of Object.entries(opts.extra)) redirect.searchParams.set(k, v);

  // 2) FormMailer 側のURLに、redirect_url を渡す
  const url = safeURL(baseUrl, origin);
  const lineIdQueryKeys = opts.lineIdQueryKeys && opts.lineIdQueryKeys.length > 0 ? opts.lineIdQueryKeys : ['line_id'];
  const caseIdQueryKeys = opts.caseIdQueryKeys && opts.caseIdQueryKeys.length > 0 ? opts.caseIdQueryKeys : ['case_id'];
  lineIdQueryKeys.forEach((key) => {
    if (key) url.searchParams.set(key, lineId);
  });
  caseIdQueryKeys.forEach((key) => {
    if (key) url.searchParams.set(key, normalizedCaseId);
  });
  url.searchParams.set('redirect_url', redirect.toString());
  url.searchParams.set('referrer', origin ?? '');
  return url.toString();
}

/**
 * intake 用。caseId は空文字で署名。
 * intakeBase: intake の FormMailer URL
 * intakeRedirect: intake 完了後に戻す /done のURL
 */
export function makeIntakeUrl(
  intakeBase: string,
  intakeRedirect: string,
  lineId: string,
  opts: {
    formId?: string;
    lineIdQueryKeys?: string[];
  } = {}
): string {
  const ts = tsSec(); // UNIX秒
  const normalizedCaseId = '';
  const payload = makePayload(lineId || '', normalizedCaseId, ts);
  const sig = signV2(payload);
  const p = b64urlFromString(payload);

  const origin = getOriginSafe();
  const redirect = safeURL(intakeRedirect, origin);
  redirect.searchParams.set('lineId', lineId || '');
  redirect.searchParams.set('caseId', normalizedCaseId);
  redirect.searchParams.set('ts', ts);
  redirect.searchParams.set('sig', sig);
  redirect.searchParams.set('p', p);
  if (opts.formId) redirect.searchParams.set('formId', opts.formId);

  const url = safeURL(intakeBase, origin);
  const lineIdQueryKeys =
    opts.lineIdQueryKeys && opts.lineIdQueryKeys.length > 0 ? opts.lineIdQueryKeys : ['line_id'];
  lineIdQueryKeys.forEach((key) => {
    if (key && lineId) url.searchParams.set(key, lineId);
  });
  url.searchParams.set('redirect_url', redirect.toString());
  url.searchParams.set('referrer', origin ?? '');
  return url.toString();
}

// 新: intake 用の redirect_url を ts=秒固定で生成（lineId/caseId を明示指定可能）
export function buildIntakeRedirectUrl(
  intakeBase: string,
  intakeRedirect: string,
  lineId: string,
  caseId: string
): string {
  const ts = tsSec();
  const normalizedCaseId = normalizeCaseId(caseId || '');
  const payload = makePayload(lineId || '', normalizedCaseId, ts);
  const sig = signV2(payload);
  const p = b64urlFromString(payload);

  const origin = getOriginSafe();
  const redirect = safeURL(intakeRedirect, origin);
  redirect.searchParams.set('lineId', lineId || '');
  redirect.searchParams.set('caseId', normalizedCaseId);
  redirect.searchParams.set('ts', ts); // 秒
  redirect.searchParams.set('sig', sig);
  redirect.searchParams.set('p', p);

  const url = safeURL(intakeBase, origin);
  url.searchParams.set('redirect_url', redirect.toString());
  url.searchParams.set('referrer', origin ?? '');
  return url.toString();
}

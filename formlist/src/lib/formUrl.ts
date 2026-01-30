// formlist/src/lib/formUrl.ts
// 役割: GAS と連携する署名付き FormMailer URL を生成し、リダイレクトやプレフィルを制御する。
// 注意: PII をクエリに出さない方針を維持すること、allowlist 以外の項目を追加しないこと。
import 'server-only';
import crypto from 'crypto';

const getSecret = () => {
  const secret = process.env.BOOTSTRAP_SECRET ?? process.env.TOKEN_SECRET;
  if (!secret) {
    throw new Error('Missing BOOTSTRAP_SECRET/TOKEN_SECRET');
  }
  return secret;
};

const PREFILL_ALLOWLIST: Record<string, string[]> = {
  s2002_userform: ['case_id', 'seq', 'email'],
  s2005_creditors: ['case_id', 'seq', 'email'],
};

const b64url = (s: string) => s.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
const b64urlFromString = (s: string) => b64url(Buffer.from(s, 'utf8').toString('base64'));
const makePayload = (lineId: string, caseId: string, ts: string) => `${lineId}|${caseId}|${ts}`;
const signV2 = (payload: string) =>
  b64url(crypto.createHmac('sha256', getSecret()).update(payload, 'utf8').digest('base64'));
const normalizeCaseId = (raw: string) => {
  if (!raw) return '';
  const onlyDigits = raw.replace(/\D/g, '');
  if (!onlyDigits) return '';
  return onlyDigits.padStart(4, '0');
};
const normalizeUserKey = (lineId: string) => {
  if (!lineId) return '';
  if (/^staff\d{2}$/i.test(lineId)) return lineId.toLowerCase();
  const cleaned = lineId.toLowerCase().replace(/[^a-z0-9]/g, '');
  return cleaned.slice(0, 6);
};
const makeCaseKey = (lineId: string, caseId: string) => {
  const userKey = normalizeUserKey(lineId);
  const normalizedCaseId = normalizeCaseId(caseId);
  if (userKey.length !== 6 || !normalizedCaseId) return '';
  return `${userKey}-${normalizedCaseId}`;
};
const toIdx0 = (name?: string) => (name ? (/\[\d+\]$/.test(name) ? name : `${name}[0]`) : undefined);
const assertNoLegacyParams = (u: URL) => {
  const bad = ['case_id', 'line_id', 'redirect_url', 'referrer'].filter((k) => u.searchParams.has(k));
  if (bad.length) throw new Error(`FormMailer baseUrl must not contain plain params: ${bad.join(', ')}`);
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

type PrefillValue = string | number | boolean | null | undefined;

type MakeFormUrlOptions = {
  redirectUrl?: string; // 例: https://app.example/done?formId=309542
  formId?: string; // 表示用（署名には使わない）
  extra?: Record<string, string>;
  formKey?: string;
  prefill?: boolean | Record<string, PrefillValue>;
  extraPrefill?: Record<string, PrefillValue>;
  lineIdQueryKeys?: string[]; // フォームが参照できるクエリ名（デフォルト: line_id）
  caseIdQueryKeys?: string[]; // フォームが参照できるクエリ名（デフォルト: case_id）
  caseKeyQueryKeys?: string[]; // フォームが参照できるクエリ名（例: case_key）
  redirectUrlQueryKey?: string; // FormMailer 側の redirect_url 指定キー（デフォルト: redirect_url[0]）
  referrerQueryKey?: string; // FormMailer 側の referrer 指定キー（デフォルト: referrer[0]）
};

export function prefillParamsBuilder(
  formKey: string | undefined,
  params?: Record<string, PrefillValue>
): Record<string, string> {
  if (!formKey || !params) return {};
  const allow = PREFILL_ALLOWLIST[formKey];
  if (!allow || allow.length === 0) return {};
  const out: Record<string, string> = {};
  for (const key of allow) {
    if (!Object.prototype.hasOwnProperty.call(params, key)) continue;
    const raw = params[key];
    if (raw === null || raw === undefined) continue;
    const value = String(raw).trim();
    if (value) out[key] = value;
  }
  return out;
}

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
  redirect.searchParams.set('caseId', normalizedCaseId);
  redirect.searchParams.set('ts', ts);
  redirect.searchParams.set('sig', sig);
  redirect.searchParams.set('p', p);
  if (opts?.formId) redirect.searchParams.set('formId', opts.formId);
  if (opts?.extra) for (const [k, v] of Object.entries(opts.extra)) redirect.searchParams.set(k, v);
  redirect.searchParams.set('bust', ts);

  // 2) FormMailer 側のURLに、redirect_url を渡す
  const url = safeURL(baseUrl, origin);
  assertNoLegacyParams(url);
  const lineIdKeys = (opts.lineIdQueryKeys ?? ['line_id']).map(toIdx0).filter((k): k is string => Boolean(k));
  const caseIdKeys = (opts.caseIdQueryKeys ?? ['case_id']).map(toIdx0).filter((k): k is string => Boolean(k));
  const caseKeyKeys = (opts.caseKeyQueryKeys ?? []).map(toIdx0).filter((k): k is string => Boolean(k));
  const redirectKey = toIdx0(opts.redirectUrlQueryKey ?? 'redirect_url');
  const referrerKey = toIdx0(opts.referrerQueryKey ?? 'referrer');
  lineIdKeys.forEach((key) => {
    if (lineId) url.searchParams.set(key, lineId);
  });
  caseIdKeys.forEach((key) => {
    if (normalizedCaseId) url.searchParams.set(key, normalizedCaseId);
  });
  caseKeyKeys.forEach((key) => {
    const caseKey = makeCaseKey(lineId, normalizedCaseId);
    if (caseKey) url.searchParams.set(key, caseKey);
  });
  let prefillSource: Record<string, PrefillValue> | undefined;
  if (opts.prefill) {
    prefillSource = {};
    if (opts.prefill === true) {
      if (normalizedCaseId) prefillSource['case_id'] = normalizedCaseId;
    } else if (typeof opts.prefill === 'object') {
      prefillSource = { ...opts.prefill };
    }
    if (opts.extraPrefill) {
      prefillSource = { ...(prefillSource ?? {}), ...opts.extraPrefill };
    }
  } else if (opts.extraPrefill) {
    prefillSource = { ...opts.extraPrefill };
  }
  const allowedPrefill = prefillParamsBuilder(opts.formKey, prefillSource);
  Object.entries(allowedPrefill).forEach(([key, value]) => {
    url.searchParams.set(key, value);
  });
  if (allowedPrefill.email) {
    url.searchParams.set('メールアドレス[0]', String(allowedPrefill.email));
  }
  if (redirectKey) url.searchParams.set(redirectKey, redirect.toString());
  if (referrerKey) url.searchParams.set(referrerKey, origin ?? '');
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
    redirectUrlQueryKey?: string;
    referrerQueryKey?: string;
  } = {}
): string {
  const ts = tsSec(); // UNIX秒
  const normalizedCaseId = '';
  const payload = makePayload(lineId || '', normalizedCaseId, ts);
  const sig = signV2(payload);
  const p = b64urlFromString(payload);

  const origin = getOriginSafe();
  const redirect = safeURL(intakeRedirect, origin);
  redirect.searchParams.set('caseId', normalizedCaseId);
  redirect.searchParams.set('ts', ts);
  redirect.searchParams.set('sig', sig);
  redirect.searchParams.set('p', p);
  if (opts.formId) redirect.searchParams.set('formId', opts.formId);

  const url = safeURL(intakeBase, origin);
  assertNoLegacyParams(url);
  const lineIdKeys = (opts.lineIdQueryKeys ?? ['line_id']).map(toIdx0).filter((k): k is string => Boolean(k));
  const redirectKey = toIdx0(opts.redirectUrlQueryKey ?? 'redirect_url');
  const referrerKey = toIdx0(opts.referrerQueryKey ?? 'referrer');
  lineIdKeys.forEach((key) => {
    if (lineId) url.searchParams.set(key, lineId);
  });
  if (redirectKey) url.searchParams.set(redirectKey, redirect.toString());
  if (referrerKey) url.searchParams.set(referrerKey, origin ?? '');
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
  redirect.searchParams.set('caseId', normalizedCaseId);
  redirect.searchParams.set('ts', ts); // 秒
  redirect.searchParams.set('sig', sig);
  redirect.searchParams.set('p', p);

  const url = safeURL(intakeBase, origin);
  url.searchParams.set('redirect_url[0]', redirect.toString());
  url.searchParams.set('referrer[0]', origin ?? '');
  return url.toString();
}

import crypto from 'node:crypto';

export const tsSec = () => Math.floor(Date.now() / 1000).toString();

export const base64url = (b64: string) => b64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');

export const payloadV2 = (lineId: string, caseId?: string, ts?: string) => [lineId, caseId ?? '', ts ?? tsSec()].join('|');

export const signV2 = (payload: string) =>
  base64url(
    crypto.createHmac('sha256', process.env.BAS_API_HMAC_SECRET || process.env.BOOTSTRAP_SECRET || '').update(payload, 'utf8').digest('base64')
  );

export function b64UrlFromString(s: string) {
  if (typeof Buffer !== 'undefined') {
    return base64url(Buffer.from(s, 'utf8').toString('base64'));
  }
  const enc = new TextEncoder().encode(s);
  const b64 = typeof btoa !== 'undefined' ? btoa(String.fromCharCode(...(enc as Uint8Array))) : '';
  return base64url(b64);
}


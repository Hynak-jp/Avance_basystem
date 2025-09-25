import crypto from 'crypto';

type BuildSignedQueryParams = {
  lineId?: string | null;
  caseId?: string | null;
  action?: string;
};

export function buildSignedQuery(params: BuildSignedQueryParams = {}): string {
  const ts = Date.now().toString();
  const lineId = params.lineId ?? '';
  const caseId = params.caseId ?? '';
  const action = params.action ?? 'status';
  const secret = process.env.BAS_API_HMAC_SECRET;
  if (!secret) {
    throw new Error('BAS_API_HMAC_SECRET is not configured');
  }
  const message = `${ts}.${lineId}.${caseId}`;
  const sig = crypto.createHmac('sha256', secret).update(message).digest('hex');
  const searchParams = new URLSearchParams({ action, ts, sig });
  if (lineId) searchParams.set('lineId', lineId);
  if (caseId) searchParams.set('caseId', caseId);
  return searchParams.toString();
}

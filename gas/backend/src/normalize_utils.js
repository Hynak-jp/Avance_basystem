/** case_id を常に "0001" 形式にする */
function normalizeCaseId_(value) {
  const digits = String(value == null ? '' : value).replace(/\D/g, '');
  if (!digits) return '';
  return digits.slice(-4).padStart(4, '0');
}

/** user_key は6桁小文字 */
function normalizeUserKey_(value) {
  const normalized = String(value == null ? '' : value).trim().toLowerCase();
  if (!normalized) return '';
  return normalized;
}

// 共通正規化・一致判定ヘルパ（GAS グローバル）

/** '1' と '0001' を等価扱い（4桁ゼロ埋め） */
function normCaseId_(s) {
  s = String(s || '').trim();
  var n = s.replace(/^0+/, '');
  if (!n) return '';
  var num = parseInt(n, 10);
  if (!isFinite(num)) return '';
  return Utilities.formatString('%04d', num);
}

/** 'UC13DF-1' → 'uc13df-0001' のように正規化 */
function normCaseKey_(s) {
  s = String(s || '').trim().toLowerCase();
  var m = s.match(/^([a-z0-9]{2,})-(\d{1,})$/);
  if (!m) return s;
  return m[1] + '-' + normCaseId_(m[2]);
}

function normLineId_(s) {
  return String(s || '').trim();
}

/** intake__*.json 名かどうか */
function isIntakeJsonName_(name) {
  return /^intake__\d+\.json$/i.test(String(name || ''));
}

/**
 * 一致判定（優先度: case_key → case_id → line_id）
 * fileMeta: { case_key, case_id, line_id }
 * known: { case_key, case_id, line_id }
 */
function matchMetaToCase_(fileMeta, known) {
  var fm = fileMeta || {};
  var fk = normCaseKey_(fm.case_key || fm.caseKey || '');
  var fid = normCaseId_(fm.case_id || fm.caseId || '');
  var fl = normLineId_(fm.line_id || fm.lineId || '');

  var kk = normCaseKey_(known && known.case_key);
  var kid = normCaseId_(known && known.case_id);
  var kl = normLineId_(known && known.line_id);

  if (fk && kk && fk === kk) return { ok: true, by: 'case_key' };
  if (fid && kid && fid === kid) return { ok: true, by: 'case_id' };
  if (fl && kl && fl === kl) return { ok: true, by: 'line_id' };
  return { ok: false, by: '' };
}


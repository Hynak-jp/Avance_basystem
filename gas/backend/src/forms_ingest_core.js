/** ================== forms_ingest_core.js ==================
 * 汎用フォーム保存パイプライン
 * - メール本文の解析（parseFormMail_ に委譲）
 * - submission_id 正規化（normalizeSubmissionIdStrict_ があれば優先）
 * - case_id → case_key フォルダ解決（resolveCaseByCaseId_ / ensureCaseFolderId_）
 * - <form_key>__<submission_id>.json をケース直下に保存（saveSubmissionJson_）
 * 依存関数（既存プロジェクト内にある前提）:
 *   parseFormMail_, resolveCaseByCaseId_, ensureCaseFolderId_, saveSubmissionJson_, updateCasesRow_
 * --------------------------------------------------------- */

/** mapper レジストリ（多重定義でも上書きしないガード） */
var FORM_MAPPERS =
  this.FORM_MAPPERS && typeof this.FORM_MAPPERS === 'object' ? this.FORM_MAPPERS : {};
this.FORM_MAPPERS = FORM_MAPPERS;

var FORM_INGEST_SECRET = (function () {
  try {
    if (
      typeof PropertiesService !== 'undefined' &&
      PropertiesService.getScriptProperties &&
      PropertiesService.getScriptProperties().getProperty
    ) {
      var v = PropertiesService.getScriptProperties().getProperty('NOTIFY_SECRET');
      if (v) return v;
    }
  } catch (_) {}
  return 'FM-BAS';
})();

/** 各フォームの「FIELDS → model」変換関数を登録 */
function registerFormMapper(formKey, mapperFn) {
  if (!formKey || typeof mapperFn !== 'function') return;
  FORM_MAPPERS[String(formKey).trim()] = mapperFn;
}

/** 既存の normalizeSubmissionIdStrict_ があれば利用、なければフォールバック */
function ensureSubmissionIdDigits_(sid, tsKey, meta) {
  try {
    if (typeof normalizeSubmissionIdStrict_ === 'function') {
      var normalized = normalizeSubmissionIdStrict_(sid, tsKey, meta);
      if (normalized) return normalized;
    }
  } catch (_) {}
  var digits = String(sid || '').replace(/\D/g, '');
  if (digits.length >= 3) return digits;
  var fromMeta = '';
  if (meta && typeof meta === 'object') {
    fromMeta = String(meta.submitted_at || meta.submittedAt || '').trim();
  }
  var fromMetaDigits = fromMeta.replace(/[^\d]/g, '').slice(0, 14);
  if (fromMetaDigits.length >= 8) return fromMetaDigits;
  var tz =
    (typeof Session !== 'undefined' &&
      Session.getScriptTimeZone &&
      Session.getScriptTimeZone()) ||
    'Asia/Tokyo';
  return Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmmss');
}

/** subject から form_key を推測（METAが無い時の保険） */
function guessFormKeyFromSubject_(subject) {
  var s = String(subject || '').toLowerCase();
  var match = s.match(/s\d{4}/);
  if (match) {
    var prefix = match[0];
    var keys = Object.keys(FORM_MAPPERS || {});
    for (var i = 0; i < keys.length; i++) {
      var key = String(keys[i]).toLowerCase();
      if (key.indexOf(prefix) === 0) return keys[i];
    }
  }
  if (/s2006/.test(s)) return 's2006_creditors_public';
  if (/s2002/.test(s)) return 's2002_userform';
  // TODO: 新フォームはここに追記
  return 'unknown_form';
}

/** フィールド配列→素直なオブジェクト（汎用フォールバック） */
function mapFieldsGeneric_(fields) {
  var obj = {};
  (fields || []).forEach(function (row, idx) {
    var key = row && row.label ? String(row.label).trim() : 'field_' + idx;
    obj[key] = (row && row.value) || '';
  });
  return { fields: obj };
}

/**
 * 【中核】メール（件名・本文）を解析→cases台帳で case_key フォルダを解決→
 * ケース直下へ <form_key>__<submission_id>.json 保存。
 * opts: { form_key?: string, mapper?: function, case_id?: string }
 */
function ingestFormMailToCase_(subject, body, opts) {
  opts = opts || {};
  var lock = null;
  var lockAcquired = false;
  try {
    if (typeof LockService !== 'undefined' && LockService.getScriptLock) {
      lock = LockService.getScriptLock();
      if (lock && lock.tryLock(5000)) {
        lockAcquired = true;
      } else if (lock) {
        throw new Error('ingestFormMailToCase_: lock timeout');
      }
    }
  } catch (lockErr) {
    throw new Error('ingestFormMailToCase_: failed to acquire lock :: ' + (lockErr && lockErr.message ? lockErr.message : lockErr));
  }

  try {
    var parsed = parseFormMail_(subject, body);
    parsed.meta = parsed.meta || {};

    var expectedSecret = String(FORM_INGEST_SECRET || '').trim().toLowerCase();
    var gotSecret = String((parsed.meta && parsed.meta.secret) || '').trim().toLowerCase();
    if (!expectedSecret || !gotSecret || gotSecret !== expectedSecret) {
      throw new Error('Invalid secret');
    }

    parsed.meta.form_key = String(
      opts.form_key || parsed.meta.form_key || guessFormKeyFromSubject_(subject)
    ).trim();

    parsed.meta.submission_id = ensureSubmissionIdDigits_(
      parsed.meta.submission_id,
      parsed.meta.tsKey,
      parsed.meta
    );

    var caseId = String(parsed.meta.case_id || opts.case_id || '').trim();
    if (!caseId) throw new Error('META.case_id is required');

    var mapper = opts.mapper || FORM_MAPPERS[parsed.meta.form_key] || mapFieldsGeneric_;
    parsed.model = mapper(parsed.fieldsRaw || [], parsed.meta);

    var caseInfo = resolveCaseByCaseId_(caseId);
    if (!caseInfo) throw new Error('Unknown case_id: ' + caseId);
    parsed.meta.case_key =
      (caseInfo && caseInfo.caseKey) ||
      parsed.meta.case_key ||
      parsed.meta.caseKey ||
      '';
    var folderId = ensureCaseFolderId_(caseInfo);
    var file = saveSubmissionJson_(folderId, parsed);

    try {
      updateCasesRow_(caseId, {
        last_activity: new Date(),
        last_form_key: parsed.meta.form_key,
      });
    } catch (_) {}

    try {
      Logger.log(
        '[INGEST] %s → %s under case_key=%s',
        parsed.meta.form_key,
        file.getName(),
        caseInfo.caseKey || '(unknown)'
      );
    } catch (_) {}

    return {
      fileId: file.getId(),
      name: file.getName(),
      caseKey: caseInfo.caseKey,
      form_key: parsed.meta.form_key,
    };
  } finally {
    if (lock && lockAcquired) {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  }
}

/**
 * フォルダの「それらしさ」を採点。
 * s2002_userform__*.json: +5（≥1件で加点）
 * s2006_creditors_public__*.json: +3（≥1件で加点）
 * 任意の直下 .json: +2（≥1件で加点）
 * drafts/ サブフォルダ: +3
 * attachments/ or staff_inputs/ サブフォルダ: +1 ずつ
 */
function scoreCaseFolder_(folder) {
  var s2 = 0;
  var s6 = 0;
  var anyJson = 0;
  if (!folder) return { score: 0, s2: 0, s6: 0, anyJson: 0, drafts: 0, attach: 0, staff: 0 };
  var itF = folder.getFiles();
  var scanned = 0;
  while (itF.hasNext() && scanned < 500) {
    var f = itF.next();
    scanned++;
    var name = f.getName && f.getName();
    if (!name) continue;
    if (/\.json$/i.test(name)) anyJson = 1;
    if (/^s2002_userform__/i.test(String(name).trim())) s2 = 1;
    if (/^s2006_creditors_public__/i.test(String(name).trim())) s6 = 1;
  }
  var drafts = 0;
  var attach = 0;
  var staff = 0;
  var itD = folder.getFolders();
  while (itD.hasNext()) {
    var d = itD.next();
    var dn = d.getName && d.getName();
    if (!dn) continue;
    dn = String(dn).trim().toLowerCase();
    if (dn === 'drafts') drafts = 1;
    if (dn === 'attachments') attach = 1;
    if (dn === 'staff_inputs') staff = 1;
  }
  var score = s2 * 5 + s6 * 3 + anyJson * 2 + drafts * 3 + attach * 1 + staff * 1;
  return { score: score, s2: s2, s6: s6, anyJson: anyJson, drafts: drafts, attach: attach, staff: staff };
}

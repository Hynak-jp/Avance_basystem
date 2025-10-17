/**
 * status_api.js
 *  - WebApp doGet ルート（フォーム受領状況・再開操作）
 *  - 署名方式（互換）
 *    V1: sig = HEX(HMAC_SHA256(`${ts}.${lineId}.${caseId}`, SECRET))
 *    V2: p = base64url(`${lineId}|${caseId}|${ts}`), sig = base64url(HMAC_SHA256(payload, SECRET))
 */

const STATUS_API_PROPS = PropertiesService.getScriptProperties();
var STATUS_API_CASES_CASE_ID_FORMATTED = false;
var STATUS_API_SUBMISSIONS_CASE_ID_FORMATTED = false;
const STATUS_API_SECRET =
  STATUS_API_PROPS.getProperty('HMAC_SECRET') ||
  STATUS_API_PROPS.getProperty('BAS_API_HMAC_SECRET') ||
  '';
const STATUS_API_NONCE_WINDOW_SECONDS = 600; // 10 分

// === case_id 正規化（"0001" など4桁固定、常に文字列） ===
function statusApi_normCaseId_(v) {
  return normalizeCaseId_(v);
}

function statusApi_formsFromSubmissions_(caseId) {
  try {
    if (!caseId || typeof bs_getSheet_ !== 'function') return [];
    var sheetName = typeof SHEETS_REPO_SUBMISSIONS !== 'undefined' && SHEETS_REPO_SUBMISSIONS
      ? SHEETS_REPO_SUBMISSIONS
      : 'submissions';
    var sh = bs_getSheet_(sheetName);
    if (!sh) return [];
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return [];
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (v) {
      return String(v || '').trim();
    });
    var idx = typeof bs_toIndexMap_ === 'function' ? bs_toIndexMap_(headers) : {};
    var ci = idx['case_id'] != null ? idx['case_id'] : idx['caseId'];
    var fi = idx['form_key'] != null ? idx['form_key'] : idx['formKey'];
    if (!(ci >= 0 && fi >= 0)) return [];
    var seqIdx = idx['seq'] != null ? idx['seq'] : idx['Seq'];
    var caseKeyIdx = idx['case_key'] != null ? idx['case_key'] : idx['caseKey'];
    var statusIdx = idx['status'] != null ? idx['status'] : idx['Status'];
    var reopenAtIdx = idx['reopened_at'] != null ? idx['reopened_at'] : idx['reopenedAt'];
    var reopenUntilIdx = idx['reopen_until'] != null ? idx['reopen_until'] : idx['reopenUntil'];
    var lockedReasonIdx = idx['locked_reason'] != null ? idx['locked_reason'] : idx['lockedReason'];
    var reopenedByIdx = idx['reopened_by'] != null ? idx['reopened_by'] : idx['reopenedBy'];
    var canEditIdx = idx['can_edit'] != null ? idx['can_edit'] : idx['canEdit'];
    var reopenAtEpochIdx = idx['reopened_at_epoch'] != null ? idx['reopened_at_epoch'] : idx['reopenedAtEpoch'];
    var reopenUntilEpochIdx = idx['reopen_until_epoch'] != null ? idx['reopen_until_epoch'] : idx['reopenUntilEpoch'];
    var rows = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var aliasMap = {
      s2002: 's2002_userform',
      s2002_form: 's2002_userform',
      s2002_userform: 's2002_userform',
    };
    var want = statusApi_normCaseId_(caseId);
    var map = Object.create(null);
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var rowCase = statusApi_normCaseId_(row[ci]);
      if (!rowCase || rowCase !== want) continue;
      var rawFormKey = String(row[fi] || '').trim();
      if (!rawFormKey) continue;
      var normalizedKey = rawFormKey.toLowerCase().replace(/[\s-]+/g, '_');
      var canonicalKey = aliasMap[normalizedKey] || normalizedKey;
      if (!canonicalKey) continue;
      var formKey = canonicalKey;
      var seq = seqIdx != null ? Number(row[seqIdx] || 0) || 0 : 0;
      var status = statusIdx != null ? String(row[statusIdx] || '').trim() : '';
      if (!status || status.toLowerCase() === formKey.toLowerCase()) status = 'submitted';
      else if (status.toLowerCase() === 'received') status = 'submitted';
      var reopenedAtEpoch = null;
      if (reopenAtEpochIdx != null && reopenAtEpochIdx >= 0) {
        var rawEpoch = Number(row[reopenAtEpochIdx]);
        if (Number.isFinite(rawEpoch) && rawEpoch > 0) reopenedAtEpoch = Math.floor(rawEpoch);
      }
      var reopenUntilEpoch = null;
      if (reopenUntilEpochIdx != null && reopenUntilEpochIdx >= 0) {
        var rawUntilEpoch = Number(row[reopenUntilEpochIdx]);
        if (Number.isFinite(rawUntilEpoch) && rawUntilEpoch > 0) reopenUntilEpoch = Math.floor(rawUntilEpoch);
      }
      var reopenedAt = reopenAtIdx != null ? String(row[reopenAtIdx] || '').trim() : '';
      var reopenUntil = reopenUntilIdx != null ? String(row[reopenUntilIdx] || '').trim() : '';
      var lockedReason = lockedReasonIdx != null ? String(row[lockedReasonIdx] || '').trim() : '';
      var reopenedBy = reopenedByIdx != null ? String(row[reopenedByIdx] || '').trim() : '';
      var canEdit = canEditIdx != null ? statusApi_normalizeBool_(row[canEditIdx]) : status.toLowerCase() === 'reopened';
      var caseKeyValue = caseKeyIdx != null ? String(row[caseKeyIdx] || '').trim() : '';
      if (reopenedAtEpoch == null && reopenedAt) {
        var parsedReopened = Date.parse(reopenedAt);
        if (Number.isFinite(parsedReopened)) reopenedAtEpoch = Math.floor(parsedReopened / 1000);
      }
      if (reopenUntilEpoch == null && reopenUntil) {
        var parsedReopenUntil = Date.parse(reopenUntil);
        if (Number.isFinite(parsedReopenUntil)) reopenUntilEpoch = Math.floor(parsedReopenUntil / 1000);
      }
      var existing = map[formKey];
      if (!existing || seq >= existing._seq) {
        map[formKey] = {
          case_id: want,
          form_key: formKey,
          case_key: caseKeyValue,
          status: status,
          can_edit: canEdit,
          reopened_at: reopenedAt || null,
          reopen_until: reopenUntil || null,
          locked_reason: lockedReason || null,
          reopened_by: reopenedBy || null,
          reopened_at_epoch: reopenedAtEpoch,
          reopen_until_epoch: reopenUntilEpoch,
          last_seq: seq,
          _seq: seq,
        };
      }
    }
    var keys = Object.keys(map);
    if (typeof getLastSeq_ === 'function') {
      keys.forEach(function (key) {
        var item = map[key];
        if (!item) return;
        var status = String(item.status || '').trim().toLowerCase();
        if (!(item.last_seq > 0) && (!status || status === 'received')) {
          try { item.last_seq = Number(getLastSeq_(item.case_id, key)) || 0; } catch (_) {}
        }
      });
    }
    return keys.map(function (key) {
      var item = map[key];
      if (!item) {
        return {
          case_id: statusApi_normCaseId_(caseId),
          form_key: key,
          case_key: '',
          status: '',
          can_edit: false,
          reopened_at: null,
          reopen_until: null,
          locked_reason: null,
          reopened_by: null,
          reopened_at_epoch: null,
          reopen_until_epoch: null,
          last_seq: 0,
        };
      }
      delete item._seq;
      return item;
    });
  } catch (_) {
    return [];
  }
}

// contacts.active_case_id を "0001" 文字列で強制保存（数値化防止）
function statusApi_ensureActiveCaseIdString_(lineId, caseId) {
  try {
    var sid =
      PropertiesService.getScriptProperties().getProperty('BAS_MASTER_SPREADSHEET_ID') ||
      PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!sid) return;
    var ss = SpreadsheetApp.openById(sid);
    var sh = ss.getSheetByName('contacts');
    if (!sh) return;

    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    var idx = {};
    headers.forEach(function (h, i) {
      idx[h] = i;
      idx[h.replace(/_([a-z])/g, (_, c) => c.toUpperCase())] = i;
    });

    var data = sh.getDataRange().getValues(); // 1行目ヘッダ込み
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var lid = String(row[idx.line_id != null ? idx.line_id : idx.lineId] || '').trim();
      if (lid && lineId && lid === lineId) {
        var aci = idx.active_case_id != null ? idx.active_case_id : idx.activeCaseId;
        if (aci != null) {
          var cur = String(row[aci] == null ? '' : row[aci]).trim();
          if (cur !== caseId) {
            // プレーンテキストで書く
            sh.getRange(r + 1, aci + 1)
              .setNumberFormat('@')
              .setValue(caseId);
          }
        }
        break;
      }
    }
  } catch (_) {}
}

function statusApi_hex_(bytes) {
  return bytes
    .map(function (b) {
      const v = b & 0xff;
      return (v < 16 ? '0' : '') + v.toString(16);
    })
    .join('');
}

function statusApi_hmac_(message, secret) {
  const raw = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, message, secret);
  return statusApi_hex_(raw).toLowerCase();
}

function statusApi_getSecret_() {
  var s =
    (STATUS_API_PROPS.getProperty('HMAC_SECRET') || STATUS_API_PROPS.getProperty('BAS_API_HMAC_SECRET') || '')
      .replace(/[\r\n\s]+$/g, '');
  return s;
}

function statusApi_verify_(params) {
  var secret = statusApi_getSecret_();
  if (!secret) throw new Error('hmac_secret_not_configured');
  const tsRaw = String(params.ts || '').trim();
  const sig = String(params.sig || '').trim();
  const lineId = String(params.lineId || '').trim();
  const caseId = String(params.caseId || params.case_id || '').trim();
  if (!tsRaw || !sig) throw new Error('missing ts or sig');
  const tsNum = Number(tsRaw);
  if (!Number.isFinite(tsNum)) throw new Error('invalid timestamp');
  const skewMs = Math.abs(Date.now() - tsNum * 1000); // ts は秒
  if (skewMs > 10 * 60 * 1000) throw new Error('timestamp too far');
  const message = tsRaw + '.' + lineId + '.' + caseId;
  const expected = statusApi_hmac_(message, secret);
  if (expected !== sig.toLowerCase()) {
    throw new Error('invalid signature');
  }
}

function statusApi_assertNonce_(lineId, tsRaw, sig) {
  if (!sig) return;
  const cache = CacheService.getScriptCache();
  const key = ['status_nonce', lineId || '', tsRaw || '', sig || ''].join(':');
  if (cache.get(key)) {
    throw new Error('nonce_reused');
  }
  cache.put(key, '1', STATUS_API_NONCE_WINDOW_SECONDS);
}

function statusApi_jsonOut_(obj, status) {
  const out = ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
  if (status && typeof out.setStatusCode === 'function') {
    out.setStatusCode(status);
  }
  return out;
}

/** ===== meta 補完ユーティリティ（staging 作成前に必ず通す） ===== **/
function contacts_lookupByEmail_(email) {
  try {
    email = String(email || '').trim().toLowerCase();
    if (!email) return null;
    var sid =
      PropertiesService.getScriptProperties().getProperty('BAS_MASTER_SPREADSHEET_ID') ||
      PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!sid) return null;
    var ss = SpreadsheetApp.openById(sid);
    var sh = ss.getSheetByName('contacts');
    if (!sh) return null;
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var idx = {};
    headers.forEach(function (h, i) { idx[String(h).trim()] = i; });

    var rc = sh.getLastRow() - 1;
    if (rc <= 0) return null;

    var rows = sh.getRange(2, 1, rc, headers.length).getValues();
    var emailIdx = idx['email'];
    var lineIdIdx = idx['line_id'] != null ? idx['line_id'] : idx['lineId'];
    var userKeyIdx = idx['user_key'] != null ? idx['user_key'] : idx['userKey'];
    var activeCaseIdx = idx['active_case_id'] != null ? idx['active_case_id'] : idx['activeCaseId'];

    if (emailIdx == null) return null;

    for (var i = rows.length - 1; i >= 0; i--) {
      var e = String(rows[i][emailIdx] || '').trim().toLowerCase();
      if (e && e === email) {
        return {
          line_id: lineIdIdx != null ? String(rows[i][lineIdIdx] || '').trim() : '',
          user_key: userKeyIdx != null ? String(rows[i][userKeyIdx] || '').trim() : '',
          active_case_id: activeCaseIdx != null ? String(rows[i][activeCaseIdx] || '').trim() : '',
        };
      }
    }
    return null;
  } catch (err) {
    try { Logger.log('[contacts_lookupByEmail_] ' + err); } catch(_) {}
    return null;
  }
}

function intake_fillMeta_(obj, hint) {
  obj = obj && typeof obj === 'object' ? obj : {};
  obj.meta = obj.meta || {};

  var hintLineId = String((hint && hint.line_id) || '').trim();
  var hintCaseId = String((hint && hint.case_id) || '').trim();

  var lineId = String(obj.meta.line_id || obj.meta.lineId || '').trim();
  if (!lineId && hintLineId) lineId = hintLineId;

  var cinfo = null;
  if (!lineId || !obj.meta.user_key || !hintCaseId) {
    var email = String(
      (obj.fields && (obj.fields.email || obj.fields['メールアドレス'])) ||
        (obj.model && (obj.model.email || obj.model['email'])) ||
        ''
    ).trim();
    if (email) cinfo = contacts_lookupByEmail_(email);
    // 任意: 後段の救済でも使えるよう、email を保持
    try {
      if (email) {
        obj.model = obj.model || {};
        if (!obj.model.email) obj.model.email = email;
      }
    } catch (_) {}
  }

  var userKey = String(obj.meta.user_key || obj.meta.userKey || '').trim();
  if (!userKey && lineId && typeof drive_userKeyFromLineId_ === 'function') {
    userKey = drive_userKeyFromLineId_(lineId);
  }
  if (!userKey && cinfo && cinfo.user_key) userKey = cinfo.user_key;

  var caseId = String(obj.meta.case_id || obj.meta.caseId || '').trim();
  if (!caseId && hintCaseId) caseId = hintCaseId;
  if (!caseId && cinfo && cinfo.active_case_id) caseId = cinfo.active_case_id;
  caseId = caseId ? String(caseId).replace(/\D/g, '').padStart(4, '0') : '';

  if (lineId && !obj.meta.line_id) obj.meta.line_id = lineId;
  if (userKey && !obj.meta.user_key) obj.meta.user_key = userKey;
  if (caseId && !obj.meta.case_id) obj.meta.case_id = caseId;
  if (!obj.meta.case_key && userKey && caseId) obj.meta.case_key = userKey + '-' + caseId;

  return obj;
}

function statusApi_normalizeBool_(v) {
  if (typeof v === 'boolean') return v;
  const s = String(v || '')
    .trim()
    .toLowerCase();
  return s === 'true' || s === '1' || s === 'yes';
}

// ===== 互換ユーティリティ（snake/camel 両対応） =====
function statusApi_toCamel_(s) {
  return String(s || '').replace(/_([a-z0-9])/g, function (_, c) {
    return c.toUpperCase();
  });
}

function statusApi_addCamelMirrors_(obj) {
  var out = Object.assign({}, obj);
  Object.keys(obj || {}).forEach(function (k) {
    if (k && k.indexOf('_') >= 0) {
      var camel = statusApi_toCamel_(k);
      if (!(camel in out)) out[camel] = obj[k];
    }
  });
  return out;
}

// ===== submissions 追記ユーティリティ（列名ベース、snake/camel 両対応） =====
function submissions_headerIndexMap_(headers) {
  var m = {};
  (headers || []).forEach(function (h, i) {
    var s = String(h == null ? '' : h);
    m[s] = i; // そのまま
    // camel <-> snake 双方向キーをざっくり登録
    var camel = s.replace(/_([a-z])/g, function (_, c) { return (c || '').toUpperCase(); });
    m[camel] = i;
    var snake = s.replace(/([A-Z])/g, '_$1').toLowerCase();
    m[snake] = i;
  });
  return m;
}

function submissions_appendRow(obj) {
  var props = PropertiesService.getScriptProperties();
  var sid = props.getProperty('BAS_MASTER_SPREADSHEET_ID') || props.getProperty('SHEET_ID');
  if (!sid) throw new Error('submissions_appendRow: missing BAS_MASTER_SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(sid);
  var sh = ss.getSheetByName('submissions');
  if (!sh) sh = ss.insertSheet('submissions');
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 10).setValues([[
      'submission_id','form_key','case_id','user_key','line_id',
      'submitted_at','seq','referrer','redirect_url','status'
    ]]);
  }

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idx = submissions_headerIndexMap_(headers);

  if (!submissions_requireHeadersOrWarn_(idx)) {
    try { Logger.log('[submissions] append skipped (please add headers: submission_id, form_key, case_id)'); } catch (_) {}
    return;
  }

  var meta = obj && typeof obj.meta === 'object' ? obj.meta : {};
  var rawLineId = obj.line_id || meta.line_id || '';
  var lineId = String(rawLineId || '').trim();
  var inferredUserKey = '';
  try {
    if (!obj.user_key && !meta.user_key && lineId && typeof drive_userKeyFromLineId_ === 'function') {
      inferredUserKey = drive_userKeyFromLineId_(lineId) || '';
    }
  } catch (_) {}
  var userKey = normalizeUserKey_(obj.user_key || meta.user_key || inferredUserKey || '');
  var caseId = normalizeCaseId_(obj.case_id || meta.case_id || '');
  var caseKey = obj.case_key || meta.case_key || (userKey && caseId ? userKey + '-' + caseId : '');
  if (meta) {
    if (lineId && !meta.line_id) meta.line_id = lineId;
    meta.user_key = userKey;
    meta.case_id = caseId;
    if (caseKey) meta.case_key = caseKey;
    obj.meta = meta;
  }

  var row = new Array(headers.length).fill('');
  var payload = {
    submission_id: obj.submission_id || '',
    form_key: obj.form_key || '',
    case_id: caseId,
    user_key: userKey,
    line_id: lineId,
    submitted_at: obj.submitted_at || new Date().toISOString(),
    seq: obj.seq || '',
    referrer: obj.referrer || '',
    redirect_url: obj.redirect_url || '',
    status: obj.status || 'received',
    case_key: caseKey,
    json_path: obj.json_path || '',
  };

  Object.keys(payload).forEach(function (k) {
    if (idx[k] != null) row[idx[k]] = payload[k];
  });

  var targetRow = sh.getLastRow() + 1;
  sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
  if (idx.case_id != null) {
    // 列全体をテキスト書式（初期化的に一度かけても害はない）
    sh.getRange(1, idx.case_id + 1, sh.getMaxRows(), 1).setNumberFormat('@');
  }
}

function extractSubmissionId_(name) {
  var m = String(name || '').match(/__(\d+)\.json$/);
  return m ? m[1] : '';
}

function submissions_requireHeadersOrWarn_(idx) {
  try {
    var must = ['submission_id', 'form_key', 'case_id'];
    var miss = must.filter(function (k) { return idx[k] == null; });
    if (miss.length) Logger.log('[submissions] missing headers: ' + miss.join(','));
    return miss.length === 0;
  } catch (_) {
    return true;
  }
}

function submissions_hasRow_(submissionId, formKey) {
  try {
    var props = PropertiesService.getScriptProperties();
    var sid = props.getProperty('BAS_MASTER_SPREADSHEET_ID') || props.getProperty('SHEET_ID');
    if (!sid) return false;
    var ss = SpreadsheetApp.openById(sid);
    var sh = ss.getSheetByName('submissions');
    if (!sh) return false;
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var colCount = headers.length; // ヘッダの実列数に合わせる（右端のゴミ列を無視）
    var idx = submissions_headerIndexMap_(headers);
    if (!submissions_requireHeadersOrWarn_(idx)) return false;
    var rc = sh.getLastRow() - 1;
    if (rc < 1) return false;
    var range = sh.getRange(2, 1, rc, colCount).getValues();
    var cSub = idx['submission_id'];
    var cForm = idx['form_key'];
    if (cSub == null || cForm == null) return false;
    for (var i = 0; i < range.length; i++) {
      var sidv = String(range[i][cSub] || '').trim();
      var fkv = String(range[i][cForm] || '').trim();
      if (sidv && fkv && sidv === String(submissionId) && fkv === String(formKey)) return true;
    }
  } catch (_) {}
  return false;
}

// ===== V2 署名サポート（base64url） =====
function statusApi_b64u_(s) {
  return String(s || '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
}

function statusApi_constEq_(a, b) {
  a = String(a || '');
  b = String(b || '');
  if (a.length !== b.length) return false;
  var d = 0;
  for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}

/**
 * V2 検証: {p, ts, sig} → { ok, lineId, caseId, ts }
 */
function statusApi_verifyV2_(params) {
  try {
    var p = String(params && params.p ? params.p : '').trim();
    if (!p) return { ok: false };
    var decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(p)).getDataAsString();
    var parts = decoded.split('|');
    var li = parts[0] || '';
    var ci = parts[1] || '';
    var t = parts[2] || '';
    var ts = String((params && params.ts) || t || '').trim();
    if (!ts) return { ok: false };

    var payload = [li, ci, ts].join('|');
    var sec = statusApi_getSecret_();
    var mac = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, payload, sec);
    var expected = statusApi_b64u_(Utilities.base64EncodeWebSafe(mac));
    var provided = statusApi_b64u_(String((params && params.sig) || ''));
    var ok = statusApi_constEq_(expected, provided);
    // 追加: 時刻スキュー検証（±600秒）
    try {
      var nowSec = Math.floor(Date.now() / 1000);
      var tsNum = Number(ts);
      if (!Number.isFinite(tsNum) || Math.abs(tsNum - nowSec) > 600) {
        return { ok: false };
      }
    } catch (_) {}
    return { ok: ok, lineId: li, caseId: ci, ts: ts };
  } catch (_) {
    return { ok: false };
  }
}

// NOTE: doGet/status 経路では絶対にフォルダ新規作成しないこと。
//       既存フォルダが無ければ return。限定救済は“本人の intake 実在時のみ”。
function statusApi_collectStaging_(lineId, caseId) {
  const lid = String(lineId || '').trim();
  const cid = String(caseId || '').trim();
  // NOTE: caseId が空でも進める（staging のファイル側 meta から救済）
  try {
    var appended = 0; // submissions 追記件数（診断ログ用）
    var moved = 0; // staging からの移送件数
    var attachReason = ''; // フェイルセーフ時の一致理由（line_id|email|case_id|case_key）
    var foundNames = [];
    var ensureProp = (STATUS_API_PROPS.getProperty('ALLOW_GET_ENSURE_CASE_FOLDER') || '').trim();
    var allowGetEnsure = ensureProp ? ensureProp === '1' : false;
    // userKey → 既存ケースフォルダの解決（作成はしない）
    var uk = '';
    try {
      var row = typeof drive_lookupCaseRow_ === 'function' ? drive_lookupCaseRow_({ caseId: cid, lineId: lid }) : null;
      if (row && row.userKey) uk = row.userKey;
    } catch (_) {}
    if (!uk && typeof drive_userKeyFromLineId_ === 'function') uk = drive_userKeyFromLineId_(lid);

    // NOTE: ケースフォルダの解決は「既存参照のみ」。
    //       ここでは絶対に getOrCreate 系は使わない（doGet/status 経路で新規作成を防止）。
    var folderId = '';
    try {
      if (typeof bs_resolveCaseFolderId_ === 'function') {
        folderId = bs_resolveCaseFolderId_(uk, cid, lid); // 既存参照のみ（作成しない）
      } else if (typeof resolveCaseFolderId_ === 'function') {
        // 呼び出し規約: createIfMissing は必ず false 固定（作成しない）
        folderId = resolveCaseFolderId_(lid, cid, /*createIfMissing=*/false);
      }
    } catch (_) {}

    // 既存フォルダがある場合も、staging に残っている intake__*.json をケース直下へ移送（ensure はしない）
    if (folderId) {
      try {
        var caseFolder = DriveApp.getFolderById(String(folderId));
        var foundMove = null;
        foundNames = foundNames || [];
        var stagingListMove = [];
        try {
          var rootMv = (typeof drive_getRootFolder_ === 'function') ? drive_getRootFolder_() : null;
          ['_email_staging', '_staging'].forEach(function (name) {
            try {
              var it = rootMv ? rootMv.getFoldersByName(name) : DriveApp.getFoldersByName(name);
              if (it && it.hasNext()) stagingListMove.push(it.next());
            } catch (_) {}
          });
        } catch (_) {}
        var ckeyMv = (uk ? uk + '-' + cid : '');
        stagingListMove.forEach(function (staging) {
          if (foundMove) return;
          var itf = staging.getFiles();
          while (!foundMove && itf.hasNext()) {
            var f = itf.next();
            var nm = f.getName && f.getName();
            if (!nm || !isIntakeJsonName_(nm)) continue;
            try {
              var txt = '';
              try { txt = f.getBlob().getDataAsString('utf-8'); } catch (_) {}
              var js = null;
              try { js = JSON.parse(txt); } catch (_) {}
              var m = (js && js.meta) || {};
              var res_mv = matchMetaToCase_(m, { case_key: ckeyMv, case_id: cid, line_id: lid });
              if (res_mv && res_mv.ok) { attachReason = res_mv.by || attachReason; foundMove = { file: f, name: nm }; }
            } catch (_) {}
          }
        });
        if (foundMove) {
          try {
            try { Logger.log('[collectStaging] move name=%s by=%s → folderId=%s', foundMove.name || '', attachReason || '', String(folderId)); } catch (_) {}
            caseFolder.createFile(foundMove.file.getBlob());
            try { foundMove.file.setTrashed(true); } catch (_) {}
            try { moved++; } catch (_) {}
          } catch (_) {}
          // submissions 補填（重複ガード）
          try {
            var subIdMv = extractSubmissionId_(foundMove.name);
            if (subIdMv && typeof submissions_hasRow_ === 'function' && typeof submissions_appendRow === 'function') {
              if (!submissions_hasRow_(subIdMv, 'intake')) {
                submissions_appendRow({
                  submission_id: subIdMv,
                  form_key: 'intake',
                  case_id: cid,
                  user_key: uk,
                  line_id: lid,
                  submitted_at: new Date().toISOString(),
                  status: 'received',
                  seq: '',
                  referrer: foundMove && foundMove.name || '',
                  redirect_url: '',
                });
                try { appended++; } catch (_) {}
              }
            }
          } catch (_) {}
        }
      } catch (_) {}
    }

    // 既存フォルダが無い場合、限定的に救済：この lineId の intake__*.json が staging に存在する時だけ ensure
    if (!folderId) {
      try { Logger.log('[collectStaging] no case folder yet; probing staging for lid=%s cid=%s', lid, cid); } catch (_) {}
      var foundForThis = null;
      try {
        // まず _email_staging と _staging を順に探す
        var stagingList = [];
        try {
          var root = (typeof drive_getRootFolder_ === 'function') ? drive_getRootFolder_() : null;
          ['_email_staging', '_staging'].forEach(function(name){
            try {
              var it = root ? root.getFoldersByName(name) : DriveApp.getFoldersByName(name);
              if (it && it.hasNext()) stagingList.push(it.next());
            } catch(_){}
          });
        } catch(_){}
        stagingList.forEach(function(staging){
          if (foundForThis) return;
          var ukey = typeof drive_userKeyFromLineId_ === 'function' ? drive_userKeyFromLineId_(lid) : '';
          var ckey = (ukey ? ukey + '-' : '') + cid;
          var itf = staging.getFiles();
          while (!foundForThis && itf.hasNext()) {
            var f = itf.next();
            var nm = f.getName && f.getName();
            if (!nm) continue;
            if (isIntakeJsonName_(nm)) { try { foundNames.push(nm); } catch(_) {} } else { continue; }
            try {
              var txt = '';
              try { txt = f.getBlob().getDataAsString('utf-8'); } catch (_) {}
              var js = null;
              try { js = JSON.parse(txt); } catch (_) {}
              var m = (js && js.meta) || {};
              var res_nf = matchMetaToCase_(m, { case_key: ckey, case_id: cid, line_id: lid });
              if (res_nf && res_nf.ok) {
                attachReason = res_nf.by || attachReason;
                foundForThis = {
                  file: f,
                  js: js,
                  name: nm,
                  fileCKey: String(m.case_key || js.case_key || '').trim(),
                  fileCID: String(m.case_id || js.case_id || '').trim(),
                };
              }
            } catch (_) {}
          }
        });
      } catch (_) {}

      if (foundForThis) {
        if (!allowGetEnsure) {
          try { Logger.log('[collectStaging] ensure skipped by policy (ALLOW_GET_ENSURE_CASE_FOLDER=0) lid=%s cid=%s', lid, cid); } catch (_) {}
          return;
        } else {
          try {
            // 限定 ensure: intake が本人のものと断定できた時だけ getOrCreate 可
            // - file.case_key が正ならそれを優先
            // - それが無ければ、file.case_id と ukey から生成
            var ensureKey = '';
            var validCaseKey = function (s) { s = String(s || '').trim(); return !!s && /^[a-z0-9]{2,}-\d{4}$/i.test(s); };
            var fileCIDn = statusApi_normCaseId_(String((foundForThis && foundForThis.fileCID) || ''));
            if (validCaseKey(foundForThis && foundForThis.fileCKey)) {
              ensureKey = foundForThis.fileCKey;
            } else if (uk && fileCIDn) {
              ensureKey = uk + '-' + fileCIDn;
            }
            // "-0001" など不正キーは使わない
            if (!validCaseKey(ensureKey)) return;
            var cf = drive_getOrCreateCaseFolderByKey_(ensureKey);
            var caseFolder = (cf && typeof cf.getId === 'function') ? cf : DriveApp.getFolderById(String(cf));
            folderId = caseFolder.getId();
            // ファイルを案件直下へコピー（原本は削除）
            try {
              try { Logger.log('[collectStaging] move name=%s by=%s → folderId=%s', foundForThis.name || '', attachReason || '', String(folderId)); } catch (_) {}
              caseFolder.createFile(foundForThis.file.getBlob());
              try { foundForThis.file.setTrashed(true); } catch (_) {}
            } catch (_) {}
            try { moved++; } catch (_) {}
            // submissions upsert（重複ガード）
            try {
            var nm2 = foundForThis.name;
            var subId = nm2 ? extractSubmissionId_(nm2) : '';
            if (subId && typeof upsertSubmission_ === 'function') {
              var ensuredCid = (ensureKey.match(/-(\d{4})$/) || [])[1] || cid;
              var normEnsuredCid = statusApi_normCaseId_(ensuredCid);
              upsertSubmission_({
                submission_id: subId,
                form_key: 'intake',
                case_id: normEnsuredCid,
                user_key: uk,
                line_id: lid,
                submitted_at: new Date().toISOString(),
                status: 'received',
                referrer: nm2 || '',
                redirect_url: '',
              });
              try { appended++; } catch (_) {}
            }
            } catch (_) {}
            // 限定ensureによりフォルダを起こした直後に、contacts/cases も揃える
            try { statusApi_ensureActiveCaseIdString_(lid, (ensureKey.match(/-(\d{4})$/) || [])[1] || cid); } catch (_) {}
            if (typeof updateCasesRow_ === 'function') {
              try {
                var cidEnsured = (ensureKey.match(/-(\d{4})$/) || [])[1] || cid;
                updateCasesRow_(cidEnsured, { case_key: ensureKey, folder_id: folderId, status: 'intake', updated_at: new Date() });
              } catch (_) {}
            }
          } catch (_) {}
        }
      } else {
        // intake がまだ無ければ、ここでは何もしない（フォルダも作らない）
        try { Logger.log('[collectStaging] no intake json for lid=%s, skip ensure', lid); } catch (_) {}
        return;
      }
    }

    // 既存フォルダがあっても、自前で staging を走査して intake__*.json を移送（委譲に頼らない）
    try {
      var caseFolder = DriveApp.getFolderById(folderId);
      var root = (typeof drive_getRootFolder_ === 'function') ? drive_getRootFolder_() : null;
      var stagingList = [];
      ['_email_staging', '_staging'].forEach(function (name) {
        try {
          var it = root ? root.getFoldersByName(name) : DriveApp.getFoldersByName(name);
          if (it && it.hasNext()) stagingList.push(it.next());
        } catch (_) {}
      });
      var ukey = typeof drive_userKeyFromLineId_ === 'function' ? drive_userKeyFromLineId_(lid) : uk;
      var ckey = (ukey ? (ukey + '-' + cid) : '');

      var found = null;
      stagingList.forEach(function (staging) {
        if (found) return;
        var itf = staging.getFiles();
        while (!found && itf.hasNext()) {
            var f = itf.next();
            var nm = f.getName && f.getName();
            if (!nm) continue;
            if (isIntakeJsonName_(nm)) { try { foundNames.push(nm); } catch(_) {} } else { continue; }
          try {
            var txt = f.getBlob().getDataAsString('utf-8');
            var js = JSON.parse(txt);
            var m = (js && js.meta) || {};
            var fileCKey = String(m.case_key || js.case_key || '').trim();
            var fileCID  = String(m.case_id  || js.case_id  || '').trim();
            var fileLID  = String(m.line_id  || m.lineId    || '').trim();
            var matched = false;
            if (fileCKey && fileCKey === ckey) matched = true;
            else if (fileCID && statusApi_normCaseId_(fileCID) === cid) matched = true;
            else if (fileLID && fileLID === lid) matched = true;
            if (matched) { found = { file: f, name: nm, js: js, ckey: fileCKey }; }
          } catch (_) {}
        }
      });
      if (found) {
        try {
          // そのままコピーすると BOM/改行扱いが変わることがあるので一度 JSON stringify（meta 補完も可）
          var obj = found.js || {};
          obj.meta = obj.meta || {};
          if (!obj.meta.case_id)  obj.meta.case_id  = cid;
          if (!obj.meta.user_key) obj.meta.user_key = ukey || uk;
          if (!obj.meta.case_key && (ukey || uk)) obj.meta.case_key = (ukey || uk) + '-' + cid;
          var saved = JSON.stringify(obj);
          try { Logger.log('[collectStaging] move name=%s by=%s → folderId=%s', found.name || '', attachReason || '', String(folderId)); } catch (_) {}
          caseFolder.createFile(Utilities.newBlob(saved, 'application/json', found.name));
          try { found.file.setTrashed(true); } catch (_) {}
          try { moved++; } catch (_) {}
        } catch (_) {}
      }
    } catch (_) {}

    // submissions ログ（intake）の追記：案件直下の intake__*.json を検出
    try {
      try {
        var caseFolder = DriveApp.getFolderById(folderId);
        var it = caseFolder.getFiles();
        while (it.hasNext()) {
          var f = it.next();
          var nm = f.getName && f.getName();
          if (!nm || !/^intake__/i.test(nm)) continue;
          var subId = extractSubmissionId_(nm);
          if (subId && typeof upsertSubmission_ === 'function') {
            var normCid3 = statusApi_normCaseId_(cid);
            var nextSeq3 = 1;
            try { nextSeq3 = (getLastSeq_(normCid3, 'intake') | 0) + 1; } catch (_) { nextSeq3 = 1; }
            upsertSubmission_({
              submission_id: subId,
              form_key: 'intake',
              case_id: normCid3,
              user_key: uk,
              line_id: lid,
              submitted_at: new Date().toISOString(),
              status: 'received',
              seq: nextSeq3,
              referrer: nm,
              redirect_url: '',
            });
            try { appended++; } catch (_) {}
          }
        }
      } catch (_) {}
    } catch (_) {}
    // --- 限定フェイルセーフ（ポリシー §4.3）: moved=0 のときだけ、直近10分以内の最新1件を救済 ---
    try {
      if ((moved | 0) === 0) {
        var rootFS = (typeof drive_getRootFolder_ === 'function') ? drive_getRootFolder_() : null;
        var stagingFS = null;
        try {
          ['_email_staging', '_staging'].some(function (name) {
            try {
              var it = rootFS ? rootFS.getFoldersByName(name) : DriveApp.getFoldersByName(name);
              if (it && it.hasNext()) {
                stagingFS = it.next();
                return true;
              }
            } catch (_) {}
            return false;
          });
        } catch (_) {}
        if (stagingFS) {
          var newest = null;
          var newestTime = 0;
          var itf2 = stagingFS.getFiles();
          while (itf2.hasNext()) {
            var ff = itf2.next();
            var nmf = ff.getName && ff.getName();
            if (!isIntakeJsonName_(String(nmf || ''))) continue;
            var mt = (ff.getLastUpdated && ff.getLastUpdated()) || new Date(0);
            var t = +mt;
            if (t > newestTime) {
              newestTime = t;
              newest = ff;
            }
          }
          // 直近5分以内のみ救済（衝突低減）
          if (newest && (Date.now() - newestTime) <= 5 * 60 * 1000) {
            // 誤アタッチ防止の追加一致判定
          var txtFS = '', jsFS = {};
          try { txtFS = newest.getBlob().getDataAsString('utf-8'); } catch(_){}
          try { jsFS = JSON.parse(txtFS||'{}'); } catch(_) { jsFS = {}; }
          var mFS = (jsFS && jsFS.meta) || {};
          var cidFS = String(mFS.case_id || jsFS.case_id || '').trim();
          var cid4 = statusApi_normCaseId_(cidFS || cid);
          var okAttach = false;
          var ukeyFS_chk = (typeof drive_userKeyFromLineId_==='function') ? drive_userKeyFromLineId_(lid) : '';
          var mt3 = matchMetaToCase_(
            { case_key: String(mFS.case_key || jsFS.case_key || ''), case_id: cidFS, line_id: String(mFS.line_id || mFS.lineId || '') },
            { case_key: (ukeyFS_chk ? (ukeyFS_chk + '-' + cid4) : ''), case_id: cid4, line_id: lid }
          );
          if (mt3 && mt3.ok) { okAttach = true; attachReason = mt3.by || attachReason; }
          var emailFS = String(
              (jsFS.fields && (jsFS.fields.email || jsFS.fields['メールアドレス'])) ||
              (jsFS.model  && (jsFS.model.email  || jsFS.model['email'])) || ''
            ).trim();
          if (!okAttach && emailFS && typeof contacts_lookupByEmail_ === 'function') {
            var cinfoFS = contacts_lookupByEmail_(emailFS);
            if (cinfoFS && cinfoFS.line_id && cinfoFS.line_id === lid) { okAttach = true; attachReason = 'email'; }
          }
            if (!okAttach) {
              if (cidFS && statusApi_normCaseId_(cidFS) === cid4) { okAttach = true; attachReason = 'case_id'; }
              var ukeyFS_chk2 = (typeof drive_userKeyFromLineId_==='function') ? drive_userKeyFromLineId_(lid) : '';
              var ckeyFS0 = String(mFS.case_key || jsFS.case_key || '').trim().toLowerCase();
              if (!okAttach && ckeyFS0 && ukeyFS_chk2) {
                if (ckeyFS0 === (ukeyFS_chk2 + '-' + cid4)) { okAttach = true; attachReason = 'case_key'; }
              }
            }
            // 追加: 候補がユニーク1件なら開発用に救済（プロパティ ALLOW_UNIQUE_RESCUE=1 のときだけ）
            if (!okAttach) {
              var allowUnique = (STATUS_API_PROPS.getProperty('ALLOW_UNIQUE_RESCUE') || '').trim() === '1';
              if (allowUnique) {
                var total = 0;
                try {
                  ['_email_staging', '_staging'].forEach(function (name) {
                    var itx = rootFS ? rootFS.getFoldersByName(name) : DriveApp.getFoldersByName(name);
                    if (itx && itx.hasNext()) {
                      var fdr = itx.next();
                      var itf3 = fdr.getFiles();
                      while (itf3.hasNext()) {
                        var f3 = itf3.next();
                        var nm3 = f3.getName && f3.getName();
                        if (/^intake__\d+\.json$/i.test(String(nm3 || ''))) total++;
                      }
                    }
                  });
                } catch (_) {}
                if (total === 1) { okAttach = true; attachReason = 'unique'; }
              }
            }
            if (!okAttach) return; // 救済しない
            var folderIdFS = '';
            try {
              if (allowGetEnsure && typeof resolveCaseFolderId_ === 'function') {
                folderIdFS = resolveCaseFolderId_(lid, cid4, /*createIfMissing=*/true);
              } else if (!allowGetEnsure) {
                try { Logger.log('[collectStaging] resolveCaseFolderId_ skipped by policy for lid=%s cid=%s', lid, cid4); } catch (_) {}
              }
            } catch (_) {}
            if (folderIdFS) {
              var ukeyFS = (typeof drive_userKeyFromLineId_ === 'function') ? drive_userKeyFromLineId_(lid) : '';
              var ckeyFS = (ukeyFS ? ukeyFS + '-' : '') + cid4;
              var subIdFS = (String(newest.getName()).match(/__(\d+)\.json$/) || [])[1] || String(Date.now());
              var obj = {};
              try { obj = JSON.parse((newest.getBlob() && newest.getBlob().getDataAsString()) || '{}'); } catch (_) { obj = {}; }
              obj = obj && typeof obj === 'object' ? obj : {};
              obj.meta = obj.meta || {};
              if (!obj.meta.line_id) obj.meta.line_id = lid;
              if (!obj.meta.user_key) obj.meta.user_key = ukeyFS;
              if (!obj.meta.case_id) obj.meta.case_id = cid4;
              if (!obj.meta.case_key) obj.meta.case_key = ckeyFS;
              var jsonNameFS = 'intake__' + subIdFS + '.json';
              try {
                DriveApp.getFolderById(folderIdFS).createFile(
                  Utilities.newBlob(JSON.stringify(obj), 'application/json', jsonNameFS)
                );
                try { newest.setTrashed(true); } catch (_) {}
                moved++;
                if (
                  typeof submissions_hasRow_ === 'function' &&
                  typeof submissions_appendRow === 'function'
                ) {
                  if (!submissions_hasRow_(subIdFS, 'intake')) {
                    submissions_appendRow({
                      submission_id: subIdFS,
                      form_key: 'intake',
                      case_id: cid4,
                      user_key: ukeyFS,
                      line_id: lid,
                      submitted_at: new Date().toISOString(),
                      status: 'received',
                      seq: '',
                      referrer: jsonNameFS,
                      redirect_url: '',
                    });
                    appended++;
                  }
                }
              } catch (_) {}
            }
          }
        }
      }
    } catch (_) {}
    try { Logger.log('[collectStaging] lid=%s cid=%s moved=%s appended=%s by=%s', lid, cid, moved, appended, attachReason); } catch (_) {}
    if (!moved || moved === 0) {
      try { Logger.log('[collectStaging] candidates(name)=%s', JSON.stringify((foundNames || []).slice(0, 10))); } catch (_) {}
      // 直近60秒以内に staging にファイルが来たら一度だけ再試行
      try {
        var cache = CacheService.getScriptCache();
        var onceKey = ['collect_retry_once', lid || '', cid || ''].join(':');
        if (!cache.get(onceKey) && tryFindRecentIntakeJson_(lid, cid, 60)) {
          cache.put(onceKey, '1', 120); // 2分の間は再入禁止
          try { Logger.log('[collectStaging] retry-after-recent-staging lid=%s cid=%s', lid, cid); } catch(_){}
          try { statusApi_collectStaging_(lid, cid); } catch(_){}
        }
      } catch (_) {}
    }
  } catch (err) {
    try {
      Logger.log('[status_api] staging sweep error: %s', (err && err.stack) || err);
    } catch (_) {}
  }
}

// 共通ルート：lineId / caseIdHint から caseId を確定→"0001" 正規化→保存→吸い上げ→submissions 補填
function statusApi_routeStatus_(lineIdRaw, caseIdHintRaw) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) return; // 軽量ロックで同時多発を抑止
  try {
    var lineId = String(lineIdRaw || '').trim();
    var caseIdHintRawStr = String(caseIdHintRaw || '').trim();
    // caseId 決定（優先: hint → contacts.lookup）
    var caseId = (function () {
      if (typeof statusApi_computeCaseId_ === 'function') {
        return statusApi_computeCaseId_(lineId, caseIdHintRawStr || '');
      }
      var cid = caseIdHintRawStr;
      if (!cid && typeof lookupCaseIdByLineId_ === 'function') {
        cid = lookupCaseIdByLineId_(lineId) || '';
      }
      return cid; // 既存が無ければ空のまま
    })();
    caseId = statusApi_normCaseId_(caseId);

    try { if (!lineId && !caseId) Logger.log("BAS:status:common { lineId:'' , caseId:'' }"); } catch (_) {}

    // contacts.active_case_id をテキストで保存（数値化防止）
    if (caseId) statusApi_ensureActiveCaseIdString_(lineId, caseId);
    // cases.case_id の列書式を固定
    ensureCasesCaseIdTextFormat_();
    ensureSubmissionsCaseIdTextFormat_();
    // staging 吸い上げ + submissions 補填
    var allowSweep = (STATUS_API_PROPS.getProperty('ALLOW_STAGING_SWEEP_ON_STATUS') || '').trim() === '1';
    if (allowSweep) statusApi_collectStaging_(lineId, caseId || '');

    try {
      if (typeof logCtx_ === 'function') logCtx_('status:route:common', { lineId: lineId, caseId: caseId });
      else Logger.log('[status:route:common] ' + JSON.stringify({ lineId: lineId, caseId: caseId }));
    } catch (_) {}
  } catch (err) {
    try { Logger.log('[statusApi_routeStatus_] ' + String((err && err.stack) || err)); } catch (_) {}
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// cases.case_id 列をテキスト('@')に固定（数値化防止）
function ensureCasesCaseIdTextFormat_() {
  try {
    if (STATUS_API_CASES_CASE_ID_FORMATTED) return;
    var prop = (STATUS_API_PROPS.getProperty('CASES_CASE_ID_FORMATTED') || '').trim();
    if (prop === '1') {
      STATUS_API_CASES_CASE_ID_FORMATTED = true;
      return;
    }
    var props = PropertiesService.getScriptProperties();
    var sid = props.getProperty('BAS_MASTER_SPREADSHEET_ID') || props.getProperty('SHEET_ID');
    if (!sid) return;
    var ss = SpreadsheetApp.openById(sid);
    var sh = ss.getSheetByName('cases');
    if (!sh) return;
    var headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    var col = -1;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i] || '').trim();
      if (!h) continue;
      var low = h.toLowerCase();
      if (low === 'case_id' || low === 'caseid') {
        col = i + 1;
        break;
      }
    }
    if (col > 0) {
      sh.getRange(1, col, sh.getMaxRows(), 1).setNumberFormat('@');
      STATUS_API_CASES_CASE_ID_FORMATTED = true;
      STATUS_API_PROPS.setProperty('CASES_CASE_ID_FORMATTED', '1');
    }
  } catch (_) {}
}

function ensureSubmissionsCaseIdTextFormat_() {
  try {
    if (STATUS_API_SUBMISSIONS_CASE_ID_FORMATTED) return;
    var prop = (STATUS_API_PROPS.getProperty('SUBMISSIONS_CASE_ID_FORMATTED') || '').trim();
    if (prop === '1') {
      STATUS_API_SUBMISSIONS_CASE_ID_FORMATTED = true;
      return;
    }
    var props = PropertiesService.getScriptProperties();
    var sid = props.getProperty('BAS_MASTER_SPREADSHEET_ID') || props.getProperty('SHEET_ID');
    if (!sid) return;
    var ss = SpreadsheetApp.openById(sid);
    var name = (typeof SHEETS_REPO_SUBMISSIONS !== 'undefined' && SHEETS_REPO_SUBMISSIONS) || 'submissions';
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    var col = -1;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i] || '').trim();
      if (!h) continue;
      var low = h.toLowerCase();
      if (low === 'case_id' || low === 'caseid') {
        col = i + 1;
        break;
      }
    }
    if (col > 0) {
      sh.getRange(1, col, sh.getMaxRows(), 1).setNumberFormat('@');
      STATUS_API_SUBMISSIONS_CASE_ID_FORMATTED = true;
      STATUS_API_PROPS.setProperty('SUBMISSIONS_CASE_ID_FORMATTED', '1');
    }
  } catch (_) {}
}

/** GET /exec?action=status&caseId=&lineId=&ts=&sig= */
function doGet(e) {
  const p = (e && e.parameter) || {};
  // 追加: ルートIDの観測（設定ズレの早期検知）
  try {
    var PROPS = PropertiesService.getScriptProperties();
    Logger.log('[config] ROOT_FOLDER_ID=%s DRIVE_ROOT_FOLDER_ID=%s', PROPS.getProperty('ROOT_FOLDER_ID') || '', PROPS.getProperty('DRIVE_ROOT_FOLDER_ID') || '');
  } catch (_) {}
  try {
    if (typeof logCtx_ === 'function') {
      logCtx_('router:in', {
        keys: Object.keys(p),
        action: String(p.action || ''),
        has_ts: typeof p.ts !== 'undefined',
        has_sig: typeof p.sig !== 'undefined',
        has_p: typeof p.p !== 'undefined',
      });
    } else {
      Logger.log('[router:in] ' + JSON.stringify({
        keys: Object.keys(p),
        action: String(p.action || ''),
        has_ts: typeof p.ts !== 'undefined',
        has_sig: typeof p.sig !== 'undefined',
        has_p: typeof p.p !== 'undefined',
      }));
    }
  } catch (_) {
    try { Logger.log('[router:in] log failed'); } catch (_) {}
  }
  const action = String(p.action || '').trim();
  const debugMode = String(p.debug || '').trim();
  // 1) 署名不要の ping は即返す（疎通確認用）
  if (p.ping === '1') {
    return statusApi_jsonOut_({ ok: true, via: 'status_api', ping: true }, 200);
  }
  if (debugMode === 'sig' || debugMode === 'sigcheck') {
    var allowDebugSig = (STATUS_API_PROPS.getProperty('ALLOW_DEBUG_SIG') || '').trim() === '1';
    if (!allowDebugSig) {
      return statusApi_jsonOut_({ ok: false, error: 'debug_sig_disabled' }, 403);
    }
    try {
      var dbgLine = String(p.lineId || p.line_id || '').trim();
      var dbgCase = String(p.caseId || p.case_id || '').trim();
      var dbgTs = String(p.ts || Math.floor(Date.now() / 1000)).trim();
      var secret = statusApi_getSecret_();
      if (!secret) return statusApi_jsonOut_({ ok: false, error: 'secret_not_configured' }, 400);
      var v1Msg = dbgTs + '.' + dbgLine + '.' + dbgCase;
      var expectedV1 = statusApi_hmac_(v1Msg, secret);
      var payload = [dbgLine, dbgCase, dbgTs].join('|');
      var mac = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, payload, secret);
      var expectedV2 = statusApi_b64u_(Utilities.base64EncodeWebSafe(mac));
      return statusApi_jsonOut_({ ok: true, lineId: dbgLine, caseId: dbgCase, ts: dbgTs, expectedSigV1: expectedV1, expectedSigV2: expectedV2 }, 200);
    } catch (err) {
      return statusApi_jsonOut_({ ok: false, error: String(err) }, 400);
    }
  }
  // ★ bootstrap は try/catch で包んで JSON エラーに変換（＋共通処理を起動）
  if (action === 'bootstrap') {
    var resp;
    try {
      resp = bootstrap_(e);
    } catch (err) {
      return statusApi_jsonOut_({ ok: false, error: String(err) }, 400);
    }
    // ここで cases.case_id の列書式を固定
    ensureCasesCaseIdTextFormat_();
    ensureSubmissionsCaseIdTextFormat_();
    // 可能なら V2 から lineId/caseId を復元し、共通ルートへ（staging 吸い上げ・submissions 追記）
    try {
      var pv2 = statusApi_verifyV2_(p) || {};
      var qp = (e && e.parameter) || {};
      var getField = function (obj, k) { try { var v = obj && obj[k]; return (v != null ? String(v).trim() : ''); } catch (_) { return ''; } };
      var lineId2 = pv2.ok ? (pv2.lineId || pv2.line_id || '') : (getField(qp, 'lineId') || getField(qp, 'line_id'));
      var caseIdHint2 = pv2.ok ? (pv2.caseId || pv2.case_id || '') : (getField(qp, 'caseId') || getField(qp, 'case_id'));
      try {
        Logger.log(
          'BAS:bootstrap:post-status-route { lineId:%s, caseIdHint:%s }',
          lineId2 ? '[ok]' : "''",
          caseIdHint2 ? '[ok]' : "''"
        );
      } catch (_) {}
      // 副作用目的で呼ぶ（レスポンスは返さない）
      try { statusApi_routeStatus_(lineId2, caseIdHint2); } catch (e2) {
        try { Logger.log('[bootstrap->routeStatus] ' + String((e2 && e2.stack) || e2)); } catch (_) {}
      }
    } catch (_) {}
    return resp;
  }
  // 3) それ以外（status 等）は最初に署名検証
  // まず V2（p/ts/sig）を試す。NG でも V1（HEX）は許容（フェイルセーフ）。
  var v2 = statusApi_verifyV2_(p);
  if (v2.ok) {
    try {
      if (typeof logCtx_ === 'function') {
        logCtx_('status:route', { via: 'v2', lineId: v2.lineId, caseId: v2.caseId });
      } else {
        Logger.log('[status:route] ' + JSON.stringify({ via: 'v2', lineId: v2.lineId, caseId: v2.caseId }));
      }
    } catch (_) {}
    if (!p.lineId) p.lineId = v2.lineId || '';
    if (!p.caseId) p.caseId = v2.caseId || '';
    if (!p.ts) p.ts = v2.ts || '';
  } else {
    // プロパティで公開可否を制御（既定は拒否）
    var allowAnon = (STATUS_API_PROPS.getProperty('ALLOW_ANON_STATUS') || '').trim() === '1';
    try { statusApi_verify_(p); }
    catch (eVerify) {
      if (!allowAnon) {
        return statusApi_jsonOut_({ ok: false, error: 'unauthorized' }, 401);
      }
      // allowAnon=1 のときは従来どおり permissive
    }
    try {
      var lidLog = String((e && e.parameter && e.parameter.lineId) || p.lineId || '');
      if (typeof logCtx_ === 'function') {
        logCtx_('status:route', { via: 'v1-fallback', lineId: lidLog, allowAnon: allowAnon });
      } else {
        Logger.log('[status:route] ' + JSON.stringify({ via: 'v1-fallback', lineId: lidLog, allowAnon: allowAnon }));
      }
    } catch (_) {}
  }
  try {
    if (action === 'status') {
      // 列書式の固定（case_id）
      ensureCasesCaseIdTextFormat_();
      ensureSubmissionsCaseIdTextFormat_();
      const lineId = String((e.parameter || {}).lineId || p.lineId || '').trim();
      let caseId = String((e.parameter || {}).caseId || p.caseId || '').trim();
      if (!caseId) caseId = lookupCaseIdByLineId_(lineId) || '';

      // ★ ここから追加：必ず "0001" 形式へ正規化（文字列）
      caseId = statusApi_normCaseId_(caseId);
      if (!caseId) {
        return statusApi_jsonOut_({ ok: false, error: 'caseId not found' }, 404);
      }

      var bustParam = String((e.parameter || {}).bust || p.bust || '').trim();
      var cache = CacheService.getScriptCache();
      var cacheKey = ['status', caseId].join(':');
      if (bustParam !== '1') {
        var cached = cache.get(cacheKey);
        if (cached) {
          try {
            var parsed = JSON.parse(cached);
            if (parsed && typeof parsed === 'object') {
              try { Logger.log('[status:cache] hit caseId=%s', caseId); } catch (_) {}
              return statusApi_jsonOut_(parsed, 200);
            }
          } catch (_) {}
        }
        try { Logger.log('[status:cache] miss caseId=%s', caseId); } catch (_) {}
      }

      // ★ contacts.active_case_id を "0001" で強制保存（数値化対策）
      statusApi_ensureActiveCaseIdString_(lineId, caseId);

      // ★ 正規化後のIDでstaging吸い上げを実行（intake__*.json → <user_key>-<case_id>/）
      var allowSweep = (STATUS_API_PROPS.getProperty('ALLOW_STAGING_SWEEP_ON_STATUS') || '').trim() === '1';
      if (allowSweep) statusApi_collectStaging_(lineId, caseId);

      const sourceForms = statusApi_formsFromSubmissions_(caseId);
      const forms = (sourceForms || []).map(function (row) {
        const formKey = String(row.form_key || '').trim();
        const snake = {
          case_id: String(row.case_id || row.caseId || caseId || ''),
          form_key: formKey,
          case_key: row.case_key || '',
          status: row.status || '',
          can_edit: statusApi_normalizeBool_(row.can_edit != null ? row.can_edit : row.canEdit),
          reopened_at: row.reopened_at || null,
          locked_reason: row.locked_reason || null,
          reopen_until: row.reopen_until || null,
          reopened_by: row.reopened_by || null,
          reopened_at_epoch:
            row.reopened_at_epoch !== undefined && row.reopened_at_epoch !== null && row.reopened_at_epoch !== ''
              ? Number(row.reopened_at_epoch)
              : null,
          reopen_until_epoch:
            row.reopen_until_epoch !== undefined && row.reopen_until_epoch !== null && row.reopen_until_epoch !== ''
              ? Number(row.reopen_until_epoch)
              : null,
          last_seq: row.last_seq != null ? Number(row.last_seq) || 0 : 0,
        };
        if (snake.status && String(snake.status).toLowerCase() === 'received') snake.status = 'submitted';
        if (!snake.status) snake.status = '';
        if (snake.reopened_at_epoch != null && !Number.isFinite(snake.reopened_at_epoch)) snake.reopened_at_epoch = null;
        if (snake.reopen_until_epoch != null && !Number.isFinite(snake.reopen_until_epoch)) snake.reopen_until_epoch = null;
        const compat = statusApi_addCamelMirrors_(snake);
        compat.caseId = String(row.caseId || caseId || '');
        if (!('canEdit' in compat)) compat.canEdit = compat.can_edit;
        return compat;
      });
      var resp = { ok: true, caseId: caseId, case_id: caseId, forms: forms };
      if (bustParam !== '1') {
        try { cache.put(cacheKey, JSON.stringify(resp), 10); } catch (_) {}
        try { Logger.log('[status:cache] store caseId=%s', caseId); } catch (_) {}
      }
      return statusApi_jsonOut_(resp, 200);
    }

  if (action === 'intake_ack') {
    return statusApi_handleIntakeAck_(p);
  }

  if (action === 'form_ack') {
    return statusApi_handleFormAck_(p);
  }

  if (action === 'markReopen') {
    return statusApi_jsonOut_({ ok: false, error: 'use_post' }, 405);
  }

    return statusApi_jsonOut_({ ok: false, error: 'unknown action' }, 400);
  } catch (err) {
    return statusApi_jsonOut_({ ok: false, error: String(err) }, 400);
  }
}

function statusApi_handleIntakeAck_(params) {
  var p = params || {};
  var v2 = statusApi_verifyV2_(p);
  var lineId = String(p.lineId || p.line_id || '').trim();
  var caseIdHint = String(p.caseId || p.case_id || '').trim();
  var paramCaseIdNorm = statusApi_normCaseId_(String(p.caseId || p.case_id || ''));
  var tsRaw = String(p.ts || '').trim();
  var sig = String(p.sig || '').trim();

  if (v2 && v2.ok) {
    if (!lineId) lineId = String(v2.lineId || '').trim();
    if (!caseIdHint) caseIdHint = String(v2.caseId || '').trim();
    if (!tsRaw) tsRaw = String(v2.ts || '').trim();
  } else {
    try {
      statusApi_verify_(p);
    } catch (err) {
      return statusApi_jsonOut_({ ok: false, error: String(err) }, 400);
    }
  }

  if (!lineId) {
    return statusApi_jsonOut_({ ok: false, error: 'missing lineId' }, 400);
  }
  if (!tsRaw) {
    return statusApi_jsonOut_({ ok: false, error: 'missing ts' }, 400);
  }
  if (!sig) sig = String(p.sig || '').trim();
  if (!sig) {
    return statusApi_jsonOut_({ ok: false, error: 'missing sig' }, 400);
  }

  try {
    statusApi_assertNonce_(lineId, tsRaw, sig);
  } catch (err) {
    return statusApi_jsonOut_({ ok: false, error: String(err) }, 409);
  }

  var caseId = caseIdHint;
  if (!caseId && paramCaseIdNorm) caseId = paramCaseIdNorm;
  if (!caseId) caseId = lookupCaseIdByLineId_(lineId);
  if (!caseId) {
    return statusApi_jsonOut_({ ok: false, error: 'case_not_found' }, 404);
  }
  var normCaseId = statusApi_normCaseId_(caseId);
  if (!normCaseId) {
    return statusApi_jsonOut_({ ok: false, error: 'invalid caseId' }, 400);
  }

  statusApi_ensureActiveCaseIdString_(lineId, normCaseId);

  var userKey = '';
  try {
    if (typeof drive_userKeyFromLineId_ === 'function') userKey = drive_userKeyFromLineId_(lineId) || '';
  } catch (_) {}
  if (!userKey && typeof drive_lookupCaseRow_ === 'function') {
    try {
      var caseRow = drive_lookupCaseRow_({ caseId: normCaseId, lineId: lineId });
      if (caseRow && caseRow.userKey) userKey = String(caseRow.userKey || '').trim();
    } catch (_) {}
  }

  var caseKey = userKey ? userKey + '-' + normCaseId : '';
  var folderEnsured = false;
  if (caseKey && typeof drive_getOrCreateCaseFolderByKey_ === 'function') {
    try {
      var folder = drive_getOrCreateCaseFolderByKey_(caseKey);
      var folderId = '';
      if (folder && typeof folder.getId === 'function') folderId = folder.getId();
      else if (folder) folderId = String(folder);
      if (folderId && typeof updateCasesRow_ === 'function') {
        try {
          updateCasesRow_(normCaseId, { case_key: caseKey, folder_id: folderId, status: 'intake', updated_at: new Date().toISOString() });
          folderEnsured = true;
        } catch (_) {}
      }
    } catch (_) {}
  }
  if (!folderEnsured && typeof updateCasesRow_ === 'function' && caseKey) {
    try {
      updateCasesRow_(normCaseId, { case_key: caseKey, user_key: userKey || '', updated_at: new Date().toISOString() });
    } catch (_) {}
  }

  var submissionId = String(p.submission_id || p.submissionId || '').trim();
  if (!submissionId) submissionId = 'ack:' + normCaseId + ':intake';
  var existingAck = false;
  try {
    if (typeof sheetsRepo_hasSubmission_ === 'function') {
      existingAck = sheetsRepo_hasSubmission_(normCaseId, 'intake', submissionId);
    }
  } catch (_) {}

  var nextSeq = 1;
  if (!existingAck) {
    try {
      if (typeof getLastSeq_ === 'function') {
        nextSeq = (getLastSeq_(normCaseId, 'intake') | 0) + 1;
        if (!(nextSeq > 0)) nextSeq = 1;
      }
    } catch (_) {
      nextSeq = 1;
    }
  }
  var nowIso = new Date().toISOString();

  if (typeof upsertSubmission_ === 'function') {
    var payloadAck = {
      submission_id: submissionId,
      form_key: 'intake',
      case_id: normCaseId,
      user_key: userKey,
      case_key: caseKey,
      line_id: lineId,
      status: 'received',
      submitted_at: nowIso,
      received_at: nowIso,
      referrer: 'intake_ack',
      redirect_url: '',
      reopened_at: '',
      reopened_at_epoch: '',
      reopen_until: '',
      reopen_until_epoch: '',
      locked_reason: '',
      can_edit: false,
      reopened_by: '',
    };
    if (!existingAck) {
      payloadAck.seq = nextSeq;
      if (nextSeq > 1) payloadAck.supersedes_seq = String(nextSeq - 1);
    }
    upsertSubmission_(payloadAck);
  }

  ensureSubmissionsCaseIdTextFormat_();

  try {
    CacheService.getScriptCache().remove('status:' + normCaseId);
    try { Logger.log('[status:cache] remove caseId=%s', normCaseId); } catch (_) {}
  } catch (_) {}

  try { Logger.log('[ack] type=intake case=%s submission=%s status=received seq=%s', normCaseId, submissionId, existingAck ? 'kept' : String(nextSeq)); } catch (_) {}

  return statusApi_jsonOut_({ ok: true, caseId: normCaseId, case_id: normCaseId }, 200);
}

function statusApi_handleFormAck_(params) {
  var p = params || {};
  var formKey = String(p.form_key || p.formKey || '').trim();
  if (!formKey) return statusApi_jsonOut_({ ok: false, error: 'missing form_key' }, 400);

  var v2 = statusApi_verifyV2_(p);
  var lineId = String(p.lineId || p.line_id || '').trim();
  var caseIdHint = String(p.caseId || p.case_id || '').trim();
  var tsRaw = String(p.ts || '').trim();
  var sigRaw = String(p.sig || '').trim();
  if (v2 && v2.ok) {
    if (!lineId) lineId = String(v2.lineId || '').trim();
    if (!caseIdHint) caseIdHint = String(v2.caseId || '').trim();
    if (!tsRaw) tsRaw = String(v2.ts || '').trim();
  } else {
    try {
      statusApi_verify_(p);
    } catch (err) {
      return statusApi_jsonOut_({ ok: false, error: String(err) }, 400);
    }
  }

  if (!sigRaw) sigRaw = String(p.sig || '').trim();
  if (!lineId) return statusApi_jsonOut_({ ok: false, error: 'missing lineId' }, 400);
  if (!tsRaw) return statusApi_jsonOut_({ ok: false, error: 'missing ts' }, 400);
  if (!sigRaw) return statusApi_jsonOut_({ ok: false, error: 'missing sig' }, 400);

  try {
    statusApi_assertNonce_(lineId, tsRaw, sigRaw);
  } catch (err) {
    return statusApi_jsonOut_({ ok: false, error: String(err) }, 409);
  }

  var caseId = caseIdHint;
  var paramCaseIdNorm = statusApi_normCaseId_(String(p.caseId || p.case_id || ''));
  if (!caseId && paramCaseIdNorm) caseId = paramCaseIdNorm;
  if (!caseId) caseId = lookupCaseIdByLineId_(lineId);
  if (!caseId) return statusApi_jsonOut_({ ok: false, error: 'case_not_found' }, 404);
  var normCaseId = statusApi_normCaseId_(caseId);
  if (!normCaseId) return statusApi_jsonOut_({ ok: false, error: 'invalid caseId' }, 400);

  statusApi_ensureActiveCaseIdString_(lineId, normCaseId);

  var userKey = '';
  try {
    if (typeof drive_userKeyFromLineId_ === 'function') userKey = drive_userKeyFromLineId_(lineId) || '';
  } catch (_) {}
  if (!userKey && typeof drive_lookupCaseRow_ === 'function') {
    try {
      var caseRow = drive_lookupCaseRow_({ caseId: normCaseId, lineId: lineId });
      if (caseRow && caseRow.userKey) userKey = String(caseRow.userKey || '').trim();
    } catch (_) {}
  }

  var caseKey = userKey ? userKey + '-' + normCaseId : '';
  var submissionId = String(p.submission_id || p.submissionId || '').trim();
  if (!submissionId) submissionId = 'ack:' + normCaseId + ':' + formKey;
  var existingAck = false;
  try {
    if (typeof sheetsRepo_hasSubmission_ === 'function') {
      existingAck = sheetsRepo_hasSubmission_(normCaseId, formKey, submissionId);
    }
  } catch (_) {}

  var nextSeq = 1;
  if (!existingAck) {
    try {
      if (typeof getLastSeq_ === 'function') {
        nextSeq = (getLastSeq_(normCaseId, formKey) | 0) + 1;
        if (!(nextSeq > 0)) nextSeq = 1;
      }
    } catch (_) {
      nextSeq = 1;
    }
  }
  var nowIso = new Date().toISOString();
  if (typeof upsertSubmission_ === 'function') {
    var payloadAck = {
      submission_id: submissionId,
      form_key: formKey,
      case_id: normCaseId,
      user_key: userKey,
      case_key: caseKey,
      line_id: lineId,
      status: 'received',
      submitted_at: nowIso,
      received_at: nowIso,
      referrer: 'form_ack',
      redirect_url: '',
      reopened_at: '',
      reopened_at_epoch: '',
      reopen_until: '',
      reopen_until_epoch: '',
      locked_reason: '',
      can_edit: false,
      reopened_by: '',
    };
    if (!existingAck) {
      payloadAck.seq = nextSeq;
      if (nextSeq > 1) payloadAck.supersedes_seq = String(nextSeq - 1);
    }
    upsertSubmission_(payloadAck);
  }

  ensureSubmissionsCaseIdTextFormat_();

  try {
    CacheService.getScriptCache().remove('status:' + normCaseId);
    try { Logger.log('[status:cache] remove caseId=%s', normCaseId); } catch (_) {}
  } catch (_) {}

  try { Logger.log('[ack] type=form case=%s form=%s submission=%s status=received seq=%s', normCaseId, formKey, submissionId, existingAck ? 'kept' : String(nextSeq)); } catch (_) {}

  return statusApi_jsonOut_({ ok: true, caseId: normCaseId, case_id: normCaseId }, 200);
}

function statusApi_handleMarkReopenPost_(body) {
  var params = body || {};
  var v2 = statusApi_verifyV2_(params);
  var lineId = String(params.lineId || params.line_id || '').trim();
  var caseIdHint = String(params.caseId || params.case_id || '').trim();
  var tsRaw = String(params.ts || '').trim();
  var sig = String(params.sig || '').trim();

  if (v2 && v2.ok) {
    if (!lineId) lineId = String(v2.lineId || '').trim();
    if (!caseIdHint) caseIdHint = String(v2.caseId || '').trim();
    if (!tsRaw) tsRaw = String(v2.ts || '').trim();
  } else {
    try {
      statusApi_verify_(params || {});
    } catch (err) {
      return statusApi_jsonOut_({ ok: false, error: String(err) }, 400);
    }
  }

  if (!sig) sig = String(params.sig || '').trim();

  try {
    statusApi_assertNonce_(lineId, tsRaw, sig);
  } catch (err) {
    return statusApi_jsonOut_({ ok: false, error: String(err) }, 409);
  }

  if (!lineId) {
    return statusApi_jsonOut_({ ok: false, error: 'missing lineId' }, 400);
  }

  if (!tsRaw) {
    return statusApi_jsonOut_({ ok: false, error: 'missing ts' }, 400);
  }
  if (!sig) {
    return statusApi_jsonOut_({ ok: false, error: 'missing sig' }, 400);
  }

  const formKey = String((body || {}).form_key || (body || {}).formKey || '').trim();
  if (!formKey) {
    return statusApi_jsonOut_({ ok: false, error: 'missing form_key' }, 400);
  }

  var caseId = caseIdHint;
  if (!caseId) caseId = lookupCaseIdByLineId_(lineId);
  if (!caseId) {
    return statusApi_jsonOut_({ ok: false, error: 'case_not_found' }, 404);
  }
  var normCaseId = statusApi_normCaseId_(caseId);
  var allowSweepOnStatus = (STATUS_API_PROPS.getProperty('ALLOW_STAGING_SWEEP_ON_STATUS') || '').trim() === '1';
  if (allowSweepOnStatus) statusApi_collectStaging_(lineId, normCaseId);

  var now = new Date();
  var reopenedAt = now.toISOString();
  var reopenUntil = String((body || {}).reopen_until || (body || {}).reopenUntil || '').trim();
  var lockedReason = String((body || {}).locked_reason || (body || {}).lockedReason || '').trim();
  var reopeningStaff = ((body || {}).staff || 'staff').toString();
  var userKeyForLine = '';
  try {
    if (typeof drive_userKeyFromLineId_ === 'function') {
      userKeyForLine = drive_userKeyFromLineId_(lineId) || '';
    }
  } catch (_) {}
  var caseKey = userKeyForLine ? userKeyForLine + '-' + normCaseId : '';
  var nextSeq = 1;
  try {
    if (typeof getLastSeq_ === 'function') {
      nextSeq = (getLastSeq_(normCaseId, formKey) | 0) + 1;
      if (!(nextSeq > 0)) nextSeq = 1;
    }
  } catch (_) {
    nextSeq = 1;
  }
  var reopenedAtEpoch = Math.floor(now.getTime() / 1000);
  var reopenUntilEpoch = null;
  if (reopenUntil) {
    var parsedUntil = Date.parse(reopenUntil);
    if (Number.isFinite(parsedUntil)) reopenUntilEpoch = Math.floor(parsedUntil / 1000);
  }
  var reopenSubmissionId = 'reopen:' + formKey;
  if (typeof upsertSubmission_ === 'function') {
    upsertSubmission_({
      submission_id: reopenSubmissionId,
      form_key: formKey,
      case_id: normCaseId,
      user_key: userKeyForLine,
      case_key: caseKey,
      line_id: lineId,
      seq: nextSeq,
      supersedes_seq: nextSeq > 1 ? String(nextSeq - 1) : '',
      status: 'reopened',
      submitted_at: reopenedAt,
      received_at: reopenedAt,
      referrer: '',
      redirect_url: '',
      reopened_at: reopenedAt,
      reopened_at_epoch: Number.isFinite(reopenedAtEpoch) ? reopenedAtEpoch : '',
      reopen_until: reopenUntil,
      reopen_until_epoch: reopenUntilEpoch != null ? reopenUntilEpoch : '',
      locked_reason: lockedReason,
      can_edit: true,
      reopened_by: reopeningStaff,
    });
  }

  ensureSubmissionsCaseIdTextFormat_();

  try {
    CacheService.getScriptCache().remove('status:' + normCaseId);
  } catch (_) {}

  return statusApi_jsonOut_({ ok: true, caseId: normCaseId, case_id: normCaseId }, 200);
}

/**
 * JSON 受領後に呼び出し、submissions / cases_forms を更新
 * @param {{caseId:string, form_key:string, submission_id?:string, json_path?:string, received_at?:Date, locked_reason?:string, meta?:Object}} payload
 */
function recordSubmission_(payload) {
  if (!payload) return;
  const caseIdRaw = payload.case_id || payload.caseId;
  const formKeyRaw = payload.form_key;
  if (!caseIdRaw || !formKeyRaw) return;
  const caseIdNorm = statusApi_normCaseId_(caseIdRaw);
  const formKey = String(formKeyRaw).trim();
  if (!caseIdNorm || !formKey) return;

  const submissionId = String(payload.submission_id || '').trim();
  const fallbackInfo = {
    caseKey: payload.case_key || payload.caseKey,
    caseId: caseIdNorm,
    userKey: payload.user_key || payload.userKey,
    lineId: payload.line_id || payload.lineId,
  };
  var ukCandidate = '';
  let caseKey = '';
  try {
    caseKey = drive_resolveCaseKeyFromMeta_(payload.meta || {}, fallbackInfo);
  } catch (_) {
    var metaLineId = fallbackInfo.lineId || '';
    if (!metaLineId && payload.meta && payload.meta.line_id) metaLineId = payload.meta.line_id;
    if (!metaLineId && payload.meta && payload.meta.lineId) metaLineId = payload.meta.lineId;
    const row =
      typeof drive_lookupCaseRow_ === 'function'
        ? drive_lookupCaseRow_({ caseId: caseIdNorm, lineId: metaLineId })
        : null;
    if (row && row.userKey) {
      fallbackInfo.userKey = row.userKey;
      caseKey = row.userKey + '-' + caseIdNorm;
    }
  }
  if (!caseKey && fallbackInfo.userKey) {
    caseKey = fallbackInfo.userKey + '-' + caseIdNorm;
  }
  if (!caseKey && fallbackInfo.lineId) {
    const inferredKey =
      typeof drive_userKeyFromLineId_ === 'function'
        ? drive_userKeyFromLineId_(fallbackInfo.lineId)
        : '';
    if (inferredKey) caseKey = inferredKey + '-' + caseIdNorm;
  }

  if (caseKey && caseKey.indexOf('-') >= 0) {
    const parts = caseKey.split('-');
    const head = parts[0] || '';
    caseKey = head + '-' + caseIdNorm;
  }

  var normalizedUserKey = normalizeUserKey_(payload.user_key || payload.userKey || '');
  if (normalizedUserKey) fallbackInfo.userKey = normalizedUserKey;
  payload.user_key = normalizedUserKey;

  // 既存フォルダの参照のみ（作成しない）
  let caseFolderId = '';
  try {
    var userKeyFromKey = '';
    if (caseKey && caseKey.indexOf('-') >= 0) {
      userKeyFromKey = String(caseKey.split('-')[0] || '').trim();
    }
    ukCandidate = userKeyFromKey || fallbackInfo.userKey || '';
    if (!ukCandidate && fallbackInfo.lineId) ukCandidate = drive_userKeyFromLineId_(fallbackInfo.lineId);
    if (ukCandidate && typeof bs_resolveCaseFolderId_ === 'function') {
      caseFolderId = bs_resolveCaseFolderId_(ukCandidate, caseIdNorm, fallbackInfo.lineId || '');
    }
    if (!caseFolderId && typeof resolveCaseFolderId_ === 'function') {
      caseFolderId = resolveCaseFolderId_(fallbackInfo.lineId || '', caseIdNorm, /*createIfMissing=*/false);
    }
  } catch (_) {}
  if (caseFolderId && typeof updateCasesRow_ === 'function') {
    try {
      const patch = { case_key: caseKey || '', folder_id: caseFolderId };
      if (!patch.case_key && caseFolderId && ukCandidate) patch.case_key = ukCandidate + '-' + caseIdNorm;
      if (!fallbackInfo.userKey && patch.case_key.indexOf('-') >= 0) {
        patch.user_key = patch.case_key.split('-')[0];
      }
      updateCasesRow_(caseIdNorm, patch);
    } catch (_) {}
  }
  try {
    if (
      submissionId &&
      typeof sheetsRepo_hasSubmission_ === 'function' &&
      sheetsRepo_hasSubmission_(caseIdNorm, formKey, submissionId)
    ) {
      return;
    }
  } catch (err) {
    Logger.log('[Intake] recordSubmission_ duplicate check error: %s', (err && err.stack) || err);
  }

  let last = 0;
  try {
    if (typeof getLastSeq_ === 'function') last = Number(getLastSeq_(caseIdNorm, formKey)) || 0;
  } catch (err) {
    last = 0;
  }
  const receivedAt = payload.received_at
    ? new Date(payload.received_at).toISOString()
    : new Date().toISOString();
  const submittedAt = payload.submitted_at
    ? new Date(payload.submitted_at).toISOString()
    : receivedAt;
  if (typeof upsertSubmission_ === 'function') {
    let statusValue = '';
    if (payload.status != null) statusValue = String(payload.status).trim();
    if (!statusValue || statusValue.toLowerCase() === formKey.toLowerCase()) {
      statusValue = 'received';
    }
    upsertSubmission_({
      case_id: caseIdNorm,
      case_key: caseKey,
      form_key: formKey,
      seq: (payload.seq != null && payload.seq !== '') ? Number(payload.seq) : (last + 1),
      submission_id: submissionId,
      received_at: receivedAt,
      supersedes_seq: last || '',
      json_path: payload.json_path || '',
      // 追加カラム（あれば埋める）
      user_key: normalizedUserKey,
      line_id: payload.line_id || '',
      submitted_at: submittedAt,
      status: statusValue,
      referrer: (payload && payload.meta && payload.meta.referrer) || payload.referrer || '',
      redirect_url: (payload && payload.meta && payload.meta.redirect_url) || payload.redirect_url || '',
    });
  }

  try {
    CacheService.getScriptCache().remove('status:' + caseIdNorm);
    try { Logger.log('[status:cache] remove caseId=%s', caseIdNorm); } catch (_) {}
  } catch (_) {}

  if (submissionId && !/^ack:/i.test(submissionId) && typeof sheetsRepo_deleteAckRow_ === 'function') {
    try { sheetsRepo_deleteAckRow_(caseIdNorm, formKey); } catch (_) {}
  }
  if (typeof sheetsRepo_sweepSubmissions_ === 'function') {
    try { sheetsRepo_sweepSubmissions_(); } catch (_) {}
  }
}

/** POST /exec でのルーティング */
function statusApi_doPost_(e) {
  var action = String((e && e.parameter && e.parameter.action) || '').trim();
  var body = {};
  try {
    body = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
  } catch (_) {}
  if (!action) action = String((body && body.action) || '').trim();

  if (action === 'markReopen') {
    return statusApi_handleMarkReopenPost_(body);
  }
  return statusApi_jsonOut_({ ok: false, error: 'unknown action' }, 400);
}

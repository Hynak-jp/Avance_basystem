/**
 * bootstrap.js (BAS / Google Apps Script)
 *  - 認証方式: URL クエリ (?sig, ?ts)
 *  - 署名対象: lineId + '|' + ts（UNIX秒）
 *  - アルゴリズム: HMAC-SHA256 → Base64URL
 *  - ts: UNIX秒 / 許容スキュー ±300秒 / 本文の ts とクエリ ts を整合確認
 *  - 処理: contacts upsert → caseId 採番（ロック）→ Drive フォルダ作成 → cases/contacts 更新
 */

/** ---------- 設定 ---------- **/
// ★ グローバル PROPS は使わない（重複定義ガード）
if (typeof props_ !== 'function') {
  var props_ = function () {
    return PropertiesService.getScriptProperties();
  };
}

if (typeof getSecret_ !== 'function') {
  var getSecret_ = function () {
    var s = props_().getProperty('BOOTSTRAP_SECRET') || props_().getProperty('TOKEN_SECRET') || '';
    if (s && typeof s.replace === 'function') s = s.replace(/[\r\n]+$/g, '');
    if (!s) throw new Error('missing secret');
    return s;
  };
}

// フォールバック: userKey 推定（lineId 先頭6小文字）
if (typeof drive_userKeyFromLineId_ !== 'function') {
  var drive_userKeyFromLineId_ = function (lineId) {
    return String(lineId || '').slice(0, 6).toLowerCase();
  };
}

// 追加：タイミング安全な比較
function safeCompare_(a, b) {
  a = String(a || '');
  b = String(b || '');
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}
function b64url_(s) {
  return String(s || '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
}

// Avoid global name collision across GAS files: use a file‑scoped alias
const BS_MASTER_SPREADSHEET_ID = props_().getProperty('BAS_MASTER_SPREADSHEET_ID'); // BAS_master
// 両対応: DRIVE_ROOT_FOLDER_ID / ROOT_FOLDER_ID
const DRIVE_ROOT_ID =
  props_().getProperty('DRIVE_ROOT_FOLDER_ID') || props_().getProperty('ROOT_FOLDER_ID'); // BAS_提出書類 ルート
const LOG_SHEET_ID = props_().getProperty('GAS_LOG_SHEET_ID') || ''; // 任意（空で無効）
const ALLOW_DEBUG = (props_().getProperty('ALLOW_DEBUG') || '').toLowerCase() === '1'; //

const SHEET_CONTACTS = 'contacts';
const SHEET_CASES = 'cases';

function bs_normCaseId_(raw) {
  const digits = String(raw == null ? '' : raw)
    .trim()
    .replace(/\D/g, '');
  if (!digits) return '';
  return digits.padStart(4, '0');
}

function bs_caseFolderName_(userKey, rawCaseId) {
  const cid = bs_normCaseId_(rawCaseId);
  if (!cid) return `${userKey}-0000`;
  return `${userKey}-${cid}`;
}

/** ===== ログユーティリティ（安全に短く） ===== */

/** ========= 1) 先頭：共通ユーティリティ（ファイルの上のほうに） ========= */
function redact_(s, show = 4) {
  if (!s) return '';
  const str = String(s);
  if (str.length <= show * 2) return '*'.repeat(str.length);
  return str.slice(0, show) + '...' + str.slice(-show);
}
function logCtx_(label, obj) {
  try {
    const safe = JSON.parse(
      JSON.stringify(obj, (k, v) => {
        if (k === 'sig' || k === 'token' || k === 'secret') return redact_(v);
        if (k === 'p') return '[base64 payload omitted]';
        return v;
      })
    );
    console.info(`BAS:${label}`, safe);
  } catch (e) {
    console.info(`BAS:${label}:<unserializable>`);
  }
}
function normalizeQuery_(obj) {
  const out = {};
  for (const [k, v] of Object.entries(obj || {})) {
    let key = (k || '').toString();
    key = key.replace(/[A-Z]/g, (m) => '_' + m.toLowerCase()).toLowerCase(); // camel→snake
    if (key === 'lineid') key = 'line_id';
    if (key === 'caseid') key = 'case_id';
    out[key] = v;
  }
  return out;
}
function makeCanonicalPayload_(line_id, case_id, ts) {
  return [String(line_id || ''), String(case_id || ''), String(ts || '')].join('|');
}

// base64url の表記揺れだけ吸収してログするための補助（比較には使わない）
function __b64peek(s) {
  s = String(s || '');
  return {
    len: s.length,
    head: s.slice(0, 6),
    tail: s.slice(-6),
  };
}

function verifySigV2_(line_id, case_id, ts, sig) {
  const payload = makeCanonicalPayload_(line_id, case_id, ts);
  const secret = getSecret_();
  const mac = Utilities.computeHmacSha256Signature(payload, secret);
  const expectedRaw = Utilities.base64EncodeWebSafe(mac);
  const expected = b64url_(expectedRaw);
  const provided = b64url_(sig);
  const ok = safeCompare_(expected, provided);
  logCtx_('sig:verify', {
    ok,
    payload_preview: payload.slice(0, 40) + (payload.length > 40 ? '...' : ''),
    ts: ts,
    exp_len: expected.length,
    got_len: provided.length,
    exp_head: expected.slice(0, 6),
    exp_tail: expected.slice(-6),
    got_head: provided.slice(0, 6),
    got_tail: provided.slice(-6),
  });
  return ok;
}
function parseAuth_(q) {
  if (q && q.p) {
    const decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(q.p)).getDataAsString();
    const parts = decoded.split('|');
    const li = parts[0] || '';
    const ci = parts[1] || '';
    const t = parts[2] || '';
    const ts = q.ts || t; // フォールバック
    if (q.ts && t !== q.ts) {
      logCtx_('sig:ts-mismatch', { q_ts: q.ts, p_ts: t });
      return { ok: false, line_id: '', case_id: '' };
    }
    const ok = verifySigV2_(li, ci, ts, q.sig);
    logCtx_('sig:p-decoded', {
      has_p: true,
      t_len: String(t).length,
      ts_len: String(ts).length,
      payload_preview: decoded.slice(0, 40) + (decoded.length > 40 ? '...' : ''),
    });
    return { ok, line_id: li, case_id: ci || '' };
  } else {
    logCtx_('sig:p-decoded', { has_p: false });
    return { ok: false, line_id: '', case_id: '' };
  }
}

/** ========= 2) 入口関数：bootstrap_ / doGet の最初に差し込む ========= */
function bootstrap_(e) {
  // [LOG-1] 入口（生 and 正規化後）
  logCtx_('bootstrap:raw', { query: e && e.parameter });
  const q = normalizeQuery_(e && e.parameter);
  logCtx_('bootstrap:norm', {
    keys: Object.keys(q),
    has_p: !!q.p,
    ts_len: String(q.ts || '').length,
    sig_len: String(q.sig || '').length,
    line_id: !!q.line_id,
    case_id: !!q.case_id,
  });

  // C. p を使っているなら “中身の形” だけ観測
  try {
    if (q.p) {
      var decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(q.p)).getDataAsString();
      var parts = decoded.split('|');
      var t = parts[2] || '';
      logCtx_('sig:p-decoded', {
        has_p: true,
        parts_len: parts.length,
        t_len: String(t).length,
        payload_preview: decoded.slice(0, 40) + (decoded.length > 40 ? '…' : ''),
      });
    } else {
      logCtx_('sig:p-decoded', { has_p: false });
    }
  } catch (err) {
    logCtx_('sig:p-decode-error', { error: String(err) });
  }

  // [LOG-2] 署名検証（p有無どちらでも観測）
  const auth = parseAuth_(q);
  logCtx_('bootstrap:auth', { ok: auth.ok, li_present: !!auth.line_id, ci_present: !!auth.case_id });
  if (!auth.ok) return bs_jsonResponse_({ ok: false, error: 'invalid sig' }, 400);

  // [LOG-3] マスタ/Drive の存在確認
  try { bs_ensureMaster_(); } catch (err) { return bs_jsonResponse_({ ok:false, error: String(err) }, 500); }
  try { bs_ensureDriveRoot_(); } catch (err) { return bs_jsonResponse_({ ok:false, error: String(err) }, 500); }

  // [LOG-4] contacts upsert（userKeyは lineId から推定）
  const lineId = String(auth.line_id || '').trim();
  const userKey = drive_userKeyFromLineId_(lineId);
  const up = bs_upsertContact_({ userKey, lineId, displayName: '', email: '' });
  logCtx_('contacts:upsert:done', { row: up.row, userKey });

  // [LOG-5] caseId 確保（既存が無ければ採番）
  const sh = bs_getSheet_(SHEET_CONTACTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = bs_toIndexMap_(headers);
  const rowVals = sh.getRange(up.row, 1, 1, headers.length).getValues()[0];
  const activeIdx = typeof idx['active_case_id'] === 'number' ? idx['active_case_id'] : idx['activeCaseId'];
  let caseId = activeIdx >= 0 ? bs_normCaseId_(rowVals[activeIdx] || '') : '';
  if (!caseId) {
    const issued = bs_issueCaseId_(userKey, lineId);
    caseId = bs_normCaseId_(issued.caseId);
    if (activeIdx >= 0) sh.getRange(up.row, activeIdx + 1).setValue(caseId);
  } else {
    // 既存 caseId の場合でも、cases 行が無ければ作っておく
    try { bs_ensureCaseRow_(caseId, userKey, lineId); } catch (_) {}
  }

  // [LOG-6] ケースフォルダは intake 完了時にのみ作成する（ここでは作らない）
  const resp = { ok: true, case_id: caseId, caseFolderReady: false };
  logCtx_('bootstrap:resp', resp);
  return bs_jsonResponse_(resp, 200);
}

/** ---------- ユーティリティ ---------- **/
function bs_jsonResponse_(obj, status) {
  const out = ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
  if (status && out.setStatusCode) out.setStatusCode(status);
  return out;
}

function bs_appendLog_(arr) {
  if (!LOG_SHEET_ID) return;
  try {
    const ss = bs_openSpreadsheet_(LOG_SHEET_ID);
    const sh = ss.getSheetByName('logs') || ss.insertSheet('logs');
    sh.appendRow(arr);
  } catch (_) {}
}

function bs_hmacBase64FromRaw_(raw, secret) {
  // Apps ScriptのHMACはbyte配列が無難。rawは文字列、secretも文字列でOK。
  return Utilities.base64Encode(
    Utilities.computeHmacSha256Signature(
      Utilities.newBlob(raw).getBytes(),
      Utilities.newBlob(secret).getBytes()
    )
  );
}

function bs_openSpreadsheet_(id) {
  const spreadsheet = SpreadsheetApp.openById(id);
  if (!spreadsheet || typeof spreadsheet.getSheetByName !== 'function') {
    throw new Error('invalid_spreadsheet_id: ' + id);
  }
  return spreadsheet;
}

// 念のための型ガード（Sheet を渡されても親を辿って Spreadsheet を返す）
function bs_ensureSpreadsheet_(obj) {
  if (obj && typeof obj.getSheetByName === 'function') return obj; // Spreadsheet
  if (obj && typeof obj.getParent === 'function') {
    const p = obj.getParent && obj.getParent();
    if (p && typeof p.getSheetByName === 'function') return p;
  }
  throw new Error('not_a_spreadsheet');
}

function bs_driveListFilesV3_(options) {
  if (!Drive || !Drive.Files || typeof Drive.Files.list !== 'function') {
    throw new Error('Drive.Files.list unavailable');
  }
  var base = {
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
    fields: 'files(id,name,parents)',
    corpora: 'allDrives',
  };
  var payload = Object.assign({}, base, options || {});
  if ('driveId' in payload) delete payload.driveId;
  if (!payload.pageSize && !payload.pageToken) payload.pageSize = 10;
  var res = Drive.Files.list(payload);
  return (res && res.files) || [];
}

function bs_driveCreateFolderV3_(resource) {
  if (
    typeof Drive === 'undefined' ||
    !Drive ||
    !Drive.Files ||
    typeof Drive.Files.create !== 'function'
  ) {
    throw new Error('Drive.Files.create unavailable');
  }
  var meta = Object.assign({}, resource || {});
  if (!meta.mimeType) meta.mimeType = 'application/vnd.google-apps.folder';
  var params = { supportsAllDrives: true };
  return Drive.Files.create(meta, null, params);
}

function bs_getSheet_(name) {
  const ss = bs_ensureSpreadsheet_(bs_openSpreadsheet_(BS_MASTER_SPREADSHEET_ID)); // Spreadsheet
  let sh = ss.getSheetByName(name); // Sheet
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/** cases 行を保証（無ければ追記して row# を返す） */
function bs_ensureCaseRow_(caseId, userKey, lineId) {
  const sh = bs_getSheet_(SHEET_CASES);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 6).setValues([[
      'case_id', 'user_key', 'line_id', 'status', 'folder_id', 'created_at',
    ]]);
  }
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = bs_toIndexMap_(headers);
  const caseIdx = idx['case_id'] != null ? idx['case_id'] : idx['caseId'];
  const userIdx = idx['user_key'] != null ? idx['user_key'] : idx['userKey'];
  const lineIdx = idx['line_id']  != null ? idx['line_id']  : idx['lineId'];
  const statusIdx = idx['status'];
  const folderIdx = idx['folder_id'] != null ? idx['folder_id'] : idx['folderId'];
  const createdIdx = idx['created_at'] != null ? idx['created_at'] : idx['createdAt'];
  const rc = sh.getLastRow() - 1;
  if (rc > 0 && caseIdx != null) {
    const rows = sh.getRange(2, 1, rc, headers.length).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      if (bs_normCaseId_(rows[i][caseIdx]) === bs_normCaseId_(caseId)) return i + 2;
    }
  }
  const now = new Date().toISOString();
  const row = new Array(headers.length).fill('');
  if (caseIdx   != null) row[caseIdx]   = bs_normCaseId_(caseId);
  if (userIdx   != null) row[userIdx]   = userKey || '';
  if (lineIdx   != null) row[lineIdx]   = lineId  || '';
  if (statusIdx != null) row[statusIdx] = 'draft';
  if (createdIdx!= null) row[createdIdx]= now;
  sh.appendRow(row);
  return sh.getLastRow();
}

function bs_headerAliases_(raw) {
  const value = String(raw == null ? '' : raw).trim();
  if (!value) return [];
  const normalizedSpace = value.replace(/[\s\u3000]+/g, ' ').trim();
  const lower = normalizedSpace.toLowerCase();
  const snake = lower
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
  const camel = snake.replace(/_([a-z0-9])/g, function (_, c) {
    return c.toUpperCase();
  });
  const pascal = camel ? camel.charAt(0).toUpperCase() + camel.slice(1) : '';
  const flat = snake.replace(/_/g, '');
  const aliases = new Set([
    value,
    normalizedSpace,
    lower,
    snake,
    camel,
    pascal,
    flat,
    flat.toLowerCase(),
  ]);
  return Array.from(aliases).filter(function (k) {
    return !!k;
  });
}

function bs_toIndexMap_(headers) {
  const m = {};
  (headers || []).forEach(function (h, i) {
    const aliases = bs_headerAliases_(h);
    aliases.forEach(function (key) {
      if (!(key in m)) m[key] = i;
    });
  });
  return m;
}

/** ---------- 初期化（存在チェック＆自動作成） ---------- **/
function bs_ensureMaster_() {
  const sid = BS_MASTER_SPREADSHEET_ID;
  if (!sid) throw new Error('BAS_MASTER_SPREADSHEET_ID is empty');
  const ss = bs_ensureSpreadsheet_(bs_openSpreadsheet_(sid));
  const contacts = ss.getSheetByName(SHEET_CONTACTS) || ss.insertSheet(SHEET_CONTACTS);
  const cases = ss.getSheetByName(SHEET_CASES) || ss.insertSheet(SHEET_CASES);
  return { ss, contacts, cases };
}

function bs_ensureDriveRoot_() {
  const rid = DRIVE_ROOT_ID;
  if (!rid) throw new Error('DRIVE_ROOT_FOLDER_ID/ROOT_FOLDER_ID is empty');
  DriveApp.getFolderById(rid); // 存在チェック。権限が無ければここで例外。
}

/** ---------- contacts upsert ---------- **/
function bs_upsertContact_(payload) {
  const sh = bs_getSheet_(SHEET_CONTACTS);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 7).setValues([
      ['user_key', 'line_id', 'display_name', 'email', 'active_case_id', 'updated_at', 'intake_at'],
    ]);
  }
  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = bs_toIndexMap_(headers);
  const required = [
    'user_key',
    'line_id',
    'display_name',
    'email',
    'active_case_id',
    'updated_at',
    'intake_at',
  ];
  let changed = false;
  required.forEach(function (key) {
    if (!(key in idx)) {
      headers.push(key);
      idx[key] = headers.length - 1;
      changed = true;
    }
  });
  if (changed) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // 互換用 alias を再設定
  idx['userKey'] = idx['user_key'];
  idx['lineId'] = idx['line_id'];
  idx['displayName'] = idx['display_name'];
  idx['activeCaseId'] = idx['active_case_id'];
  idx['updatedAt'] = idx['updated_at'];
  if (typeof idx['intake_at'] === 'number') idx['intakeAt'] = idx['intake_at'];

  const alias = function (key, fallbacks) {
    const keys = [key].concat(fallbacks || []);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (typeof idx[k] === 'number') return idx[k];
    }
    return -1;
  };

  const user_key = payload.user_key ?? payload.userKey;
  const line_id = payload.line_id ?? payload.lineId;
  const display_name = payload.display_name ?? payload.displayName ?? '';
  const email = payload.email ?? '';
  if (!user_key) throw new Error('user_key required');
  if (!line_id) throw new Error('line_id required');

  const colUserKey = alias('user_key', ['userKey']);
  const colLineId = alias('line_id', ['lineId']);
  const colDisplayName = alias('display_name', ['displayName']);
  const colEmail = alias('email', []);
  const colActiveCaseId = alias('active_case_id', ['activeCaseId']);
  const colUpdatedAt = alias('updated_at', ['updatedAt']);
  const colIntakeAt = alias('intake_at', ['intakeAt']);

  const rowCount = sh.getLastRow() - 1;
  const rows = rowCount > 0 ? sh.getRange(2, 1, rowCount, headers.length).getValues() : [];
  let found = -1;
  if (colUserKey >= 0) {
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][colUserKey]) === String(user_key)) {
        found = i;
        break;
      }
    }
  }

  const now = new Date().toISOString();
  if (found >= 0) {
    const row = 2 + found;
    const current = sh.getRange(row, 1, 1, headers.length).getValues()[0];
    if (colLineId >= 0) current[colLineId] = line_id;
    if (colDisplayName >= 0 && display_name != null) current[colDisplayName] = display_name;
    if (colEmail >= 0 && email != null) current[colEmail] = email;
    if (colUpdatedAt >= 0) current[colUpdatedAt] = now;
    sh.getRange(row, 1, 1, headers.length).setValues([current]);
    return { row, headers, idx };
  } else {
    const values = new Array(headers.length).fill('');
    if (colUserKey >= 0) values[colUserKey] = user_key;
    if (colLineId >= 0) values[colLineId] = line_id;
    if (colDisplayName >= 0) values[colDisplayName] = display_name || '';
    if (colEmail >= 0) values[colEmail] = email || '';
    if (colActiveCaseId >= 0) values[colActiveCaseId] = '';
    if (colUpdatedAt >= 0) values[colUpdatedAt] = now;
    if (colIntakeAt >= 0) values[colIntakeAt] = '';
    sh.appendRow(values);
    return { row: sh.getLastRow(), headers, idx };
  }
}

function bs_setContactIntakeAt_(lineId, when) {
  const sh = bs_getSheet_(SHEET_CONTACTS);
  if (sh.getLastRow() < 1) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = bs_toIndexMap_(headers);
  const alias = function (key, fallbacks) {
    const keys = [key].concat(fallbacks || []);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (typeof idx[k] === 'number') return idx[k];
    }
    return -1;
  };
  const intakeIdx = alias('intake_at', ['intakeAt']);
  const lineIdx = alias('line_id', ['lineId']);
  if (intakeIdx < 0 || lineIdx < 0) return;
  const rowCount = sh.getLastRow() - 1;
  if (rowCount < 1) return;
  const rows = sh.getRange(2, 1, rowCount, headers.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][lineIdx]) === String(lineId)) {
      sh.getRange(i + 2, intakeIdx + 1).setValue(when || new Date().toISOString());
      return;
    }
  }
}

function bs_setCaseStatus_(caseId, status) {
  const sh = bs_getSheet_(SHEET_CASES);
  if (sh.getLastRow() < 1) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = bs_toIndexMap_(headers);
  const alias = function (key, fallbacks) {
    const keys = [key].concat(fallbacks || []);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (typeof idx[k] === 'number') return idx[k];
    }
    return -1;
  };
  const caseIdx = alias('case_id', ['caseId']);
  const statusIdx = alias('status', []);
  const lastIdx = alias('last_activity', ['lastActivity']);
  if (caseIdx < 0 || statusIdx < 0) return;
  const rowCount = sh.getLastRow() - 1;
  if (rowCount < 1) return;
  const rows = sh.getRange(2, 1, rowCount, headers.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][caseIdx]) === String(caseId)) {
      sh.getRange(i + 2, statusIdx + 1).setValue(status);
      if (lastIdx >= 0) {
        sh.getRange(i + 2, lastIdx + 1).setValue(new Date().toISOString());
      }
      return;
    }
  }
}

/** ---------- caseId 採番（ロック付き） ---------- **/
function bs_issueCaseId_(userKey, lineId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) throw new Error('lock_timeout: bs_issueCaseId_');
  try {
    const sh = bs_getSheet_(SHEET_CASES);
    if (sh.getLastRow() < 1) {
      sh.getRange(1, 1, 1, 6).setValues([
        ['case_id', 'user_key', 'line_id', 'status', 'folder_id', 'created_at'],
      ]);
    }
    let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const idx = bs_toIndexMap_(headers);
    const required = ['case_id', 'user_key', 'line_id', 'status', 'folder_id', 'created_at'];
    let changed = false;
    required.forEach(function (key) {
      if (!(key in idx)) {
        headers.push(key);
        idx[key] = headers.length - 1;
        changed = true;
      }
    });
    if (changed) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    idx['caseId'] = idx['case_id'];
    idx['userKey'] = idx['user_key'];
    idx['lineId'] = idx['line_id'];
    idx['folderId'] = idx['folder_id'];
    idx['createdAt'] = idx['created_at'];

    const alias = function (key, fallbacks) {
      const keys = [key].concat(fallbacks || []);
      for (var i = 0; i < keys.length; i++) {
        var k = keys[i];
        if (typeof idx[k] === 'number') return idx[k];
      }
      return -1;
    };

    const colCaseId = alias('case_id', ['caseId']);
    const colUserKey = alias('user_key', ['userKey']);
    const colLineId = alias('line_id', ['lineId']);
    const colStatus = alias('status', []);
    const colFolderId = alias('folder_id', ['folderId']);
    const colCreatedAt = alias('created_at', ['createdAt']);

    const rowCount = sh.getLastRow() - 1;
    const rows = rowCount > 0 ? sh.getRange(2, 1, rowCount, headers.length).getValues() : [];
    let maxNum = 0;
    if (colCaseId >= 0) {
      for (let i = 0; i < rows.length; i++) {
        const n = parseInt(String(rows[i][colCaseId] || ''), 10);
        if (Number.isFinite(n)) maxNum = Math.max(maxNum, n);
      }
    }
    const next = String(maxNum + 1).padStart(4, '0');
    const now = new Date().toISOString();

    const values = new Array(headers.length).fill('');
    if (colCaseId >= 0) values[colCaseId] = next;
    if (colUserKey >= 0) values[colUserKey] = userKey;
    if (colLineId >= 0) values[colLineId] = lineId;
    if (colStatus >= 0) values[colStatus] = 'draft';
    if (colFolderId >= 0) values[colFolderId] = '';
    if (colCreatedAt >= 0) values[colCreatedAt] = now;
    sh.appendRow(values);

    return { caseId: next, row: sh.getLastRow(), headers, idx };
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

/** ---------- Drive フォルダ作成 / 取得 ---------- **/
function bs_ensureCaseFolder_(userKey, caseId) {
  const rootId = DRIVE_ROOT_ID;
  const normalizedCaseId = bs_normCaseId_(caseId);
  if (!normalizedCaseId) throw new Error('invalid_case_id');
  const title = bs_caseFolderName_(userKey, normalizedCaseId);

  // DriveApp のみで検索→作成（v2/v3 差異の影響を避ける）
  const root = DriveApp.getFolderById(rootId);
  const it = root.getFoldersByName(title);
  if (it.hasNext()) return it.next().getId();
  return root.createFolder(title).getId();
}

/**
 * cases シートからフォルダ ID を取得（フォルダ新規作成はしない）
 */
function bs_resolveCaseFolderId_(userKey, caseId, lineId) {
  const normalizedCaseId = bs_normCaseId_(caseId);
  if (!normalizedCaseId) return '';
  const sh = bs_getSheet_(SHEET_CASES);
  if (sh.getLastRow() < 2) return '';
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = bs_toIndexMap_(headers);
  const alias = function (key, fallbacks) {
    const keys = [key].concat(fallbacks || []);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (typeof idx[k] === 'number') return idx[k];
    }
    return -1;
  };
  const colFolderId = alias('folder_id', ['folderId']);
  const colCaseId = alias('case_id', ['caseId']);
  const colUserKey = alias('user_key', ['userKey']);
  const colLineId = alias('line_id', ['lineId']);
  if (colFolderId < 0 || colCaseId < 0) return '';
  const rowCount = sh.getLastRow() - 1;
  const rows = sh.getRange(2, 1, rowCount, headers.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    const rowCaseId = bs_normCaseId_(rows[i][colCaseId] || '');
    if (rowCaseId !== normalizedCaseId) continue;
    if (colUserKey >= 0 && userKey) {
      if (String(rows[i][colUserKey] || '') !== String(userKey)) continue;
    } else if (colLineId >= 0 && lineId) {
      if (String(rows[i][colLineId] || '') !== String(lineId)) continue;
    }
    const folderId = String(rows[i][colFolderId] || '').trim();
    if (folderId) return folderId;
    return '';
  }
  return '';
}

/**
 * intake JSON が案件フォルダに存在するかを DriveApp で確認
 */
function bs_isIntakeJsonReady_(caseFolderId) {
  if (!caseFolderId) return false;
  const folder = DriveApp.getFolderById(caseFolderId);
  const it = folder.getFiles();
  while (it.hasNext()) {
    const file = it.next();
    const name = file.getName && file.getName();
    if (name && /^intake__/i.test(name)) return true;
  }
  return false;
}

function resolveCaseFolderId_(lineId, caseId) {
  const direct = bs_resolveCaseFolderId_('', caseId, lineId);
  if (direct) return direct;

  const userKey = lineId ? drive_userKeyFromLineId_(lineId) : '';
  if (userKey) {
    const viaUser = bs_resolveCaseFolderId_(userKey, caseId, lineId);
    if (viaUser) return viaUser;
    try {
      const ensuredId = bs_ensureCaseFolder_(userKey, caseId);
      if (ensuredId) return ensuredId;
    } catch (err) {
      try {
        Logger.log('[resolveCaseFolderId_] ensure error: %s', (err && err.stack) || err);
      } catch (_) {}
    }
  }

  return '';
}

function bs_collectIntakeFromStaging_(lineId, caseId) {
  const normalizedCaseId = bs_normCaseId_(caseId);
  if (!lineId || !normalizedCaseId) return;

  let caseFolderId = resolveCaseFolderId_(lineId, normalizedCaseId);
  if (!caseFolderId) return;

  let caseFolder;
  try {
    caseFolder = DriveApp.getFolderById(caseFolderId);
  } catch (err) {
    try {
      Logger.log('[bs_collectIntakeFromStaging_] case folder error: %s', (err && err.stack) || err);
    } catch (_) {}
    return;
  }

  try {
    const root = DriveApp.getFolderById(DRIVE_ROOT_ID);
    ['_staging', '_email_staging'].forEach(function (name) {
      const iter = root.getFoldersByName(name);
      if (!iter.hasNext()) return;
      moveAll(iter.next());
    });
  } catch (err) {
    try {
      Logger.log('[bs_collectIntakeFromStaging_] error: %s', (err && err.stack) || err);
    } catch (_) {}
  }

  function moveAll(folder) {
    const stack = [folder];
    while (stack.length) {
      const current = stack.pop();
      const files = current.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        if (file.getMimeType && file.getMimeType() === 'application/json') {
          const name = file.getName ? file.getName() : '';
          if (/^intake__/i.test(name)) {
            try {
              caseFolder.addFile(file);
              const parents = file.getParents();
              while (parents.hasNext()) {
                const parent = parents.next();
                if (parent.getId && parent.getId() !== caseFolderId) {
                  try {
                    parent.removeFile(file);
                  } catch (_) {}
                }
              }
              // 受領記録（submissions/cases_forms 更新）
              try {
                var text = '';
                var data = null;
                try {
                  text = file.getBlob().getDataAsString('utf-8');
                } catch (_e) {}
                try {
                  data = JSON.parse(text);
                } catch (_e2) {}
                var sid = '';
                if (data && typeof data === 'object') {
                  sid = String(
                    data.submission_id || (data.meta && data.meta.submission_id) || ''
                  ).trim();
                }
                if (typeof recordSubmission_ === 'function') {
                  recordSubmission_({
                    case_id: normalizedCaseId,
                    form_key: 'intake',
                    submission_id: sid,
                    json_path: file.getName(),
                    meta:
                      data && data.meta
                        ? data.meta
                        : { case_id: normalizedCaseId, line_id: lineId },
                  });
                }
                if (typeof bs_setCaseStatus_ === 'function') {
                  try {
                    bs_setCaseStatus_(normalizedCaseId, 'intake');
                  } catch (_) {}
                }
              } catch (e3) {
                try {
                  Logger.log(
                    '[bs_collectIntakeFromStaging_] recordSubmission error: %s',
                    (e3 && e3.stack) || e3
                  );
                } catch (_) {}
              }
            } catch (err) {
              try {
                Logger.log(
                  '[bs_collectIntakeFromStaging_] move error: %s',
                  (err && err.stack) || err
                );
              } catch (_) {}
            }
          }
        }
      }
      const children = current.getFolders();
      while (children.hasNext()) stack.push(children.next());
    }
  }
}

function handleStatus_(lineId, caseId, hasIntake, userKey) {
  const intakeFlag = !!hasIntake;
  const normalizedCaseId = bs_normCaseId_(caseId);
  if (!normalizedCaseId) {
    return ContentService.createTextOutput(
      JSON.stringify({ ok: true, caseId: '', intakeReady: false, hasIntake: intakeFlag })
    ).setMimeType(ContentService.MimeType.JSON);
  }
  let folderId = resolveCaseFolderId_(lineId, normalizedCaseId);
  if (!folderId && userKey) {
    try {
      folderId = bs_ensureCaseFolder_(userKey, normalizedCaseId);
    } catch (_) {}
  }
  const caseFolderReady = Boolean(folderId);
  if (!caseFolderReady) {
    try {
      Logger.log(
        '[handleStatus_] no folderId yet, userKey=%s caseId=%s',
        userKey || '',
        normalizedCaseId
      );
    } catch (_) {}
  }
  let intakeReady = folderId ? bs_isIntakeJsonReady_(folderId) : false;
  if (!intakeReady) {
    bs_collectIntakeFromStaging_(lineId, normalizedCaseId);
    if (!folderId && userKey) {
      try {
        folderId = bs_resolveCaseFolderId_(userKey, normalizedCaseId, lineId);
      } catch (_) {}
    }
    intakeReady = folderId ? bs_isIntakeJsonReady_(folderId) : false;
  }
  return ContentService.createTextOutput(
    JSON.stringify({
      ok: true,
      caseId: normalizedCaseId,
      intakeReady,
      caseFolderReady: caseFolderReady || Boolean(folderId),
      hasIntake: intakeFlag,
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

/** ---------- doPost ---------- **/
function doPost(e) {
  let ST = 'enter';
  try {
    const VER = 'dbg-final';
    const qs = e?.parameter || {};
    const body = e?.postData?.contents ? JSON.parse(e.postData.contents) : {};

    const action = String((body || {}).action || '').trim();
    const lineId = String((body || {}).lineId ?? (qs || {}).lineId ?? '').trim();
    const ts = Number(String((body || {}).ts ?? (qs || {}).ts ?? '0').trim());
    const providedSigRaw = String(
      (body || {}).sig ?? (qs || {}).sig ?? (body || {}).signature ?? (qs || {}).signature ?? ''
    ).trim();
    const providedSig = providedSigRaw.replace(/=+$/, '');

    // 確定デバッグ: ALLOW_DEBUG=1 のときのみ有効
    const wantDebug = String((qs || {}).debug ?? (body || {}).debug) === '1';
    if (wantDebug && !ALLOW_DEBUG) {
      return ContentService.createTextOutput(
        JSON.stringify({ ok: false, error: 'debug_disabled', VER })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    if (wantDebug && ALLOW_DEBUG) {
      const SECRET = getSecret_();
      const base = lineId + '|' + ts;
      const raw = Utilities.computeHmacSha256Signature(base, SECRET, Utilities.Charset.UTF_8);
      const expect = Utilities.base64EncodeWebSafe(raw).replace(/=+$/, '');
      const secretFP = Utilities.base64EncodeWebSafe(
        Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SECRET, Utilities.Charset.UTF_8)
      )
        .replace(/=+$/, '')
        .slice(0, 16);

      return ContentService.createTextOutput(
        JSON.stringify({ ok: true, VER, base, lineId, ts, providedSig, expect, secretLen: SECRET.length, secretFP })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'markReopen') {
      const contentType = (e && e.postData && e.postData.type) || '';
      if (!contentType || contentType.toLowerCase().indexOf('application/json') === -1) {
        return statusApi_jsonOut_({ ok: false, error: 'content_type_must_be_json' }, 415);
      }
      return statusApi_handleMarkReopenPost_(body || {});
    }

    // 通常フロー: 署名チェック（返却点はここだけ）
    const SECRET = getSecret_();
    const base = lineId + '|' + ts;
    const raw = Utilities.computeHmacSha256Signature(base, SECRET, Utilities.Charset.UTF_8);
    const expectB64Url = Utilities.base64EncodeWebSafe(raw).replace(/=+$/, '');
    const expectHex = raw
      .map(function (b) {
        var s = (b & 0xff).toString(16);
        return s.length === 1 ? '0' + s : s;
      })
      .join('');

    ST = 'sig_checked';
    if (!(providedSig === expectB64Url || providedSig.toLowerCase() === expectHex)) {
      try {
        if (String(action) === 'intake_complete') {
          bs_appendLog_([new Date(), 'fail_intake_bad_sig', lineId]);
        }
      } catch (_) {}
      return ContentService.createTextOutput(
        JSON.stringify({
          error: 'bad_sig',
          VER,
          base,
          providedSig: providedSigRaw,
          expectB64Url,
          expectHex,
          secretLen: SECRET.length,
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // 新フロー: action でルーティング（デフォルトは副作用なし）
    ST = 'ensure_master';
    bs_ensureMaster_();

    const aliasIdx = function (map, key, fallbacks) {
      const keys = [key].concat(fallbacks || []);
      for (var i = 0; i < keys.length; i++) {
        var k = keys[i];
        if (map && typeof map[k] === 'number') return map[k];
      }
      return -1;
    };

    const userKey = String(
      (body || {}).userKey ?? (qs || {}).userKey ?? lineId.slice(0, 6).toLowerCase()
    ).trim();
    const displayName = String((body || {}).displayName ?? (qs || {}).displayName ?? '');
    const email = String((body || {}).email ?? (qs || {}).email ?? '');

    ST = 'upsert_contact';
    const up = bs_upsertContact_({ userKey, lineId, displayName, email });

    if (action === 'status') {
      const sh = bs_getSheet_(SHEET_CONTACTS);
      const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idx = bs_toIndexMap_(headers);
      const rowVals = sh.getRange(up.row, 1, 1, headers.length).getValues()[0];
      const intakeIdx = aliasIdx(idx, 'intake_at', ['intakeAt']);
      const activeIdx = aliasIdx(idx, 'active_case_id', ['activeCaseId']);
      const intakeAt = intakeIdx >= 0 ? rowVals[intakeIdx] : '';
      const activeCaseIdRaw = activeIdx >= 0 ? rowVals[activeIdx] : '';
      const activeCaseId = bs_normCaseId_(activeCaseIdRaw);
      return handleStatus_(lineId, activeCaseId, !!intakeAt, userKey);
    }

    if (action === 'intake_complete') {
      ST = 'ensure_drive_root';
      bs_ensureDriveRoot_();
      const sh = bs_getSheet_(SHEET_CONTACTS);
      const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idx = bs_toIndexMap_(headers);
      const rowVals = sh.getRange(up.row, 1, 1, headers.length).getValues()[0];
      const activeIdx = aliasIdx(idx, 'active_case_id', ['activeCaseId']);
      let activeCaseId = bs_normCaseId_(activeIdx >= 0 ? rowVals[activeIdx] || '' : '');
      if (!activeCaseId) {
        ST = 'issue_case_id';
        const issued = bs_issueCaseId_(userKey, lineId);
        activeCaseId = bs_normCaseId_(issued.caseId);
      }
      ST = 'ensure_case_folder';
      const folderId = bs_ensureCaseFolder_(userKey, activeCaseId);
      ST = 'writeback_case_folder';
      const shCases = bs_getSheet_(SHEET_CASES);
      // cases シートの該当行に folderId を書く（最後に追加された or 探索）
      const h2 = shCases.getRange(1, 1, 1, shCases.getLastColumn()).getValues()[0];
      const i2 = bs_toIndexMap_(h2);
      const caseIdx = aliasIdx(i2, 'case_id', ['caseId']);
      const folderIdx = aliasIdx(i2, 'folder_id', ['folderId']);
      const statusIdx = aliasIdx(i2, 'status', []);
      const rc = shCases.getLastRow() - 1;
      if (rc > 0 && caseIdx >= 0) {
        const rows = shCases.getRange(2, 1, rc, h2.length).getValues();
        for (let i = rows.length - 1; i >= 0; i--) {
          if (bs_normCaseId_(rows[i][caseIdx]) === activeCaseId) {
            if (folderIdx >= 0) shCases.getRange(i + 2, folderIdx + 1).setValue(folderId);
            if (statusIdx >= 0) shCases.getRange(i + 2, statusIdx + 1).setValue('intake');
            break;
          }
        }
      }
      ST = 'writeback_contacts';
      const shContacts = bs_getSheet_(SHEET_CONTACTS);
      const contactActiveIdx = aliasIdx(up.idx, 'active_case_id', ['activeCaseId']);
      if (contactActiveIdx >= 0) {
        shContacts.getRange(up.row, contactActiveIdx + 1).setValue(activeCaseId);
      }
      bs_setContactIntakeAt_(lineId, new Date().toISOString());
      // _staging にある intake JSON を案件直下へ移送（存在すれば）
      bs_collectIntakeFromStaging_(lineId, activeCaseId);
      const res = {
        ok: true,
        VER,
        activeCaseId,
        caseKey: userKey + '-' + activeCaseId,
        folderId,
        ts: new Date().toISOString(),
      };
      ST = 'respond_ok';
      bs_appendLog_([new Date(), 'ok_intake', activeCaseId, folderId, userKey]);
      return bs_jsonResponse_(res, 200);
    }

    // デフォルト: 従来のような副作用は行わず、軽いステータスのみ返却
    const sh = bs_getSheet_(SHEET_CONTACTS);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const idx = bs_toIndexMap_(headers);
    const rowVals = sh.getRange(up.row, 1, 1, headers.length).getValues()[0];
    const intakeIdx = aliasIdx(idx, 'intake_at', ['intakeAt']);
    const activeIdx = aliasIdx(idx, 'active_case_id', ['activeCaseId']);
    return bs_jsonResponse_(
      {
        ok: true,
        VER,
        hasIntake: !!(intakeIdx >= 0 && rowVals[intakeIdx]),
        activeCaseId: activeIdx >= 0 ? rowVals[activeIdx] || null : null,
      },
      200
    );
  } catch (err) {
    try {
      bs_appendLog_([new Date(), 'ERR', ST + ':' + String(err)]);
      if (String(action) === 'intake_complete') {
        bs_appendLog_([new Date(), 'fail_intake', String(err)]);
      }
    } catch (_) {}
    return ContentService.createTextOutput(
      JSON.stringify({ ok: false, error: String(err), stage: ST })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

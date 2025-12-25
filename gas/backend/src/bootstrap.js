/**
 * bootstrap.js (BAS / Google Apps Script)
 *  - 認証方式: URL クエリ (?sig, ?ts)
 *  - 署名対象: lineId + '|' + ts（UNIX秒）
 *  - アルゴリズム: HMAC-SHA256 → Base64URL
 *  - ts: UNIX秒 / 許容スキュー ±300秒 / 本文の ts とクエリ ts を整合確認
 *  - 処理: contacts upsert のみ（caseId 採番／ケースフォルダ作成は intake_complete 側で実施）
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

/** ===== 共通マッチャー（必要なら定義）: case_key → case_id → line_id ===== */
if (typeof normCaseId_ !== 'function') {
  function normCaseId_(s) {
    s = String(s || '').trim();
    var n = s.replace(/^0+/, '');
    if (!n) return '';
    var num = parseInt(n, 10);
    if (!isFinite(num)) return '';
    return Utilities.formatString('%04d', num);
  }
}
if (typeof normCaseKey_ !== 'function') {
  function normCaseKey_(s) {
    s = String(s || '').trim().toLowerCase();
    var m = s.match(/^([a-z0-9]{2,})-(\d{1,})$/);
    if (!m) return s;
    return m[1] + '-' + normCaseId_(m[2]);
  }
}
if (typeof normLineId_ !== 'function') {
  function normLineId_(s) { return String(s || '').trim(); }
}
if (typeof matchMetaToCase_ !== 'function') {
  function matchMetaToCase_(fileMeta, known) {
    var fm = fileMeta || {};
    var fk = normCaseKey_(fm.case_key || fm.caseKey || '');
    var fid = normCaseId_(fm.case_id || fm.caseId || '');
    var fl = normLineId_(fm.line_id || fm.lineId || '');
    var kk = normCaseKey_(known && known.case_key || '');
    var kid = normCaseId_(known && known.case_id || '');
    var kl = normLineId_(known && known.line_id || '');
    if (fk && kk && fk === kk) return { ok: true, by: 'case_key' };
    if (fid && kid && fid === kid) return { ok: true, by: 'case_id' };
    if (fl && kl && fl === kl) return { ok: true, by: 'line_id' };
    return { ok: false, by: '' };
  }
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
    // 追加: 時刻スキュー検証（±600秒）
    const nowSec = Math.floor(Date.now() / 1000);
    const tsNum = Number(ts);
    if (!Number.isFinite(tsNum) || Math.abs(tsNum - nowSec) > 600) {
      logCtx_('sig:ts-skew', { now: nowSec, ts: tsNum, diff: tsNum - nowSec });
      return { ok: false, line_id: '', case_id: '' };
    }
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

  // [LOG-5] caseId 読み出しのみ（既存が無ければ空のまま）
  const sh = bs_getSheet_(SHEET_CONTACTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = bs_toIndexMap_(headers);
  const rowVals = sh.getRange(up.row, 1, 1, headers.length).getValues()[0];
  const activeIdx = typeof idx['active_case_id'] === 'number' ? idx['active_case_id'] : idx['activeCaseId'];
  const caseId = activeIdx >= 0 ? bs_normCaseId_(rowVals[activeIdx] || '') : '';

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
function bs_contactsSchema_() {
  if (
    typeof SCHEMA === 'object' &&
    SCHEMA &&
    Array.isArray(SCHEMA.contacts) &&
    SCHEMA.contacts.length
  ) {
    return SCHEMA.contacts.slice();
  }
  // fallback（最低限）
  return ['line_id', 'user_key', 'active_case_id', 'intake_at', 'updated_at'];
}

function bs_upsertContact_(payload) {
  const sh = bs_getSheet_(SHEET_CONTACTS);
  if (sh.getLastRow() < 1) {
    const headers = bs_contactsSchema_();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = bs_toIndexMap_(headers);
  const required = bs_contactsSchema_();
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

function resolveCaseFolderId_(lineId, caseId, createIfMissing /* = false */) {
  const direct = bs_resolveCaseFolderId_('', caseId, lineId);
  if (direct) return direct;

  const userKey = lineId ? drive_userKeyFromLineId_(lineId) : '';
  if (userKey) {
    const viaUser = bs_resolveCaseFolderId_(userKey, caseId, lineId);
    if (viaUser) return viaUser;
    if (createIfMissing) {
      try {
        const ensuredId = bs_ensureCaseFolder_(userKey, caseId);
        if (ensuredId) return ensuredId;
      } catch (err) {
        try {
          Logger.log('[resolveCaseFolderId_] ensure error: %s', (err && err.stack) || err);
        } catch (_) {}
      }
    }
  }

  return '';
}

function bs_collectIntakeFromStaging_(lineId, caseId, knownFolderId /* optional */) {
  const normalizedCaseId = bs_normCaseId_(caseId);
  if (!lineId || !normalizedCaseId) return;

  let caseFolderId = knownFolderId || '';
  if (!caseFolderId) {
    caseFolderId = resolveCaseFolderId_(lineId, normalizedCaseId, /*createIfMissing=*/false);
  }
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

  // submissions 追記の件数（診断用）
  var appended = 0;
  var moved = 0;
  var lastSavedName = '';

  try {
    const root = DriveApp.getFolderById(DRIVE_ROOT_ID);
    try { logCtx_('collectStaging', { lineId, caseId: normalizedCaseId, folderId: caseFolderId }); } catch (_) {}
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
              // フィルタ: case_key → case_id → line_id の順で一致判定
              var text0 = '';
              var data0 = null;
              try { text0 = file.getBlob().getDataAsString('utf-8'); } catch (_) {}
              try { data0 = JSON.parse(text0); } catch (_) {}
              var ukey0 = drive_userKeyFromLineId_(lineId);
              var cid0 = normalizedCaseId;
              var ckey0 = (ukey0 ? (ukey0 + '-') : '') + cid0; // uk 不明時は "-0001" を作らない
              var meta0 = (data0 && data0.meta) || {};
              var known0 = { case_key: (ukey0 ? (ukey0 + '-' + cid0) : ''), case_id: cid0, line_id: String(lineId) };
              var res0 = (typeof matchMetaToCase_ === 'function') ? matchMetaToCase_(meta0, known0) : { ok: false };
              var matched0 = !!(res0 && res0.ok);
              // 採用条件: 一致済み or (メタが空 かつ 直近15分内)
              var updatedAt = +(file.getLastUpdated && file.getLastUpdated());
              var recent = Number.isFinite(updatedAt) ? (Date.now() - updatedAt <= 15 * 60 * 1000) : false;
              var noMeta = !(fileCKey0 || fileCID0 || fileLID0);
              if (!(matched0 || (noMeta && recent))) continue;

              // 一致した場合のみ、内容を meta 補完して「書き直し」で案件直下へ保存
              try {
                var raw = '';
                try { raw = file.getBlob().getDataAsString('utf-8'); } catch (_) {}
                var obj = {};
                try { obj = JSON.parse(raw || '{}'); } catch (_) { obj = {}; }
                try {
                  if (typeof intake_fillMeta_ === 'function') {
                    obj = intake_fillMeta_(obj, { line_id: lineId, case_id: cid0 }) || obj;
                  }
                } catch (_) {}
                // 呼び出し時の lineId/caseId を権威として最終上書き
                try {
                  obj = obj && typeof obj === 'object' ? obj : {};
                  obj.meta = obj.meta || {};
                  obj.meta.line_id = String(lineId);
                  var ukeyFill = '';
                  try { ukeyFill = drive_userKeyFromLineId_(lineId) || ''; } catch (_) {}
                  if (!obj.meta.user_key && ukeyFill) obj.meta.user_key = ukeyFill;
                  obj.meta.case_id = String(cid0);
                  if (!obj.meta.case_key && ukeyFill && cid0) obj.meta.case_key = ukeyFill + '-' + cid0;
                } catch (_) {}
                var subIdNew = '';
                try { if (typeof extractSubmissionId_ === 'function') subIdNew = extractSubmissionId_(name) || ''; } catch (_) {}
                if (!subIdNew) { var mm = String(name || '').match(/__(\d+)\.json$/); if (mm) subIdNew = mm[1]; }
                if (!subIdNew) subIdNew = String(Date.now());
                var newName = 'intake__' + subIdNew + '.json';
                var jsonStr = '';
                try { jsonStr = JSON.stringify(obj); } catch (_) { jsonStr = raw || '{}'; }
                try {
                  caseFolder.createFile(Utilities.newBlob(jsonStr, 'application/json', newName));
                  lastSavedName = newName;
                } catch (_) {}
                try { file.setTrashed(true); } catch (_) {}
                try { moved++; } catch (_) {}
              } catch (_) {}
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
                // submission_id の補完順序: 1) ファイル名 → 2) JSON → 3) 現在時刻
                var sid = '';
                try {
                  if (typeof extractSubmissionId_ === 'function') sid = extractSubmissionId_(name) || '';
                } catch (_) {}
                if (!sid) {
                  var m = String(name || '').match(/__(\d+)\.json$/);
                  if (m) sid = m[1];
                }
                if (!sid && data && typeof data === 'object') {
                  sid = String(
                    data.submission_id || (data.meta && data.meta.submission_id) || ''
                  ).trim();
                }
                if (typeof submissions_upsert_ === 'function') {
                  submissions_upsert_({
                    submission_id: sid || String(Date.now()),
                    form_key: 'intake',
                    case_id: normalizedCaseId,
                    user_key: drive_userKeyFromLineId_(lineId),
                    line_id: lineId,
                    submitted_at: new Date().toISOString(),
                    referrer: lastSavedName || file.getName(),
                    status: 'received',
                  });
                  try { appended++; } catch (_) {}
                }
                // 旧API互換（最小限・最後の手段）
                else if (typeof recordSubmission_ === 'function') {
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

  // 開発用: moved=0 かつ ALLOW_UNIQUE_RESCUE=1 のとき、直近5分・ユニーク1件を救済して meta を充填して書き込み
  try {
    if ((moved | 0) === 0) {
      var allowUnique = (props_().getProperty('ALLOW_UNIQUE_RESCUE') || '').trim() === '1';
      if (allowUnique) {
        var root = DriveApp.getFolderById(DRIVE_ROOT_ID);
        var candidates = [];
        ['_email_staging', '_staging'].forEach(function (name) {
          try {
            var it = root.getFoldersByName(name);
            if (!it.hasNext()) return;
            var st = it.next();
            var itf = st.getFiles();
            while (itf.hasNext()) {
              var f = itf.next();
              var nm = f.getName && f.getName();
              if (!/^intake__\d+\.json$/i.test(String(nm || ''))) continue;
              var t = +((f.getLastUpdated && f.getLastUpdated()) || new Date(0));
              candidates.push({ f: f, t: t, nm: nm });
            }
          } catch (_) {}
        });
        if (candidates.length === 1 && (Date.now() - candidates[0].t) <= 5 * 60 * 1000) {
          try {
            var rawU = '';
            var jsU = {};
            try { rawU = candidates[0].f.getBlob().getDataAsString('utf-8'); } catch (_) {}
            try { jsU = JSON.parse(rawU || '{}'); } catch (_) { jsU = {}; }
            jsU = jsU && typeof jsU === 'object' ? jsU : {};
            jsU.meta = jsU.meta || {};
            var ukeyU = drive_userKeyFromLineId_(lineId) || '';
            if (!jsU.meta.line_id) jsU.meta.line_id = String(lineId);
            if (!jsU.meta.user_key) jsU.meta.user_key = String(ukeyU || '');
            if (!jsU.meta.case_id) jsU.meta.case_id = String(normalizedCaseId);
            if (!jsU.meta.case_key && ukeyU && normalizedCaseId) jsU.meta.case_key = ukeyU + '-' + normalizedCaseId;
            var sidU = '';
            try { if (typeof extractSubmissionId_ === 'function') sidU = extractSubmissionId_(candidates[0].nm) || ''; } catch (_) {}
            if (!sidU) { var mmu = String(candidates[0].nm || '').match(/__(\d+)\.json$/); if (mmu) sidU = mmu[1]; }
            if (!sidU) sidU = String(Date.now());
            var newNameU = 'intake__' + sidU + '.json';
            var outU = '';
            try { outU = JSON.stringify(jsU); } catch (_) { outU = rawU || '{}'; }
            try { caseFolder.createFile(Utilities.newBlob(outU, 'application/json', newNameU)); } catch (_) {}
            try { candidates[0].f.setTrashed(true); } catch (_) {}
            moved++;
            if (typeof submissions_upsert_ === 'function') {
              submissions_upsert_({
                submission_id: sidU,
                form_key: 'intake',
                case_id: normalizedCaseId,
                user_key: ukeyU,
                line_id: lineId,
                submitted_at: new Date().toISOString(),
                referrer: newNameU,
                status: 'received',
              });
              appended++;
            }
            if (typeof bs_setCaseStatus_ === 'function') {
              try { bs_setCaseStatus_(normalizedCaseId, 'intake'); } catch (_) {}
            }
            try { Logger.log('[collectStaging] unique rescue lid=%s cid=%s', lineId, normalizedCaseId); } catch (_) {}
          } catch (_) {}
        }
      }
    }
  } catch (_) {}

  // 最終的な補填件数を出力（診断性向上）
  try {
    Logger.log('[collectStaging] lid=%s cid=%s moved=%s appended=%s', lineId, normalizedCaseId, moved, appended);
  } catch (_) {}
}

function handleStatus_(lineId, caseId, hasIntake, userKey) {
  const intakeFlag = !!hasIntake;
  const normalizedCaseId = bs_normCaseId_(caseId);
  if (!normalizedCaseId) {
    return ContentService.createTextOutput(
      JSON.stringify({ ok: true, caseId: '', intakeReady: false, hasIntake: intakeFlag })
    ).setMimeType(ContentService.MimeType.JSON);
  }
  let folderId = resolveCaseFolderId_(lineId, normalizedCaseId, /*createIfMissing=*/false);
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
  if (!intakeReady && folderId) {
    // 案件フォルダが存在するときのみ _staging を吸い上げる
    bs_collectIntakeFromStaging_(lineId, normalizedCaseId);
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
      // status_api.js 側の POST ルータへ委譲（エントリ一本化）
      if (typeof statusApi_doPost_ === 'function') return statusApi_doPost_(e);
      // フォールバック: 直接ハンドラ呼び出し（互換）
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
    // 追加: 時刻スキュー検証（±600秒）
    try {
      const nowSec = Math.floor(Date.now() / 1000);
      if (!Number.isFinite(ts) || Math.abs(Number(ts) - nowSec) > 600) {
        try { logCtx_('sig:ts-skew', { now: nowSec, ts: Number(ts), diff: Number(ts) - nowSec, action }); } catch (_) {}
        return ContentService.createTextOutput(
          JSON.stringify({ ok: false, error: 'ts_skew', now: nowSec, ts: Number(ts) })
        ).setMimeType(ContentService.MimeType.JSON);
      }
    } catch (_) {}
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
      try { PropertiesService.getScriptProperties().setProperty('LAST_LINE_ID', lineId); } catch (_) {}
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
      try { PropertiesService.getScriptProperties().setProperty('LAST_LINE_ID', lineId); } catch (_) {}
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
      // cases の更新は updateCasesRow_ があればそれを使い、無ければ手動で書き込み
      try {
        if (typeof updateCasesRow_ === 'function') {
          updateCasesRow_(activeCaseId, {
            case_key: userKey + '-' + activeCaseId,
            folder_id: folderId,
            status: 'intake',
          });
        } else {
          const shCases = bs_getSheet_(SHEET_CASES);
          // cases シートの該当行に folderId 等を書き込み（最後に追加された or 探索）
          const h2 = shCases.getRange(1, 1, 1, shCases.getLastColumn()).getValues()[0];
          const i2 = bs_toIndexMap_(h2);
          const caseIdx = aliasIdx(i2, 'case_id', ['caseId']);
          const folderIdx = aliasIdx(i2, 'folder_id', ['folderId']);
          const statusIdx = aliasIdx(i2, 'status', []);
          const caseKeyIdx = aliasIdx(i2, 'case_key', ['caseKey']);
          // cases.case_id / case_key をテキスト（'@'）に固定（安全のため直前でも実施）
          try {
            if (typeof ensureCasesCaseIdTextFormat_ === 'function') ensureCasesCaseIdTextFormat_();
            else {
              const col = h2.indexOf('case_id') + 1;
              if (col > 0) shCases.getRange(1, col, shCases.getMaxRows(), 1).setNumberFormat('@');
            }
            const ckCol = h2.indexOf('case_key') + 1;
            if (ckCol > 0) shCases.getRange(1, ckCol, shCases.getMaxRows(), 1).setNumberFormat('@');
          } catch (_) {}
          const rc = shCases.getLastRow() - 1;
          if (rc > 0 && caseIdx >= 0) {
            const rows = shCases.getRange(2, 1, rc, h2.length).getValues();
            for (let i = rows.length - 1; i >= 0; i--) {
              if (bs_normCaseId_(rows[i][caseIdx]) === activeCaseId) {
                if (folderIdx >= 0) shCases.getRange(i + 2, folderIdx + 1).setValue(String(folderId));
                if (statusIdx >= 0) shCases.getRange(i + 2, statusIdx + 1).setValue('intake');
                if (caseKeyIdx >= 0)
                  shCases.getRange(i + 2, caseKeyIdx + 1).setNumberFormat('@').setValue(userKey + '-' + activeCaseId);
                break;
              }
            }
          }
        }
      } catch (_) {}
      ST = 'writeback_contacts';
      const shContacts = bs_getSheet_(SHEET_CONTACTS);
      const contactActiveIdx = aliasIdx(up.idx, 'active_case_id', ['activeCaseId']);
      if (contactActiveIdx >= 0) {
        shContacts.getRange(up.row, contactActiveIdx + 1).setNumberFormat('@').setValue(activeCaseId);
      }
      bs_setContactIntakeAt_(lineId, new Date().toISOString());
      // 追加: payload があれば intake__*.json を案件直下へ直接保存（任意）
      try {
        const payloadRaw = (body || {}).payload;
        const submissionIdRaw =
          (body || {}).submissionId ?? (body || {}).submission_id ?? (body || {}).subId;
        const submissionId = String(submissionIdRaw || '').trim() || String(Date.now());
        if (payloadRaw != null) {
          const cid = ('0000' + String(activeCaseId)).slice(-4);
          const ukey = userKey;
          const ckey = ukey + '-' + cid;
          let obj = {};
          try { obj = typeof payloadRaw === 'string' ? JSON.parse(payloadRaw) : (payloadRaw || {}); } catch (_) { obj = {}; }
          obj = obj && typeof obj === 'object' ? obj : {};
          obj.meta = obj.meta || {};
          if (!obj.meta.line_id) obj.meta.line_id = (obj.meta.lineId || lineId || '');
          if (!obj.meta.user_key) obj.meta.user_key = ukey;
          if (!obj.meta.case_id) obj.meta.case_id = cid;
          if (!obj.meta.case_key) obj.meta.case_key = ckey;

          const jsonName = 'intake__' + submissionId + '.json';
          try {
            const jsonStr = JSON.stringify(obj);
            DriveApp.getFolderById(folderId).createFile(
              Utilities.newBlob(jsonStr, 'application/json', jsonName)
            );
          } catch (_) {}
          // submissions 追記
          try {
            if (typeof submissions_upsert_ === 'function') {
              submissions_upsert_({
                submission_id: String(submissionId),
                form_key: 'intake',
                case_id: cid,
                user_key: ukey,
                line_id: lineId,
                submitted_at: new Date().toISOString(),
                referrer: jsonName,
                status: 'received',
              });
            }
          } catch (_) {}
        }
      } catch (_) {}

      // _staging にある intake JSON を案件直下へ移送（存在すれば）
      bs_collectIntakeFromStaging_(lineId, activeCaseId, folderId);
      const res = {
        ok: true,
        VER,
        activeCaseId,
        caseKey: userKey + '-' + activeCaseId,
        folderId,
        ts: new Date().toISOString(),
      };
      try {
        PropertiesService.getScriptProperties().setProperty('LAST_CASE_ID', activeCaseId);
      } catch (_) {}
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

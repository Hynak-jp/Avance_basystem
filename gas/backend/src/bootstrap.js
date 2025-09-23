/**
 * bootstrap.js (BAS / Google Apps Script)
 *  - 認証方式: URL クエリ (?sig, ?ts)
 *  - 署名対象: lineId + '|' + ts（UNIX秒）
 *  - アルゴリズム: HMAC-SHA256 → Base64URL
 *  - ts: UNIX秒 / 許容スキュー ±300秒 / 本文の ts とクエリ ts を整合確認
 *  - 処理: contacts upsert → caseId 採番（ロック）→ Drive フォルダ作成 → cases/contacts 更新
 */

/** ---------- 設定 ---------- **/
const PROP = PropertiesService.getScriptProperties();
// Avoid global name collision across GAS files: use a file‑scoped alias
const BS_MASTER_SPREADSHEET_ID = PROP.getProperty('BAS_MASTER_SPREADSHEET_ID'); // BAS_master
// 両対応: DRIVE_ROOT_FOLDER_ID / ROOT_FOLDER_ID
const DRIVE_ROOT_ID =
  PROP.getProperty('DRIVE_ROOT_FOLDER_ID') || PROP.getProperty('ROOT_FOLDER_ID'); // BAS_提出書類 ルート
const BAS_BOOTSTRAP_SECRET = PROP.getProperty('BOOTSTRAP_SECRET') || ''; // 衝突回避のためユニーク名
const LOG_SHEET_ID = PROP.getProperty('GAS_LOG_SHEET_ID') || ''; // 任意（空で無効）
const ALLOW_DEBUG = (PROP.getProperty('ALLOW_DEBUG') || '').toLowerCase() === '1'; //

const SHEET_CONTACTS = 'contacts';
const SHEET_CASES = 'cases';

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

function bs_getSheet_(name) {
  const ss = bs_ensureSpreadsheet_(bs_openSpreadsheet_(BS_MASTER_SPREADSHEET_ID)); // Spreadsheet
  let sh = ss.getSheetByName(name); // Sheet
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function bs_toIndexMap_(headers) {
  const m = {};
  headers.forEach((h, i) => (m[h] = i));
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
  if (!rid) throw new Error('ROOT_FOLDER_ID is empty');
  DriveApp.getFolderById(rid); // 存在チェック。権限が無ければここで例外。
}

/** ---------- contacts upsert ---------- **/
function bs_upsertContact_(payload) {
  const sh = bs_getSheet_(SHEET_CONTACTS);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 7).setValues([
      ['userKey', 'lineId', 'displayName', 'email', 'activeCaseId', 'updatedAt', 'intakeAt'],
    ]);
  }
  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = bs_toIndexMap_(headers);

  const need = [
    'userKey',
    'lineId',
    'displayName',
    'email',
    'activeCaseId',
    'updatedAt',
    'intakeAt',
  ];
  let changed = false;
  need.forEach((col) => {
    if (!(col in idx)) {
      headers.push(col);
      idx[col] = headers.length - 1;
      changed = true;
    }
  });
  if (changed) sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const { userKey, lineId, displayName, email } = payload;
  if (!userKey) throw new Error('userKey required');
  if (!lineId) throw new Error('lineId required');

  const rowCount = sh.getLastRow() - 1;
  const rows = rowCount > 0 ? sh.getRange(2, 1, rowCount, headers.length).getValues() : [];
  let found = -1;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][idx['userKey']] === userKey) {
      found = i;
      break;
    }
  }

  const now = new Date().toISOString();
  if (found >= 0) {
    const row = 2 + found;
    const current = sh.getRange(row, 1, 1, headers.length).getValues()[0];
    current[idx['lineId']] = lineId;
    if (displayName != null) current[idx['displayName']] = displayName;
    if (email != null) current[idx['email']] = email;
    current[idx['updatedAt']] = now;
    sh.getRange(row, 1, 1, headers.length).setValues([current]);
    return { row, headers, idx };
  } else {
    const values = new Array(headers.length).fill('');
    values[idx['userKey']] = userKey;
    values[idx['lineId']] = lineId;
    values[idx['displayName']] = displayName || '';
    values[idx['email']] = email || '';
    values[idx['activeCaseId']] = '';
    values[idx['updatedAt']] = now;
    if ('intakeAt' in idx) values[idx['intakeAt']] = '';
    sh.appendRow(values);
    return { row: sh.getLastRow(), headers, idx };
  }
}

function bs_setContactIntakeAt_(lineId, when) {
  const sh = bs_getSheet_(SHEET_CONTACTS);
  if (sh.getLastRow() < 1) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = bs_toIndexMap_(headers);
  if (!('intakeAt' in idx)) return; // 列がまだ無い場合は黙って無視
  const rowCount = sh.getLastRow() - 1;
  if (rowCount < 1) return;
  const rows = sh.getRange(2, 1, rowCount, headers.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][idx['lineId']]) === String(lineId)) {
      sh.getRange(i + 2, idx['intakeAt'] + 1).setValue(when || new Date().toISOString());
      return;
    }
  }
}

function bs_setCaseStatus_(caseId, status) {
  const sh = bs_getSheet_(SHEET_CASES);
  if (sh.getLastRow() < 1) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = bs_toIndexMap_(headers);
  const rowCount = sh.getLastRow() - 1;
  if (rowCount < 1) return;
  const rows = sh.getRange(2, 1, rowCount, headers.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][idx['caseId']]) === String(caseId)) {
      sh.getRange(i + 2, idx['status'] + 1).setValue(status);
      sh.getRange(i + 2, idx['lastActivity'] ? idx['lastActivity'] + 1 : 5).setValue(
        new Date().toISOString()
      );
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
        ['caseId', 'userKey', 'lineId', 'status', 'folderId', 'createdAt'],
      ]);
    }
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const idx = bs_toIndexMap_(headers);
    const rowCount = sh.getLastRow() - 1;
    const rows = rowCount > 0 ? sh.getRange(2, 1, rowCount, headers.length).getValues() : [];

    let maxNum = 0;
    const cidIdx = idx['caseId'];
    for (let i = 0; i < rows.length; i++) {
      const n = parseInt(String(rows[i][cidIdx] || ''), 10);
      if (Number.isFinite(n)) maxNum = Math.max(maxNum, n);
    }
    const next = String(maxNum + 1).padStart(4, '0');
    const now = new Date().toISOString();

    const values = new Array(headers.length).fill('');
    values[idx['caseId']] = next;
    values[idx['userKey']] = userKey;
    values[idx['lineId']] = lineId;
    values[idx['status']] = 'draft';
    values[idx['folderId']] = '';
    values[idx['createdAt']] = now;
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
  const root = DriveApp.getFolderById(DRIVE_ROOT_ID);
  const name = `${userKey}-${caseId}`;
  const it = root.getFoldersByName(name);
  if (it.hasNext()) return it.next().getId();
  return root.createFolder(name).getId();
}

/**
 * intake JSON が案件フォルダに移動済みかを確認（Drive には副作用なし）
 */
function bs_isIntakeJsonReady_(userKey, caseId) {
  if (!DRIVE_ROOT_ID || !userKey || !caseId) return false;
  try {
    const root = DriveApp.getFolderById(DRIVE_ROOT_ID);
    const folders = root.getFoldersByName(`${userKey}-${caseId}`);
    if (!folders.hasNext()) return false;
    const caseFolder = folders.next();
    const files = caseFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const name = file && file.getName ? file.getName() : '';
      if (/^intake__/i.test(name) && /\.json$/i.test(name)) return true;
    }
  } catch (e) {
    try {
      Logger.log('[bs_isIntakeJsonReady_] error: %s', (e && e.stack) || e);
    } catch (_) {}
  }
  return false;
}

/** ---------- doGet / doPost ---------- **/
function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ ok: true, VER: 'dbg-final' })
  ).setMimeType(ContentService.MimeType.JSON);
}

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
      const SECRET = PropertiesService.getScriptProperties().getProperty('BOOTSTRAP_SECRET') || '';
      const base = lineId + '|' + ts;
      const raw = Utilities.computeHmacSha256Signature(base, SECRET, Utilities.Charset.UTF_8);
      const expect = Utilities.base64EncodeWebSafe(raw).replace(/=+$/, '');
      const secretFP = Utilities.base64EncodeWebSafe(
        Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SECRET, Utilities.Charset.UTF_8)
      )
        .replace(/=+$/, '')
        .slice(0, 16);

      return ContentService.createTextOutput(
        JSON.stringify({
          ok: true,
          VER,
          base,
          lineId,
          ts,
          providedSig,
          expect,
          secretLen: SECRET.length,
          secretFP,
          // 本番でも漏れないレベルの指紋のみ（長さを出すのが嫌なら消してOK）
          secretLen: SECRET.length,
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // 通常フロー: 署名チェック（返却点はここだけ）
    const SECRET = PropertiesService.getScriptProperties().getProperty('BOOTSTRAP_SECRET') || '';
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
    const { contacts } = bs_ensureMaster_();

    const userKey = String(
      (body || {}).userKey ?? (qs || {}).userKey ?? lineId.slice(0, 6).toLowerCase()
    ).trim();
    const displayName = String((body || {}).displayName ?? (qs || {}).displayName ?? '');
    const email = String((body || {}).email ?? (qs || {}).email ?? '');

    ST = 'upsert_contact';
    const up = bs_upsertContact_({ userKey, lineId, displayName, email });

    if (action === 'status') {
      // 受付済みかの確認（副作用なし）
      const sh = bs_getSheet_(SHEET_CONTACTS);
      const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idx = bs_toIndexMap_(headers);
      const rowVals = sh.getRange(up.row, 1, 1, headers.length).getValues()[0];
      const intakeAt = 'intakeAt' in idx ? rowVals[idx['intakeAt']] : '';
      const activeCaseId = 'activeCaseId' in idx ? rowVals[idx['activeCaseId']] : '';
      const hasIntake = !!intakeAt;
      const intakeReady = hasIntake && activeCaseId ? bs_isIntakeJsonReady_(userKey, activeCaseId) : false;
      return bs_jsonResponse_(
        {
          ok: true,
          VER,
          hasIntake,
          activeCaseId: activeCaseId || null,
          intakeReady,
        },
        200
      );
    }

    if (action === 'intake_complete') {
      ST = 'ensure_drive_root';
      bs_ensureDriveRoot_();
      const sh = bs_getSheet_(SHEET_CONTACTS);
      const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idx = bs_toIndexMap_(headers);
      const rowVals = sh.getRange(up.row, 1, 1, headers.length).getValues()[0];
      let activeCaseId = rowVals[idx['activeCaseId']] || '';
      if (!activeCaseId) {
        ST = 'issue_case_id';
        const issued = bs_issueCaseId_(userKey, lineId);
        activeCaseId = issued.caseId;
      }
      ST = 'ensure_case_folder';
      const folderId = bs_ensureCaseFolder_(userKey, activeCaseId);
      ST = 'writeback_case_folder';
      const shCases = bs_getSheet_(SHEET_CASES);
      // cases シートの該当行に folderId を書く（最後に追加された or 探索）
      const h2 = shCases.getRange(1, 1, 1, shCases.getLastColumn()).getValues()[0];
      const i2 = bs_toIndexMap_(h2);
      const rc = shCases.getLastRow() - 1;
      if (rc > 0) {
        const rows = shCases.getRange(2, 1, rc, h2.length).getValues();
        for (let i = rows.length - 1; i >= 0; i--) {
          if (String(rows[i][i2['caseId']]) === String(activeCaseId)) {
            shCases.getRange(i + 2, i2['folderId'] + 1).setValue(folderId);
            shCases.getRange(i + 2, i2['status'] + 1).setValue('intake');
            break;
          }
        }
      }
      ST = 'writeback_contacts';
      const shContacts = bs_getSheet_(SHEET_CONTACTS);
      shContacts.getRange(up.row, up.idx['activeCaseId'] + 1).setValue(activeCaseId);
      bs_setContactIntakeAt_(lineId, new Date().toISOString());
      // _staging にある intake JSON を案件直下へ移送（存在すれば）
      try {
        if (typeof moveStagingIntakeJsonToCase_ === 'function') {
          moveStagingIntakeJsonToCase_(lineId, activeCaseId);
        }
      } catch (_) {}
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
    return bs_jsonResponse_(
      {
        ok: true,
        VER,
        hasIntake: !!(idx['intakeAt'] != null && rowVals[idx['intakeAt']]),
        activeCaseId: rowVals[idx['activeCaseId']] || null,
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

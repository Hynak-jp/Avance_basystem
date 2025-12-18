/******** Schema Guard for BAS Master (public funcs: no trailing underscore) ********/

const SCHEMA = {
  contacts: [
    // 必須（最小運用）
    'line_id',
    'user_key',
    'active_case_id',
    'updated_at',
    'intake_at',
  ],
  cases: [
    'line_id',
    'case_id',
    'user_key',
    'case_key',
    'created_at',
    'status',
    'last_activity',
    'folder_id',
    'intake_at',
    'updated_at',
  ],
  submissions: [
    'submission_id',
    'form_key',
    'case_id',
    'user_key',
    'line_id',
    'submitted_at',
    'seq',
    'referrer',
    'redirect_url',
    'status',
  ],
};

// ====== 公開関数（UIから実行／トリガー対象）======

function getMasterSheetId_() {
  const props = PropertiesService.getScriptProperties();
  return (
    props.getProperty('SHEET_ID') ||
    props.getProperty('BAS_MASTER_SPREADSHEET_ID') ||
    ''
  ).trim();
}

// 1) SHEET_ID 設定
function setSheetId() {
  const NEW_SHEET_ID = '1G1IPbmGM1USdpRb9T6-56qggBEzyo-DF0zMUAn0ui6o'; // ←必要に応じて更新
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', NEW_SHEET_ID);
  Logger.log('SHEET_ID updated: ' + NEW_SHEET_ID);
}

// 2) スキーマ検証（手動実行用）
function validateMaster() {
  return validateMasterInternal_();
}

// 3) 夜間監視ジョブ（トリガーから呼ばれる）
function nightlySchemaWatch() {
  try {
    validateMasterInternal_();
    Logger.log('Schema OK: ' + new Date().toISOString());
  } catch (e) {
    MailApp.sendEmail({
      to: 'you@example.com',
      subject: '[BAS] Schema error on master',
      htmlBody: '<pre>' + String(e) + '</pre>',
    });
    throw e;
  }
}

// 4) トリガーのインストール／アンインストール
function installNightlyTrigger() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === 'nightlySchemaWatch') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('nightlySchemaWatch').timeBased().atHour(2).everyDays(1).create();
  Logger.log('Installed nightly schema trigger.');
}

function removeNightlyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((t) => {
    if (t.getHandlerFunction() === 'nightlySchemaWatch') {
      ScriptApp.deleteTrigger(t);
    }
  });
  Logger.log('Removed nightly schema trigger(s).');
}

// ====== 内部ヘルパー（末尾 _ OK）======

function validateMasterInternal_() {
  const sid = getMasterSheetId_();
  if (!sid) throw new Error('SHEET_ID / BAS_MASTER_SPREADSHEET_ID not set.');
  const ss = SpreadsheetApp.openById(sid);

  const results = [];
  Object.keys(SCHEMA).forEach((name) => {
    const sh = ss.getSheetByName(name);
    if (!sh) {
      results.push({ sheet: name, ok: false, error: 'missing sheet' });
      return;
    }
    const lastCol = Math.max(1, sh.getLastColumn());
    const header = (sh.getRange(1, 1, 1, lastCol).getValues()[0] || [])
      .map((v) => String(v || '').trim())
      .filter((v) => v);
    const want = SCHEMA[name];
    const missing = want.filter((c) => !header.includes(c));
    const extras = header.filter((c) => !want.includes(c));
    const dups = header.filter((c, i) => header.indexOf(c) !== i);

    results.push({
      sheet: name,
      ok: missing.length === 0 && dups.length === 0, // extras は許容
      missing: missing.join(', '),
      extras: extras.join(', '),
      duplicates: dups.join(', '),
    });
  });

  Logger.log(JSON.stringify(results, null, 2));
  const ng = results.filter((r) => !r.ok);
  if (ng.length) throw new Error('Schema violation: ' + JSON.stringify(ng));
  return results;
}

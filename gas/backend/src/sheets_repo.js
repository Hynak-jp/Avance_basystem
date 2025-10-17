/**
 * sheets_repo.js
 *  - BAS_master の cases_forms / submissions / cases シートを扱う薄いリポジトリ
 *  - 既存の bootstrap.js 内ユーティリティ（bs_getSheet_ など）を利用
 */

const SHEETS_REPO_PROPS = PropertiesService.getScriptProperties();

const SHEETS_REPO_CASES_FORMS = SHEETS_REPO_PROPS.getProperty('SHEET_CASES_FORMS') || 'cases_forms';
const SHEETS_REPO_SUBMISSIONS = SHEETS_REPO_PROPS.getProperty('SHEET_SUBMISSIONS') || 'submissions';
const SHEETS_REPO_CASES = SHEETS_REPO_PROPS.getProperty('SHEET_CASES') || 'cases';

const SHEETS_REPO_CASES_FORMS_HEADERS = [
  'case_id',
  'form_key',
  'status',
  'can_edit',
  'reopened_at',
  'reopened_by',
  'locked_reason',
  'reopen_until',
  'updated_at',
];

const SHEETS_REPO_SUBMISSIONS_HEADERS = [
  'case_id',
  'case_key',
  'form_key',
  'seq',
  'submission_id',
  'received_at',
  'supersedes_seq',
  'json_path',
  // 追加カラム（存在しなければ自動で増設）
  'referrer',
  'redirect_url',
  'status',
  'user_key',
  'line_id',
  'submitted_at',
  'reopened_at',
  'reopened_by',
  'reopen_until',
  'locked_reason',
  'can_edit',
  'reopened_at_epoch',
  'reopen_until_epoch',
];

function sheetsRepo_aliases_(key) {
  if (typeof bs_headerAliases_ === 'function') return bs_headerAliases_(key);
  const value = String(key == null ? '' : key).trim();
  if (!value) return [];
  const lower = value.toLowerCase();
  const snake = lower
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
  const camel = snake.replace(/_([a-z0-9])/g, function (_, c) {
    return c.toUpperCase();
  });
  const pascal = camel ? camel.charAt(0).toUpperCase() + camel.slice(1) : '';
  const flat = snake.replace(/_/g, '');
  const aliases = new Set([value, lower, snake, camel, pascal, flat, flat.toLowerCase()]);
  return Array.from(aliases).filter(function (v) {
    return !!v;
  });
}

function sheetsRepo_getValue_(obj, key) {
  if (!obj) return undefined;
  const aliases = sheetsRepo_aliases_(key);
  for (let i = 0; i < aliases.length; i++) {
    const alias = aliases[i];
    if (Object.prototype.hasOwnProperty.call(obj, alias)) {
      return obj[alias];
    }
  }
  return undefined;
}

function sheetsRepo_assignAliases_(target, key, value) {
  const aliases = sheetsRepo_aliases_(key);
  for (let i = 0; i < aliases.length; i++) {
    const alias = aliases[i];
    if (!Object.prototype.hasOwnProperty.call(target, alias)) {
      target[alias] = value;
    }
  }
  if (aliases.length && !Object.prototype.hasOwnProperty.call(target, key)) {
    target[key] = value;
  }
}

function sheetsRepo_ensureSheet_(name, headers) {
  const sheet = bs_getSheet_(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const headerRange = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length));
    const current = headerRange.getValues()[0].map(function (v) {
      return String(v || '').trim();
    });
    if (current.length < headers.length) {
      sheet.insertColumnsAfter(current.length || 1, headers.length - current.length);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      const merged = headers.map(function (key, idx) {
        return current[idx] && current[idx].length ? current[idx] : key;
      });
      sheet.getRange(1, 1, 1, merged.length).setValues([merged]);
    }
  }
  return sheet;
}

function sheetsRepo_readAll_(sheetName, headers) {
  const sheet = sheetsRepo_ensureSheet_(sheetName, headers);
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return values.map(function (row) {
    const obj = {};
    headers.forEach(function (key, idx) {
      sheetsRepo_assignAliases_(obj, key, row[idx]);
    });
    return obj;
  });
}

function sheetsRepo_buildIndexMap_(headers) {
  const map = {};
  (headers || []).forEach(function (header, idx) {
    const aliases = sheetsRepo_aliases_(header || '');
    aliases.forEach(function (alias) {
      if (!alias) return;
      if (!Object.prototype.hasOwnProperty.call(map, alias)) {
        map[alias] = idx;
      }
    });
  });
  return map;
}

function sheetsRepo_getSubmissionsSheet_() {
  return sheetsRepo_ensureSheet_(SHEETS_REPO_SUBMISSIONS, SHEETS_REPO_SUBMISSIONS_HEADERS);
}

function upsertCasesForms_(row) {
  try { Logger.log('upsertCasesForms_: skipped (cases_forms disabled) %s', JSON.stringify(row || {})); } catch (_) {}
  return null;
}

function getCaseForms_(caseId) {
  if (!caseId) return [];
  const want = String(caseId).trim();
  try {
    const props = PropertiesService.getScriptProperties();
    const sid = props.getProperty('BAS_MASTER_SPREADSHEET_ID') || props.getProperty('SHEET_ID');
    if (!sid) return [];
    const ss = SpreadsheetApp.openById(sid);
    const sheet = ss.getSheetByName(SHEETS_REPO_CASES_FORMS);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const idx = (typeof bs_toIndexMap_ === 'function') ? bs_toIndexMap_(headers) : {};
    const ci = idx['case_id'] != null ? idx['case_id'] : idx['caseId'];
    if (!(ci >= 0)) return [];
    const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    return rows
      .filter(function (row) {
        return String(row[ci] || '').trim() === want;
      })
      .map(function (row) {
        const out = {};
        headers.forEach(function (key, idxCol) {
          sheetsRepo_assignAliases_(out, key, row[idxCol]);
        });
        return out;
      });
  } catch (_) {
    return [];
  }
}

function getLastSeq_(caseId, form_key) {
  if (!caseId || !form_key) return 0;
  const sheet = sheetsRepo_ensureSheet_(SHEETS_REPO_SUBMISSIONS, SHEETS_REPO_SUBMISSIONS_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const idx = (typeof bs_toIndexMap_ === 'function') ? bs_toIndexMap_(header) : {};
  const ci = idx['case_id'], fi = idx['form_key'], qi = idx['seq'];
  if (!(fi >= 0 && qi >= 0)) return 0;
  const wantCase = normalizeCaseId_(caseId);
  const wantForm = String(form_key).trim();
  const rows = sheet.getRange(2, 1, lastRow - 1, header.length).getValues();
  let max = 0;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const okForm = String(row[fi] || '').trim() === wantForm;
    const okCase = (ci == null)
      ? true
      : normalizeCaseId_(row[ci] || '') === wantCase;
    if (okForm && okCase) {
      const n = Number(row[qi] || 0);
      if (Number.isFinite(n) && n > max) max = n;
    }
  }
  return max;
}

function insertSubmission_(row) {
  const caseIdRaw = sheetsRepo_getValue_(row, 'case_id');
  const formKeyRaw = sheetsRepo_getValue_(row, 'form_key');
  if (!caseIdRaw || !formKeyRaw) {
    throw new Error('insertSubmission_: missing case_id or form_key');
  }
  const normalizedCaseId = normalizeCaseId_(caseIdRaw);
  if (row && typeof normalizeCaseId_ === 'function') {
    row.case_id = normalizeCaseId_(row.case_id || row.caseId || caseIdRaw);
  }
  const userKeyRaw = sheetsRepo_getValue_(row, 'user_key');
  const normalizedUserKey = normalizeUserKey_(userKeyRaw);
  const sheet = sheetsRepo_getSubmissionsSheet_();
  ensureSubmissionColumns_(sheet, SHEETS_REPO_SUBMISSIONS_HEADERS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const normalized = Object.assign({}, row, {
    case_id: normalizedCaseId,
    form_key: String(formKeyRaw).trim(),
    user_key: normalizedUserKey,
  });
  if (!normalized.case_key && normalized.case_id && normalized.user_key) {
    normalized.case_key = normalized.user_key + '-' + normalized.case_id;
  }
  const ordered = headers.map(function (key) {
    const value = sheetsRepo_getValue_(normalized, key);
    return value !== undefined && value !== null ? value : '';
  });
  sheet.appendRow(ordered);
  if (typeof ensureSubmissionsCaseIdTextFormat_ === 'function') {
    try { ensureSubmissionsCaseIdTextFormat_(); } catch (_) {}
  }
}

/** (submission_id, form_key) で upsert し、任意の追加カラムも反映 */
function upsertSubmission_(row) {
  const caseIdRaw = sheetsRepo_getValue_(row, 'case_id');
  const formKeyRaw = sheetsRepo_getValue_(row, 'form_key');
  const submissionIdRaw = sheetsRepo_getValue_(row, 'submission_id');
  if (!caseIdRaw || !formKeyRaw || !submissionIdRaw) {
    throw new Error('upsertSubmission_: missing case_id/form_key/submission_id');
  }
  const normalizedCaseId = normalizeCaseId_(caseIdRaw);
  if (row && typeof normalizeCaseId_ === 'function') {
    row.case_id = normalizeCaseId_(row.case_id || row.caseId || caseIdRaw);
  }
  const userKeyRaw = sheetsRepo_getValue_(row, 'user_key');
  const normalizedUserKey = normalizeUserKey_(userKeyRaw);
  const sheet = sheetsRepo_getSubmissionsSheet_();
  // 足りない列を保証（列ズレ防止）
  ensureSubmissionColumns_(sheet, SHEETS_REPO_SUBMISSIONS_HEADERS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(v){return String(v||'').trim();});
  const idx = bs_toIndexMap_(headers);
  const lastRow = sheet.getLastRow();
  let foundRow = -1;
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (let i = 0; i < values.length; i++) {
      const r = values[i];
      const fk  = String(r[idx['form_key']] || '').trim();
      const sid = String(r[idx['submission_id']] || '').trim();
      if (fk === String(formKeyRaw).trim() && sid === String(submissionIdRaw).trim()) {
        foundRow = i + 2; // 1-based + header
        break;
      }
    }
  }
  // 正規化（case_id/form_key は snake に揃える）
  const normalized = Object.assign({}, row, {
    case_id: normalizedCaseId,
    form_key: String(formKeyRaw).trim(),
    submission_id: String(submissionIdRaw).trim(),
    user_key: normalizedUserKey,
  });
  if (!normalized.case_key && normalized.case_id && normalized.user_key) {
    normalized.case_key = normalized.user_key + '-' + normalized.case_id;
  }
  if (foundRow > 0) {
    const current = sheet.getRange(foundRow, 1, 1, headers.length).getValues()[0];
    headers.forEach(function (key, colIdx) {
      const val = sheetsRepo_getValue_(normalized, key);
      if (val != null) current[colIdx] = val;
    });
    sheet.getRange(foundRow, 1, 1, headers.length).setValues([current]);
  } else {
    const ordered = headers.map(function (key) {
      const value = sheetsRepo_getValue_(normalized, key);
      return value !== undefined && value !== null ? value : '';
    });
    sheet.appendRow(ordered);
  }
  if (typeof ensureSubmissionsCaseIdTextFormat_ === 'function') {
    try { ensureSubmissionsCaseIdTextFormat_(); } catch (_) {}
  }
}

function ensureSubmissionColumns_(sheet, needed) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(s){return String(s||'').trim();});
  const have = new Set(header.filter(Boolean));
  const add = (needed || []).filter(function(k){ return !have.has(String(k).trim()); });
  if (add.length) {
    sheet.insertColumnsAfter(header.length || 1, add.length);
    const merged = header.concat(add);
    sheet.getRange(1, 1, 1, merged.length).setValues([merged]);
  }
}

function sheetsRepo_deleteAckRow_(caseId, formKey) {
  try {
    var normalizedCaseId = normalizeCaseId_(caseId || '');
    if (!normalizedCaseId || !formKey) return false;
    var sheet = sheetsRepo_getSubmissionsSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    var idx = (typeof bs_toIndexMap_ === 'function') ? bs_toIndexMap_(headers) : sheetsRepo_buildIndexMap_(headers);
    var sidIdx = idx['submission_id'] != null ? idx['submission_id'] : idx['submissionId'];
    if (!(sidIdx >= 0)) return false;
    var rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var target = 'ack:' + normalizedCaseId + ':' + String(formKey || '').trim();
    for (var r = rows.length - 1; r >= 0; r--) {
      var sid = String(rows[r][sidIdx] || '').trim();
      if (!sid) continue;
      if (sid.toLowerCase() === target.toLowerCase()) {
        sheet.deleteRow(r + 2);
        return true;
      }
    }
  } catch (_) {}
  return false;
}

function sheetsRepo_sweepSubmissions_() {
  try {
    var sheet = sheetsRepo_getSubmissionsSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    var idx = (typeof bs_toIndexMap_ === 'function') ? bs_toIndexMap_(headers) : sheetsRepo_buildIndexMap_(headers);
    var sidIdx = idx['submission_id'] != null ? idx['submission_id'] : idx['submissionId'];
    if (!(sidIdx >= 0)) return 0;
    var removed = 0;
    var rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (var r = rows.length - 1; r >= 0; r--) {
      var sid = String(rows[r][sidIdx] || '').trim();
      if (!/^(ack:[\w:-]+|\d+)$/.test(sid)) {
        sheet.deleteRow(r + 2);
        removed++;
      }
    }
    return removed;
  } catch (_) {}
  return 0;
}

function lookupCaseIdByLineId_(lineId) {
  if (!lineId) return null;
  const sheet = bs_getSheet_(SHEETS_REPO_CASES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = (typeof bs_toIndexMap_ === 'function') ? bs_toIndexMap_(headers) : {};
  const li = idx['line_id'] != null ? idx['line_id'] : idx['lineId'];
  const ci = idx['case_id'] != null ? idx['case_id'] : idx['caseId'];
  if (!(li >= 0 && ci >= 0)) return null;
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const want = String(lineId).trim();
  for (let i = 0; i < rows.length; i++) {
    const rowLine = String(rows[i][li] || '').trim();
    if (!rowLine) continue;
    if (rowLine === want) {
      const rawCase = rows[i][ci];
      return typeof bs_normCaseId_ === 'function' ? bs_normCaseId_(rawCase) : normalizeCaseId_(rawCase);
    }
  }
  return null;
}

function sheetsRepo_hasSubmission_(caseId, formKey, submissionId) {
  if (!caseId || !formKey || !submissionId) return false;
  try {
    const sheet = bs_getSheet_(SHEETS_REPO_SUBMISSIONS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const idx = (typeof bs_toIndexMap_ === 'function') ? bs_toIndexMap_(headers) : {};
    const ci = idx['case_id'] != null ? idx['case_id'] : idx['caseId'];
    const fi = idx['form_key'] != null ? idx['form_key'] : idx['formKey'];
    const si = idx['submission_id'] != null ? idx['submission_id'] : idx['submissionId'];
    if (!(ci >= 0 && fi >= 0 && si >= 0)) return false;
    const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    const wantCase = normalizeCaseId_(caseId);
    const wantForm = String(formKey).trim();
    const wantSub = String(submissionId).trim();
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (
        normalizeCaseId_(row[ci] || '') === wantCase &&
        String(row[fi] || '').trim() === wantForm &&
        String(row[si] || '').trim() === wantSub
      ) {
        return true;
      }
    }
    return false;
  } catch (_) {
    return false;
  }
}


/**
 * contacts シートから email 一致で user_key を取得（見つからなければ空文字）。
 */
function lookupUserKeyByEmail_(email) {
  try {
    const sh = bs_getSheet_('contacts');
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return '';
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const idx = (typeof buildHeaderIndexMap_ === 'function') ? buildHeaderIndexMap_(headers) : bs_toIndexMap_(headers);
    const colEmail = idx['email'] != null ? idx['email'] : (idx['Email'] != null ? idx['Email'] : idx['mail']);
    const colUserKey = (idx['user_key'] != null ? idx['user_key'] : idx['userKey']);
    if (colEmail == null || colUserKey == null) return '';
    const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const want = String(email || '').trim().toLowerCase();
    for (let i = 0; i < rows.length; i++) {
      const ev = String(rows[i][colEmail] || '').trim().toLowerCase();
      if (ev && ev === want) {
        const uk = String(rows[i][colUserKey] || '').trim();
        if (uk) return uk;
      }
    }
  } catch (_) {}
  return '';
}

/**
 * contacts シートから user_key 一致で line_id を取得（見つからなければ空文字）。
 */
function lookupLineIdByUserKey_(userKey) {
  try {
    const sh = bs_getSheet_('contacts');
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return '';
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const idx = (typeof buildHeaderIndexMap_ === 'function') ? buildHeaderIndexMap_(headers) : bs_toIndexMap_(headers);
    const colUserKey = (idx['user_key'] != null ? idx['user_key'] : idx['userKey']);
    const colLineId = (idx['line_id'] != null ? idx['line_id'] : idx['lineId']);
    if (colUserKey == null || colLineId == null) return '';
    const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const want = String(userKey || '').trim();
    for (let i = 0; i < rows.length; i++) {
      const uk = String(rows[i][colUserKey] || '').trim();
      if (uk && uk === want) {
        const lid = String(rows[i][colLineId] || '').trim();
        if (lid) return lid;
      }
    }
  } catch (_) {}
  return '';
}

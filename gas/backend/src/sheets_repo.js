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

function upsertCasesForms_(row) {
  const caseIdRaw = sheetsRepo_getValue_(row, 'case_id');
  const formKeyRaw = sheetsRepo_getValue_(row, 'form_key');
  if (!caseIdRaw || !formKeyRaw) {
    throw new Error('upsertCasesForms_: case_id and form_key are required');
  }
  const caseId = String(caseIdRaw).trim();
  const formKey = String(formKeyRaw).trim();
  const sheet = sheetsRepo_ensureSheet_(SHEETS_REPO_CASES_FORMS, SHEETS_REPO_CASES_FORMS_HEADERS);
  const headers = SHEETS_REPO_CASES_FORMS_HEADERS;
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  let targetRow = -1;
  if (lastRow >= 2) {
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = range.getValues();
    const caseIdx = headers.indexOf('case_id');
    const formIdx = headers.indexOf('form_key');
    for (let i = 0; i < values.length; i++) {
      const rowCase = String(values[i][caseIdx] || '').trim();
      const rowForm = String(values[i][formIdx] || '').trim();
      if (rowCase === caseId && rowForm === formKey) {
        targetRow = i + 2;
        break;
      }
    }
  }

  const baseRecord = {
    status: '',
    can_edit: false,
    reopened_at: '',
    reopened_by: '',
    locked_reason: '',
    reopen_until: '',
    updated_at: new Date().toISOString(),
  };

  const normalized = {};
  headers.forEach(function (key) {
    const incoming = sheetsRepo_getValue_(row, key);
    if (incoming !== undefined && incoming !== null) {
      normalized[key] = incoming;
    }
  });

  normalized['case_id'] = caseId;
  normalized['form_key'] = formKey;
  normalized['updated_at'] = new Date().toISOString();

  const payload = Object.assign({}, baseRecord, normalized);

  if (targetRow > 0) {
    const currentValues = sheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];
    const merged = headers.map(function (key, idx) {
      const incoming = Object.prototype.hasOwnProperty.call(payload, key) ? payload[key] : undefined;
      if (incoming === undefined || incoming === null || incoming === '') {
        return currentValues[idx];
      }
      return incoming;
    });
    sheet.getRange(targetRow, 1, 1, headers.length).setValues([merged]);
  } else {
    const ordered = headers.map(function (key) {
      return Object.prototype.hasOwnProperty.call(payload, key) ? payload[key] : '';
    });
    sheet.appendRow(ordered);
  }
}

function getCaseForms_(caseId) {
  if (!caseId) return [];
  const want = String(caseId).trim();
  return sheetsRepo_readAll_(SHEETS_REPO_CASES_FORMS, SHEETS_REPO_CASES_FORMS_HEADERS).filter(function (row) {
    return String(sheetsRepo_getValue_(row, 'case_id') || '').trim() === want;
  });
}

function getLastSeq_(caseId, form_key) {
  if (!caseId || !form_key) return 0;
  const wantCase = String(caseId).trim();
  const wantForm = String(form_key).trim();
  const rows = sheetsRepo_readAll_(SHEETS_REPO_SUBMISSIONS, SHEETS_REPO_SUBMISSIONS_HEADERS).filter(function (row) {
    return (
      String(sheetsRepo_getValue_(row, 'case_id') || '').trim() === wantCase &&
      String(sheetsRepo_getValue_(row, 'form_key') || '').trim() === wantForm
    );
  });
  if (!rows.length) return 0;
  const seqs = rows
    .map(function (row) {
      return Number(row.seq || 0);
    })
    .filter(function (num) {
      return Number.isFinite(num);
    });
  if (!seqs.length) return 0;
  return Math.max.apply(null, seqs);
}

function insertSubmission_(row) {
  const caseIdRaw = sheetsRepo_getValue_(row, 'case_id');
  const formKeyRaw = sheetsRepo_getValue_(row, 'form_key');
  if (!caseIdRaw || !formKeyRaw) {
    throw new Error('insertSubmission_: missing case_id or form_key');
  }
  const sheet = sheetsRepo_ensureSheet_(SHEETS_REPO_SUBMISSIONS, SHEETS_REPO_SUBMISSIONS_HEADERS);
  const headers = SHEETS_REPO_SUBMISSIONS_HEADERS;
  const normalized = Object.assign({}, row, {
    case_id: String(caseIdRaw).trim(),
    form_key: String(formKeyRaw).trim(),
  });
  const ordered = headers.map(function (key) {
    const value = sheetsRepo_getValue_(normalized, key);
    return value !== undefined && value !== null ? value : '';
  });
  sheet.appendRow(ordered);
}

function lookupCaseIdByLineId_(lineId) {
  if (!lineId) return null;
  const sheet = bs_getSheet_(SHEETS_REPO_CASES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = bs_toIndexMap_(headers);
  if (!(idx.lineId >= 0 && idx.caseId >= 0)) return null;
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (let i = 0; i < rows.length; i++) {
    const rowLine = String(rows[i][idx.lineId] || '').trim();
    if (!rowLine) continue;
    if (rowLine === String(lineId).trim()) {
      const rawCase = rows[i][idx.caseId];
      return typeof bs_normCaseId_ === 'function' ? bs_normCaseId_(rawCase) : rawCase;
    }
  }
  return null;
}

function sheetsRepo_hasSubmission_(caseId, formKey, submissionId) {
  if (!caseId || !formKey || !submissionId) return false;
  const rows = sheetsRepo_readAll_(SHEETS_REPO_SUBMISSIONS, SHEETS_REPO_SUBMISSIONS_HEADERS);
  return rows.some(function (row) {
    return (
      String(sheetsRepo_getValue_(row, 'case_id') || '').trim() === String(caseId).trim() &&
      String(sheetsRepo_getValue_(row, 'form_key') || '').trim() === String(formKey).trim() &&
      String(sheetsRepo_getValue_(row, 'submission_id') || '').trim() === String(submissionId).trim()
    );
  });
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

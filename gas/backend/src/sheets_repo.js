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
  'caseId',
  'form_key',
  'status',
  'canEdit',
  'reopened_at',
  'reopened_by',
  'locked_reason',
  'reopen_until',
  'updated_at',
];

const SHEETS_REPO_SUBMISSIONS_HEADERS = [
  'caseId',
  'form_key',
  'seq',
  'submission_id',
  'received_at',
  'supersedes_seq',
  'json_path',
];

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
      obj[key] = row[idx];
    });
    return obj;
  });
}

function upsertCasesForms_(row) {
  if (!row || !row.caseId || !row.form_key) {
    throw new Error('upsertCasesForms_: caseId and form_key are required');
  }
  const sheet = sheetsRepo_ensureSheet_(SHEETS_REPO_CASES_FORMS, SHEETS_REPO_CASES_FORMS_HEADERS);
  const headers = SHEETS_REPO_CASES_FORMS_HEADERS;
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  let targetRow = -1;
  if (lastRow >= 2) {
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = range.getValues();
    const caseIdx = headers.indexOf('caseId');
    const formIdx = headers.indexOf('form_key');
    for (let i = 0; i < values.length; i++) {
      const rowCase = String(values[i][caseIdx] || '').trim();
      const rowForm = String(values[i][formIdx] || '').trim();
      if (rowCase === String(row.caseId).trim() && rowForm === String(row.form_key).trim()) {
        targetRow = i + 2;
        break;
      }
    }
  }

  const baseRecord = {
    status: '',
    canEdit: false,
    reopened_at: '',
    reopened_by: '',
    locked_reason: '',
    reopen_until: '',
    updated_at: new Date(),
  };
  const payload = Object.assign({}, baseRecord, row, { updated_at: new Date() });

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
  return sheetsRepo_readAll_(SHEETS_REPO_CASES_FORMS, SHEETS_REPO_CASES_FORMS_HEADERS).filter(function (row) {
    return String(row.caseId || '').trim() === String(caseId).trim();
  });
}

function getLastSeq_(caseId, form_key) {
  if (!caseId || !form_key) return 0;
  const rows = sheetsRepo_readAll_(SHEETS_REPO_SUBMISSIONS, SHEETS_REPO_SUBMISSIONS_HEADERS).filter(function (row) {
    return (
      String(row.caseId || '').trim() === String(caseId).trim() &&
      String(row.form_key || '').trim() === String(form_key).trim()
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
  if (!row || !row.caseId || !row.form_key) {
    throw new Error('insertSubmission_: missing caseId or form_key');
  }
  const sheet = sheetsRepo_ensureSheet_(SHEETS_REPO_SUBMISSIONS, SHEETS_REPO_SUBMISSIONS_HEADERS);
  const headers = SHEETS_REPO_SUBMISSIONS_HEADERS;
  const ordered = headers.map(function (key) {
    return Object.prototype.hasOwnProperty.call(row, key) ? row[key] : '';
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

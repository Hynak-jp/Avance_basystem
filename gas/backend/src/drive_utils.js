/**
 * drive_utils.js
 *  - Drive/Sheets 共通ユーティリティ
 *  - ここで定義した関数は他ファイルからグローバルに参照できます
 */

const DRIVE_UTILS_PROPS = PropertiesService.getScriptProperties();

function drive_getRootFolderId_() {
  const id =
    DRIVE_UTILS_PROPS.getProperty('DRIVE_ROOT_FOLDER_ID') ||
    DRIVE_UTILS_PROPS.getProperty('ROOT_FOLDER_ID');
  if (!id) throw new Error('DRIVE_ROOT_FOLDER_ID/ROOT_FOLDER_ID is not configured');
  return id;
}

function drive_getRootFolder_() {
  return DriveApp.getFolderById(drive_getRootFolderId_());
}

function drive_normalizePathPart_(part) {
  return String(part || '')
    .replace(/[\\/:*?"<>|]/g, '_')
    .trim();
}

function drive_getOrCreatePath_(parent, path) {
  if (!parent) throw new Error('drive_getOrCreatePath_: parent is required');
  const parts = String(path || '')
    .split('/')
    .map((p) => drive_normalizePathPart_(p))
    .filter(Boolean);
  let current = parent;
  for (const name of parts) {
    const it = current.getFoldersByName(name);
    current = it.hasNext() ? it.next() : current.createFolder(name);
  }
  return current;
}

function drive_userKeyFromLineId_(lineId) {
  const lid = String(lineId || '').trim();
  if (!lid) return '';
  try {
    const hit = drive_lookupCaseRow_({ lineId: lid });
    const userKey = hit && hit.userKey ? String(hit.userKey).trim() : '';
    if (userKey) return userKey.toLowerCase();
  } catch (_) {}
  return lid.slice(0, 6).toLowerCase();
}

function drive_lookupCaseRow_(criteria) {
  const SHEET_CASES = DRIVE_UTILS_PROPS.getProperty('SHEET_CASES') || 'cases';
  const sh = bs_getSheet_(SHEET_CASES);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = bs_toIndexMap_(headers);
  const rows = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const caseId = criteria.caseId ? bs_normCaseId_(criteria.caseId) : '';
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowCaseId = bs_normCaseId_(row[idx['caseId']] || row[idx['case_id']] || '');
    if (caseId && rowCaseId !== caseId) continue;
    if (criteria.lineId && idx['lineId'] >= 0) {
      if (String(row[idx['lineId']] || '').trim() !== String(criteria.lineId).trim()) continue;
    }
    return {
      caseId: rowCaseId,
      userKey:
        idx['userKey'] >= 0
          ? String(row[idx['userKey']] || row[idx['user_key']] || '').trim()
          : idx['user_key'] >= 0
          ? String(row[idx['user_key']] || '').trim()
          : '',
      lineId:
        idx['lineId'] >= 0
          ? String(row[idx['lineId']] || row[idx['line_id']] || '').trim()
          : idx['line_id'] >= 0
          ? String(row[idx['line_id']] || '').trim()
          : '',
    };
  }
  return null;
}

function drive_resolveCaseKeyFromMeta_(meta, fallback) {
  const m = meta || {};
  const fb = fallback || {};
  const normCase = bs_normCaseId_ || function (v) { return String(v || '').trim(); };
  const caseId = normCase(m.case_id || m.caseId || fb.case_id || fb.caseId || '');
  if (!caseId) throw new Error('drive_resolveCaseKeyFromMeta_: case_id is required');

  const direct = m.case_key || m.caseKey || fb.case_key || fb.caseKey;
  if (direct && direct.indexOf('-') >= 0) return direct;

  let userKey = m.user_key || m.userKey || fb.user_key || fb.userKey || '';
  let lineId = m.line_id || m.lineId || fb.line_id || fb.lineId || '';

  if (!userKey && lineId) {
    userKey = drive_userKeyFromLineId_(lineId);
  }

  if (!userKey) {
    const row = drive_lookupCaseRow_({ caseId: caseId, lineId: lineId });
    if (row && row.userKey) userKey = row.userKey;
  }

  if (userKey) return userKey + '-' + caseId;

  throw new Error('drive_resolveCaseKeyFromMeta_: cannot resolve caseKey');
}

function drive_getOrCreateCaseFolderByKey_(caseKey) {
  const root = drive_getRootFolder_();
  return drive_getOrCreatePath_(root, caseKey);
}

function drive_getOrCreateEmailStagingFolder_() {
  const root = drive_getRootFolder_();
  const it = root.getFoldersByName('_email_staging');
  return it.hasNext() ? it.next() : root.createFolder('_email_staging');
}

function drive_moveFileToFolder_(file, targetFolder) {
  if (!file || !targetFolder) return file;
  const fileId = typeof file === 'string' ? file : file.getId();
  const targetId = typeof targetFolder === 'string' ? targetFolder : targetFolder.getId();
  let removeParents = '';
  try {
    const parents = DriveApp.getFileById(fileId).getParents();
    const ids = [];
    while (parents.hasNext()) ids.push(parents.next().getId());
    removeParents = ids.filter((id) => id !== targetId).join(',');
  } catch (_) {}
  const request = {
    addParents: targetId,
    supportsAllDrives: true,
    supportsTeamDrives: true,
  };
  if (removeParents) request.removeParents = removeParents;
  Drive.Files.update({}, fileId, null, request);
  return DriveApp.getFileById(fileId);
}

function drive_placeFileIntoCase_(file, meta, fallback) {
  if (!file) return;
  const caseKey = drive_resolveCaseKeyFromMeta_(meta || {}, fallback || {});
  const caseFolder = drive_getOrCreateCaseFolderByKey_(caseKey);
  return drive_moveFileToFolder_(file, caseFolder);
}

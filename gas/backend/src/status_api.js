/**
 * status_api.js
 *  - WebApp doGet ルート（フォーム受領状況・再開操作）
 *  - HMAC 署名: sig = HMAC_SHA256(`${ts}.${lineId}.${caseId}`, HMAC_SECRET)
 */

const STATUS_API_PROPS = PropertiesService.getScriptProperties();
const STATUS_API_SECRET =
  STATUS_API_PROPS.getProperty('HMAC_SECRET') || STATUS_API_PROPS.getProperty('BAS_API_HMAC_SECRET') || '';
const STATUS_API_NONCE_WINDOW_SECONDS = 600; // 10 分

function statusApi_hex_(bytes) {
  return bytes
    .map(function (b) {
      const v = b & 0xff;
      return (v < 16 ? '0' : '') + v.toString(16);
    })
    .join('');
}

function statusApi_hmac_(message, secret) {
  const raw = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256,
    message,
    secret
  );
  return statusApi_hex_(raw).toLowerCase();
}

function statusApi_verify_(params) {
  if (!STATUS_API_SECRET) throw new Error('hmac_secret_not_configured');
  const tsRaw = String(params.ts || '').trim();
  const sig = String(params.sig || '').trim();
  const lineId = String(params.lineId || '').trim();
  const caseId = String(params.caseId || params.case_id || '').trim();
  if (!tsRaw || !sig) throw new Error('missing ts or sig');
  const tsNum = Number(tsRaw);
  if (!Number.isFinite(tsNum)) throw new Error('invalid timestamp');
  const skew = Math.abs(Date.now() - tsNum);
  if (skew > 10 * 60 * 1000) throw new Error('timestamp too far');
  const message = tsRaw + '.' + lineId + '.' + caseId;
  const expected = statusApi_hmac_(message, STATUS_API_SECRET);
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

function statusApi_normalizeBool_(v) {
  if (typeof v === 'boolean') return v;
  const s = String(v || '').trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'yes';
}

function statusApi_collectStaging_(lineId, caseId) {
  const lid = String(lineId || '').trim();
  const cid = String(caseId || '').trim();
  if (!cid) return;
  try {
    if (typeof bs_collectIntakeFromStaging_ === 'function') {
      bs_collectIntakeFromStaging_(lid, cid);
    } else if (typeof moveStagingIntakeJsonToCase_ === 'function') {
      moveStagingIntakeJsonToCase_(lid, cid);
    }
  } catch (err) {
    try {
      Logger.log('[status_api] staging sweep error: %s', (err && err.stack) || err);
    } catch (_) {}
  }
}

/** GET /exec?action=status&caseId=&lineId=&ts=&sig= */
function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || 'status');
  try {
    statusApi_verify_(e.parameter || {});
  } catch (err) {
    return statusApi_jsonOut_({ ok: false, error: String(err) }, 400);
  }

  try {
    if (action === 'status') {
      const lineId = String((e.parameter || {}).lineId || '').trim();
      let caseId = String((e.parameter || {}).caseId || '').trim();
      if (!caseId) caseId = lookupCaseIdByLineId_(lineId) || '';
      if (!caseId) {
        return statusApi_jsonOut_({ ok: false, error: 'caseId not found' }, 404);
      }
      statusApi_collectStaging_(lineId, caseId);
      const forms = getCaseForms_(caseId).map(function (row) {
        const canEdit = statusApi_normalizeBool_(row.canEdit);
        const formKey = String(row.form_key || '').trim();
        return {
          caseId: String(row.caseId || caseId || ''),
          form_key: formKey,
          status: row.status || '',
          canEdit: canEdit,
          reopened_at: row.reopened_at || null,
          locked_reason: row.locked_reason || null,
          reopen_until: row.reopen_until || null,
          last_seq: formKey ? getLastSeq_(caseId, formKey) : 0,
        };
      });
      return statusApi_jsonOut_({ ok: true, caseId: caseId, forms: forms }, 200);
    }

    if (action === 'markReopen') {
      return statusApi_jsonOut_({ ok: false, error: 'use_post' }, 405);
    }

    return statusApi_jsonOut_({ ok: false, error: 'unknown action' }, 400);
  } catch (err) {
    return statusApi_jsonOut_({ ok: false, error: String(err) }, 400);
  }
}

function statusApi_handleMarkReopenPost_(body) {
  try {
    statusApi_verify_(body || {});
  } catch (err) {
    return statusApi_jsonOut_({ ok: false, error: String(err) }, 400);
  }

  const lineId = String((body || {}).lineId || '').trim();
  const tsRaw = String((body || {}).ts || '').trim();
  const sig = String((body || {}).sig || '').trim();

  try {
    statusApi_assertNonce_(lineId, tsRaw, sig);
  } catch (err) {
    return statusApi_jsonOut_({ ok: false, error: String(err) }, 409);
  }

  if (!lineId) {
    return statusApi_jsonOut_({ ok: false, error: 'missing lineId' }, 400);
  }

  const formKey = String((body || {}).form_key || (body || {}).formKey || '').trim();
  if (!formKey) {
    return statusApi_jsonOut_({ ok: false, error: 'missing form_key' }, 400);
  }

  const caseId = lookupCaseIdByLineId_(lineId);
  if (!caseId) {
    return statusApi_jsonOut_({ ok: false, error: 'case_not_found' }, 404);
  }
  statusApi_collectStaging_(lineId, caseId);

  upsertCasesForms_({
    case_id: caseId,
    form_key: formKey,
    status: 'reopened',
    can_edit: true,
    reopened_at: new Date().toISOString(),
    reopened_by: ((body || {}).staff || 'staff').toString(),
    reopen_until: (body || {}).reopen_until || (body || {}).reopenUntil || '',
  });
  return statusApi_jsonOut_({ ok: true, caseId: caseId }, 200);
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
  const caseId = String(caseIdRaw).trim();
  const formKey = String(formKeyRaw).trim();
  if (!caseId || !formKey) return;

  const submissionId = String(payload.submission_id || '').trim();
  const fallbackInfo = {
    caseKey: payload.case_key || payload.caseKey,
    caseId: caseId,
    userKey: payload.user_key || payload.userKey,
    lineId: payload.line_id || payload.lineId,
  };
  let caseKey = '';
  try {
    caseKey = drive_resolveCaseKeyFromMeta_(payload.meta || {}, fallbackInfo);
  } catch (_) {
    var metaLineId = fallbackInfo.lineId || '';
    if (!metaLineId && payload.meta && payload.meta.line_id) metaLineId = payload.meta.line_id;
    if (!metaLineId && payload.meta && payload.meta.lineId) metaLineId = payload.meta.lineId;
    const row = drive_lookupCaseRow_({ caseId: caseId, lineId: metaLineId });
    if (row && row.userKey) {
      fallbackInfo.userKey = row.userKey;
      caseKey = row.userKey + '-' + caseId;
    }
  }
  if (!caseKey && fallbackInfo.userKey) {
    caseKey = fallbackInfo.userKey + '-' + caseId;
  }
  if (!caseKey && fallbackInfo.lineId) {
    const inferredKey = drive_userKeyFromLineId_(fallbackInfo.lineId);
    if (inferredKey) caseKey = inferredKey + '-' + caseId;
  }

  let caseFolderId = '';
  if (caseKey) {
    try {
      const caseFolder = drive_getOrCreateCaseFolderByKey_(caseKey);
      caseFolderId = caseFolder.getId();
      if (typeof updateCasesRow_ === 'function') {
        const patch = { case_key: caseKey, folder_id: caseFolderId };
        if (!fallbackInfo.userKey && caseKey.indexOf('-') >= 0) {
          patch.user_key = caseKey.split('-')[0];
        }
        updateCasesRow_(caseId, patch);
      }
    } catch (err) {
      Logger.log('[Intake] ensure case folder error: %s', (err && err.stack) || err);
    }
  }
  try {
    if (submissionId && sheetsRepo_hasSubmission_(caseId, formKey, submissionId)) {
      upsertCasesForms_({
        case_id: caseId,
        form_key: formKey,
        status: 'submitted',
        can_edit: false,
        locked_reason: payload.locked_reason || '',
        updated_at: new Date().toISOString(),
      });
      return;
    }
  } catch (err) {
    Logger.log('[Intake] recordSubmission_ duplicate check error: %s', (err && err.stack) || err);
  }

  let last = 0;
  try {
    last = Number(getLastSeq_(caseId, formKey)) || 0;
  } catch (err) {
    last = 0;
  }
  const receivedAt = payload.received_at
    ? new Date(payload.received_at).toISOString()
    : new Date().toISOString();
  insertSubmission_({
    case_id: caseId,
    case_key: caseKey,
    form_key: formKey,
    seq: last + 1,
    submission_id: submissionId,
    received_at: receivedAt,
    supersedes_seq: last || '',
    json_path: payload.json_path || '',
  });
  upsertCasesForms_({
    case_id: caseId,
    form_key: formKey,
    status: 'submitted',
    can_edit: false,
    locked_reason: payload.locked_reason || '',
    updated_at: new Date().toISOString(),
  });
}

/** POST /exec でのルーティング */
function doPost(e) {
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

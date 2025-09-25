/**
 * status_api.js
 *  - WebApp doGet ルート（フォーム受領状況・再開操作）
 *  - HMAC 署名: sig = HMAC_SHA256(`${ts}.${lineId}.${caseId}`, HMAC_SECRET)
 */

const STATUS_API_PROPS = PropertiesService.getScriptProperties();
const STATUS_API_SECRET =
  STATUS_API_PROPS.getProperty('HMAC_SECRET') || STATUS_API_PROPS.getProperty('BAS_API_HMAC_SECRET') || '';

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
  const caseId = String(params.caseId || '').trim();
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
      const params = e.parameter || {};
      const caseId = String(params.caseId || '').trim();
      const formKey = String(params.form_key || '').trim();
      const staff = String(params.staff || 'staff').trim();
      if (!caseId || !formKey) {
        return statusApi_jsonOut_({ ok: false, error: 'missing caseId or form_key' }, 400);
      }
      upsertCasesForms_({
        caseId: caseId,
        form_key: formKey,
        status: 'reopened',
        canEdit: true,
        reopened_at: new Date(),
        reopened_by: staff,
      });
      return statusApi_jsonOut_({ ok: true }, 200);
    }

    return statusApi_jsonOut_({ ok: false, error: 'unknown action' }, 400);
  } catch (err) {
    return statusApi_jsonOut_({ ok: false, error: String(err) }, 400);
  }
}

/**
 * JSON 受領後に呼び出し、submissions / cases_forms を更新
 * @param {{caseId:string, form_key:string, submission_id?:string, json_path?:string, received_at?:Date, locked_reason?:string}} payload
 */
function recordSubmission_(payload) {
  if (!payload || !payload.caseId || !payload.form_key) return;
  const caseId = String(payload.caseId).trim();
  const formKey = String(payload.form_key).trim();
  if (!caseId || !formKey) return;
  var last = 0;
  try {
    last = Number(getLastSeq_(caseId, formKey)) || 0;
  } catch (err) {
    last = 0;
  }
  const next = last + 1;
  insertSubmission_({
    caseId: caseId,
    form_key: formKey,
    seq: next,
    submission_id: payload.submission_id || '',
    received_at: payload.received_at || new Date(),
    supersedes_seq: last || '',
    json_path: payload.json_path || '',
  });
  upsertCasesForms_({
    caseId: caseId,
    form_key: formKey,
    status: 'submitted',
    canEdit: false,
    locked_reason: payload.locked_reason || '',
  });
}

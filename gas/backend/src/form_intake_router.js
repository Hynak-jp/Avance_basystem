/**
 * 共通フォーム Intake ルーター
 *  - FormAttach/ToProcess に入った通知メールをフォーム種別ごとに振り分け、caseフォルダへ JSON 保存
 *  - FORM_REGISTRY にエントリを追加するだけで新フォームを取り込める
 */

const FORM_INTAKE_LABEL_TO_PROCESS =
  typeof LABEL_TO_PROCESS === 'string' ? LABEL_TO_PROCESS : 'FormAttach/ToProcess';
const FORM_INTAKE_LABEL_PROCESSED =
  typeof LABEL_PROCESSED === 'string' ? LABEL_PROCESSED : 'FormAttach/Processed';
const FORM_INTAKE_LABEL_LOCK = 'BAS/lock';

function resolveCaseByCaseIdSmart_(caseId) {
  const raw = String(caseId || '').trim();
  const noPad = raw.replace(/^0+/, '');
  const pad4 = noPad.padStart(4, '0');
  if (typeof resolveCaseByCaseId_ !== 'function') {
    throw new Error('resolveCaseByCaseId_ is not defined');
  }
  return (
    resolveCaseByCaseId_(raw) ||
    (noPad ? resolveCaseByCaseId_(noPad) : null) ||
    (pad4 ? resolveCaseByCaseId_(pad4) : null) ||
    null
  );
}

const FORM_INTAKE_REGISTRY = {
  intake: {
    name: '初回受付',
    queueLabel: 'BAS/Intake/Queue',
    parser: function (subject, body) {
      if (typeof parseMetaBlock_ !== 'function') {
        throw new Error('parseMetaBlock_ is not defined');
      }
      const meta = parseMetaBlock_(body) || {};
      if (!meta.form_key) meta.form_key = 'intake';
      return { meta, fieldsRaw: [], model: {} };
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    statusAfterSave: 'intake',
    afterSave: function () { return {}; },
  },

  s2010_p1_career: {
    name: 'S2010 Part1(経歴等)',
    queueLabel: 'BAS/S2010P1/Queue',
    parser: function (subject, body) {
      if (typeof parseFormMail_S2010_P1_ !== 'function') {
        throw new Error('parseFormMail_S2010_P1_ is not defined');
      }
      return parseFormMail_S2010_P1_(subject, body);
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    statusAfterSave: 'intake',
    afterSave: function () {
      return {};
    },
  },
  s2002_userform: {
    name: 'S2002 申立',
    queueLabel: 'BAS/S2002/Queue',
    parser: function (subject, body) {
      if (typeof parseFormMail_ !== 'function') {
        throw new Error('parseFormMail_ is not defined');
      }
      return parseFormMail_(subject, body);
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    statusAfterSave: 'draft',
    afterSave: function (caseInfo, parsed) {
      if (typeof generateS2002Draft_ === 'function') {
        const draft = generateS2002Draft_(caseInfo, parsed);
        const patch = {};
        if (draft && draft.draftUrl) patch.last_draft_url = draft.draftUrl;
        return patch;
      }
      return {};
    },
  },
  s2010_p2_cause: {
    name: 'S2010 Part2(申立てに至った事情)',
    queueLabel: 'BAS/S2010P2/Queue',
    parser: function (subject, body) {
      if (typeof parseFormMail_S2010_P2_ !== 'function') {
        throw new Error('parseFormMail_S2010_P2_ is not defined');
      }
      return parseFormMail_S2010_P2_(subject, body);
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    statusAfterSave: 'intake',
    afterSave: function () {
      return {};
    },
  },
};

const FORM_INTAKE_QUEUE_LABELS = Object.freeze(
  Array.from(new Set(Object.values(FORM_INTAKE_REGISTRY).map((def) => def.queueLabel)))
);

const FORM_INTAKE_LABEL_CACHE = {};

function formIntake_normalizeCaseId_(value) {
  if (typeof bs_normCaseId_ === 'function') return bs_normCaseId_(value);
  const digits = String(value || '').replace(/\D/g, '');
  if (!digits) return '';
  return digits.padStart(4, '0');
}

function formIntake_normalizeUserKey_(value) {
  const raw = String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  if (!raw) return '';
  if (raw.length >= 6) return raw.slice(0, 6);
  return (raw + 'xxxxxx').slice(0, 6);
}

function formIntake_generateUserKey_(meta) {
  // intake はメールだけでは user_key を決めない（誤決定を防ぐ）
  try {
    const fk = String((meta && (meta.form_key || meta.formKey)) || '').trim();
    if (fk === 'intake') return '';
  } catch (_) {}

  const sources = [
    meta && meta.user_key,
    meta && meta.userKey,
    meta && meta.email,
    meta && meta.Email,
    meta && meta.mail,
    meta && meta.submission_id,
    meta && meta.submissionId,
  ];
  for (let i = 0; i < sources.length; i++) {
    const normal = formIntake_normalizeUserKey_(sources[i]);
    if (normal) return normal;
  }
  const seed =
    typeof Utilities !== 'undefined' && typeof Utilities.getUuid === 'function'
      ? Utilities.getUuid()
      : Date.now().toString(36) + Math.random().toString(36).slice(2);
  return formIntake_normalizeUserKey_(seed);
}

function formIntake_issueNewCase_(meta) {
  if (typeof bs_issueCaseId_ !== 'function') {
    throw new Error('case_id を採番できません (bs_issueCaseId_ 未定義)');
  }
  const userKeyHint = formIntake_normalizeUserKey_(meta && (meta.user_key || meta.userKey));
  const userKey = userKeyHint || formIntake_generateUserKey_(meta);
  const lineId = String((meta && (meta.line_id || meta.lineId)) || '').trim();
  const issued = bs_issueCaseId_(userKey, lineId);
  const caseId = formIntake_normalizeCaseId_(issued && issued.caseId);
  if (!caseId) throw new Error('case_id の採番に失敗しました');
  return { caseId, userKey, lineId };
}

function formIntake_prepareCaseInfo_(meta, def, parsed) {
  const metaObj = meta || {};
  if (parsed && !parsed.meta) parsed.meta = metaObj;
  let caseId = formIntake_normalizeCaseId_(metaObj.case_id || metaObj.caseId);
  let userKey = formIntake_normalizeUserKey_(metaObj.user_key || metaObj.userKey);
  let lineId = String(metaObj.line_id || metaObj.lineId || '').trim();
  if (!userKey && metaObj.email && typeof lookupUserKeyByEmail_ === 'function') {
    userKey = formIntake_normalizeUserKey_(lookupUserKeyByEmail_(metaObj.email));
  }
  if (!lineId && typeof lookupLineIdByUserKey_ === 'function' && userKey) {
    lineId = String(lookupLineIdByUserKey_(userKey) || '').trim();
  }
  if (!caseId) {
    const issued = formIntake_issueNewCase_(metaObj);
    caseId = issued.caseId;
    userKey = issued.userKey || userKey;
    lineId = issued.lineId || lineId;
    metaObj.case_id = caseId;
    metaObj.caseId = caseId;
    if (userKey && !metaObj.user_key) metaObj.user_key = userKey;
    if (userKey && !metaObj.userKey) metaObj.userKey = userKey;
  }
  const caseInfo = formIntake_resolveCase_(caseId, def);
  caseInfo.caseId = caseInfo.caseId || caseId;
  if (!caseInfo.userKey && userKey) caseInfo.userKey = userKey;
  if (!caseInfo.lineId && lineId) caseInfo.lineId = lineId;
  return {
    caseInfo,
    caseId,
    userKey: caseInfo.userKey || userKey || '',
    lineId: caseInfo.lineId || lineId || '',
  };
}

function formIntake_labelOrCreate_(name) {
  if (!name) return null;
  if (FORM_INTAKE_LABEL_CACHE[name]) return FORM_INTAKE_LABEL_CACHE[name];
  const label = GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
  FORM_INTAKE_LABEL_CACHE[name] = label;
  return label;
}

function getFormQueueLabels_() {
  return FORM_INTAKE_QUEUE_LABELS;
}

function getFormLockLabel_() {
  return FORM_INTAKE_LABEL_LOCK;
}

function run_ProcessInbox_AllForms() {
  formIntake_assignQueueLabels_();

  const lockLabel = formIntake_labelOrCreate_(FORM_INTAKE_LABEL_LOCK);
  const processedLabel = formIntake_labelOrCreate_(FORM_INTAKE_LABEL_PROCESSED);
  const toProcessLabel = formIntake_labelOrCreate_(FORM_INTAKE_LABEL_TO_PROCESS);

  Object.keys(FORM_INTAKE_REGISTRY).forEach(function (formKey) {
    const def = FORM_INTAKE_REGISTRY[formKey];
    const queueLabel = formIntake_labelOrCreate_(def.queueLabel);
    const query = `label:${def.queueLabel} -label:${FORM_INTAKE_LABEL_LOCK}`;
    const threads = GmailApp.search(query, 0, 50);
    threads.forEach(function (thread) {
      const messages = thread.getMessages();
      if (!messages || !messages.length) return;

      const msg = messages[0];
      let locked = false;
      try {
        thread.addLabel(lockLabel);
        locked = true;

        const body = msg.getPlainBody() || HtmlService.createHtmlOutput(msg.getBody()).getContent();
        const subject = msg.getSubject();
        const parsed = def.parser(subject, body);
        const meta = parsed?.meta || {};
        const actualKey = String(meta.form_key || '').trim();
        if (actualKey !== formKey) {
          thread.removeLabel(queueLabel);
          if (locked) thread.removeLabel(lockLabel);
          return;
        }

        // ===== intake はケース採番・保存の前に、まず staging へ保存して終了（重複採番防止） =====
        const formKeyForStatusEarly = String(meta.form_key || actualKey || '').trim();
        if (formKeyForStatusEarly === 'intake') {
          try {
            const P = PropertiesService.getScriptProperties();
            const EXPECT = String(P.getProperty('NOTIFY_SECRET') || '').trim();
            const ALLOW_NO_SECRET = (P.getProperty('ALLOW_NO_SECRET') || '').toLowerCase() === '1';
            const provided = String((meta && meta.secret) || '').trim();
            if (EXPECT && !ALLOW_NO_SECRET && provided !== EXPECT) {
              try { Logger.log('[Intake] secret mismatch (early): meta.secret=%s', provided || '(empty)'); } catch (_) {}
              formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, false);
              return;
            }
          } catch (_) {}

          if (!meta.submission_id) {
            meta.submission_id =
              meta.submissionId ||
              (typeof Utilities !== 'undefined' && typeof Utilities.getUuid === 'function'
                ? Utilities.getUuid()
                : String(Date.now()));
            if (!meta.submissionId) meta.submissionId = meta.submission_id;
            parsed.meta = meta;
          }

          const P = PropertiesService.getScriptProperties();
          const ROOT_ID = P.getProperty('DRIVE_ROOT_FOLDER_ID') || P.getProperty('ROOT_FOLDER_ID');
          const root = DriveApp.getFolderById(ROOT_ID);
          const itSt = root.getFoldersByName('_email_staging');
          const staging = itSt.hasNext() ? itSt.next() : root.createFolder('_email_staging');
          const fname = `intake__${meta.submission_id}.json`;
          const blob = Utilities.newBlob(JSON.stringify(parsed, null, 2), 'application/json', fname);
          staging.createFile(blob);
          try { Logger.log('[Intake] early staged %s', fname); } catch (_) {}
          formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, true);
          return;
        }

        const prepared = formIntake_prepareCaseInfo_(meta, def, parsed);
        const caseInfo = prepared.caseInfo || {};
        const caseId = prepared.caseId;
        const case_id = caseId;
        caseInfo.caseId = caseInfo.caseId || caseId;
        if (!caseInfo.case_id) caseInfo.case_id = caseInfo.caseId;

        // intake（meta.case_id 無し）への採番結果を同じ参照に反映
        if (!meta.case_id) meta.case_id = caseId;
        if (!meta.caseId) meta.caseId = caseId;
        if (!meta.submission_id) {
          meta.submission_id =
            meta.submissionId ||
            (typeof Utilities !== 'undefined' && typeof Utilities.getUuid === 'function'
              ? Utilities.getUuid()
              : String(Date.now()));
        }
        if (!meta.submissionId) meta.submissionId = meta.submission_id;
        parsed.meta = meta;

        const fallbackInfo = {
          caseId,
          case_id,
          caseKey: caseInfo.caseKey,
          case_key: caseInfo.caseKey,
          userKey: caseInfo.userKey || caseInfo.user_key || prepared.userKey || '',
          user_key: caseInfo.userKey || caseInfo.user_key || prepared.userKey || '',
          lineId: caseInfo.lineId || prepared.lineId || '',
          line_id: caseInfo.lineId || prepared.lineId || '',
        };

        let resolvedCaseKey =
          caseInfo.caseKey && String(caseInfo.caseKey) ? String(caseInfo.caseKey) : '';
        if (!resolvedCaseKey && (caseInfo.userKey || prepared.userKey)) {
          const uk = caseInfo.userKey || prepared.userKey;
          resolvedCaseKey = `${uk}-${caseId}`;
        }
        if (!resolvedCaseKey) {
          resolvedCaseKey = drive_resolveCaseKeyFromMeta_(
            parsed?.meta || parsed?.META || {},
            fallbackInfo
          );
        }
        if (!resolvedCaseKey && fallbackInfo.userKey) {
          resolvedCaseKey = `${fallbackInfo.userKey}-${caseId}`;
        }

        const resolved_case_key = resolvedCaseKey;

        fallbackInfo.caseKey = resolved_case_key;
        fallbackInfo.case_key = resolved_case_key;
        if (resolved_case_key) caseInfo.caseKey = resolved_case_key;
        if (resolved_case_key) {
          if (!meta.case_key) meta.case_key = resolved_case_key;
          if (!meta.caseKey) meta.caseKey = resolved_case_key;
        }

        if (!resolved_case_key) {
          throw new Error('Unable to resolve case folder key');
        }

        const caseFolder = drive_getOrCreateCaseFolderByKey_(resolved_case_key);
        const caseFolderId = caseFolder.getId();
        caseInfo.folderId = caseFolderId;
        caseInfo.caseKey = resolved_case_key;
        caseInfo.case_key = resolved_case_key;
        const effectiveUserKey =
          caseInfo.userKey ||
          fallbackInfo.userKey ||
          (resolved_case_key && resolved_case_key.indexOf('-') >= 0
            ? resolved_case_key.split('-')[0]
            : '');
        if (effectiveUserKey) {
          caseInfo.userKey = effectiveUserKey;
          caseInfo.user_key = effectiveUserKey;
          fallbackInfo.userKey = effectiveUserKey;
          fallbackInfo.user_key = effectiveUserKey;
        }
        if (fallbackInfo.lineId && !caseInfo.lineId) caseInfo.lineId = fallbackInfo.lineId;

        if (parsed && parsed.meta) {
          if (!parsed.meta.case_id) parsed.meta.case_id = caseId;
          if (!parsed.meta.case_key) parsed.meta.case_key = resolved_case_key;
          if (!parsed.meta.caseKey) parsed.meta.caseKey = resolved_case_key;
          if (!parsed.meta.user_key && (caseInfo.userKey || caseInfo.user_key)) {
            parsed.meta.user_key = caseInfo.userKey || caseInfo.user_key;
          }
          if (!parsed.meta.userKey && (caseInfo.userKey || caseInfo.user_key)) {
            parsed.meta.userKey = caseInfo.userKey || caseInfo.user_key;
          }
        }
        try {
          Logger.log(
            '[Intake] allocated case_id=%s case_key=%s user_key=%s',
            case_id,
            resolved_case_key,
            caseInfo.userKey || caseInfo.user_key || ''
          );
        } catch (_) {}
        try {
          Logger.log(
            '[Intake] save target folder=%s (%s)',
            caseFolder.getName && caseFolder.getName(),
            caseFolderId
          );
        } catch (_) {}
        try {
          Logger.log('[Intake] json meta=%s', JSON.stringify(parsed.meta));
        } catch (_) {}

        const formKeyForStatus = String(parsed?.meta?.form_key || actualKey || '').trim();

        // intake はケース直行せず、必ず staging へ保存（後続の status_api で吸い上げ）
        if (formKeyForStatus === 'intake') {
          try {
            const P = PropertiesService.getScriptProperties();
            const EXPECT = String(P.getProperty('NOTIFY_SECRET') || '').trim();
            const ALLOW_NO_SECRET = (P.getProperty('ALLOW_NO_SECRET') || '').toLowerCase() === '1';
            const provided = String((meta && meta.secret) || '').trim();
            if (EXPECT && !ALLOW_NO_SECRET && provided !== EXPECT) {
              try { Logger.log('[Intake] secret mismatch: meta.secret=%s', provided || '(empty)'); } catch (_) {}
              formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, false);
              return;
            }
          } catch (_) {}

          if (!meta.submission_id) {
            meta.submission_id =
              meta.submissionId ||
              (typeof Utilities !== 'undefined' && typeof Utilities.getUuid === 'function'
                ? Utilities.getUuid()
                : String(Date.now()));
            if (!meta.submissionId) meta.submissionId = meta.submission_id;
            parsed.meta = meta;
          }

          const P = PropertiesService.getScriptProperties();
          const ROOT_ID = P.getProperty('DRIVE_ROOT_FOLDER_ID') || P.getProperty('ROOT_FOLDER_ID');
          const root = DriveApp.getFolderById(ROOT_ID);
          const it = root.getFoldersByName('_email_staging');
          const staging = it.hasNext() ? it.next() : root.createFolder('_email_staging');
          const fname = `intake__${meta.submission_id}.json`;
          const blob = Utilities.newBlob(JSON.stringify(parsed, null, 2), 'application/json', fname);
          staging.createFile(blob);
          try { Logger.log('[Intake] staged %s', fname); } catch (_) {}
          formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, true);
          return;
        }

        if (formIntake_isDuplicateSubmission_(caseFolderId, actualKey, meta.submission_id)) {
          formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, true);
          return;
        }

        const savedFile = saveSubmissionJson_(caseFolderId, parsed);
        const file_path = `${caseFolderId}/${savedFile.getName()}`;
        Logger.log('[Intake] saved %s', file_path);

        try {
          drive_placeFileIntoCase_(
            savedFile,
            parsed?.meta || parsed?.META || {},
            {
              caseId: caseId,
              case_id: case_id,
              caseKey: resolved_case_key,
              case_key: resolved_case_key,
              userKey: caseInfo.userKey || caseInfo.user_key,
              user_key: caseInfo.userKey || caseInfo.user_key,
              lineId: caseInfo.lineId,
              line_id: caseInfo.lineId,
            }
          );
        } catch (placeErr) {
          Logger.log('[Intake] placeFile error: %s', (placeErr && placeErr.stack) || placeErr);
        }

        if (typeof recordSubmission_ === 'function') {
          try {
            recordSubmission_({
              case_id: case_id,
              form_key: actualKey,
              submission_id: meta.submission_id || '',
              json_path: file_path,
              meta,
              case_key: resolved_case_key,
              user_key: caseInfo.userKey || caseInfo.user_key,
              line_id: caseInfo.lineId,
            });
          } catch (recErr) {
            Logger.log('[Intake] recordSubmission_ error: %s', (recErr && recErr.stack) || recErr);
          }
        }

        if (typeof updateCasesRow_ === 'function') {
          const basePatch = {};
          basePatch.last_activity = new Date();
          if (formKeyForStatus === 'intake') {
            basePatch.status = 'intake';
          } else if (def.statusAfterSave) {
            basePatch.status = def.statusAfterSave;
          }
          if (typeof def.afterSave === 'function') {
            try {
              const extra = def.afterSave(caseInfo, parsed) || {};
              Object.keys(extra || {}).forEach(function (k) {
                basePatch[k] = extra[k];
              });
            } catch (e) {
              Logger.log('[Intake] afterSave error: %s', (e && e.stack) || e);
            }
          }
          basePatch.case_key = resolved_case_key;
          basePatch.folder_id = caseFolderId;
          if (caseInfo.userKey || caseInfo.user_key) {
            basePatch.user_key = caseInfo.userKey || caseInfo.user_key;
          }
          updateCasesRow_(case_id, basePatch);
        }

        formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, true);
      } catch (err) {
        GmailApp.createDraft(
          msg.getFrom(),
          `[BAS Intake Error] ${def.name}`,
          String(err),
          { htmlBody: `<pre>${safeHtml((err && err.stack) || err)}</pre>` }
        );
        formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, false);
      }
    });
  });
}

function formIntake_assignQueueLabels_() {
  const query = `label:${FORM_INTAKE_LABEL_TO_PROCESS} ("form_key:" OR "==== META START ====")`;
  const threads = GmailApp.search(query, 0, 100);
  threads.forEach(function (thread) {
    try {
      const messages = thread.getMessages();
      if (!messages || !messages.length) return;
      const msg = messages[0];
      const body = msg.getPlainBody() || HtmlService.createHtmlOutput(msg.getBody()).getContent();
      const meta = parseMetaBlock_(body);
      const formKey = String(meta?.form_key || '').trim();
      const def = FORM_INTAKE_REGISTRY[formKey];
      if (def) {
        thread.addLabel(formIntake_labelOrCreate_(def.queueLabel));
      }
    } catch (_) {}
  });
}

function formIntake_resolveCase_(caseId, def) {
  const resolver = def.caseResolver || (typeof resolveCaseByCaseId_ === 'function' ? resolveCaseByCaseId_ : null);
  if (!resolver) throw new Error('resolveCaseByCaseId_ is not defined');
  const info = resolver(caseId);
  if (!info) throw new Error('Unknown case_id: ' + caseId);
  return info;
}

function formIntake_ensureCaseFolder_(caseInfo, def) {
  const ensureFn =
    def.ensureCaseFolder || (typeof ensureCaseFolderId_ === 'function' ? ensureCaseFolderId_ : null);
  if (!ensureFn) throw new Error('ensureCaseFolderId_ is not defined');
  return ensureFn(caseInfo);
}

function formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, processed) {
  if (processed) {
    thread.addLabel(processedLabel);
    try {
      thread.removeLabel(queueLabel);
    } catch (_) {}
    try {
      thread.removeLabel(toProcessLabel);
    } catch (_) {}
  }
  try {
    thread.removeLabel(lockLabel);
  } catch (_) {}
}

function formIntake_isDuplicateSubmission_(folderId, formKey, submissionId) {
  if (!folderId || !formKey || !submissionId) return false;
  try {
    const folder = DriveApp.getFolderById(folderId);
    const fname = `${formKey}__${submissionId}.json`;
    return folder.getFilesByName(fname).hasNext();
  } catch (_) {
    return false;
  }
}

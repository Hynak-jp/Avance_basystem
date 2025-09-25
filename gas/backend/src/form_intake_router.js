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
        if (draft && draft.draftUrl) patch.lastDraftUrl = draft.draftUrl;
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

        const caseInfo = formIntake_resolveCase_(meta.case_id, def);
        caseInfo.folderId = formIntake_ensureCaseFolder_(caseInfo, def);

        if (formIntake_isDuplicateSubmission_(caseInfo.folderId, actualKey, meta.submission_id)) {
          formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, true);
          return;
        }

        const filePath = saveSubmissionJson_(caseInfo.folderId, parsed);
        Logger.log('[Intake] saved %s', filePath);

        if (typeof recordSubmission_ === 'function') {
          try {
            recordSubmission_({
              caseId: caseInfo.caseId || meta.case_id,
              form_key: actualKey,
              submission_id: meta.submission_id || '',
              json_path: filePath,
            });
          } catch (recErr) {
            Logger.log('[Intake] recordSubmission_ error: %s', (recErr && recErr.stack) || recErr);
          }
        }

        if (typeof updateCasesRow_ === 'function') {
          const basePatch = {};
          basePatch.lastActivity = new Date();
          if (def.statusAfterSave) basePatch.status = def.statusAfterSave;
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
          updateCasesRow_(meta.case_id, basePatch);
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

/** ================== forms_ingest_core.js ==================
 * 汎用フォーム保存パイプライン
 * - メール本文の解析（parseFormMail_ に委譲）
 * - submission_id 正規化（normalizeSubmissionIdStrict_ があれば優先）
 * - case_id → case_key フォルダ解決（resolveCaseByCaseId_ / ensureCaseFolderId_）
 * - <form_key>__<submission_id>.json をケース直下に保存（saveSubmissionJson_）
 * 依存関数（既存プロジェクト内にある前提）:
 *   parseFormMail_, resolveCaseByCaseId_, ensureCaseFolderId_, saveSubmissionJson_, updateCasesRow_
 * --------------------------------------------------------- */

/** mapper レジストリ（多重定義でも上書きしないガード） */
var FORM_MAPPERS =
  this.FORM_MAPPERS && typeof this.FORM_MAPPERS === 'object'
    ? this.FORM_MAPPERS
    : Object.create(null);
this.FORM_MAPPERS = FORM_MAPPERS;

var AUTO_DRAFT_HANDLERS =
  typeof globalThis !== 'undefined'
    ? (globalThis.AUTO_DRAFT_HANDLERS =
        globalThis.AUTO_DRAFT_HANDLERS || Object.create(null))
    : (this.AUTO_DRAFT_HANDLERS = this.AUTO_DRAFT_HANDLERS || Object.create(null));

var FORM_INGEST_SECRET_CACHE = { value: null, bust: null };
var ENFORCE_SID_14_CACHE = { value: null, bust: null };

function ensureEmailStagingFolder_(parsed) {
  var root = DriveApp.getRootFolder();
  function getOrCreate_(folder, name) {
    var it = folder.getFoldersByName(name);
    return it.hasNext() ? it.next() : folder.createFolder(name);
  }
  var base = getOrCreate_(root, '_email_staging');
  var receivedAt = null;
  try {
    receivedAt = parsed && parsed.meta && parsed.meta.received_at ? new Date(parsed.meta.received_at) : null;
  } catch (_) {
    receivedAt = null;
  }
  var tz = (parsed && parsed.meta && parsed.meta.received_tz) || 'Asia/Tokyo';
  var stamp = receivedAt && isFinite(receivedAt.getTime()) ? receivedAt : new Date();
  var monthName = Utilities.formatDate(stamp, tz, 'yyyy-MM');
  var monthFolder = getOrCreate_(base, monthName);
  var sid = (parsed && parsed.meta && parsed.meta.submission_id) || 'unknown';
  var leaf = getOrCreate_(monthFolder, 'submission_' + sid);
  return leaf.getId();
}

/**
 * intake 専用: ケース採番を行わず、メールをそのまま staging へ保存
 *  - _extractEmail_ / _fillMetaBeforeStage_ / fi_* 系ヘルパは form_intake_router.js 由来を再利用
 */
function stageIntakeMail_(thread, msg, parsed, meta, rawBody) {
  parsed = parsed || {};
  meta = meta || parsed.meta || {};
  parsed.meta = parsed.meta || meta || {};

  var props = null;
  try {
    if (typeof PropertiesService !== 'undefined' && PropertiesService.getScriptProperties) {
      props = PropertiesService.getScriptProperties();
    }
  } catch (_) {}

  try {
    var expect = (props && props.getProperty && props.getProperty('NOTIFY_SECRET')) || '';
    var allowNoSecret =
      String((props && props.getProperty && props.getProperty('ALLOW_NO_SECRET')) || '')
        .toLowerCase()
        .trim() === '1';
    var provided = String((meta && meta.secret) || '').trim();
    if (expect && !allowNoSecret && provided !== expect) {
      try {
        Logger.log('[Intake] secret mismatch (ingest_core): meta.secret=%s', provided || '(empty)');
      } catch (_) {}
      return { staged: false, rejected: true, reason: 'secret_mismatch', form_key: 'intake' };
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

  var rootId = '';
  try {
    rootId =
      (props && props.getProperty && props.getProperty('DRIVE_ROOT_FOLDER_ID')) ||
      (props && props.getProperty && props.getProperty('ROOT_FOLDER_ID')) ||
      '';
  } catch (_) {}
  var root = null;
  try {
    root = rootId ? DriveApp.getFolderById(rootId) : DriveApp.getRootFolder();
  } catch (_) {
    root = DriveApp.getRootFolder();
  }
  var staging = root;
  try {
    var itSt = root.getFoldersByName('_email_staging');
    staging = itSt.hasNext() ? itSt.next() : root.createFolder('_email_staging');
  } catch (_) {}

  var knownLineId = '';
  try {
    knownLineId =
      typeof getLineIdFromContext_ === 'function' ? getLineIdFromContext_(null, msg, parsed) || '' : '';
  } catch (_) {}
  var knownCaseId = '';
  try {
    knownCaseId = String(meta.case_id || meta.caseId || '').replace(/\D/g, '');
  } catch (_) {}

  var emailEarly = '';
  try {
    if (typeof _extractEmail_ === 'function') {
      emailEarly = _extractEmail_({ msg: msg, obj: parsed, rawText: rawBody });
    }
  } catch (_) {}
  try {
    if (typeof _fillMetaBeforeStage_ === 'function') {
      parsed = _fillMetaBeforeStage_(parsed, {
        knownLineId: knownLineId,
        knownCaseId: knownCaseId,
        email: emailEarly,
      });
      meta = parsed && parsed.meta ? parsed.meta : meta;
    }
  } catch (_) {}
  if (!(emailEarly || (parsed && parsed.meta && parsed.meta.line_id))) {
    try {
      var itQe = root.getFoldersByName('_quarantine');
      var qfe = itQe.hasNext() ? itQe.next() : root.createFolder('_quarantine');
      var qne = 'raw_intake__' + (meta.submission_id || Date.now()) + '.json';
      qfe.createFile(
        Utilities.newBlob(JSON.stringify({ raw: parsed, note: 'no_email_no_lineid' }), 'application/json', qne)
      );
      Logger.log('[Intake][drop] no identifiers; quarantined %s', qne);
    } catch (_) {}
    return {
      staged: true,
      quarantined: true,
      form_key: 'intake',
      submission_id: meta.submission_id || '',
      case_id: (parsed && parsed.meta && parsed.meta.case_id) || '',
      caseKey: (parsed && parsed.meta && (parsed.meta.case_key || parsed.meta.caseKey)) || '',
      name: '',
    };
  }

  try {
    var normCaseIdFn =
      typeof normCaseId_ === 'function'
        ? normCaseId_
        : function (s) {
            var d = String(s || '').replace(/\D/g, '');
            return d ? ('0000' + d).slice(-4) : '';
          };
    var normCaseKeyFn =
      typeof normCaseKey_ === 'function'
        ? normCaseKey_
        : function (s) {
            return String(s || '').trim().toLowerCase();
          };
    var buildCandsFn =
      typeof buildCandidates_ === 'function'
        ? buildCandidates_
        : function () {
            return [];
          };
    var ukEarly = '';
    try {
      ukEarly =
        typeof fi_userKeyFromLineId_ === 'function'
          ? fi_userKeyFromLineId_(String((parsed && parsed.meta && parsed.meta.line_id) || knownLineId || '')) || ''
          : '';
    } catch (_) {}
    var knownEarly = {
      case_key:
        ukEarly && knownCaseId ? String(ukEarly).toLowerCase() + '-' + normCaseIdFn(knownCaseId) : '',
      case_id: normCaseIdFn(knownCaseId || ''),
      line_id: String((parsed && parsed.meta && parsed.meta.line_id) || knownLineId || ''),
    };
    var fromCasesEarly =
      typeof fi_casesLookup_ === 'function'
        ? fi_casesLookup_({ lineId: knownEarly.line_id, caseId: knownEarly.case_id })
        : null;
    var fromContactsEarly =
      typeof fi_contactsLookupByEmail_ === 'function' ? fi_contactsLookupByEmail_(emailEarly) : null;
    var candEarly = {
      fromCases: fromCasesEarly,
      fromContacts: fromContactsEarly,
      fromLine: { line_id: knownEarly.line_id, user_key: ukEarly },
      metaInMail: parsed && parsed.meta,
    };
    var rEarly =
      typeof resolveMetaWithPriority_ === 'function'
        ? resolveMetaWithPriority_(candEarly, knownEarly)
        : { meta: {} };
    var mEarly = (rEarly && rEarly.meta) || {};
    parsed.meta = parsed.meta || {};
    parsed.meta.user_key = parsed.meta.user_key || mEarly.user_key || ukEarly || parsed.meta.userKey || '';
    parsed.meta.case_id = normCaseIdFn(
      parsed.meta.case_id || parsed.meta.caseId || mEarly.case_id || knownEarly.case_id || ''
    );
    parsed.meta.case_key = normCaseKeyFn(
      parsed.meta.case_key ||
        parsed.meta.caseKey ||
        mEarly.case_key ||
        (parsed.meta.user_key && parsed.meta.case_id
          ? String(parsed.meta.user_key).toLowerCase() + '-' + parsed.meta.case_id
          : '')
    );
    parsed.meta.line_id = parsed.meta.line_id || mEarly.line_id || knownEarly.line_id || '';
    try {
      Logger.log(
        '[Intake] adopt(by=%s,source=%s) candidates=%s',
        (rEarly && rEarly.by) || '',
        (rEarly && rEarly.source) || 'unknown',
        JSON.stringify(buildCandsFn(candEarly))
      );
    } catch (_) {}
  } catch (_) {}

  var fname = 'intake__' + meta.submission_id + '.json';
  var blob = Utilities.newBlob(JSON.stringify(parsed, null, 2), 'application/json', fname);
  var createdFile = staging.createFile(blob);
  try {
    if (typeof tryMoveIntakeToCase_ === 'function') {
      tryMoveIntakeToCase_(parsed && parsed.meta, createdFile, fname);
    }
  } catch (_) {}
  try {
    var candidatesEarlyJson = (function () {
      try {
        return JSON.stringify(
          typeof buildCandidates_ === 'function'
            ? buildCandidates_(typeof candEarly !== 'undefined' ? candEarly : { metaInMail: parsed && parsed.meta })
            : []
        );
      } catch (_) {
        return '[]';
      }
    })();
    Logger.log(
      '[Intake] staged name=%s meta={lid:%s, uk:%s, cid:%s, ckey:%s} candidates=%s',
      fname,
      (parsed && parsed.meta && parsed.meta.line_id) || '',
      (parsed && parsed.meta && parsed.meta.user_key) || '',
      (parsed && parsed.meta && parsed.meta.case_id) || '',
      (parsed && parsed.meta && parsed.meta.case_key) || '',
      candidatesEarlyJson
    );
  } catch (_) {}

  return {
    staged: true,
    form_key: 'intake',
    submission_id: meta.submission_id || '',
    case_id: (parsed && parsed.meta && parsed.meta.case_id) || '',
    caseKey: (parsed && parsed.meta && (parsed.meta.case_key || parsed.meta.caseKey)) || '',
    fileId: createdFile && createdFile.getId && createdFile.getId(),
    name: fname,
  };
}

/** 各フォームの「FIELDS → model」変換関数を登録 */
function registerFormMapper(formKey, mapperFn) {
  if (!formKey || typeof mapperFn !== 'function') return;
  FORM_MAPPERS[String(formKey).trim()] = mapperFn;
}

function registerAutoDraft(formKey, fn) {
  try {
    if (!formKey || typeof fn !== 'function') return;
    AUTO_DRAFT_HANDLERS[formKey] = fn;
    try {
      Logger.log('[INGEST] auto-draft handler registered: %s', formKey);
    } catch (_) {}
  } catch (_) {}
}

function ensureFormMapperRegistered_(formKey) {
  if (!formKey) return false;
  if (FORM_MAPPERS[formKey]) return true;
  var factories = typeof globalThis !== 'undefined' ? globalThis.FORM_MAPPER_FACTORIES : null;
  if (factories && typeof factories === 'object') {
    var factory = factories[formKey];
    if (typeof factory === 'function') {
      registerFormMapper(formKey, factory);
      try {
        Logger.log('[INGEST] mapper (re)registered: %s', formKey);
      } catch (_) {}
      return true;
    }
  }
  if (formKey === 's2005_creditors' && typeof mapS2005FieldsToModel_ === 'function') {
    registerFormMapper('s2005_creditors', mapS2005FieldsToModel_);
    return true;
  }
  if (formKey === 's2006_creditors_public' && typeof mapS2006FieldsToModel_ === 'function') {
    registerFormMapper('s2006_creditors_public', mapS2006FieldsToModel_);
    return true;
  }
  try {
    Logger.log('[INGEST] warning: mapper not found for form_key=%s (using generic)', formKey);
  } catch (_) {}
  return false;
}
function getIngestSecret_() {
  var props = null;
  try {
    if (
      typeof PropertiesService !== 'undefined' &&
      PropertiesService.getScriptProperties &&
      PropertiesService.getScriptProperties().getProperty
    ) {
      props = PropertiesService.getScriptProperties();
    }
  } catch (_) {}
  var bust =
    props && typeof props.getProperty === 'function'
      ? props.getProperty('INGEST_CONFIG_BUST') || ''
      : '';
  if (FORM_INGEST_SECRET_CACHE.value !== null && FORM_INGEST_SECRET_CACHE.bust === bust) {
    return FORM_INGEST_SECRET_CACHE.value;
  }
  var secret = '';
  if (props && typeof props.getProperty === 'function') {
    try {
      secret = props.getProperty('NOTIFY_SECRET') || '';
    } catch (_) {
      secret = '';
    }
  }
  if (!secret) {
    secret = 'FM-BAS';
  }
  FORM_INGEST_SECRET_CACHE = { value: secret, bust: bust };
  return secret;
}

function getEnforceSid14_() {
  var props = null;
  try {
    if (
      typeof PropertiesService !== 'undefined' &&
      PropertiesService.getScriptProperties &&
      PropertiesService.getScriptProperties().getProperty
    ) {
      props = PropertiesService.getScriptProperties();
    }
  } catch (_) {}
  var bust =
    props && typeof props.getProperty === 'function'
      ? props.getProperty('INGEST_CONFIG_BUST') || ''
      : '';
  if (ENFORCE_SID_14_CACHE.value !== null && ENFORCE_SID_14_CACHE.bust === bust) {
    return ENFORCE_SID_14_CACHE.value;
  }
  var enforce = false;
  if (props && typeof props.getProperty === 'function') {
    try {
      enforce = props.getProperty('ENFORCE_SID_14') === '1';
    } catch (_) {
      enforce = false;
    }
  }
  ENFORCE_SID_14_CACHE = { value: enforce, bust: bust };
  return enforce;
}

function normalizeMeta_(meta) {
  var m = meta || {};
  var ztrim = function (s) {
    return String(s == null ? '' : s).replace(/^[\s\u3000]+|[\s\u3000]+$/g, '');
  };
  if (!m.form_key && m.formKey) m.form_key = m.formKey;
  if (!m.submission_id && m.submissionId) m.submission_id = m.submissionId;
  if (!m.case_id && (m.caseId || m.caseID)) {
    m.case_id = m.caseId || m.caseID;
  }
  m.form_key = ztrim(m.form_key);
  m.submission_id = ztrim(m.submission_id).replace(/[^\d]/g, '');
  m.case_id = ztrim(m.case_id).replace(/[^\d]/g, '');
  try {
    if (m.submission_id && typeof normalizeSubmissionIdStrict_ === 'function') {
      var normSid = normalizeSubmissionIdStrict_(m.submission_id, m.tsKey, m);
      if (normSid) m.submission_id = String(normSid);
    }
  } catch (_) {}
  try {
    if (m.case_id) {
      if (typeof normalizeCaseIdString_ === 'function') {
        m.case_id = normalizeCaseIdString_(m.case_id);
      } else {
        m.case_id = ('0000' + m.case_id).slice(-4);
      }
    }
  } catch (_) {
    if (m.case_id) {
      m.case_id = ('0000' + m.case_id).slice(-4);
    }
  }
  return m;
}

function normalizeSecret_(s) {
  if (s == null) return '';
  var out = String(s);
  try {
    out = out.replace(/\u00A0/g, ' '); // nbsp
    out = out.replace(/^\uFEFF/, ''); // BOM
    out = out.replace(/[\u200B\u200C\u200D]/g, ''); // zero width variants
    out = out.normalize('NFKC'); // 全角→半角 等
  } catch (_) {}
  out = out.replace(/[‐‑–—−─―]/g, '-'); // 各種ハイフン
  out = out.replace(/\s+/g, '');
  return out.toLowerCase();
}

/**
 * FormAttach/NoMeta ラベルのスレッドを 24h 以上経過したら FormAttach/NoMeta/Archive へ退避
 */
function run_AgeOutNoMetaThreads() {
  if (typeof GmailApp === 'undefined' || !GmailApp.getUserLabelByName) return;
  var labNoMeta = GmailApp.getUserLabelByName('FormAttach/NoMeta');
  if (!labNoMeta) return;
  var labArchive =
    GmailApp.getUserLabelByName('FormAttach/NoMeta/Archive') ||
    GmailApp.createLabel('FormAttach/NoMeta/Archive');
  var cutoff = Date.now() - 24 * 60 * 60 * 1000;
  var threads = labNoMeta.getThreads(0, 200);
  for (var i = 0; i < threads.length; i++) {
    var th = threads[i];
    try {
      var lastDate = th.getLastMessageDate();
      if (lastDate && lastDate.getTime() < cutoff) {
        th.addLabel(labArchive);
        th.removeLabel(labNoMeta);
        try {
          Logger.log('[Router] archive no-meta tid=%s', th.getId());
        } catch (_) {}
      }
    } catch (_) {}
  }
}

/** 既存の normalizeSubmissionIdStrict_ があれば利用、なければフォールバック */
function ensureSubmissionIdDigits_(sid, tsKey, meta) {
  function enforceLengthIfNeeded_(digits) {
    if (!digits) return digits;
    if (getEnforceSid14_() && /^\d{3,13}$/.test(digits)) {
      digits = (digits + '00000000000000').slice(0, 14);
    }
    return digits;
  }
  try {
    if (typeof normalizeSubmissionIdStrict_ === 'function') {
      var normalized = normalizeSubmissionIdStrict_(sid, tsKey, meta);
      if (normalized) return enforceLengthIfNeeded_(normalized);
    }
  } catch (_) {}
  var originalDigits = String(sid || '').replace(/\D/g, '');
  var digits = originalDigits;
  if (digits.length > 32) {
    var truncated = digits.slice(-14);
    try {
      Logger.log('[INGEST] submission_id truncated sid=%s -> %s', digits, truncated);
    } catch (_) {}
    digits = truncated;
  }
  digits = enforceLengthIfNeeded_(digits);
  if (digits.length >= 3) return digits;
  var fromMeta = '';
  if (meta && typeof meta === 'object') {
    fromMeta = String(meta.submitted_at || meta.submittedAt || '').trim();
  }
  var fromMetaDigits = fromMeta.replace(/[^\d]/g, '').slice(0, 14);
  if (fromMetaDigits.length === 12) {
    fromMetaDigits = fromMetaDigits + '00';
  } else if (fromMetaDigits.length > 0 && fromMetaDigits.length < 14) {
    fromMetaDigits = (fromMetaDigits + '00000000000000').slice(0, 14);
  }
  fromMetaDigits = enforceLengthIfNeeded_(fromMetaDigits);
  if (fromMetaDigits.length === 14) return fromMetaDigits;
  if (fromMetaDigits.length >= 8) return (fromMetaDigits + '00000000000000').slice(0, 14);
  var tz =
    (typeof Session !== 'undefined' &&
      Session.getScriptTimeZone &&
      Session.getScriptTimeZone()) ||
    'Asia/Tokyo';
  var fallback = Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmmss');
  return enforceLengthIfNeeded_(fallback);
}

/** subject から form_key を推測（METAが無い時の保険） */
function guessFormKeyFromSubject_(subject) {
  var raw = String(subject || '');
  var s = raw.toLowerCase();
  var match = s.match(/s\d{4}/);
  if (match) {
    var prefix = match[0];
    var keys = Object.keys(FORM_MAPPERS || {});
    for (var i = 0; i < keys.length; i++) {
      var key = String(keys[i]).toLowerCase();
      if (key.indexOf(prefix) === 0) return keys[i];
    }
  }
  // 日本語件名のみでも適切に判定
  if (/s2006/.test(s) || /公租公課/.test(raw) || /債権者一覧表（公租公課用）/.test(raw)) {
    return 's2006_creditors_public';
  }
  if (/s2005/.test(s) || (/債権者一覧表/.test(raw) && !/公租公課/.test(raw))) {
    return 's2005_creditors';
  }
  if (/s2002/.test(s)) return 's2002_userform';
  // TODO: 新フォームはここに追記
  return 'unknown_form';
}

/** フィールド配列→素直なオブジェクト（汎用フォールバック） */
function normalizeLabelForKey_(s) {
  return String(s || '')
    .replace(/^[\s\u3000]+|[\s\u3000]+$/g, '')
    .replace(/[（）]/g, function (m) {
      return m === '（' ? '(' : ')';
    })
    .replace(/\s+/g, '')
    .toLowerCase();
}

function mapFieldsGeneric_(fields) {
  var obj = {};
  var seen = Object.create(null);
  (fields || []).forEach(function (row) {
    var base = row && row.label ? String(row.label).trim() : 'field';
    var norm = normalizeLabelForKey_(base) || 'field';
    var count = seen[norm] || 0;
    var key = count ? base + '_' + count : base;
    seen[norm] = count + 1;
    obj[key] = row && row.value != null ? row.value : '';
  });
  return { fields: obj, fields_indexed: fields || [] };
}

/**
 * 【中核】メール（件名・本文）を解析→cases台帳で case_key フォルダを解決→
 * ケース直下へ <form_key>__<submission_id>.json 保存。
 * opts: { form_key?: string, mapper?: function, case_id?: string }
 */
function ingestFormMailToCase_(subject, body, opts) {
  opts = opts || {};
  var lock = null;
  var lockAcquired = false;
  try {
    if (typeof LockService !== 'undefined' && LockService.getScriptLock) {
      lock = LockService.getScriptLock();
      if (lock && lock.tryLock(5000)) {
        lockAcquired = true;
      } else if (lock) {
        throw new Error('ingestFormMailToCase_: lock timeout');
      }
    }
  } catch (lockErr) {
    throw new Error('ingestFormMailToCase_: failed to acquire lock :: ' + (lockErr && lockErr.message ? lockErr.message : lockErr));
  }

  try {
    var parsed = parseFormMail_(subject, body);
    parsed.meta = normalizeMeta_(parsed.meta || {});

    var formKey = String(
      opts.form_key || parsed.meta.form_key || guessFormKeyFromSubject_(subject)
    ).trim();
    var expectedSecretNorm = normalizeSecret_(getIngestSecret_());
    var gotSecretNorm = normalizeSecret_((parsed.meta && parsed.meta.secret) || '');
    var allowNoSecret = false;
    try {
      allowNoSecret =
        formKey === 'intake' &&
        typeof PropertiesService !== 'undefined' &&
        PropertiesService.getScriptProperties &&
        String(PropertiesService.getScriptProperties().getProperty('ALLOW_NO_SECRET') || '')
          .toLowerCase()
          .trim() === '1';
    } catch (_) {}
    var secretOk = gotSecretNorm === expectedSecretNorm || (allowNoSecret && !gotSecretNorm);
    if (!expectedSecretNorm || !secretOk) {
      throw new Error(
        'Invalid secret (subject: ' +
          subject +
          ', form_key: ' +
          formKey +
          ', case_id: ' +
          (parsed.meta.case_id || '') +
          ')'
      );
    }

    parsed.meta.form_key = formKey;
    ensureFormMapperRegistered_(parsed.meta.form_key);
    if (parsed.meta.form_key === 'unknown_form') {
      try {
        Logger.log(
          '[INGEST] warning: unknown form_key inferred from subject=%s (fallback generic)',
          subject
        );
      } catch (_) {}
    }

    parsed.meta.submission_id = ensureSubmissionIdDigits_(
      parsed.meta.submission_id,
      parsed.meta.tsKey,
      parsed.meta
    );
    if (opts && opts.message_id && !parsed.meta.message_id) {
      parsed.meta.message_id = opts.message_id;
    }

    if (opts && Object.prototype.hasOwnProperty.call(opts, 'case_id')) {
      var override = normalizeMeta_({ case_id: opts.case_id }).case_id;
      if (override) {
        parsed.meta.case_id = override;
      }
    }

    var caseId = String(parsed.meta.case_id || '').trim();
    parsed.meta.case_id = caseId;
    if (formKey === 'intake') {
      return stageIntakeMail_(
        opts && opts.thread,
        (opts && opts.message) || (opts && opts.msg),
        parsed,
        parsed.meta,
        body
      );
    }
    if (!caseId)
      throw new Error(
        'META.case_id is required (subject: ' +
          subject +
          ', form_key: ' +
          parsed.meta.form_key +
          ', submission_id: ' +
          (parsed.meta.submission_id || '') +
          ')'
      );

    var mapper = opts.mapper || FORM_MAPPERS[parsed.meta.form_key];
    var effectiveMapper = mapper;
    if (!mapper) {
      try {
        Logger.log(
          '[INGEST] warning: no mapper for form_key=%s (subject=%s) fallback=generic',
          parsed.meta.form_key,
          subject
        );
      } catch (_) {}
      mapper = mapFieldsGeneric_;
    }
    try {
      Logger.log(
        '[INGEST] mapping start form=%s case=%s sid=%s',
        parsed.meta.form_key,
        caseId,
        parsed.meta.submission_id || ''
      );
    } catch (_) {}
    try {
      parsed.model = mapper(parsed.fieldsRaw || [], parsed.meta);
    } catch (mapperErr) {
      try {
        Logger.log(
          '[INGEST] mapper error form=%s case=%s err=%s (fallback generic)',
          parsed.meta.form_key,
          caseId,
          (mapperErr && mapperErr.stack) || mapperErr
        );
      } catch (_) {}
      parsed.model = mapFieldsGeneric_(parsed.fieldsRaw || [], parsed.meta);
      effectiveMapper = mapFieldsGeneric_;
    }
    try {
      Logger.log(
        '[INGEST] mapping done  form=%s sid=%s',
        parsed.meta.form_key,
        parsed.meta.submission_id || ''
      );
    } catch (_) {}
    parsed.meta.mapper_used = effectiveMapper === mapFieldsGeneric_ ? 'generic' : parsed.meta.form_key;

    var caseInfo = null;
    try {
      caseInfo = resolveCaseByCaseId_(caseId);
    } catch (_) {
      caseInfo = null;
    }
    if (caseInfo) {
      parsed.meta.case_key =
        (caseInfo && caseInfo.caseKey) ||
        parsed.meta.case_key ||
        parsed.meta.caseKey ||
        '';
    } else {
      parsed.meta.case_key = parsed.meta.case_key || parsed.meta.caseKey || '';
    }
    if (!parsed.meta.received_at) {
      parsed.meta.received_at = new Date().toISOString();
    }
    if (!parsed.meta.received_tz) {
      parsed.meta.received_tz =
        (typeof Session !== 'undefined' &&
          Session.getScriptTimeZone &&
          Session.getScriptTimeZone()) ||
        'Asia/Tokyo';
    }
    if (!parsed.meta.ingest_subject) {
      parsed.meta.ingest_subject = String(subject || '');
    }
    var folderId = '';
    var staged = false;
    if (caseInfo) {
      try {
        folderId = ensureCaseFolderId_(caseInfo);
      } catch (ensureErr) {
        folderId = '';
        try {
          Logger.log('[INGEST] warning: ensureCaseFolderId_ failed case=%s err=%s', caseId, (ensureErr && ensureErr.message) || ensureErr);
        } catch (_) {}
      }
    }
    if (!folderId) {
      staged = true;
      folderId = ensureEmailStagingFolder_(parsed);
    }
    var file;
    try {
      file = saveSubmissionJson_(folderId, parsed);
    } catch (saveErr) {
      var stagedFolderId = ensureEmailStagingFolder_(parsed);
      var fallbackName = (parsed.meta.form_key || 'unknown_form') + '__' + (parsed.meta.submission_id || 'unknown') + '.json';
      var fallbackBlob = Utilities.newBlob(JSON.stringify(parsed, null, 2), 'application/json', fallbackName);
      file = DriveApp.getFolderById(stagedFolderId).createFile(fallbackBlob);
      try {
        Logger.log(
          '[INGEST] warn: save to case folder failed case=%s sid=%s err=%s (staged)',
          caseId,
          parsed.meta.submission_id || '',
          (saveErr && saveErr.stack) || saveErr
        );
      } catch (_) {}
      staged = true;
    }

    if (!staged && caseInfo) {
      try {
        updateCasesRow_(caseId, {
          last_activity: new Date(),
          last_form_key: parsed.meta.form_key,
        });
      } catch (_) {}
    }

    try {
      var fieldsCount = (parsed && parsed.fieldsRaw && parsed.fieldsRaw.length) || 0;
      if (staged) {
        Logger.log(
          '[INGEST] staged: case=%s form=%s sid=%s fields=%s file=%s mid=%s',
          caseId,
          parsed.meta.form_key,
          parsed.meta.submission_id || '',
          fieldsCount,
          file.getName(),
          parsed.meta.message_id || ''
        );
      } else {
      Logger.log(
        '[INGEST] saved: case=%s form=%s sid=%s fields=%s file=%s caseKey=%s mid=%s',
        caseId,
        parsed.meta.form_key,
        parsed.meta.submission_id || '',
        fieldsCount,
        file.getName(),
        (caseInfo && caseInfo.caseKey) || '(unknown)',
        parsed.meta.message_id || ''
      );
      }
    } catch (_) {}

    return {
      fileId: file.getId(),
      name: file.getName(),
      caseKey: staged ? '' : (caseInfo && caseInfo.caseKey) || '',
      form_key: parsed.meta.form_key,
      case_id: staged ? caseId : (caseInfo && caseInfo.caseId) || caseId,
      submission_id: parsed.meta.submission_id || '',
      staged: staged
    };
  } finally {
    if (lock && lockAcquired) {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  }
}

/**
 * フォルダの「それらしさ」を採点。
 * s2002_userform__*.json: +5（≥1件で加点）
 * s2006_creditors_public__*.json: +3（≥1件で加点）
 * 任意の直下 .json: +2（≥1件で加点）
 * drafts/ サブフォルダ: +3
 * attachments/ or staff_inputs/ サブフォルダ: +1 ずつ
 */
function scoreCaseFolder_(folder) {
  var s2 = 0;
  var s6 = 0;
  var s5 = 0;
  var anyJson = 0;
  if (!folder)
    return { score: 0, s2: 0, s6: 0, s5: 0, anyJson: 0, drafts: 0, attach: 0, staff: 0 };
  var itF = folder.getFiles();
  var scanned = 0;
  while (itF.hasNext() && scanned < 500) {
    var f = itF.next();
    scanned++;
    var name = f.getName && f.getName();
    if (!name) continue;
    if (/\.json$/i.test(name)) anyJson = 1;
    if (/^s2002_userform__/i.test(String(name).trim())) s2 = 1;
    if (/^s2006_creditors_public__/i.test(String(name).trim())) s6 = 1;
    if (/^s2005_creditors__/i.test(String(name).trim())) s5 = 1;
  }
  var drafts = 0;
  var attach = 0;
  var staff = 0;
  var itD = folder.getFolders();
  while (itD.hasNext()) {
    var d = itD.next();
    var dn = d.getName && d.getName();
    if (!dn) continue;
    dn = String(dn).trim().toLowerCase();
    if (dn === 'drafts') drafts = 1;
    if (dn === 'attachments') attach = 1;
    if (dn === 'staff_inputs') staff = 1;
  }
  var score = s2 * 5 + s6 * 3 + s5 * 3 + anyJson * 2 + drafts * 3 + attach * 1 + staff * 1;
  return {
    score: score,
    s2: s2,
    s6: s6,
    s5: s5,
    anyJson: anyJson,
    drafts: drafts,
    attach: attach,
    staff: staff,
  };
}

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

// ===== 共通マッチャ（必要なら定義）: case_key → case_id → line_id =====
if (typeof normCaseId_ !== 'function') {
  function normCaseId_(s) {
    s = String(s || '').trim();
    var n = s.replace(/^0+/, '');
    if (!n) return '';
    var num = parseInt(n, 10);
    if (!isFinite(num)) return '';
    return Utilities.formatString('%04d', num);
  }
}
if (typeof normCaseKey_ !== 'function') {
  function normCaseKey_(s) {
    s = String(s || '').trim().toLowerCase();
    var m = s.match(/^([a-z0-9]{2,})-(\d{1,})$/);
    if (!m) return s;
    return m[1] + '-' + normCaseId_(m[2]);
  }
}
if (typeof normLineId_ !== 'function') {
  function normLineId_(s) { return String(s || '').trim(); }
}
if (typeof matchMetaToCase_ !== 'function') {
  function matchMetaToCase_(fileMeta, known) {
    var fm = fileMeta || {};
    var fk = normCaseKey_(fm.case_key || fm.caseKey || '');
    var fid = normCaseId_(fm.case_id || fm.caseId || '');
    var fl = normLineId_(fm.line_id || fm.lineId || '');
    var kk = normCaseKey_(known && known.case_key || '');
    var kid = normCaseId_(known && known.case_id || '');
    var kl = normLineId_(known && known.line_id || '');
    if (fk && kk && fk === kk) return { ok: true, by: 'case_key' };
    if (fid && kid && fid === kid) return { ok: true, by: 'case_id' };
    if (fl && kl && fl === kl) return { ok: true, by: 'line_id' };
    return { ok: false, by: '' };
  }
}

// 候補群から "採用源" を決める（優先: case_key → case_id → line_id）
// ソース提供内容の要約（機微値は短縮）
if (typeof describeSource_ !== 'function') {
  function describeSource_(label, m) {
    if (!m) return null;
    return {
      src: label,
      has: {
        case_key: !!(m.case_key || m.caseKey),
        case_id:  !!(m.case_id  || m.caseId),
        line_id:  !!(m.line_id  || m.lineId),
      },
      sample: {
        case_key: String(m.case_key || m.caseKey || '').slice(0, 24),
        case_id:  String(m.case_id  || m.caseId  || '').slice(0, 8),
        line_id:  String(m.line_id  || m.lineId  || '').slice(0, 12),
      }
    };
  }
}
if (typeof buildCandidates_ !== 'function') {
  function buildCandidates_(cands) {
    var arr = [];
    function push(lbl, m) { var d = describeSource_(lbl, m); if (d) arr.push(d); }
    push('cases',    cands && cands.fromCases);
    push('contacts', cands && cands.fromContacts);
    push('line',     cands && cands.fromLine);
    push('mail',     cands && cands.metaInMail);
    return arr;
  }
}
if (typeof resolveMetaWithPriority_ !== 'function') {
  function resolveMetaWithPriority_(cands, known) {
    var candidates = buildCandidates_(cands);
    // 優先度順で一致探索
    var ordered = [];
    for (var i = 0; i < candidates.length; i++) if (candidates[i].has.case_key) ordered.push(candidates[i]);
    for (var j = 0; j < candidates.length; j++) if (!candidates[j].has.case_key && candidates[j].has.case_id) ordered.push(candidates[j]);
    for (var k = 0; k < candidates.length; k++) if (!candidates[k].has.case_key && !candidates[k].has.case_id && candidates[k].has.line_id) ordered.push(candidates[k]);
    var decided = { meta: {}, by: 'fallback', source: 'unknown', candidates: candidates };
    for (var t = 0; t < ordered.length; t++) {
      var c = ordered[t];
      var raw = (c.src === 'cases')    ? (cands && cands.fromCases)
              : (c.src === 'contacts') ? (cands && cands.fromContacts)
              : (c.src === 'line')     ? (cands && cands.fromLine)
              :                           (cands && cands.metaInMail);
      var res = matchMetaToCase_(raw || {}, known || {});
      if (res && res.ok) { decided = { meta: raw || {}, by: res.by, source: c.src, candidates: candidates }; break; }
    }
    if (decided.by === 'fallback') {
      for (var u = 0; u < candidates.length; u++) {
        var cc = candidates[u];
        if (cc.has.case_key || cc.has.case_id || cc.has.line_id) {
          decided.source = cc.src;
          decided.meta = (cc.src === 'cases')    ? (cands && cands.fromCases)
                      : (cc.src === 'contacts') ? (cands && cands.fromContacts)
                      : (cc.src === 'line')     ? (cands && cands.fromLine)
                      :                           (cands && cands.metaInMail);
          break;
        }
      }
    }
    return decided;
  }
}

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

// ===== 受付メール本文の正規化・email抽出・meta補完（staging 保存前に通す） =====
function _htmlToText_(html) {
  return String(html || '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<br\s*\/?>(?!\n)/gi, '\n')
    .replace(/<\/(p|li|div|h\d)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/\r/g, '')
    .trim();
}

function _extractEmail_({ msg, obj, rawText }) {
  let email =
    (obj && obj.model && (obj.model.email || obj.model['メールアドレス'])) ||
    (obj && obj.fields && (obj.fields.email || obj.fields['メールアドレス'])) ||
    '';
  if (!email) {
    try {
      const html = (msg && msg.getBody && msg.getBody()) || '';
      const text = (msg && msg.getPlainBody && msg.getPlainBody()) || '';
      const body = _htmlToText_(html) || text || rawText || '';
      const m =
        body.match(/メール\s*アドレス\s*[:：]\s*([^\s<>]+@[^\s<>]+)/i) ||
        body.match(/\b([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})\b/i);
      email = m ? String(m[1]).trim() : '';
    } catch (_) {}
  }
  if (!email && msg && typeof msg.getAttachments === 'function') {
    try {
      const atts = msg.getAttachments();
      for (var i = 0; i < atts.length; i++) {
        const att = atts[i];
        const ct = String(att.getContentType() || '').toLowerCase();
        const nm = String(att.getName() || '').toLowerCase();
        if (/json/.test(ct) || /\.json$/.test(nm)) {
          try {
            const j = JSON.parse(att.copyBlob().getDataAsString('utf-8')) || {};
            email = j.email || (j.model && j.model.email) || (j.fields && j.fields.email) || '';
            if (email) {
              obj.model = Object.assign({}, obj.model || {}, j.model || {});
              obj.fields = Object.assign({}, obj.fields || {}, j.fields || {});
              break;
            }
          } catch (_) {}
        }
      }
    } catch (_) {}
  }
  // Gmail 正規化（ドット無視・+以降切り落とし）
  email = normalizeGmail_(email);
  return String(email || '').trim();
}

// Gmail アドレスの正規化（ドット無視・+以降切り落とし）
function normalizeGmail_(addr) {
  addr = String(addr || '').trim();
  const m = addr.toLowerCase().match(/^([^@+]+)(\+[^@]*)?@gmail\.com$/);
  if (!m) return addr;
  const local = m[1].replace(/\./g, '');
  return local + '@gmail.com';
}

// cases から line_id（or case_id）で事実を引く
function fi_casesLookup_(keys) {
  keys = keys || {};
  var lineId = String(keys.lineId || '').trim();
  var caseId = String(keys.caseId || '').replace(/\D/g, '');
  if (caseId) caseId = ('0000' + caseId).slice(-4);
  try {
    var sp = PropertiesService.getScriptProperties();
    var sid = String(
      sp.getProperty('BAS_MASTER_SPREADSHEET_ID') ||
        sp.getProperty('SHEET_ID') ||
        sp.getProperty('MASTER_SPREADSHEET_ID') ||
        ''
    ).trim();
    if (!sid) return null;
    var ss = SpreadsheetApp.openById(sid);
    var sh = ss.getSheetByName('cases');
    if (!sh) return null;
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) return null;
    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
    var cLine = headers.indexOf('line_id');
    var cCid  = headers.indexOf('case_id');
    var cUk   = headers.indexOf('user_key');
    var cCk   = headers.indexOf('case_key');
    var cFld  = headers.indexOf('folder_id');
    var cStat = headers.indexOf('status');
    var rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      var lid = cLine >= 0 ? String(rows[i][cLine] || '').trim() : '';
      var cid = cCid  >= 0 ? String(rows[i][cCid]  || '').replace(/\D/g, '') : '';
      if (cid) cid = ('0000' + cid).slice(-4);
      if ((lineId && lid && lid === lineId) || (caseId && cid && cid === caseId)) {
        var uk   = cUk  >= 0 ? String(rows[i][cUk ] || '').trim() : '';
        var ck   = cCk  >= 0 ? String(rows[i][cCk ] || '').trim() : '';
        var fid  = cFld >= 0 ? String(rows[i][cFld] || '').trim() : '';
        var stat = cStat>= 0 ? String(rows[i][cStat]|| '').trim() : '';
        if (!ck && uk && cid) ck = uk + '-' + cid;
        return { line_id: lid, user_key: uk, case_id: cid, case_key: ck, folder_id: fid, status: stat };
      }
    }
  } catch (e) {
    try { Logger.log('[fi_casesLookup_] ' + e); } catch (_) {}
  }
  return null;
}

// ★内部専用：このファイルだけで完結する contacts 逆引き
function fi_contactsLookupByEmail_(email) {
  function canon(s) {
    try { s = String(s || '').normalize('NFKC'); }
    catch (_) { s = String(s || ''); }
    s = s
      .toLowerCase()
      .replace(/[\u00A0\u200B\u200C\u200D\uFEFF\u2060]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
    var at = s.lastIndexOf('@');
    if (at > 0) {
      var local = s.slice(0, at);
      var domain = s.slice(at + 1);
      if (domain === 'googlemail.com') domain = 'gmail.com';
      if (domain === 'gmail.com') {
        local = local.replace(/[.\uFF0E]/g, '').replace(/[+＋].*$/, '');
      }
      s = local + '@' + domain;
    }
    return s;
  }
  var needle = canon(email);
  if (!needle) return null;

  var sp = PropertiesService.getScriptProperties();
  var sid = String(
    sp.getProperty('BAS_MASTER_SPREADSHEET_ID') ||
      sp.getProperty('SHEET_ID') ||
      sp.getProperty('MASTER_SPREADSHEET_ID') ||
      ''
  ).trim();
  if (!sid) {
    try { Logger.log('[fi_contacts] NO_SPREADSHEET_ID'); } catch (_) {}
    return null;
  }
  function sha256hex(str) {
    try {
      var raw = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        str,
        Utilities.Charset.UTF_8
      );
      var out = '';
      for (var i = 0; i < raw.length; i++) {
        var b = (raw[i] + 256) % 256;
        out += ('0' + b.toString(16)).slice(-2);
      }
      return out;
    } catch (_) {
      return '';
    }
  }
  var ss, sh;
  try {
    ss = SpreadsheetApp.openById(sid);
  } catch (e) {
    try { Logger.log('[fi_contacts] openById error: %s', e); } catch (_) {}
    return null;
  }
  sh = ss.getSheetByName('contacts');
  if (!sh) {
    try { Logger.log('[fi_contacts] NO_CONTACTS_SHEET in %s', ss.getName()); } catch (_) {}
    return null;
  }
  var headers = sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function (h) {
      return String(h).trim();
    });
  var cEmail = headers.indexOf('email');
  var cHash = headers.indexOf('email_hash');
  var cLine = headers.indexOf('line_id');
  var cUser = headers.indexOf('user_key');
  var cAci = headers.indexOf('active_case_id');
  var last = sh.getLastRow();
  if (last < 2) return null;
  var rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  var needleHash = sha256hex(needle);
  for (var i = rows.length - 1; i >= 0; i--) {
    var ok = false;
    var by = '';
    if (cEmail >= 0) {
      var e = canon(rows[i][cEmail]);
      if (e && e === needle) { ok = true; by = 'email'; }
    }
    if (!ok && cHash >= 0) {
      var hv = String(rows[i][cHash] || '').trim().toLowerCase();
      if (hv && hv === needleHash) { ok = true; by = 'email_hash'; }
    }
    if (!ok) continue;
    var line_id = cLine >= 0 ? String(rows[i][cLine] || '').trim() : '';
    var user_key = cUser >= 0 ? String(rows[i][cUser] || '').trim() : '';
    var aci = cAci >= 0 ? String(rows[i][cAci] || '').trim() : '';
    if (aci) aci = aci.replace(/\D/g, '').padStart(4, '0');
    try { Logger.log('[fi_contacts] HIT lid=%s uk=%s aci=%s (by=%s)', line_id, user_key, aci, by); } catch (_) {}
    return { line_id: line_id, user_key: user_key, active_case_id: aci };
  }
  try {
    var samples = [];
    for (var j = rows.length - 1; j >= 0 && samples.length < 3; j--) {
      var raw = String(rows[j][cEmail >= 0 ? cEmail : 0] || '');
      samples.push(canon(raw));
    }
    Logger.log('[fi_contacts] NO_HIT email=%s hash=%s; samples=%s', needle, needleHash, samples.join(', '));
  } catch (_) {}
  return null;
}

function fi_userKeyFromLineId_(lineId) {
  var lid = String(lineId || '').trim();
  if (!lid) return '';
  try {
    var sp = PropertiesService.getScriptProperties();
    var sid = String(
      sp.getProperty('BAS_MASTER_SPREADSHEET_ID') ||
        sp.getProperty('SHEET_ID') ||
        sp.getProperty('MASTER_SPREADSHEET_ID') ||
        ''
    ).trim();
    if (!sid) return '';
    var ss = SpreadsheetApp.openById(sid);
    var sh = ss.getSheetByName('contacts');
    if (!sh) return '';
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    var lineIdx = headers.indexOf('line_id');
    var userIdx = headers.indexOf('user_key');
    if (lineIdx < 0 || userIdx < 0) return '';
    var last = sh.getLastRow();
    if (last < 2) return '';
    var rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      var lv = String(rows[i][lineIdx] || '').trim();
      if (lv && lv === lid) return String(rows[i][userIdx] || '').trim();
    }
  } catch (_) {}
  return '';
}

function _fillMetaBeforeStage_(obj, payload) {
  payload = payload || {};
  var knownLineId = String(payload.knownLineId || '').trim();
  var knownCaseId = String(payload.knownCaseId || '').trim();
  var email = String(payload.email || '').trim();

  obj = obj && typeof obj === 'object' ? obj : {};
  obj.meta = obj.meta || {};

  // 0) cases シートを最優先に参照して "事実" を反映
  var sourceLine = '';
  var sourceCase = '';
  try {
    var lidHint = String(obj.meta.line_id || knownLineId || obj.model?.line_id || obj.fields?.line_id || '').trim();
    var cidHintRaw = String(obj.meta.case_id || knownCaseId || '').replace(/\D/g, '');
    var cidHint = cidHintRaw ? ('0000' + cidHintRaw).slice(-4) : '';
    var fromCases = (typeof fi_casesLookup_ === 'function') ? fi_casesLookup_({ lineId: lidHint, caseId: cidHint }) : null;
    if (fromCases) {
      if (fromCases.line_id   && !obj.meta.line_id)  obj.meta.line_id  = fromCases.line_id;
      if (fromCases.line_id) sourceLine = 'cases';
      if (fromCases.user_key  && !obj.meta.user_key) obj.meta.user_key = fromCases.user_key;
      if (fromCases.case_id   && !obj.meta.case_id)  obj.meta.case_id  = fromCases.case_id; // 常に4桁
      if (fromCases.case_id) sourceCase = 'cases';
      if (fromCases.case_key  && !obj.meta.case_key) obj.meta.case_key = fromCases.case_key;
      if (fromCases.folder_id) { obj.model = obj.model || {}; obj.model.folder_id = fromCases.folder_id; }
      if (fromCases.status)    { obj.model = obj.model || {}; obj.model.case_status = fromCases.status; }
    }
  } catch (_) {}

  // 1) line_id
  var lid = String(obj.meta.line_id || knownLineId || '').trim();
  if (!sourceLine && lid && knownLineId && lid === knownLineId) sourceLine = 'ctx';
  if (!lid && email) {
    try {
      // 外部実装に依存せず、内部版を無条件に使用
      var c = fi_contactsLookupByEmail_(email);
      if (c && c.line_id) {
        lid = c.line_id;
        if (!knownCaseId && c.active_case_id) knownCaseId = c.active_case_id;
        if (!sourceLine) sourceLine = 'contacts';
        if (!sourceCase && c.active_case_id) sourceCase = 'contacts';
      }
    } catch (_) {}
  }
  // 2) user_key
  var ukey = String(obj.meta.user_key || '').trim();
  if (!ukey && lid) {
    try { ukey = fi_userKeyFromLineId_(lid) || ''; } catch (_) {}
  }
  // 3) case_id（4桁）
  var cid = String(obj.meta.case_id || knownCaseId || '').replace(/\D/g, '');
  if (cid) cid = ('0000' + cid).slice(-4);
  if (cid === '0000') cid = '';
  // 4) 書き戻し
  if (lid && !obj.meta.line_id) obj.meta.line_id = lid;
  if (ukey && !obj.meta.user_key) obj.meta.user_key = ukey;
  if (cid && !obj.meta.case_id) obj.meta.case_id = cid;
  if (!obj.meta.case_key && ukey && cid) obj.meta.case_key = ukey + '-' + cid;
  // source 記録（診断用）
  try {
    obj.model = obj.model || {};
    if (!obj.model._source_line) obj.model._source_line = sourceLine || (obj.meta.line_id ? 'unknown' : 'none');
    if (!obj.model._source_case) obj.model._source_case = sourceCase || (obj.meta.case_id ? 'unknown' : 'none');
  } catch (_) {}
  // email も保持
  if (email) {
    obj.model = obj.model || {};
    if (!obj.model.email) obj.model.email = email;
  }
  return obj;
}

// === まず「line_id を最優先で確定」する集約関数 ===
function getLineIdFromContext_(req, msg, obj) {
  try {
    var p = (req && req.parameter) || {};
    var lid = String(p.lineId || p.lid || '').trim();
    if (lid) return lid;
  } catch (_) {}

  try {
    var m = (obj && obj.meta) || {};
    var lid2 = String(m.line_id || m.lineId || '').trim();
    if (lid2) return lid2;
    var model = (obj && obj.model) || {};
    lid2 = String(model.line_id || model.lineId || '').trim();
    if (lid2) return lid2;
    var fields = (obj && obj.fields) || {};
    lid2 = String(fields.line_id || fields.lineId || '').trim();
    if (lid2) return lid2;
  } catch (_) {}

  try {
    var hdrs = (req && req.headers) || {};
    var lid3 = String(hdrs['x-line-user-id'] || hdrs['X-Line-User-Id'] || '').trim();
    if (lid3) return lid3;
  } catch (_) {}

  try {
    var cookie = String((hdrs && (hdrs.cookie || hdrs.Cookie)) || '').trim();
    var m2 = cookie && cookie.match(/(?:^|;\s*)lid=([^;]+)/);
    if (m2) {
      var lid4 = decodeURIComponent(m2[1] || '').trim();
      if (lid4) return lid4;
    }
  } catch (_) {}

  try {
    var sp = PropertiesService.getScriptProperties();
    var lid5 = String(sp.getProperty('LAST_LINE_ID') || '').trim();
    if (lid5) return lid5;
  } catch (_) {}

  return '';
}

// ===== contacts 逆引きヘルパ（このデプロイに無い場合のフォールバック） =====
if (typeof contacts_lookupByEmail_ !== 'function') {
  function contacts_lookupByEmail_(email) {
    email = String(email || '').trim().toLowerCase();
    if (!email) return null;

    var sid =
      PropertiesService.getScriptProperties().getProperty('BAS_MASTER_SPREADSHEET_ID') ||
      PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!sid) {
      try { Logger.log('[contacts_lookup] NO_SPREADSHEET_ID'); } catch (_) {}
      return null;
    }
    var ss, sh;
    try {
      ss = SpreadsheetApp.openById(sid);
    } catch (e) {
      try { Logger.log('[contacts_lookup] openById error: ' + e); } catch (_) {}
      return null;
    }
    sh = ss.getSheetByName('contacts');
    if (!sh) {
      try { Logger.log('[contacts_lookup] NO_CONTACTS_SHEET'); } catch (_) {}
      return null;
    }
    var headers = sh
      .getRange(1, 1, 1, sh.getLastColumn())
      .getValues()[0]
      .map(function (h) {
        return String(h).trim();
      });
    var emailIdx = headers.indexOf('email');
    var lineIdx = headers.indexOf('line_id');
    var userIdx = headers.indexOf('user_key');
    var aciIdx = headers.indexOf('active_case_id');
    if (emailIdx < 0) {
      try { Logger.log('[contacts_lookup] NO email column'); } catch (_) {}
      return null;
    }
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return null;
    var rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    var needle = email.toLowerCase();
    for (var i = rows.length - 1; i >= 0; i--) {
      var e = String(rows[i][emailIdx] || '').trim().toLowerCase();
      if (!e) continue;
      if (e === needle) {
        var aci = aciIdx >= 0 ? String(rows[i][aciIdx] || '').trim() : '';
        aci = aci ? aci.replace(/\D/g, '').padStart(4, '0') : '';
        return {
          line_id: lineIdx >= 0 ? String(rows[i][lineIdx] || '').trim() : '',
          user_key: userIdx >= 0 ? String(rows[i][userIdx] || '').trim() : '',
          active_case_id: aci,
        };
      }
    }
    return null;
  }
}

if (typeof drive_userKeyFromLineId_ !== 'function') {
  function drive_userKeyFromLineId_(lineId) {
    var lid = String(lineId || '').trim();
    if (!lid) return '';
    try {
      var sid =
        PropertiesService.getScriptProperties().getProperty('BAS_MASTER_SPREADSHEET_ID') ||
        PropertiesService.getScriptProperties().getProperty('SHEET_ID');
      if (!sid) return '';
      var ss = SpreadsheetApp.openById(sid);
      var sh = ss.getSheetByName('contacts');
      if (!sh) return '';
      var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      var lineIdx = headers.indexOf('line_id');
      var userIdx = headers.indexOf('user_key');
      if (lineIdx < 0 || userIdx < 0) return '';
      var last = sh.getLastRow();
      if (last < 2) return '';
      var rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
      for (var i = rows.length - 1; i >= 0; i--) {
        var lv = String(rows[i][lineIdx] || '').trim();
        if (lv && lv === lid) return String(rows[i][userIdx] || '').trim();
      }
    } catch (_) {}
    return '';
  }
}

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
          // line_id を多系統から最優先で抽出
          var knownLineIdEarly = '';
          try { knownLineIdEarly = getLineIdFromContext_(null /*req*/, msg, parsed) || ''; } catch (_) {}
          // caseId はパラメータ or 既存meta
          var knownCaseIdEarly = '';
          try { knownCaseIdEarly = String(meta.case_id || meta.caseId || '').replace(/\D/g, ''); } catch (_) {}

          // メタ補完（email抽出＋contacts 逆引き）→ バリデーション（隔離）
          const emailEarly = (function () {
            try { return _extractEmail_({ msg, obj: parsed, rawText: body }); } catch (_) { return ''; }
          })();
          try {
            parsed = _fillMetaBeforeStage_(parsed, {
              knownLineId: knownLineIdEarly,
              knownCaseId: knownCaseIdEarly,
              email: emailEarly,
            });
          } catch (_) {}
          if (!(emailEarly || (parsed && parsed.meta && parsed.meta.line_id))) {
            try {
              const itQe = root.getFoldersByName('_quarantine');
              const qfe = itQe.hasNext() ? itQe.next() : root.createFolder('_quarantine');
              const qne = `raw_intake__${meta.submission_id || Date.now()}.json`;
              qfe.createFile(
                Utilities.newBlob(
                  JSON.stringify({ raw: parsed, note: 'no_email_no_lineid' }),
                  'application/json',
                  qne
                )
              );
              Logger.log('[Intake][drop] no identifiers; quarantined %s', qne);
            } catch (_) {}
            formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, true);
            return;
          }
          // 追加: 保存直前に、候補群から採用源と優先キーを確定
          try {
            var ukEarly = '';
            try { ukEarly = fi_userKeyFromLineId_(String(parsed?.meta?.line_id || knownLineIdEarly || '')) || ''; } catch (_) {}
            var knownEarly = {
              case_key: (ukEarly && knownCaseIdEarly) ? (String(ukEarly).toLowerCase() + '-' + normCaseId_(knownCaseIdEarly)) : '',
              case_id: normCaseId_(knownCaseIdEarly || ''),
              line_id: String(parsed?.meta?.line_id || knownLineIdEarly || ''),
            };
            var fromCasesEarly = (typeof fi_casesLookup_ === 'function') ? fi_casesLookup_({ lineId: knownEarly.line_id, caseId: knownEarly.case_id }) : null;
            var fromContactsEarly = (typeof fi_contactsLookupByEmail_ === 'function') ? fi_contactsLookupByEmail_(emailEarly) : null;
            var candEarly = {
              fromCases: fromCasesEarly,
              fromContacts: fromContactsEarly,
              fromLine: { line_id: knownEarly.line_id, user_key: ukEarly },
              metaInMail: parsed && parsed.meta,
            };
            var rEarly = resolveMetaWithPriority_(candEarly, knownEarly);
            var mEarly = rEarly && rEarly.meta || {};
            parsed.meta = parsed.meta || {};
            parsed.meta.user_key = parsed.meta.user_key || mEarly.user_key || ukEarly || parsed.meta.userKey || '';
            parsed.meta.case_id  = normCaseId_(parsed.meta.case_id || parsed.meta.caseId || mEarly.case_id || knownEarly.case_id || '');
            parsed.meta.case_key = normCaseKey_(parsed.meta.case_key || parsed.meta.caseKey || mEarly.case_key || ((parsed.meta.user_key && parsed.meta.case_id) ? (String(parsed.meta.user_key).toLowerCase() + '-' + parsed.meta.case_id) : ''));
            parsed.meta.line_id  = parsed.meta.line_id || mEarly.line_id || knownEarly.line_id || '';
            try { Logger.log('[Intake] adopt(by=%s,source=%s) candidates=%s', rEarly.by || '', rEarly.source || '', JSON.stringify(buildCandidates_(candEarly))); } catch (_) {}
          } catch (_) {}
          // 追加: 保存直前に、候補群から採用源と優先キーを確定（後段保存でも同様の形に）
          try {
            var uk2 = '';
            try { uk2 = fi_userKeyFromLineId_(String(parsed?.meta?.line_id || knownLineId2 || '')) || ''; } catch (_) {}
            var known2 = {
              case_key: (uk2 && knownCaseId2) ? (String(uk2).toLowerCase() + '-' + normCaseId_(knownCaseId2)) : '',
              case_id: normCaseId_(knownCaseId2 || ''),
              line_id: String(parsed?.meta?.line_id || knownLineId2 || ''),
            };
            var fromCases2 = (typeof fi_casesLookup_ === 'function') ? fi_casesLookup_({ lineId: known2.line_id, caseId: known2.case_id }) : null;
            var fromContacts2 = (typeof fi_contactsLookupByEmail_ === 'function') ? fi_contactsLookupByEmail_(email2) : null;
            var cand2 = {
              fromCases: fromCases2,
              fromContacts: fromContacts2,
              fromLine: { line_id: known2.line_id, user_key: uk2 },
              metaInMail: parsed && parsed.meta,
            };
            var r2 = resolveMetaWithPriority_(cand2, known2);
            var m2 = r2 && r2.meta || {};
            parsed.meta = parsed.meta || {};
            parsed.meta.user_key = parsed.meta.user_key || m2.user_key || uk2 || parsed.meta.userKey || '';
            parsed.meta.case_id  = normCaseId_(parsed.meta.case_id || parsed.meta.caseId || m2.case_id || known2.case_id || '');
            parsed.meta.case_key = normCaseKey_(parsed.meta.case_key || parsed.meta.caseKey || m2.case_key || ((parsed.meta.user_key && parsed.meta.case_id) ? (String(parsed.meta.user_key).toLowerCase() + '-' + parsed.meta.case_id) : ''));
            parsed.meta.line_id  = parsed.meta.line_id || m2.line_id || known2.line_id || '';
            try { Logger.log('[Intake] adopt(by=%s,source=%s) candidates=%s', r2.by || '', r2.source || '', JSON.stringify(buildCandidates_(cand2))); } catch (_) {}
          } catch (_) {}
          const fname = `intake__${meta.submission_id}.json`;
          const blob = Utilities.newBlob(JSON.stringify(parsed, null, 2), 'application/json', fname);
          staging.createFile(blob);
          try {
            var candidatesEarlyJson = (function(){ try { return JSON.stringify(buildCandidates_(candEarly)); } catch(_) { return '[]'; } })();
            Logger.log(
              '[Intake] staged name=%s meta={lid:%s, uk:%s, cid:%s, ckey:%s} by=%s source=%s candidates=%s',
              fname,
              (parsed && parsed.meta && parsed.meta.line_id) || '',
              (parsed && parsed.meta && parsed.meta.user_key) || '',
              (parsed && parsed.meta && parsed.meta.case_id) || '',
              (parsed && parsed.meta && parsed.meta.case_key) || '',
              (rEarly && rEarly.by) || '',
              (rEarly && rEarly.source) || 'unknown',
              candidatesEarlyJson
            );
          } catch (_) {}
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

        // メタ確定順: case_key → case_id → line_id（case_key を最優先で決定）
        let resolvedCaseKey = caseInfo.caseKey && String(caseInfo.caseKey) ? String(caseInfo.caseKey) : '';
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
          // line_id を多系統から抽出 → メタ補完（email抽出＋contacts逆引き）→ バリデーション（隔離）
          var knownLineId2 = '';
          try { knownLineId2 = getLineIdFromContext_(null /*req*/, msg, parsed) || ''; } catch (_) {}
          var knownCaseId2 = '';
          try { knownCaseId2 = String(meta.case_id || meta.caseId || '').replace(/\D/g, ''); } catch (_) {}
          const email2 = (function () {
            try { return _extractEmail_({ msg, obj: parsed, rawText: body }); } catch (_) { return ''; }
          })();
          try {
            parsed = _fillMetaBeforeStage_(parsed, {
              knownLineId: knownLineId2,
              knownCaseId: knownCaseId2,
              email: email2,
            });
          } catch (_) {}
          if (!(email2 || (parsed && parsed.meta && parsed.meta.line_id))) {
            try {
              const itQ2 = root.getFoldersByName('_quarantine');
              const qf2 = itQ2.hasNext() ? itQ2.next() : root.createFolder('_quarantine');
              const qn2 = `raw_intake__${meta.submission_id || Date.now()}.json`;
              qf2.createFile(
                Utilities.newBlob(
                  JSON.stringify({ raw: parsed, note: 'no_email_no_lineid' }),
                  'application/json',
                  qn2
                )
              );
              Logger.log('[Intake][drop] no identifiers; quarantined %s', qn2);
            } catch (_) {}
            formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, true);
            return;
          }
          const fname = `intake__${meta.submission_id}.json`;
          const blob = Utilities.newBlob(JSON.stringify(parsed, null, 2), 'application/json', fname);
          staging.createFile(blob);
          try {
            var candidates2Json = (function(){ try { return JSON.stringify(buildCandidates_(cand2)); } catch(_) { return '[]'; } })();
            Logger.log(
              '[Intake] staged name=%s meta={lid:%s, uk:%s, cid:%s, ckey:%s} by=%s source=%s candidates=%s',
              fname,
              (parsed && parsed.meta && parsed.meta.line_id) || '',
              (parsed && parsed.meta && parsed.meta.user_key) || '',
              (parsed && parsed.meta && parsed.meta.case_id) || '',
              (parsed && parsed.meta && parsed.meta.case_key) || '',
              (r2 && r2.by) || '',
              (r2 && r2.source) || 'unknown',
              candidates2Json
            );
          } catch (_) {}
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

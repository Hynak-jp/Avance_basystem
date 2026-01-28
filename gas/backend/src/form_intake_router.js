/**
 * 共通フォーム Intake ルーター
 *  - FormAttach/ToProcess に入った通知メールをフォーム種別ごとに振り分け、caseフォルダへ JSON 保存
 *  - FORM_REGISTRY にエントリを追加するだけで新フォームを取り込める
 */

const FORM_INTAKE_LABEL_TO_PROCESS = 'FormAttach/ToProcess';
const FORM_INTAKE_LABEL_PROCESSED = 'FormAttach/Processed';
const FORM_INTAKE_LABEL_ERROR = 'FormAttach/Error';
const FORM_INTAKE_LABEL_NO_META = 'FormAttach/NoMeta';
const FORM_INTAKE_LABEL_REJECTED = 'FormAttach/Rejected';
const FORM_INTAKE_LABEL_LOCK = ''; // ScriptLockで排他するためラベルロックは廃止
const FORM_INTAKE_LABEL_ATTACH_SAVED = 'FormAttach/AttachmentsSaved';

// ===== FIELDS ブロック抽出（メール本文から） =====
if (typeof parseFieldsBlock_ !== 'function') {
  function parseFieldsBlock_(text) {
    try {
      text = String(text || '');
      var m = text.match(/====\s*FIELDS START\s*====([\s\S]*?)====\s*FIELDS END\s*====/i);
      if (!m) return [];
      var body = m[1].replace(/\r/g, '');
      var lines = body.split('\n');
      var out = [];
      var curLabel = '', curValue = [];
      function flush() {
        if (!curLabel) return;
        var val = curValue.join('\n').replace(/^\s+|\s+$/g, '');
        out.push({ label: curLabel, value: val });
        curLabel = ''; curValue = [];
      }
      lines.forEach(function (ln) {
        var l = ln.replace(/\s+$/, '');
        var m2 = l.match(/^([【].+?[】])\s*$/);
        if (m2) { flush(); curLabel = m2[1]; curValue = []; }
        else {
          // 先頭の全角スペースは維持
          curValue.push(l.replace(/^\u3000/, '　'));
        }
      });
      flush();
      return out;
    } catch (_) {
      return [];
    }
  }
}

// _email_staging の月別/ハッシュ配下にフォルダを作成して ID を返す
function formIntake_getOrCreateEmailStagingFolderId_(emailRaw) {
  const root = drive_getRootFolder_();
  const yyyymm = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
  const email = String(emailRaw || '').trim().toLowerCase();
  const hash = email
    ? Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email, Utilities.Charset.UTF_8)
        .map(function (b) {
          const v = b < 0 ? b + 256 : b;
          return v.toString(16).padStart(2, '0');
        })
        .join('')
    : 'unknown';
  const folder = drive_getOrCreatePath_(root, ['_email_staging', yyyymm, hash].join('/'));
  return folder.getId();
}
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

// write-through mover: staging 保存直後に、即ケース直下へ移送を試みる
if (typeof tryMoveIntakeToCase_ !== 'function') {
  function tryMoveIntakeToCase_(meta, file, fileName) {
    try {
      // --- FIELDS ブロックの整形ヘルパ ---
      function toStr_(v) { return v == null ? '' : String(v); }
      function buildFieldsBlock_(m) {
        m = m || {};
        var name   = toStr_(m.name || m.full_name || m['氏名']);
        var email  = toStr_(m.email || m['メールアドレス']);
        var phone  = toStr_(m.phone || m.tel || m['電話番号']);
        var gender = toStr_(m.gender || m['性別']);
        var dob    = toStr_(m.birth || m.birthdate || m['生年月日']);
        // 住所は項目が分割されていても結合
        var postal = toStr_(m.postal || m.zip || m['郵便番号']);
        var addr   = [
          toStr_(m.pref || m['都道府県']),
          toStr_(m.city || m['市区町村']),
          toStr_(m.addr1 || m['番地以降']),
          toStr_(m.addr2 || m['建物名・部屋番号']),
        ].filter(Boolean).join(' ');
        var address = [postal ? postal : '', addr].filter(Boolean).join(' ');
        return [
          '==== FIELDS START ====','【名前】','　' + name,
          '【メールアドレス】','　' + email,
          '【電話番号】','　' + phone,
          '【性別】','　' + gender,
          '【生年月日】','　' + dob,
          '【住所】','　' + address,
          '','',
          '==== FIELDS END ===='
        ].join('\n');
      }
      meta = meta || {};
      var uk  = String(meta.user_key || meta.userKey || '').trim();
      var cid = normCaseId_(meta.case_id || meta.caseId || '');
      var ckey= normCaseKey_(meta.case_key || meta.caseKey || (uk && cid ? (String(uk).toLowerCase() + '-' + cid) : ''));
      var lid = String(meta.line_id || meta.lineId || '').trim();
      if (lid && typeof fi_casesLookup_ === 'function') {
        try {
          var hit = fi_casesLookup_({ lineId: lid, caseId: cid });
          if (hit) {
            if (!cid && hit.case_id) cid = normCaseId_(hit.case_id);
            if (!uk && hit.user_key) uk = String(hit.user_key || '').trim();
            if (!ckey && hit.case_key) ckey = normCaseKey_(hit.case_key);
          }
        } catch (_) {}
      }
      if (!cid && lid && typeof lookupCaseIdByLineId_ === 'function') {
        try { cid = normCaseId_(lookupCaseIdByLineId_(lid)); } catch (_) {}
      }
      if (!uk && lid && typeof drive_userKeyFromLineId_ === 'function') {
        try { uk = String(drive_userKeyFromLineId_(lid) || '').trim(); } catch (_) {}
      }
      if (!ckey && uk && cid) {
        ckey = normCaseKey_(String(uk).toLowerCase() + '-' + cid);
      }
      var known = { case_key: ckey, case_id: cid, line_id: lid };

      // ケースフォルダ解決（case_key 優先 → 予備として case_id 名の直一致）
      var caseFolder = null;
      try {
        if (ckey) {
          var it = DriveApp.getFoldersByName(ckey);
          if (it && it.hasNext()) caseFolder = it.next();
        }
        if (!caseFolder && cid) {
          var it2 = DriveApp.getFoldersByName(cid);
          if (it2 && it2.hasNext()) caseFolder = it2.next();
        }
      } catch (_) {}
      if (!caseFolder) {
        try { Logger.log('[router.move] no case folder for %s (ckey=%s,cid=%s,lid=%s)', fileName || '', ckey || '', cid || '', lid || ''); } catch(_){ }
        return false;
      }

      // 一致確認（冪等・安全）
      var res = matchMetaToCase_(meta, known);
      if (!res || !res.ok) {
        try { Logger.log('[router.move] skip: meta not matched to known for %s', fileName || ''); } catch(_){ }
        return false;
      }

      // JSON 整形・補完して保存（原本は削除）
      var txt = '';
      try { txt = file && file.getBlob && file.getBlob().getDataAsString('utf-8'); } catch(_){}
      var obj = {};
      try { obj = JSON.parse(txt || '{}') || {}; } catch(_) { obj = {}; }
      obj = obj && typeof obj === 'object' ? obj : {};
      obj.meta = obj.meta || {};
      if (!obj.meta.case_id)  obj.meta.case_id  = cid;
      if (!obj.meta.user_key) obj.meta.user_key = uk;
      if (!obj.meta.case_key) obj.meta.case_key = ckey;
      if (!obj.meta.line_id)  obj.meta.line_id  = lid;
      // ---- intake の正規化＆出力物（fields_block / fieldsRaw / model.schema） ----
      obj.model = obj.model || {};
      // 値取りのエイリアス（最初の非空を採用）
      function firstNonEmpty() {
        for (var i = 0; i < arguments.length; i++) {
          var v = arguments[i];
          if (v != null && String(v).trim() !== '') return String(v).trim();
        }
        return '';
      }
      var M = obj.model, ME = obj.meta || {};
      var FR = Array.isArray(obj.fieldsRaw) ? obj.fieldsRaw : [];
      var fMap = {};
      try { FR.forEach(function (e) { if (e && e.label) fMap[String(e.label)] = String(e.value == null ? '' : e.value).trim(); }); } catch(_){}
      var name   = firstNonEmpty(fMap['【名前】'],           M.name, M.full_name, M['氏名'], (M.applicant && M.applicant.name));
      var email  = firstNonEmpty(fMap['【メールアドレス】'], M.email, (M.applicant && M.applicant.email), ME.email);
      var phone  = firstNonEmpty(fMap['【電話番号】'],       M.phone, M.tel, M['電話番号'], (M.applicant && M.applicant.phone));
      var gender = firstNonEmpty(fMap['【性別】'],           M.gender, M['性別'], (M.applicant && M.applicant.gender));
      var birth  = firstNonEmpty(fMap['【生年月日】'],       M.birth, M.birthdate, M['生年月日'], (M.applicant && M.applicant.birth));
      var addrLineFromFields = firstNonEmpty(fMap['【住所】']);
      // postal はまず明示項目から。住所行に含まれる場合は後で補完
      var postal = firstNonEmpty(M.postal, M.zip, M['郵便番号'], (M.address && M.address.postal));
      var pref   = firstNonEmpty(M.pref, M['都道府県'], (M.address && M.address.pref));
      var city   = firstNonEmpty(M.city, M['市区町村'], (M.address && M.address.city));
      var addr1  = firstNonEmpty(M.addr1, M['番地以降'], (M.address && M.address.addr1));
      var addr2  = firstNonEmpty(M.addr2, M['建物名・部屋番号'], (M.address && M.address.addr2));
      var addressLine = '';
      if (addrLineFromFields) {
        addressLine = addrLineFromFields;
        // 住所行の先頭が郵便番号なら分離して補完（ハイフン無しも許容）
        try {
          var mm = addressLine.match(/^\s*([0-9]{3}-?[0-9]{4})\s*(.*)$/);
          if (mm) {
            if (!postal) postal = (mm[1] || '').replace(/^(\d{3})(\d{4})$/, '$1-$2');
            // 残部は rest として保持（pref/city 等は既存項目があればそちらを優先）
          }
        } catch (_) {}
      } else {
        addressLine = [postal ? postal : '', [pref, city, addr1, addr2].filter(Boolean).join(' ')].filter(Boolean).join(' ');
      }

      function buildFieldsBlock() {
        var lines = [
          '==== FIELDS START ====',
          '【名前】',
          '　' + (name || ''),
          '【メールアドレス】',
          '　' + (email || ''),
          '【電話番号】',
          '　' + (phone || ''),
          '【性別】',
          '　' + (gender || ''),
          '【生年月日】',
          '　' + (birth || ''),
          '【住所】',
          '　' + (addressLine || ''),
          '',
          '',
          '==== FIELDS END ===='],
          out = lines.join('\n');
        return out;
      }
      obj.model.fields_block = buildFieldsBlock();

      // 他フォームと同じ [{label,value}] 形式
      obj.fieldsRaw = [
        { label: '【名前】',           value: name },
        { label: '【メールアドレス】', value: email },
        { label: '【電話番号】',       value: phone },
        { label: '【性別】',           value: gender },
        { label: '【生年月日】',       value: birth },
        { label: '【住所】',           value: addressLine },
      ];

      // model の軽い正規化
      obj.model.schema = obj.model.schema || 'intake@v1';
      obj.model.applicant = obj.model.applicant || {};
      if (name)   obj.model.applicant.name   = name;
      if (email)  obj.model.applicant.email  = email;
      if (phone)  obj.model.applicant.phone  = phone;
      if (gender) obj.model.applicant.gender = gender;
      if (birth)  obj.model.applicant.birth  = birth;
      obj.model.address = obj.model.address || {};
      if (postal) obj.model.address.postal = postal;
      if (pref)   obj.model.address.pref   = pref;
      if (city)   obj.model.address.city   = city;
      if (addr1)  obj.model.address.addr1  = addr1;
      if (addr2)  obj.model.address.addr2  = addr2;
      if (addressLine) obj.model.address.line = addressLine;

      var saved = '';
      try { saved = JSON.stringify(obj, null, 2) + '\n'; } catch(_) { saved = (txt && (txt + '\n')) || '{}\n'; }
      var outBlob = Utilities.newBlob(saved, 'application/json; charset=utf-8', fileName || (file && file.getName && file.getName()) || ('intake__' + Date.now() + '.json'));
      try { caseFolder.createFile(outBlob); } catch(_){ }
      try { Logger.log('[router.move] saved bytes=%s name=%s', (outBlob && outBlob.getBytes && outBlob.getBytes().length) || -1, fileName || ''); } catch(_){ }
      try { file && file.setTrashed && file.setTrashed(true); } catch(_){}
      try { Logger.log('[router.move] moved name=%s by=%s → folder=%s', fileName || '', res.by || '', (caseFolder && caseFolder.getName && caseFolder.getName()) || ''); } catch(_){}

      // submissions 追記
      try {
        var sid = (obj && obj.meta && obj.meta.submission_id) || ((String(fileName || '').match(/^intake__(\d+)\.json$/i) || [])[1] || '');
        if (!sid) { try { var m2 = String(fileName || '').match(/__(\d+)\.json$/); if (m2) sid = m2[1]; } catch(_){}
        }
        if (!sid) sid = String(Date.now());
        if (typeof submissions_upsert_ === 'function') {
          submissions_upsert_({
            submission_id: sid,
            form_key: 'intake',
            case_id: cid,
            user_key: uk,
            line_id: lid,
            submitted_at: (obj && obj.meta && obj.meta.submitted_at) || new Date().toISOString(),
            status: 'intake',
            // 拡張カラム
            seq: (obj && obj.meta && obj.meta.seq) || '',
            referrer: (obj && obj.meta && obj.meta.referrer) || '',
            redirect_url: (obj && obj.meta && obj.meta.redirect_url) || ''
          });
        }
      } catch (_) {}
      return true;
    } catch (e) {
      try { Logger.log('[router.move] error %s', (e && e.message) || e); } catch(_){ }
      return false;
    }
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
    allowIssueNewCase: false,
    parser: function (subject, body) {
      if (typeof parseMetaBlock_ !== 'function') {
        throw new Error('parseMetaBlock_ is not defined');
      }
      const meta = parseMetaBlock_(body) || {};
      if (!meta.form_key) meta.form_key = 'intake';
      // FIELDS ブロックを本文から抽出
      const fieldsRaw = (typeof parseFieldsBlock_ === 'function') ? parseFieldsBlock_(body) : [];
      return { meta, fieldsRaw, model: {} };
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    statusAfterSave: 'intake',
    afterSave: function () { return {}; },
  },
  doc_payslip: {
    name: '書類提出（給与明細）',
    parser: function (subject, body) {
      if (typeof parseMetaBlock_ !== 'function') {
        throw new Error('parseMetaBlock_ is not defined');
      }
      const meta = parseMetaBlock_(body) || {};
      if (!meta.form_key) meta.form_key = 'doc_payslip';
      const fieldsRaw = (typeof parseFieldsBlock_ === 'function') ? parseFieldsBlock_(body) : [];
      return { meta, fieldsRaw, model: {} };
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    requireCaseId: true,
    afterSave: function () { return {}; },
  },

  s2010_p1_career: {
    name: 'S2010 Part1(経歴等)',
    parser: function (subject, body) {
      if (typeof parseFormMail_S2010_P1_ !== 'function') {
        throw new Error('parseFormMail_S2010_P1_ is not defined');
      }
      return parseFormMail_S2010_P1_(subject, body);
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    statusAfterSave: 'intake',
    requireCaseId: true,
    afterSave: function (caseInfo, parsed) {
      const caseId = caseInfo.caseId || caseInfo.case_id || (parsed && parsed.meta && parsed.meta.case_id) || '';
      const folderId = caseInfo.folderId || caseInfo.folder_id || '';
      if (!caseId || !folderId) return {};
      if (typeof haveAllPartsS2010_ !== 'function' || typeof run_GenerateS2010DraftMergedByCaseId !== 'function') {
        return {};
      }
      const prefixes =
        typeof S2010_PART_PREFIXES !== 'undefined' && S2010_PART_PREFIXES
          ? S2010_PART_PREFIXES
          : ['s2010_p1_', 's2010_p2_'];
      if (haveAllPartsS2010_(folderId, prefixes)) {
        run_GenerateS2010DraftMergedByCaseId(caseId);
      }
      return {};
    },
  },
  s2010_p1_intake: {
    name: 'S2010 Part1(経歴等)',
    parser: function (subject, body) {
      if (typeof parseFormMail_S2010_P1_ !== 'function') {
        throw new Error('parseFormMail_S2010_P1_ is not defined');
      }
      const parsed = parseFormMail_S2010_P1_(subject, body);
      if (parsed && parsed.meta && !parsed.meta.form_key) {
        parsed.meta.form_key = 's2010_p1_intake';
      }
      return parsed;
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    statusAfterSave: 'intake',
    requireCaseId: true,
    afterSave: function () {
      return {};
    },
  },
  s2002_userform: {
    name: 'S2002 申立',
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
        // ドラフト生成前に folderId を強制補正（未設定時）
        try {
          if (!caseInfo.folderId && typeof ensureCaseFolderId_ === 'function') {
            caseInfo.folderId = ensureCaseFolderId_(caseInfo);
          }
        } catch (_) {}
        const draft = generateS2002Draft_(caseInfo, parsed);
        const patch = {};
        if (draft && draft.draftUrl) patch.last_draft_url = draft.draftUrl;
        return patch;
      }
      return {};
    },
  },
  s2005_creditors: {
    name: 'S2005 債権者一覧表',
    parser: function (subject, body) {
      if (typeof parseMetaBlock_ !== 'function') {
        throw new Error('parseMetaBlock_ is not defined');
      }
      const meta = parseMetaBlock_(body) || {};
      meta.form_key = meta.form_key || 's2005_creditors';
      if (!meta.submission_id) {
        const m = subject.match(/submission_id:(\d+)/) || subject.match(/提出\s+(\d{6,})/);
        if (m) meta.submission_id = m[1];
      }
      const fields = (typeof parseFieldsBlock_ === 'function') ? parseFieldsBlock_(body) : [];
      const model =
        typeof mapS2005FieldsToModel_ === 'function' ? mapS2005FieldsToModel_(fields) : {};
      return { meta, fieldsRaw: fields, model: model };
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    statusAfterSave: 'draft',
    requireCaseId: true,
    afterSave: function (caseInfo, parsed) {
      if (typeof run_GenerateS2005DraftBySubmissionId !== 'function') return {};
      try {
        if (!caseInfo.folderId && typeof ensureCaseFolderId_ === 'function') {
          caseInfo.folderId = ensureCaseFolderId_(caseInfo);
        }
      } catch (_) {}
      const submissionId =
        (parsed && parsed.meta && (parsed.meta.submission_id || parsed.meta.submissionId)) || '';
      if (!submissionId) return {};
      const caseId = caseInfo.caseId || caseInfo.case_id || '';
      const draft = run_GenerateS2005DraftBySubmissionId(caseId, String(submissionId));
      const patch = {};
      if (draft && draft.url) patch.last_draft_url = draft.url;
      return patch;
    },
  },
  s2006_creditors_public: {
    name: 'S2006 債権者一覧表（公租公課用）',
    parser: function (subject, body) {
      if (typeof parseMetaBlock_ !== 'function') {
        throw new Error('parseMetaBlock_ is not defined');
      }
      const meta = parseMetaBlock_(body) || {};
      meta.form_key = meta.form_key || 's2006_creditors_public';
      if (!meta.submission_id) {
        const m = subject.match(/submission_id:(\d+)/) || subject.match(/提出\s+(\d{6,})/);
        if (m) meta.submission_id = m[1];
      }
      const fields = (typeof parseFieldsBlock_ === 'function') ? parseFieldsBlock_(body) : [];
      const model =
        typeof mapS2006FieldsToModel_ === 'function' ? mapS2006FieldsToModel_(fields) : {};
      return { meta, fieldsRaw: fields, model: model };
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    statusAfterSave: 'draft',
    requireCaseId: true,
    afterSave: function (caseInfo, parsed) {
      if (typeof run_GenerateS2006DraftBySubmissionId !== 'function') return {};
      try {
        if (!caseInfo.folderId && typeof ensureCaseFolderId_ === 'function') {
          caseInfo.folderId = ensureCaseFolderId_(caseInfo);
        }
      } catch (_) {}
      const submissionId =
        (parsed && parsed.meta && (parsed.meta.submission_id || parsed.meta.submissionId)) || '';
      if (!submissionId) return {};
      const caseId = caseInfo.caseId || caseInfo.case_id || '';
      const draft = run_GenerateS2006DraftBySubmissionId(caseId, String(submissionId));
      const patch = {};
      if (draft && draft.url) patch.last_draft_url = draft.url;
      return patch;
    },
  },
  s2010_p2_cause: {
    name: 'S2010 Part2(申立てに至った事情)',
    parser: function (subject, body) {
      if (typeof parseFormMail_S2010_P2_ !== 'function') {
        throw new Error('parseFormMail_S2010_P2_ is not defined');
      }
      return parseFormMail_S2010_P2_(subject, body);
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    statusAfterSave: 'intake',
    requireCaseId: true,
    afterSave: function (caseInfo, parsed) {
      const caseId = caseInfo.caseId || caseInfo.case_id || (parsed && parsed.meta && parsed.meta.case_id) || '';
      const folderId = caseInfo.folderId || caseInfo.folder_id || '';
      if (!caseId || !folderId) return {};
      if (typeof haveAllPartsS2010_ !== 'function' || typeof run_GenerateS2010DraftMergedByCaseId !== 'function') {
        return {};
      }
      const prefixes =
        typeof S2010_PART_PREFIXES !== 'undefined' && S2010_PART_PREFIXES
          ? S2010_PART_PREFIXES
          : ['s2010_p1_', 's2010_p2_'];
      if (haveAllPartsS2010_(folderId, prefixes)) {
        run_GenerateS2010DraftMergedByCaseId(caseId);
      }
      return {};
    },
  },
  s2010_userform: {
    name: 'S2010 申立（統合）',
    parser: function (subject, body) {
      if (typeof parseMetaBlock_ !== 'function') {
        throw new Error('parseMetaBlock_ is not defined');
      }
      const meta = parseMetaBlock_(body) || {};
      meta.form_key = meta.form_key || 's2010_userform';
      if (!meta.submission_id) {
        const m = subject.match(/submission_id:(\d+)/) || subject.match(/提出\s+(\d{6,})/);
        if (m) meta.submission_id = m[1];
      }
      const fields = (typeof parseFieldsBlock_ === 'function') ? parseFieldsBlock_(body) : [];
      return { meta, fieldsRaw: fields, model: {} };
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    statusAfterSave: 'draft',
    requireCaseId: true,
    afterSave: function (caseInfo) {
      if (typeof run_GenerateS2010DraftByCaseId !== 'function') return {};
      const caseId = caseInfo.caseId || caseInfo.case_id || '';
      if (!caseId) return {};
      const draft = run_GenerateS2010DraftByCaseId(caseId);
      const patch = {};
      if (draft && draft.draftUrl) patch.last_draft_url = draft.draftUrl;
      return patch;
    },
  },
  s2011_income_m1: {
    name: 'S2011 家計収支（1か月目）',
    parser: function (subject, body) {
      if (typeof parseMetaBlock_ !== 'function') {
        throw new Error('parseMetaBlock_ is not defined');
      }
      const meta = parseMetaBlock_(body) || {};
      meta.form_key = meta.form_key || 's2011_income_m1';
      if (!meta.submission_id) {
        const m = subject.match(/submission_id:(\d+)/) || subject.match(/提出\s+(\d{6,})/);
        if (m) meta.submission_id = m[1];
      }
      const fields = (typeof parseFieldsBlock_ === 'function') ? parseFieldsBlock_(body) : [];
      const model =
        typeof mapS2011_TableDriven_ === 'function' ? mapS2011_TableDriven_(fields, meta) : {};
      return { meta, fieldsRaw: fields, model: model };
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    statusAfterSave: 'draft',
    requireCaseId: true,
    afterSave: function (caseInfo, parsed) {
      if (typeof run_GenerateS2011DraftBySubmissionId !== 'function') return {};
      const submissionId =
        (parsed && parsed.meta && (parsed.meta.submission_id || parsed.meta.submissionId)) || '';
      if (!submissionId) return {};
      const caseId = caseInfo.caseId || caseInfo.case_id || '';
      const draft = run_GenerateS2011DraftBySubmissionId(caseId, String(submissionId), 's2011_income_m1');
      const patch = {};
      if (draft && draft.url) patch.last_draft_url = draft.url;
      return patch;
    },
  },
  s2011_income_m2: {
    name: 'S2011 家計収支（2か月目）',
    parser: function (subject, body) {
      if (typeof parseMetaBlock_ !== 'function') {
        throw new Error('parseMetaBlock_ is not defined');
      }
      const meta = parseMetaBlock_(body) || {};
      meta.form_key = meta.form_key || 's2011_income_m2';
      if (!meta.submission_id) {
        const m = subject.match(/submission_id:(\d+)/) || subject.match(/提出\s+(\d{6,})/);
        if (m) meta.submission_id = m[1];
      }
      const fields = (typeof parseFieldsBlock_ === 'function') ? parseFieldsBlock_(body) : [];
      const model =
        typeof mapS2011_TableDriven_ === 'function' ? mapS2011_TableDriven_(fields, meta) : {};
      return { meta, fieldsRaw: fields, model: model };
    },
    caseResolver: resolveCaseByCaseIdSmart_,
    statusAfterSave: 'draft',
    requireCaseId: true,
    afterSave: function (caseInfo, parsed) {
      if (typeof run_GenerateS2011DraftBySubmissionId !== 'function') return {};
      const submissionId =
        (parsed && parsed.meta && (parsed.meta.submission_id || parsed.meta.submissionId)) || '';
      if (!submissionId) return {};
      const caseId = caseInfo.caseId || caseInfo.case_id || '';
      const draft = run_GenerateS2011DraftBySubmissionId(caseId, String(submissionId), 's2011_income_m2');
      const patch = {};
      if (draft && draft.url) patch.last_draft_url = draft.url;
      return patch;
    },
  },
};

const FORM_INTAKE_QUEUE_LABELS = Object.freeze([]); // フォーム別Queueラベルは廃止

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
    var allowLast = String(sp.getProperty('ALLOW_LAST_LINE_ID_FALLBACK') || '').trim().toLowerCase();
    if (allowLast === '1' || allowLast === 'true') {
      var lid5 = String(sp.getProperty('LAST_LINE_ID') || '').trim();
      if (lid5) return lid5;
    }
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

function formIntake_normalizeFormKey_(rawKey, subject, meta) {
  var key = String(rawKey || '').trim();
  if (key && FORM_INTAKE_REGISTRY[key]) return key;
  var low = key.toLowerCase();
  if (low === 'intake_form' || low === 'intake') return 'intake';
  if (low === 'supporting_documents' || low === 'doc_payslip') return 'doc_payslip';
  if (/^s2010_p1/.test(low)) return 's2010_p1_career';
  if (/^s2010_p2/.test(low)) return 's2010_p2_cause';
  if (/^s2010/.test(low)) return 's2010_userform';
  if (/^s2011_income_m1/.test(low)) return 's2011_income_m1';
  if (/^s2011_income_m2/.test(low)) return 's2011_income_m2';
  if (key && FORM_INTAKE_REGISTRY[key]) return key;

  var subj = String(subject || '');
  var subjLow = subj.toLowerCase();
  var formName = String((meta && meta.form_name) || '').trim();
  if (/s2006/.test(subjLow) || /公租公課/.test(subj) || /債権者一覧表（公租公課用）/.test(formName)) {
    return 's2006_creditors_public';
  }
  if (/s2005/.test(subjLow) || (/債権者一覧表/.test(subj) && !/公租公課/.test(subj))) {
    return 's2005_creditors';
  }
  if (/s2010/.test(subjLow) || /S2010/.test(formName)) {
    if (/part\s*1/i.test(subj) || /経歴/.test(subj) || /経歴/.test(formName)) return 's2010_p1_career';
    if (/part\s*2/i.test(subj) || /事情/.test(subj) || /事情/.test(formName)) return 's2010_p2_cause';
    return 's2010_userform';
  }
  if (/s2011/.test(subjLow) || /S2011/.test(formName)) {
    if (/m1|1か月|１か月|1ヶ月/.test(subj) || /1か月|１か月|1ヶ月/.test(formName)) return 's2011_income_m1';
    if (/m2|2か月|２か月|2ヶ月/.test(subj) || /2か月|２か月|2ヶ月/.test(formName)) return 's2011_income_m2';
    return '';
  }
  if (/書類提出/.test(subj) || /書類提出/.test(formName) || /給与明細/.test(subj) || /給与明細/.test(formName)) {
    return 'doc_payslip';
  }
  if (/初回受付/.test(subj) || /初回受付/.test(formName) || /受付フォーム/.test(formName)) {
    return 'intake';
  }
  return '';
}

function formIntake_normalizeMetaAliases_(meta) {
  if (!meta || typeof meta !== 'object') return meta || {};
  var map = [
    { dst: 'line_id', srcs: ['line_id[0]', 'line_id[]', 'lineId', 'lineId[0]'] },
    { dst: 'case_id', srcs: ['case_id[0]', 'case_id[]', 'caseId', 'caseId[0]'] },
    { dst: 'user_key', srcs: ['user_key[0]', 'user_key[]', 'userKey', 'userKey[0]'] },
    { dst: 'form_key', srcs: ['form_key[0]', 'form_key[]', 'formKey', 'formKey[0]'] },
    { dst: 'submission_id', srcs: ['submission_id[0]', 'submission_id[]', 'submissionId', 'submissionId[0]'] },
  ];
  for (var i = 0; i < map.length; i++) {
    var item = map[i];
    if (meta[item.dst]) continue;
    for (var j = 0; j < item.srcs.length; j++) {
      var key = item.srcs[j];
      if (meta[key]) {
        meta[item.dst] = meta[key];
        break;
      }
    }
  }
  return meta;
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
  let email = String(metaObj.email || metaObj.mail || '').trim();
  if (!email && parsed && parsed.model) {
    email =
      String(
        parsed.model.email ||
          (parsed.model.applicant && parsed.model.applicant.email) ||
          ''
      ).trim();
  }
  if (email && !metaObj.email) metaObj.email = email;
  let caseId = formIntake_normalizeCaseId_(metaObj.case_id || metaObj.caseId);
  let userKey = formIntake_normalizeUserKey_(metaObj.user_key || metaObj.userKey);
  let lineId = String(metaObj.line_id || metaObj.lineId || '').trim();
  if (!userKey && email && typeof lookupUserKeyByEmail_ === 'function') {
    userKey = formIntake_normalizeUserKey_(lookupUserKeyByEmail_(email));
  }
  if (!lineId && typeof lookupLineIdByUserKey_ === 'function' && userKey) {
    lineId = String(lookupLineIdByUserKey_(userKey) || '').trim();
  }
  if (!caseId && metaObj.case_key) {
    const m = String(metaObj.case_key || '').match(/-(\d{4})$/);
    if (m) caseId = formIntake_normalizeCaseId_(m[1]);
  }
  if (!caseId && email && typeof fi_contactsLookupByEmail_ === 'function') {
    const hit = fi_contactsLookupByEmail_(email);
    if (hit) {
      if (!caseId && hit.active_case_id) caseId = formIntake_normalizeCaseId_(hit.active_case_id);
      if (!lineId && hit.line_id) lineId = String(hit.line_id || '').trim();
      if (!userKey && hit.user_key) userKey = formIntake_normalizeUserKey_(hit.user_key);
    }
  }
  if (!caseId && lineId && typeof lookupCaseIdByLineId_ === 'function') {
    caseId = formIntake_normalizeCaseId_(lookupCaseIdByLineId_(lineId));
  }
  // case_id 未確定時の採番可否をフォーム定義に従って判定
  const fkLow = String(metaObj.form_key || metaObj.formKey || '').trim().toLowerCase();
  const allowIssueNewCase = !(def && def.allowIssueNewCase === false);
  if (!caseId && (fkLow === 'intake' || !allowIssueNewCase)) {
    return {
      caseInfo: {},
      caseId: '',
      userKey: userKey || '',
      lineId: lineId || '',
      needsStaging: true,
    };
  }
  if (!caseId) {
    if (def && def.requireCaseId) {
      throw new Error('case_id is required for form_key=' + String(metaObj.form_key || ''));
    }
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

// saveSubmissionJson_ が無い/失敗する場合のフォールバック保存
function formIntake_saveSubmissionJsonFallback_(folderId, parsed, formKey, submissionId) {
  const folder = DriveApp.getFolderById(folderId);
  const meta = (parsed && parsed.meta) || {};
  const fk =
    String(formKey || meta.form_key || meta.formKey || '').trim() ||
    String(meta.form_name || meta.formName || '').trim() ||
    'intake';
  const sid =
    String(submissionId || meta.submission_id || meta.submissionId || '').trim() ||
    (typeof Utilities !== 'undefined' && typeof Utilities.getUuid === 'function'
      ? Utilities.getUuid()
      : String(Date.now()));
  const fileName = `${fk}__${sid}.json`;

  const existing = folder.getFilesByName(fileName);
  if (existing.hasNext()) return existing.next();

  const blob = Utilities.newBlob(JSON.stringify(parsed || {}, null, 2), 'application/json', fileName);
  return folder.createFile(blob);
}

function run_ProcessInbox_AllForms(opts) {
  opts = opts || {};
  const skipLock = !!opts.skipLock;
  const scriptLock = LockService.getScriptLock();
  let lockAcquired = false;
  if (!skipLock) {
    if (!scriptLock.tryLock(30000)) return;
    lockAcquired = true;
  }
  try {
    try { Logger.log('[router] start caller=%s', opts.caller || 'unknown'); } catch (_) {}
    const lockLabel = null;
    const processedLabelDefault = formIntake_labelOrCreate_(FORM_INTAKE_LABEL_PROCESSED);
    const toProcessLabel = formIntake_labelOrCreate_(FORM_INTAKE_LABEL_TO_PROCESS);
    const errorLabel = formIntake_labelOrCreate_(FORM_INTAKE_LABEL_ERROR);
    const noMetaLabel = formIntake_labelOrCreate_(FORM_INTAKE_LABEL_NO_META);
    const rejectedLabel = formIntake_labelOrCreate_(FORM_INTAKE_LABEL_REJECTED);

    const query = `label:${FORM_INTAKE_LABEL_TO_PROCESS}`;
    while (true) {
      const threads = GmailApp.search(query, 0, 50);
      if (!threads.length) break;
      threads.forEach(function (thread) {
        const messages = thread.getMessages();
        if (!messages || !messages.length) return;

        const msg = messages[messages.length - 1];
        let def = null;
        try {
          const body = msg.getPlainBody() || _htmlToText_(msg.getBody());
          const subject = msg.getSubject();
          let metaHint = {};
          try {
            metaHint = formIntake_normalizeMetaAliases_(parseMetaBlock_(body) || {});
          } catch (_) {
            metaHint = {};
          }
          const initialKeyRaw = String(metaHint.form_key || metaHint.formKey || '').trim();
          const initialKey = formIntake_normalizeFormKey_(initialKeyRaw, subject, metaHint);
          if (!initialKey) {
            formIntake_markFailed_(thread, lockLabel, toProcessLabel, noMetaLabel);
            return;
          }
          def = FORM_INTAKE_REGISTRY[initialKey];
          if (!def) {
            formIntake_markFailed_(thread, lockLabel, toProcessLabel, noMetaLabel);
            return;
          }

          let parsed = def.parser(subject, body);
          let meta = formIntake_normalizeMetaAliases_(parsed?.meta || {});
          parsed.meta = meta;
          let actualKeyRaw = String(meta.form_key || meta.formKey || '').trim();
          let actualKey = formIntake_normalizeFormKey_(actualKeyRaw, subject, meta);
          if (actualKey && actualKey !== actualKeyRaw) {
            meta.form_key = actualKey;
            parsed.meta = meta;
          }
          if (actualKey && actualKey !== initialKey) {
            const actualDef = FORM_INTAKE_REGISTRY[actualKey];
            if (!actualDef) {
              formIntake_markFailed_(thread, lockLabel, toProcessLabel, noMetaLabel);
              return;
            }
            if (actualDef !== def) {
              def = actualDef;
              parsed = def.parser(subject, body);
              meta = parsed?.meta || meta;
              actualKeyRaw = String(meta.form_key || meta.formKey || '').trim();
              actualKey = formIntake_normalizeFormKey_(actualKeyRaw, subject, meta);
              if (actualKey && actualKey !== actualKeyRaw) {
                meta.form_key = actualKey;
                parsed.meta = meta;
              }
            }
          }
          if (!actualKey) {
            formIntake_markFailed_(thread, lockLabel, toProcessLabel, noMetaLabel);
            return;
          }

          // 通知secretガード（NOTIFY_SECRET または ENFORCE_EMAIL_GUARD=1 のときだけ）
          try {
            const P = PropertiesService.getScriptProperties();
            const EXPECT = String(P.getProperty('NOTIFY_SECRET') || '').trim();
            const ENFORCE = (P.getProperty('ENFORCE_EMAIL_GUARD') || '').trim() === '1';
            if (ENFORCE || EXPECT) {
              if (!EXPECT) {
                try { Logger.log('[Guard] NOTIFY_SECRET missing; reject thread=%s', thread.getId && thread.getId()); } catch (_) {}
                formIntake_markFailed_(thread, lockLabel, toProcessLabel, rejectedLabel);
                return;
              }
              const provided = String((meta && meta.secret) || '').trim();
              if (provided !== EXPECT) {
                try { Logger.log('[Guard] secret mismatch'); } catch (_) {}
                formIntake_markFailed_(thread, lockLabel, toProcessLabel, rejectedLabel);
                return;
              }
            }
          } catch (e) {
            try { Logger.log('[Guard] error: %s', (e && e.stack) || e); } catch (_) {}
          }

          const queueLabel = toProcessLabel;
          const processedLabel = def.processedLabel
            ? formIntake_labelOrCreate_(def.processedLabel)
            : processedLabelDefault;

          const formKeyForStatusEarly = String(meta.form_key || actualKey || '').trim();
          if (formKeyForStatusEarly === 'intake') {
            try {
              const P = PropertiesService.getScriptProperties();
              const EXPECT = String(P.getProperty('NOTIFY_SECRET') || '').trim();
              const ALLOW_NO_SECRET = (P.getProperty('ALLOW_NO_SECRET') || '').toLowerCase() === '1';
              const provided = String((meta && meta.secret) || '').trim();
              if (EXPECT && !ALLOW_NO_SECRET && provided !== EXPECT) {
                try { Logger.log('[Intake] secret mismatch: meta.secret=%s', provided || '(empty)'); } catch (_) {}
                formIntake_markFailed_(thread, lockLabel, toProcessLabel, rejectedLabel);
                return;
              }
            } catch (_) {}
          }

          // doc_ 系フォームの添付保存保険（AttachmentsSaved ラベルが無い場合のみ）
          try {
            if (/^doc_/i.test(actualKey)) {
              const labels = thread.getLabels().map(function (l) { return l.getName(); });
              if (labels.indexOf(FORM_INTAKE_LABEL_ATTACH_SAVED) < 0) {
                if (typeof parseMetaAndFields === 'function' && typeof saveAttachmentsAndJson === 'function') {
                  const parsedAttach = parseMetaAndFields(msg);
                  if (!parsedAttach.form_name) parsedAttach.form_name = String(meta.form_name || '');
                  saveAttachmentsAndJson(parsedAttach, msg, { skipJson: true });
                  const attachLabel = formIntake_labelOrCreate_(FORM_INTAKE_LABEL_ATTACH_SAVED);
                  if (attachLabel) thread.addLabel(attachLabel);
                }
              }
            }
          } catch (e) {
            try { Logger.log('[doc_attach] error: %s', (e && e.stack) || e); } catch (_) {}
          }

        const prepared = formIntake_prepareCaseInfo_(meta, def, parsed);
        const caseInfo = prepared.caseInfo || {};
        const caseId = prepared.caseId;
        const case_id = caseId;
        caseInfo.caseId = caseInfo.caseId || caseId;
        if (!caseInfo.case_id && caseId) caseInfo.case_id = caseInfo.caseId;

        // intake（meta.case_id 無し）への採番結果を同じ参照に反映
        if (!meta.case_id && caseId) meta.case_id = caseId;
        if (!meta.caseId && caseId) meta.caseId = caseId;
        if (!meta.submission_id) {
          meta.submission_id =
            meta.submissionId ||
            (typeof Utilities !== 'undefined' && typeof Utilities.getUuid === 'function'
              ? Utilities.getUuid()
              : String(Date.now()));
        }
        if (!meta.submissionId) meta.submissionId = meta.submission_id;
        parsed.meta = meta;

        // case 未確定は staging へ保存して終了（case フォルダは作らない）
        if (prepared.needsStaging) {
          try {
            if (typeof stageIntakeMail_ === 'function') {
              const staged = stageIntakeMail_(thread, msg, parsed, meta, body) || {};
              if (staged.rejected) {
                formIntake_markFailed_(thread, lockLabel, toProcessLabel, rejectedLabel);
                return;
              }
              if (staged.quarantined) {
                formIntake_markFailed_(thread, lockLabel, toProcessLabel, noMetaLabel);
                return;
              }
              formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, true);
              return;
            }
            const stagingFolder =
              typeof drive_getOrCreateEmailStagingFolder_ === 'function'
                ? drive_getOrCreateEmailStagingFolder_()
                : drive_getRootFolder_();
            const stId = stagingFolder ? stagingFolder.getId() : drive_getRootFolder_().getId();
            let savedFile = null;
            try {
              if (typeof saveSubmissionJson_ === 'function') {
                savedFile = saveSubmissionJson_(stId, parsed);
              } else {
                savedFile = formIntake_saveSubmissionJsonFallback_(stId, parsed, actualKey, meta.submission_id);
              }
            } catch (saveErr) {
              try {
                Logger.log('[Intake] staging save failed: %s', (saveErr && saveErr.stack) || saveErr);
              } catch (_) {}
              savedFile = formIntake_saveSubmissionJsonFallback_(stId, parsed, actualKey, meta.submission_id);
            }
            try {
              Logger.log(
                '[Intake] staged to %s (%s)',
                stagingFolder ? stagingFolder.getName() : '_email_staging',
                savedFile && savedFile.getName && savedFile.getName()
              );
            } catch (_) {}
            formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, true);
            return;
          } catch (_) {
            formIntake_markFailed_(thread, lockLabel, toProcessLabel, noMetaLabel);
            return;
          }
        }

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
          try {
            resolvedCaseKey = drive_resolveCaseKeyFromMeta_(
              parsed?.meta || parsed?.META || {},
              fallbackInfo
            );
          } catch (err) {
            try {
              Logger.log(
                '[Intake] case_key resolve failed: %s',
                (err && err.message) || err
              );
            } catch (_) {}
            resolvedCaseKey = '';
          }
        }
        if (!resolvedCaseKey && fallbackInfo.userKey) {
          resolvedCaseKey = `${fallbackInfo.userKey}-${caseId}`;
        }

        // 2) 妥当性ガード（^[a-z0-9]{2,6}-\d{4}$ 以外は再生成）
        function isValidCaseKey_(s) { return /^[a-z0-9]{2,6}-\d{4}$/.test(String(s||'')); }
        let finalCaseKey = resolvedCaseKey;
        if (!isValidCaseKey_(finalCaseKey)) {
          const ukFix = (caseInfo.userKey || prepared.userKey || fallbackInfo.userKey || '').toString().slice(0,6).toLowerCase();
          const cidFix = (typeof normCaseId_ === 'function') ? normCaseId_(caseId) : String(caseId||'').replace(/\D/g,'').padStart(4,'0');
          if (ukFix && cidFix) finalCaseKey = `${ukFix}-${cidFix}`;
        }
        if (!isValidCaseKey_(finalCaseKey)) {
          const cidFix = (typeof normCaseId_ === 'function') ? normCaseId_(caseId) : String(caseId||'').replace(/\D/g,'').padStart(4,'0');
          // intake はここに来ない想定。フォールバック採番はしない。
          if (cidFix && formKeyForStatusEarly === 'intake') {
            const lid = caseInfo.lineId || prepared.lineId || fallbackInfo.lineId || fallbackInfo.line_id || '';
            let ukFix = lid ? formIntake_normalizeUserKey_(lid) : '';
            if (!ukFix && fallbackInfo.userKey) ukFix = formIntake_normalizeUserKey_(fallbackInfo.userKey);
            finalCaseKey = ukFix && cidFix ? `${ukFix}-${cidFix}` : '';
          }
        }
  if (!isValidCaseKey_(finalCaseKey)) throw new Error('Unable to resolve case folder key');

  // 3) 以降は finalCaseKey を唯一の真実に
  const resolved_case_key = finalCaseKey;

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

        const caseFolder = drive_getOrCreateCaseFolderByKey_(finalCaseKey);
        const caseFolderId = caseFolder.getId();
        caseInfo.folderId = caseFolderId;
        caseInfo.caseKey = finalCaseKey;
        caseInfo.case_key = finalCaseKey;
        const effectiveUserKey =
          caseInfo.userKey ||
          fallbackInfo.userKey ||
          (finalCaseKey && finalCaseKey.indexOf('-') >= 0
            ? finalCaseKey.split('-')[0]
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
          if (!parsed.meta.case_key) parsed.meta.case_key = finalCaseKey;
          if (!parsed.meta.caseKey) parsed.meta.caseKey = finalCaseKey;
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
            finalCaseKey,
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

        if (formIntake_isDuplicateSubmission_(case_id, caseFolderId, actualKey, meta.submission_id)) {
          formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, true);
          return;
        }

        let savedFile = null;
        try {
          if (typeof saveSubmissionJson_ === 'function') {
            savedFile = saveSubmissionJson_(caseFolderId, parsed);
          } else {
            savedFile = formIntake_saveSubmissionJsonFallback_(caseFolderId, parsed, actualKey, meta.submission_id);
          }
        } catch (saveErr) {
          try {
            Logger.log('[Intake] saveSubmissionJson_ failed: %s', (saveErr && saveErr.stack) || saveErr);
          } catch (_) {}
          savedFile = formIntake_saveSubmissionJsonFallback_(caseFolderId, parsed, actualKey, meta.submission_id);
        }
        let placedFile = null;
        try {
          placedFile = drive_placeFileIntoCase_(
            savedFile,
            parsed?.meta || parsed?.META || {},
            {
              caseId: caseId,
              case_id: case_id,
              caseKey: finalCaseKey,
              case_key: finalCaseKey,
              userKey: caseInfo.userKey || caseInfo.user_key,
              user_key: caseInfo.userKey || caseInfo.user_key,
              lineId: caseInfo.lineId,
              line_id: caseInfo.lineId,
            }
          );
        } catch (placeErr) {
          Logger.log('[Intake] placeFile error: %s', (placeErr && placeErr.stack) || placeErr);
        }
        const file_path = formIntake_buildFilePath_(placedFile || savedFile, caseFolderId);
        Logger.log('[Intake] saved %s', file_path);

        if (typeof recordSubmission_ === 'function') {
          try {
            recordSubmission_({
              case_id: case_id,
              form_key: actualKey,
              submission_id: meta.submission_id || '',
              json_path: file_path,
              meta,
              case_key: finalCaseKey,
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
          basePatch.case_key = finalCaseKey;
          basePatch.folder_id = caseFolderId;
          if (caseInfo.userKey || caseInfo.user_key) {
            basePatch.user_key = caseInfo.userKey || caseInfo.user_key;
          }
          updateCasesRow_(case_id, basePatch);
        }

        try {
          const attachSaved = formIntake_labelOrCreate_(FORM_INTAKE_LABEL_ATTACH_SAVED);
          if (attachSaved) thread.removeLabel(attachSaved);
        } catch (_) {}
        try {
          const errLabel = formIntake_labelOrCreate_(FORM_INTAKE_LABEL_ERROR);
          if (errLabel) thread.removeLabel(errLabel);
        } catch (_) {}
        formIntake_cleanupLabels_(thread, queueLabel, lockLabel, processedLabel, toProcessLabel, true);
      } catch (err) {
        const defName = def && def.name ? def.name : 'unknown';
        let errText = '';
        try {
          errText = String((err && err.stack) || err || '');
        } catch (_) {
          errText = 'unknown error';
        }
        let htmlBody = '';
        try {
          const safe = typeof safeHtml === 'function' ? safeHtml(errText) : errText;
          htmlBody = `<pre>${safe}</pre>`;
        } catch (_) {
          htmlBody = `<pre>${errText}</pre>`;
        }
        try {
          const draftTo = (() => {
            try {
              return msg.getFrom();
            } catch (_) {
              return 'me';
            }
          })();
          GmailApp.createDraft(draftTo, `[BAS Intake Error] ${defName}`, errText, { htmlBody });
        } catch (_) {}
        try {
          formIntake_markFailed_(thread, lockLabel, toProcessLabel, errorLabel);
        } catch (_) {}
      }
    });
    }
  } finally {
    if (lockAcquired) {
      try {
        scriptLock.releaseLock();
      } catch (_) {}
    }
  }
}

function formIntake_assignQueueLabels_() {
  // フォーム別Queueラベルは使用しない
  return;
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
    if (processedLabel) thread.addLabel(processedLabel);
    try {
      if (queueLabel) thread.removeLabel(queueLabel);
    } catch (_) {}
    try {
      if (toProcessLabel && toProcessLabel !== queueLabel) thread.removeLabel(toProcessLabel);
    } catch (_) {}
  }
  try {
    if (lockLabel) thread.removeLabel(lockLabel);
  } catch (_) {}
}

function formIntake_markFailed_(thread, lockLabel, toProcessLabel, failedLabel) {
  try {
    if (failedLabel) thread.addLabel(failedLabel);
  } catch (_) {}
  try {
    if (toProcessLabel) thread.removeLabel(toProcessLabel);
  } catch (_) {}
  try {
    if (lockLabel) thread.removeLabel(lockLabel);
  } catch (_) {}
}

function formIntake_buildFilePath_(file, fallbackFolderId) {
  if (!file) return '';
  let folderId = String(fallbackFolderId || '');
  try {
    const parents = file.getParents();
    if (parents && parents.hasNext()) {
      folderId = parents.next().getId();
    }
  } catch (_) {}
  let name = '';
  try {
    name = file.getName();
  } catch (_) {}
  if (folderId && name) return `${folderId}/${name}`;
  if (name) return name;
  return folderId;
}

function formIntake_isDuplicateSubmission_(caseId, folderId, formKey, submissionId) {
  if (!formKey || !submissionId) return false;
  try {
    const normCaseId =
      typeof normCaseId_ === 'function'
        ? normCaseId_(caseId)
        : String(caseId || '').replace(/\D/g, '').padStart(4, '0');
    if (normCaseId && typeof sheetsRepo_hasSubmission_ === 'function') {
      if (sheetsRepo_hasSubmission_(normCaseId, formKey, submissionId)) return true;
    }
  } catch (e) {
    try { Logger.log('[Intake] duplicate check (sheet) error: %s', (e && e.stack) || e); } catch (_) {}
  }
  if (!folderId) return false;
  try {
    const folder = DriveApp.getFolderById(folderId);
    const fname = `${formKey}__${submissionId}.json`;
    return folder.getFilesByName(fname).hasNext();
  } catch (_) {
    return false;
  }
}

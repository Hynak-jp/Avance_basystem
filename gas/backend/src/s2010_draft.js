/** ====== 設定（S2010） ====== **
 * 必要な Script Properties:
 * - BAS_MASTER_SPREADSHEET_ID : cases台帳シートのID（S2002と共通）
 * - S2010_TEMPLATE_GDOC_ID    : S2010差し込み用gdocテンプレのファイルID
 */
const PROP_S2010 = PropertiesService.getScriptProperties();
const S2010_SPREADSHEET_ID = PROP_S2010.getProperty('BAS_MASTER_SPREADSHEET_ID') || '';
const S2010_TPL_GDOC_ID = PROP_S2010.getProperty('S2010_TEMPLATE_GDOC_ID') || '';
const S2010_LABEL_TO_PROCESS = 'FormAttach/ToProcess';
const S2010_LABEL_PROCESSED = 'FormAttach/Processed';
// S2010 の分割フォームが揃っているか判定するための form_key 接頭辞リスト
const S2010_PART_PREFIXES = ['s2010_p1_', 's2010_p2_', 's2010_p3_'];

// チェック記号はS2002と統一
const CHECKED = '☑';
const UNCHECKED = '□';
function renderCheck_(b) {
  return b ? CHECKED : UNCHECKED;
}

/** ====== パブリック・エントリ ====== **/

/**
 * 受信箱から S2010 通知メールを取り込み → JSON 保存 → 必須パートが揃えば統合 JSON を保存
 * ラベル: FormAttach/ToProcess → 処理後に FormAttach/Processed を付与
 */
function run_ProcessInbox_S2010() {
  const query = `label:${S2010_LABEL_TO_PROCESS} subject:#FM-BAS subject:S2010`;
  const threads = GmailApp.search(query, 0, 50);
  try {
    Logger.log('[S2010] tick: query=%s threads=%s', query, threads.length);
  } catch (_) {}
  if (!threads.length) return;

  threads.forEach((th) => {
    th.getMessages().forEach((msg) => {
      try {
        const body = msg.getPlainBody() || HtmlService.createHtmlOutput(msg.getBody()).getContent();
        const subject = msg.getSubject();
        const parsed = parseFormMail_(subject, body);
        const formKey = String(parsed.meta?.form_key || '').trim();
        if (!/^s2010_/.test(formKey)) return; // 関係ない通知は無視

        const caseInfo = resolveCaseByCaseId_(parsed.meta.case_id) || {};
        const fallbackInfo = {
          caseId: caseInfo.caseId || parsed.meta?.case_id,
          case_id: caseInfo.caseId || parsed.meta?.case_id,
          caseKey: caseInfo.caseKey,
          case_key: caseInfo.caseKey,
          userKey:
            (caseInfo.caseKey && String(caseInfo.caseKey).indexOf('-') >= 0
              ? String(caseInfo.caseKey).split('-')[0]
              : caseInfo.userKey) || '',
          user_key:
            (caseInfo.caseKey && String(caseInfo.caseKey).indexOf('-') >= 0
              ? String(caseInfo.caseKey).split('-')[0]
              : caseInfo.userKey) || '',
          lineId: caseInfo.lineId || parsed.meta?.line_id || parsed.meta?.lineId || '',
          line_id: caseInfo.lineId || parsed.meta?.line_id || parsed.meta?.lineId || '',
        };
        const resolvedCaseKey = drive_resolveCaseKeyFromMeta_(parsed.meta || {}, fallbackInfo);
        const caseFolder = drive_getOrCreateCaseFolderByKey_(resolvedCaseKey);
        const caseFolderId = caseFolder.getId();
        caseInfo.folderId = caseFolderId;
        caseInfo.caseKey = resolvedCaseKey;
        caseInfo.userKey = resolvedCaseKey.split('-')[0];
        caseInfo.user_key = caseInfo.userKey;
        if (!caseInfo.lineId && fallbackInfo.lineId) caseInfo.lineId = fallbackInfo.lineId;

        const savedFile = saveSubmissionJson_(caseFolderId, parsed);
        try {
          drive_placeFileIntoCase_(savedFile, parsed.meta || {}, {
            caseId: fallbackInfo.caseId,
            case_id: fallbackInfo.case_id,
            caseKey: resolvedCaseKey,
            case_key: resolvedCaseKey,
            userKey: caseInfo.userKey,
            user_key: caseInfo.userKey,
            lineId: caseInfo.lineId,
            line_id: caseInfo.lineId,
          });
        } catch (err) {
          Logger.log('[S2010] placeFile error: %s', (err && err.stack) || err);
        }

        if (typeof updateCasesRow_ === 'function') {
          const patch = {
            case_key: resolvedCaseKey,
            folder_id: caseFolderId,
            user_key: caseInfo.userKey,
            last_activity: new Date(),
          };
          updateCasesRow_(parsed.meta.case_id, patch);
        }

        if (haveAllPartsS2010_(caseInfo.folderId, S2010_PART_PREFIXES)) {
          run_GenerateS2010MergedJsonByCaseId(caseInfo.caseId);
        }

        msg.addLabel(GmailApp.getUserLabelByName(S2010_LABEL_PROCESSED));
        msg.removeLabel(GmailApp.getUserLabelByName(S2010_LABEL_TO_PROCESS));
      } catch (e) {
        GmailApp.createDraft('me', '[BAS Intake Error]', String(e), {
          htmlBody: `<pre>${safeHtml(e.stack || e)}</pre>`,
        });
      }
    });
  });
}

/** 直近の S2010 JSON を読み込んでドラフト生成（ケースID指定） */
function run_GenerateS2010DraftByCaseId(caseId) {
  const info = resolveCaseByCaseId_(caseId);
  if (!info) throw new Error(`Unknown case_id: ${caseId}`);
  info.folderId = ensureCaseFolderId_(info);

  const parsed = loadLatestFormJson_(info.folderId, 's2010_userform');
  if (!parsed) throw new Error(`No S2010 JSON found under case folder: ${info.folderId}`);

  const draft = generateS2010Draft_(info, parsed);
  updateCasesRow_(info.caseId || caseId, {
    status: 'draft',
    last_activity: new Date(),
    last_draft_url: draft.draftUrl,
  });
  try {
    Logger.log('[S2010] draft created: %s', draft.draftUrl);
  } catch (_) {}
  return draft;
}

/** s2010_p*_ をマージして統合JSONのみ保存（ドラフト生成はしない） */
function run_GenerateS2010MergedJsonByCaseId(caseId) {
  const info = resolveCaseByCaseId_(caseId);
  if (!info) throw new Error(`Unknown case_id: ${caseId}`);
  info.folderId = ensureCaseFolderId_(info);

  const parts = loadLatestPartsByPrefix_(info.folderId, 's2010_');
  if (!parts.length) throw new Error('No s2010_* json found under case folder.');

  if (!haveAllPartsS2010_(info.folderId, S2010_PART_PREFIXES)) {
    throw new Error('Not all required S2010 parts are present yet.');
  }

  const merged = mergeS2010Parts_(parts, { caseId: info.caseId });
  const fname = `s2010_userform__merged_${merged.meta.merged_at}.json`;
  DriveApp.getFolderById(info.folderId).createFile(
    Utilities.newBlob(JSON.stringify(merged, null, 2), 'application/json', fname)
  );

  try {
    Logger.log('[S2010] merged json saved: %s', fname);
  } catch (_) {}
  return merged;
}

/** デバッグ：'0001' などで呼び出し */
function debug_GenerateS2010_for_0001() {
  return run_GenerateS2010DraftByCaseId('0001');
}

/** ===== マルチフォーム統合 → ドラフト生成（入口） ===== */
function run_GenerateS2010DraftMergedByCaseId(caseId) {
  // 1) cases解決 & フォルダ解決
  const info = resolveCaseByCaseId_(caseId);
  if (!info) throw new Error(`Unknown case_id: ${caseId}`);
  info.folderId = ensureCaseFolderId_(info);

  // 2) ケース直下から最新の各partを収集
  const parts = loadLatestPartsByPrefix_(info.folderId, 's2010_'); // s2010_ で始まるform_keyを全部拾う
  if (!parts.length) throw new Error('No S2010 part-jsons found.');

  // 3) マージして統合JSONを作る
  const merged = mergeS2010Parts_(parts, { caseId: info.caseId }); // { meta, fieldsRaw, model }

  // 4) 統合JSONをケース直下に保存（監査用）
  const mergedName = `s2010_userform__merged_${merged.meta.merged_at}.json`;
  DriveApp.getFolderById(info.folderId)
    .createFile(Utilities.newBlob(JSON.stringify(merged, null, 2), 'application/json', mergedName));

  // 5) 既存のS2010生成器でドラフト化
  const draft = generateS2010Draft_(info, merged);

  // 6) ステータス更新
  updateCasesRow_(info.caseId || caseId, {
    status: 'draft',
    last_activity: new Date(),
    last_draft_url: draft.draftUrl,
  });

  try { Logger.log('[S2010] merged draft created: %s', draft.draftUrl); } catch (_){ }
  return draft;
}

/** ===== ケース直下から "prefix" に合う form_key の最新を集める ===== */
function loadLatestPartsByPrefix_(caseFolderId, prefix) {
  const id = normalizeFolderId_(caseFolderId);
  const folder = DriveApp.getFolderById(id);

  // form_key => {json, t}（最新のみ保持）
  const latestByKey = {};
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const name = f.getName();
    if (!/\.json$/i.test(name || '')) continue;
    let j; try { j = JSON.parse(f.getBlob().getDataAsString('UTF-8')); } catch (_){ continue; }

    const key = (j && j.meta && j.meta.form_key) || j.form_key || '';
    if (!key || !String(key).startsWith(prefix)) continue;        // s2010_... だけ
    if (key === 's2010_userform') continue;                       // 既存の統合JSONは除外

    const t = f.getLastUpdated().getTime();
    const cur = latestByKey[key];
    if (!cur || t > cur.t) latestByKey[key] = { json: j, t, fileId: f.getId() };
  }

  // 収集結果を配列に
  return Object.keys(latestByKey).map(k => ({ form_key: k, ...latestByKey[k] }));
}

/** ===== パーツ配列を 1つの "s2010_userform" JSON に統合 ===== */
function mergeS2010Parts_(parts, opt) {
  // 1) すべての fields を配列 [{label, value}] に正規化
  const arrays = [];
  const meta_list = [];
  parts.forEach(p => {
    const j = p.json || {};
    const arr = normalizeFieldsArrayForAny_(j); // 既存の正規化
    if (arr && arr.length) arrays.push(arr);
    meta_list.push({
      form_key: (j.meta && j.meta.form_key) || p.form_key,
      submission_id: j.meta && j.meta.submission_id,
      fileId: p.fileId
    });
  });

  // 2) ラベル単位でマージ（後勝ち or 非空優先）。ここは「後勝ち＋非空」でシンプルに。
  const mergedMap = new Map(); // label -> value
  arrays.forEach(arr => {
    arr.forEach(({label, value}) => {
      const v = String(value || '').trim();
      if (!v) return;                 // 空は無視（既存値を消さない）
      mergedMap.set(label, v);        // 同一ラベルは後から来た値で上書き
    });
  });

  // 3) fieldsRawを配列化（元の順序にこだわらない場合はMap→配列でOK）
  const fieldsRaw = Array.from(mergedMap.entries()).map(([label, value]) => ({ label, value }));

  // 4) S2010のモデリングへ
  const model = mapFieldsToModel_S2010_(fieldsRaw); // 既存のS2010マッパ

  // 5) 統合メタ
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const meta = {
    form_key: 's2010_userform',
    merged_at: now,
    merged_from: meta_list,
    case_id: (opt && opt.caseId) || ''
  };

  return { meta, fieldsRaw, model };
}

/** 必要な各パートが揃っているか確認する */
function haveAllPartsS2010_(caseFolderId, prefixes) {
  const got = new Set();
  const folder = DriveApp.getFolderById(caseFolderId);
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (!/\.json$/i.test(f.getName() || '')) continue;
    let j;
    try {
      j = JSON.parse(f.getBlob().getDataAsString('UTF-8'));
    } catch (_) {
      continue;
    }
    const key = (j.meta && j.meta.form_key) || j.form_key || '';
    prefixes.forEach((p) => {
      if (String(key).startsWith(p)) got.add(p);
    });
  }
  return prefixes.every((p) => got.has(p));
}

/** ====== ローダ（ケース直下 .json を form_key で選別） ====== **/

function loadLatestFormJson_(caseFolderId, wantFormKey) {
  const id = normalizeFolderId_(caseFolderId);
  if (!id) throw new Error('[S2010] cases.folderId is empty.');
  let parent;
  try {
    parent = DriveApp.getFolderById(id);
  } catch (e) {
    throw new Error('[S2010] invalid folderId: ' + id + ' :: ' + (e.message || e));
  }

  const candidates = [];
  const it = parent.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const n = f.getName && f.getName();
    if (!n || !/\.json$/i.test(n)) continue;
    try {
      const j = JSON.parse(f.getBlob().getDataAsString('UTF-8'));
      const formKey = (j && j.meta && j.meta.form_key) || j.form_key || '';
      if (String(formKey).trim() === String(wantFormKey)) {
        candidates.push({ file: f, json: j, t: f.getLastUpdated().getTime() });
      }
    } catch (_) {}
  }
  if (!candidates.length) return null;
  candidates.sort((a, b) => b.t - a.t);
  const latest = candidates[0].json || {};

  if (latest.model && latest.meta) return latest;

  const fieldsArr = normalizeFieldsArrayForAny_(latest);
  if (!fieldsArr) throw new Error('Invalid submission JSON shape (no fields/model)');
  return {
    meta: latest.meta || {},
    fieldsRaw: fieldsArr,
    model: mapFieldsToModel_S2010_(fieldsArr),
  };
}

function normalizeFolderId_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  const m = s.match(/[-\w]{25,}/);
  return m ? m[0] : s;
}

function normalizeFieldsArrayForAny_(json) {
  if (!json || typeof json !== 'object') return null;
  if (Array.isArray(json.fieldsRaw)) return json.fieldsRaw;
  const f = json.fields;
  if (Array.isArray(f)) return f;
  if (f && typeof f === 'object') {
    const arr = [];
    Object.keys(f).forEach((k) => {
      const label = /^【.*】$/.test(k) ? k : `【${k}】`;
      arr.push({ label, value: f[k] });
    });
    return arr;
  }
  return null;
}

/** ====== パーサ → S2010 モデル整形 ====== **/

/** FIELDS配列を S2010 用のモデルにマッピング。S2002の mapFieldsToModel_ とは独立に最小限だけ。 */
function mapFieldsToModel_S2010_(fields) {
  const out = { app: {}, addr: {}, ref: {}, stmt: {} };

  // 基本系（S2002同様のラベル拾い・必要最小限）
  fields.forEach(({ label, value }) => {
    if (/^【?名前（ふりがな）】?/.test(label)) {
      const { name, kana } = parseNameKanaValue_(value);
      out.app.name = normSpace_(name);
      out.app.kana = kana;
      return;
    }
    if (/^【?名前】?$/.test(label)) {
      out.app.name = normSpace_(value);
      return;
    }
    if (/^【?メールアドレス】?/.test(label)) {
      out.app.email = String(value || '').trim();
      return;
    }
    if (/^【?連絡先】?/.test(label)) {
      out.app.phone = normPhone_(value);
      return;
    }
    if (/^【?生年月日】?/.test(label)) {
      out.app.birth = String(value || '').trim();
      return;
    }
    if (/^【?住居所（住民票と異なる[^】]*）】?$/.test(label)) {
      out.addr.alt_full = normJaSpace_(value);
      return;
    }
    if (/^【?住居所】?/.test(label)) {
      out.addr.full = normJaSpace_(value);
      return;
    }
  });

  // ---- 配偶者・同居/別居の家族（最大4人） ----
  const HH_MAX = 4;
  out.hh = out.hh || [];

  const gv = (regexes) => {
    const hit = (fields || []).find((f) => regexes.some((re) => re.test(String(f.label || ''))));
    return hit ? String(hit.value || '').trim() : '';
  };
  const numLead = (i, key) =>
    new RegExp(`^【?\\s*${i}\\s*[：:）)]?\\s*${key}\\s*】?$`);
  const toCohab = (v) => {
    const s = String(v || '').trim();
    if (/^同(居)?$/.test(s)) return true;
    if (/^別(居)?$/.test(s)) return false;
    return null;
  };
  const toDigits = (v) => {
    const t = String(v || '').replace(/[^\d]/g, '');
    return t ? t : '';
  };

  for (let i = 1; i <= HH_MAX; i++) {
    const idx = i - 1;
    out.hh[idx] = {
      name: gv([numLead(i, '氏名')]),
      relation: gv([numLead(i, '続柄')]),
      age: gv([numLead(i, '年齢')]),
      occupation: gv([numLead(i, '職業・?学年')]),
      cohab: toCohab(gv([numLead(i, '同別'), numLead(i, '同居・別居')])),
      income: toDigits(gv([numLead(i, '平均月収\\(円\\)')])),
    };
  }

  // 後処理（年齢・和暦）
  if (out.app && out.app.birth) {
    const iso = toIsoBirth_(out.app.birth);
    out.app.birth_iso = iso;
    out.app.age = iso ? calcAge_(iso) : '';
    const w = toWareki_(iso);
    if (w) out.app.birth_wareki = w;
  }

  return out;
}

/** ここからS2010専用の意味付け（理由/きっかけ/日付/職歴1） */
function ensureS2010Model_(m) {
  const out = JSON.parse(JSON.stringify(m || {}));
  if (!out.app) out.app = {};
  if (!out.addr) out.addr = {};
  if (!out.ref) out.ref = {};
  if (!out.stmt) out.stmt = {};
  if (!out.hh) out.hh = [];

  // 自由記述（拾えるだけ）
  if (!out.stmt.free && Array.isArray((m && m.fieldsRaw) || [])) {
    const lines = m.fieldsRaw
      .filter((f) => /陳述|事情|経緯|理由|生活状況|支払不能|収支/i.test(f.label))
      .map((f) => String(f.value || '').trim())
      .filter(Boolean);
    if (lines.length) out.stmt.free = lines.join('\n\n');
  }

  // 理由（複数選択想定）
  const reasonsRaw = findFieldValue_(m, /理由は.*とおり|借金.*理由/);
  const reasons = splitMulti_(reasonsRaw);
  out.reason = {
    living: hasAny_(reasons, /生活費/),
    mortgage: hasAny_(reasons, /住宅ローン|住宅/),
    education: hasAny_(reasons, /教育/),
    waste: hasAny_(reasons, /浪費|飲食|飲酒|投資|投機|商品購入|ギャンブル/),
    business: hasAny_(reasons, /事業|経営破綻|マルチ|ネットワーク/),
    guarantee: hasAny_(reasons, /保証/),
    other: hasAny_(reasons, /その他/),
    other_text: pickOtherText_(reasonsRaw),
  };

  // きっかけ（複数選択想定）
  const triggersRaw = findFieldValue_(m, /きっかけ.*とおり|返済.*できなく.*きっかけ/);
  const triggers = splitMulti_(triggersRaw);
  out.trigger = {
    overpay: hasAny_(triggers, /収入以上|返済金額/),
    dismiss: hasAny_(triggers, /解雇/),
    paycut: hasAny_(triggers, /減額/),
    hospital: hasAny_(triggers, /病気|入院/),
    other: hasAny_(triggers, /その他/),
    other_text: pickOtherText_(triggersRaw),
  };

  // 支払不能の時期・約定返済合計・受任日（任意）
  const unable = findFieldValue_(m, /支払不能.*時期/);
  const unableYM = toYMP_(unable);
  out.unable_yyyy = unableYM.yyyy || '';
  out.unable_mm = unableYM.mm || '';
  const monthly = findFieldValue_(m, /約定返済額|月々.*返済額/);
  out.unable_monthly_total = toNumberText_(monthly);

  const notice = findFieldValue_(m, /受任通知.*発送日/);
  const nd = toYMD_(notice);
  out.notice_yyyy = nd.yyyy || '';
  out.notice_mm = nd.mm || '';
  out.notice_dd = nd.dd || '';

  // ===== 職歴1（開始は常に入力あり。終了/現在は入力無しでもOK。無職でも開始だけ出す） =====
  const companyRaw = findFieldValue_(m, /職歴1.*(就業先|会社名|勤務先)/);
  const kindRaw = findFieldValue_(m, /職歴1.*種\s*別|種\s*別/);
  const monthlyRaw = findFieldValue_(
    m,
    /職歴1.*平均月収(\s*[（(]円[）)]\s*)?|平均月収(\s*[（(]円[）)]\s*)?/
  );
  const severanceRaw = findFieldValue_(m, /職歴1.*退職金.*有|退職金.*有/);

  const kind = String(kindRaw || '');
  const kindFlags = {
    employee: /勤め|正社員|従業員/i.test(kind),
    part: /パート|ｱﾙﾊﾞｲﾄ|アルバイト|派遣|契約/i.test(kind),
    self: /自営|自営業|個人事業/i.test(kind),
    representative: /法人代表者|代表|社長|取締役/i.test(kind),
    unemployed: /無職|失業/i.test(kind),
  };

  // 開始は常に拾う（「職歴1：開始」「職歴1：就業開始」「無職開始」など許容）
  const startRaw = findFieldValue_(
    m,
    /職歴1[:：]?\s*(開始|就業開始|無職開始)|就業期間.*開始|職歴.*1.*開始/
  );
  const endRaw = findFieldValue_(m, /職歴1[:：]?\s*(就業終了|終了)|就業期間.*終了|職歴.*1.*終了/);
  const curRaw = findFieldValue_(m, /職歴1[:：]?\s*現\s*在|職歴1.*現在|現在就業中/);

  const fromYMD = toYMD_(startRaw); // {yyyy,mm,dd}
  const toYMDv = toYMD_(endRaw); // {yyyy,mm,dd}

  let isCurrent =
    /はい|現.?在/i.test(String(curRaw || '')) || !endRaw || /現.?在/.test(String(endRaw || ''));
  if (kindFlags.unemployed) {
    // 無職：現在フラグは立てない。終了は表示しない（空）。
    isCurrent = false;
    toYMDv.yyyy = '';
    toYMDv.mm = '';
    toYMDv.dd = '';
  }

  const pad2 = (s) => (s ? String(s).padStart(2, '0') : '');
  out.jobs1 = {
    company: normJaSpace_(companyRaw || ''),
    kind: kindFlags,
    avg_monthly: kindFlags.unemployed ? '' : toNumberText_(monthlyRaw),
    severance: toBoolJaLoose_(severanceRaw),

    from_yyyy: fromYMD.yyyy || '',
    from_mm: pad2(fromYMD.mm || ''),
    from_dd: pad2(fromYMD.dd || ''),

    end_yyyy: isCurrent ? '' : toYMDv.yyyy || '',
    end_mm: isCurrent ? '' : pad2(toYMDv.mm || ''),
    current: !!isCurrent,
  };
  out.jobs1.end_text = out.jobs1.current
    ? '現　在'
    : out.jobs1.end_yyyy && out.jobs1.end_mm
    ? `${out.jobs1.end_yyyy}年${out.jobs1.end_mm}月`
    : '';

  return out;
}

/** ====== ドラフト生成（S2010） ====== **/
function generateS2010Draft_(caseInfo, parsed) {
  if (!S2010_TPL_GDOC_ID) throw new Error('S2010_TEMPLATE_GDOC_ID not set');
  const drafts = getOrCreateSubfolder_(DriveApp.getFolderById(caseInfo.folderId), 'drafts');

  const subId =
    parsed.meta?.submission_id ||
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const draftName = `S2010_${caseInfo.caseId}_draft_${subId}`;
  const gdocId = DriveApp.getFileById(S2010_TPL_GDOC_ID).makeCopy(draftName, drafts).getId();

  const doc = DocumentApp.openById(gdocId);
  const body = doc.getBody();
  const M = ensureS2010Model_(parsed.model || {});

  // 基本情報（テンプレに置いてあれば出ます）
  replaceAll_(body, '{{app.name}}', M.app.name || '');
  replaceAll_(body, '{{app.kana}}', M.app.kana || '');
  if (M.app.birth_iso) {
    const [yy, mm, dd] = M.app.birth_iso.split('-');
    replaceAll_(body, '{{app.birth_yyyy}}', yy || '');
    replaceAll_(body, '{{app.birth_mm}}', String(mm || ''));
    replaceAll_(body, '{{app.birth_dd}}', String(dd || ''));
  }
  const w = M.app.birth_wareki || null;
  replaceAll_(body, '{{app.birth_wareki_gengo}}', (w && w.gengo) || '');
  replaceAll_(body, '{{app.birth_wareki_yy}}', (w && String(w.yy)) || '');
  replaceAll_(body, '{{app.age}}', (M.app.age ?? '') + '');
  replaceAll_(body, '{{addr.full}}', normJaSpace_(M.addr.full || ''));
  replaceAll_(body, '{{addr.alt_full}}', normJaSpace_(M.addr.alt_full || ''));
  replaceAll_(body, '{{stmt.free}}', (M.stmt.free || '').trim());

  for (let i = 1; i <= 4; i++) {
    const h = (M.hh && M.hh[i - 1]) || {};
    replaceAll_(body, `{{hh${i}.name}}`, h.name || '');
    replaceAll_(body, `{{hh${i}.relation}}`, h.relation || '');
    replaceAll_(body, `{{hh${i}.age}}`, h.age || '');
    replaceAll_(body, `{{hh${i}.occupation}}`, h.occupation || '');
    replaceAll_(body, `{{hh${i}.avg_monthly}}`, h.income || '');
    replaceAll_(body, `{{chk.hh${i}.cohab}}`, renderCheck_(h.cohab === true));
    replaceAll_(body, `{{chk.hh${i}.separate}}`, renderCheck_(h.cohab === false));
  }

  // 第2-1 理由
  replaceAll_(body, '{{chk.reason.living}}', renderCheck_(!!M.reason?.living));
  replaceAll_(body, '{{chk.reason.mortgage}}', renderCheck_(!!M.reason?.mortgage));
  replaceAll_(body, '{{chk.reason.education}}', renderCheck_(!!M.reason?.education));
  replaceAll_(body, '{{chk.reason.waste}}', renderCheck_(!!M.reason?.waste));
  replaceAll_(body, '{{chk.reason.business}}', renderCheck_(!!M.reason?.business));
  replaceAll_(body, '{{chk.reason.guarantee}}', renderCheck_(!!M.reason?.guarantee));
  replaceAll_(body, '{{chk.reason.other}}', renderCheck_(!!M.reason?.other));
  replaceAll_(body, '{{reason.other_text}}', M.reason?.other_text || '');

  // 第2-2 きっかけ
  replaceAll_(body, '{{chk.trigger.overpay}}', renderCheck_(!!M.trigger?.overpay));
  replaceAll_(body, '{{chk.trigger.dismiss}}', renderCheck_(!!M.trigger?.dismiss));
  replaceAll_(body, '{{chk.trigger.paycut}}', renderCheck_(!!M.trigger?.paycut));
  replaceAll_(body, '{{chk.trigger.hospital}}', renderCheck_(!!M.trigger?.hospital));
  replaceAll_(body, '{{chk.trigger.other}}', renderCheck_(!!M.trigger?.other));
  replaceAll_(body, '{{trigger.other_text}}', M.trigger?.other_text || '');

  // 支払不能時期・受任・金額
  replaceAll_(body, '{{unable_yyyy}}', M.unable_yyyy || '');
  replaceAll_(body, '{{unable_mm}}', M.unable_mm || '');
  replaceAll_(body, '{{unable_monthly_total}}', M.unable_monthly_total || '');
  replaceAll_(body, '{{notice_yyyy}}', M.notice_yyyy || '');
  replaceAll_(body, '{{notice_mm}}', M.notice_mm || '');
  replaceAll_(body, '{{notice_dd}}', M.notice_dd || '');

  // 第1-1 職歴1（上段1ブロック）
  replaceAll_(body, '{{jobs1.from_yyyy}}', M.jobs1?.from_yyyy || '');
  replaceAll_(body, '{{jobs1.from_mm}}', M.jobs1?.from_mm || '');
  replaceAll_(body, '{{jobs1.from_dd}}', M.jobs1?.from_dd || ''); // テンプレに置いた場合だけ出る
  replaceAll_(body, '{{jobs1.end_text}}', M.jobs1?.end_text || '');
  replaceAll_(body, '{{jobs1.company}}', M.jobs1?.company || '');
  replaceAll_(body, '{{jobs1.avg_monthly}}', M.jobs1?.avg_monthly || '');
  replaceAll_(body, '{{chk.jobs1.severance}}', renderCheck_(!!M.jobs1?.severance));

  replaceAll_(body, '{{chk.jobs1.kind.employee}}', renderCheck_(!!M.jobs1?.kind?.employee));
  replaceAll_(body, '{{chk.jobs1.kind.part}}', renderCheck_(!!M.jobs1?.kind?.part));
  replaceAll_(body, '{{chk.jobs1.kind.self}}', renderCheck_(!!M.jobs1?.kind?.self));
  replaceAll_(
    body,
    '{{chk.jobs1.kind.representative}}',
    renderCheck_(!!M.jobs1?.kind?.representative)
  );
  replaceAll_(body, '{{chk.jobs1.kind.unemployed}}', renderCheck_(!!M.jobs1?.kind?.unemployed));

  doc.saveAndClose();
  return { gdocId, draftUrl: doc.getUrl() };
}

/** ====== ヘルパ群（共通が無い場合のみ使われます） ====== **/

// すでにS2002側に同名があるなら、そちらが使われます（同実装）。
function replaceAll_(body, token, value) {
  const safe = token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  body.replaceText(safe, String(value ?? ''));
}
function normJaSpace_(s) {
  return String(s || '')
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
function findFieldValue_(m, labelRegex) {
  const arr = (m && m.fieldsRaw) || [];
  const hit = arr.find((f) => labelRegex.test(String(f.label || '')));
  return hit ? String(hit.value || '') : '';
}
function splitMulti_(s) {
  return String(s || '')
    .split(/[\r\n、，,・;；]+/)
    .map((v) => normJaSpace_(v))
    .filter(Boolean);
}
function hasAny_(arr, re) {
  return (arr || []).some((v) => re.test(v));
}
function pickOtherText_(raw) {
  const m = String(raw || '').match(/その他[：: ]*([^\n]+)/);
  return m ? normJaSpace_(m[1]) : '';
}
function toYMP_(s) {
  const a = toYMD_(s);
  return { yyyy: a.yyyy, mm: a.mm };
}
function toYMD_(s) {
  let str = String(s || '').trim();
  const era = str.match(/(令和|平成|昭和)\s*(\d+)年(?:\s*(\d+)月)?(?:\s*(\d+)日)?/);
  if (era) {
    const y = { 令和: 2018, 平成: 1988, 昭和: 1925 }[era[1]] + parseInt(era[2], 10);
    return { yyyy: String(y), mm: era[3] ? String(era[3]) : '', dd: era[4] ? String(era[4]) : '' };
  }
  const m = str
    .replace(/[年月]/g, '-')
    .replace(/[.\/]/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .split('-');
  return { yyyy: m[0] || '', mm: m[1] || '', dd: m[2] || '' };
}
function toNumberText_(s) {
  const n = String(s || '').replace(/[^\d]/g, '');
  return n ? n.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : '';
}
function toBoolJaLoose_(v) {
  const s = String(v || '').trim();
  return /^(はい|有|あり|有り|true|当|○|チェック)$/i.test(s);
}
function parseNameKanaValue_(raw) {
  const s = String(raw || '').trim();
  const isKana = (t) => /^[\p{Script=Hiragana}\p{Script=Katakana}\u30FC\s・･ｰﾞﾟ\-]+$/u.test(t);
  const hasKanji = (t) => /[\p{Script=Han}]/u.test(t);
  let m = s.match(/^(.+?)\s*[（(]\s*([^)）]+?)\s*[)）]\s*$/);
  if (m) {
    let [_, a, b] = m;
    a = a.trim();
    b = b.trim();
    if (hasKanji(a) || !isKana(a)) return { name: a, kana: b };
    if (hasKanji(b) && isKana(a)) return { name: b, kana: a };
    return { name: a, kana: b };
  }
  m = s.match(/^(.+?)\s*[/｜\|]\s*(.+)$/);
  if (m) {
    let a = m[1].trim(),
      b = m[2].trim();
    if (hasKanji(a) || !isKana(a)) return { name: a, kana: b };
    if (hasKanji(b) && isKana(a)) return { name: b, kana: a };
    return { name: a, kana: b };
  }
  return isKana(s) ? { name: '', kana: s } : { name: s, kana: '' };
}
function normSpace_(s) {
  return String(s || '')
    .replace(/\s+/g, ' ')
    .trim();
}
function normPhone_(s) {
  const t = String(s || '').replace(/[^\d]/g, '');
  return t ? t.replace(/(\d{2,4})(\d{2,4})(\d{3,4})/, '$1-$2-$3') : '';
}
function toIsoBirth_(ja) {
  let s = String(ja)
    .trim()
    .replace(/[年月]/g, '-')
    .replace(/[日]/g, '')
    .replace(/[.\/\s]/g, '-')
    .replace(/-+/g, '-');
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return '';
  const y = +m[1],
    mo = ('0' + m[2]).slice(-2),
    d = ('0' + m[3]).slice(-2);
  return `${y}-${mo}-${d}`;
}
function calcAge_(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  const today = new Date(),
    b = new Date(y, m - 1, d);
  let age = today.getFullYear() - y;
  const md = (today.getMonth() + 1) * 100 + today.getDate(),
    bd = m * 100 + d;
  if (md < bd) age--;
  return age;
}
function toWareki_(iso) {
  if (!iso) return null;
  const d = new Date(iso);
  const eras = [
    { gengo: '令和', start: new Date('2019-05-01'), offset: 2018 },
    { gengo: '平成', start: new Date('1989-01-08'), offset: 1988 },
    { gengo: '昭和', start: new Date('1926-12-25'), offset: 1925 },
  ];
  for (const e of eras)
    if (d >= e.start)
      return {
        gengo: e.gengo,
        yy: d.getFullYear() - e.offset,
        mm: d.getMonth() + 1,
        dd: d.getDate(),
      };
  return null;
}
function safeHtml(s) {
  return String(s).replace(
    /[&<>"']/g,
    (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m])
  );
}

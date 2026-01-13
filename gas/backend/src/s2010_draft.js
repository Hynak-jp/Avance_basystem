/** ====== 設定（S2010） ====== **
 * 必要な Script Properties:
 * - BAS_MASTER_SPREADSHEET_ID : cases台帳シートのID（S2002と共通）
 * - S2010_TEMPLATE_FILE_ID    : S2010差し込み用テンプレのファイルID（docx/gdoc）
 * - S2010_TPL_GDOC_ID (or S2010_TEMPLATE_GDOC_ID) : 旧gdocテンプレのファイルID（移行期間のみ）
 */
const PROP_S2010 = PropertiesService.getScriptProperties();
const S2010_SPREADSHEET_ID = PROP_S2010.getProperty('BAS_MASTER_SPREADSHEET_ID') || '';
const S2010_TPL_FILE_ID = PROP_S2010.getProperty('S2010_TEMPLATE_FILE_ID') || '';
const S2010_TPL_GDOC_ID =
  PROP_S2010.getProperty('S2010_TPL_GDOC_ID') ||
  PROP_S2010.getProperty('S2010_TEMPLATE_GDOC_ID') ||
  '';
const S2010_LABEL_TO_PROCESS = 'FormAttach/ToProcess';
const S2010_LABEL_PROCESSED = 'FormAttach/Processed';
// S2010 の分割フォームが揃っているか判定するための form_key 接頭辞リスト（p1/p2のみ）
const S2010_PART_PREFIXES = ['s2010_p1_', 's2010_p2_'];
const S2010_LABEL_CACHE = {};
const S2010_MIME_GDOC = MimeType.GOOGLE_DOCS;
const S2010_MIME_DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const S2010_UNRESOLVED_ALLOWLIST = [];

// チェック記号はテンプレ仕様に合わせる
const CHECKED = '■';
const UNCHECKED = '□';
function s2010_renderCheck_(b) {
  return b ? CHECKED : UNCHECKED;
}

function s2010_labelOrCreate_(name) {
  if (!name) return null;
  if (S2010_LABEL_CACHE[name]) return S2010_LABEL_CACHE[name];
  const label = GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
  S2010_LABEL_CACHE[name] = label;
  return label;
}

/** ====== パブリック・エントリ ====== **/

/**
 * 受信箱から S2010 通知メールを取り込み → JSON 保存 → 必須パートが揃えば統合 JSON 保存 + ドラフト生成
 * ラベル: FormAttach/ToProcess → 処理後に FormAttach/Processed を付与
 */
// Manual test checklist:
// 1) ToProcess -> Processed label transition happens on success.
// 2) P1 only does not generate draft; P1+P2 merges and generates.
// 3) unresolved placeholders log is 0 or expected tokens.
function run_ProcessInbox_S2010() {
  const query = `label:${S2010_LABEL_TO_PROCESS} subject:#FM-BAS subject:S2010`;
  const threads = GmailApp.search(query, 0, 50);
  const processedLabel = s2010_labelOrCreate_(S2010_LABEL_PROCESSED);
  const toProcessLabel = s2010_labelOrCreate_(S2010_LABEL_TO_PROCESS);
  try {
    Logger.log('[S2010] tick: query=%s threads=%s', query, threads.length);
  } catch (_) {}
  if (!threads.length) return;

  threads.forEach((th) => {
    th.getMessages().forEach((msg) => {
      let processedOk = false;
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
        let resolvedCaseKey = '';
        try {
          resolvedCaseKey = drive_resolveCaseKeyFromMeta_(parsed.meta || {}, fallbackInfo);
        } catch (err) {
          Logger.log('[S2010] caseKey resolve failed: %s', (err && err.message) || err);
          return; // caseKey が取れない場合は保留（ラベルは維持）
        }
        const caseFolderId = s2010_resolveExistingCaseFolderId_({
          ...caseInfo,
          caseKey: resolvedCaseKey,
          case_key: resolvedCaseKey,
          caseId: fallbackInfo.caseId,
          case_id: fallbackInfo.case_id,
        });
        if (!caseFolderId) {
          Logger.log('[S2010] case folder not found. key=%s', resolvedCaseKey);
          return; // intake前などのため、ToProcess に残しておく
        }
        caseInfo.folderId = caseFolderId;
        caseInfo.caseKey = resolvedCaseKey;
        caseInfo.userKey = resolvedCaseKey.split('-')[0];
        caseInfo.user_key = caseInfo.userKey;
        if (!caseInfo.lineId && fallbackInfo.lineId) caseInfo.lineId = fallbackInfo.lineId;

        saveSubmissionJson_(caseFolderId, parsed);

        const caseIdForProcess = fallbackInfo.caseId || parsed.meta?.case_id || parsed.meta?.caseId || '';
        const hasCaseIdForProcess = !!caseIdForProcess;
        if (typeof updateCasesRow_ === 'function' && hasCaseIdForProcess) {
          const patch = {
            case_key: resolvedCaseKey,
            folder_id: caseFolderId,
            user_key: caseInfo.userKey,
            last_activity: new Date(),
          };
          updateCasesRow_(caseIdForProcess, patch);
        }

        if (haveAllPartsS2010_(caseInfo.folderId, S2010_PART_PREFIXES)) {
          if (hasCaseIdForProcess) {
            run_GenerateS2010DraftMergedByCaseId(caseIdForProcess);
          } else {
            Logger.log('[S2010] case_id is empty. draft generation skipped.');
          }
        }

        processedOk = true;
      } catch (e) {
        try {
          const alertTo = Session.getActiveUser().getEmail();
          if (alertTo) {
            GmailApp.createDraft(alertTo, '[BAS Intake Error]', String(e), {
              htmlBody: `<pre>${s2010_safeHtml_(e.stack || e)}</pre>`,
            });
          }
        } catch (_) {}
      }
      if (processedOk) {
        if (processedLabel) th.addLabel(processedLabel);
        if (toProcessLabel) th.removeLabel(toProcessLabel);
      }
    });
  });
}

/** 直近の S2010 JSON を読み込んでドラフト生成（ケースID指定） */
function run_GenerateS2010DraftByCaseId(caseId) {
  const info = resolveCaseByCaseId_(caseId);
  if (!info) throw new Error(`Unknown case_id: ${caseId}`);
  info.folderId = s2010_resolveExistingCaseFolderId_(info);
  if (!info.folderId) throw new Error(`Case folder not found for case_id: ${caseId}`);

  const parsed = s2010_loadLatestFormJson_(info.folderId, 's2010_userform');
  if (!parsed) throw new Error(`No S2010 JSON found under case folder: ${info.folderId}`);

  const draft = generateS2010Draft_(info, parsed);
  const patch = { last_activity: new Date() };
  if (draft && draft.draftUrl) {
    patch.status = 'draft';
    patch.last_draft_url = draft.draftUrl;
  }
  if (draft && draft.docxUrl) patch.last_draft_docx_url = draft.docxUrl;
  updateCasesRow_(info.caseId || caseId, patch);
  try {
    if (draft && draft.draftUrl) Logger.log('[S2010] draft created: %s', draft.draftUrl);
    if (draft) {
      Logger.log('[S2010] draft urls: gdoc=%s docx=%s unresolved=%s', draft.draftUrl || '', draft.docxUrl || '', draft.unresolvedCount || 0);
    }
  } catch (_) {}
  s2010_logDraftSmoke_(draft);
  return draft;
}

/** s2010_p*_ をマージして統合JSONのみ保存（ドラフト生成はしない） */
function run_GenerateS2010MergedJsonByCaseId(caseId) {
  const info = resolveCaseByCaseId_(caseId);
  if (!info) throw new Error(`Unknown case_id: ${caseId}`);
  info.folderId = s2010_resolveExistingCaseFolderId_(info);
  if (!info.folderId) throw new Error(`Case folder not found for case_id: ${caseId}`);

  const parts = s2010_loadLatestPartsByPrefix_(info.folderId, 's2010_');
  const requiredParts = sortS2010PartsByPrefixes_(
    filterS2010PartsByPrefixes_(parts, S2010_PART_PREFIXES),
    S2010_PART_PREFIXES
  );
  if (!requiredParts.length) throw new Error('No required s2010 parts found under case folder.');

  if (!haveAllPartsS2010_(info.folderId, S2010_PART_PREFIXES)) {
    throw new Error('Not all required S2010 parts are present yet.');
  }

  const draftKey = s2010_makeDraftKeyFromParts_(requiredParts);
  const merged = mergeS2010Parts_(requiredParts, { caseId: info.caseId, draftKey: draftKey });
  const fname = `s2010_userform__merged_${draftKey}.json`;
  const parent = DriveApp.getFolderById(info.folderId);
  const existing = s2010_findFileByName_(parent, fname);
  if (!existing) {
    parent.createFile(Utilities.newBlob(JSON.stringify(merged, null, 2), 'application/json', fname));
    try {
      Logger.log('[S2010] merged json saved: %s', fname);
    } catch (_) {}
  } else {
    try {
      Logger.log('[S2010] merged json exists: %s', fname);
    } catch (_) {}
  }
  return merged;
}

/** デバッグ：'0001' などで呼び出し */
function debug_GenerateS2010_for_0001() {
  return run_GenerateS2010DraftByCaseId('0001');
}

function s2010_logDraftSmoke_(draft) {
  const draftUrl = (draft && draft.draftUrl) || '';
  const docxUrl = (draft && draft.docxUrl) || '';
  const unresolved = Array.isArray(draft && draft.unresolvedKeys) ? draft.unresolvedKeys : [];
  const allow = Array.isArray(S2010_UNRESOLVED_ALLOWLIST) ? S2010_UNRESOLVED_ALLOWLIST : [];
  const blocked = unresolved.filter((key) => !allow.includes(key));
  try {
    Logger.log('[S2010][smoke] draftUrl=%s docxUrl=%s', draftUrl, docxUrl);
    Logger.log('[S2010][smoke] urls_ok=%s', !!draftUrl && !!docxUrl);
    Logger.log('[S2010][smoke] unresolved=%s blocked=%s', unresolved.length, blocked.length);
    if (blocked.length) {
      const head = blocked.slice(0, 20).join(', ');
      Logger.log('[S2010][smoke] unresolved list (%s): %s', blocked.length, head);
    }
  } catch (_) {}
}

/** ===== マルチフォーム統合 → ドラフト生成（入口） ===== */
function run_GenerateS2010DraftMergedByCaseId(caseId) {
  // 1) cases解決 & フォルダ解決
  const info = resolveCaseByCaseId_(caseId);
  if (!info) throw new Error(`Unknown case_id: ${caseId}`);
  info.folderId = s2010_resolveExistingCaseFolderId_(info);
  if (!info.folderId) throw new Error(`Case folder not found for case_id: ${caseId}`);

  // 2) ケース直下から最新の各partを収集
  const parts = s2010_loadLatestPartsByPrefix_(info.folderId, 's2010_'); // s2010_ で始まるform_keyを全部拾う
  const requiredParts = sortS2010PartsByPrefixes_(
    filterS2010PartsByPrefixes_(parts, S2010_PART_PREFIXES),
    S2010_PART_PREFIXES
  );
  if (!requiredParts.length) throw new Error('No required S2010 part-jsons found.');

  // 3) マージして統合JSONを作る
  const draftKey = s2010_makeDraftKeyFromParts_(requiredParts);
  const merged = mergeS2010Parts_(requiredParts, { caseId: info.caseId, draftKey: draftKey }); // { meta, fieldsRaw, model }

  // 4) 統合JSONをケース直下に保存（監査用）
  const mergedName = `s2010_userform__merged_${draftKey}.json`;
  const parent = DriveApp.getFolderById(info.folderId);
  if (!s2010_findFileByName_(parent, mergedName)) {
    parent.createFile(Utilities.newBlob(JSON.stringify(merged, null, 2), 'application/json', mergedName));
  }

  // 5) 既存のS2010生成器でドラフト化
  const draft = generateS2010Draft_(info, merged);

  // 6) ステータス更新
  const patch = { last_activity: new Date() };
  if (draft && draft.draftUrl) {
    patch.status = 'draft';
    patch.last_draft_url = draft.draftUrl;
  }
  if (draft && draft.docxUrl) patch.last_draft_docx_url = draft.docxUrl;
  updateCasesRow_(info.caseId || caseId, patch);

  try {
    if (draft && draft.draftUrl) Logger.log('[S2010] merged draft created: %s', draft.draftUrl);
    if (draft) {
      Logger.log('[S2010] merged draft urls: gdoc=%s docx=%s unresolved=%s', draft.draftUrl || '', draft.docxUrl || '', draft.unresolvedCount || 0);
    }
  } catch (_){ }
  s2010_logDraftSmoke_(draft);
  return draft;
}

/** ===== ケース直下から "prefix" に合う form_key の最新を集める ===== */
function s2010_loadLatestPartsByPrefix_(caseFolderId, prefix) {
  const id = s2010_normalizeFolderId_(caseFolderId);
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

/** S2010の必須パートのみを残す（旧パート混入の防止） */
function filterS2010PartsByPrefixes_(parts, prefixes) {
  return (parts || []).filter((p) => {
    const key = getS2010FormKey_(p);
    return prefixes.some((prefix) => String(key).startsWith(prefix));
  });
}

function sortS2010PartsByPrefixes_(parts, prefixes) {
  const ordered = (parts || []).slice();
  ordered.sort((a, b) => {
    const ra = getS2010PrefixRank_(getS2010FormKey_(a), prefixes);
    const rb = getS2010PrefixRank_(getS2010FormKey_(b), prefixes);
    return ra - rb;
  });
  return ordered;
}

function getS2010PrefixRank_(key, prefixes) {
  const s = String(key || '');
  for (let i = 0; i < prefixes.length; i++) {
    if (s.startsWith(prefixes[i])) return i;
  }
  return prefixes.length;
}

function getS2010FormKey_(part) {
  return (part && part.json && part.json.meta && part.json.meta.form_key) || part.form_key || '';
}

/** ===== パーツ配列を 1つの "s2010_userform" JSON に統合 ===== */
function mergeS2010Parts_(parts, opt) {
  // 1) すべての fields を配列 [{label, value}] に正規化
  const arrays = [];
  const meta_list = [];
  parts.forEach(p => {
    const j = p.json || {};
    const arr = s2010_normalizeFieldsArrayForAny_(j); // 既存の正規化
    if (arr && arr.length) arrays.push(arr);
    meta_list.push({
      form_key: (j.meta && j.meta.form_key) || p.form_key,
      submission_id: j.meta && j.meta.submission_id,
      fileId: p.fileId
    });
  });

  // 2) fieldsRawは潰さず連結（後勝ち互換は検索側で担保）
  const fieldsRaw = [];
  arrays.forEach(arr => {
    arr.forEach(({label, value}) => {
      const v = String(value || '').trim();
      if (!v) return; // 空は無視（既存値を消さない）
      fieldsRaw.push({ label, value: v });
    });
  });

  // 3) S2010のモデリングへ
  const model = mapFieldsToModel_S2010_(fieldsRaw); // 既存のS2010マッパ
  const p1Model = s2010_pickP1Model_(parts);
  if (p1Model) {
    ['jobs', 'marital_events', 'household', 'inheritances', 'housing'].forEach((k) => {
      if (p1Model[k]) model[k] = p1Model[k];
    });
  }

  // 5) 統合メタ
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const meta = {
    form_key: 's2010_userform',
    merged_at: now,
    merged_from: meta_list,
    case_id: (opt && opt.caseId) || '',
    draft_key: (opt && opt.draftKey) || ''
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

function s2010_loadLatestFormJson_(caseFolderId, wantFormKey) {
  const id = s2010_normalizeFolderId_(caseFolderId);
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

  const fieldsArr = s2010_normalizeFieldsArrayForAny_(latest);
  if (!fieldsArr) throw new Error('Invalid submission JSON shape (no fields/model)');
  return {
    meta: latest.meta || {},
    fieldsRaw: fieldsArr,
    model: mapFieldsToModel_S2010_(fieldsArr),
  };
}

function s2010_normalizeFolderId_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  const m = s.match(/[-\w]{25,}/);
  return m ? m[0] : s;
}

function s2010_normalizeFieldsArrayForAny_(json) {
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
      const { name, kana } = s2010_parseNameKanaValue_(value);
      out.app.name = s2010_normSpace_(name);
      out.app.kana = kana;
      return;
    }
    if (/^【?名前】?$/.test(label)) {
      out.app.name = s2010_normSpace_(value);
      return;
    }
    if (/^【?メールアドレス】?/.test(label)) {
      out.app.email = String(value || '').trim();
      return;
    }
    if (/^【?連絡先】?/.test(label)) {
      out.app.phone = s2010_normPhone_(value);
      return;
    }
    if (/^【?生年月日】?/.test(label)) {
      out.app.birth = String(value || '').trim();
      return;
    }
    if (/^【?住居所（住民票と異なる[^】]*）】?$/.test(label)) {
      out.addr.alt_full = s2010_normJaSpace_(value);
      return;
    }
    if (/^【?住居所】?/.test(label)) {
      out.addr.full = s2010_normJaSpace_(value);
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
    const iso = s2010_toIsoBirth_(out.app.birth);
    out.app.birth_iso = iso;
    out.app.age = iso ? s2010_calcAge_(iso) : '';
    const w = s2010_toWareki_(iso);
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
  const reasonsRaw = s2010_findFieldValue_(m, /理由は.*とおり|借金.*理由/);
  const reasons = s2010_splitMulti_(reasonsRaw);
  out.reason = {
    living: s2010_hasAny_(reasons, /生活費/),
    mortgage: s2010_hasAny_(reasons, /住宅ローン|住宅/),
    education: s2010_hasAny_(reasons, /教育/),
    waste: s2010_hasAny_(reasons, /浪費|飲食|飲酒|投資|投機|商品購入|ギャンブル/),
    business: s2010_hasAny_(reasons, /事業|経営破綻|マルチ|ネットワーク/),
    guarantee: s2010_hasAny_(reasons, /保証/),
    other: s2010_hasAny_(reasons, /その他/),
    other_text: s2010_pickOtherText_(reasonsRaw),
  };

  // きっかけ（複数選択想定）
  const triggersRaw = s2010_findFieldValue_(m, /きっかけ.*とおり|返済.*できなく.*きっかけ/);
  const triggers = s2010_splitMulti_(triggersRaw);
  out.trigger = {
    overpay: s2010_hasAny_(triggers, /収入以上|返済金額/),
    dismiss: s2010_hasAny_(triggers, /解雇/),
    paycut: s2010_hasAny_(triggers, /減額/),
    hospital: s2010_hasAny_(triggers, /病気|入院/),
    other: s2010_hasAny_(triggers, /その他/),
    other_text: s2010_pickOtherText_(triggersRaw),
  };

  // 支払不能の時期・約定返済合計・受任日（任意）
  const unable = s2010_findFieldValue_(m, /支払不能.*時期/);
  const unableYM = s2010_toYMP_(unable);
  out.unable_yyyy = unableYM.yyyy || '';
  out.unable_mm = unableYM.mm || '';
  const monthly = s2010_findFieldValue_(m, /約定返済額|月々.*返済額/);
  out.unable_monthly_total = s2010_toNumberText_(monthly);

  const notice = s2010_findFieldValue_(m, /受任通知.*発送日/);
  const nd = s2010_toYMD_(notice);
  out.notice_yyyy = nd.yyyy || '';
  out.notice_mm = nd.mm || '';
  out.notice_dd = nd.dd || '';

  // ===== 職歴1（開始は常に入力あり。終了/現在は入力無しでもOK。無職でも開始だけ出す） =====
  const companyRaw = s2010_findFieldValue_(m, /職歴1.*(就業先|会社名|勤務先)/);
  const kindRaw = s2010_findFieldValue_(m, /職歴1.*種\s*別|種\s*別/);
  const monthlyRaw = s2010_findFieldValue_(
    m,
    /職歴1.*平均月収(\s*[（(]円[）)]\s*)?|平均月収(\s*[（(]円[）)]\s*)?/
  );
  const severanceRaw = s2010_findFieldValue_(m, /職歴1.*退職金.*有|退職金.*有/);

  const kind = String(kindRaw || '');
  const kindFlags = {
    employee: /勤め|正社員|従業員/i.test(kind),
    part: /パート|ｱﾙﾊﾞｲﾄ|アルバイト|派遣|契約/i.test(kind),
    self: /自営|自営業|個人事業/i.test(kind),
    representative: /法人代表者|代表|社長|取締役/i.test(kind),
    unemployed: /無職|失業/i.test(kind),
  };

  // 開始は常に拾う（「職歴1：開始」「職歴1：就業開始」「無職開始」など許容）
  const startRaw = s2010_findFieldValue_(
    m,
    /職歴1[:：]?\s*(開始|就業開始|無職開始)|就業期間.*開始|職歴.*1.*開始/
  );
  const endRaw = s2010_findFieldValue_(m, /職歴1[:：]?\s*(就業終了|終了)|就業期間.*終了|職歴.*1.*終了/);
  const curRaw = s2010_findFieldValue_(m, /職歴1[:：]?\s*現\s*在|職歴1.*現在|現在就業中/);

  const fromYMD = s2010_toYMD_(startRaw); // {yyyy,mm,dd}
  const toYMDv = s2010_toYMD_(endRaw); // {yyyy,mm,dd}

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
    company: s2010_normJaSpace_(companyRaw || ''),
    kind: kindFlags,
    avg_monthly: kindFlags.unemployed ? '' : s2010_toNumberText_(monthlyRaw),
    severance: s2010_toBoolJaLoose_(severanceRaw),

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

  const p1Jobs = Array.isArray(m.jobs) ? m.jobs : [];
  if (p1Jobs.length) {
    out.jobs1 = s2010_buildJobFromP1_(p1Jobs[0]);
  }
  const p1Household = Array.isArray(m.household) ? m.household : [];
  if (p1Household.length) {
    out.hh = s2010_buildHouseholdFromP1_(p1Household);
  }

  return out;
}

/** ====== ドラフト生成（S2010） ====== **/
function generateS2010Draft_(caseInfo, parsed) {
  const templateId = s2010_getTemplateFileId_();
  if (!templateId) {
    Logger.log('[S2010] S2010_TEMPLATE_FILE_ID is not set. draft skipped.');
    return { gdocId: '', draftUrl: '', docxId: '', docxUrl: '' };
  }
  const drafts = getOrCreateSubfolder_(DriveApp.getFolderById(caseInfo.folderId), 'drafts');

  const caseId = caseInfo.caseId || caseInfo.case_id || 'unknown';
  const subId =
    (parsed.meta && (parsed.meta.draft_key || parsed.meta.submission_id)) ||
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const logCtx = `caseId=${caseId} draftKey=${subId}`;
  const draftName = `S2010_${caseId}_draft_${subId}`;
  const buildExisting = (gdocFile) => {
    let unresolved = [];
    try {
      const existingDoc = DocumentApp.openById(gdocFile.getId());
      unresolved = s2010_findUnresolvedPlaceholders_(existingDoc);
      if (unresolved.length) {
        const head = unresolved.slice(0, 20).join(', ');
        Logger.log('[S2010] unresolved placeholders (%s) %s: %s', unresolved.length, logCtx, head);
      }
      existingDoc.saveAndClose();
    } catch (_) {}
    const docx = s2010_exportDocxFromGdoc_(gdocFile.getId(), draftName, drafts);
    return {
      gdocId: gdocFile.getId(),
      draftUrl: gdocFile.getUrl(),
      docxId: docx.docxId,
      docxUrl: docx.docxUrl,
      unresolvedCount: unresolved.length,
      unresolvedKeys: unresolved,
    };
  };

  const existingGdoc = s2010_findFileByName_(drafts, draftName);
  if (existingGdoc) {
    return buildExisting(existingGdoc);
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    Logger.log('[S2010] lock wait failed: %s', (e && e.message) || e);
    return { gdocId: '', draftUrl: '', docxId: '', docxUrl: '' };
  }

  try {
    const existingAfterLock = s2010_findFileByName_(drafts, draftName);
    if (existingAfterLock) {
      return buildExisting(existingAfterLock);
    }

    const gdocId = s2010_copyTemplateToDrafts_(templateId, draftName, drafts);
    const doc = DocumentApp.openById(gdocId);
    const M = ensureS2010Model_({
      ...(parsed.model || {}),
      fieldsRaw: parsed.fieldsRaw || [],
    });

    const placeholderMap = s2010_buildS2010PlaceholderMap_(M, parsed, logCtx);
    try {
      Logger.log('[S2010] placeholder keys=%s %s', Object.keys(placeholderMap || {}).length, logCtx);
    } catch (_) {}
    s2010_applyS2010Placeholders_(doc, placeholderMap);

    const unresolved = s2010_findUnresolvedPlaceholders_(doc);
    if (unresolved.length) {
      const head = unresolved.slice(0, 20).join(', ');
      Logger.log('[S2010] unresolved placeholders (%s) %s: %s', unresolved.length, logCtx, head);
      try {
        Logger.log('[S2010] unresolved debug keys=%s %s: %s', Object.keys(placeholderMap || {}).length, logCtx, head);
      } catch (_) {}
    }

    doc.saveAndClose();
    const docx = s2010_exportDocxFromGdoc_(gdocId, draftName, drafts);
    const draftUrl = DriveApp.getFileById(gdocId).getUrl();
    const docxUrl = docx.docxId ? DriveApp.getFileById(docx.docxId).getUrl() : docx.docxUrl;
    return {
      gdocId,
      draftUrl,
      docxId: docx.docxId,
      docxUrl,
      unresolvedCount: unresolved.length,
      unresolvedKeys: unresolved,
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

/** ====== ヘルパ群（共通が無い場合のみ使われます） ====== **/

function s2010_getTemplateFileId_() {
  if (S2010_TPL_FILE_ID) return S2010_TPL_FILE_ID;
  if (S2010_TPL_GDOC_ID) {
    try {
      Logger.log('[S2010] using legacy template id (S2010_TPL_GDOC_ID)');
    } catch (_) {}
  }
  return S2010_TPL_GDOC_ID || '';
}

function s2010_copyTemplateToDrafts_(templateId, draftName, draftsFolder) {
  const file = DriveApp.getFileById(templateId);
  const mime = file.getMimeType();
  if (mime === S2010_MIME_GDOC) {
    return file.makeCopy(draftName, draftsFolder).getId();
  }
  if (mime === S2010_MIME_DOCX) {
    if (!Drive || !Drive.Files || typeof Drive.Files.copy !== 'function') {
      throw new Error('[S2010] Drive.Files.copy unavailable for docx template');
    }
    try {
      Logger.log('[S2010] docx template detected. converting to gdoc: %s', templateId);
    } catch (_) {}
    const resource = {
      title: draftName,
      mimeType: S2010_MIME_GDOC,
      parents: [{ id: draftsFolder.getId() }],
    };
    const copied = Drive.Files.copy(resource, templateId, { convert: true });
    return copied.id;
  }
  throw new Error('[S2010] unsupported template mime: ' + mime);
}

function s2010_exportDocxFromGdoc_(gdocId, draftName, draftsFolder) {
  if (!gdocId) return { docxId: '', docxUrl: '' };
  const docxName = draftName + '.docx';
  const existing = s2010_findFileByName_(draftsFolder, docxName);
  if (existing) {
    const existingId = existing.getId();
    return { docxId: existingId, docxUrl: DriveApp.getFileById(existingId).getUrl() };
  }
  try {
    const url = 'https://www.googleapis.com/drive/v3/files/' + gdocId + '/export?mimeType=' + encodeURIComponent(S2010_MIME_DOCX);
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    });
    if (resp.getResponseCode() === 200) {
      const blob = resp.getBlob().setName(docxName);
      const docx = draftsFolder.createFile(blob);
      const docxId = docx.getId();
      return { docxId: docxId, docxUrl: DriveApp.getFileById(docxId).getUrl() };
    }
    Logger.log('[S2010] DOCX export failed: %s', resp.getContentText());
  } catch (e) {
    Logger.log('[S2010] DOCX export error: %s', e);
  }
  return { docxId: '', docxUrl: '' };
}

function s2010_findFileByName_(folder, name) {
  if (!folder || !name) return null;
  const it = folder.getFilesByName(name);
  return it.hasNext() ? it.next() : null;
}

function s2010_makeDraftKeyFromParts_(parts) {
  const list = (parts || []).map((p) => {
    const key = getS2010FormKey_(p);
    const meta = (p && p.json && p.json.meta) || {};
    const sid = meta.submission_id || meta.submissionId || p.fileId || '';
    return key + ':' + String(sid || p.fileId || '');
  });
  const raw = list.join('|');
  if (!raw) {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  }
  return s2010_sha1Hex_(raw);
}

function s2010_sha1Hex_(input) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_1,
    String(input || ''),
    Utilities.Charset.UTF_8
  );
  return bytes
    .map(function (b) {
      const v = (b < 0 ? b + 256 : b).toString(16);
      return v.length === 1 ? '0' + v : v;
    })
    .join('');
}

function s2010_resolveExistingCaseFolderId_(caseInfo) {
  const rawId =
    typeof extractDriveIdMaybe_ === 'function'
      ? extractDriveIdMaybe_(caseInfo.folderId || '')
      : s2010_extractDriveIdMaybe_(caseInfo.folderId || '');
  if (rawId) {
    try {
      const f = DriveApp.getFolderById(rawId);
      if (f && f.getId()) return rawId;
    } catch (_) {}
  }

  const caseId = caseInfo.caseId || caseInfo.case_id || '';
  let caseKey = caseInfo.caseKey || caseInfo.case_key || '';
  if (!caseKey) {
    const lid = String(caseInfo.lineId || caseInfo.line_id || '').trim();
    const normCase = typeof normalizeCaseIdString_ === 'function'
      ? normalizeCaseIdString_
      : function (v) { return String(v || '').trim(); };
    const cid = normCase(caseId);
    if (lid && cid) caseKey = lid.slice(0, 6).toLowerCase() + '-' + cid;
  }

  if (caseKey) {
    let foundId = '';
    if (typeof findBestCaseFolderUnderRoot_ === 'function') {
      const best = findBestCaseFolderUnderRoot_(caseKey);
      if (best && best.id) foundId = best.id;
    }
    if (!foundId && typeof drive_getRootFolder_ === 'function') {
      const it = drive_getRootFolder_().getFoldersByName(caseKey);
      if (it.hasNext()) foundId = it.next().getId();
    }
    if (foundId) {
      try {
        if (caseId && typeof updateCasesRow_ === 'function') {
          updateCasesRow_(caseId, { folder_id: foundId, case_key: caseKey });
        }
      } catch (_) {}
      return foundId;
    }
  }
  return '';
}

function s2010_extractDriveIdMaybe_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  let m = s.match(/\/folders\/([A-Za-z0-9_-]{10,})/);
  if (m) return m[1];
  m = s.match(/\/d\/([A-Za-z0-9_-]{10,})/);
  if (m) return m[1];
  if (/^[A-Za-z0-9_-]{10,}$/.test(s)) return s;
  return '';
}

function s2010_pickP1Model_(parts) {
  const found = (parts || []).find((p) => String(getS2010FormKey_(p)).startsWith('s2010_p1_'));
  return found && found.json && found.json.model ? found.json.model : null;
}

function s2010_pad2_(v) {
  return v ? String(v).padStart(2, '0') : '';
}

function s2010_parseYmdParts_(v) {
  const d = s2010_toYMD_(v);
  return { yyyy: d.yyyy || '', mm: s2010_pad2_(d.mm || ''), dd: s2010_pad2_(d.dd || '') };
}

function s2010_formatYmdJa_(v) {
  const d = s2010_parseYmdParts_(v);
  if (!d.yyyy) return '';
  if (d.mm && d.dd) return `${d.yyyy}年${d.mm}月${d.dd}日`;
  if (d.mm) return `${d.yyyy}年${d.mm}月`;
  return `${d.yyyy}年`;
}

function s2010_jobKindFlags_(kindRaw) {
  const kind = String(kindRaw || '');
  return {
    employee: /勤め|正社員|従業員/i.test(kind),
    part: /パート|ｱﾙﾊﾞｲﾄ|アルバイト|派遣|契約/i.test(kind),
    self: /自営|自営業|個人事業/i.test(kind),
    representative: /法人代表者|代表|社長|取締役/i.test(kind),
    unemployed: /無職|失業/i.test(kind),
  };
}

function s2010_buildJobFromP1_(job) {
  const start = s2010_parseYmdParts_(job.start_iso || '');
  const end = s2010_parseYmdParts_(job.end_iso || '');
  const kind = s2010_jobKindFlags_(job.type || '');
  const isUnemployed = !!kind.unemployed;
  const hasEnd = !!(end.yyyy || end.mm);
  const isCurrent = !isUnemployed && !hasEnd;
  const endText = isCurrent ? '現　在' : end.yyyy && end.mm ? `${end.yyyy}年${end.mm}月` : '';
  return {
    company: s2010_normJaSpace_(job.employer || ''),
    from_yyyy: start.yyyy || '',
    from_mm: start.mm || '',
    from_dd: start.dd || '',
    end_yyyy: isUnemployed ? '' : end.yyyy || '',
    end_mm: isUnemployed ? '' : end.mm || '',
    end_dd: isUnemployed ? '' : end.dd || '',
    to_yyyy: isUnemployed ? '' : end.yyyy || '',
    to_mm: isUnemployed ? '' : end.mm || '',
    to_dd: isUnemployed ? '' : end.dd || '',
    end_text: isUnemployed ? '' : endText,
    avg_monthly: isUnemployed ? '' : s2010_toNumberText_(job.avg_month_income),
    severance: s2010_toBoolJaLoose_(job.severance),
    kind: kind,
  };
}

function s2010_buildJobFromLegacy_(job) {
  return {
    company: s2010_normJaSpace_((job && job.company) || ''),
    from_yyyy: (job && job.from_yyyy) || '',
    from_mm: (job && job.from_mm) || '',
    from_dd: (job && job.from_dd) || '',
    end_yyyy: (job && job.end_yyyy) || '',
    end_mm: (job && job.end_mm) || '',
    end_dd: (job && job.end_dd) || '',
    to_yyyy: (job && (job.to_yyyy || job.end_yyyy)) || '',
    to_mm: (job && (job.to_mm || job.end_mm)) || '',
    to_dd: (job && (job.to_dd || job.end_dd)) || '',
    end_text: (job && job.end_text) || '',
    avg_monthly: (job && job.avg_monthly) || '',
    severance: !!(job && job.severance),
    kind: (job && job.kind) || s2010_jobKindFlags_(''),
  };
}

function s2010_emptyJob_() {
  return {
    company: '',
    from_yyyy: '',
    from_mm: '',
    from_dd: '',
    end_yyyy: '',
    end_mm: '',
    end_dd: '',
    to_yyyy: '',
    to_mm: '',
    to_dd: '',
    end_text: '',
    avg_monthly: '',
    severance: false,
    kind: s2010_jobKindFlags_(''),
  };
}

function s2010_toCohab_(v) {
  const s = String(v || '').trim();
  if (/^同(居)?$/.test(s)) return true;
  if (/^別(居)?$/.test(s)) return false;
  return null;
}

function s2010_buildHouseholdFromP1_(rows) {
  const out = [];
  for (let i = 0; i < 4; i++) {
    const row = rows[i] || {};
    out.push({
      name: s2010_normSpace_(row.name || ''),
      relation: String(row.relation || '').trim(),
      age: row.age != null ? String(row.age) : '',
      occupation: String(row.occupation_or_grade || row.occupation || '').trim(),
      cohab: s2010_toCohab_(row.living),
      income: s2010_toNumberText_(row.avg_month_income),
    });
  }
  return out;
}

function s2010_maritalFlags_(reason) {
  const s = String(reason || '');
  return {
    marriage: /結婚/.test(s),
    divorce: /離婚/.test(s),
    commonlaw: /内縁(?!解消|終了)/.test(s),
    commonlaw_end: /内縁解消|内縁終了/.test(s),
    death: /死別|死亡/.test(s),
  };
}

function s2010_inheritanceStatusFlags_(statusText) {
  const s = String(statusText || '');
  const waive = /放棄/.test(s);
  const dispose = /遺産.*処理|処分/.test(s);
  const other = !!s && !waive && !dispose;
  return { waive: waive, dispose: dispose, other: other };
}

function s2010_inheritanceStatusFlagsV2_(statusText, logCtx) {
  const s = String(statusText || '').trim();
  const flags = { waive: false, divided: false, pending: false, none: false };
  if (!s) return flags;
  if (/放棄/.test(s)) {
    flags.waive = true;
    return flags;
  }
  if (/分割済|分割完了|分割|協議成立|協議済|解決/.test(s)) {
    flags.divided = true;
    return flags;
  }
  if (/未分割|協議中|手続中|調停|審判|裁判|係争|保留|pending/i.test(s)) {
    flags.pending = true;
    return flags;
  }
  if (/なし|無|該当なし|特になし/.test(s)) {
    flags.none = true;
    return flags;
  }
  try {
    Logger.log('[S2010] unknown inheritance status: %s %s', s, logCtx || '');
  } catch (_) {}
  return flags;
}

function s2010_housingTypeFlags_(type) {
  const s = String(type || '');
  return {
    private_rent: /民間賃貸住宅/.test(s),
    public_rent: /公営住宅/.test(s),
    owned: /持ち家/.test(s),
    other_person: /申立人以外の者/.test(s),
    other: /その他/.test(s),
  };
}

function s2010_ownOrRentFlags_(value) {
  const s = String(value || '');
  return {
    own: /所有|持ち家/.test(s),
    rent: /賃借|賃貸/.test(s),
  };
}

function s2010_housingOwnedDetailFlags_(detail) {
  const s = String(detail || '');
  return {
    detached: /一戸建|戸建/.test(s),
    mansion: /マンション|分譲/.test(s),
  };
}

function s2010_pickFieldValue_(parsed, re) {
  return s2010_findFieldValue_(parsed || {}, re);
}

function s2010_pickHousingDesc_(parsed) {
  return s2010_pickFieldValue_(parsed, /現在の住居の状況（居住する家屋の形態等）/);
}

function s2010_buildS2010PlaceholderMap_(m, parsed, logCtx) {
  const M = m || {};
  const map = {};
  const setText = function (token, value) {
    map[token] = value == null ? '' : String(value);
  };
  const setCheck = function (token, value) {
    map[token] = s2010_renderCheck_(!!value);
  };

  const app = M.app || {};
  const addr = M.addr || {};
  const stmt = M.stmt || {};
  const birthParts = String(app.birth_iso || '').split('-');
  const w = app.birth_wareki || null;

  setText('{{app.name}}', app.name || '');
  setText('{{app.kana}}', app.kana || '');
  setText('{{app.birth_yyyy}}', birthParts[0] || '');
  setText('{{app.birth_mm}}', birthParts[1] || '');
  setText('{{app.birth_dd}}', birthParts[2] || '');
  setText('{{app.birth_wareki_gengo}}', (w && w.gengo) || '');
  setText('{{app.birth_wareki_yy}}', (w && String(w.yy)) || '');
  setText('{{app.age}}', app.age != null ? String(app.age) : '');
  setText('{{addr.full}}', s2010_normJaSpace_(addr.full || ''));
  setText('{{addr.alt_full}}', s2010_normJaSpace_(addr.alt_full || ''));
  setText('{{stmt.free}}', (stmt.free || '').trim());

  // household (hh)
  let hhRows = Array.isArray(M.hh) ? M.hh : [];
  const p1Household = Array.isArray(M.household) ? M.household : [];
  if (p1Household.length) hhRows = s2010_buildHouseholdFromP1_(p1Household);
  for (let i = 1; i <= 4; i++) {
    const h = hhRows[i - 1] || {};
    const cohabSame = h.cohab === true || /同居/.test(String(h.cohab || ''));
    const cohabSeparate = h.cohab === false || /別居/.test(String(h.cohab || ''));
    const cohabText =
      h.cohab === true ? '同居' : h.cohab === false ? '別居' : String(h.cohab || '');

    setText(`{{hh${i}.name}}`, h.name || '');
    setText(`{{hh${i}.relation}}`, h.relation || '');
    setText(`{{hh${i}.age}}`, h.age || '');
    setText(`{{hh${i}.occupation}}`, h.occupation || '');
    setText(`{{hh${i}.job_or_grade}}`, h.occupation || '');
    setText(`{{hh${i}.cohab}}`, cohabText);
    setText(`{{hh${i}.living}}`, cohabText);
    setText(`{{hh${i}.income}}`, h.income || '');
    setText(`{{hh${i}.avg_monthly}}`, h.income || '');
    setText(`{{hh${i}.monthly}}`, h.income || '');
    setCheck(`{{chk.hh${i}.cohab.same}}`, cohabSame);
    setCheck(`{{chk.hh${i}.cohab.separate}}`, cohabSeparate);
    setCheck(`{{chk.hh${i}.cohab}}`, cohabSame);
    setCheck(`{{chk.hh${i}.separate}}`, cohabSeparate);
    setCheck(`{{chk.hh${i}.cohab.live}}`, cohabSame);
    setCheck(`{{chk.hh${i}.cohab.separate}}`, cohabSeparate);
  }

  // reason/trigger/unable/notice
  setCheck('{{chk.reason.living}}', !!M.reason?.living);
  setCheck('{{chk.reason.mortgage}}', !!M.reason?.mortgage);
  setCheck('{{chk.reason.education}}', !!M.reason?.education);
  setCheck('{{chk.reason.business}}', !!M.reason?.business);
  setCheck('{{chk.reason.guarantee}}', !!M.reason?.guarantee);
  setCheck('{{chk.reason.waste}}', !!M.reason?.waste);
  setCheck('{{chk.reason.other}}', !!M.reason?.other);
  setText('{{reason.other_text}}', M.reason?.other_text || '');

  setCheck('{{chk.trigger.overpay}}', !!M.trigger?.overpay);
  setCheck('{{chk.trigger.dismiss}}', !!M.trigger?.dismiss);
  setCheck('{{chk.trigger.paycut}}', !!M.trigger?.paycut);
  setCheck('{{chk.trigger.hospital}}', !!M.trigger?.hospital);
  setCheck('{{chk.trigger.other}}', !!M.trigger?.other);
  setText('{{trigger.other_text}}', M.trigger?.other_text || '');

  setText('{{unable_yyyy}}', M.unable_yyyy || '');
  setText('{{unable_mm}}', M.unable_mm || '');
  setText('{{unable_monthly_total}}', M.unable_monthly_total || '');
  setText('{{notice_yyyy}}', M.notice_yyyy || '');
  setText('{{notice_mm}}', M.notice_mm || '');
  setText('{{notice_dd}}', M.notice_dd || '');

  // jobs
  const jobs = Array.isArray(M.jobs) ? M.jobs : [];
  for (let i = 1; i <= 8; i++) {
    const src = jobs[i - 1];
    const job = src
      ? s2010_buildJobFromP1_(src)
      : i === 1 && M.jobs1
      ? s2010_buildJobFromLegacy_(M.jobs1)
      : s2010_emptyJob_();
    setText(`{{jobs${i}.company}}`, job.company || '');
    setText(`{{jobs${i}.from_yyyy}}`, job.from_yyyy || '');
    setText(`{{jobs${i}.from_mm}}`, job.from_mm || '');
    setText(`{{jobs${i}.from_dd}}`, job.from_dd || '');
    setText(`{{jobs${i}.end_yyyy}}`, job.end_yyyy || '');
    setText(`{{jobs${i}.end_mm}}`, job.end_mm || '');
    setText(`{{jobs${i}.end_dd}}`, job.end_dd || '');
    setText(`{{jobs${i}.to_yyyy}}`, job.to_yyyy || '');
    setText(`{{jobs${i}.to_mm}}`, job.to_mm || '');
    setText(`{{jobs${i}.to_dd}}`, job.to_dd || '');
    setText(`{{jobs${i}.end_text}}`, job.end_text || '');
    setText(`{{jobs${i}.avg_monthly}}`, job.avg_monthly || '');
    setCheck(`{{chk.jobs${i}.severance}}`, !!job.severance);
    setCheck(`{{jobs${i}.chk.severance}}`, !!job.severance);
    setCheck(`{{chk.jobs${i}.kind.employee}}`, !!job.kind.employee);
    setCheck(`{{chk.jobs${i}.kind.part}}`, !!job.kind.part);
    setCheck(`{{chk.jobs${i}.kind.self}}`, !!job.kind.self);
    setCheck(`{{chk.jobs${i}.kind.representative}}`, !!job.kind.representative);
    setCheck(`{{chk.jobs${i}.kind.unemployed}}`, !!job.kind.unemployed);
    setCheck(`{{jobs${i}.chk.kind.employee}}`, !!job.kind.employee);
    setCheck(`{{jobs${i}.chk.kind.part}}`, !!job.kind.part);
    setCheck(`{{jobs${i}.chk.kind.self}}`, !!job.kind.self);
    setCheck(`{{jobs${i}.chk.kind.representative}}`, !!job.kind.representative);
    setCheck(`{{jobs${i}.chk.kind.unemployed}}`, !!job.kind.unemployed);
  }

  // marital
  const marital = Array.isArray(M.marital_events) ? M.marital_events : [];
  for (let i = 1; i <= 4; i++) {
    const row = marital[i - 1] || {};
    const dt = s2010_parseYmdParts_(row.date_iso || '');
    const flags = s2010_maritalFlags_(row.reason || '');
    const reasonText = s2010_normJaSpace_(row.reason || '');
    setText(`{{marital${i}.time}}`, s2010_formatYmdJa_(row.date_iso || ''));
    setText(`{{marital${i}.yyyy}}`, dt.yyyy || '');
    setText(`{{marital${i}.mm}}`, dt.mm || '');
    setText(`{{marital${i}.dd}}`, dt.dd || '');
    setText(`{{marital${i}.partner}}`, s2010_normJaSpace_(row.partner_name || ''));
    setText(`{{marital${i}.reason}}`, reasonText);
    setText(`{{marital${i}.reason_text}}`, reasonText);
    setCheck(`{{chk.marital${i}.kind.marriage}}`, !!flags.marriage);
    setCheck(`{{chk.marital${i}.kind.divorce}}`, !!flags.divorce);
    setCheck(`{{chk.marital${i}.kind.commonlaw}}`, !!flags.commonlaw);
    setCheck(`{{chk.marital${i}.kind.commonlaw_end}}`, !!flags.commonlaw_end);
    setCheck(`{{chk.marital${i}.kind.death}}`, !!flags.death);
  }

  // inheritances
  const inheritances = Array.isArray(M.inheritances) ? M.inheritances : [];
  for (let i = 1; i <= 4; i++) {
    const row = inheritances[i - 1] || {};
    const dt = s2010_parseYmdParts_(row.date_iso || '');
    const statusText = s2010_normJaSpace_(row.status || '');
    const statusFlags = s2010_inheritanceStatusFlagsV2_(statusText, logCtx);
    const legacyFlags = s2010_inheritanceStatusFlags_(statusText);
    setText(`{{inh${i}.decedent_name}}`, s2010_normJaSpace_(row.decedent_name || ''));
    setText(`{{inh${i}.relation}}`, s2010_normJaSpace_(row.relation || ''));
    setText(`{{inh${i}.yyyy}}`, dt.yyyy || '');
    setText(`{{inh${i}.mm}}`, dt.mm || '');
    setText(`{{inh${i}.dd}}`, dt.dd || '');
    setText(`{{inh${i}.status}}`, statusText);
    setText(`{{inh${i}.status_text}}`, statusText);
    setText(`{{inh${i}.name}}`, s2010_normJaSpace_(row.decedent_name || ''));
    setText(`{{inh${i}.date}}`, s2010_formatYmdJa_(row.date_iso || ''));
    setCheck(`{{chk.inh${i}.status.waive}}`, !!statusFlags.waive);
    setCheck(`{{chk.inh${i}.status.divided}}`, !!statusFlags.divided);
    setCheck(`{{chk.inh${i}.status.pending}}`, !!statusFlags.pending);
    setCheck(`{{chk.inh${i}.status.none}}`, !!statusFlags.none);
    setCheck(`{{chk.inh${i}.status.dispose}}`, !!legacyFlags.dispose);
    setCheck(`{{chk.inh${i}.status.other}}`, !!legacyFlags.other);
  }

  // housing
  const housing = M.housing || {};
  const housingType = String(housing.type || '');
  const housingFlags = s2010_housingTypeFlags_(housingType);
  const housingDesc = s2010_pickHousingDesc_(parsed) || housingType || '';
  const ownedRaw = s2010_pickFieldValue_(parsed, /持ち家/);
  const ownOrRent = housing.other_own_or_rent || s2010_pickFieldValue_(parsed, /家屋は所有か賃借/);
  const otherText = s2010_pickFieldValue_(parsed, /^【\s*その他\s*】$/);
  const housingOwned = !!String(ownedRaw || '').trim();
  const ownOrRentFlags = s2010_ownOrRentFlags_(ownOrRent);
  const otherOwnRentFlags = s2010_ownOrRentFlags_(housing.other_own_or_rent || ownOrRent || '');
  const ownedDetailFlags = s2010_housingOwnedDetailFlags_(housing.detail || '');
  const housingOther = s2010_normJaSpace_(housing.notes || otherText || '');

  setText('{{housing.type}}', s2010_normJaSpace_(housingType));
  setText('{{housing.owner}}', s2010_normJaSpace_(housing.owner || ''));
  setText('{{housing.detail}}', s2010_normJaSpace_(housing.detail || ''));
  setText('{{housing.notes}}', s2010_normJaSpace_(housing.notes || ''));
  setText('{{housing.other_name}}', s2010_normJaSpace_(housing.other_name || ''));
  setText('{{housing.other_relation}}', s2010_normJaSpace_(housing.other_relation || ''));
  setText('{{housing.other_own_or_rent}}', s2010_normJaSpace_(housing.other_own_or_rent || ''));
  setText('{{housing.desc}}', s2010_normJaSpace_(housingDesc));
  setText('{{housing.other_owner_name}}', s2010_normJaSpace_(housing.other_name || ''));
  setText('{{housing.other_owner_relation}}', s2010_normJaSpace_(housing.other_relation || ''));
  setText('{{housing.own_or_rent}}', s2010_normJaSpace_(ownOrRent || ''));
  setText('{{housing.other_text}}', s2010_normJaSpace_(otherText || ''));
  setText('{{housing.other}}', housingOther);
  setCheck('{{chk.housing.owned}}', housingOwned);
  setCheck('{{chk.housing.rent}}', !!ownOrRentFlags.rent);
  setCheck('{{chk.housing.own}}', !!ownOrRentFlags.own);
  setCheck('{{chk.housing.type.private_rent}}', !!housingFlags.private_rent);
  setCheck('{{chk.housing.type.public_rent}}', !!housingFlags.public_rent);
  setCheck('{{chk.housing.type.owned}}', !!housingFlags.owned);
  setCheck('{{chk.housing.type.other_person}}', !!housingFlags.other_person);
  setCheck('{{chk.housing.type.other}}', !!housingFlags.other);
  setCheck('{{chk.housing.owned.detached}}', !!ownedDetailFlags.detached);
  setCheck('{{chk.housing.owned.mansion}}', !!ownedDetailFlags.mansion);
  setCheck('{{chk.housing.other_own}}', !!otherOwnRentFlags.own);
  setCheck('{{chk.housing.other_rent}}', !!otherOwnRentFlags.rent);

  return map;
}

function s2010_applyS2010Placeholders_(doc, map) {
  if (!doc || !map) return;
  Object.keys(map)
    .sort()
    .forEach(function (token) {
      s2010_replaceAllEverywhere_(doc, token, map[token]);
    });
}

function s2010_findUnresolvedPlaceholders_(doc) {
  if (!doc) return [];
  const parts = [];
  try { parts.push(doc.getBody().getText()); } catch (_) {}
  try {
    const header = doc.getHeader();
    if (header) parts.push(header.getText());
  } catch (_) {}
  try {
    const footer = doc.getFooter();
    if (footer) parts.push(footer.getText());
  } catch (_) {}
  const text = parts.join('\n');
  const matches = text.match(/{{[^}]+}}/g) || [];
  const uniq = {};
  matches.forEach(function (m) { uniq[m] = true; });
  return Object.keys(uniq);
}

// すでにS2002側に同名があるなら、そちらが使われます（同実装）。
function s2010_replaceAllEverywhere_(doc, token, value) {
  if (!doc) return;
  s2010_replaceAllText_(doc.getBody(), token, value);
  try {
    const header = doc.getHeader();
    if (header) s2010_replaceAllText_(header, token, value);
  } catch (_) {}
  try {
    const footer = doc.getFooter();
    if (footer) s2010_replaceAllText_(footer, token, value);
  } catch (_) {}
}

// すでにS2002側に同名があるなら、そちらが使われます（同実装）。
function s2010_replaceAllText_(element, token, value) {
  if (!element) return;
  const safe = token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const replacement = String(value ?? '').replace(/\\/g, '\\\\').replace(/\$/g, '$$$$');
  element.replaceText(safe, replacement);
}
function s2010_normJaSpace_(s) {
  return String(s || '')
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
function s2010_findFieldValue_(m, labelRegex) {
  const arr = (m && m.fieldsRaw) || [];
  for (let i = arr.length - 1; i >= 0; i--) {
    const f = arr[i];
    if (labelRegex.test(String(f.label || ''))) return String(f.value || '');
  }
  return '';
}
function s2010_splitMulti_(s) {
  return String(s || '')
    .split(/[\r\n、，,・;；]+/)
    .map((v) => s2010_normJaSpace_(v))
    .filter(Boolean);
}
function s2010_hasAny_(arr, re) {
  return (arr || []).some((v) => re.test(v));
}
function s2010_pickOtherText_(raw) {
  const m = String(raw || '').match(/その他[：: ]*([^\n]+)/);
  return m ? s2010_normJaSpace_(m[1]) : '';
}
function s2010_toYMP_(s) {
  const a = s2010_toYMD_(s);
  return { yyyy: a.yyyy, mm: a.mm };
}
function s2010_toYMD_(s) {
  let str = String(s || '').trim();
  const era = str.match(/(令和|平成|昭和)\s*(\d+|元)年(?:\s*(\d+)月)?(?:\s*(\d+)日)?/);
  if (era) {
    const yy = era[2] === '元' ? 1 : parseInt(era[2], 10);
    const y = { 令和: 2018, 平成: 1988, 昭和: 1925 }[era[1]] + yy;
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
function s2010_toNumberText_(s) {
  const n = String(s || '').replace(/[^\d]/g, '');
  return n ? n.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : '';
}
function s2010_toBoolJaLoose_(v) {
  const s = String(v || '').trim();
  return /^(はい|有|あり|有り|true|当|○|チェック)$/i.test(s);
}
function s2010_parseNameKanaValue_(raw) {
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
function s2010_normSpace_(s) {
  return String(s || '')
    .replace(/\s+/g, ' ')
    .trim();
}
function s2010_normPhone_(s) {
  const t = String(s || '').replace(/[^\d]/g, '');
  return t ? t.replace(/(\d{2,4})(\d{2,4})(\d{3,4})/, '$1-$2-$3') : '';
}
function s2010_toIsoBirth_(ja) {
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
function s2010_calcAge_(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  const today = new Date(),
    b = new Date(y, m - 1, d);
  let age = today.getFullYear() - y;
  const md = (today.getMonth() + 1) * 100 + today.getDate(),
    bd = m * 100 + d;
  if (md < bd) age--;
  return age;
}
function s2010_toWareki_(iso) {
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
function s2010_safeHtml_(s) {
  return String(s).replace(
    /[&<>"']/g,
    (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m])
  );
}

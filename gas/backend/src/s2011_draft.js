/** ================== auto_draft_s2011.js ================== **/

// 登録（m1 / m2 とも同じハンドラでOK）
try {
  // forms_ingest_core.js からは (case_id, submission_id) で呼ばれるため、フォームキー別のアダプタを登録
  registerAutoDraft('s2011_income_m1', function (caseId, sid) {
    return handleS2011AutoDraftByKey_('s2011_income_m1', caseId, sid);
  });
  registerAutoDraft('s2011_income_m2', function (caseId, sid) {
    return handleS2011AutoDraftByKey_('s2011_income_m2', caseId, sid);
  });
} catch (_) {}

function run_GenerateS2011DraftBySubmissionId(caseId, submissionId, formKey) {
  if (!caseId || !submissionId) throw new Error('caseId/submissionId required');
  var fk = formKey || 's2011_income_m2';
  return handleS2011AutoDraftByKey_(fk, String(caseId).trim(), String(submissionId).trim());
}

/**
 * ctx 例：
 * {
 *   case_id: "0001",
 *   form_key: "s2011_income_m2",
 *   saved_file_id: "...",             // 今回保存したJSONのDrive File ID
 *   meta: { submission_id, submitted_at, ... }
 * }
 */
function handleS2011AutoDraft_(ctx) {
  if (!ctx || !ctx.case_id || !ctx.form_key) return;
  var sid =
    (ctx.meta && ctx.meta.submission_id) ||
    (ctx.meta && ctx.meta.sid) ||
    (ctx.meta && ctx.meta._meta && ctx.meta._meta.sid) ||
    '';
  return handleS2011AutoDraftByKey_(ctx.form_key, ctx.case_id, sid);
}

function handleS2011AutoDraftByKey_(formKey, caseId, sid) {
  if (!caseId || !sid) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20 * 1000)) {
    Utilities.sleep(1500);
    if (!lock.tryLock(10 * 1000)) {
      try {
        Logger.log('[S2011] WARN: lock not acquired (caseId=%s, sid=%s)', caseId, sid);
      } catch (_) {}
      return;
    }
  }
  try {
    const sourceFile = findS2011JsonFile_(caseId, formKey, sid);
    if (!sourceFile)
      throw new Error('S2011 source json not found (formKey=' + formKey + ', sid=' + sid + ')');

    const raw = JSON.parse(sourceFile.getBlob().getDataAsString('UTF-8'));
    const model = raw && typeof raw === 'object' && raw.model ? raw.model : raw;
    const submittedAt = normalizeDateTime_(
      (raw && raw.meta && raw.meta.submitted_at) ||
        (raw && raw._meta && raw._meta.submitted_at) ||
        Utilities.formatDate(sourceFile.getLastUpdated(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm')
    );

    const agg = loadS2011Agg_(caseId) || {};
    const periodKey =
      formKey === 's2011_income_m1' ? 'm1' : formKey === 's2011_income_m2' ? 'm2' : '';
    if (!periodKey) return;

    const stamped = JSON.parse(JSON.stringify(model || {}));
    stamped._meta = stamped._meta || {};
    stamped._meta.sid = String(sid);
    stamped._meta.submitted_at = submittedAt;

    if (shouldOverwrite_(agg[periodKey], stamped)) {
      agg[periodKey] = stamped;
    }

    agg.status = agg.status || {};
    agg.status.m1 = !!agg.m1;
    agg.status.m2 = !!agg.m2;
    agg.status.complete = !!(agg.m1 && agg.m2);
    agg.status.updated_at = new Date().toISOString();

    const draft = ensureS2011DraftSheet_(caseId, agg.draft);
    agg.draft = draft;
    renameS2011DraftFile_(draft.sheet_id, caseId, sid);
  const newHash = renderS2011DraftSheet_(draft.sheet_id, agg);
  if (newHash) agg.draft.last_hash = newHash;

    saveS2011Agg_(caseId, agg);
    updateCaseS2011Status_(caseId, {
      m1_received: !!agg.m1,
      m2_received: !!agg.m2,
      complete: agg.status.complete,
      sheet_id: draft.sheet_id,
    });

  const url = SpreadsheetApp.openById(draft.sheet_id).getUrl();
  try {
    Logger.log('[S2011] draft url: %s', url);
  } catch (_) {}
  return { sheetId: draft.sheet_id, url: url, complete: agg.status.complete };
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

/* ========== helpers (必要に応じて既存実装に合わせて置換) ========== */

// 集約ファイルの場所・命名は運用に合わせて変更
function loadS2011Agg_(caseId) {
  const f = findCaseFile_(caseId, 's2011_income_agg.json'); // 既存: ケース直下検索
  if (!f) return {};
  try {
    return JSON.parse(f.getBlob().getDataAsString('UTF-8'));
  } catch (_) {
    return {};
  }
}

function saveS2011Agg_(caseId, obj) {
  const json = JSON.stringify(obj, null, 2);
  let f = findCaseFile_(caseId, 's2011_income_agg.json');
  const folder = getCaseFolder_(caseId);
  const blob = Utilities.newBlob(json, 'application/json', 's2011_income_agg.json');
  if (!f) {
    folder.createFile(blob);
  } else {
    const mime = f.getMimeType && f.getMimeType();
    if (mime && /^application\/vnd\.google-apps\./.test(String(mime))) {
      try {
        f.setTrashed(true);
      } catch (_) {}
      folder.createFile(blob);
    } else {
      f.setContent(json);
    }
  }
}

function attachMeta_(model, ctx) {
  const m = JSON.parse(JSON.stringify(model || {}));
  m._meta = m._meta || {};
  m._meta.sid = (ctx.meta && ctx.meta.submission_id) || '';
  m._meta.submitted_at = normalizeDateTime_((ctx.meta && ctx.meta.submitted_at) || '');
  return m;
}

function shouldOverwrite_(existing, incoming) {
  if (!incoming) return false;
  if (!existing) return true;
  const a = Date.parse((existing._meta && existing._meta.submitted_at) || '') || 0;
  const b = Date.parse((incoming._meta && incoming._meta.submitted_at) || '') || Date.now();
  return b >= a; // 新しい方を採用
}

function normalizeDateTime_(s) {
  // "2025年10月22日 08時05分" → ISO
  s = String(s || '')
    .replace(/年|\/|\./g, '-')
    .replace(/月/g, '-')
    .replace(/日/g, '')
    .replace(/時/g, ':')
    .replace(/分/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?([.+-]\d{2}:\d{2}|Z)?$/.test(s)) {
    return s;
  }
  // 例: "2025-10-22 08:05" → "2025-10-22T08:05:00+09:00"
  if (/^\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}$/.test(s)) {
    return s.replace(' ', 'T') + ':00+09:00';
  }
  return new Date().toISOString();
}

function ensureS2011DraftSheet_(caseId, prev) {
  if (prev && prev.sheet_id) {
    try {
      SpreadsheetApp.openById(prev.sheet_id);
      return prev;
    } catch (_) {}
  }
  const caseFolder = getCaseFolder_(caseId);
  const draftsFolder =
    typeof getOrCreateSubfolder_ === 'function'
      ? getOrCreateSubfolder_(caseFolder, 'drafts')
      : caseFolder;
  const props = PropertiesService.getScriptProperties();
  const tplIdTrimmed = String(
    (props && props.getProperty('S2011_TEMPLATE_GSHEET_ID')) ||
      (props && props.getProperty('S2011_TEMPLATE_SSID')) ||
      ''
  ).trim();
  if (tplIdTrimmed) {
    const ssId = DriveApp.getFileById(tplIdTrimmed)
      .makeCopy(`S2011_家計収支_${caseId}`, draftsFolder)
      .getId();
    return { sheet_id: ssId, doc_id: ssId };
  }
  throw new Error('S2011_TEMPLATE_GSHEET_ID 未設定。テンプレからのコピーに失敗しました。');
}

function buildS2011Skeleton_(ss) {
  const sh = ss.getActiveSheet();
  sh.clear();
  sh.getRange(1, 1).setValue('S2011 家計収支（ドラフト）');
  sh.getRange(2, 1, 1, 3).setValues([['項目', 'm-2', 'm-1']]);
  const rows = [
    '前月繰越',
    '給与(申立人)',
    '給与(配偶者)',
    '給与(同居家族その他)',
    '自営(申立人)',
    '自営(その他)',
    '年金(申立人)',
    '年金(その他)',
    '雇用保険',
    '生活保護',
    '児童手当',
    '援助',
    '借入れ',
    'その他収入',
    '---',
    '住居費',
    '駐車場',
    '食費',
    '嗜好品',
    '外食',
    '電気',
    'ガス',
    '水道',
    '携帯(合計)',
    '通信その他',
    '日用品',
    '新聞',
    '国保/国年',
    '任意保険',
    'ガソリン',
    '交通',
    '医療',
    '被服',
    '教育',
    '交際',
    '娯楽',
    '債務返済1',
    '債務返済2',
    'その他1',
    'その他2',
  ];
  rows.forEach((r, i) => sh.getRange(3 + i, 1).setValue(r));
  sh.autoResizeColumns(1, 3);
}

function getTopLeftCell_(cell) {
  if (!cell) return null;
  if (cell.isPartOfMerge()) {
    const m = cell.getMergedRanges();
    if (m && m.length) {
      return m[0].getCell(1, 1);
    }
  }
  return cell;
}

function getRowLabelText_(sh, row) {
  let cellB = sh.getRange(row, 2);
  let topB = getTopLeftCell_(cellB);
  let v = topB ? topB.getDisplayValue() : '';
  if (v) return v;
  let cellA = sh.getRange(row, 1);
  let topA = getTopLeftCell_(cellA);
  v = topA ? topA.getDisplayValue() : '';
  return v || '';
}

function listOtherRows_(sh) {
  const rows = [];
  const last = Math.max(1, sh.getLastRow());
  for (let r = 1; r <= last; r++) {
    const raw = getRowLabelText_(sh, r);
    if (!raw) continue;
    if (normalizeNoParen_(raw) === 'その他') rows.push(r);
  }
  return rows;
}

function getExpenseStartRow_(sh, idx) {
  idx = idx || indexSheetByBcol_(sh);
  let r = idx.get('住居費') || idx.get(normalizeKeepParen_('住居費'));
  if (r) return r;
  const last = sh.getLastRow();
  for (let i = 1; i <= last; i++) {
    const raw = getRowLabelText_(sh, i);
    if (normalizeNoParen_(raw) === '住居費') return i;
  }
  let lastTotal = null;
  for (let i = 1; i <= last; i++) {
    const raw = getRowLabelText_(sh, i);
    if (normalizeNoParen_(raw) === '合計') lastTotal = i;
  }
  return (lastTotal || 1) + 1;
}

function listOtherRowsExpense_(sh, idx) {
  const start = getExpenseStartRow_(sh, idx);
  const rows = [];
  const last = sh.getLastRow();
  for (let r = start; r <= last; r++) {
    const raw = getRowLabelText_(sh, r);
    if (!raw) continue;
    if (normalizeNoParen_(raw) === 'その他') rows.push(r);
  }
  if (!rows.length) {
    for (let r = 1; r <= last; r++) {
      const raw = getRowLabelText_(sh, r);
      if (!raw) continue;
      if (normalizeNoParen_(raw) === 'その他') rows.push(r);
    }
    return rows.slice(-2);
  }
  return rows;
}

function canonKeysFromRow_(sh, row) {
  const raw = getRowLabelText_(sh, row);
  if (!raw) return [];
  const k1 = normalizeNoParen_(raw);
  const k2 = normalizeKeepParen_(raw);
  const k2c = k2.replace(/□.*$/, '');
  return Array.from(
    new Set([k1, k2, k2c].map((k) => (S2011_LABEL_ALIAS_MAP_[k] || k || '').trim()))
  ).filter(Boolean);
}

function getCheckedGlyph_() {
  const p = PropertiesService.getScriptProperties().getProperty('S2011_CHECKED_GLYPH') || '';
  return p.trim() || '■';
}

function getUncheckedGlyph_() {
  return '□';
}

function normalizePaymentMethod_(s) {
  const t = String(s || '').trim().replace(/\s+/g, '');
  if (/^(口座)?振[り込込]$/.test(t) || /銀行振込/.test(t)) return 'transfer';
  if (/^現金(支給)?$/.test(t)) return 'cash';
  return '';
}

function markPaymentMethodInB_(sh, row, rawValue) {
  const tl = getTopLeftCell_(sh.getRange(row, 2));
  if (!tl) return;
  let s = String(tl.getDisplayValue() || '');
  s = s.replace(/（\s*(振込|現金)\s*）/g, '');
  const mode = normalizePaymentMethod_(rawValue);
  const C = getCheckedGlyph_();
  const U = getUncheckedGlyph_();
  const hasChoices = /[□■☑]\s*振込/.test(s) || /[□■☑]\s*現金/.test(s);
  let next = s;
  const setTransfer = () => {
    next = next.replace(/[□■☑]\s*振込/g, C + '振込');
    next = next.replace(/[□■☑]\s*現金/g, U + '現金');
  };
  const setCash = () => {
    next = next.replace(/[□■☑]\s*振込/g, U + '振込');
    next = next.replace(/[□■☑]\s*現金/g, C + '現金');
  };
  const setUnknown = () => {
    next = next.replace(/[□■☑]\s*振込/g, U + '振込');
    next = next.replace(/[□■☑]\s*現金/g, U + '現金');
  };
  if (mode === 'transfer') setTransfer();
  else if (mode === 'cash') setCash();
  else setUnknown();
  if (!hasChoices) {
    if (mode === 'transfer') next = (s + '  ' + C + '振込 ' + U + '現金').trim();
    else if (mode === 'cash') next = (s + '  ' + U + '振込 ' + C + '現金').trim();
    else next = (s + '  ' + U + '振込 ' + U + '現金').trim();
  }
  if (next !== s) tl.setValue(next);
  const e = sh.getRange(row, 5);
  if (String(e.getValue() || '') !== '') e.clearContent();
}

/** 文字列正規化（括弧内削除版） */
function normalizeNoParen_(s) {
  s = String(s || '');
  try {
    s = s.normalize('NFKC');
  } catch (_) {}
  return s
    .replace(/[【】\[\]]/g, '')
    .replace(/（.*?）/g, '')
    .replace(/[()（）]/g, '')
    .replace(/[：:]/g, ':')
    .replace(/\s|　/g, '')
    .toLowerCase();
}

/** 文字列正規化（括弧内テキスト保持） */
function normalizeKeepParen_(s) {
  s = String(s || '');
  try {
    s = s.normalize('NFKC');
  } catch (_) {}
  return s
    .replace(/[【】\[\]]/g, '')
    .replace(/[()（）]/g, '')
    .replace(/[：:]/g, ':')
    .replace(/\s|　/g, '')
    .toLowerCase();
}

const S2011_LABEL_ALIAS_MAP_ = (function () {
  const pairs = [
    ['前月繰越', '前月からの繰越'],
    ['翌月繰越', '翌月への繰越'],
    ['携帯(合計)', '携帯電話料金'],
    ['通信その他', 'その他通話料・通信料・CATV等'],
    ['給与その他同居家族分', '給与分'],
    ['自営収入申立人以外分', '自営収入分'],
    ['年金申立人以外分', '年金分'],
    ['携帯料金', '携帯電話料金'],
    ['通信費', 'その他通話料・通信料・CATV等'],
    ['債務返済実額1', '債務返済1'],
    ['債務返済実額2', '債務返済2'],
    ['その他1', 'その他支出1'],
    ['その他2', 'その他支出2'],
    ['その他１', 'その他支出1'],
    ['その他２', 'その他支出2'],
  ];
  const map = {};
  pairs.forEach(([from, to]) => {
    map[normalizeKeepParen_(from)] = normalizeKeepParen_(to);
  });
  return map;
})();

/** "住居費:種別" → { base: "住居費", sub: "種別" } */
function splitBaseAndSub_(k) {
  const a = String(k || '').split(':');
  const rawBase = a[0] || '';
  let base = rawBase.replace(/(金額|実額|費|料金)(\d+)?$/, '');
  if (!base) base = rawBase;
  const sub = a.slice(1).join(':') || '';
  return { base, sub };
}

/** B列のインデックス作成（正規化キー → 行番号） */
function indexSheetByBcol_(sh) {
  const last = Math.max(1, sh.getLastRow());
  const idx = new Map();
  const buckets = new Map();
  for (let i = 1; i <= last; i++) {
    const raw = getRowLabelText_(sh, i);
    if (!raw) continue;
    const key1 = normalizeNoParen_(raw);
    const key2 = normalizeKeepParen_(raw);
    const key2Clean = key2.replace(/□.*$/, '');
    const canons = Array.from(
      new Set([key1, key2, key2Clean].map((k) => (S2011_LABEL_ALIAS_MAP_[k] || k || '').trim()))
    ).filter(Boolean);
    if (!canons.length) continue;
    const canon = canons[0];
    if (!buckets.has(canon)) buckets.set(canon, []);
    buckets.get(canon).push(i);
  }
  for (const [canon, rows] of buckets.entries()) {
    if (!rows.length) continue;
    if (!idx.has(canon)) idx.set(canon, rows[0]);
    rows.forEach((row, i) => {
      idx.set(canon + String(i + 1), row);
    });
  }
  return idx;
}

function resolveRowByBaseNumberedFirst_(sh, idx, base) {
  if (!idx || !base) return null;
  let b = base.replace(/□.*$/, '');
  try {
    b = b.normalize('NFKC');
  } catch (_) {}
  let row = idx.get(b) || idx.get(base);
  if (row) return row;
  const m = String(b).match(/^(.*?)(\d+)$/);
  if (m) {
    const stem = m[1];
    const ordNum = Number(m[2]);
    if (Number.isFinite(ordNum) && ordNum > 0) {
      const direct = idx.get(stem + String(ordNum));
      if (direct) return direct;
      const last = Math.max(1, sh.getLastRow());
      const bucket = [];
      for (let r = 1; r <= last; r++) {
        const raw = getRowLabelText_(sh, r);
        if (!raw) continue;
        let k = normalizeNoParen_(raw);
        if (k === stem) bucket.push(r);
      }
      if (bucket.length >= ordNum) return bucket[ordNum - 1];
    }
  }
  let bestKey = null;
  let bestLen = -1;
  for (const key of idx.keys()) {
    let k = key.replace(/□.*$/, '');
    try {
      k = k.normalize('NFKC');
    } catch (_) {}
    if (!(b.startsWith(k) || k.startsWith(b))) continue;
    if (k.length > bestLen) {
      bestLen = k.length;
      bestKey = key;
    }
  }
  return bestKey ? idx.get(bestKey) : null;
}

/** 括弧差し込み（マージ対応） */
function putParenTextIntoB_(sh, row, text) {
  if (!text) return;
  let cell = sh.getRange(row, 2);
  if (cell.isPartOfMerge()) {
    const merged = cell.getMergedRanges();
    if (merged && merged.length) {
      cell = merged[0].getCell(1, 1);
    }
  }
  let s = String(cell.getDisplayValue() || '')
    .replace(/\(/g, '（')
    .replace(/\)/g, '）');
  const next = /（.*?）/.test(s) ? s.replace(/（.*?）/, '（' + text + '）') : s + '（' + text + '）';
  if (next !== s) cell.setValue(next);
}

/** "￥12,345" → 12345（空は空のまま） */
function coerceNumber_(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  let s = String(v || '').trim();
  try {
    s = s.normalize('NFKC');
  } catch (_) {}
  s = s
    .replace(/[￥¥,円\s]/g, '')
    .replace(/[－ー—–]/g, '-')
    .replace(/^△|^▲/, '-')
    .replace(/^\((.*)\)$/, '-$1');
  if (!s) return '';
  const n = Number(s);
  return Number.isFinite(n) ? n : '';
}

/** JSON → ラベル辞書 */
function buildLabelMapFromJson_(json) {
  if (!json) return {};
  if (Array.isArray(json.fields_indexed)) {
    const map = {};
    for (const it of json.fields_indexed) {
      const label = it && (it.label || it.key || it.name);
      const value = it && (it.value != null ? it.value : it.text);
      if (label) map[String(label)] = value == null ? '' : value;
    }
    return map;
  }
  if (json.fields_indexed && typeof json.fields_indexed === 'object') return json.fields_indexed;
  if (json.fields_map && typeof json.fields_map === 'object') return json.fields_map;
  if (json.model && typeof json.model === 'object') {
    const m = json.model;
    if (m.fields_indexed) {
      if (Array.isArray(m.fields_indexed)) {
        const map = {};
        for (const it of m.fields_indexed) {
          const label = (it && (it.label || it.key || it.name)) || '';
          const value = it && (it.value != null ? it.value : it.text);
          if (label) map[String(label)] = value == null ? '' : value;
        }
        return map;
      }
      if (typeof m.fields_indexed === 'object') return m.fields_indexed;
    }
    if (m.fields && typeof m.fields === 'object') return m.fields;
  }
  if (Array.isArray(json.fieldsRaw)) {
    const map = {};
    for (const it of json.fieldsRaw) {
      const label = (it && (it.label || it.key || it.name)) || '';
      const value = it && (it.value != null ? it.value : it.text);
      if (label) map[String(label)] = value == null ? '' : value;
    }
    return map;
  }
  const f = json.fields || json.data;
  if (Array.isArray(f)) {
    const map = {};
    for (const it of f) {
      const label = (it && (it.label || it.key || it.name)) || '';
      const value = it && (it.value != null ? it.value : it.text);
      if (label) map[String(label)] = value == null ? '' : value;
    }
    return map;
  }
  if (f && typeof f === 'object') return f;
  return {};
}

/** ドラフトの親（ケース）フォルダ取得（drafts/ 経由にも対応） */
function getParentFolder_(fileId) {
  const file = DriveApp.getFileById(fileId);
  const parents = file.getParents();
  if (!parents.hasNext()) throw new Error('Parent folder not found for draft: ' + fileId);
  let folder = parents.next();
  if (String(folder.getName() || '').toLowerCase() === 'drafts') {
    const upper = folder.getParents();
    if (upper.hasNext()) folder = upper.next();
  }
  return folder;
}

/** ドラフト名 → { caseId, sid } */
function parseCaseAndSidFromDraftName_(name) {
  const m = String(name || '').match(/^S2011_([^_]+)_draft_([0-9A-Za-z-]+)$/);
  return m ? { caseId: m[1], sid: m[2] } : { caseId: '', sid: '' };
}

/** ケースフォルダから原本 JSON を読み出す */
function loadRawJsonFromCase_(caseFolder, formKey, sid) {
  if (!caseFolder) throw new Error('caseFolder is required');
  const normalizedKey = String(formKey || '').trim();
  if (!normalizedKey) throw new Error('formKey is required');
  const normalizedSid = String(sid || '').replace(/[^\d]/g, '');
  if (normalizedSid) {
    const fname = normalizedKey + '__' + normalizedSid + '.json';
    const it = caseFolder.getFilesByName(fname);
    if (it.hasNext()) {
      return JSON.parse(it.next().getBlob().getDataAsString('UTF-8'));
    }
  }
  const files = caseFolder.getFiles();
  let pick = null;
  let pickTs = 0;
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (name.indexOf(normalizedKey + '__') === 0 && /\.json$/i.test(name)) {
      const t = file.getLastUpdated().getTime();
      if (t > pickTs) {
        pick = file;
        pickTs = t;
      }
    }
  }
  if (!pick) {
    throw new Error('Case JSON not found: formKey=' + normalizedKey + ' sid=' + normalizedSid);
  }
  return JSON.parse(pick.getBlob().getDataAsString('UTF-8'));
}

/** m2/m1 1セクション分を書き込み */
function writeValuesForSection_(sh, labelsDict, monthCol, idx) {
  if (!labelsDict) return 0;
  if (Array.isArray(labelsDict)) {
    labelsDict = buildLabelMapFromJson_({ fields_indexed: labelsDict });
  }
  idx = idx || indexSheetByBcol_(sh);
  const TEXT_SUBS =
    /^(種別|内訳|品名|車名義|名義|名義人|何人分か|対象者|受給対象者|対象児|対象児童|誰からか|誰の分か|相手|宛先|支払先|内容|用途|使途|摘要|備考|受取方法|メモ)$/;
  const BASES_TO_PAREN =
    /^(住居費|駐車場代|嗜好品代|教育費|ガソリン代|援助|借入れ|児童.*手当|債務返済実額\d|雇用保険|給与(?:分)?|年金(?:分)?|自営収入(?:分)?|その他支出(?:\d+)?)$/;
  let wrote = 0;
  Object.keys(labelsDict).forEach((rawKey) => {
    const value = labelsDict[rawKey];
    const normKey = normalizeKeepParen_(rawKey);
    const { base, sub } = splitBaseAndSub_(normKey);
    const rowKey = S2011_LABEL_ALIAS_MAP_[base] || base;
    if (!rowKey) return;
    let row = null;
    if (/^その他支出\d+$/.test(rowKey)) {
      row = idx.get(rowKey) || null;
    }
    const mOrd = /^その他(\d+)$/.exec(rowKey);
    if (mOrd) {
      const ord = Number(mOrd[1]);
      if (Number.isFinite(ord) && ord > 0) {
        const others = listOtherRowsExpense_(sh, idx);
        row = others[ord - 1] || null;
      }
    }
    if (!row) {
      row = resolveRowByBaseNumberedFirst_(sh, idx, rowKey);
    }
    if (!row) return;
    if (sub) {
      if (sub === '受取方法') {
        markPaymentMethodInB_(sh, row, value);
        wrote++;
        return;
      }
      if (TEXT_SUBS.test(sub)) {
        if (value != null && value !== '') {
          putParenTextIntoB_(sh, row, String(value));
        }
        const eCell = sh.getRange(row, 5);
        if (String(eCell.getValue() || '') !== '') eCell.clearContent();
        wrote++;
        return;
      }
      if (sub === '金額') {
        sh.getRange(row, monthCol).setValue(coerceNumber_(value));
        wrote++;
        return;
      }
      const str = value == null ? '' : String(value);
      const num = coerceNumber_(str);
      if (!num && BASES_TO_PAREN.test(rowKey)) {
        if (str) putParenTextIntoB_(sh, row, str);
        const eCell = sh.getRange(row, 5);
        if (String(eCell.getValue() || '') !== '') eCell.clearContent();
      } else {
        sh.getRange(row, 5).setValue(str);
      }
      wrote++;
      return;
    }
    if (/(金額|実額|費|料金)(\d+)?$/.test(normKey)) {
      sh.getRange(row, monthCol).setValue(coerceNumber_(value));
    } else {
      const str = value == null ? '' : String(value);
      const num = coerceNumber_(str);
      if (!num && BASES_TO_PAREN.test(rowKey)) {
        if (str) putParenTextIntoB_(sh, row, str);
        const eCell = sh.getRange(row, 5);
        if (String(eCell.getValue() || '') !== '') eCell.clearContent();
      } else {
        sh.getRange(row, 5).setValue(str);
      }
    }
    wrote++;
  });
  return wrote;
}

/** セクション情報から raw ラベル辞書を取得（無ければケースJSON読み直し） */
function loadLabelDictFallback_(caseFolder, section, defaultFormKey, draftSid) {
  if (section && typeof section === 'object') {
    if (section.fields_indexed || section.fields_map) {
      return buildLabelMapFromJson_(section);
    }
    if (section.raw_fields && typeof section.raw_fields === 'object') {
      return section.raw_fields;
    }
  }
  if (!caseFolder) return null;
  const meta = (section && section._meta) || {};
  const formKey = String(meta.form_key || defaultFormKey || '').trim();
  const sid = String(meta.sid || meta.submission_id || draftSid || '').replace(/[^\d]/g, '') || '';
  if (!formKey) return null;
  const raw = loadRawJsonFromCase_(caseFolder, formKey, sid);
  return buildLabelMapFromJson_(raw);
}

// どのタブに書き込むか決める（プロパティ優先 → ヒューリスティック）
function pickS2011DataSheet_(ss) {
  const props = PropertiesService.getScriptProperties();
  const explicit = String((props && props.getProperty('S2011_TEMPLATE_SHEET_NAME')) || '').trim();
  if (explicit) {
    const sh = ss.getSheetByName(explicit);
    if (sh) return sh;
  }
  const keys = [
    /前月(からの)?繰越/,
    /翌月(への)?繰越/,
    /給与/,
    /自営/,
    /年金/,
    /児童.*手当/,
    /援助/,
    /借入/,
    /住居費/,
    /教育費/,
    /携帯(電話)?料金/,
    /医療費/,
    /交際費/,
    /娯楽費/,
  ];
  let best = null;
  let bestScore = -1;
  ss.getSheets().forEach((sheet) => {
    try {
      const rows = Math.max(sheet.getLastRow(), 150);
      const colB = sheet
        .getRange(1, 2, rows, 1)
        .getDisplayValues()
        .map((r) => String(r[0] || ''))
        .join('\n');
      const colA = sheet
        .getRange(1, 1, rows, 1)
        .getDisplayValues()
        .map((r) => String(r[0] || ''))
        .join('\n');
      const blob = colA + '\n' + colB;
      const score = keys.reduce((acc, rx) => acc + (rx.test(blob) ? 1 : 0), 0);
      if (score > bestScore) {
        bestScore = score;
        best = sheet;
      }
    } catch (_) {}
  });
  if (!best || bestScore <= 0) throw new Error('S2011: データタブが特定できません（スコア=0）');
  return best;
}

function renderS2011DraftSheet_(sheetOrId, agg) {
  let id = sheetOrId;
  if (!id) throw new Error('renderS2011DraftSheet_: sheetId is empty');
  if (typeof id !== 'string') {
    if (id && typeof id.getId === 'function') {
      id = id.getId();
    } else {
      throw new Error('renderS2011DraftSheet_: invalid arg type=' + typeof id);
    }
  }
  id = id.replace(/^https?:\/\/.*\/d\/([a-zA-Z0-9-_]+)(?:\/|$).*/, '$1');
  if (!/^[a-zA-Z0-9-_]{20,}$/.test(id)) {
    throw new Error('renderS2011DraftSheet_: looks not an ID -> ' + id);
  }
  Logger.log('[S2011] render: openById=%s', id);
  const ss = SpreadsheetApp.openById(id);
  const sh = pickS2011DataSheet_(ss);

  const file = DriveApp.getFileById(id);
  const parsedName = parseCaseAndSidFromDraftName_(file.getName());
  const draftSid = String(parsedName.sid || '').replace(/[^\d]/g, '');
  const caseFolder = getParentFolder_(id);

  const COL_M2 = 3;
  const COL_M1 = 4;

  const m2Labels =
    loadLabelDictFallback_(caseFolder, agg && agg.m2, 's2011_income_m2', draftSid) || {};
  const m1Labels =
    agg && agg.m1
      ? loadLabelDictFallback_(caseFolder, agg.m1, 's2011_income_m1', draftSid) || {}
      : null;

  const baseIndex = indexSheetByBcol_(sh);
  let wrote = 0;
  wrote += writeValuesForSection_(sh, m2Labels, COL_M2, baseIndex);
  wrote += writeValuesForSection_(sh, m1Labels, COL_M1, baseIndex);

  Logger.log(
    '[S2011] sheet=%s case=%s rows=%s wrote=%s',
    sh.getName(),
    parsedName.caseId || '',
    sh.getLastRow(),
    wrote
  );

  if (!wrote) {
    const limit = Math.min(30, sh.getLastRow());
    const bvals = Array.from({ length: limit }, (_, i) => getRowLabelText_(sh, i + 1));
    const sampleM2 = Object.keys(m2Labels || {}).slice(0, 20);
    const sampleM1 = Object.keys(m1Labels || {}).slice(0, 20);
    Logger.log(
      JSON.stringify({
        error: 'S2011: 書込み対象0件',
        sheet: { name: sh.getName(), rows: sh.getLastRow() },
        bcol_samples_keep: bvals.slice(0, 10).map((x) => normalizeKeepParen_(x)),
        json_m2_keys_keep: sampleM2.map((x) => normalizeKeepParen_(x)),
        json_m1_keys_keep: sampleM1.map((x) => normalizeKeepParen_(x)),
      })
    );
    throw new Error('S2011: 書込み対象0件（データタブ/ラベル不一致の可能性）');
  }

  try {
    const idx = indexSheetByBcol_(sh);
    const prevKeys = ['前月からの繰越', '前月繰越'].map((s) => normalizeKeepParen_(s));
    const nextKeys = ['翌月への繰越', '翌月繰越'].map((s) => normalizeKeepParen_(s));
    let rowPrevCarry = null;
    let rowNextCarry = null;
    for (const key of prevKeys) {
      const hit = idx.get(key);
      if (hit) {
        rowPrevCarry = hit;
        break;
      }
    }
    for (const key of nextKeys) {
      const hit = idx.get(key);
      if (hit) {
        rowNextCarry = hit;
        break;
      }
    }
    if (rowPrevCarry && rowNextCarry) {
      sh.getRange(rowPrevCarry, COL_M1).setFormula('=C' + rowNextCarry);
    }
  } catch (_) {}

  sh.getRange(1, COL_M2, sh.getMaxRows(), 2).setNumberFormat('#,##0');
  sh.getRange(1, 5).setValue(agg.status && agg.status.complete ? '【完了】' : '【片側未提出】');

  const hashSrc = JSON.stringify({ m1: agg.m1 || {}, m2: agg.m2 || {} });
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, hashSrc)
  );
}

function updateCaseS2011Status_(caseId, state) {
  // 例：cases_forms シートに upsert。既存の status_api.js 等に合わせて差し替え
  try {
    if (typeof appendSubmissionLog_ === 'function') {
      appendSubmissionLog_({
        case_id: caseId,
        form_key: 's2011_income',
        note: JSON.stringify(state),
      });
    }
  } catch (_) {}
}

function findS2011JsonFile_(caseId, formKey, sidRaw) {
  const sid = String(sidRaw || '').replace(/[^\d]/g, '');
  const exact = String(formKey || '').trim() + '__' + sid + '.json';

  let f = null;
  try {
    f = findCaseFile_(caseId, exact);
  } catch (_) {}
  if (f) return f;

  const folder = getCaseFolder_(caseId);
  const prefix = String(formKey || '').trim() + '__';
  const pred = (name) => name === exact || name.indexOf(prefix) === 0;

  const files = folder.getFiles();
  while (files.hasNext()) {
    const candidate = files.next();
    if (pred(candidate.getName()) && s2011_fileHasSid_(candidate, sid)) return candidate;
  }

  const hit = findCaseFileRecursive_(
    folder,
    (file) => pred(file.getName()) && s2011_fileHasSid_(file, sid)
  );
  if (hit) return hit;

  const latest = findCaseFileRecursive_(folder, (file) => pred(file.getName()), true);
  if (latest) {
    try {
      Logger.log(
        '[S2011] WARN: fallback to latest prefix file without SID match: %s',
        latest.getName()
      );
    } catch (_) {}
  }
  return latest || null;
}

function s2011_fileHasSid_(file, sid) {
  try {
    const raw = JSON.parse(file.getBlob().getDataAsString('UTF-8')) || {};
    const pointer = [
      raw && raw.meta && raw.meta.submission_id,
      raw && raw._meta && raw._meta.sid,
      raw && raw.submission_id,
    ];
    return pointer.some((v) => String(v || '') === sid);
  } catch (_) {
    return false;
  }
}

function findCaseFileRecursive_(folder, predicate, pickLatest) {
  let best = null;
  let bestTs = 0;

  const fit = folder.getFiles();
  while (fit.hasNext()) {
    const file = fit.next();
    if (predicate(file)) {
      if (!pickLatest) return file;
      const ts = file.getLastUpdated().getTime();
      if (ts > bestTs) {
        bestTs = ts;
        best = file;
      }
    }
  }

  const dit = folder.getFolders();
  while (dit.hasNext()) {
    const sub = dit.next();
    const found = findCaseFileRecursive_(sub, predicate, pickLatest);
    if (found) {
      if (!pickLatest) return found;
      const ts = found.getLastUpdated().getTime();
      if (ts > bestTs) {
        bestTs = ts;
        best = found;
      }
    }
  }
  return best;
}

function debug_list_s2011_files(caseId) {
  const folder = getCaseFolder_(caseId);
  const out = [];
  (function walk(dir, path) {
    const files = dir.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      if (/^s2011_income_/.test(name)) {
        out.push([path + '/' + name, new Date(file.getLastUpdated()).toISOString()]);
      }
    }
    const subs = dir.getFolders();
    while (subs.hasNext()) {
      const sub = subs.next();
      walk(sub, path + '/' + sub.getName());
    }
  })(folder, '');
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

/** ケースID→案件フォルダを s2006 と同じ流れで解決 */
function getCaseFolder_(caseId) {
  var raw = String(caseId || '').trim();
  if (!/^\d+$/.test(raw)) throw new Error('getCaseFolder_: invalid caseId=' + raw);
  if (typeof resolveCaseByCaseId_ !== 'function' || typeof ensureCaseFolderId_ !== 'function') {
    throw new Error(
      'getCaseFolder_: resolver not available (resolveCaseByCaseId_/ensureCaseFolderId_)'
    );
  }
  var info = resolveCaseByCaseId_(raw);
  if (!info) throw new Error('getCaseFolder_: unknown case_id=' + raw);
  var folderId = ensureCaseFolderId_(info);
  if (!folderId) throw new Error('getCaseFolder_: ensureCaseFolderId_ failed for ' + raw);
  return DriveApp.getFolderById(folderId);
}

/** ケース直下で名前完全一致のファイルを1件返す（なければ null） */
function findCaseFile_(caseId, name) {
  var folder = getCaseFolder_(caseId);
  var it = folder.getFilesByName(String(name || ''));
  return it.hasNext() ? it.next() : null;
}

function renameS2011DraftFile_(sheetId, caseId, sid) {
  try {
    var cidMatch = String(caseId || '').match(/\d+/);
    var cidDigits = cidMatch ? cidMatch[0] : '';
    var cidFormatted = cidDigits ? String(cidDigits).padStart(4, '0') : String(caseId || '');
    var sidDigits = String(sid || '').replace(/[^\d]/g, '');
    var name = 'S2011_' + (cidFormatted || '0000') + '_draft_' + sidDigits;
    DriveApp.getFileById(sheetId).setName(name);
  } catch (_) {}
}

/**
 * ケースフォルダ内の既存 JSON（例: s2011_income_m2__49832654.json）だけを使って、
 * テンプレから S2011_0001_draft_49832654 を「作り直し」ます。
 *
 * 使い方（GASエディタから実行 or clasp run）:
 *   run_RebuildS2011DraftFromCaseJson('0001', '49832654', 's2011_income_m2', true)
 */
function run_RebuildS2011DraftFromCaseJson(caseId, submissionId, formKey, overwrite) {
  formKey = formKey || 's2011_income_m2';
  overwrite = overwrite !== false; // 省略時は上書き削除して作り直す

  const props = PropertiesService.getScriptProperties();
  const rootId = props.getProperty('DRIVE_ROOT_ID'); // BASのルート
  if (!rootId) throw new Error('ScriptProperty DRIVE_ROOT_ID が未設定です。');

  const caseFolder = findCaseFolder_(rootId, caseId);
  if (!caseFolder) throw new Error('ケースフォルダが見つかりません: ' + caseId);
  const draftsFolder = getOrCreateSubfolder_(caseFolder, 'drafts');

  const jsonName = formKey + '__' + submissionId + '.json';
  const jsonFile = findFileInFolderByName_(caseFolder, jsonName);
  if (!jsonFile) throw new Error('ケース内に JSON が見つかりません: ' + jsonName);

  const jsonText = jsonFile.getBlob().getDataAsString('UTF-8');
  const data = JSON.parse(jsonText);

  // agg を最小形で再構築（m2のみ）
  const agg = {
    m2: data,
    m1: undefined,
    status: { m2: true, m1: false, complete: false, updated_at: new Date().toISOString() },
    draft: {},
  };

  // 既存ドラフトがあれば処理方針に従う
  const draftName = 'S2011_' + caseId + '_draft_' + submissionId;
  const existingDraft =
    findFileInFolderByName_(draftsFolder, draftName) ||
    findFileInFolderByName_(caseFolder, draftName);
  if (existingDraft) {
    if (overwrite) {
      existingDraft.setTrashed(true); // ごみ箱へ（完全削除はしない）
    } else {
      // 上書き方式：既存シートへそのままレンダリング
      const id = existingDraft.getId();
      renderS2011DraftSheet_(id, agg);
      Logger.log('[S2011] re-rendered into existing draft: %s', draftName);
      // agg も保存更新
      upsertJson_(
        caseFolder,
        's2011_income_agg.json',
        Object.assign(loadAggIfAny_(caseFolder), agg, { draft: { sheet_id: id } })
      );
      return;
    }
  }

  // テンプレからコピーして新規作成 → レンダリング
  const templateId =
    props.getProperty('S2011_TEMPLATE_GSHEET_ID') || props.getProperty('S2011_TEMPLATE_SSID');
  if (!templateId)
    throw new Error(
      'ScriptProperty S2011_TEMPLATE_GSHEET_ID (or S2011_TEMPLATE_SSID) が未設定です。'
    );
  const template = SpreadsheetApp.openById(templateId);
  const draft = template.copy(draftName);
  renderS2011DraftSheet_(draft.getId(), agg);

  // テンプレのコピーは My Drive 直下にできるため、drafts へ移動
  DriveApp.getFileById(draft.getId()).moveTo(draftsFolder);

  // agg を保存（シートIDと共に）
  const mergedAgg = Object.assign(loadAggIfAny_(caseFolder), agg, {
    draft: { sheet_id: draft.getId() },
  });
  upsertJson_(caseFolder, 's2011_income_agg.json', mergedAgg);

  Logger.log('[S2011] rebuilt draft: %s', draftName);
}

function setS2011Props_() {
  const p = PropertiesService.getScriptProperties();
  p.setProperty('S2011_TEMPLATE_GSHEET_ID', '1EMGuXuPbxuCSgv-YW2aM8fH4pyuKlbMZTBnRFasvsS8');
  p.setProperty('S2011_TEMPLATE_SHEET_NAME', '家計収支表'); // ←タブの実名に置き換え
}

/** ケースフォルダ探索（まず cases/ 配下→無ければルート直下で名称一致） */
function findCaseFolder_(rootId, caseId) {
  const root = DriveApp.getFolderById(rootId);
  // /cases/<caseId>
  const cases = getOrNull_(root.getFoldersByName('cases'));
  if (cases && cases.hasNext()) {
    const f = getOrNull_(cases.next().getFoldersByName(caseId));
    if (f && f.hasNext()) return f.next();
  }
  // ルート直下に <caseId>
  const direct = getOrNull_(root.getFoldersByName(caseId));
  if (direct && direct.hasNext()) return direct.next();
  return null;
}

/** フォルダ内で完全一致のファイルを探す（サブフォルダは見ない） */
function findFileInFolderByName_(folder, name) {
  const it = folder.getFilesByName(name);
  return it.hasNext() ? it.next() : null;
}

function getOrNull_(iter) {
  try {
    return iter;
  } catch (_) {
    return null;
  }
}

/** 既存の agg があれば読み込む（なければ {}） */
function loadAggIfAny_(caseFolder) {
  const f = findFileInFolderByName_(caseFolder, 's2011_income_agg.json');
  if (!f) return {};
  try {
    const txt = f.getBlob().getDataAsString('UTF-8');
    return JSON.parse(txt) || {};
  } catch (_) {
    return {};
  }
}

/** JSON を upsert（存在すれば置換、無ければ新規） */
function upsertJson_(folder, name, obj) {
  const f = findFileInFolderByName_(folder, name);
  const blob = Utilities.newBlob(JSON.stringify(obj, null, 2), 'application/json', name);
  if (f) {
    f.setTrashed(true);
  }
  folder.createFile(blob);
}

function run_GenerateS2011DraftBySubmissionId_dev() {
  return run_GenerateS2011DraftBySubmissionId('0001', '49832654', 's2011_income_m2');
}

function run_GenerateS2011Draft_forceNew_dev() {
  const caseId = '0001';
  const sid = '49832654';
  const formKey = 's2011_income_m2';
  const agg = loadS2011Agg_(caseId) || {};
  if (agg.draft && agg.draft.sheet_id) {
    try {
      DriveApp.getFileById(agg.draft.sheet_id).setTrashed(true);
    } catch (_) {}
    delete agg.draft;
    saveS2011Agg_(caseId, agg);
  }
  return run_GenerateS2011DraftBySubmissionId(caseId, sid, formKey);
}

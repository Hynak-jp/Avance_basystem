/** ====== 設定 ====== **
 * 必要な Script Properties:
 * - BAS_MASTER_SPREADSHEET_ID : cases台帳シートのID
 * - S2002_TEMPLATE_GDOC_ID   : 差し込み用gdocテンプレのファイルID
 */
const PROP_S2002 = PropertiesService.getScriptProperties();
const S2002_SPREADSHEET_ID = PROP_S2002.getProperty('BAS_MASTER_SPREADSHEET_ID') || ''; // cases台帳
const S2002_TPL_GDOC_ID = PROP_S2002.getProperty('S2002_TEMPLATE_GDOC_ID') || ''; // gdocテンプレ
const S2002_LABEL_TO_PROCESS = 'FormAttach/ToProcess';
const S2002_LABEL_PROCESSED = 'FormAttach/Processed';

/** ====== パブリック・エントリ ====== **/

/**
 * 受信箱から S2002 通知メールを取り込み → JSON保存 → S2002ドラフト生成（gdoc）
 * ラベル: FormAttach/ToProcess → 処理後に FormAttach/Processed を付与
 */
function run_ProcessInbox_S2002() {
  // レガシー実行は共通ルーターへ委譲（Queue → ルーターの一括処理）
  try {
    Logger.log('[S2002] delegate to common router (run_ProcessInbox_AllForms)');
  } catch (_) {}
  try {
    if (typeof run_ProcessInbox_AllForms === 'function') {
      run_ProcessInbox_AllForms();
    }
  } catch (e) {
    try {
      Logger.log('[S2002] delegation error: %s', (e && e.stack) || e);
    } catch (_) {}
  }
}

/** ====== 既存JSON→ドラフト生成ユーティリティ ====== **/

/**
 * 直近の s2002_userform__*.json を読み込んで S2002ドラフトを生成
 * @param {string} caseId - 例: "0045"
 */

function debug_GenerateS2002_for_0001() {
  // 仕様上の正式表記は4桁ゼロ埋め（例: 0001）。内部では '1' でも正規化しますが、明示的に '0001' を使います。
  return run_GenerateS2002DraftByCaseId('0001');
}

function run_GenerateS2002DraftByCaseId(caseId) {
  const caseInfo = resolveCaseByCaseId_(caseId);
  if (!caseInfo) throw new Error(`Unknown case_id: ${caseId}`);
  // folderId が未設定/URL等なら補正
  caseInfo.folderId = ensureCaseFolderId_(caseInfo);

  const parsed = loadLatestSubmissionJson_(caseInfo.folderId, 's2002_userform__');
  if (!parsed) throw new Error(`No S2002 JSON found under case folder: ${caseInfo.folderId}`);

  const draft = generateS2002Draft_(caseInfo, parsed);
  updateCasesRow_(caseId, {
    status: 'draft',
    last_activity: new Date(),
    last_draft_url: draft.draftUrl,
  });
  try {
    Logger.log('[S2002] draft created: %s', draft.draftUrl);
  } catch (_) {}
  return draft;
}

/**
 * caseフォルダIDを直接指定してS2002ドラフトを生成（テスト用）
 * @param {string} caseFolderId
 */
function run_GenerateS2002DraftByFolder(caseFolderId) {
  const parsed = loadLatestSubmissionJson_(caseFolderId, 's2002_userform__');
  if (!parsed) throw new Error(`No S2002 JSON found under case folder: ${caseFolderId}`);

  // cases行を逆引き
  const ss = SpreadsheetApp.openById(S2002_SPREADSHEET_ID);
  const sh = ss.getSheetByName('cases');
  const vals = sh.getDataRange().getValues();
  const header = vals[0];
  const idxMap = bs_toIndexMap_(header);
  const idxFolder = idxMap.folder_id;
  const idxCaseId = idxMap.case_id;
  const idxCaseKey = idxMap.case_key;
  const idxLineId = idxMap.line_id;
  if (!(idxFolder >= 0 && idxCaseId >= 0)) {
    throw new Error('cases sheet is missing folder_id/case_id columns');
  }
  let caseInfo = null;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][idxFolder]) === String(caseFolderId)) {
      caseInfo = {
        caseId: vals[i][idxCaseId],
        caseKey: idxCaseKey >= 0 ? vals[i][idxCaseKey] : '',
        lineId: idxLineId >= 0 ? vals[i][idxLineId] : '',
        folderId: caseFolderId,
        rowIndex: i + 1,
      };
      break;
    }
  }
  if (!caseInfo) throw new Error(`Folder not found in cases: ${caseFolderId}`);

  const draft = generateS2002Draft_(caseInfo, parsed);
  updateCasesRow_(caseInfo.caseId, {
    status: 'draft',
    last_activity: new Date(),
    last_draft_url: draft.draftUrl,
  });
  return draft;
}

/**
 * ケース直下の S2002 JSON（meta.form_key === 's2002_userform'）のうち最新を読み込む
 * 返り値は { meta, fieldsRaw, model }（メール経路と同スキーマ）
 */
function normalizeFolderId_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  const m = s.match(/[-\w]{25,}/); // Drive IDのざっくりパターン
  return m ? m[0] : s;
}

function loadLatestSubmissionJson_(caseFolderId, _filePrefixIgnored) {
  const id = normalizeFolderId_(caseFolderId);
  if (!id) throw new Error('[S2002] cases.folderId is empty.');
  let parent;
  try {
    parent = DriveApp.getFolderById(id);
  } catch (e) {
    throw new Error('[S2002] invalid folderId: ' + id + ' :: ' + (e.message || e));
  }
  // ケース直下の .json を走査し、meta.form_key === 's2002_userform' の最新を選ぶ
  const candidates = [];
  const it = parent.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const name = f.getName && f.getName();
    if (!name || !/\.json$/i.test(name)) continue;
    try {
      const j = JSON.parse(f.getBlob().getDataAsString('UTF-8'));
      const formKey = (j && j.meta && j.meta.form_key) || j.form_key || '';
      if (String(formKey).trim() === 's2002_userform') {
        candidates.push({ file: f, json: j, t: f.getLastUpdated().getTime() });
      }
    } catch (_) {}
  }
  if (!candidates.length) return null;
  candidates.sort((a, b) => b.t - a.t);
  const latest = candidates[0].json || {};

  // そのまま model があれば採用
  if (latest.model && latest.meta) return latest;

  // さまざまな形の fields を配列 [{label, value}] に正規化
  const fieldsArr = normalizeFieldsArray_(latest);
  if (!fieldsArr) throw new Error('Invalid submission JSON shape (no fields/model)');
  return { meta: latest.meta || {}, fieldsRaw: fieldsArr, model: mapFieldsToModel_(fieldsArr) };
}

/**
 * ケース直下の prefix*.json のうち「最初の intake」を選ぶ。
 * 優先: intake__<submissionId>.json の submissionId 小さい順 → 作成日時が古い順
 */
function loadEarliestJsonByPrefix_(caseFolder, prefix) {
  if (!caseFolder || !prefix) return null;
  let folder = caseFolder;
  if (typeof caseFolder === 'string') {
    const id = normalizeFolderId_(caseFolder);
    if (!id) return null;
    try {
      folder = DriveApp.getFolderById(id);
    } catch (_) {
      return null;
    }
  }
  const it = folder.getFiles();
  const picks = [];
  const escaped = String(prefix).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const nameRe = new RegExp('^' + escaped + '(\\d+).*\\.json$', 'i');
  while (it.hasNext()) {
    const f = it.next();
    const name = f.getName && f.getName();
    if (!name || !/\.json$/i.test(name)) continue;
    if (String(name).indexOf(prefix) !== 0) continue;
    const m = String(name).match(nameRe);
    const sid = m ? parseInt(m[1], 10) : NaN;
    let created = 0;
    try {
      created = f.getDateCreated().getTime();
    } catch (_) {
      created = f.getLastUpdated().getTime();
    }
    picks.push({ file: f, sid: Number.isFinite(sid) ? sid : null, created: created });
  }
  if (!picks.length) return null;
  picks.sort(function (a, b) {
    const aHas = a.sid != null;
    const bHas = b.sid != null;
    if (aHas && bHas) {
      if (a.sid !== b.sid) return a.sid - b.sid;
    } else if (aHas && !bHas) {
      return -1;
    } else if (!aHas && bHas) {
      return 1;
    }
    return a.created - b.created;
  });
  const pick = picks[0];
  try {
    return { file: pick.file, json: JSON.parse(pick.file.getBlob().getDataAsString('UTF-8')) };
  } catch (_) {
    return null;
  }
}

function s2002_extractIntakeBaseInfo_(json) {
  if (!json || typeof json !== 'object') return {};
  const name = s2002_getByLabel_(
    json,
    [/申立人氏名/, /氏名/, /名前/],
    [['app', 'name'], ['model', 'app', 'name'], ['applicant', 'name'], ['name'], ['applicant_name']]
  );
  const birth = s2002_getByLabel_(
    json,
    [/生年月日/],
    [['app', 'birth'], ['model', 'app', 'birth'], ['birth'], ['birth_iso']]
  );
  let zip = s2002_getByLabel_(
    json,
    [/郵便/, /〒/],
    [['addr', 'postal'], ['model', 'addr', 'postal'], ['postal'], ['zip'], ['zipcode']]
  );
  const addressPaths = [
    ['addr', 'full'],
    ['model', 'addr', 'full'],
    ['address'],
    ['addr', 'address'],
    ['addr', 'residence'],
  ];
  const addressExact = s2002_getByLabel_(json, [/^【?\s*住居所\s*】?$/], []);
  const addressFallback = s2002_getByLabel_(json, [/^(?!.*異なる).*住所/], []);
  let address = addressExact || addressFallback;
  if (!address) {
    const directAddress = s2002_pickDirect_(json, addressPaths);
    if (directAddress && !/異なる/.test(directAddress)) address = directAddress;
  }
  if (address) {
    const p = parseAddressLine_(address);
    if (!zip && p.postal) zip = p.postal;
    address = s2002_stripZipFromAddress_(address);
  }
  const sameAsRaw = s2002_getByLabel_(
    json,
    [/(?=.*住民票)(?=.*記載)/, /住民票記載のとおり/, /住所.*住民票/],
    []
  );
  const same_as_resident = toBoolJa_(sameAsRaw);
  const phone = s2002_getByLabel_(
    json,
    [/電話番号/, /電話/, /連絡先/],
    [['app', 'phone'], ['model', 'app', 'phone'], ['phone'], ['tel'], ['telephone']]
  );
  return {
    name: String(name || '').trim(),
    birth: String(birth || '').trim(),
    zip: String(zip || '').trim(),
    address: String(address || '').trim(),
    same_as_resident: same_as_resident,
    phone: String(phone || '').trim(),
  };
}

function s2002_getByLabel_(json, regexes, directPaths) {
  const list = Array.isArray(regexes) ? regexes : [regexes];
  const direct = s2002_pickDirect_(json, directPaths);
  if (direct) return direct;

  const sources = [];
  if (json.fields_indexed) sources.push(json.fields_indexed);
  if (json.model && json.model.fields_indexed) sources.push(json.model.fields_indexed);
  if (json.fieldsRaw) sources.push(json.fieldsRaw);
  if (json.fields) sources.push(json.fields);
  if (json.model && json.model.fields) sources.push(json.model.fields);
  if (json.data) sources.push(json.data);

  for (let i = 0; i < sources.length; i++) {
    const hit = s2002_pickFromSource_(sources[i], list);
    if (hit) return hit;
  }
  return '';
}

function s2002_pickDirect_(json, paths) {
  const list = Array.isArray(paths) ? paths : [];
  for (let i = 0; i < list.length; i++) {
    const path = list[i];
    let cur = json;
    for (let j = 0; j < path.length; j++) {
      if (!cur || typeof cur !== 'object') {
        cur = null;
        break;
      }
      cur = cur[path[j]];
    }
    if (cur != null && String(cur).trim()) return String(cur).trim();
  }
  return '';
}

function s2002_pickFromSource_(src, regexes) {
  if (!src) return '';
  if (Array.isArray(src)) {
    for (let i = 0; i < src.length; i++) {
      const it = src[i] || {};
      const label = it.label || it.key || it.name || '';
      const value = it.value != null ? it.value : it.text != null ? it.text : it.answer;
      if (s2002_labelMatch_(label, regexes) && String(value || '').trim()) {
        return String(value || '').trim();
      }
    }
    return '';
  }
  if (typeof src === 'object') {
    const keys = Object.keys(src);
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      const v = src[k];
      const hasObj = v && typeof v === 'object';
      const label = hasObj && v.label ? v.label : k;
      let val = v;
      if (hasObj) {
        if (v.value != null) val = v.value;
        else if (v.text != null) val = v.text;
        else if (v.answer != null) val = v.answer;
      }
      if (
        (s2002_labelMatch_(k, regexes) || s2002_labelMatch_(label, regexes)) &&
        String(val || '').trim()
      ) {
        return String(val || '').trim();
      }
    }
  }
  return '';
}

function s2002_labelMatch_(label, regexes) {
  const raw = String(label || '').trim();
  const inner = raw.replace(/^【\s*|\s*】$/g, '').trim();
  for (let i = 0; i < regexes.length; i++) {
    const re = regexes[i];
    if (re.test(raw) || re.test(inner)) return true;
  }
  return false;
}

function s2002_stripPostal_(v) {
  let s = String(v || '').trim();
  if (!s) return '';
  s = s.replace(/[〒\s()（）]/g, '');
  s = s.replace(/[０-９]/g, function (ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0xfee0);
  });
  return s;
}

function s2002_stripZipFromAddress_(v) {
  let s = String(v || '').trim();
  if (!s) return '';
  s = s.replace(/〒/g, ' ');
  s = s.replace(/[0-9０-９]{3}[-－―‐ー]?[0-9０-９]{4}/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

function s2002_normalizeZip_(v) {
  const s = s2002_stripPostal_(v);
  if (!s) return '';
  const m = s.match(/(\d{3})-?(\d{4})/);
  if (!m) return '';
  return m[1] + '-' + m[2];
}

function s2002_buildResidenceBlock_(opts) {
  const zip = s2002_normalizeZip_(opts && opts.zip);
  const address = String((opts && opts.address) || '').trim();
  if (zip && address) return '〒(' + zip + ') ' + address;
  if (zip) return '〒(' + zip + ')';
  return address;
}

function s2002_appendResidenceAfterZipIfMissing_(body, zip, address) {
  const z = s2002_normalizeZip_(zip);
  const addr = String(address || '').trim();
  if (!body || !z || !addr) return;
  const text = body.getText();
  const textNorm = text.replace(/\s+/g, '');
  const addrNorm = addr.replace(/\s+/g, '');
  if (text.indexOf(addr) !== -1 || (addrNorm && textNorm.indexOf(addrNorm) !== -1)) return;
  const token = '〒(' + z + ')';
  if (text.indexOf(token) === -1) return;
  // テンプレに住所プレースホルダが無い場合の保険（〒だけ残る不具合対策）
  const safe = token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  body.replaceText(safe, escapeReplaceTextValue_(token + ' ' + addr));
}

/**
 * s2002 JSON に含まれる fields を配列 [{label, value}] に正規化
 * - fieldsRaw が配列ならそのまま
 * - fields が配列ならそのまま
 * - fields がオブジェクト {label:value} なら配列へ変換（label は '【...】' 形式に）
 */
function normalizeFieldsArray_(json) {
  if (!json || typeof json !== 'object') return null;
  const fr = json.fieldsRaw;
  if (Array.isArray(fr)) return fr;
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

/** ケース直下にある S2002 JSON を列挙（デバッグ用） */
function debug_ListS2002LogsByCaseId(caseId) {
  const info = resolveCaseByCaseId_(caseId);
  if (!info) throw new Error('Unknown caseId: ' + caseId);
  const folderId = ensureCaseFolderId_(info);
  const parent = DriveApp.getFolderById(folderId);
  const out = [];
  // ケース直下のみを探索し、各 JSON の formKey を推定して表示
  let it = parent.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const n = f.getName && f.getName();
    if (!n || !/\.json$/i.test(n)) continue;
    try {
      const j = JSON.parse(f.getBlob().getDataAsString('UTF-8'));
      const formKey = (j && j.meta && j.meta.form_key) || j.form_key || '';
      out.push({ where: 'case_root', name: n, id: f.getId(), formKey: String(formKey).trim() });
    } catch (_) {}
  }
  Logger.log('[S2002] logs for case %s: %s', caseId, JSON.stringify(out));
  return out;
}

/**
 * ケースフォルダの健全性を点検（folderId, フォルダ名, 直下ファイル一覧）
 * 使い方: debug_InspectCaseFolder('0001')
 */
function debug_InspectCaseFolder(caseId) {
  const info = resolveCaseByCaseId_(caseId);
  if (!info) throw new Error('Unknown caseId: ' + caseId);
  const folderId = ensureCaseFolderId_(info);
  const folder = DriveApp.getFolderById(folderId);
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    files.push({ name: f.getName(), id: f.getId(), updatedAt: f.getLastUpdated() });
  }
  const res = {
    caseId: normalizeCaseIdString_(caseId),
    resolvedFolderId: folderId,
    folderName: folder.getName(),
    files,
  };
  Logger.log('[S2002] inspect case %s: %s', caseId, JSON.stringify(res));
  return res;
}

/**
 * cases シートの folderId を手動修正する（Drive で確認した正しいIDを書き戻す）。
 * 使い方: debug_SetCaseFolderId('0001', 'xxxxxxxxxxxxxxxxxxxx')
 */
function debug_SetCaseFolderId(caseId, folderId) {
  updateCasesRow_(caseId, { folder_id: String(folderId || '').trim() });
  const info = resolveCaseByCaseId_(caseId);
  Logger.log('[S2002] folderId updated: caseId=%s folderId=%s', caseId, info && info.folderId);
}

/**
 * 任意のファイルIDの親フォルダを表示（Drive 上での実体確認に使用）
 * 使い方: debug_ShowParents('<FILE_ID>')
 */
function debug_ShowParents(fileId) {
  const f = DriveApp.getFileById(String(fileId || '').trim());
  const parents = [];
  const it = f.getParents();
  while (it.hasNext()) {
    const p = it.next();
    parents.push({ id: p.getId(), name: p.getName() });
  }
  const res = { fileId: f.getId(), name: f.getName(), parents };
  Logger.log('[S2002] parents: %s', JSON.stringify(res));
  return res;
}

/** ====== デバッグ ====== **/

function debug_ParseSample_() {
  const sampleSubject =
    '[#FM-BAS] S2002 破産手続開始申立書 提出 submission_id:48029097 〔2025年09月07日 18時37分〕';
  const sampleBody = `
==== META START ====
form_name: S2002 破産手続開始申立書
form_key: s2002_userform
submission_id: 48029097
line_id: Uc13df94016ee50eb9dd5552bffbe6624
case_id: 0020
submitted_at: 2025年09月07日 18時37分
secret: FM-BAS
==== META END ====

==== FIELDS START ====
【名前（ふりがな）】
テスト 太郎（てすと たろう）
【旧姓（ふりがな）】
ダミー 太郎（だみー たろう）
【メールアドレス】
dummy@address.com
【連絡先】
000-0000-0000
【生年月日】
2000年01月01日
【国籍】
日本
【住居所は住民票記載のとおりですか？】
はい
【住居所】
001-0000 北海道 札幌市北区テスト町0-0-0 000号
【住居所（住民票と異なる）】

【個人事業者か】
いいえ
【法人の代表者ですか？】
はい
【その法人の破産申立て】
同時申立て
【生活保護受給世帯に属しているか】
いいえ
【申立前７年以内に個人再生の認可決定が確定したことがありますか？】
はい
【申立前７年以内に破産免責決定が確定したことがありますか？】
いいえ
==== FIELDS END ====
`;
  const parsed = parseFormMail_(sampleSubject, sampleBody);
  Logger.log(JSON.stringify(parsed, null, 2));
}

/** ====== 取り込み・パーサ ====== **/

function parseFormMail_(subject, body) {
  const meta = parseMetaBlock_(body);
  // 件名から submission_id の予備抽出（subject変更にも対応）
  meta.submission_id =
    meta.submission_id ||
    (subject.match(/submission_id:(\d+)/) || subject.match(/提出\s+(\d{6,})/))?.[1] ||
    '';

  const fields = parseFieldsBlock_(body);
  const out = mapFieldsToModel_(fields);
  return { meta, fieldsRaw: fields, model: out };
}

function parseMetaBlock_(body) {
  const m = body.match(/====\s*META START\s*====([\s\S]*?)====\s*META END\s*====/);
  if (!m) throw new Error('META block not found');
  const lines = m[1]
    .split(/\r?\n/)
    .map((s) => s.trim())
    .filter(Boolean);
  const meta = {};
  lines.forEach((line) => {
    const mm = line.match(/^([^:]+):\s*(.*)$/);
    if (mm) meta[mm[1].trim()] = mm[2].trim();
  });
  return meta;
}

function parseFieldsBlock_(body) {
  const m = body.match(/====\s*FIELDS START\s*====([\s\S]*?)====\s*FIELDS END\s*====/);
  if (!m) throw new Error('FIELDS block not found');
  const block = m[1];

  // 形式: 「【ラベル】\n値(複数行OK)\n【ラベル】...」。行頭の空白（全角含む）にも耐性を持たせる
  const parts = block
    .split(/\r?\n(?=[ \t\u3000]*【)/)
    .map((s) => s.replace(/^[ \t\u3000]+/, '')) // 行頭の全角/半角空白を除去
    .map((s) => s.trimEnd())
    .filter(Boolean);
  const out = [];
  parts.forEach((part) => {
    // 行頭空白を許容してラベル抽出
    const lm = part.match(/^[ \t\u3000]*【(.+?)】\s*([\s\S]*)$/);
    if (!lm) return;
    const label = `【${lm[1].trim()}】`;
    let value = (lm[2] || '')
      .replace(/^\r?\n+/, '') // 先頭の改行を除去
      .replace(/^[ \t\u3000]+/, ''); // 先頭の空白を除去

    // 値が“純粋なラベル行”だけなら空値とみなす（次ラベルの取り込みを防止）
    if (/^[ \t\u3000]*【[^】]+】[ \t\u3000]*$/.test(value)) value = '';

    out.push({ label, value: String(value || '').trim() });
  });
  return out; // [{label:'【...】', value:'...'}]
}

/** ====== マッピング ====== **/

function mapFieldsToModel_(fields) {
  const out = { app: {}, addr: {}, ref: {} };
  fields.forEach(({ label, value }) => applyFieldLine_(label, value, out));

  // 後処理（正規化）
  if (out.app && out.app.birth) {
    const iso = toIsoBirth_(out.app.birth);
    out.app.birth_iso = iso;
    out.app.age = iso ? calcAge_(iso) : '';
    const w = toWareki_(iso);
    if (w) out.app.birth_wareki = w;
  }
  // mapFieldsToModel_ の後処理に追加（fields.forEach の後）
  if (out.addr) {
    // 住民票どおりのときの通常パース
    if (out.addr.full) {
      const p = parseAddressLine_(out.addr.full);
      out.addr.postal = out.addr.postal || p.postal || '';
      out.addr.pref_city = out.addr.pref_city || p.pref_city || '';
      out.addr.street = out.addr.street || p.street || '';
    }

    // 「住民票と異なる」なら alt_full を必ず埋める
    if (out.addr.same_as_resident === false) {
      // ① 取りこぼし再取得（必ず value を使う！）
      if (!out.addr.alt_full || /^【/.test(out.addr.alt_full)) {
        const rec = fields.find((f) => /^【\s*住居所（住民票と異なる[^】]*）\s*】$/.test(f.label));
        if (rec && rec.value && rec.value.trim()) {
          out.addr.alt_full = sanitizeAltAddressValue_(rec.value);
        }
      }
      // ② それでも空なら full を代入（空出力回避の最終手段）— ただし full がラベル見えなら禁止
      const looksLikeLabel = (s) => /^【[^】]*】$/.test(String(s || '').trim());
      if (!out.addr.alt_full && out.addr.full && !looksLikeLabel(out.addr.full)) {
        out.addr.alt_full = normJaSpace_(out.addr.full);
      }
    } else if (out.addr.same_as_resident === true && !out.addr.alt_full) {
      out.addr.alt_full = ''; // 住民票どおりなら空に寄せる
    }
  }

  return out;
}
function normJaSpace_(s) {
  return String(s || '')
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function applyFieldLine_(label, value, out) {
  // 追加: 漢字のみの氏名フィールドに対応
  if (/^【?名前】?$/.test(label)) {
    out.app.name = normSpace_(value);
    return;
  }
  if (/^【?旧姓】?$/.test(label)) {
    out.app.maiden_name = normSpace_(value);
    return;
  }
  if (/^【?名前（ふりがな）】?/.test(label)) {
    const { name, kana } = parseNameKanaValue_(value);
    out.app.name = normSpace_(name);
    out.app.kana = kana;
    return;
  }
  if (/^【?旧姓（ふりがな）】?/.test(label)) {
    const { name, kana } = parseNameKanaValue_(value);
    out.app.maiden_name = normSpace_(name);
    out.app.maiden_kana = kana;
    return;
  }
  if (/^【?メールアドレス】?/.test(label)) {
    out.app.email = value.trim();
    return;
  }
  if (/^【?連絡先】?/.test(label)) {
    out.app.phone = normPhone_(value);
    return;
  }
  if (/^【?生年月日】?/.test(label)) {
    out.app.birth = value.trim();
    return;
  }
  if (/^【?国籍】?/.test(label)) {
    out.app.nationality = value.trim();
    return;
  }

  if (/^【?住居所は住民票記載のとおりですか？】?/.test(label)) {
    out.addr.same_as_resident = toBoolJa_(value);
    return;
  }
  // applyFieldLine_ の住所まわりをこの順に
  if (/^【\s*住居所（住民票と異なる[^】]*）\s*】$/.test(label)) {
    const v = sanitizeAltAddressValue_(value);
    if (v) out.addr.alt_full = v; // 値が空やラベルのみなら代入しない
    return;
  }
  // フォールバック: 「住民票」かつ「異なる」を含むラベル全般
  if (/^【[^】]*住民票[^】]*異なる[^】]*】$/.test(label)) {
    out.addr.alt_full = sanitizeAltAddressValue_(value);
    return;
  }
  if (/^【\s*住居所\s*】$/.test(label)) {
    out.addr.full = normJaSpace_(value);
    return;
  }

  if (/^【?個人事業者か】?/.test(label)) {
    const v = String(value || '').trim();
    if (!v) return;
    if (v === 'いいえ') {
      out.ref.self_employed = 'none';
    } else if (/[6６]/.test(v)) {
      out.ref.self_employed = 'past6m';
    } else {
      out.ref.self_employed = 'current';
    }
    return;
  }
  if (/^【?法人の代表者ですか？】?/.test(label)) {
    out.ref.corp_representative = toBoolJa_(value);
    return;
  }
  if (/^【?その法人の破産申立て】?/.test(label)) {
    out.ref.corp_bankruptcy_status = value.trim(); // 申立済み / 同時申立て / 申立て予定 / 予定なし
    return;
  }
  if (/^【?生活保護受給世帯に属しているか】?/.test(label)) {
    out.ref.welfare = toBoolJa_(value);
    return;
  }
  if (/^【?申立前７年以内に個人再生.*】?/.test(label)) {
    out.ref.pr_rehab_within7y = toBoolJa_(value);
    return;
  }
  if (/^【?申立前７年以内に破産免責決定.*】?/.test(label)) {
    out.ref.bankruptcy_discharge_within7y = toBoolJa_(value);
    return;
  }
}

/** ====== ユーティリティ（正規化） ====== **/

// 置換: 名前/かなの抽出を堅牢化
function parseNameKanaValue_(raw) {
  const s = String(raw || '').trim();
  const isKana = (t) => /^[\p{Script=Hiragana}\p{Script=Katakana}\u30FC\s・･ｰﾞﾟ\-]+$/u.test(t);
  const hasKanji = (t) => /[\p{Script=Han}]/u.test(t);

  // パターン1: 「A（B）」全角/半角カッコ
  let m = s.match(/^(.+?)\s*[（(]\s*([^)）]+?)\s*[)）]\s*$/);
  if (m) {
    let left = m[1].trim(),
      right = m[2].trim();
    // 期待: 「漢字（かな）」→ そのまま
    if (hasKanji(left) || !isKana(left)) return { name: left, kana: right };
    // 逆転: 「かな（漢字）」→ スワップ
    if (hasKanji(right) && isKana(left)) return { name: right, kana: left };
    return { name: left, kana: right };
  }

  // パターン2: 区切り記号「/｜|」
  m = s.match(/^(.+?)\s*[/｜\|]\s*(.+)$/);
  if (m) {
    let a = m[1].trim(),
      b = m[2].trim();
    if (hasKanji(a) || !isKana(a)) return { name: a, kana: b };
    if (hasKanji(b) && isKana(a)) return { name: b, kana: a };
    return { name: a, kana: b };
  }

  // 単独値: かなだけなら name は空、kana のみ（氏名は別ラベルで漢字を取得）
  return isKana(s) ? { name: '', kana: s } : { name: s, kana: '' };
}

function normSpace_(s) {
  return String(s || '')
    .replace(/\s+/g, ' ')
    .trim();
}
function normPhone_(s) {
  const t = String(s || '').replace(/[^\d]/g, '');
  if (!t) return '';
  // ざっくり3分割
  return t.replace(/(\d{2,4})(\d{2,4})(\d{3,4})/, '$1-$2-$3');
}
function toBoolJa_(v) {
  const s = String(v || '').trim();
  if (!s) return null;
  return s === 'はい';
}
function toIsoBirth_(ja) {
  // 例: 2000年01月01日 / 2000/1/1 / 2000.1.1 など
  let s = String(ja || '')
    .trim()
    .replace(/生\s*$/, '')
    .replace(/[年月]/g, '-')
    .replace(/[日]/g, '')
    .replace(/[.\/\s]/g, '-')
    .replace(/-+/g, '-');
  s = s.replace(/[^0-9-]/g, '');
  const m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return '';
  const y = +m[1],
    mo = ('0' + m[2]).slice(-2),
    d = ('0' + m[3]).slice(-2);
  return `${y}-${mo}-${d}`;
}
function calcAge_(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  const today = new Date();
  const b = new Date(y, m - 1, d);
  let age = today.getFullYear() - y;
  const md = (today.getMonth() + 1) * 100 + today.getDate();
  const bd = m * 100 + d;
  if (md < bd) age--;
  return age;
}
// 置換: 和暦対応を拡張
function toWareki_(iso) {
  if (!iso) return null;
  const m = String(iso).match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return null;
  const y = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10);
  const d0 = parseInt(m[3], 10);
  const d = new Date(y, mo - 1, d0);
  if (d.getFullYear() !== y || d.getMonth() + 1 !== mo || d.getDate() !== d0) return null;
  const eras = [
    { gengo: '令和', start: new Date(2019, 4, 1), offset: 2018 },
    { gengo: '平成', start: new Date(1989, 0, 8), offset: 1988 },
    { gengo: '昭和', start: new Date(1926, 11, 25), offset: 1925 },
  ];
  for (const e of eras) {
    if (d >= e.start)
      return {
        gengo: e.gengo,
        yy: d.getFullYear() - e.offset,
        mm: d.getMonth() + 1,
        dd: d.getDate(),
      };
  }
  return null;
}

function parseAddressLine_(oneLine) {
  const s = String(oneLine || '').trim();
  if (!s) return { postal: '', pref_city: '', street: '' };
  const pm = s.match(/(\d{3}-?\d{4})/);
  const postal = pm ? pm[1].replace(/-/, '').replace(/(\d{3})(\d{4})/, '$1-$2') : '';
  const rest = pm ? s.replace(pm[1], '').trim() : s;
  const m = rest.match(/^(.{2,7}[都道府県].*?[市区町村郡].*?)(.+)$/);
  if (m) return { postal, pref_city: m[1].trim(), street: m[2].trim() };
  return { postal, pref_city: rest, street: '' };
}

// 住民票と異なる住所の値に紛れたラベル/注記を除去
function sanitizeAltAddressValue_(v) {
  let s = String(v || '').trim();
  // ラベル風の「【...】」を除去
  s = s.replace(/【[^】]*】/g, ' ');
  // 注記「（住民票と異なる場合）」を除去
  s = s.replace(/（\s*住民票と異なる場合\s*）/g, ' ');
  // 余分な空白を整形
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

// 差し込み直前の最終クリーニング（全角スペースも半角へ、空白整形）
function cleanAltAddress_(s) {
  const cleaned = normJaSpace_(sanitizeAltAddressValue_(s));
  return cleaned.replace(/^〒\s*/, '');
}

/** caseId を 4 桁ゼロ埋めの文字列に正規化（"1"→"0001"）。*/
function normalizeCaseIdString_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  // 数値化できる部分（数字のみ）を抽出してからゼロ埋め
  const digits = s.replace(/\D/g, '');
  if (!digits) return '';
  const n = parseInt(digits, 10);
  if (!isFinite(n)) return '';
  return String(n).padStart(4, '0');
}

/** ====== 保存・台帳アクセス ====== **/

function resolveCaseByCaseId_(caseId) {
  const ss = SpreadsheetApp.openById(S2002_SPREADSHEET_ID);
  const sh = ss.getSheetByName('cases');
  const vals = sh.getDataRange().getValues();
  const header = vals[0];
  const idxMap = bs_toIndexMap_(header);
  const idx = {
    caseId: idxMap.case_id,
    lineId: idxMap.line_id,
    caseKey: idxMap.case_key,
    folderId: idxMap.folder_id,
  };
  if (!(idx.caseId >= 0)) {
    throw new Error('cases sheet does not contain case_id column');
  }
  const want = normalizeCaseIdString_(caseId);
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    const rowCase = normalizeCaseIdString_(row[idx.caseId]);
    if (rowCase && rowCase === want) {
      return {
        caseId: rowCase,
        lineId: idx.lineId >= 0 ? row[idx.lineId] : '',
        caseKey: idx.caseKey >= 0 ? row[idx.caseKey] : '',
        folderId: idx.folderId >= 0 ? row[idx.folderId] : '',
        rowIndex: i + 1,
      };
    }
  }
  return null;
}

function updateCasesRow_(caseId, patch) {
  const ss = SpreadsheetApp.openById(S2002_SPREADSHEET_ID);
  const sh = ss.getSheetByName('cases');
  const vals = sh.getDataRange().getValues();
  const header = vals[0];
  const idxMap = bs_toIndexMap_(header);
  const idxCase = idxMap.case_id;
  const idxKey = idxMap.case_key;
  const idxLine = idxMap.line_id;
  // 列書式: case_id をテキスト('@')に固定（数値化防止）
  try {
    if (idxCase >= 0) sh.getRange(1, idxCase + 1, sh.getMaxRows(), 1).setNumberFormat('@');
  } catch (_) {}

  const wantCid = normalizeCaseIdString_(caseId);
  let rowIdx = -1;
  // 1) case_key 一致優先
  if (idxKey >= 0 && patch && patch.case_key) {
    const wantKey = String(patch.case_key || '').trim();
    rowIdx = vals.findIndex((r, i) => i > 0 && String(r[idxKey] || '').trim() === wantKey);
  }
  // 2) case_id 一致
  if (rowIdx < 0 && idxCase >= 0) {
    rowIdx = vals.findIndex((r, i) => i > 0 && normalizeCaseIdString_(r[idxCase]) === wantCid);
  }
  // 3) line_id 一致
  if (rowIdx < 0 && idxLine >= 0 && patch && patch.line_id) {
    const wantLine = String(patch.line_id || '').trim();
    rowIdx = vals.findIndex((r, i) => i > 0 && String(r[idxLine] || '').trim() === wantLine);
  }
  if (rowIdx < 1) return;
  const r = rowIdx + 1;
  Object.keys(patch).forEach((k) => {
    let c = idxMap[k];
    if (!(c >= 0) && typeof bs_headerAliases_ === 'function') {
      const aliases = bs_headerAliases_(k);
      for (let i = 0; i < aliases.length; i++) {
        const candidate = idxMap[aliases[i]];
        if (candidate >= 0) {
          c = candidate;
          break;
        }
      }
    }
    if (c >= 0) sh.getRange(r, c + 1).setValue(patch[k]);
  });
  // case_id の正規化書き戻し
  if (idxCase >= 0) sh.getRange(r, idxCase + 1).setValue(wantCid);
}

function saveSubmissionJson_(caseFolderId, parsed) {
  const parent = DriveApp.getFolderById(caseFolderId);
  const tz =
    (typeof Session !== 'undefined' && Session.getScriptTimeZone && Session.getScriptTimeZone()) ||
    'Asia/Tokyo';
  let sid =
    parsed && parsed.meta && parsed.meta.submission_id ? String(parsed.meta.submission_id) : '';
  sid = sid.replace(/[^\d]/g, '');
  if (!sid) {
    sid = Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmmss');
  }
  const safeKey = String(
    parsed && parsed.meta && parsed.meta.form_key ? parsed.meta.form_key : 'unknown'
  ).replace(/[^a-z0-9_]/gi, '_');
  const fname = `${safeKey}__${sid}.json`;
  const content = JSON.stringify(parsed, null, 2);
  let existing = null;
  const dupIt = parent.getFilesByName(fname);
  while (dupIt.hasNext()) {
    const candidate = dupIt.next();
    if (!existing) {
      try {
        candidate.setContent(content);
      } catch (_) {
        try {
          candidate.setTrashed(true);
          continue;
        } catch (_) {}
      }
      existing = candidate;
    } else {
      try {
        candidate.setTrashed(true);
      } catch (_) {}
    }
  }
  if (existing) return existing;
  const blob = Utilities.newBlob(content, 'application/json', fname);
  return parent.createFile(blob); // 仕様: ケース直下に保存
}

function getOrCreateSubfolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** ====== フォルダID補正ユーティリティ ====== **/
/**
 * cases 行の folderId を検証し、未設定/URL/誤値なら補正する（caseKey があれば Drive 直下で検索/作成）。
 * @param {{caseId:string, caseKey?:string, folderId?:string}} caseInfo
 * @returns {string} 正常な Drive フォルダID
 */
function ensureCaseFolderId_(caseInfo) {
  function isValidCaseKey_(s) {
    return /^[a-z0-9]{2,6}-\d{4}$/.test(String(s || ''));
  }
  // 1) 既存 folderId を URL→ID 抽出して試す
  const norm = extractDriveIdMaybe_(caseInfo.folderId || '');
  if (norm) {
    try {
      const f = DriveApp.getFolderById(norm);
      if (f && f.getId()) {
        // 既存IDが空フォルダ等で重複候補に「より良い」ものがある場合は乗り換える
        const currentScore = scoreCaseFolder_(f);
        const nameGuess =
          caseInfo.caseKey ||
          String(caseInfo.lineId || '')
            .slice(0, 6)
            .toLowerCase() +
            '-' +
            normalizeCaseIdString_(caseInfo.caseId || '');
        const better = findBestCaseFolderUnderRoot_(nameGuess, f.getId());
        if (better && better.id !== f.getId() && better.score > currentScore.score) {
          updateCasesRow_(caseInfo.caseId, { folder_id: better.id, case_key: nameGuess });
          return better.id;
        }
        // 現行のIDで続行
        if (norm !== caseInfo.folderId) updateCasesRow_(caseInfo.caseId, { folder_id: norm });
        return norm;
      }
    } catch (_) {
      // 続行して補正
    }
  }

  // 2) caseKey から Drive ルート直下で検索/作成（妥当性チェックを追加）
  const ROOT_ID =
    PROP_S2002.getProperty('DRIVE_ROOT_FOLDER_ID') ||
    PROP_S2002.getProperty('ROOT_FOLDER_ID') ||
    '';
  if (!ROOT_ID) throw new Error('ROOT_FOLDER_ID/DRIVE_ROOT_FOLDER_ID が未設定です');
  const root = DriveApp.getFolderById(ROOT_ID);
  let name = String(caseInfo.caseKey || '').trim();
  if (!name) {
    // cases に caseKey 列が無い場合は lineId+caseId から推定（userKey = lineId 先頭6文字）
    const lid = String(caseInfo.lineId || '').trim();
    const cid = normalizeCaseIdString_(caseInfo.caseId || '');
    const userKey = lid ? lid.slice(0, 6).toLowerCase() : '';
    if (userKey && cid) name = userKey + '-' + cid;
  }
  // 妥当な case_key 以外は拒否して再生成
  if (!isValidCaseKey_(name)) {
    const lid = String(caseInfo.lineId || '').trim();
    const cid = normalizeCaseIdString_(caseInfo.caseId || '');
    const uk = lid ? lid.slice(0, 6).toLowerCase() : '';
    if (uk && cid) name = uk + '-' + cid;
  }
  if (!isValidCaseKey_(name)) {
    throw new Error('case_key を生成できません（line_id または case_id が不足）');
  }
  if (!name)
    throw new Error(
      'caseKey が無く、lineId からも生成できません（cases に caseKey か lineId 列が必要）'
    );
  // 重複フォルダが複数存在する可能性があるため、内容がある方を優先
  const best = findBestCaseFolderUnderRoot_(name);
  const id = best ? best.id : root.createFolder(name).getId();
  // folderId は必ず書き戻し。正しい case_key も書き戻す。
  updateCasesRow_(caseInfo.caseId, { folder_id: id, case_key: name });
  return id;
}

/** URL/名前混在の値から Drive ID らしきものを抽出（見つからなければ空文字）。 */
function extractDriveIdMaybe_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  // URLパターン優先
  let m = s.match(/\/folders\/([A-Za-z0-9_-]{10,})/);
  if (m) return m[1];
  m = s.match(/\/d\/([A-Za-z0-9_-]{10,})/);
  if (m) return m[1];
  // 生IDらしきもの（10文字以上の [-_A-Za-z0-9]）
  if (/^[A-Za-z0-9_-]{10,}$/.test(s)) return s;
  return '';
}

/**
 * ルート直下の同名フォルダ（<caseKey>）候補から「中身がそれらしい」ものを選ぶ。
 * 優先度: s2002_userform__*.json の有無 > 直下jsonの有無 > drafts/ や attachments/ の存在。
 * @param {string} name caseKey（例: uc13df-0001）
 * @param {string=} preferId 既存folderId（スコア比較用）
 * @returns {{id:string, score:number}|null}
 */
function findBestCaseFolderUnderRoot_(name, preferId) {
  if (!name) return null;
  const ROOT_ID =
    PROP_S2002.getProperty('DRIVE_ROOT_FOLDER_ID') ||
    PROP_S2002.getProperty('ROOT_FOLDER_ID') ||
    '';
  if (!ROOT_ID) return null;
  const root = DriveApp.getFolderById(ROOT_ID);
  const it = root.getFoldersByName(name);
  const candidates = [];
  while (it.hasNext()) {
    const f = it.next();
    const sc = scoreCaseFolder_(f);
    candidates.push({ id: f.getId(), score: sc.score, detail: sc });
  }
  if (!candidates.length) return null;
  // preferId があり、そのスコアが最大と同点以上ならそれを返す
  candidates.sort((a, b) => b.score - a.score);
  if (preferId) {
    const cur = candidates.find((c) => c.id === preferId);
    if (cur && cur.score >= candidates[0].score) return cur;
  }
  return candidates[0];
}

/** ====== S2002 ドラフト生成（gdoc） ====== **/

function generateS2002Draft_(caseInfo, parsed) {
  if (!S2002_TPL_GDOC_ID) throw new Error('S2002_TEMPLATE_GDOC_ID not set');
  // generateS2002Draft_ 内
  const caseFolder = DriveApp.getFolderById(caseInfo.folderId);
  const drafts = getOrCreateSubfolder_(caseFolder, 'drafts');
  try {
    Logger.log('[S2002] draftsFolderId=%s caseFolderId=%s', drafts.getId(), caseInfo.folderId);
  } catch (_) {}

  const M = (parsed && parsed.model) || { app: {}, addr: {}, ref: {} };
  const intake = loadEarliestJsonByPrefix_(caseFolder, 'intake__');
  const intakeInfo = intake ? s2002_extractIntakeBaseInfo_(intake.json) : {};
  const intakePickName = intake && intake.file && intake.file.getName ? intake.file.getName() : '';
  try {
    Logger.log('[S2002] intakePickName=%s intakePickFound=%s', intakePickName, !!intake);
    Logger.log(
      '[S2002] intakeSources flags name=%s birth=%s addr=%s phone=%s',
      !!intakeInfo.name,
      !!intakeInfo.birth,
      !!intakeInfo.address,
      !!intakeInfo.phone
    );
  } catch (_) {}

  const app = Object.assign({}, M.app || {});
  const addr = Object.assign({}, M.addr || {});

  if (intakeInfo.name) app.name = normSpace_(intakeInfo.name);
  if (intakeInfo.phone) app.phone = normPhone_(intakeInfo.phone);
  if (intakeInfo.birth) {
    const iso = toIsoBirth_(intakeInfo.birth);
    if (iso) {
      app.birth = intakeInfo.birth.trim();
      app.birth_iso = iso;
      app.age = calcAge_(iso);
      const w0 = toWareki_(iso);
      if (w0) app.birth_wareki = w0;
    } else if (!app.birth) {
      app.birth = intakeInfo.birth.trim();
    }
  }

  const intakeZip = s2002_normalizeZip_(intakeInfo.zip);
  const intakeAddr = intakeInfo.address ? normJaSpace_(intakeInfo.address) : '';
  if (intakeAddr) addr.full = intakeAddr;
  if (intakeZip) addr.postal = intakeZip;
  if (addr.same_as_resident == null && intakeInfo.same_as_resident != null) {
    addr.same_as_resident = intakeInfo.same_as_resident;
  }
  let residentFull = '';
  if (addr.full) {
    const p = parseAddressLine_(addr.full);
    if (!addr.postal) addr.postal = p.postal || '';
    if (p.pref_city) addr.pref_city = p.pref_city;
    if (p.street) addr.street = p.street;
    const fromParts = (p.pref_city || '') + (p.street ? ' ' + p.street : '');
    residentFull = fromParts.trim() || normJaSpace_(addr.full);
  }
  if (!residentFull) {
    const fallback = [addr.pref_city, addr.street].filter(Boolean).join(' ');
    if (fallback) residentFull = fallback;
  }
  residentFull = s2002_stripZipFromAddress_(residentFull);
  const residenceBlock = s2002_buildResidenceBlock_({ zip: addr.postal, address: residentFull });
  let birthIsoForParts = app.birth_iso;
  if (!birthIsoForParts) {
    const srcBirth = app.birth || intakeInfo.birth || '';
    birthIsoForParts = toIsoBirth_(srcBirth);
    if (birthIsoForParts) app.birth_iso = birthIsoForParts;
  }
  try {
    Logger.log(
      '[S2002] afterMerge flags name=%s birth_iso=%s postal=%s full=%s phone=%s',
      !!app.name,
      !!app.birth_iso,
      !!addr.postal,
      !!addr.full,
      !!app.phone
    );
  } catch (_) {}

  // テンプレ複製（履歴保全のため submission_id を付与推奨）
  const subId =
    parsed.meta?.submission_id ||
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const draftName = `S2002_${caseInfo.caseId}_draft_${subId}`;
  const gdocId = DriveApp.getFileById(S2002_TPL_GDOC_ID).makeCopy(draftName, drafts).getId();
  try {
    Logger.log('[S2002] created gdocId=%s url=%s', gdocId, DocumentApp.openById(gdocId).getUrl());
  } catch (_) {}
  const doc = DocumentApp.openById(gdocId);
  const body = doc.getBody();

  // 申立人情報
  replaceAll_(body, '{{app.name}}', app.name || '');
  replaceAll_(body, '{{app.kana}}', app.kana || '');
  replaceAll_(body, '{{app.maiden_name}}', app.maiden_name || '');
  replaceAll_(body, '{{app.maiden_kana}}', app.maiden_kana || '');
  replaceAll_(body, '{{app.nationality}}', app.nationality || '');
  replaceAll_(body, '{{app.phone}}', app.phone || '');
  replaceAll_(body, '{{app.age}}', (app.age ?? '') + '');

  // 生年月日（西暦/元号）
  let birthY = '';
  let birthM = '';
  let birthD = '';
  if (birthIsoForParts) {
    const parts = birthIsoForParts.split('-');
    birthY = parts[0] || '';
    birthM = parts[1] || '';
    birthD = parts[2] || '';
  }
  replaceAll_(body, '{{app.birth_yyyy}}', birthY);
  replaceAll_(body, '{{app.birth_mm}}', String(birthM || ''));
  replaceAll_(body, '{{app.birth_dd}}', String(birthD || ''));
  try {
    Logger.log('[S2002] birthParts yyyy=%s mm=%s dd=%s', !!birthY, !!birthM, !!birthD);
  } catch (_) {}
  // ★ここに追加（旧: if (M.app.birth_wareki) { ... } は削除）
  const w = app.birth_wareki || (birthIsoForParts ? toWareki_(birthIsoForParts) : null);
  replaceAll_(body, '{{app.birth_wareki_gengo}}', (w && w.gengo) || '');
  replaceAll_(body, '{{app.birth_wareki_yy}}', (w && String(w.yy)) || '');

  // 住所
  const altClean0 = cleanAltAddress_(addr.alt_full);
  let altOut = altClean0;
  if (!altOut && addr && addr.same_as_resident === false) {
    addr.same_as_resident = null;
  }
  try {
    Logger.log(
      '[S2002] addr debug: same_as=%s alt_raw=%s alt_out=%s',
      addr && addr.same_as_resident,
      !!(addr && addr.alt_full),
      !!altOut
    );
  } catch (_) {}

  replaceAll_(body, '{{addr.postal}}', addr.postal || '');
  replaceAll_(body, '{{addr.pref_city}}', addr.pref_city || '');
  replaceAll_(body, '{{addr.street}}', addr.street || '');
  replaceAll_(body, '{{addr.alt_full}}', altOut || '');
  replaceAll_(body, '{{addr.same_as_resident}}', renderCheck_(addr.same_as_resident === true));
  replaceAll_(body, '{{addr.same_as_resident_no}}', renderCheck_(addr.same_as_resident === false));
  replaceAll_(body, '{{addr.zip}}', addr.postal || '');
  replaceAll_(body, '{{addr.address}}', addr.full || '');
  replaceAll_(body, '{{addr.full}}', addr.full || '');
  replaceAll_(body, '{{addr.resident_full}}', residentFull || '');
  replaceAll_(body, '{{addr.residence_block}}', residenceBlock || '');
  replaceAll_(body, '{{addr.residence}}', residenceBlock || '');
  s2002_appendResidenceAfterZipIfMissing_(body, addr.postal, residentFull);

  // 本籍・国籍（本籍は固定チェック、国籍は入力があれば☑）
  replaceAll_(body, '{{ref.domicile_resident}}', renderCheck_(true));
  replaceAll_(body, '{{ref.nationality}}', renderCheck_(!!app.nationality));

  // 個人事業者か（3択）
  replaceAll_(body, '{{ref.self_employed_none}}', renderCheck_(M.ref.self_employed === 'none'));
  replaceAll_(body, '{{ref.self_employed_past6m}}', renderCheck_(M.ref.self_employed === 'past6m'));
  replaceAll_(
    body,
    '{{ref.self_employed_current}}',
    renderCheck_(M.ref.self_employed === 'current')
  );

  // 法人代表者（有／無）
  replaceAll_(
    body,
    '{{ref.corp_representative}}',
    renderCheck_(M.ref.corp_representative === true)
  );
  replaceAll_(
    body,
    '{{ref.corp_representative_no}}',
    renderCheck_(M.ref.corp_representative === false)
  );

  // その法人の破産申立て（4択）
  replaceAll_(
    body,
    '{{ref.corp_bankruptcy_filed}}',
    renderCheck_(M.ref.corp_bankruptcy_status === '申立済み')
  );
  replaceAll_(
    body,
    '{{ref.corp_bankruptcy_same}}',
    renderCheck_(M.ref.corp_bankruptcy_status === '同時申立て')
  );
  replaceAll_(
    body,
    '{{ref.corp_bankruptcy_planned}}',
    renderCheck_(M.ref.corp_bankruptcy_status === '申立て予定')
  );
  replaceAll_(
    body,
    '{{ref.corp_bankruptcy_none}}',
    renderCheck_(M.ref.corp_bankruptcy_status === '予定なし')
  );

  // 生活保護（有／無）
  replaceAll_(body, '{{ref.welfare}}', renderCheck_(M.ref.welfare === true));
  replaceAll_(body, '{{ref.welfare_no}}', renderCheck_(M.ref.welfare === false));

  // 個人再生7年内（有／無）
  replaceAll_(body, '{{ref.pr_rehab_within7y}}', renderCheck_(M.ref.pr_rehab_within7y === true));
  replaceAll_(
    body,
    '{{ref.pr_rehab_within7y_no}}',
    renderCheck_(M.ref.pr_rehab_within7y === false)
  );

  // 免責7年内（有／無）
  replaceAll_(
    body,
    '{{ref.bankruptcy_discharge_within7y}}',
    renderCheck_(M.ref.bankruptcy_discharge_within7y === true)
  );
  replaceAll_(
    body,
    '{{ref.bankruptcy_discharge_within7y_no}}',
    renderCheck_(M.ref.bankruptcy_discharge_within7y === false)
  );

  // 文末
  doc.saveAndClose();
  return { gdocId, draftUrl: doc.getUrl() };
}

/** ====== 置換ユーティリティ ====== **/
function replaceAll_(body, token, value) {
  const safe = token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  body.replaceText(safe, escapeReplaceTextValue_(value));
}
function escapeReplaceTextValue_(value) {
  return String(value ?? '').replace(/\$/g, '$$$$');
}
function renderCheck_(b) {
  return b ? '☑' : '□';
}

/** ====== util ====== **/
// startsWithI_ は不要になったため削除（名前での前方一致ではなく JSON 中の form_key で判定）
function safeHtml(s) {
  return String(s).replace(
    /[&<>"']/g,
    (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m])
  );
}
// s2002_draft.js 内だけに定義して使う
function sanitizeAltAddressValue_S2002_(s) {
  return normJaSpace_(
    String(s || '')
      .replace(/【[^】]*】/g, ' ') // ラベル除去
      .replace(/（\s*住民票と異なる場合\s*）/g, ' ') // ← S2002固有の注記除去
  );
}

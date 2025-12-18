/********** 設定 **********/
// Script Properties を最優先に使用（DRIVE_ROOT_FOLDER_ID → ROOT_FOLDER_ID の順）

function getSecret_() {
  // 両方試す：BOOTSTRAP_SECRET 優先、なければ TOKEN_SECRET
  let s = props_().getProperty('BOOTSTRAP_SECRET') || props_().getProperty('TOKEN_SECRET') || '';
  if (s && typeof s.replace === 'function') s = s.replace(/[\r\n]+$/g, '');
  if (!s) throw new Error('missing secret');
  return s;
}

const ROOT_FOLDER_ID = (function () {
  const id = props_().getProperty('DRIVE_ROOT_FOLDER_ID') || props_().getProperty('ROOT_FOLDER_ID');
  if (!id) throw new Error('DRIVE_ROOT_FOLDER_ID/ROOT_FOLDER_ID is not configured');
  return id;
})();

const SHEET_ID = ''; // 既存の台帳シートを使うならIDを入れる。空なら後で自動生成可能
const LABEL_TO_PROCESS = 'FormAttach/ToProcess';
const LABEL_PROCESSED = 'FormAttach/Processed';
const LABEL_ERROR = 'FormAttach/Error';
// 通知の安全確認まわり（任意で切替可能）
const SUBJECT_TAG = '[#FM-BAS]'; // 件名に含まれるタグ（空なら無効）
const NOTIFY_SECRET = 'FM-BAS'; // META の secret 値（プロパティ未設定時のフォールバック）
const REQUIRE_SECRET = true; // true: タグとsecret両方必須 / false: どちらか片方でOK（共通ルール: 両方必須）
const ASIA_TOKYO = 'Asia/Tokyo';

// Script Properties 推奨（プロパティ未設定ならトンネル直URLを既定）
const BASE = (
  props_().getProperty('NEXT_BASE_URL') || 'https://depot-heath-television-ga.trycloudflare.com'
).replace(/\/+$/, '');
const NEXT_API_URL = `${BASE}/api/extract`;
const SECRET = props_().getProperty('SECRET') || 'FM-BAS';
const ENABLE_FORM_LOG_SHORTCUTS = false; // ← ショートカット不要なら false

/** 書類種別：フォルダは日本語(type)、ファイルは英数字コード(code) */
const DOC_MAP = [
  { labels: ['給与明細', '給料明細', '給料の明細', '給与'], type: '給与明細', code: 'PAY' },
  { labels: ['銀行通帳の写し', '銀行通帳', '通帳', 'bank'], type: '銀行通帳', code: 'BANK' },
  { labels: ['家計収支表', '家計簿', '収支表', 'budget'], type: '家計収支表', code: 'BUDG' },
];

/** 対象年月(YYYYMM)抽出 */
const PERIOD_PATS = [
  { pat: /(\d{4})[-./年](\d{1,2})(?:月)?/, to: (y, m) => `${y}${('0' + m).slice(-2)}` }, // 2025-08 / 2025年8月
  { pat: /(\d{2})[-./](\d{1,2})/, to: (y, m) => `20${y}${('0' + m).slice(-2)}` }, // 25-08
  { pat: /令和(\d+)年(\d{1,2})月/, to: (ry, m) => `${2018 + Number(ry)}${('0' + m).slice(-2)}` }, // 令和n年m月
];

// 既に名前付きフォルダがある場合の方針: 'first'（最初の名前を尊重）or 'latest'（常に最新に更新）
const RENAME_STRATEGY = 'latest';
const FOLDER_NAME_MAX = 80; // Drive の見やすさ用に適度に丸める

// 「書類提出」を1つのフォルダにまとめる
const DOCS_BUCKET_NAME = '書類提出';
const DOCS_FORM_KEYWORDS = ['書類提出']; // フォーム名にこの語が含まれたら集約
function isDocsSubmission(formName) {
  const s = String(formName || '');
  return DOCS_FORM_KEYWORDS.some((k) => s.includes(k));
}

// 原本（回答シート）をルートから探す
function findFormSheetInMyDriveRoot_(formName) {
  // ルート直下にあり、かつ Spreadsheet のものを名前一致で拾う
  // DriveApp は「場所での検索」が弱いので、名前一致で妥協（同名が複数あるケースは稀）
  const files = DriveApp.getFilesByName(formName);
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() === MimeType.GOOGLE_SHEETS) {
      // ルート直下かどうかは判定が難しいので、まずは最初の一致を採用（必要があれば厳密化）
      return f;
    }
  }
  return null;
}

//ショートカット作成（可能なら）＋台帳 upsert
function upsertFormLogRegistry_(formName) {
  const ss = openOrCreateMaster_();
  const sh = getOrCreateSheet(ss, 'form_logs');
  const data = sh.getDataRange().getValues();
  const now = Utilities.formatDate(new Date(), ASIA_TOKYO, 'yyyy-MM-dd HH:mm:ss');

  let row = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(formName)) {
      row = i + 1;
      break;
    }
  }

  const src = findFormSheetInMyDriveRoot_(formName);
  if (!src) {
    if (row > 0) sh.getRange(row, 8).setValue(now); // last_seen
    return null;
  }
  const sheetId = src.getId();
  const sheetUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;

  let shortcutId = '',
    shortcutUrl = '',
    mirrored = 'no';
  if (ENABLE_FORM_LOG_SHORTCUTS) {
    try {
      if (Drive && Drive.Files) {
        const folder = ensureFormLogsFolder_(); // _form_logs を本当に使う時だけ作る
        const maybe = folder.getFilesByName(formName + '（原本へのショートカット）');
        if (maybe.hasNext()) {
          const sc = maybe.next();
          shortcutId = sc.getId();
          shortcutUrl = `https://drive.google.com/file/d/${shortcutId}/view`;
          mirrored = 'yes';
        } else {
          const resource = {
            title: formName + '（原本へのショートカット）',
            mimeType: 'application/vnd.google-apps.shortcut',
            parents: [{ id: folder.getId() }],
            shortcutDetails: { targetId: sheetId },
          };
          const sc = Drive.Files.insert(resource);
          shortcutId = sc.id;
          shortcutUrl = `https://drive.google.com/file/d/${shortcutId}/view`;
          mirrored = 'yes';
        }
      }
    } catch (_) {
      mirrored = 'no';
    }
  }

  if (row < 0) {
    sh.appendRow([
      formName,
      sheetId,
      sheetUrl,
      fmt(new Date(src.getDateCreated())),
      shortcutId,
      shortcutUrl,
      mirrored,
      now,
    ]);
  } else {
    if (!data[row - 1][1]) sh.getRange(row, 2).setValue(sheetId);
    if (!data[row - 1][2]) sh.getRange(row, 3).setValue(sheetUrl);
    sh.getRange(row, 5).setValue(shortcutId);
    sh.getRange(row, 6).setValue(shortcutUrl);
    sh.getRange(row, 7).setValue(mirrored);
    sh.getRange(row, 8).setValue(now);
  }
  return { sheetId, sheetUrl, shortcutId, shortcutUrl, mirrored };
}

// 件名パース（テンプレに合わせた正規表現）
const RX_SUBJECT_SECRET = /\[#FM-BAS\]/i;
const RX_SUBJECT_LINE = /(?:^|\s)line:([A-Za-z0-9_-]+)/i;
const RX_SUBJECT_SID = /(?:^|\s)sid:([A-Za-z0-9_-]+)/i;

// META/FIELDSブロック
const RX_META_BLOCK = /====\s*META START\s*====([\s\S]*?)====\s*META END\s*====/i;
const RX_FIELDS_BLOCK = /====\s*FIELDS START\s*====([\s\S]*?)====\s*FIELDS END\s*====/i;

// META内のkey:value行
const RX_META_KV = /^\s*([A-Za-z0-9_]+)\s*:\s*(.*?)\s*$/gm;
// FIELDSの行（例: 「【ラベル】 値」 を抽出）
const RX_FIELD_LINE = /^\s*【(.+?)】\s*(.*)$/gm;

// ==== OCR PoC 追加設定 ====
const OCR_TARGET_FORMS = []; // []テスト用に空です。運用時はフォーム名に['書類', '給与明細']などを含む場合だけOCR実行
const OCR_LANG_HINTS = ['ja', 'en']; // 日本語＋英語
const OCR_PROCESSED_PROP = 'ocr_processed'; // Driveファイルのプロパティ名

/** ===== JSON保存: Google Drive (案件 or staging) ===== */
/** METAブロック抽出（==== META START ==== ... ==== META END ====） */
function parseMetaBlock_(text) {
  if (!text) return {};
  const m = text.match(/====\s*META START\s*====([\s\S]*?)(?:====\s*META END\s*====|$)/i);
  if (!m) return {};
  const meta = {};
  for (const raw of m[1].split(/\r?\n/)) {
    const line = String(raw).trim();
    if (!line || line.startsWith('#')) continue;
    const p = line.indexOf(':');
    if (p < 0) continue;
    meta[line.slice(0, p).trim().toLowerCase()] = line.slice(p + 1).trim();
  }
  return meta;
}

/** 英数_ 正規化 */
function normalizeKey_(s) {
  s = String(s || '');
  s = s.replace(/[！-～]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0)); // 全角→半角
  s = s.replace(/[^\x20-\x7E]/g, ' ');
  s = s.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
  return s || 'unknown';
}

/** 件名からフォームキーを推定（フォールバック）
 * 優先度: [BAS:xxx] > 既知の科目トークン（例:S2002）
 */
function resolveFormKeyFromSubject_(subject) {
  const s = String(subject || '');
  // 1) 明示指定: [BAS:xxx]
  let m = s.match(/\[\s*BAS\s*:\s*([A-Za-z0-9_]+)\s*\]/i);
  if (m) return normalizeKey_(m[1]);

  // 2) 既知の科目名トークン（件名に S2002 が含まれる等）
  const KNOWN = [{ re: /\bS2002\b/i, key: 's2002_userform' }];
  for (const k of KNOWN) {
    if (k.re.test(s)) return k.key;
  }

  return '';
}

/** 最終 form_key を決定：META > hidden field > 件名タグ > form_name 正規化 > unknown */
function resolveFormKeyFinal_(mailPlainBody, fields, subject) {
  const meta = parseMetaBlock_(mailPlainBody);
  let k = (meta.form_key || '').trim();
  if (!k) k = String(fields['form_key'] || fields['フォームキー'] || '').trim();
  if (!k) k = resolveFormKeyFromSubject_(subject);
  if (!k && meta.form_name) k = normalizeKey_(meta.form_name);
  return normalizeKey_(k || 'unknown');
}

/** ユーザーフォルダ直下に JSON 保存（submission_logs シートはログ専用） */
function saveSubmissionJsonShallow_(
  userFolder,
  submissionId,
  mailPlainBody,
  fields,
  subject,
  lineId
) {
  const formKey = resolveFormKeyFinal_(mailPlainBody, fields, subject);
  const caseId = resolveCaseId_(mailPlainBody, subject, lineId);
  const name = `${formKey}__${submissionId}.json`;
  const payload = JSON.stringify(
    { submission_id: submissionId, meta: { form_key: formKey, case_id: caseId, subject }, fields },
    null,
    2
  );
  const blob = Utilities.newBlob(payload, 'application/json', name);
  return userFolder.createFile(blob);
}

/** ===== 添付ファイル保存: 月フォルダ廃止・ファイル名YYYYMM化 ===== */
/** カテゴリ別の保存先フォルダ名とTYPEコード */
const ATTACH_RULE = {
  PAY: { folder: '給与明細', exts: ['png', 'jpg', 'jpeg', 'pdf', 'heic', 'webp'] },
  BANK: { folder: '銀行通帳', exts: ['png', 'jpg', 'jpeg', 'pdf', 'heic', 'webp'] },
  BUDG: { folder: '家計収支表', exts: ['png', 'jpg', 'jpeg', 'pdf', 'heic', 'webp'] },
  TAX: { folder: '税金関連', exts: ['png', 'jpg', 'jpeg', 'pdf', 'heic', 'webp'] },
  ETC: { folder: 'その他', exts: ['png', 'jpg', 'jpeg', 'pdf', 'heic', 'webp'] },
};

/** 元ファイル名や推定から TYPE を決める（必要に応じて拡張） */
function detectTypeCode_(name, hintedType) {
  const n = String(name || '').toLowerCase();
  const h = String(hintedType || '').toUpperCase();
  if (h && ATTACH_RULE[h]) return h;
  if (/pay|給与|給料|salary/.test(n)) return 'PAY';
  if (/bank|通帳|口座|statement/.test(n)) return 'BANK';
  if (/budg|家計|収支|budget/.test(n)) return 'BUDG';
  if (/tax|住民税|市県民税|国保|年金|保険料/.test(n)) return 'TAX';
  return 'ETC';
}

/** yyyymm を決める（メール受信日時 or 画像EXIFから決めてもOK。ここは受信日基準） */
function yyyymmFromDate_(d) {
  const pad = (x) => String(x).padStart(2, '0');
  return d.getFullYear() + pad(d.getMonth() + 1);
}

/** 同名があれば _2, _3... を付与して衝突回避 */
function uniqueNameInFolder_(folder, baseName, ext) {
  let name = `${baseName}.${ext}`;
  let i = 2;
  while (folder.getFilesByName(name).hasNext()) {
    name = `${baseName}_${i}.${ext}`;
    i++;
  }
  return name;
}

/**
 * 添付ファイルを「カテゴリ/ YYYYMM_TYPE[_連番].ext」で保存（浅い構造）
 * @param {Folder} userFolder - ユーザーのルートフォルダ（または staging のベース）
 * @param {Blob} blob - Gmailから取得した添付Blob
 * @param {Object} opts - { hintedType?: 'PAY'|'BANK'|'BUDG'|'TAX'|'ETC', receivedAt?: Date }
 */
function saveAttachmentShallow_(userFolder, blob, opts = {}) {
  const receivedAt = opts.receivedAt || new Date();
  const yyyymm = yyyymmFromDate_(receivedAt);

  const rawName = blob.getName() || 'file';
  const ext = (rawName.split('.').pop() || 'bin').toLowerCase();
  const type = detectTypeCode_(rawName, opts.hintedType);
  const rule = ATTACH_RULE[type] || ATTACH_RULE.ETC;

  const catFolder = getOrCreateFolder(userFolder, rule.folder);
  const base = `${yyyymm}_${type}`;
  const finalName = uniqueNameInFolder_(catFolder, base, ext);

  const file = catFolder.createFile(blob);
  file.setName(finalName);
  return file;
}

// _form_logs フォルダ（BAS配下）の確保
function ensureFormLogsFolder_() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID); // BAS_提出書類 のID
  return getOrCreateFolder(root, '_form_logs');
}

// 画像を“公開直リンク”に
function ensurePublicImageUrl_(file) {
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id = file.getId();
  // usercontent ドメインが安定
  return `https://drive.usercontent.google.com/u/0/uc?id=${id}&export=download`;
}

// 抽出API呼び出しの共通関数
function postExtract_(fileId, imageUrl, ocrText, lineId) {
  const url = NEXT_API_URL;
  const secret = SECRET || 'FM-BAS';
  const payload = { imageUrl, ocrText, lineId, sourceId: fileId };

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-Secret': secret },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    followRedirects: true,
    validateHttpsCertificates: true,
    escaping: false,
    // ← timeout はサポートされないので入れない
  });

  const status = resp.getResponseCode();
  const text = resp.getContentText();

  try {
    const json = JSON.parse(text);
    return { status, raw: text, ...json };
  } catch (e) {
    return { ok: false, status, error: String(e), raw: text };
  }
}

// LINE IDが取れない場合のステージング取得関数
function ensureStagingPath_(submitYm, submissionId, docTypeName) {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const sRoot = getOrCreateFolder(root, '_staging');
  const ym = getOrCreateFolder(sRoot, submitYm);
  const sub = getOrCreateFolder(ym, `submission_${submissionId || 'noid'}`);
  return getOrCreateFolder(sub, docTypeName || '未分類');
}

// Staging のベース（カテゴリはこの下で作る）
function ensureStagingBasePath_(submitYm, submissionId) {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const sRoot = getOrCreateFolder(root, '_staging');
  const ym = getOrCreateFolder(sRoot, submitYm);
  return getOrCreateFolder(ym, `submission_${submissionId || 'noid'}`);
}

/********** ユーティリティ **********/
function fmt(d) {
  return Utilities.formatDate(d, ASIA_TOKYO, 'yyyy-MM-dd HH:mm:ss');
}
function tsKey(d) {
  return Utilities.formatDate(d, ASIA_TOKYO, 'yyyyMMdd_HHmmss');
}
function sanitize(s) {
  return String(s || '')
    .replace(/[\\/:*?"<>|]/g, '_')
    .trim();
}

function normalizeSubmissionIdStrict_(s) {
  s = String(s || '')
    .trim()
    .replace(/[！-～]/g, function (ch) {
      return String.fromCharCode(ch.charCodeAt(0) - 0xfee0);
    })
    .replace(/[\u200B-\u200D\uFEFF]/g, '');
  if (!s) return '';
  if (/^\d{3,}$/.test(s)) return s;
  return '';
}

function getOrCreateFolder(parent, name) {
  const nm = sanitize(name || 'unknown');
  const it = parent.getFoldersByName(nm);
  return it.hasNext() ? it.next() : parent.createFolder(nm);
}

/** userKey = lineId 先頭6小文字 */
function userKeyFromLineId_(lineId) {
  return String(lineId || '')
    .slice(0, 6)
    .toLowerCase();
}
/** caseKey = userKey-caseId(4桁ゼロ埋め) */
function caseKey_(lineId, caseId) {
  const uk = userKeyFromLineId_(lineId);
  const cid = String(caseId || '').padStart(4, '0');
  return `${uk}-${cid}`;
}
/** <userKey-caseId>/ を BAS ルート直下に作成、標準サブフォルダも用意 */
function ensureCaseFolder_(lineId, caseId) {
  const key = caseKey_(lineId, caseId);
  const folder = drive_getOrCreateCaseFolderByKey_(key);
  getOrCreateFolder(folder, 'attachments');
  getOrCreateFolder(folder, 'staff_inputs');
  getOrCreateFolder(folder, 'drafts');
  return folder;
}

// ===== JSON バインド支援ユーティリティ =====
function needsCaseBinding_(obj) {
  try {
    const m = (obj && obj.meta) || {};
    return !('case_id' in m) || m.case_id == null || String(m.case_id) === '';
  } catch (_) {
    return false;
  }
}
function normalizeJsonName_(name) {
  const n = String(name || '')
    .replace(/[ \u3000]/g, '_')
    .replace(/[()（）]/g, '_');
  return /\.json$/i.test(n) ? n : n + '.json';
}
function sha1_(blob) {
  try {
    const dig = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, blob.getBytes());
    return Utilities.base64Encode(dig);
  } catch (e) {
    return '';
  }
}
function getStagingFolders_() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const out = [];
  const it1 = root.getFoldersByName('_staging');
  if (it1.hasNext()) out.push(it1.next());
  const it2 = root.getFoldersByName('_email_staging');
  if (it2.hasNext()) out.push(it2.next());
  return out;
}

function moveJsonsUnderFolderToCase_(srcFolder, caseFolder, caseId, lineId) {
  // 1) ファイルを処理
  const files = srcFolder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const nm = f.getName();
    if (!/\.json$/i.test(nm)) continue;
    let obj = null;
    let text = '';
    try {
      text = f.getBlob().getDataAsString('utf-8');
      obj = JSON.parse(text);
    } catch (_) {}

    const isIntakeByName = (typeof isIntakeJsonName_ === 'function' ? isIntakeJsonName_(nm) : /^intake__/i.test(nm));
    const isIntakeByMeta = String(obj?.meta?.form_key || '').toLowerCase() === 'intake';
    const needsBinding = obj ? needsCaseBinding_(obj) : false;
    if (!(isIntakeByName || isIntakeByMeta || needsBinding)) continue;

    // 追加ガード: 指定 lineId/caseId のものだけを対象にする（誤移送防止）
    try {
      const m = (obj && obj.meta) || {};
      const uk = userKeyFromLineId_(lineId || '');
      const cid4 = String(caseId || '').padStart(4, '0');
      const wantCK = uk ? uk + '-' + cid4 : '';
      const res = (typeof matchMetaToCase_ === 'function')
        ? matchMetaToCase_(m, { case_key: wantCK, case_id: cid4, line_id: String(lineId || '') })
        : { ok: true };
      if (!res.ok) continue;
    } catch (_) {}

    // meta.case_id を補完
    obj = obj || {};
    obj.meta = obj.meta || {};
    obj.meta.case_id = String(caseId);
    const outText = JSON.stringify(obj, null, 2);
    const outBlob = Utilities.newBlob(outText, 'application/json', 'tmp.json');
    const outHash = sha1_(outBlob);

    // ファイル名整形＋重複回避
    let baseName = normalizeJsonName_(nm);
    let finalName = baseName;
    let i = 2;
    while (caseFolder.getFilesByName(finalName).hasNext()) {
      // 既存同名があれば内容ハッシュで同一性チェック
      const ex = caseFolder.getFilesByName(finalName).next();
      const exHash = sha1_(ex.getBlob());
      if (exHash && exHash === outHash) {
        // 同一なら staging から外すだけ
        try {
          srcFolder.removeFile(f);
        } catch (_) {}
        finalName = '';
        break;
      }
      finalName = baseName.replace(/(\.json)$/i, `_${i}$1`);
      i++;
    }
    if (!finalName) continue; // 同一だった

    // 擬似アトミック書き込み（tmp名→本名）
    const tmpName = '._tmp_' + finalName;
    const tmpFile = caseFolder.createFile(Utilities.newBlob(outText, 'application/json', tmpName));
    try {
      tmpFile.setName(finalName);
    } catch (_) {}

    // staging から除去
    try {
      srcFolder.removeFile(f);
    } catch (_) {}

    // 構造化ログ
    try {
      Logger.log(
        JSON.stringify({
          at: 'moveStagingToCase',
          caseId: caseId,
          src: srcFolder.getName() + '/' + nm,
          dst: caseFolder.getName() + '/' + finalName,
          mode: 'rewrite+move',
        })
      );
    } catch (_) {}
  }

  // 2) サブフォルダを再帰
  const subs = srcFolder.getFolders();
  while (subs.hasNext()) {
    const sub = subs.next();
    moveJsonsUnderFolderToCase_(sub, caseFolder, caseId);
    // 空なら片付け
    if (!sub.getFiles().hasNext() && !sub.getFolders().hasNext()) {
      try {
        srcFolder.removeFolder(sub);
      } catch (_) {}
    }
  }
}
function ensureLabels() {
  GmailApp.getUserLabelByName(LABEL_TO_PROCESS) || GmailApp.createLabel(LABEL_TO_PROCESS);
  GmailApp.getUserLabelByName(LABEL_PROCESSED) || GmailApp.createLabel(LABEL_PROCESSED);
  GmailApp.getUserLabelByName(LABEL_ERROR) || GmailApp.createLabel(LABEL_ERROR);
}
// 安全に Spreadsheet を確保（Sheet を誤って渡された場合でも親を辿る）
function ensureSpreadsheet_(obj) {
  if (obj && typeof obj.getSheetByName === 'function') return obj; // Spreadsheet
  if (obj && typeof obj.getParent === 'function') {
    const parent = obj.getParent && obj.getParent();
    if (parent && typeof parent.getSheetByName === 'function') return parent;
  }
  throw new Error('not_a_spreadsheet');
}

function getOrCreateSheet(ss, name, header) {
  ss = ensureSpreadsheet_(ss);
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0 && header && header.length) sh.appendRow(header);
  return sh;
}

// ---- email secondary key & dedupe ----
function normalizeEmail_(s) {
  return String(s || '')
    .trim()
    .toLowerCase();
}
function _hex(bytes) {
  return bytes.map((b) => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function emailHash_(email) {
  const e = normalizeEmail_(email);
  if (!e) return '';
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, e);
  return _hex(digest).slice(0, 16);
}
function pickEmail(fields) {
  const keys = ['メール', 'メールアドレス', 'email', 'e-mail', 'mail', 'Email'];
  for (const k of keys) {
    if (fields && fields[k]) return normalizeEmail_(fields[k]);
  }
  return '';
}

function ensureEmailStagingPath_(submitYm, emailHash, docTypeName) {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const eRoot = getOrCreateFolder(root, '_email_staging');
  const ym = getOrCreateFolder(eRoot, submitYm);
  const eh = getOrCreateFolder(ym, emailHash || 'noemail');
  return getOrCreateFolder(eh, docTypeName || '未分類');
}

// email staging のベース（カテゴリはこの下で作る）
function ensureEmailStagingBasePath_(submitYm, emailHash) {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const eRoot = getOrCreateFolder(root, '_email_staging');
  const ym = getOrCreateFolder(eRoot, submitYm);
  return getOrCreateFolder(ym, emailHash || 'noemail');
}
function contentHashHex_(blob) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, blob.getBytes());
  return _hex(digest);
}
function findExistingByHash_(folder, contentHash) {
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if ((f.getDescription() || '').indexOf(`content_hash=${contentHash}`) >= 0) return f;
  }
  return null;
}
function countExistingWithBase_(folder, base) {
  let c = 0;
  const re = new RegExp(
    '^' + base.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?:_\\d{2})?\\.[A-Za-z0-9]{1,6}$',
    'i'
  );
  const it = folder.getFiles();
  while (it.hasNext()) {
    if (re.test(it.next().getName())) c++;
  }
  return c;
}

/** _email_staging からユーザーフォルダへ統合 */
function reconcileEmailToUser_(lineId, displayName, email, submitYm) {
  const e = normalizeEmail_(email || '');
  if (!e) return;
  const eh = emailHash_(e);
  if (!eh) return;

  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const stagingRoot = getOrCreateFolder(root, '_email_staging');
  const ymFolder = getOrCreateFolder(stagingRoot, submitYm);
  // email_hash 階層が無ければ終わり
  const itEh = ymFolder.getFoldersByName(eh);
  if (!itEh.hasNext()) return;

  const srcEmailFolder = itEh.next(); // <YYYY-MM>/<email_hash>
  const userFolder = getOrCreateUserFolder(root, lineId, displayName || '');

  // docType 階層ごとに移動（年月フォルダは作らない）
  const docTypes = srcEmailFolder.getFolders();
  while (docTypes.hasNext()) {
    const docTypeFolder = docTypes.next(); // 例: "給与明細"
    const destType = getOrCreateFolder(userFolder, docTypeFolder.getName());

    // ファイルごとに重複チェックして移動（宛先はカテゴリ直下）
    const files = docTypeFolder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      const desc = f.getDescription() || '';
      const m = desc.match(/content_hash=([0-9a-f]{64})/i);
      const ch = m ? m[1] : '';
      const dup = ch ? findExistingByHash_(destType, ch) : null;
      if (!dup) {
        destType.addFile(f);
        srcEmailFolder.removeFile(f);
      }
    }
    // 空になったら片付け（任意）
    if (!docTypeFolder.getFiles().hasNext()) srcEmailFolder.removeFolder(docTypeFolder);
  }
  // email_hash フォルダが空なら片付け（任意）
  if (!srcEmailFolder.getFolders().hasNext() && !srcEmailFolder.getFiles().hasNext()) {
    ymFolder.removeFolder(srcEmailFolder);
  }
}

/** _email_staging から <caseKey>/attachments/<カテゴリ> へ統合 */
function reconcileEmailStagingToCase_(lineId, caseId, email, submitYm) {
  const e = normalizeEmail_(email || '');
  const eh = emailHash_(e);
  if (!eh) return;

  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const stagingRoot = getOrCreateFolder(root, '_email_staging');
  const ymFolder = getOrCreateFolder(stagingRoot, submitYm);
  const itEh = ymFolder.getFoldersByName(eh);
  if (!itEh.hasNext()) return;

  const srcEmailFolder = itEh.next(); // <YYYY-MM>/<email_hash>
  const caseFolder = ensureCaseFolder_(lineId, caseId);
  const attachRoot = getOrCreateFolder(caseFolder, 'attachments');

  const docTypes = srcEmailFolder.getFolders();
  while (docTypes.hasNext()) {
    const docTypeFolder = docTypes.next(); // 例: "給与明細"
    const destType = getOrCreateFolder(attachRoot, docTypeFolder.getName());

    const files = docTypeFolder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      const desc = f.getDescription() || '';
      const m = desc.match(/content_hash=([0-9a-f]{64})/i);
      const ch = m ? m[1] : '';
      const dup = ch ? findExistingByHash_(destType, ch) : null;
      if (!dup) {
        destType.addFile(f);
        srcEmailFolder.removeFile(f);
      }
    }
    if (!docTypeFolder.getFiles().hasNext()) srcEmailFolder.removeFolder(docTypeFolder);
  }
  if (!srcEmailFolder.getFolders().hasNext() && !srcEmailFolder.getFiles().hasNext()) {
    ymFolder.removeFolder(srcEmailFolder);
  }
}

/** 旧フォルダ（氏名__LINEID）→ 新フォルダ（<userKey-caseId>）へ移送（必要時に実行） */
function migrateLegacyUserFolderToCaseKey_(lineId, caseId, legacyDisplayName) {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const legacyBase = String(legacyDisplayName || '').trim()
    ? `${sanitize(legacyDisplayName)}__${sanitize(lineId)}`
    : sanitize(lineId);
  const it = root.getFoldersByName(legacyBase);
  if (!it.hasNext()) return;

  const legacy = it.next();
  const caseFolder = ensureCaseFolder_(lineId, caseId);
  const attachRoot = getOrCreateFolder(caseFolder, 'attachments');

  // 旧直下の「カテゴリ」フォルダを attachments 配下に移す（重複は content_hash で回避）
  const cats = legacy.getFolders();
  while (cats.hasNext()) {
    const cat = cats.next();
    const destCat = getOrCreateFolder(attachRoot, cat.getName());
    const files = cat.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      const m = (f.getDescription() || '').match(/content_hash=([0-9a-f]{64})/i);
      const dup = m ? findExistingByHash_(destCat, m[1]) : null;
      if (!dup) {
        destCat.addFile(f);
        legacy.removeFile(f);
      }
    }
  }
}

/**
 * _staging 配下に一時保存された intake JSON（例: intake__<sid>.json）を
 * 新規作成された <caseKey>/ 直下へ移動し、meta.case_id を書き戻す。
 *
 * - 対象: _staging/<YYYY-MM>/submission_/intake__.json または meta.form_key==='intake'
 * - 動作: JSON を読み込み meta.case_id を上書き → 案件直下へ新規作成 → 元ファイルを staging から外す
 */

function moveStagingIntakeJsonToCase_(lineId, caseId) {
  try {
    try {
      Logger.log(
        '[moveStagingIntakeJsonToCase_] start lineId=%s caseId=%s root=%s',
        lineId,
        caseId,
        ROOT_FOLDER_ID
      );
    } catch (_) {}
    const caseFolder = ensureCaseFolder_(lineId, caseId);
    const stagingRoots = getStagingFolders_();
    try {
      Logger.log(
        '[moveStagingIntakeJsonToCase_] stagingRoots=%s caseFolder=%s',
        stagingRoots
          .map(function (f) {
            return (f && f.getName && f.getName()) || '?';
          })
          .join(','),
        (caseFolder && caseFolder.getName && caseFolder.getName()) || '?'
      );
    } catch (_) {}
    stagingRoots.forEach(function (st) {
      const ymIter = st.getFolders();
      while (ymIter.hasNext()) {
        const ym = ymIter.next(); // 'YYYY-MM'
        const unitIter = ym.getFolders(); // submission_* or email_hash
        while (unitIter.hasNext()) {
          const unit = unitIter.next();
          moveJsonsUnderFolderToCase_(unit, caseFolder, caseId, lineId);
          if (!unit.getFiles().hasNext() && !unit.getFolders().hasNext()) {
            try {
              ym.removeFolder(unit);
            } catch (_) {}
          }
        }
        if (!ym.getFiles().hasNext() && !ym.getFolders().hasNext()) {
          try {
            st.removeFolder(ym);
          } catch (_) {}
        }
      }
    });
  } catch (e) {
    Logger.log('[moveStagingIntakeJsonToCase_] error: %s', (e && e.stack) || e);
  }
}

/** 手動デバッグ用: 指定 lineId/caseId で mover を実行し、結果をログに出力 */
function debug_moveStagingToCase(lineId = 'Uc13df94016ee50eb9dd5552bffbe6624', caseId = '0001') {
  try {
    Logger.log('[debug_moveStagingToCase] ROOT_FOLDER_ID=%s', ROOT_FOLDER_ID);
    moveStagingIntakeJsonToCase_(lineId, caseId);
    // 結果確認（案件直下の .json を列挙）
    const cf = ensureCaseFolder_(lineId, caseId);
    const names = [];
    const it = cf.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (/\.json$/i.test(f.getName())) names.push(f.getName());
    }
    Logger.log('[debug_moveStagingToCase] caseFolder=%s jsons=%s', cf.getName(), names.join(','));
  } catch (e) {
    Logger.log('[debug_moveStagingToCase] error: %s', (e && e.stack) || e);
  }
}

/** 本文から「【項目名】↵（先頭が空白）ファイル名」を列挙 */
function extractLabelFilenamePairsFromBody_(body) {
  const lines = String(body || '').split(/\r?\n/);
  const out = [];
  let current = null;
  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i];
    const t = raw.trim();
    const m = t.match(/^【(.+?)】$/);
    if (m) {
      current = m[1].trim();
      continue;
    }
    if (current && /^[　\s]/.test(raw)) {
      const fname = t;
      if (fname) out.push({ label: current, filename: fname });
    }
  }
  return out; // [{label:'給与明細', filename:'テスト給料明細1.PNG'}]
}

/** DOC_MAP から辞書を作る */
function buildDocDict_() {
  const norm = (s) => (s || '').toLowerCase().replace(/\s+/g, '').replace(/[　]+/g, '');
  const labelToTypeCode = {};
  const typeToCode = {};
  const normTypeToType = {};
  DOC_MAP.forEach((e) => {
    typeToCode[e.type] = e.code;
    normTypeToType[norm(e.type)] = e.type;
    (e.labels || []).forEach((l) => {
      labelToTypeCode[norm(l)] = { type: e.type, code: e.code };
      normTypeToType[norm(l)] = e.type;
    });
  });
  return { labelToTypeCode, typeToCode, normTypeToType, norm };
}

/** ペア照合（完全一致→拡張子抜き一致→部分一致） */
function matchByPairs_(originalName, pairs, labelToTypeCode) {
  const norm = (s) => (s || '').toLowerCase().replace(/\s+/g, '').replace(/[　]+/g, '');
  const on = norm(originalName);
  const rmExt = (s) => s.replace(/\.[A-Za-z0-9]{1,6}$/, '');
  let p = pairs.find((pr) => norm(pr.filename) === on);
  if (p) return { ...(labelToTypeCode[norm(p.label)] || {}), fromName: p.filename };
  p = pairs.find(
    (pr) => rmExt(norm(pr.filename)) === rmExt(on) || on.includes(rmExt(norm(pr.filename)))
  );
  if (p) return { ...(labelToTypeCode[norm(p.label)] || {}), fromName: p.filename };
  return null;
}

/** ファイル名ヒューリスティック（保険） */
function guessTypeNameByFilename_(name) {
  const s = String(name || '').toLowerCase();
  if (/salary|pay.?slip|給与|給料|給與|明細/.test(s)) return '給与明細';
  if (/bank|通帳|残高|入出金|statement/.test(s)) return '銀行通帳';
  if (/家計|収支|budget/.test(s)) return '家計収支表';
  return null;
}

/** 対象YYYYMM抽出 */
function guessPeriod_(s) {
  if (!s) return '';
  for (const p of PERIOD_PATS) {
    const m = String(s).match(p.pat);
    if (m) return p.to.apply(null, m.slice(1));
  }
  return '';
}

/********** パース **********/
function parseMetaAndFields(msg) {
  const subject = msg.getSubject() || '';
  const bodyPlain = msg.getPlainBody() || '';
  const bodyHtml = msg.getBody() || '';
  const body = bodyPlain || bodyHtml;

  const subjectLine = (subject.match(RX_SUBJECT_LINE) || [])[1] || '';
  const subjectSid = (subject.match(RX_SUBJECT_SID) || [])[1] || '';
  const subjectSecretOK = RX_SUBJECT_SECRET.test(subject);

  const metaBlock = (body.match(RX_META_BLOCK) || [])[1] || '';
  const meta = {};
  let m;
  while ((m = RX_META_KV.exec(metaBlock)) !== null) {
    meta[m[1].toLowerCase()] = m[2];
  }

  // まずは従来の FIELDS ブロック
  const fields = {};
  let f;
  const fieldsBlock = (body.match(RX_FIELDS_BLOCK) || [])[1] || '';
  while ((f = RX_FIELD_LINE.exec(fieldsBlock)) !== null) {
    const label = sanitize(f[1]);
    fields[label] = f[2] || '';
  }

  // ★フォールバック：メール本文全体から「【ラベル】 値」を拾う（FIELDSブロックが無いとき）
  if (Object.keys(fields).length === 0) {
    const RX_INLINE = /【(.+?)】\s*([^\n\r]+)/g;
    let g;
    while ((g = RX_INLINE.exec(body)) !== null) {
      const k = sanitize(g[1]);
      const v = (g[2] || '').trim();
      if (k && v) fields[k] = v;
    }
  }

  const line_id = sanitize(meta.line_id || subjectLine || '');
  const form_name = sanitize(meta.form_name || '');
  const submission_id_raw = meta.submission_id || subjectSid || '';
  const submission_id = normalizeSubmissionIdStrict_(submission_id_raw);
  const submitted_at = meta.submitted_at || '';
  const seq = meta.seq || '';
  const referrer = meta.referrer || '';
  const secretOK = (meta.secret || '').trim() === SECRET && subjectSecretOK;

  return {
    subject,
    body,
    line_id,
    form_name,
    submission_id,
    submitted_at,
    seq,
    referrer,
    secretOK,
    fields,
  };
}

/********** 簡易METAパース & ガード **********/
function parseMeta_(body) {
  try {
    const m = String(body || '').match(
      /====\s*META START\s*====\s*([\s\S]*?)\s*====\s*META END\s*====/
    );
    if (!m) return {};
    const kv = {};
    String(m[1] || '')
      .split(/\r?\n/)
      .forEach((line) => {
        const i = line.indexOf(':');
        if (i < 0) return;
        const k = String(line.slice(0, i)).trim();
        const v = String(line.slice(i + 1)).trim();
        if (k) kv[k] = v;
      });
    return kv; // { form_key, case_id, secret, ... }
  } catch (_) {
    return {};
  }
}

// 共通トークン正規化：全角→半角、ダッシュ類統一、ゼロ幅除去、空白除去、小文字化
function _normToken(s) {
  s = String(s || '');
  // 全角 → 半角
  s = s.replace(/[！-～]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0));
  // よく混入する各種ハイフン/ダッシュ/長音を ASCII ハイフンへ統一
  s = s.replace(/[-‐‒–—−ー－]/g, '-');
  // ゼロ幅スペース除去
  s = s.replace(/[\u200B-\u200D\uFEFF]/g, '');
  // 空白全削除＋小文字
  return s.replace(/\s+/g, '').toLowerCase();
}

function checkNotificationGuard_(msg, meta) {
  const subjRaw = String((msg && msg.getSubject && msg.getSubject()) || '');
  const subj = _normToken(subjRaw);

  // 件名タグ（[#FM-BAS]）の柔軟判定：全角記号・各種ハイフン差異・空白混入を吸収
  // 例: [#FM-BAS], ［#FM–BAS］, [＃ＦＭ－ＢＡＳ] などもOKにする
  const tagNorm = _normToken(SUBJECT_TAG || '[#FM-BAS]'); // 期待タグ
  // 期待タグからブラケットと#を取り除いた“核”も作る（本文側に #FM-BAS があるだけでもOK）
  const tagCore = tagNorm.replace(/[\[\]［］#＃]/g, '');

  const hasSubjectTag = (() => {
    const subjCore = subj.replace(/[\[\]［］]/g, ''); // ブラケット差異を潰す
    return (
      subj.includes(tagNorm) || // 完全一致（正規化後）
      subjCore.includes(tagCore) || // ブラケット無しでも“fm-bas”が入っていればOK
      /\bfm-bas\b/i.test(subjRaw) // 生テキストにも一応保険（ASCIIで）
    );
  })();

  // META キーをゆるくローワー化して拾う
  const metaLower = {};
  try {
    Object.keys(meta || {}).forEach((k) => (metaLower[String(k).toLowerCase()] = meta[k]));
  } catch (_) {}

  // 期待シークレット：プロパティ優先、無ければ定数
  const expectedSecretRaw =
    (props_() && (props_().getProperty('NOTIFY_SECRET') || props_().getProperty('SECRET'))) ||
    (typeof NOTIFY_SECRET !== 'undefined' ? NOTIFY_SECRET : '') ||
    '';
  const expected = _normToken(expectedSecretRaw);
  const provided = _normToken(metaLower['secret'] || '');

  const hasSecret = !!provided && provided === expected;
  const ok = REQUIRE_SECRET ? hasSubjectTag && hasSecret : hasSubjectTag || hasSecret;

  // デバッグ用ログ
  Logger.log(
    '[guard] subjTag=%s secret=%s (subjRaw=%s, meta.secret=%s, expected=%s)',
    hasSubjectTag,
    hasSecret,
    subjRaw,
    metaLower['secret'] || '',
    expectedSecretRaw
  );

  return { ok, hasTag: hasSubjectTag, hasSecret };
}

/********** 添付保存 + JSON保存 **********/
function saveAttachmentsAndJson(meta, msg) {
  const when = msg.getDate();
  const submitYm = Utilities.formatDate(when, ASIA_TOKYO, 'yyyy-MM'); // 提出月（送信日時）
  const normalizedSubmissionId = normalizeSubmissionIdStrict_(meta.submission_id);
  let submissionId = normalizedSubmissionId;
  if (!submissionId) {
    submissionId = String(Date.now());
    try {
      Logger.log(
        '[email-intake] sid_fallback from "%s" -> "%s"',
        meta.submission_id || '',
        submissionId
      );
    } catch (_) {}
  }
  meta.submission_id = submissionId;
  const display = pickDisplayName(meta.fields);
  const body = meta.body || msg.getPlainBody() || msg.getBody() || '';

  // LINEが分かった & メールも分かる場合は、先に email_staging を案件フォルダ(attachments/)へ統合
  const emailNow = pickEmail(meta.fields);
  if (meta.line_id && emailNow) {
    try {
      const caseIdForMerge = resolveCaseId_(body, meta.subject || '', meta.line_id || '');
      if (caseIdForMerge)
        reconcileEmailStagingToCase_(meta.line_id, caseIdForMerge, emailNow, submitYm);
    } catch (_) {}
  }

  // 本文の「【項目名】↔ファイル名」ペアと辞書
  const pairs = extractLabelFilenamePairsFromBody_(body);
  const maps = buildDocDict_();

  // 添付ごとに「日本語フォルダ」「英数字ファイル名」で保存
  const saved = [];
  const seqCounter = {}; // docCode|ymForName -> seq
  const atts = msg.getAttachments({ includeInlineImages: false, includeAttachments: true }) || [];

  atts.forEach((att) => {
    if (att.isGoogleType()) return;
    const original = att.getName();
    const ext = (original.match(/(\.[A-Za-z0-9]{1,6})$/)?.[1] || '').toLowerCase();
    const blob = att.copyBlob();
    const cHash = contentHashHex_(blob);

    // 1) TYPE推定（本文ペア→ファイル名ヒューリスティック）
    const hit = matchByPairs_(original, pairs, maps.labelToTypeCode);
    const hinted = hit?.code || '';
    const typeCode = detectTypeCode_(original, hinted);
    const rule = ATTACH_RULE[typeCode] || ATTACH_RULE.ETC;
    const typeName = rule.folder;
    const period = yyyymmFromDate_(when);

    // 2) 保存先のベース（caseKey フォルダに統一）
    const email = pickEmail(meta.fields);
    const emailH = emailHash_(email);
    const caseId = resolveCaseId_(body, meta.subject || '', meta.line_id || '');
    let caseFolder, baseFolder;
    if (meta.line_id && caseId) {
      caseFolder = ensureCaseFolder_(meta.line_id, caseId);
      baseFolder = getOrCreateFolder(caseFolder, 'attachments');
    } else if (emailH) {
      baseFolder = ensureEmailStagingBasePath_(submitYm, emailH);
    } else {
      baseFolder = ensureStagingBasePath_(submitYm, submissionId);
    }

    // 3) カテゴリフォルダ（カテゴリ直下）
    const dest = getOrCreateFolder(baseFolder, typeName);

    // 4) 重複チェック（content_hash）
    const existed = findExistingByHash_(dest, cHash);
    if (existed) {
      // ★ デデュープでも submissions / contacts を記録する
      appendSubmissionLog_({
        line_id: meta.line_id || '',
        email: pickEmail(meta.fields) || '',
        submission_id: submissionId,
        form_name: meta.form_name || '',
        submit_ym: submitYm,
        doc_code: typeCode,
        period_yyyymm: period || '',
        drive_file_id: existed.getId(),
        file_name: existed.getName(),
        folder_path: dest.getName(),
        content_hash: cHash || '',
      });
      upsertContactLegacy_({
        line_id: meta.line_id || '',
        email: pickEmail(meta.fields) || '',
        display_name: display || '',
        last_form: meta.form_name || '',
        last_submit_ym: submitYm,
      });
      saved.push({
        id: existed.getId(),
        name: existed.getName(),
        size: existed.getSize(),
        folderId: dest.getId(),
        doc_type: typeName,
        doc_code: typeCode,
        period_yyyymm: period || '',
        dedup: true,
      });
      return;
    }
    // 5) 保存（saveAttachmentShallow_を使用）
    const file = saveAttachmentShallow_(baseFolder, blob, {
      hintedType: typeCode,
      receivedAt: when,
    });
    file.setDescription(
      [
        `original=${original}`,
        `doc_type=${typeName}`,
        `doc_code=${typeCode}`,
        `period=${period || ''}`,
        `submitted_at=${Utilities.formatDate(when, ASIA_TOKYO, "yyyy-MM-dd'T'HH:mm:ssXXX")}`,
        `submission_id=${submissionId}`,
        `form_id=${meta.form_name || ''}`,
        `line_id=${meta.line_id || ''}`,
        `email=${email || ''}`,
        `email_hash=${emailH || ''}`,
        `staged=${meta.line_id ? 'false' : email ? 'email' : 'true'}`,
        `content_hash=${cHash}`,
      ].join('\n')
    );

    appendSubmissionLog_({
      line_id: meta.line_id || '',
      email: pickEmail(meta.fields) || '',
      submission_id: submissionId,
      form_name: meta.form_name || '',
      submit_ym: submitYm,
      doc_code: typeCode,
      period_yyyymm: period || '',
      drive_file_id: file.getId(),
      file_name: file.getName(),
      folder_path: file.getParents().hasNext() ? file.getParents().next().getName() : '',
      content_hash: cHash || '', // 既に計算済みならその値、無ければ空でも動く
    });

    upsertContactLegacy_({
      line_id: meta.line_id || '',
      email: pickEmail(meta.fields) || '',
      display_name: display || '',
      last_form: meta.form_name || '',
      last_submit_ym: submitYm,
    });

    saved.push({
      id: file.getId(),
      name: file.getName(),
      size: att.getSize(),
      folderId: file.getParents().hasNext() ? file.getParents().next().getId() : '',
      doc_type: typeName,
      doc_code: typeCode,
      period_yyyymm: period || '',
    });
  });

  // 7) JSON保存：案件フォルダ直下へ <formkey>__<submissionId>.json（LINEなし時は従来どおり）
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  let jsonParent;
  let folderIdForReturn = '';
  const caseIdForJson = resolveCaseId_(body, meta.subject || '', meta.line_id || '');
  if (meta.line_id && caseIdForJson) {
    const cf = ensureCaseFolder_(meta.line_id, caseIdForJson);
    jsonParent = cf;
    folderIdForReturn = cf.getId();
  } else {
    // LINEなし時は、最初の保存先フォルダ（email staging or _staging）に置く
    const first = saved[0];
    jsonParent = first
      ? DriveApp.getFolderById(first.folderId)
      : ensureStagingPath_(submitYm, submissionId, '未分類');
    folderIdForReturn = first ? first.folderId : jsonParent.getId();
  }
  const jsonFile = saveSubmissionJsonShallow_(
    jsonParent,
    submissionId,
    body,
    meta.fields,
    meta.subject || '',
    meta.line_id || ''
  );

  return { folderId: folderIdForReturn, savedFiles: saved, jsonId: jsonFile.getId() };
}

/********** 台帳＆顧客表 更新 **********/
function updateLedgers(meta, saved) {
  const ss = openOrCreateMaster_();

  // submissions
  const shSub = getOrCreateSheet(ss, 'submission_logs', [
    'ts_saved',
    'line_id',
    'form_name',
    'submission_id',
    'seq',
    'saved_file_ids',
    'json_id',
  ]);
  shSub.appendRow([
    fmt(new Date()),
    meta.line_id,
    meta.form_name,
    normalizeSubmissionIdStrict_(meta.submission_id) || String(Date.now()),
    meta.seq,
    (saved.savedFiles || []).map((s) => s.id).join(','),
    saved.jsonId,
  ]);

  // customers（display_name は今回 fields から推定：候補ラベルを優先順で）
  const nameCandidates = ['名前', '氏名', 'お名前', 'フルネーム'];
  let display = '';
  for (const key of nameCandidates) {
    if (meta.fields[key]) {
      display = meta.fields[key];
      break;
    }
  }
  const shCus = getOrCreateSheet(ss, 'customers', [
    'line_id',
    'display_name',
    'first_seen_at',
    'last_seen_at',
    'last_form',
  ]);
  const lastRow = shCus.getLastRow();
  const all = lastRow > 1 ? shCus.getRange(2, 1, lastRow - 1, 5).getValues() : [];
  let found = false;
  for (let i = 0; i < all.length; i++) {
    if (String(all[i][0]) === meta.line_id) {
      // 更新
      shCus.getRange(i + 2, 2).setValue(display || all[i][1]); // display_name
      shCus.getRange(i + 2, 4).setValue(fmt(new Date())); // last_seen_at
      shCus.getRange(i + 2, 5).setValue(meta.form_name); // last_form
      found = true;
      break;
    }
  }
  if (!found) {
    shCus.appendRow([
      meta.line_id,
      display || '',
      fmt(new Date()),
      fmt(new Date()),
      meta.form_name,
    ]);
  }
}

/********** メイン処理 **********/
function processLabel(labelName) {
  const labelToProcess = GmailApp.getUserLabelByName(labelName);
  if (!labelToProcess) return;

  const labelProcessed =
    GmailApp.getUserLabelByName(LABEL_PROCESSED) || GmailApp.createLabel(LABEL_PROCESSED);
  const labelError = GmailApp.getUserLabelByName(LABEL_ERROR) || GmailApp.createLabel(LABEL_ERROR);

  const threads = labelToProcess.getThreads(0, 50); // 1回で最大50スレッド
  threads.forEach((thread) => {
    // すでに処理済み/エラーのラベルが付いていたら ToProcess を外してスキップ
    const labs = thread.getLabels().map((l) => l.getName());
    const queueLabels = typeof getFormQueueLabels_ === 'function' ? getFormQueueLabels_() : [];
    const lockLabelName =
      typeof getFormLockLabel_ === 'function' ? getFormLockLabel_() : 'BAS/lock';
    if (queueLabels.some((name) => labs.indexOf(name) >= 0) || labs.indexOf(lockLabelName) >= 0) {
      return;
    }
    if (labs.indexOf(LABEL_PROCESSED) >= 0 || labs.indexOf(LABEL_ERROR) >= 0) {
      try {
        thread.removeLabel(labelToProcess);
      } catch (_) {}
      return;
    }

    // スレッドの最後のメッセージのみ見る
    const msgs = thread.getMessages();
    const msg = msgs[msgs.length - 1];

    try {
      const body = msg.getPlainBody() || msg.getBody() || '';
      const metaKV = parseMeta_(body);

      // 通知の安全確認（タグ/secret）。REQUIRE_SECRETで必須/任意を切替
      const guard = checkNotificationGuard_(msg, metaKV);
      if (!guard.ok) {
        Logger.log(
          '[skip->Error] guard fail: hasTag=%s hasSecret=%s subj=%s',
          guard.hasTag,
          guard.hasSecret,
          msg.getSubject()
        );
        thread.removeLabel(labelToProcess).addLabel(labelError);
        return;
      }

      // 以降は既存の保存処理（添付→Drive、JSON保存、OCR、各種台帳更新）
      const parsed = parseMetaAndFields(msg);
      if (parsed.form_name) upsertFormLogRegistry_(parsed.form_name);
      if (!parsed.form_name) parsed.form_name = 'unknown_form';

      const saved = saveAttachmentsAndJson(parsed, msg);
      ocr_processSaved_(saved, parsed);

      thread.removeLabel(labelToProcess).addLabel(labelProcessed);
      Logger.log(
        '[ok] processed thread=%s form=%s',
        thread.getId && thread.getId(),
        parsed.form_name
      );
    } catch (e) {
      Logger.log('[err->Error] %s', (e && e.stack) || e);
      try {
        thread.removeLabel(labelToProcess).addLabel(labelError);
      } catch (_) {}
    }
  });
}

/**
 * Vision API 同期OCR（画像のみ：jpg/jpeg/png）
 * 生成物:
 *  - <元ファイル名>.ocr.txt
 *  - <元ファイル名>.ocr.json
 */
/** 画像(JPG/PNG)にOCRをかけて .ocr.txt / .ocr.json を隣に保存（REST版） */
function ocr_runForImageFile_(file) {
  // 拡張子ではなく MIME で判定（大文字/無拡張でもOK）
  const mt = String(file.getMimeType() || '').toLowerCase();
  if (!/^image\/(png|jpe?g)$/i.test(mt)) return null; // PDF/HEICは別レーン

  // 既に .ocr.txt / .ocr.json があればスキップ
  const parent = file.getParents().hasNext() ? file.getParents().next() : null;
  const base = file.getName();
  if (parent) {
    const hasTxt = parent.getFilesByName(base + '.ocr.txt').hasNext();
    const hasJson = parent.getFilesByName(base + '.ocr.json').hasNext();
    if (hasTxt && hasJson) return 'skipped';
  }

  // Vision OCR 実行
  const body = {
    requests: [
      {
        image: { content: Utilities.base64Encode(file.getBlob().getBytes()) },
        features: [{ type: 'DOCUMENT_TEXT_DETECTION' }],
        imageContext: { languageHints: OCR_LANG_HINTS },
      },
    ],
  };

  const resp = UrlFetchApp.fetch('https://vision.googleapis.com/v1/images:annotate', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  const json = JSON.parse(resp.getContentText());
  const text = json?.responses?.[0]?.fullTextAnnotation?.text || '';

  if (parent) {
    parent.createFile(Utilities.newBlob(text, 'text/plain', base + '.ocr.txt'));
    parent.createFile(
      Utilities.newBlob(JSON.stringify(json, null, 2), 'application/json', base + '.ocr.json')
    );
  }
  return text;
}

/**
 * OCRテキストから給与明細の主要項目をザックリ抽出（PoC用）
 * ※ 後でパターン強化＆正規化予定
 */
function _toHalf(s) {
  return String(s || '').replace(/[！-～]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
}
function _num(s) {
  if (!s) return '';
  s = _toHalf(s)
    .replace(/[,\s　]/g, '')
    .replace(/(\d+)\.(\d{3})(?!\d)/g, '$1$2')
    .replace(/円.*$/, '')
    .replace(/[^\d\-]/g, '');
  return s ? String(parseInt(s, 10)) : '';
}
function _eraToISO(s) {
  s = _toHalf(s);
  const m = s.match(/(令和|平成)\s*(\d{1,2})年\s*(\d{1,2})月\s*(\d{1,2})?日/);
  if (!m) return '';
  const y = (m[1] === '令和' ? 2018 : 1988) + +m[2];
  const pad = (n) => ('0' + n).slice(-2);
  return `${y}-${pad(+m[3])}-${pad(+(m[4] || '1'))}`;
}
function _pickDate(text) {
  const t = text || '';
  let m = t.match(
    /(令和|平成)\s*[0-9０-９]{1,2}年\s*[0-9０-９]{1,2}月\s*[0-9０-９]{1,2}日[^。\n]*支給/
  );
  if (m) {
    const iso = _eraToISO(m[0]);
    if (iso) return iso;
  }
  m = t.match(
    /(支給日|支給年月日)[：:\s]*([0-9０-９]{4}[/\-年][0-9０-９]{1,2}[/\-月][0-9０-９]{1,2}日?)/
  );
  if (m) {
    return _toHalf(m[2])
      .replace(/[年月]/g, '-')
      .replace(/日/g, '')
      .replace(/\/+/g, '-');
  }
  m = t.match(/(令和|平成)\s*[0-9０-９]{1,2}年\s*[0-9０-９]{1,2}月\s*[0-9０-９]{1,2}日/);
  if (m) {
    const iso = _eraToISO(m[0]);
    if (iso) return iso;
  }
  return '';
}
function ocr_extractPayslip_(text) {
  const t = text || '';
  const pick = (alts) => {
    for (const re of alts) {
      const m = t.match(re);
      if (m) return (m[2] || m[1] || '').trim();
    }
    return '';
  };
  const obj = {
    company: pick([/(会社名|勤務先名|事業所名)[：:]\s*([^\n]+)/i, /^(.+?株式会社)[^\n]*$/m]),
    employee: pick([/(氏名|従業員名|社員名|お名前)[：:]\s*([^\n]+)/i]),
    payday: _pickDate(t),
    gross: _num(pick([/(実総支給額|総支給額|総支給|支給合計)[：:\s]*([0-9０-９,，\.]+)円?/i])),
    deduction: _num(pick([/(総控除額|控除合計|控除額)[：:\s]*([0-9０-９,，\.]+)円?/i])),
    net: _num(pick([/(差引支給額|差引額|手取|支給額)[：:\s]*([0-9０-９,，\.]+)円?/i])),
    baseSalary: _num(pick([/(基本給)[：:\s]*([0-9０-９,，\.]+)円?/i])),
    commute: _num(
      pick([/(通勤手当|通勤費|通勤手当\(課\)|通勤手当\(非\))[：:\s]*([0-9０-９,，\.]+)円?/i])
    ),
    healthIns: _num(pick([/(健康保険|健保)[：:\s]*([0-9０-９,，\.]+)円?/i])),
    pension: _num(pick([/(厚生年金|年金)[：:\s]*([0-9０-９,，\.]+)円?/i])),
    tax: _num(pick([/(源泉所得税|所得税)[：:\s]*([0-9０-９,，\.]+)円?/i])),
    residentTax: _num(pick([/(住民税)[：:\s]*([0-9０-９,，\.]+)円?/i])),
  };
  if (obj.payday && +obj.payday.slice(0, 4) < 2015) obj.payday = ''; // 明らかな誤年を弾く
  return obj;
}

/**
 * 抽出結果を <元ファイル名>.extracted.json として保存
 */
function ocr_saveExtraction_(file, extracted) {
  if (!extracted) return;
  const parent = file.getParents().hasNext() ? file.getParents().next() : null;
  if (!parent) return;
  parent.createFile(
    Utilities.newBlob(
      JSON.stringify(extracted, null, 2),
      'application/json',
      file.getName() + '.extracted.json'
    )
  );
}

/**
 * 保存直後のファイル群に対しOCR→抽出まで実行
 * @param {Object} saved saveAttachmentsAndJson の戻り値
 * @param {Object} meta  parseMetaAndFields の結果
 */
function ocr_processSaved_(saved, meta) {
  try {
    const form = String(meta.form_name || '');
    if (OCR_TARGET_FORMS.length > 0 && !OCR_TARGET_FORMS.some((k) => form.includes(k))) return;

    (saved.savedFiles || []).forEach((s) => {
      const file = DriveApp.getFileById(s.id);
      const parent = file.getParents().hasNext() ? file.getParents().next() : null;
      if (!parent) return;

      // 既に _model.json があればスキップ（再生成したい時は削除してから）
      if (parent.getFilesByName(s.id + '_model.json').hasNext()) return;

      // まずOCRを試みる（既存あれば 'skipped'）
      let text = ocr_runForImageFile_(file);

      // skippedなら既存の .ocr.txt を読み込む（なければ空でOK→画像のみで抽出）
      if (text === 'skipped') {
        const it = parent.getFilesByName(file.getName() + '.ocr.txt');
        text = it.hasNext() ? it.next().getBlob().getDataAsString('utf-8') : '';
      }

      // （任意）PoC抽出：text があれば .extracted.json を更新
      if (text) {
        const extracted = ocr_extractPayslip_(text);
        ocr_saveExtraction_(file, extracted);
      }

      // 画像URLを公開にして抽出APIにPOST（ocrTextは空でもOK）
      const imageUrl = ensurePublicImageUrl_(file);
      const out = postExtract_(s.id, imageUrl, text || '', meta.line_id);

      // 成果物保存
      if (out && out.ok && out.data) {
        parent.createFile(
          Utilities.newBlob(
            JSON.stringify(out.data, null, 2),
            'application/json',
            s.id + '_model.json'
          )
        );
      } else {
        parent.createFile(
          Utilities.newBlob(
            `status=${out?.status}\n${out?.error || ''}\n\n${out?.raw || ''}`,
            'text/plain',
            s.id + '_extract_error.txt'
          )
        );
      }
    });
  } catch (e) {
    console.error('OCR処理エラー:', e);
  }
}

/********** エントリ **********/
function cron_1min() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) return;
  try {
    if (typeof run_ProcessInbox_AllForms === 'function') {
      run_ProcessInbox_AllForms();
    } else {
      ensureLabels();
      processLabel(LABEL_TO_PROCESS);
    }
  } catch (e) {
    console.error('cron_1min failed:', e);
  } finally {
    lock.releaseLock();
  }
}

function pickDisplayName(fields) {
  const keys = ['名前', '氏名', 'お名前', 'フルネーム', 'Name'];
  for (const k of keys) {
    if (fields && fields[k]) return sanitize(fields[k]);
  }
  return '';
}

/** ====== caseId 導入: ブートストラップ最小実装 ====== */
function json_(obj, code) {
  const out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  if (typeof out.setStatusCode === 'function') {
    out.setStatusCode(code || 200);
    return out;
  }
  return out; // Apps Script では setStatusCode が使えない WebApp 環境があるため冪等
}

function verifySig_(base, sigHex) {
  const secret = PropertiesService.getScriptProperties().getProperty('BOOTSTRAP_SECRET') || '';
  const raw = Utilities.computeHmacSha256Signature(String(base || ''), secret);
  const hex = raw.map((b) => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  return hex === String(sigHex || '');
}

function makeUserKey_(lineId) {
  return String(lineId || '')
    .slice(0, 6)
    .toLowerCase();
}

function ensureUserRootFolder_(displayName, userKey) {
  const parent = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const name = `${sanitize(displayName || '')}__${sanitize(userKey || '')}`;
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) {
    const f = it.next();
    // attachments サブフォルダが無ければ作成
    const att = f.getFoldersByName('attachments');
    if (!att.hasNext()) f.createFolder('attachments');
    return f;
  }
  // attachments を作ってから親に戻る＝親はユーザフォルダ
  return parent.createFolder(name).createFolder('attachments').getParents().next();
}

function upsertContact_(ss, lineId, displayName) {
  const sh = ss.getSheetByName('contacts') || ss.insertSheet('contacts');
  // 新規は snake_case で初期化
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'line_id',
      'display_name',
      'user_key',
      'active_case_id',
    ]);
  }
  const header = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
  let idxMap =
    typeof buildHeaderIndexMap_ === 'function'
      ? buildHeaderIndexMap_(header)
      : header.reduce((m, v, i) => ((m[String(v)] = i), m), {});

  const HEADERS = {};
  HEADERS[K.lineId] = 'line_id';
  HEADERS[K.displayName] = 'display_name';
  HEADERS[K.userKey] = 'user_key';
  HEADERS[K.activeCaseId] = 'active_case_id';
  Object.keys(HEADERS).forEach((canon) => {
    if (idxMap[canon] === undefined) {
      const col = sh.getLastColumn() + 1;
      sh.getRange(1, col).setValue(HEADERS[canon]);
      header.push(HEADERS[canon]);
      idxMap =
        typeof buildHeaderIndexMap_ === 'function'
          ? buildHeaderIndexMap_(header)
          : header.reduce((m, v, i) => ((m[String(v)] = i), m), {});
    }
  });

  const rowCount = sh.getLastRow() - 1;
  const lastCol = sh.getLastColumn();
  const values = rowCount > 0 ? sh.getRange(2, 1, rowCount, lastCol).getValues() : [];

  let targetRow = -1; // 1-based
  const colLine0 = idxMap[K.lineId];
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][colLine0] || '') === String(lineId)) {
      targetRow = i + 2;
      break;
    }
  }

  const userKey = makeUserKey_(lineId);
  if (targetRow === -1) {
    sh.appendRow(new Array(lastCol).fill(''));
    targetRow = sh.getLastRow();
    const root = ensureUserRootFolder_(displayName, userKey);
    setCellByKey_(sh, targetRow, idxMap, K.lineId, lineId);
    setCellByKey_(sh, targetRow, idxMap, K.displayName, displayName || '');
    setCellByKey_(sh, targetRow, idxMap, K.userKey, userKey);
    setCellByKey_(sh, targetRow, idxMap, K.rootFolderId, root.getId());
    setCellByKey_(sh, targetRow, idxMap, K.nextCaseSeq, 0);
    setCellByKey_(sh, targetRow, idxMap, K.activeCaseId, '');
    return {
      lineId,
      displayName,
      userKey,
      rootFolderId: root.getId(),
      nextCaseSeq: 0,
      activeCaseId: '',
    };
  } else {
    const rowVals = sh.getRange(targetRow, 1, 1, lastCol).getValues()[0];
    const currentUserKey = String(getCellByKey_(rowVals, idxMap, K.userKey) || '') || userKey;
    let rootId = String(getCellByKey_(rowVals, idxMap, K.rootFolderId) || '');
    if (!rootId) rootId = ensureUserRootFolder_(displayName, currentUserKey).getId();
    if (!getCellByKey_(rowVals, idxMap, K.displayName) && displayName) {
      setCellByKey_(sh, targetRow, idxMap, K.displayName, displayName);
    }
    setCellByKey_(sh, targetRow, idxMap, K.userKey, currentUserKey);
    setCellByKey_(sh, targetRow, idxMap, K.rootFolderId, rootId);
    return {
      lineId,
      displayName: String(getCellByKey_(rowVals, idxMap, K.displayName) || displayName || ''),
      userKey: currentUserKey,
      rootFolderId: rootId,
      nextCaseSeq: Number(getCellByKey_(rowVals, idxMap, K.nextCaseSeq) || 0) || 0,
      activeCaseId: String(getCellByKey_(rowVals, idxMap, K.activeCaseId) || ''),
    };
  }
}

function setActiveCaseId_(ss, lineId, caseId) {
  const sh = ss.getSheetByName('contacts');
  if (!sh) return;
  const header = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
  let idxMap =
    typeof buildHeaderIndexMap_ === 'function'
      ? buildHeaderIndexMap_(header)
      : header.reduce((m, v, i) => ((m[String(v)] = i), m), {});
  if (idxMap[K.lineId] === undefined) {
    const col = sh.getLastColumn() + 1;
    sh.getRange(1, col).setValue('line_id');
    header.push('line_id');
    idxMap =
      typeof buildHeaderIndexMap_ === 'function'
        ? buildHeaderIndexMap_(header)
        : header.reduce((m, v, i) => ((m[String(v)] = i), m), {});
  }
  if (idxMap[K.activeCaseId] === undefined) {
    const col = sh.getLastColumn() + 1;
    sh.getRange(1, col).setValue('active_case_id');
    header.push('active_case_id');
    idxMap =
      typeof buildHeaderIndexMap_ === 'function'
        ? buildHeaderIndexMap_(header)
        : header.reduce((m, v, i) => ((m[String(v)] = i), m), {});
  }
  const rowCount = sh.getLastRow() - 1;
  if (rowCount < 1) return;
  const lastCol = sh.getLastColumn();
  const rows = sh.getRange(2, 1, rowCount, lastCol).getValues();
  const colLine0 = idxMap[K.lineId];
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][colLine0] || '') === String(lineId)) {
      setCellByKey_(sh, i + 2, idxMap, K.activeCaseId, caseId);
      return;
    }
  }
}

function casesAppend_(lineId, caseId) {
  const ss = openOrCreateMaster_();
  const sh = ss.getSheetByName('cases') || ss.insertSheet('cases');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['line_id', 'case_id', 'created_at', 'status', 'last_activity']);
  }
  const header = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
  const idxMap =
    typeof buildHeaderIndexMap_ === 'function'
      ? buildHeaderIndexMap_(header)
      : header.reduce((m, v, i) => ((m[String(v)] = i), m), {});
  sh.appendRow(new Array(sh.getLastColumn()).fill(''));
  const row1 = sh.getLastRow();
  const now = new Date();
  setCellByKey_(sh, row1, idxMap, K.lineId, lineId);
  setCellByKey_(sh, row1, idxMap, K.caseId, caseId);
  setCellByKey_(sh, row1, idxMap, K.createdAt, now);
  setCellByKey_(sh, row1, idxMap, K.status, 'draft');
  setCellByKey_(sh, row1, idxMap, K.lastActivity, now);
}

function resolveCaseId_(mailPlainBody, subject, lineId) {
  const meta = parseMetaBlock_(mailPlainBody);
  if (meta.case_id) return String(meta.case_id).trim();
  const ss = openOrCreateMaster_();
  const sh = ss.getSheetByName('contacts');
  const header = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
  const idxMap =
    typeof buildHeaderIndexMap_ === 'function'
      ? buildHeaderIndexMap_(header)
      : header.reduce((m, v, i) => ((m[String(v)] = i), m), {});
  const rowCount = sh.getLastRow() - 1;
  if (rowCount < 1) return '';
  const rows = sh.getRange(2, 1, rowCount, sh.getLastColumn()).getValues();
  const colLine0 = idxMap[K.lineId];
  const colActive0 = idxMap[K.activeCaseId];
  if (colLine0 === undefined || colActive0 === undefined) return '';
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][colLine0] || '') === String(lineId)) {
      return String(rows[i][colActive0] || '');
    }
  }
  return '';
}

function doPost_drive(e) {
  try { Logger.log('[DEPRECATED] doPost_drive called'); } catch (_) {}
  return ContentService.createTextOutput(
    JSON.stringify({
      ok: false,
      error: 'deprecated',
      hint: 'drive endpoint retired. Use intake/bootstrap flow.',
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

function inferDisplayNameFromLedger_(lineId) {
  try {
    const ss = openOrCreateMaster_();
    const sh = ss.getSheetByName('customers');
    if (!sh) return '';
    const vals = sh.getDataRange().getValues(); // [line_id, display_name, ...]
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === String(lineId)) return sanitize(vals[i][1] || '');
    }
  } catch (_) {}
  return '';
}

function getOrCreateUserFolder(root, lineId, displayName) {
  const base = sanitize(lineId || 'unknown');

  // displayName が空なら台帳から補完を試みる
  let disp = sanitize(displayName || '') || inferDisplayNameFromLedger_(base);
  const named = disp ? (disp + '__' + base).slice(0, FOLDER_NAME_MAX) : '';

  // 1) 完全一致をまず探す（named → base）
  if (named) {
    const itNamed = root.getFoldersByName(named);
    if (itNamed.hasNext()) return itNamed.next();
  }
  const itBase = root.getFoldersByName(base);
  if (itBase.hasNext()) {
    const f = itBase.next();
    // 氏名が分かっていて、方針が 'latest' ならリネーム
    if (disp && RENAME_STRATEGY === 'latest') {
      if (!root.getFoldersByName(named).hasNext()) f.setName(named);
    }
    return f;
  }

  // 2) "*__lineId" を総当たりで探す（フォームB→Aパターン対策）
  //   ※ フォルダ数が極端に多い環境で重ければ early return に変更してください
  const all = root.getFolders();
  let match = null;
  while (all.hasNext()) {
    const f = all.next();
    const nm = f.getName();
    if (nm.endsWith('__' + base)) {
      match = f;
      break;
    }
  }
  if (match) {
    // 氏名が取れていて 'latest' なら名前を更新
    if (disp && RENAME_STRATEGY === 'latest') {
      const target = named;
      if (target && match.getName() !== target && !root.getFoldersByName(target).hasNext()) {
        match.setName(target);
      }
    }
    return match;
  }

  // 3) どれも無ければ新規作成（氏名があれば named、無ければ base）
  return root.createFolder(named || base);
}

/** /root/<氏名__LINEID>/<日本語書類名>/ を返す（年月フォルダなし） */
function ensurePathJapanese_(lineId, displayName, docTypeName, submitYm) {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const userFolder = getOrCreateUserFolder(root, lineId || 'unknown', displayName || '');
  const typeFolder = getOrCreateFolder(userFolder, docTypeName || '未分類');
  return typeFolder; // 年月フォルダは廃止（浅い構造）
}

/** =========================
 *  BAS_master（状態管理ブック）
 *  contacts / submissions / form_logs をこの1冊に集約
 *  ========================= */
// ❶ マスターSpreadsheetを開く（必ず Spreadsheet を返す）
function openOrCreateMaster_() {
  const sid = PropertiesService.getScriptProperties().getProperty('BAS_MASTER_SPREADSHEET_ID');
  if (!sid) throw new Error('BAS_MASTER_SPREADSHEET_ID is empty');

  const spreadsheet = SpreadsheetApp.openById(sid);
  if (!spreadsheet || typeof spreadsheet.getSheetByName !== 'function') {
    throw new Error('invalid_spreadsheet_id: ' + sid);
  }

  // 必要シートの最低限を用意（足りなければ作る）
  ['contacts', 'cases', 'submissions', 'form_logs', 'logs', 'customers'].forEach(function (n) {
    if (!spreadsheet.getSheetByName(n)) spreadsheet.insertSheet(n);
  });

  return spreadsheet; // ←ここが超重要：Sheetを返さない
}

// ❷ 台帳系も Spreadsheet を返す（必要なら中でシートを用意）
function openOrCreateLedger() {
  // もし別IDならプロパティ名に合わせて取得
  const sid = PropertiesService.getScriptProperties().getProperty('BAS_MASTER_SPREADSHEET_ID');
  if (!sid) throw new Error('BAS_MASTER_SPREADSHEET_ID is empty');

  const spreadsheet = SpreadsheetApp.openById(sid);
  if (!spreadsheet || typeof spreadsheet.getSheetByName !== 'function') {
    throw new Error('invalid_spreadsheet_id: ' + sid);
  }

  // 必要なら 'ledger' シートを作成しておく
  if (!spreadsheet.getSheetByName('ledger')) spreadsheet.insertSheet('ledger');

  return spreadsheet; // ←SheetではなくSpreadsheetを返す
}

/** contacts を upsert（line_id 優先、なければ email） */
function upsertContactLegacy_({ line_id, email, display_name, last_form, last_submit_ym }) {
  const ss = openOrCreateMaster_();
  const sh = ss.getSheetByName('contacts');
  const vals = sh.getDataRange().getValues();
  const now = Utilities.formatDate(new Date(), ASIA_TOKYO, 'yyyy-MM-dd HH:mm:ss');
  const e = normalizeEmail_(email || '');
  const eh = emailHash_(e);

  // 既存行探索：line_id → email
  let row = -1;
  for (let i = 1; i < vals.length; i++) {
    if (line_id && String(vals[i][0]) === String(line_id)) {
      row = i + 1;
      break;
    }
  }
  if (row < 0 && e) {
    for (let i = 1; i < vals.length; i++) {
      if (normalizeEmail_(vals[i][1]) === e) {
        row = i + 1;
        break;
      }
    }
  }

  if (row < 0) {
    sh.appendRow([
      line_id || '',
      e,
      eh,
      display_name || '',
      now,
      now,
      last_form || '',
      last_submit_ym || '',
      '',
    ]);
  } else {
    if (line_id && !vals[row - 1][0]) sh.getRange(row, 1).setValue(line_id);
    if (e && normalizeEmail_(vals[row - 1][1]) !== e) sh.getRange(row, 2).setValue(e);
    if (eh && vals[row - 1][2] !== eh) sh.getRange(row, 3).setValue(eh);
    if (display_name && (!vals[row - 1][3] || vals[row - 1][3] !== display_name))
      sh.getRange(row, 4).setValue(display_name);
    sh.getRange(row, 6).setValue(now); // last_seen_at
    if (last_form) sh.getRange(row, 7).setValue(last_form);
    if (last_submit_ym) sh.getRange(row, 8).setValue(last_submit_ym);
  }
}

/** submissions へ1行追記 */
function appendSubmissionLog_({
  line_id,
  email,
  submission_id,
  form_name,
  submit_ym,
  doc_code,
  period_yyyymm,
  drive_file_id,
  file_name,
  folder_path,
  content_hash,
}) {
  const ss = openOrCreateMaster_();
  const sh = getOrCreateSheet(ss, 'submission_logs', [
    'ts_saved',
    'line_id',
    'email',
    'submission_id',
    'form_name',
    'submit_ym',
    'doc_code',
    'period_yyyymm',
    'drive_file_id',
    'file_name',
    'folder_path',
    'content_hash',
  ]);
  const now = Utilities.formatDate(new Date(), ASIA_TOKYO, 'yyyy-MM-dd HH:mm:ss');
  const normalizedSid = normalizeSubmissionIdStrict_(submission_id);
  let sid = normalizedSid;
  if (!sid) {
    sid = String(Date.now());
    try {
      Logger.log('[email-intake] sid_fallback(log) from "%s" -> "%s"', submission_id || '', sid);
    } catch (_) {}
  }
  sh.appendRow([
    now,
    line_id || '',
    normalizeEmail_(email || ''),
    sid,
    form_name || '',
    submit_ym || '',
    doc_code || '',
    period_yyyymm || '',
    drive_file_id || '',
    file_name || '',
    folder_path || '',
    content_hash || '',
  ]);
}

function sweepInvalidSubmissionRows_() {
  const ss = openOrCreateMaster_();
  const sh = ss.getSheetByName('submission_logs');
  if (!sh) return;
  const range = sh.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return;
  const header = values[0];
  const indexMap = header.reduce(function (acc, key, idx) {
    acc[String(key)] = idx;
    return acc;
  }, {});
  const sidIdx =
    indexMap.submission_id != null
      ? indexMap.submission_id
      : indexMap.submissionId != null
      ? indexMap.submissionId
      : null;
  if (sidIdx == null) return;
  for (let r = values.length - 1; r >= 1; r--) {
    const sid = String(values[r][sidIdx] || '').trim();
    if (!/^(ack:[\w:-]+|\d{3,})$/.test(sid)) {
      sh.deleteRow(r + 1);
    }
  }
}

/** 指定フォルダ以下のツリーをログに出力 */
function debug_showBasTree() {
  logFolderTree(ROOT_FOLDER_ID);
}

function logFolderTree(folderId, depth = 0) {
  const folder = DriveApp.getFolderById(folderId);
  const prefix = '  '.repeat(depth);
  Logger.log(prefix + '📁 ' + folder.getName());

  // ファイル
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    Logger.log(prefix + '  📄 ' + f.getName());
  }

  // サブフォルダ
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    const sub = subs.next();
    logFolderTree(sub.getId(), depth + 1);
  }
}

function ping_() {
  Logger.log('BAS ping OK');
}

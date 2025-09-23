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
  const query = `label:${S2002_LABEL_TO_PROCESS} subject:#FM-BAS subject:S2002`;
  const threads = GmailApp.search(query, 0, 50);
  try {
    Logger.log('[S2002] tick: query=%s threads=%s', query, threads.length);
  } catch (_) {}
  if (!threads.length) return;

  threads.forEach((th) => {
    const msgs = th.getMessages();
    msgs.forEach((msg) => {
      try {
        const body = msg.getPlainBody() || HtmlService.createHtmlOutput(msg.getBody()).getContent();
        const subject = msg.getSubject();
        const parsed = parseFormMail_(subject, body);

        if (parsed.meta.form_key !== 's2002_userform') return; // S2002のみ処理

        // cases解決（caseId→caseKey/folderId取得）
        const caseInfo = resolveCaseByCaseId_(parsed.meta.case_id);
        if (!caseInfo) throw new Error(`Unknown case_id: ${parsed.meta.case_id}`);
        // folderId が未設定/URL等なら補正（caseKey から検索/作成）
        caseInfo.folderId = ensureCaseFolderId_(caseInfo);

        // 1) JSON保存（<caseFolder> 直下に保存）
        saveSubmissionJson_(caseInfo.folderId, parsed);

        // 2) S2002 ドラフト生成（gdocコピー→差し込み→drafts保存）
        const draft = generateS2002Draft_(caseInfo, parsed);
        try {
          Logger.log(
            '[S2002] draft created: caseId=%s url=%s',
            parsed.meta.case_id,
            draft && draft.draftUrl
          );
        } catch (_) {}

        // 3) ステータス更新
        updateCasesRow_(parsed.meta.case_id, {
          status: 'draft',
          lastActivity: new Date(),
          lastDraftUrl: draft.draftUrl,
        });

        // ラベル付替え
        msg.addLabel(GmailApp.getUserLabelByName(S2002_LABEL_PROCESSED));
        msg.removeLabel(GmailApp.getUserLabelByName(S2002_LABEL_TO_PROCESS));
      } catch (e) {
        // 失敗時はスレッドにノートを残す
        GmailApp.createDraft(msg.getFrom(), '[BAS Intake Error]', String(e), {
          htmlBody: `<pre>${safeHtml(e.stack || e)}</pre>`,
        });
      }
    });
  });
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
    lastActivity: new Date(),
    lastDraftUrl: draft.draftUrl,
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
  const idxFolder = header.indexOf('folderId');
  const idxCaseId = header.indexOf('caseId');
  let caseInfo = null;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][idxFolder]) === String(caseFolderId)) {
      caseInfo = {
        caseId: vals[i][idxCaseId],
        caseKey: vals[i][header.indexOf('caseKey')],
        lineId: vals[i][header.indexOf('lineId')],
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
    lastActivity: new Date(),
    lastDraftUrl: draft.draftUrl,
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
  updateCasesRow_(caseId, { folderId: String(folderId || '').trim() });
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
    const v = value.trim();
    out.ref.self_employed = v === 'いいえ' ? 'none' : v.includes('6') ? 'past6m' : 'current';
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
  return String(v).trim() === 'はい';
}
function toIsoBirth_(ja) {
  // 例: 2000年01月01日 / 2000/1/1 / 2000.1.1 など
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
  const d = new Date(iso);
  const eras = [
    { gengo: '令和', start: new Date('2019-05-01'), offset: 2018 },
    { gengo: '平成', start: new Date('1989-01-08'), offset: 1988 },
    { gengo: '昭和', start: new Date('1926-12-25'), offset: 1925 },
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
  return normJaSpace_(sanitizeAltAddressValue_(s));
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
  const idx = {
    caseId: header.indexOf('caseId'),
    lineId: header.indexOf('lineId'),
    caseKey: header.indexOf('caseKey'),
    folderId: header.indexOf('folderId'),
  };
  const want = normalizeCaseIdString_(caseId);
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    const rowCase = normalizeCaseIdString_(row[idx.caseId]);
    if (rowCase && rowCase === want) {
      return {
        caseId: rowCase,
        lineId: row[idx.lineId],
        caseKey: row[idx.caseKey],
        folderId: row[idx.folderId],
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
  const idxCase = header.indexOf('caseId');
  const want = normalizeCaseIdString_(caseId);
  const rowIdx = vals.findIndex((r, i) => i > 0 && normalizeCaseIdString_(r[idxCase]) === want);
  if (rowIdx < 1) return;
  const r = rowIdx + 1;
  Object.keys(patch).forEach((k) => {
    const c = header.indexOf(k);
    if (c >= 0) sh.getRange(r, c + 1).setValue(patch[k]);
  });
}

function saveSubmissionJson_(caseFolderId, parsed) {
  const parent = DriveApp.getFolderById(caseFolderId);
  const fname = `${parsed.meta.form_key}__${parsed.meta.submission_id || Date.now()}.json`;
  const blob = Utilities.newBlob(JSON.stringify(parsed, null, 2), 'application/json', fname);
  parent.createFile(blob); // 仕様: ケース直下に保存
  return `${caseFolderId}/${fname}`;
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
          updateCasesRow_(caseInfo.caseId, { folderId: better.id, caseKey: nameGuess });
          return better.id;
        }
        // 現行のIDで続行
        if (norm !== caseInfo.folderId) updateCasesRow_(caseInfo.caseId, { folderId: norm });
        return norm;
      }
    } catch (_) {
      // 続行して補正
    }
  }

  // 2) caseKey から Drive ルート直下で検索/作成
  const props = PropertiesService.getScriptProperties();
  const ROOT_ID =
    props.getProperty('DRIVE_ROOT_FOLDER_ID') || props.getProperty('ROOT_FOLDER_ID') || '';
  if (!ROOT_ID) throw new Error('ROOT_FOLDER_ID/DRIVE_ROOT_FOLDER_ID が未設定です');
  const root = DriveApp.getFolderById(ROOT_ID);
  let name = caseInfo.caseKey || '';
  if (!name) {
    // cases に caseKey 列が無い場合は lineId+caseId から推定（userKey = lineId 先頭6文字）
    const lid = String(caseInfo.lineId || '').trim();
    const cid = normalizeCaseIdString_(caseInfo.caseId || '');
    const userKey = lid ? lid.slice(0, 6).toLowerCase() : '';
    if (userKey && cid) name = userKey + '-' + cid;
  }
  if (!name)
    throw new Error(
      'caseKey が無く、lineId からも生成できません（cases に caseKey か lineId 列が必要）'
    );
  // 重複フォルダが複数存在する可能性があるため、内容がある方を優先
  const best = findBestCaseFolderUnderRoot_(name);
  const id = best ? best.id : root.createFolder(name).getId();
  // folderId は必ず書き戻し。caseKey 列が存在すれば caseKey も書き戻される（無ければスキップ）。
  updateCasesRow_(caseInfo.caseId, { folderId: id, caseKey: name });
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
  const props = PropertiesService.getScriptProperties();
  const ROOT_ID =
    props.getProperty('DRIVE_ROOT_FOLDER_ID') || props.getProperty('ROOT_FOLDER_ID') || '';
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

/**
 * フォルダの「それらしさ」を採点。
 * s2002_userform__*.json: +5（≥1件で加点）
 * 任意の直下 .json: +2（≥1件で加点）
 * drafts/ サブフォルダ: +3
 * attachments/ or staff_inputs/ サブフォルダ: +1 ずつ
 */
function scoreCaseFolder_(folder) {
  let s2 = 0;
  let anyJson = 0;
  const itF = folder.getFiles();
  let scanned = 0;
  while (itF.hasNext() && scanned < 500) {
    const f = itF.next();
    scanned++;
    const n = f.getName && f.getName();
    if (!n) continue;
    if (/\.json$/i.test(n)) anyJson = 1;
    if (/^s2002_userform__/i.test(String(n).trim())) s2 = 1;
  }
  let drafts = 0,
    attach = 0,
    staff = 0;
  const itD = folder.getFolders();
  while (itD.hasNext()) {
    const d = itD.next();
    const dn = d.getName && d.getName();
    if (dn === 'drafts') drafts = 1;
    if (dn === 'attachments') attach = 1;
    if (dn === 'staff_inputs') staff = 1;
  }
  const score = s2 * 5 + anyJson * 2 + drafts * 3 + attach * 1 + staff * 1;
  return { score, s2, anyJson, drafts, attach, staff };
}

/** ====== S2002 ドラフト生成（gdoc） ====== **/

function generateS2002Draft_(caseInfo, parsed) {
  if (!S2002_TPL_GDOC_ID) throw new Error('S2002_TEMPLATE_GDOC_ID not set');
  // generateS2002Draft_ 内
  const drafts = getOrCreateSubfolder_(DriveApp.getFolderById(caseInfo.folderId), 'drafts');
  try {
    Logger.log('[S2002] draftsFolderId=%s caseFolderId=%s', drafts.getId(), caseInfo.folderId);
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

  // 差し込み（最低限：ここから増やす）
  const M = parsed.model;

  // 申立人情報
  replaceAll_(body, '{{app.name}}', M.app.name || '');
  replaceAll_(body, '{{app.kana}}', M.app.kana || '');
  replaceAll_(body, '{{app.maiden_name}}', M.app.maiden_name || '');
  replaceAll_(body, '{{app.maiden_kana}}', M.app.maiden_kana || '');
  replaceAll_(body, '{{app.nationality}}', M.app.nationality || '');
  replaceAll_(body, '{{app.phone}}', M.app.phone || '');
  replaceAll_(body, '{{app.age}}', (M.app.age ?? '') + '');

  // 生年月日（西暦/元号）
  if (M.app.birth_iso) {
    const [yy, mm, dd] = M.app.birth_iso.split('-');
    replaceAll_(body, '{{app.birth_yyyy}}', yy || '');
    replaceAll_(body, '{{app.birth_mm}}', String(mm || ''));
    replaceAll_(body, '{{app.birth_dd}}', String(dd || ''));
  }
  // ★ここに追加（旧: if (M.app.birth_wareki) { ... } は削除）
  const w = M.app.birth_wareki || null;
  replaceAll_(body, '{{app.birth_wareki_gengo}}', (w && w.gengo) || '');
  replaceAll_(body, '{{app.birth_wareki_yy}}', (w && String(w.yy)) || '');

  // 住所
  const altClean0 = cleanAltAddress_(M.addr.alt_full);
  let altOut = altClean0;
  if (!altOut && M.addr && M.addr.same_as_resident === false) {
    // 住民票と異なるのに空 → 通常住所で埋める（空出力回避）
    altOut = cleanAltAddress_(M.addr.full);
  }
  try {
    Logger.log(
      '[S2002] addr debug: same_as=%s alt_raw="%s" alt_out="%s"',
      M.addr && M.addr.same_as_resident,
      M.addr && M.addr.alt_full,
      altOut
    );
  } catch (_) {}

  replaceAll_(body, '{{addr.postal}}', M.addr.postal || '');
  replaceAll_(body, '{{addr.pref_city}}', M.addr.pref_city || '');
  replaceAll_(body, '{{addr.street}}', M.addr.street || '');
  replaceAll_(body, '{{addr.alt_full}}', altOut || '');
  replaceAll_(body, '{{addr.same_as_resident}}', renderCheck_(!!M.addr.same_as_resident));
  replaceAll_(body, '{{addr.same_as_resident_no}}', renderCheck_(!M.addr.same_as_resident));

  // 本籍・国籍（本籍は固定チェック、国籍は入力があれば☑）
  replaceAll_(body, '{{ref.domicile_resident}}', renderCheck_(true));
  replaceAll_(body, '{{ref.nationality}}', renderCheck_(!!M.app.nationality));

  // 個人事業者か（3択）
  replaceAll_(body, '{{ref.self_employed_none}}', renderCheck_(M.ref.self_employed === 'none'));
  replaceAll_(body, '{{ref.self_employed_past6m}}', renderCheck_(M.ref.self_employed === 'past6m'));
  replaceAll_(
    body,
    '{{ref.self_employed_current}}',
    renderCheck_(M.ref.self_employed === 'current')
  );

  // 法人代表者（有／無）
  replaceAll_(body, '{{ref.corp_representative}}', renderCheck_(!!M.ref.corp_representative));
  replaceAll_(body, '{{ref.corp_representative_no}}', renderCheck_(!M.ref.corp_representative));

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
  replaceAll_(body, '{{ref.welfare}}', renderCheck_(!!M.ref.welfare));
  replaceAll_(body, '{{ref.welfare_no}}', renderCheck_(!M.ref.welfare));

  // 個人再生7年内（有／無）
  replaceAll_(body, '{{ref.pr_rehab_within7y}}', renderCheck_(!!M.ref.pr_rehab_within7y));
  replaceAll_(body, '{{ref.pr_rehab_within7y_no}}', renderCheck_(!M.ref.pr_rehab_within7y));

  // 免責7年内（有／無）
  replaceAll_(
    body,
    '{{ref.bankruptcy_discharge_within7y}}',
    renderCheck_(!!M.ref.bankruptcy_discharge_within7y)
  );
  replaceAll_(
    body,
    '{{ref.bankruptcy_discharge_within7y_no}}',
    renderCheck_(!M.ref.bankruptcy_discharge_within7y)
  );

  // 文末
  doc.saveAndClose();
  return { gdocId, draftUrl: doc.getUrl() };
}

/** ====== 置換ユーティリティ ====== **/
function replaceAll_(body, token, value) {
  const safe = token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  body.replaceText(safe, String(value ?? ''));
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

// 使うとき（S2002の {{addr.alt_full}} 差し込み直前だけ）
const altOut = sanitizeAltAddressValue_S2002_(M.addr.alt_full);
replaceAll_(body, '{{addr.alt_full}}', altOut || '');

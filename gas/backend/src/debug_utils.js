// @ts-nocheck
/**
 * debug_utils.gs
 * 本番コードから切り出したデバッグ用関数群（固定IDは使わない）
 * - 破壊的操作なし（閲覧/生成のみ）
 * - fileId か Drive URL を毎回手入力/引数で渡す
 */

/********** 共通ユーティリティ（安全） **********/
function parseDriveFileId(input) {
  var s = String(input || '').trim();
  if (/^[A-Za-z0-9_-]{20,}$/.test(s)) return s;
  var m = s.match(/\/file\/d\/([^/]+)/);
  if (m) return m[1];
  m = s.match(/[?&]id=([^&]+)/);
  if (m) return m[1];
  return '';
}

function getFileSafe_(id) {
  id = String(id || '').trim();
  var lastErr;
  for (var i = 0; i < 3; i++) {
    try {
      return DriveApp.getFileById(id);
    } catch (e) {
      lastErr = e;
      Utilities.sleep(500);
    }
  }
  throw lastErr;
}

/********** 1) 同一フォルダの生成物一覧を表示 **********/
function debug_listOutputs(fileIdOrUrl) {
  var id = parseDriveFileId(fileIdOrUrl);
  if (!id) {
    Logger.log('bad fileId/url');
    return;
  }

  var file = getFileSafe_(id);
  var parent = file.getParents().hasNext() ? file.getParents().next() : null;
  if (!parent) {
    Logger.log('parent not found');
    return;
  }

  var names = [];
  var it = parent.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    var nm = f.getName();
    // 代表ID由来の成果物 or OCR派生ファイルを拾う
    if (nm.indexOf(id) === 0 || /\.ocr\.(txt|json)$/i.test(nm)) names.push(nm);
  }
  Logger.log('files:\n' + names.join('\n'));
}

/** 入力プロンプト付き（固定ID廃止） */
function debug_listOutputs_prompt() {
  var s = Browser.inputBox('Driveの fileId または URL を入力');
  if (s === 'cancel') return;
  debug_listOutputs(s);
}

/********** 2) _model.json（抽出結果）を表示 **********/
function debug_readModelJson(fileIdOrUrl) {
  var id = parseDriveFileId(fileIdOrUrl);
  if (!id) {
    Logger.log('bad fileId/url');
    return;
  }

  var file = getFileSafe_(id);
  var parent = file.getParents().hasNext() ? file.getParents().next() : null;
  if (!parent) {
    Logger.log('parent not found');
    return;
  }

  var name = id + '_model.json';
  var it = parent.getFilesByName(name);
  if (it.hasNext()) {
    var json = it.next().getBlob().getDataAsString('utf-8');
    Logger.log(json);
  } else {
    Logger.log(name + ' not found');
  }
}

function debug_readModelJson_prompt() {
  var s = Browser.inputBox('Driveの fileId または URL を入力');
  if (s === 'cancel') return;
  debug_readModelJson(s);
}

/********** 3) エラーファイル（_extract_error.txt）を表示 **********/
function debug_readExtractError(fileIdOrUrl) {
  var id = parseDriveFileId(fileIdOrUrl);
  if (!id) {
    Logger.log('bad fileId/url');
    return;
  }

  var file = getFileSafe_(id);
  var parent = file.getParents().hasNext() ? file.getParents().next() : null;
  if (!parent) {
    Logger.log('parent not found');
    return;
  }

  var it = parent.getFilesByName(id + '_extract_error.txt');
  Logger.log(it.hasNext() ? it.next().getBlob().getDataAsString('utf-8') : 'error file not found');
}

function debug_readExtractError_prompt() {
  var s = Browser.inputBox('Driveの fileId または URL を入力');
  if (s === 'cancel') return;
  debug_readExtractError(s);
}

/********** 4) 既存のOCRテキストを使って /api/extract を“だけ”叩く **********
 * 本体の ensurePublicImageUrl_ / postExtract_ を利用します。
 * 既に <fileId>_model.json がある場合はスキップします。
 **********************************************************************/
function debug_postOnly(fileIdOrUrl, lineId) {
  var id = parseDriveFileId(fileIdOrUrl);
  if (!id) {
    Logger.log('bad fileId/url');
    return;
  }

  // 本体側にあるはずの関数が無い場合はエラーにする
  if (typeof ensurePublicImageUrl_ !== 'function' || typeof postExtract_ !== 'function') {
    throw new Error(
      'ensurePublicImageUrl_ / postExtract_ が本体に見つかりません。先に本体へ定義してください。'
    );
  }

  var file = getFileSafe_(id);
  var parent = file.getParents().hasNext() ? file.getParents().next() : null;
  if (!parent) {
    Logger.log('parent not found');
    return;
  }

  // 既に結果があればスキップ
  if (parent.getFilesByName(id + '_model.json').hasNext()) {
    Logger.log('already has _model.json, skip');
    return;
  }

  // OCRテキストがあれば使用（無ければ空でOK）
  var ocrText = '';
  var it = parent.getFilesByName(file.getName() + '.ocr.txt');
  if (it.hasNext()) {
    ocrText = it.next().getBlob().getDataAsString('utf-8');
  }

  // 画像の公開URLを確保
  var imageUrl = ensurePublicImageUrl_(file);

  // 抽出APIを叩く（lineIdは任意）
  var out = postExtract_(id, imageUrl, ocrText, lineId || 'LINE_DEBUG');

  // 成果物を保存
  if (out && out.ok && out.data) {
    parent.createFile(
      Utilities.newBlob(JSON.stringify(out.data, null, 2), 'application/json', id + '_model.json')
    );
    Logger.log('saved: ' + id + '_model.json');
  } else {
    parent.createFile(
      Utilities.newBlob(
        'status=' +
          (out && out.status) +
          '\n' +
          ((out && out.error) || '') +
          '\n\n' +
          ((out && out.raw) || ''),
        'text/plain',
        id + '_extract_error.txt'
      )
    );
    Logger.log('saved: ' + id + '_extract_error.txt');
  }
}

function debug_postOnly_prompt() {
  var f = Browser.inputBox('Driveの fileId または URL を入力');
  if (f === 'cancel') return;
  var l = Browser.inputBox('lineId（任意・空可）を入力');
  if (l === 'cancel') return;
  debug_postOnly(f, l || 'LINE_DEBUG');
}

/**
 * ステージング: /_staging/<YYYY-MM>/submission_<SID>/ 以下を
 * /<氏名__LINEID>/<日本語ドキュメント名>/<YYYY-MM>/ へ移送する
 */
function fix_moveSubmissionToUser(submissionId, lineId, displayName) {
  if (!submissionId || !lineId) throw new Error('submissionId と lineId は必須');

  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const sRoot = getOrCreateFolder(root, '_staging');

  // submission_<sid> を全期間から探す（同SIDは通常1箇所）
  const ymIter = sRoot.getFolders();
  let movedCount = 0;
  while (ymIter.hasNext()) {
    const ym = ymIter.next(); // 'yyyy-MM'
    const subIter = ym.getFoldersByName(`submission_${submissionId}`);
    if (!subIter.hasNext()) continue;

    const sub = subIter.next(); // submission_<sid>
    const typeFolders = sub.getFolders(); // 日本語書類名（給与明細 等）

    while (typeFolders.hasNext()) {
      const typeFolder = typeFolders.next();
      const typeName = typeFolder.getName(); // 例: '給与明細'

      // 提出月はステージング元の yyyy-MM を採用
      const dest = ensurePathJapanese_(lineId, displayName || '', typeName, ym.getName());

      // 中身を全部移動
      const it = typeFolder.getFiles();
      while (it.hasNext()) {
        const f = it.next();
        dest.addFile(f);
        typeFolder.removeFile(f);
        movedCount++;
      }
    }

    // 空になった submission_<sid> を削除
    const leftover = sub.getFiles().hasNext() || sub.getFolders().hasNext();
    if (!leftover) ym.removeFolder(sub);
  }

  Logger.log(`moved files: ${movedCount}`);
}

function fix_moveSubmissionToUser_example() {
  var submissionId = '47862386';
  var lineId = 'Uc13df94016ee50eb9dd5552bffbe6624';
  var displayName = ''; // 任意（空でも可。台帳補完があるなら空でOK）
  fix_moveSubmissionToUser(submissionId, lineId, displayName);
}

/********** トグル用ユーティリティ（GASエディタで1回実行） **********/
function __enableDebug() {
  PropertiesService.getScriptProperties().setProperty('ALLOW_DEBUG', '1');
}
function __disableDebug() {
  PropertiesService.getScriptProperties().deleteProperty('ALLOW_DEBUG');
}
function __checkDebug() {
  var v = PropertiesService.getScriptProperties().getProperty('ALLOW_DEBUG');
  Logger.log('ALLOW_DEBUG=' + (v ? v : '(not set)'));
}
// EOF
// @ts-nocheck
/**
 * debug_utils.gs
 * 本番コードから切り出したデバッグ用関数群（固定IDは使わない）
 * - 破壊的操作なし（閲覧/生成のみ）
 * - fileId か Drive URL を毎回手入力/引数で渡す
 */

/********** 共通ユーティリティ（安全） **********/
function parseDriveFileId(input) {
  var s = String(input || '').trim();
  if (/^[A-Za-z0-9_-]{20,}$/.test(s)) return s;
  var m = s.match(/\/file\/d\/([^/]+)/);
  if (m) return m[1];
  m = s.match(/[?&]id=([^&]+)/);
  if (m) return m[1];
  return '';
}

function getFileSafe_(id) {
  id = String(id || '').trim();
  var lastErr;
  for (var i = 0; i < 3; i++) {
    try {
      return DriveApp.getFileById(id);
    } catch (e) {
      lastErr = e;
      Utilities.sleep(500);
    }
  }
  throw lastErr;
}

/********** 1) 同一フォルダの生成物一覧を表示 **********/
function debug_listOutputs(fileIdOrUrl) {
  var id = parseDriveFileId(fileIdOrUrl);
  if (!id) {
    Logger.log('bad fileId/url');
    return;
  }

  var file = getFileSafe_(id);
  var parent = file.getParents().hasNext() ? file.getParents().next() : null;
  if (!parent) {
    Logger.log('parent not found');
    return;
  }

  var names = [];
  var it = parent.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    var nm = f.getName();
    // 代表ID由来の成果物 or OCR派生ファイルを拾う
    if (nm.indexOf(id) === 0 || /\.ocr\.(txt|json)$/i.test(nm)) names.push(nm);
  }
  Logger.log('files:\n' + names.join('\n'));
}

/** 入力プロンプト付き（固定ID廃止） */
function debug_listOutputs_prompt() {
  var s = Browser.inputBox('Driveの fileId または URL を入力');
  if (s === 'cancel') return;
  debug_listOutputs(s);
}

/********** 2) _model.json（抽出結果）を表示 **********/
function debug_readModelJson(fileIdOrUrl) {
  var id = parseDriveFileId(fileIdOrUrl);
  if (!id) {
    Logger.log('bad fileId/url');
    return;
  }

  var file = getFileSafe_(id);
  var parent = file.getParents().hasNext() ? file.getParents().next() : null;
  if (!parent) {
    Logger.log('parent not found');
    return;
  }

  var name = id + '_model.json';
  var it = parent.getFilesByName(name);
  if (it.hasNext()) {
    var json = it.next().getBlob().getDataAsString('utf-8');
    Logger.log(json);
  } else {
    Logger.log(name + ' not found');
  }
}

function debug_readModelJson_prompt() {
  var s = Browser.inputBox('Driveの fileId または URL を入力');
  if (s === 'cancel') return;
  debug_readModelJson(s);
}

/*環境の健全性チェック（プロパティ＆フォルダ＆スプレッドシート）*/
function diag_checkEnv() {
  try {
    const sid = props.getProperty('BAS_MASTER_SPREADSHEET_ID');
    const did = props.getProperty('DRIVE_ROOT_FOLDER_ID') || ROOT_FOLDER_ID;

    Logger.log('[ENV] BAS_MASTER_SPREADSHEET_ID=%s', sid);
    Logger.log('[ENV] DRIVE_ROOT_FOLDER_ID or ROOT_FOLDER_ID=%s', did);

    const ss = SpreadsheetApp.openById(sid);
    Logger.log('[ENV] Spreadsheet title=%s', ss.getName());

    const root = DriveApp.getFolderById(did);
    Logger.log('[ENV] Drive root name=%s', root.getName());

    ensureLabels();
    Logger.log('[ENV] Labels ensured');
    Logger.log('[OK] diag_checkEnv passed');
  } catch (e) {
    Logger.log('[NG] diag_checkEnv failed: %s', (e && e.stack) || e);
  }
}

/*ガードを通過した後の“保存処理だけ”を強制テスト
（Gmailなしで Drive 保存と JSON 保存の経路を検証）*/

function diag_fakeSave() {
  // 想定ユーザー（テスト用ダミー）
  const lineId = 'Uc13df94016ee50eb9dd5552bffbe6624';
  const displayName = 'テスト 太郎';
  const now = new Date();
  const meta = {
    subject: '[#FM-BAS] テスト送信',
    body:
      '==== META START ====\nline_id: ' +
      lineId +
      '\nform_name: 書類提出\nsecret: FM-BAS\n==== META END ====',
    line_id: lineId,
    form_name: '書類提出',
    submission_id: 'SUB' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd_HHmmss'),
    submitted_at: now.toISOString(),
    seq: '',
    referrer: '',
    secretOK: true,
    fields: { メールアドレス: 'dummy@address.com', 名前: displayName },
  };

  // ダミー添付（PNG風バイト列）
  const png = Utilities.newBlob('fake-image', 'image/png', '給与明細.png');

  // Gmailなしなので saveAttachmentsAndJson 内の attachments を模倣するため、
  // 直接カテゴリ保存APIを叩く（PAYカテゴリとして）
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const userFolder = getOrCreateUserFolder(root, lineId, displayName);
  const savedFile = saveAttachmentShallow_(userFolder, png, { hintedType: 'PAY', receivedAt: now });

  // JSON保存（本来は saveAttachmentsAndJson がやるが、ここでは個別にテスト）
  const jsonParent = userFolder;
  const jsonFile = jsonParent.createFile(
    Utilities.newBlob(
      JSON.stringify(
        { meta: { form_key: 'test', case_id: '', subject: meta.subject }, fields: meta.fields },
        null,
        2
      ),
      'application/json',
      'testform__' + meta.submission_id + '.json'
    )
  );

  Logger.log('[OK] fake saved: %s / %s', savedFile.getName(), jsonFile.getName());
}

/*ガード直前までの Gmail パスを可視化（何で弾かれてるか）*/
function diag_checkGuardOnToProcess() {
  const label = GmailApp.getUserLabelByName(LABEL_TO_PROCESS);
  if (!label) {
    Logger.log('ToProcess label not found');
    return;
  }
  const threads = label.getThreads(0, 10);
  if (!threads.length) {
    Logger.log('No threads under ToProcess');
    return;
  }

  threads.forEach((thread) => {
    const msg = thread.getMessages().pop();
    const body = msg.getPlainBody() || msg.getBody() || '';
    const metaKV = parseMeta_(body);
    const guard = checkNotificationGuard_(msg, metaKV);
    Logger.log(
      '[guard] subj="%s" hasTag=%s hasSecret=%s ok=%s meta.secret="%s"',
      msg.getSubject(),
      guard.hasTag,
      guard.hasSecret,
      guard.ok,
      metaKV && metaKV.secret
    );
  });
}

/* ログをもう少し詳しく見る（ワンショット実行）*/
function diag_processOnceVerbose() {
  try {
    ensureLabels();
    const labelToProcess = GmailApp.getUserLabelByName(LABEL_TO_PROCESS);
    const threads = labelToProcess ? labelToProcess.getThreads(0, 1) : [];
    if (!threads.length) {
      Logger.log('No thread under ToProcess');
      return;
    }

    const thread = threads[0];
    const msg = thread.getMessages().pop();
    const body = msg.getPlainBody() || msg.getBody() || '';
    const metaKV = parseMeta_(body);
    const guard = checkNotificationGuard_(msg, metaKV);
    Logger.log('[pre] guard ok=%s', guard.ok);
    if (!guard.ok) {
      Logger.log('Guard NG');
      return;
    }

    const parsed = parseMetaAndFields(msg);
    if (!parsed.form_name) parsed.form_name = 'unknown_form';

    const saved = saveAttachmentsAndJson(parsed, msg);
    Logger.log(
      '[saved] folderId=%s jsonId=%s files=%s',
      saved && saved.folderId,
      saved && saved.jsonId,
      ((saved && saved.savedFiles) || []).map((s) => s.name + ':' + s.id).join(', ')
    );

    ocr_processSaved_(saved, parsed); // CloudflareなしでもOK（失敗時は _extract_error.txt）
    thread.removeLabel(labelToProcess).addLabel(GmailApp.getUserLabelByName(LABEL_PROCESSED));
    Logger.log('[ok] processed thread=%s', thread.getId && thread.getId());
  } catch (e) {
    Logger.log('[err] %s', (e && e.stack) || e);
    try {
      GmailApp.getUserLabelByName(LABEL_ERROR) &&
        threads[0].addLabel(GmailApp.getUserLabelByName(LABEL_ERROR));
    } catch (_) {}
  }
}

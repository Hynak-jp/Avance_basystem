// _model.json が無ければ作ってから finalize まで自動実行
function finalize_fromImage_smart(fileId, lineId){
  fileId = resolveDriveFileIdForFinalize_(fileId);
  if (!fileId) {
    throw new Error(
      'fileId が未指定です。' +
      'Apps Script の実行ボタンでは引数を渡せないため、' +
      '1) finalize_fromImage_smart_withProps を実行して Script Properties の ' +
      'FINALIZE_FILE_ID / FINALIZE_FILE_URL を使う、' +
      'または 2) finalize_fromImage_smart(\"<Drive fileId or URL>\", \"<lineId>\") をコード上から呼び出してください。'
    );
  }
  var file;
  try {
    file = DriveApp.getFileById(fileId);
  } catch (e) {
    // 失敗原因を切り分けやすくする（削除済み・権限なし・ID違い）
    throw new Error(
      'DriveApp.getFileById 失敗: fileId=' + fileId +
      ' / ファイルが存在しない・権限がない・IDが誤りの可能性があります。' +
      '（元エラー: ' + ((e && e.message) || e) + '）'
    );
  }
  var parent = file.getParents().hasNext() ? file.getParents().next() : null;
  if (!parent) throw new Error('parent folder not found');

  var modelName = fileId + '_model.json';
  var itModel = parent.getFilesByName(modelName);

  if (!itModel.hasNext()){
    // 既存の ocr.txt を使って /api/extract へPOST（= debug_postOnly と同等）
    var ocrText = '';
    var itTxt = parent.getFilesByName(file.getName() + '.ocr.txt');
    if (itTxt.hasNext()) ocrText = itTxt.next().getBlob().getDataAsString('utf-8');

    var imageUrl = ensurePublicImageUrl_(file);
    var out = postExtract_(fileId, imageUrl, ocrText, lineId || 'LINE_TEST');

    if (out && out.ok && out.data){
      parent.createFile(
        Utilities.newBlob(JSON.stringify(out.data, null, 2), 'application/json', modelName)
      );
      // 小休止（Driveの整合待ち）
      Utilities.sleep(500);
    } else {
      parent.createFile(
        Utilities.newBlob(
          'status=' + (out && out.status) + '\n' + (out && out.error || '') + '\n\n' + (out && out.raw || ''),
          'text/plain',
          fileId + '_extract_error.txt'
        )
      );
      throw new Error('extract failed; see _extract_error.txt');
    }
  }

  // ここまで来たら _model.json あり → finalize
  finalize_fromModel_(fileId, lineId || 'LINE_TEST');
}

function normalizeDriveFileId_(value){
  if (typeof parseDriveFileId === 'function') {
    var parsed = String(parseDriveFileId(value) || '').trim();
    if (parsed) return parsed;
  }
  var s = String(value || '').trim();
  if (!s) return '';
  // 素の fileId
  if (/^[A-Za-z0-9_-]{20,}$/.test(s)) return s;
  // 代表的な共有URL
  var m = s.match(/\/file\/d\/([A-Za-z0-9_-]{20,})/);
  if (m && m[1]) return m[1];
  m = s.match(/\/folders\/([A-Za-z0-9_-]{20,})/);
  if (m && m[1]) return m[1];
  m = s.match(/[?&]resourcekey=([^&]+)/);
  if (m && m[1]) {
    var first = s.match(/\/d\/([A-Za-z0-9_-]{20,})/) || s.match(/[?&]id=([A-Za-z0-9_-]{20,})/);
    if (first && first[1]) return first[1];
  }
  var m = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m && m[1]) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m && m[1]) return m[1];
  return s;
}

function resolveDriveFileIdForFinalize_(value) {
  var resolved = normalizeDriveFileId_(value);
  if (resolved) return resolved;

  // 実行ボタン（引数なし）用のフォールバック:
  // Script Properties に FINALIZE_FILE_ID / FINALIZE_FILE_URL を設定しておく
  try {
    if (typeof PropertiesService !== 'undefined' && PropertiesService.getScriptProperties) {
      var props = PropertiesService.getScriptProperties();
      var candidates = [
        props.getProperty('FINALIZE_FILE_ID'),
        props.getProperty('FINALIZE_FILE_URL'),
        props.getProperty('FINALIZE_TARGET'),
        props.getProperty('LAST_FILE_ID'),
      ];
      for (var i = 0; i < candidates.length; i++) {
        var c = normalizeDriveFileId_(candidates[i]);
        if (c) return c;
      }
    }
  } catch (_) {}

  // 手動実行（エディタから関数実行）時の補助
  try {
    if (typeof Browser !== 'undefined' && Browser && typeof Browser.inputBox === 'function') {
      var input = Browser.inputBox('Driveの fileId または URL を入力');
      if (String(input || '').toLowerCase() === 'cancel') return '';
      return normalizeDriveFileId_(input);
    }
  } catch (_) {}
  return '';
}

// 引数なしで実行したいとき用。
// Script Properties:
//   FINALIZE_FILE_ID or FINALIZE_FILE_URL (必須)
//   FINALIZE_LINE_ID (任意)
function finalize_fromImage_smart_withProps() {
  var lineId = 'LINE_TEST';
  try {
    if (typeof PropertiesService !== 'undefined' && PropertiesService.getScriptProperties) {
      var props = PropertiesService.getScriptProperties();
      lineId = String(props.getProperty('FINALIZE_LINE_ID') || lineId);
    }
  } catch (_) {}
  finalize_fromImage_smart('', lineId);
}

// 実行前にターゲットを設定するためのヘルパー。
// 例:
//   finalize_setTarget_('1AbCdEfGhIjKlMnOpQrStUvWxYz012345', 'LINE_TEST');
//   finalize_setTarget_('https://drive.google.com/file/d/....../view', 'U1234567890');
function finalize_setTarget_(fileIdOrUrl, lineId) {
  var fileId = normalizeDriveFileId_(fileIdOrUrl);
  if (!fileId) throw new Error('fileId/url が不正です');
  var props = PropertiesService.getScriptProperties();
  props.setProperty('FINALIZE_FILE_ID', fileId);
  props.setProperty('FINALIZE_FILE_URL', String(fileIdOrUrl || ''));
  if (lineId) props.setProperty('FINALIZE_LINE_ID', String(lineId));
  Logger.log(
    '[finalize_setTarget_] FINALIZE_FILE_ID=%s FINALIZE_LINE_ID=%s',
    fileId,
    String(lineId || props.getProperty('FINALIZE_LINE_ID') || 'LINE_TEST')
  );
}

function finalize_showTarget_() {
  var props = PropertiesService.getScriptProperties();
  Logger.log(
    '[finalize_showTarget_] FILE_ID=%s FILE_URL=%s LINE_ID=%s',
    props.getProperty('FINALIZE_FILE_ID') || '',
    props.getProperty('FINALIZE_FILE_URL') || '',
    props.getProperty('FINALIZE_LINE_ID') || ''
  );
}

// 実行メニュー用ラッパー（例）
function finalize_fromImage_example(){
  var fileId = '1RIBN_QusUsH3P-dMLIXnXuI25rAvUnjA';
  finalize_fromImage_smart(fileId, 'LINE_TEST');
}

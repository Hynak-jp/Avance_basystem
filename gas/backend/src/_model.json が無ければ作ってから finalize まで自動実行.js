// _model.json が無ければ作ってから finalize まで自動実行
function finalize_fromImage_smart(fileId, lineId){
  fileId = normalizeDriveFileId_(fileId);
  if (!fileId) {
    throw new Error('fileId が不正です（空文字 / URL抽出失敗）');
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
  var s = String(value || '').trim();
  if (!s) return '';
  var m = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m && m[1]) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m && m[1]) return m[1];
  return s;
}

// 実行メニュー用ラッパー（例）
function finalize_fromImage_example(){
  var fileId = '1RIBN_QusUsH3P-dMLIXnXuI25rAvUnjA';
  finalize_fromImage_smart(fileId, 'LINE_TEST');
}

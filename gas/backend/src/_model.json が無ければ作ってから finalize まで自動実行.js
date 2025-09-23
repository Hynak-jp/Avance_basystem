// _model.json が無ければ作ってから finalize まで自動実行
function finalize_fromImage_smart(fileId, lineId){
  fileId = String(fileId).trim();
  var file   = DriveApp.getFileById(fileId);
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

// 実行メニュー用ラッパー（例）
function finalize_fromImage_example(){
  var fileId = '1RIBN_QusUsH3P-dMLIXnXuI25rAvUnjA';
  finalize_fromImage_smart(fileId, 'LINE_TEST');
}

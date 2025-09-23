function quick_ocr_smoke_rest() {
  // ROOT直下の最初の jpg/png を拾う
  const folder = DriveApp.getFolderById(ROOT_FOLDER_ID);
  let target = null;
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (/\.(jpg|jpeg|png)$/i.test(f.getName())) { target = f; break; }
  }
  if (!target) throw new Error('ROOT直下に小さめの JPG/PNG を置いてください。');

  const body = {
    requests: [{
      image: { content: Utilities.base64Encode(target.getBlob().getBytes()) },
      features: [{ type: 'DOCUMENT_TEXT_DETECTION' }],
      imageContext: { languageHints: ['ja','en'] }
    }]
  };

  const res = UrlFetchApp.fetch(
    'https://vision.googleapis.com/v1/images:annotate',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    }
  );
  const obj = JSON.parse(res.getContentText());
  const text = obj.responses?.[0]?.fullTextAnnotation?.text || '(no text)';
  Logger.log(text.substring(0, 800));

  // ついでに同フォルダへ保存
  const parent = target.getParents().hasNext() ? target.getParents().next() : null;
  if (parent) parent.createFile(Utilities.newBlob(text, 'text/plain', target.getName() + '.ocr.txt'));
}

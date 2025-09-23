/** ========= 仕上げ用パッチ =========
 *  _model.json を読み込んで
 *   1) Parsed/NeedsReview を判定
 *   2) 台帳シートに1行追記
 *   3) Googleドキュメントのドラフト生成
 *   4) DOCX にエクスポート（同じフォルダへ）
 *  ※ 既存の openOrCreateLedger() はあなたのコードを利用します
 *  ================================== */

// 設定
const SHEET_PAYSLIPS = 'payslips'; // 追記先シート名
const CONF_THRESHOLD = 0.50;      // 以前: 0.60
const MAX_DIFF_YEN   = 3000;  

// ===== ユーティリティ =====
function _link(url, label){ return '=HYPERLINK("' + url + '","' + (label||'link') + '")'; }
function _viewUrl(fileId){ return 'https://drive.google.com/file/d/' + fileId + '/view'; }

// _model.json を読む
function readModelJson_(fileId){
  const file = DriveApp.getFileById(String(fileId).trim());
  const parent = file.getParents().hasNext() ? file.getParents().next() : null;
  if (!parent) throw new Error('parent folder not found');
  const name = fileId + '_model.json';
  const it = parent.getFilesByName(name);
  if (!it.hasNext()) throw new Error(name + ' not found');
  const f = it.next();
  const json = JSON.parse(f.getBlob().getDataAsString('utf-8'));
  return { json, modelFileId: f.getId(), parent, imageFile: file };
}

function ensureSubfolder_(parent, name){
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// 仕分けロジック
function classifyPaySlip_(ps){
  const reasons = [];
  const j = ps || {};
  const num = v => (typeof v === 'number' ? v : (typeof v === 'string' ? parseInt(String(v).replace(/[^\d-]/g,''),10) : NaN));
  const okDoc = j.doc_type === 'pay_slip';
  if (!okDoc) reasons.push('doc_type != pay_slip');

  const okPeriod = typeof j.pay_period === 'string' && /^\d{4}-\d{2}$/.test(j.pay_period);
  if (!okPeriod) reasons.push('pay_period invalid');

  const net = num(j.net_amount);
  if (!isFinite(net) || net <= 0) reasons.push('net_amount missing/invalid');

  // 整合チェック（緩め）
  const gross = isFinite(num(j.gross_amount)) ? num(j.gross_amount) : null;
  const ded   = isFinite(num(j.deductions_total)) ? num(j.deductions_total) : null;
  const itemsSum = Array.isArray(j.items) ? j.items.reduce((a,it)=>a + (isFinite(num(it.amount))?num(it.amount):0), 0) : null;

  if (gross != null && ded != null && isFinite(net)) {
    const diff = Math.abs(gross - ded - net);
    if (diff > 2000) reasons.push('gross - deductions - net mismatch: ' + diff);
  }
  if (gross != null && itemsSum != null) {
    const diff2 = Math.abs(gross - itemsSum);
    if (diff2 > 2000) reasons.push('items sum mismatch: ' + diff2);
  }

  const conf = typeof j.confidence === 'number' ? j.confidence : 0.0;
  if (conf < CONF_THRESHOLD) reasons.push('low confidence: ' + conf);

  const status = reasons.length ? 'NeedsReview' : 'Parsed';
  return { status, reasons, confidence: conf, gross, ded, net, itemsSum };
}

// ドキュメント作成 → DOCX書き出し
function createDocAndDocx_(parentFolder, imageFile, ps){
  const title = 'BAS_PAYSLIP_' + (imageFile.getName().replace(/\.[^.]+$/,''));
  const doc = DocumentApp.create(title);
  const docId = doc.getId();

  const b = doc.getBody();
  b.appendParagraph('破産申立 添付書類（給与明細）ドラフト').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  const p = (k,v)=> b.appendParagraph(k + '：' + (v==null?'':v));
  p('社員氏名', ps.employee_name || '');
  p('勤務先',   ps.employer || '');
  p('支給対象月', ps.pay_period || '');
  p('支給日',     ps.pay_date || '');
  p('総支給額',   typeof ps.gross_amount==='number'? yen_(ps.gross_amount) : '');
  p('控除合計',   typeof ps.deductions_total==='number'? yen_(ps.deductions_total) : '');
  p('差引支給額', typeof ps.net_amount==='number'? yen_(ps.net_amount) : '');
  b.appendParagraph('内訳').setHeading(DocumentApp.ParagraphHeading.HEADING2);

  if (Array.isArray(ps.items) && ps.items.length){
    const table = b.appendTable([['項目','金額']]);
    ps.items.forEach(it=>{
      table.appendTableRow([ String(it.label||''), String(it.amount||'') ]);
    });
  } else {
    b.appendParagraph('（内訳なし）');
  }

  // 画像へのリンクを追記
  b.appendParagraph('画像: ' + _viewUrl(imageFile.getId()));

  doc.saveAndClose();

  // DOCXへエクスポート（Advanced Drive 不要：RESTを直叩き）
  let docxId = '';
  try {
    const url = 'https://www.googleapis.com/drive/v3/files/'+docId+'/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() === 200) {
      const blob = resp.getBlob().setName(title + '.docx');
      const docx = parentFolder.createFile(blob);
      docxId = docx.getId();
    } else {
      Logger.log('DOCX export failed: ' + resp.getContentText());
    }
  } catch (e) {
    Logger.log('DOCX export error: ' + e);
  }

  // 作成したGoogleドキュメントを親フォルダに移動
  try {
    const gdocFile = DriveApp.getFileById(docId);
    parentFolder.addFile(gdocFile);
    DriveApp.getRootFolder().removeFile(gdocFile); // ルートから外す（権限OKなら）
  } catch (_) {}

  return { docId, docxId };
}

// 台帳（payslips）に1行追記
function appendToPayslipsSheet_(fileId, ps, clf, links, lineId){
  const ss = openOrCreateLedger(); // ← 既存の関数を使用
  const sh = getOrCreateSheet(ss, SHEET_PAYSLIPS, [
    'ts','line_id','source_id','pay_period','pay_date',
    'employer','employee_name','gross_amount','deductions_total','net_amount',
    'confidence','status','reasons','image','model','doc','docx'
  ]);
  sh.appendRow([
    Utilities.formatDate(new Date(), 'Asia/Tokyo','yyyy-MM-dd HH:mm:ss'),
    lineId || ps.line_id || '',
    fileId,
    ps.pay_period || '',
    ps.pay_date || '',
    ps.employer || '',
    ps.employee_name || '',
    typeof ps.gross_amount==='number'? ps.gross_amount : '',
    typeof ps.deductions_total==='number'? ps.deductions_total : '',
    typeof ps.net_amount==='number'? ps.net_amount : '',
    typeof ps.confidence==='number'? ps.confidence : '',
    clf.status,
    (clf.reasons||[]).join('; '),
    _link(links.image),
    _link(links.model),
    links.doc ? _link(links.doc,'doc') : '',
    links.docx ? _link(links.docx,'docx') : ''
  ]);
}

// --- 既存の _toHalf / _num / _pickDate / ocr_extractPayslip_ はそのまま利用可 ---

// .ocr.txt を読む（無くても空文字を返す）
function readOcrTextForFile_(imageFile){
  var parent = imageFile.getParents().hasNext() ? imageFile.getParents().next() : null;
  if (!parent) return '';
  var it = parent.getFilesByName(imageFile.getName() + '.ocr.txt');
  return it.hasNext() ? it.next().getBlob().getDataAsString('utf-8') : '';
}

// OCRテキストから “それっぽい項目” を配列化（見つかったものだけ）
// OCRテキストから「ラベル + 金額」を広く拾う（表でも行でもOK）
function buildItemsFromOcrText_(t){
  if (!t) return [];
  var lines = _toHalf(t).split(/\r?\n/).map(function(s){
    return s.replace(/[￥¥,，]/g,'').trim();
  });

  // 合計系は除外
  var ban = /(総支給|支給合計|総控除|控除合計|差引|手取|合計|小計|計)/;

  // ラベル正規化辞書（含む→統一名）
  var dict = [
    ['基本給','基本給'],
    ['時間外','時間外勤務手当'], ['残業','時間外勤務手当'],
    ['休日','休日勤務手当'],
    ['深夜','深夜勤務手当'],
    ['家族','家族手当'],
    ['住宅','住宅手当'],
    ['通勤','通勤手当'], ['交通費','通勤手当'],
    ['健康保険','健康保険'], ['健保','健康保険'],
    ['厚生年金','厚生年金'], ['年金','厚生年金'],
    ['雇用保険','雇用保険'],
    ['所得税','所得税'], ['源泉','所得税'],
    ['住民税','住民税']
  ];

  var out = [], seen = {};
  lines.forEach(function(s){
    // 例: 「基本給： 303100」「通勤手当 22500 円」「健康保険 2686」
    var m = s.match(/^(.{2,24}?)[\s:：]*(-?\d[\d\.]*)\s*(円|$)/);
    if (!m) return;
    var raw = m[1].replace(/[\s:：]+$/,'');
    if (ban.test(raw)) return;

    var n = parseInt(m[2].replace(/[^\d\-]/g,''), 10);
    if (!isFinite(n)) return;

    // ラベル正規化
    var label = raw;
    for (var i=0;i<dict.length;i++){
      if (raw.indexOf(dict[i][0]) >= 0){ label = dict[i][1]; break; }
    }

    var key = label + '@' + n;
    if (seen[key]) return;
    seen[key] = 1;

    out.push({ label: label, amount: n });
  });

  return out.slice(0, 20); // 念のため上限
}


// 数値表示を見栄え良く
function yen_(v){ if (typeof v !== 'number') return ''; return '¥' + String(v).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }

// メイン：_model.json → 仕上げ
function finalize_fromModel_(fileId, lineId){
  const { json, modelFileId, parent, imageFile } = readModelJson_(fileId);
  // OCRテキストを読んで欠損を補完
  const ocrText = readOcrTextForFile_(imageFile);
  if (ocrText){
    const p = ocr_extractPayslip_(ocrText); // 既存のPoC抽出器
    if (!json.employee_name && p.employee) json.employee_name = p.employee;
    if (!json.employer && p.company)       json.employer      = p.company;
    if (!json.pay_date && p.payday)        json.pay_date      = p.payday;

    // items が空なら OCR から組み立て
    if (!Array.isArray(json.items) || !json.items.length){
      const built = buildItemsFromOcrText_(ocrText);
      if (built.length) json.items = built;
    }
  }

  // 合計の保険（モデルが入れ忘れた場合）
  if ((json.gross_amount == null) && Array.isArray(json.items) && json.items.length){
    json.gross_amount = json.items.reduce((a,it)=>a + (isFinite(it.amount)? it.amount:0), 0);
  }

  const clf = classifyPaySlip_(json);
  const { docId, docxId } = createDocAndDocx_(parent, imageFile, json);
  const links = {
    image: _viewUrl(imageFile.getId()),
    model: _viewUrl(modelFileId),
    doc: docId ? 'https://docs.google.com/document/d/'+docId+'/edit' : '',
    docx: docxId ? _viewUrl(docxId) : ''
  };
  const targetFolder = ensureSubfolder_(parent, clf.status); // Parsed or NeedsReview
  try {
    // 対象ファイルをサブフォルダにもぶら下げる（元の場所には残す）
    // model.json
    const modelName = fileId + '_model.json';
    const itModel = parent.getFilesByName(modelName);
    if (itModel.hasNext()) targetFolder.addFile(itModel.next());

    // Googleドキュメント
    if (docId) targetFolder.addFile(DriveApp.getFileById(docId));
    // DOCX
    if (docxId) targetFolder.addFile(DriveApp.getFileById(docxId));
  } catch(e) {
    Logger.log('attach to subfolder failed: ' + e);
  }
  appendToPayslipsSheet_(fileId, json, clf, links, lineId || ''); // 追記
  Logger.log('finalized: status=%s, doc=%s, docx=%s', clf.status, links.doc, links.docx);
}

/* ---- 実行メニュー用ラッパー（IDを差し替えて使ってください） ---- */
function finalize_fromModel_example(){
  var fileId = '1RIBN_QusUsH3P-dMLIXnXuI25rAvUnjA';
  finalize_fromModel_(fileId, 'LINE_TEST');
}

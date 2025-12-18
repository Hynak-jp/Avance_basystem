/**
 * S2005（債権者一覧表）テンプレ差し込み → cases/<case>/drafts/
 * - テンプレはマージ撤廃版（スクショ準拠）
 * - F列「使途」のチェックは、テンプレ内の「□住宅ローン」「□保証」等の □ を ■ に置換して表現
 * - D25 = 住宅ローン合計、F25 = 保証債務合計（どちらも F 列のチェック状態を条件に D 列を合算）
 * - 注意: テンプレート列構成を変更する際は detectHeaderRow_S2005_ と列マッピングを必ず更新すること
 */

const S2005_PROP = PropertiesService.getScriptProperties();
const S2005_TPL_GSHEET_ID = S2005_PROP.getProperty('S2005_TEMPLATE_GSHEET_ID') || '';

/** ============== Public Entrypoints ============== **/

/** 例: run_GenerateS2005DraftBySubmissionId('0001', '49192017') */
function run_GenerateS2005DraftBySubmissionId(caseId, submissionId) {
  if (!S2005_TPL_GSHEET_ID) throw new Error('S2005_TEMPLATE_GSHEET_ID not set');
  if (!caseId) throw new Error('caseId is required');
  if (!submissionId) throw new Error('submissionId is required');

  const caseInfo = resolveCaseByCaseId_(String(caseId).trim());
  if (!caseInfo) throw new Error('Unknown case_id: ' + caseId);
  caseInfo.folderId = ensureCaseFolderId_(caseInfo);

  const jsonName = `s2005_creditors__${String(submissionId)}.json`;
  const folder = DriveApp.getFolderById(caseInfo.folderId);
  const it = folder.getFilesByName(jsonName);
  if (!it.hasNext()) throw new Error('JSON not found: ' + jsonName);

  const parsed = JSON.parse(it.next().getBlob().getDataAsString('UTF-8'));
  return generateS2005SheetDraft_(caseInfo, parsed);
}

/** ケース内の最新 S2005 JSON から生成 */
function run_GenerateS2005DraftLatest(caseId) {
  if (!S2005_TPL_GSHEET_ID) throw new Error('S2005_TEMPLATE_GSHEET_ID not set');
  if (!caseId) throw new Error('caseId is required');

  const caseInfo = resolveCaseByCaseId_(String(caseId).trim());
  if (!caseInfo) throw new Error('Unknown case_id: ' + caseId);
  caseInfo.folderId = ensureCaseFolderId_(caseInfo);

  const folder = DriveApp.getFolderById(caseInfo.folderId);
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const n = f.getName && f.getName();
    if (/^s2005_creditors__\d+\.json$/i.test(n)) {
      files.push({ f, t: f.getLastUpdated().getTime() });
    }
  }
  if (!files.length) throw new Error('No S2005 JSON in case folder');
  files.sort((a, b) => b.t - a.t);
  const parsed = JSON.parse(files[0].f.getBlob().getDataAsString('UTF-8'));
  return generateS2005SheetDraft_(caseInfo, parsed);
}

/** ============== Core ============== **/

function generateS2005SheetDraft_(caseInfo, parsed) {
  const M = pickS2005Model_(parsed); // {creditors:[], totals:{amount_sum,...}}

  // コピー作成 → drafts へ
  const parent = DriveApp.getFolderById(caseInfo.folderId);
  const drafts = getOrCreateSubfolder_(parent, 'drafts');
  const sid =
    (parsed && parsed.meta && parsed.meta.submission_id) ||
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMMddHHmmss');
  const draftName = `S2005_${caseInfo.caseId}_draft_${sid}`;
  const ssId = DriveApp.getFileById(S2005_TPL_GSHEET_ID).makeCopy(draftName, drafts).getId();
  const ss = SpreadsheetApp.openById(ssId);
  const sh = ss.getSheetByName('債権者一覧表') || ss.getSheetByName('S2005') || ss.getSheets()[0];

  // 見出し行検出（先頭10行から）
  const hdr = detectHeaderRow_S2005_(sh);
  if (!hdr || !hdr.row) throw new Error('Header row not found in template');
  const startRow = hdr.row + 1;

  // 必要行数を確保
  ensureRows_(sh, startRow, M.creditors.length || 1);
  const rowCount = Math.max(1, M.creditors.length);

  // 金額集計（全体 / 住宅ローン / 保証）
  let sumAll = 0,
    sumHouseLoan = 0,
    sumGuarantee = 0;

  // 1行分の「使途」テンプレ（□…が並んだ元テキスト）を基準にする
  // 行ごとに取得（テンプレ側に既に記載されていることを前提）
  function buildPurposeCellText_(row, purposeWord) {
    const c = hdr.map.purpose;
    if (!c) return purposeWord || '';
    let base = sh.getRange(row, c).getDisplayValue() || '';
    // 正規化（全角→半角、空白整理）
    const norm = (s) => {
      const raw = String(s || '');
      const normalized = raw.normalize ? raw.normalize('NFKC') : raw;
      return normalized.replace(/\s+/g, '');
    };
    const p = norm(purposeWord);

    // チェック対象語彙（必要なら増やせる）
    const targets = [
      { key: '住宅ローン', re: /□\s*住宅ローン/g },
      { key: '保証', re: /□\s*(?:保証|連帯保証|保証債務)/g },
      { key: '購入', re: /□\s*購入/g },
      { key: '生活費', re: /□\s*生活費/g },
      { key: '事業資金', re: /□\s*事業資金/g },
      { key: '教育資金', re: /□\s*教育資金/g },
      { key: '医療費', re: /□\s*医療費/g },
      { key: 'その他', re: /□\s*その他/g },
    ];

    targets.forEach((t) => {
      if (p.indexOf(t.key) !== -1) {
        base = base.replace(t.re, '■ ' + t.key); // □→■
      }
    });
    return base;
  }

  // 差し込み
  for (let i = 0; i < M.creditors.length; i++) {
    const r = startRow + i;
    const rec = M.creditors[i] || {};

    // A:番号
    putCell_(sh, r, hdr.map.no, i + 1);

    // B:債権者名
    putCell_(sh, r, hdr.map.creditor_name, rec.creditor_name || '');

    // C:住所（JSONに無ければテンプレの「〒」を維持）
    if (rec.address && String(rec.address).trim()) {
      putCell_(sh, r, hdr.map.address, String(rec.address).trim());
    }

    // D:現在の債務額(円)
    const amt = asIntOrBlank_(rec.current_debt_jpy);
    putCell_(sh, r, hdr.map.current_debt_jpy, amt);
    if (hdr.map.current_debt_jpy) {
      try {
        sh.getRange(r, hdr.map.current_debt_jpy).setNumberFormat('#,##0');
      } catch (_) {}
    }
    sumAll += Number(amt || 0);

    // E:借入・購入等の日（和暦略号で整形）
    if (hdr.map.date_range) {
      const eraRange = formatEraRange_(rec.date_start, rec.date_end);
      if (eraRange) {
        putCell_(sh, r, hdr.map.date_range, eraRange);
      }
    }

    // F:使途（□→■ 置換）
    const purposeText = buildPurposeCellText_(r, rec.purpose || '');
    putCell_(sh, r, hdr.map.purpose, purposeText);

    // H:調査票 / I:意見票（□に ■ を書く）
    if (hdr.map.investigation) {
      setCheckCell_(sh, r, hdr.map.investigation, !!rec.investigation_present);
    }
    if (hdr.map.opinion) {
      setCheckCell_(sh, r, hdr.map.opinion, !!rec.opinion_present);
    }

    // 住宅ローン・保証の条件合計（F列の書き込み結果で判定）
    const marked = purposeText.replace(/\s+/g, '');
    if (/■住宅ローン/.test(marked)) sumHouseLoan += Number(amt || 0);
    if (/■(?:保証|連帯保証|保証債務)/.test(marked)) sumGuarantee += Number(amt || 0);
  }

  // 折返し（名前・使途・備考・日付）
  const wrapCols = [
    hdr.map.creditor_name,
    hdr.map.purpose,
    hdr.map.memo,
    hdr.map.date_range,
  ].filter(Boolean);
  wrapCols.forEach((c) => {
    try {
      const rng = sh.getRange(startRow, c, rowCount, 1);
      rng.setWrap(true);
      if (hdr.map.date_range && c === hdr.map.date_range) {
        rng.setVerticalAlignment('top');
      }
    } catch (_) {}
  });

  // 下部集計
  //  - 債務総額（D24）はテンプレ側の数式があれば任せる。無ければ補助的に書く
  //  - うち住宅ローン（D25）、保証債務（F25）は本ロジックで明示セット
  try {
    const d24 = sh.getRange(24, 4); // D24
    if (!d24.getFormula()) d24.setValue(sumAll);
  } catch (_) {}

  // D25 (うち住宅ローン)
  try {
    sh.getRange(25, 4).setValue(asIntOrBlank_(sumHouseLoan)).setNumberFormat('#,##0');
  } catch (_) {}

  // F25 (保証債務)
  try {
    sh.getRange(25, 6).setValue(asIntOrBlank_(sumGuarantee)).setNumberFormat('#,##0');
  } catch (_) {}

  // ログ
  try {
    Logger.log(
      '[S2005] draft created case=%s sid=%s sums=%j fileId=%s url=%s',
      caseInfo.caseId || '',
      sid,
      { all: sumAll, house: sumHouseLoan, guarantee: sumGuarantee },
      ssId,
      ss.getUrl()
    );
  } catch (_) {}

  return {
    fileId: ssId,
    url: ss.getUrl(),
    name: draftName,
    caseId: caseInfo.caseId,
    caseKey: caseInfo.caseKey,
    sums: { all: sumAll, house_loan: sumHouseLoan, guarantee: sumGuarantee },
  };
}

/** ============== Header Detection & Write Helpers ============== **/

function detectHeaderRow_S2005_(sh) {
  const maxCols = Math.min(30, sh.getMaxColumns());
  const maxRows = Math.min(10, sh.getMaxRows());
  const norm = (s) => String(s || '').replace(/[ 　\t]/g, '');

  const patterns = {
    no: /^(No\.?|番号)$/i,
    creditor_name: /債権者名/,
    address: /^住所$/,
    current_debt_jpy: /(現在の債務額\(円\)|現在の債務額|債務額)/,
    date_range: /(借入・購入等の日|借入開始|開始日)/,
    purpose: /^使途$/,
    memo: /^備考$/,
    investigation: /調査票/,
    opinion: /意見票|意見/,
  };

  for (let r = 1; r <= maxRows; r++) {
    const row = sh.getRange(r, 1, 1, maxCols).getDisplayValues()[0];
    const map = {};
    for (let c = 1; c <= row.length; c++) {
      const v = norm(row[c - 1]);
      if (!v) continue;
      Object.keys(patterns).forEach((k) => {
        if (map[k]) return;
        if (patterns[k].test(v)) map[k] = c;
      });
    }
    // 必須（最低限）：債権者名/金額
    if (map.creditor_name && map.current_debt_jpy) {
      if (!map.no) map.no = findEmptyIndexColumn_(row) || 1;
      return { row: r, map };
    }
  }
  return null;
}

function ensureRows_(sh, startRow, need) {
  const remain = sh.getMaxRows() - (startRow - 1);
  if (remain < need) sh.insertRowsAfter(sh.getMaxRows(), need - remain);
}
function putCell_(sh, r, c, v) {
  if (!c) return;
  sh.getRange(r, c).setValue(v == null ? '' : v);
}
function findEmptyIndexColumn_(rowValues) {
  for (let i = 0; i < rowValues.length; i++) {
    const v = String(rowValues[i] || '').trim();
    if (!v) return i + 1;
  }
  return 1;
}
function asIntOrBlank_(v) {
  const n = parseInt(String(v ?? ''), 10);
  return isFinite(n) ? n : '';
}

/** □/■ あるいは Google チェックボックスへ安全に書き込む */
function setCheckCell_(sh, row, col, checked) {
  if (!col) return;
  var rng = sh.getRange(row, col);
  try {
    var dv = rng.getDataValidation && rng.getDataValidation();
    if (
      dv &&
      dv.getCriteriaType &&
      typeof SpreadsheetApp !== 'undefined' &&
      SpreadsheetApp.DataValidationCriteria &&
      dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX
    ) {
      rng.setValue(!!checked);
      return;
    }
  } catch (_) {}
  rng.setValue(checked ? '■' : '□');
}

/** 西暦 'YYYY-MM-DD' → 和暦略号 'H30.02.20' / 'R7.10.21' / （保険）'S64.01.07' */
function toEraShort_(ymd) {
  var s = String(ymd || '').trim();
  if (!s) return '';
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return '';
  var Y = +m[1],
    Mo = +m[2],
    D = +m[3];
  var stamp = new Date(Y, Mo - 1, D).getTime();

  // 境界（含む/含まないに注意）
  var R_START = new Date(2019, 4, 1).getTime(); // 2019-05-01
  var H_START = new Date(1989, 0, 8).getTime(); // 1989-01-08
  var S_START = new Date(1926, 11, 25).getTime(); // 1926-12-25

  function pad2(n) {
    return (n < 10 ? '0' : '') + n;
  }

  if (stamp >= R_START) {
    var r = Y - 2018; // 2019→1
    return 'R' + r + '.' + pad2(Mo) + '.' + pad2(D);
  } else if (stamp >= H_START) {
    var h = Y - 1988; // 1989→1
    return 'H' + h + '.' + pad2(Mo) + '.' + pad2(D);
  } else if (stamp >= S_START) {
    var sh = Y - 1925; // 1926→1
    return 'S' + sh + '.' + pad2(Mo) + '.' + pad2(D);
  }
  // それ以前は保険として西暦ドット表記
  return Y + '.' + pad2(Mo) + '.' + pad2(D);
}

/** 'start','end' を和暦で整形し、両方あれば "start\n～\nend"、片方だけならその1行 */
function formatEraRange_(startYmd, endYmd) {
  var s = toEraShort_(startYmd);
  var e = toEraShort_(endYmd);
  if (s && e) return s + '\n～\n' + e;
  return s || e || '';
}

/** ============== Model Extractor（保険） ============== **/

function pickS2005Model_(parsed) {
  if (parsed && parsed.model && Array.isArray(parsed.model.creditors)) {
    const creditors = parsed.model.creditors || [];
    const totals = parsed.model.totals || { amount_sum: sumCreditorAmounts_(creditors) };
    return { creditors, totals };
  }
  // マッパーが読み込まれていれば使う
  try {
    if (typeof mapS2005FieldsToModel_ === 'function') {
      const m = mapS2005FieldsToModel_(parsed.fieldsRaw || []);
      if (m && Array.isArray(m.creditors)) {
        return {
          creditors: m.creditors,
          totals: m.totals || { amount_sum: sumCreditorAmounts_(m.creditors) },
        };
      }
    }
  } catch (e) {
    try {
      Logger.log('[S2005] mapper error: %s', (e && e.stack) || e);
    } catch (_) {}
  }
  // 最後の保険（最低限）
  const rows = [];
  const src = parsed.fieldsRaw || [];
  const money = (s) => {
    const n = parseInt(String(s || '').replace(/[^\d]/g, ''), 10);
    return isFinite(n) ? n : 0;
  };
  function read(re) {
    const row = src.find((r) => re.test(String(r.label || '')));
    return row ? String(row.value || '') : '';
  }
  for (let i = 1; i <= 12; i++) {
    const name = read(new RegExp(`^【\\s*債権者${i}:債権者名`));
    const amt = read(new RegExp(`^【\\s*債権者${i}:現在の債務額`));
    const s = read(new RegExp(`^【\\s*債権者${i}:借入・購入等の日（開始日）`));
    const e = read(new RegExp(`^【\\s*債権者${i}:借入・購入等の日（終了日）`));
    const pu = read(new RegExp(`^【\\s*債権者${i}:使途`));
    const memo = read(new RegExp(`^【\\s*債権者${i}:備考`));
    if (name || amt || pu || memo || s || e) {
      rows.push({
        index: i,
        creditor_name: name,
        current_debt_jpy: money(amt),
        date_start: s,
        date_end: e,
        purpose: pu,
        memo: memo,
      });
    }
  }
  return { creditors: rows, totals: { amount_sum: sumCreditorAmounts_(rows) } };
}

function sumCreditorAmounts_(rows) {
  let t = 0;
  (rows || []).forEach((rec) => (t += rec.current_debt_jpy || 0));
  return t;
}

/** ============== Auto Draft Registration ============== **/

try {
  if (typeof registerAutoDraft === 'function') {
    registerAutoDraft('s2005_creditors', function (caseId, submissionId) {
      return run_GenerateS2005DraftBySubmissionId(caseId, submissionId);
    });
  }
} catch (_) {}

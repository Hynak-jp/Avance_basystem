/**
 * S2006（公租公課）テンプレ差し込み → drafts 保存
 *
 * 前提:
 * - 注意: シート構成や行番号を変更する場合は ROW マッピングと書き込み位置を合わせて更新すること
 * - Script Properties:
 *     S2006_TEMPLATE_GSHEET_ID : テンプレ（Googleスプレッドシート）ファイルID
 * - 既存ユーティリティ:
 *     resolveCaseByCaseId_(caseId), ensureCaseFolderId_(caseInfo), getOrCreateSubfolder_(parent, name)
 * - JSON 命名: s2006_creditors_public__<submission_id>.json （ケース直下）
 *
 * 差し込み方式:
 * - シートA列の見出し（例:「所得税」「自動車税・軽自動車税」）から行番号を動的に特定
 * - B列: 滞納額,  C列: 年度(YYYY-MM),  D列: 納付先（vehicleは登録番号も括弧書き）
 */

const S2006_PROP = PropertiesService.getScriptProperties();
const S2006_TPL_GSHEET_ID = S2006_PROP.getProperty('S2006_TEMPLATE_GSHEET_ID') || '';

function wrap_(rng) {
  try {
    rng.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  } catch (_) {}
}

function composeVehicleDCell_(rec) {
  const payer = String(rec.payer || '').trim();
  const plate = String(rec.plate_pretty || rec.plate || rec.plate_raw || '').trim();
  if (payer && plate) return payer + '\n' + '登録番号：' + plate;
  if (payer) return payer;
  if (plate) return '登録番号：' + plate;
  return '';
}

/** ============== Public Entrypoints ============== **/

/** 例: run_GenerateS2006DraftBySubmissionId('0001', '49100925') */
function run_GenerateS2006DraftBySubmissionId(caseId, submissionId) {
  if (!S2006_TPL_GSHEET_ID) throw new Error('S2006_TEMPLATE_GSHEET_ID not set');
  if (!caseId) throw new Error('caseId is required');
  if (!submissionId) throw new Error('submissionId is required');

  const caseInfo = resolveCaseByCaseId_(String(caseId).trim());
  if (!caseInfo) throw new Error('Unknown case_id: ' + caseId);
  caseInfo.folderId = ensureCaseFolderId_(caseInfo);

  const jsonName = `s2006_creditors_public__${String(submissionId)}.json`;
  const folder = DriveApp.getFolderById(caseInfo.folderId);
  const it = folder.getFilesByName(jsonName);
  if (!it.hasNext()) throw new Error('JSON not found: ' + jsonName);

  const parsed = JSON.parse(it.next().getBlob().getDataAsString('UTF-8'));
  return generateS2006SheetDraft_(caseInfo, parsed);
}

/** ケース内で最新の S2006 JSON から生成（お好みで） */
function run_GenerateS2006DraftLatest(caseId) {
  if (!S2006_TPL_GSHEET_ID) throw new Error('S2006_TEMPLATE_GSHEET_ID not set');
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
    if (/^s2006_creditors_public__\d+\.json$/i.test(n)) {
      files.push({ f, t: f.getLastUpdated().getTime() });
    }
  }
  if (!files.length) throw new Error('No S2006 JSON in case folder');
  files.sort((a, b) => b.t - a.t);
  const parsed = JSON.parse(files[0].f.getBlob().getDataAsString('UTF-8'));
  return generateS2006SheetDraft_(caseInfo, parsed);
}

/** ============== Core ============== **/

function generateS2006SheetDraft_(caseInfo, parsed) {
  // モデル抽出（専用マッパーが保存時に使われているはずだが、保険でフォールバックあり）
  const M = pickS2006Model_(parsed); // { taxes, totals }

  // drafts サブフォルダ
  const parent = DriveApp.getFolderById(caseInfo.folderId);
  const drafts = getOrCreateSubfolder_(parent, 'drafts');

  // テンプレ複製
  const sid =
    (parsed && parsed.meta && parsed.meta.submission_id) ||
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMMddHHmmss');
  const draftName = `S2006_${caseInfo.caseId}_draft_${sid}`;
  const ssId = DriveApp.getFileById(S2006_TPL_GSHEET_ID).makeCopy(draftName, drafts).getId();
  const ss = SpreadsheetApp.openById(ssId);
  const sh = ss.getSheetByName('債権者一覧表（公租公課用）') || ss.getSheets()[0];

  // A列ラベルから行番号を作る
  const ROW = buildRowMapByLabels_(sh);

  // 差し込みヘルパ
  const put = (r, c, v) => (r ? sh.getRange(r, c).setValue(v == null ? '' : v) : null);

  // 各税目: B(2)=amount, C(3)=year, D(4)=payer (+ plate)
  const keys = [
    'income_tax',
    'resident_tax',
    'property_tax',
    'business_tax',
    'national_health_insurance',
    'pension_premium',
    'vehicle_tax',
    'inheritance_tax',
  ];
  keys.forEach((key) => {
    const r = ROW[key];
    const rec = (M.taxes && M.taxes[key]) || {};
    put(r, 2, asIntOrBlank_(rec.amount)); // B
    const yearText = String(rec.year || '').replace(/^(\d{4})年(\d{1,2})月$/, function (_, y, m) {
      return y + '-' + ('0' + m).slice(-2);
    });
    put(r, 3, yearText); // C
    if (key === 'vehicle_tax') {
      const dText = composeVehicleDCell_(rec);
      put(r, 4, dText);
      wrap_(sh.getRange(r, 4));
      try {
        const nextRng = sh.getRange(r + 1, 4);
        const next = nextRng.getDisplayValue().trim();
        if (/登録番号/.test(next) || /^[()（）\s]*$/.test(next)) nextRng.clearContent();
      } catch (_) {}
    } else {
      put(r, 4, rec.payer || '');
    }
  });

  // 合計: テンプレの数式に任せるのが基本。A列に「合計」があるならそこへも書ける。
  const totalRow = ROW.total || findRowByLabel_(sh, '合計');
  if (totalRow) {
    const totalCell = sh.getRange(totalRow, 2);
    if (!totalCell.getFormula()) {
      totalCell.setValue(asIntOrBlank_(M.totals && M.totals.amount_sum));
    }
  }

  try {
    Logger.log(
      '[S2006] draft created caseKey=%s caseId=%s sid=%s fileId=%s url=%s',
      caseInfo.caseKey || '',
      caseInfo.caseId || '',
      sid,
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
  };
}

/** ============== Model Extractor（保険付き） ============== **/

function pickS2006Model_(parsed) {
  // 1) 既に model がある（保存時に専用マッパー使用）→そのまま
  if (parsed && parsed.model && parsed.model.taxes) {
    const taxes = parsed.model.taxes || {};
    const totals = parsed.model.totals || { amount_sum: sumTaxes_(taxes) };
    return { taxes, totals };
  }
  // 2) 専用マッパーがロード済みなら再マッピング
  try {
    if (typeof mapS2006FieldsToModel_ === 'function') {
      const m = mapS2006FieldsToModel_(parsed.fieldsRaw || []);
      if (m && m.taxes)
        return { taxes: m.taxes, totals: m.totals || { amount_sum: sumTaxes_(m.taxes) } };
    }
  } catch (e) {
    try {
      Logger.log('[S2006] mapper error (fallback to generic): %s', (e && e.stack) || e);
    } catch (_) {}
  }
  // 3) 最後の保険：ラベルベースでざっくり拾う
  const taxes = {};
  const read = (re) => {
    const row = (parsed.fieldsRaw || []).find((r) => re.test(String(r.label || '')));
    return row ? String(row.value || '') : '';
  };
  const money = (s) => {
    const n = parseInt(String(s || '').replace(/[^\d]/g, ''), 10);
    return isFinite(n) ? n : 0;
  };
  function setTax(key, ja) {
    taxes[key] = {
      amount: money(read(new RegExp(`^【\\s*${ja}:滞納額`))),
      year: read(new RegExp(`^【\\s*${ja}:年度`)),
      payer: read(new RegExp(`^【\\s*${ja}:納付先`)),
    };
  }
  setTax('income_tax', '所得税');
  setTax('resident_tax', '住民税');
  setTax('property_tax', '固定資産税');
  setTax('business_tax', '事業税');
  setTax('national_health_insurance', '国民健康保険料');
  setTax('pension_premium', '年金保険料');
  taxes.vehicle_tax = {
    amount: money(read(/^【\s*自動車税・軽自動車税:滞納額/)),
    year: read(/^【\s*自動車税・軽自動車税:年度/),
    payer: read(/^【\s*自動車税・軽自動車税:納付先/),
    plate_raw: read(/^【\s*自動車税・軽自動車税:登録番号/),
  };
  setTax('inheritance_tax', '相続税');
  return { taxes, totals: { amount_sum: sumTaxes_(taxes) } };
}

/** ============== Label-driven Row Mapping ============== **/

function buildRowMapByLabels_(sh) {
  // A2:A200 を走査して見出しと行番号を対応付け（テンプレが将来ズレても追従）
  const keyByLabel = {
    所得税: 'income_tax',
    住民税: 'resident_tax',
    固定資産税: 'property_tax',
    事業税: 'business_tax',
    国民健康保険料: 'national_health_insurance',
    年金保険料: 'pension_premium',
    '自動車税・軽自動車税': 'vehicle_tax',
    相続税: 'inheritance_tax',
    合計: 'total',
  };
  const A = sh
    .getRange(2, 1, 199, 1)
    .getDisplayValues()
    .map((r) => r[0]);

  const norm = (s) =>
    String(s || '')
      .replace(/[ 　\t]/g, '') // 空白除去（半/全角）
      .replace(/[・･·∙•]/g, '・'); // 中黒の揺れ統一
  const rev = {};
  Object.keys(keyByLabel).forEach((lbl) => (rev[norm(lbl)] = keyByLabel[lbl]));

  const rowMap = {};
  for (let i = 0; i < A.length; i++) {
    const label = norm(A[i]);
    const key = rev[label];
    if (key) rowMap[key] = i + 2; // 行番号（2起点）
  }
  return rowMap;
}

function findRowByLabel_(sh, labelJa) {
  const A = sh
    .getRange(1, 1, Math.min(300, sh.getMaxRows()), 1)
    .getDisplayValues()
    .map((r) => r[0]);
  const norm = (s) =>
    String(s || '')
      .replace(/[ 　\t]/g, '')
      .replace(/[・･·∙•]/g, '・');
  const target = norm(labelJa);
  for (let i = 0; i < A.length; i++) if (norm(A[i]) === target) return i + 1;
  return 0;
}

/** ============== Misc Utils ============== **/

function sumTaxes_(taxes) {
  let t = 0;
  Object.keys(taxes || {}).forEach((k) => (t += taxes[k].amount || 0));
  return t;
}
function asIntOrBlank_(v) {
  const n = parseInt(String(v ?? ''), 10);
  return isFinite(n) ? n : '';
}

try {
  if (typeof registerAutoDraft === 'function') {
    registerAutoDraft('s2006_creditors_public', function (caseId, submissionId) {
      return run_GenerateS2006DraftBySubmissionId(caseId, submissionId);
    });
  }
} catch (_) {}

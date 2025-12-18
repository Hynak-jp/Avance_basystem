/** =============== form_mapper_s2006.js =============== */

/**
 * 使い方（デバッグ）:
 * debug_SaveSampleS2006_for_0001();
 * → ケース '0001'（case_key: uc13df-0001 等）のフォルダ直下に
 *    s2006_creditors_public__49086456.json を保存（サンプル本文の場合）
 */
function debug_SaveSampleS2006_for_0001() {
  var subject =
    '[#FM-BAS] S2006 債権者一覧表（公租公課用） submission_id:49086456 〔2025年10月17日 10時45分〕';
  var body = String(raw_sample_body_S2006_()).replace(/\r?\n/g, '\n');
  return run_SaveS2006JsonByCaseId('0001', subject, body);
}

// ←通知メール本文（サンプル）
function raw_sample_body_S2006_() {
  return [
    '==== META START ====',
    'form_name: S2006 債権者一覧表（公租公課用）',
    'form_key: s2006_creditors_public',
    'secret: FM-BAS',
    'submission_id: 49086456',
    'case_id: 0001',
    'submitted_at: 2025年10月17日 10時45分',
    'seq: 0001',
    'referrer: https://business-panel.form-mailer.jp/',
    'redirect_url: https://formlist.vercel.app/done?formId=314004',
    '==== META END ====',
    '',
    '==== FIELDS START ====',
    '【メールアドレス】',
    '　design.hayashi@gmail.com',
    '【▼所得税について】',
    '　入力項目を表示',
    '【所得税:滞納額(円)】',
    '　200000',
    '【所得税:年度】',
    '　2025年03月',
    '【所得税:納付先】',
    '　大阪税務署',
    '【▼住民税について】',
    '　入力項目を表示',
    '【住民税:滞納額(円)】',
    '　80000',
    '【住民税:年度】',
    '　2024年01月',
    '【住民税:納付先】',
    '　大阪税務署',
    '【▼固定資産税について】',
    '　入力項目を表示',
    '【固定資産税:滞納額(円)】',
    '　20000',
    '【固定資産税:年度】',
    '　2024年01月',
    '【固定資産税:納付先】',
    '　大阪税務署',
    '【▼事業税】',
    '　入力項目を表示',
    '【事業税:滞納額(円)】',
    '　6000',
    '【事業税:年度】',
    '　2024年01月',
    '【事業税:納付先】',
    '　大阪税務署',
    '【▼国民健康保険料】',
    '　入力項目を表示',
    '【国民健康保険料:滞納額(円)】',
    '　130000',
    '【国民健康保険料:年度】',
    '　2025年01月',
    '【国民健康保険料:納付先】',
    '　大阪税務署',
    '【▼年金保険料】',
    '　入力項目を表示',
    '【年金保険料:滞納額(円)】',
    '　80000',
    '【年金保険料:年度】',
    '　2023年01月',
    '【年金保険料:納付先】',
    '　大阪税務署',
    '【▼自動車税・軽自動車税】',
    '　入力項目を表示',
    '【自動車税・軽自動車税:滞納額(円)】',
    '　20000',
    '【自動車税・軽自動車税:年度】',
    '　2025年01月',
    '【自動車税・軽自動車税:納付先】',
    '　大阪税務署',
    '【自動車税・軽自動車税:登録番号】',
    '　品川 123 さ 45-67',
    '【▼相続税】',
    '　入力項目を表示',
    '【相続税:滞納額(円)】',
    '　30000',
    '【相続税:年度】',
    '　2024年01月',
    '【相続税:納付先】',
    '　大阪税務署',
    '',
    '==== FIELDS END ====',
  ].join('\n');
}

/** ひらがな変換（カタカナ→ひらがな） */
function s2006_toHiragana_(s) {
  return String(s || '').replace(/[\u30A1-\u30F6]/g, function (ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0x60);
  });
}

/** 日本の車両番号 正規化 */
function s2006_normalizeVehiclePlate_(input) {
  if (!input) return { compact: '', pretty: '', parts: null };
  var s = String(input).trim();
  try {
    s = s.normalize('NFKC');
  } catch (_) {}
  s = s.replace(/[‐-–—―ー]/g, '-').replace(/[・·∙•]/g, '・').replace(/[ 　\t]+/g, ' ');
  s = s.replace(/[ 　\-]/g, '');
  s = s2006_toHiragana_(s);

  var firstDigitIdx = s.search(/[0-9]/);
  if (firstDigitIdx < 0) return { compact: s, pretty: s, parts: null };
  var area = s.slice(0, firstDigitIdx);
  var rest = s.slice(firstDigitIdx);

  var classMatch = rest.match(/^(\d{2,3})/);
  if (!classMatch) return { compact: s, pretty: s, parts: null };
  var class_no = classMatch[1];
  rest = rest.slice(class_no.length);

  var kanaMatch = rest.match(/^([ぁ-ゖ])/);
  if (!kanaMatch) return { compact: s, pretty: s, parts: null };
  var kana = kanaMatch[1];
  rest = rest.slice(kana.length);

  var numberDigits = rest.replace(/[^0-9・]/g, '').replace(/・/g, '0').replace(/\D/g, '');
  if (numberDigits.length < 4) numberDigits = ('0000' + numberDigits).slice(-4);
  if (numberDigits.length > 4) numberDigits = numberDigits.slice(-4);

  var compact = area + class_no + kana + numberDigits;
  var pretty =
    area + ' ' + class_no + ' ' + kana + ' ' + numberDigits.slice(0, 2) + '-' + numberDigits.slice(2);
  return {
    compact: compact,
    pretty: pretty,
    parts: { area: area, class_no: class_no, kana: kana, number4: numberDigits },
  };
}

/** S2006 のフィールドを {email, taxes:{...}, totals:{...}} へマッピング */
function mapS2006FieldsToModel_(fields) {
  var taxes = {};
  var email = '';

  function getKeyFromJa_(ja) {
    if (/所得税/.test(ja)) return 'income_tax';
    if (/住民税/.test(ja)) return 'resident_tax';
    if (/固定資産税/.test(ja)) return 'property_tax';
    if (/事業税/.test(ja)) return 'business_tax';
    if (/国民健康保険料/.test(ja)) return 'national_health_insurance';
    if (/年金保険料/.test(ja)) return 'pension_premium';
    if (/自動車税/.test(ja)) return 'vehicle_tax';
    if (/相続税/.test(ja)) return 'inheritance_tax';
    return '';
  }

  function yenToInt_(v) {
    var n = parseInt(String(v || '').replace(/[^\d]/g, ''), 10);
    return isFinite(n) ? n : 0;
  }

  function toYearMonth_(v) {
    var s = String(v || '')
      .trim()
      .replace(/年/g, '-')
      .replace(/月/g, '')
      .replace(/[.\/\s]/g, '-')
      .replace(/-+/g, '-');
    var m1 = s.match(/^(\d{4})-(\d{1,2})$/);
    if (m1) return m1[1] + '-' + ('0' + m1[2]).slice(-2);
    var m2 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m2) return m2[1] + '-' + ('0' + m2[2]).slice(-2);
    return s || '';
  }

  (fields || []).forEach(function (row) {
    var label = String((row && row.label) || '');
    var value = row && row.value;

    if (/メールアドレス/.test(label)) {
      email = String(value || '').trim();
      return;
    }

    var m = label.match(/^【\s*([^:】]+)\s*:\s*([^】]+)\s*】$/);
    if (!m) return;

    var catJa = m[1].trim();
    var fieldJa = m[2].trim();
    var key = getKeyFromJa_(catJa);
    if (!key) return;

    var rec = taxes[key] || (taxes[key] = {});
    if (/滞納額/.test(fieldJa)) {
      rec.amount = yenToInt_(value);
    } else if (/年度/.test(fieldJa)) {
      rec.year = toYearMonth_(value);
    } else if (/納付先/.test(fieldJa)) {
      rec.payer = String(value || '').trim();
    } else if (/登録番号/.test(fieldJa)) {
      var normalized = s2006_normalizeVehiclePlate_(value);
      rec.plate_raw = String(value || '').trim();
      rec.plate = normalized.compact || rec.plate_raw;
      rec.plate_pretty = normalized.pretty || rec.plate_raw;
      if (normalized.parts) rec.plate_parts = normalized.parts;
    }
  });

  var total = 0;
  Object.keys(taxes).forEach(function (key) {
    total += taxes[key].amount || 0;
  });
  return { email: email, taxes: taxes, totals: { amount_sum: total } };
}

// 登録（forms_ingest_core.js のレジストリへ）
var globalObj = typeof globalThis !== 'undefined' ? globalThis : this;
if (!globalObj.FORM_MAPPER_FACTORIES) globalObj.FORM_MAPPER_FACTORIES = {};
globalObj.FORM_MAPPER_FACTORIES.s2006_creditors_public = mapS2006FieldsToModel_;
try {
  registerFormMapper('s2006_creditors_public', mapS2006FieldsToModel_);
} catch (_) {}

/** 件名・本文（通知メール）から解析 → ケース直下に保存 */
function run_SaveS2006JsonByMail(subject, body) {
  return ingestFormMailToCase_(subject, body, { form_key: 's2006_creditors_public' });
}

/** caseId 指定で保存 */
function run_SaveS2006JsonByCaseId(caseId, subject, body) {
  var normalized =
    typeof normalizeCaseIdString_ === 'function'
      ? normalizeCaseIdString_(caseId)
      : String(caseId || '').trim();
  if (!normalized) throw new Error('caseId is required');
  return ingestFormMailToCase_(subject, body, {
    form_key: 's2006_creditors_public',
    case_id: normalized,
  });
}

/** caseKey 指定で保存（caseId を推定） */
function run_SaveS2006JsonByCaseKey(caseKey, subject, body) {
  if (!caseKey) throw new Error('caseKey is required');
  var parts = String(caseKey).trim().split('-');
  var caseIdCandidate = parts.length ? parts[parts.length - 1] : '';
  var normalized =
    typeof normalizeCaseIdString_ === 'function'
      ? normalizeCaseIdString_(caseIdCandidate)
      : ('0000' + String(caseIdCandidate || '').replace(/\D/g, '')).slice(-4);
  if (!normalized) throw new Error('Unable to derive case_id from caseKey');
  return ingestFormMailToCase_(subject, body, {
    form_key: 's2006_creditors_public',
    case_id: normalized,
  });
}

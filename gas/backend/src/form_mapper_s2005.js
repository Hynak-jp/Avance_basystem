/** =============== form_mapper_s2005.js ===============
 * 役割: S2005 フォームの FIELDS 配列をアプリモデルへマッピングするユーティリティ。
 * 注意: フォーム項目のラベル変更時はマッピングの正規表現を必ず更新すること。
 * 使い方（デバッグ）:
 * debug_SaveSampleS2005_for_0001();
 * → ケース '0001' のフォルダ直下に s2005_creditors__<sid>.json を保存
 */

/** デバッグ: サンプル保存（提示いただいた通知メールを内包） */
function debug_SaveSampleS2005_for_0001() {
  var subject = '[#FM-BAS] S2005 債権者一覧表 submission_id:49192017 〔2025年10月21日 14時25分〕';
  var body = String(raw_sample_body_S2005_()).replace(/\r?\n/g, '\n');
  return run_SaveS2005JsonByCaseId('0001', subject, body);
}

// ←通知メール本文（サンプル：ご提示内容をそのまま採用）
function raw_sample_body_S2005_() {
  return [
    '==== META START ====',
    'form_name: S2005 債権者一覧表',
    'form_key: s2005_creditors',
    'secret: FM-BAS',
    'submission_id: 49192017',
    'case_id:0001 ',
    'submitted_at: 2025年10月21日 14時25分',
    'seq: 0001',
    'referrer: https://business-panel.form-mailer.jp/',
    'redirect_url: https://formlist.vercel.app/done?formId=315397',
    '==== META END ====',
    '',
    '==== FIELDS START ====',
    '【メールアドレス】',
    '　design.hayashi@gmail.com',
    '【▼債権者1について】',
    '　入力項目を表示',
    '【債権者1:債権者名】',
    '　債権太郎',
    '【債権者1:現在の債務額】',
    '　120000',
    '【債権者1:借入・購入等の日（開始日）】',
    '　2023年02月01日',
    '【債権者1:借入・購入等の日（終了日）】',
    '　2026年07月21日',
    '【債権者1:使途】',
    '　購入',
    '【債権者1:備考】',
    '　電動自転車の購入',
    '【債権者1:調査票・意見の有無】',
    '　意見有り',
    '【▼債権者2について】',
    '　入力項目を表示',
    '【債権者2:債権者名】',
    '　債権ファイナンス',
    '【債権者2:現在の債務額】',
    '　600000',
    '【債権者2:借入・購入等の日（開始日）】',
    '　2020年10月01日',
    '【債権者2:借入・購入等の日（終了日）】',
    '　2024年11月30日',
    '【債権者2:使途】',
    '　住宅ローン',
    '【債権者2:備考】',
    '',
    '【債権者2:調査票・意見の有無】',
    '',
    '【▼債権者3について】',
    '　入力項目を表示',
    '【債権者3:債権者名】',
    '　債権リフォームローン',
    '【債権者3:現在の債務額】',
    '　1400000',
    '【債権者3:借入・購入等の日（開始日）】',
    '　2020年02月01日',
    '【債権者3:借入・購入等の日（終了日）】',
    '　2023年08月31日',
    '【債権者3:使途】',
    '　住宅ローン',
    '【債権者3:備考】',
    '',
    '【債権者3:調査票・意見の有無】',
    '　調査票有り',
    '【▼債権者4について】',
    '',
    '【債権者4:債権者名】',
    '',
    '【債権者4:現在の債務額】',
    '',
    '【債権者4:借入・購入等の日（開始日）】',
    '',
    '【債権者4:借入・購入等の日（終了日）】',
    '',
    '【債権者4:使途】',
    '',
    '【債権者4:備考】',
    '',
    '【債権者4:調査票・意見の有無】',
    '',
    '【▼債権者5について】',
    '',
    '【債権者5:債権者名】',
    '',
    '【債権者5:現在の債務額】',
    '',
    '【債権者5:借入・購入等の日（開始日）】',
    '',
    '【債権者5:借入・購入等の日（終了日）】',
    '',
    '【債権者5:使途】',
    '',
    '【債権者5:備考】',
    '',
    '【債権者5:調査票・意見の有無】',
    '',
    '【▼債権者6について】',
    '',
    '【債権者6:債権者名】',
    '',
    '【債権者6:現在の債務額】',
    '',
    '【債権者6:借入・購入等の日（開始日）】',
    '',
    '【債権者6:借入・購入等の日（終了日）】',
    '',
    '【債権者6:使途】',
    '',
    '【債権者6:備考】',
    '',
    '【債権者6:調査票・意見の有無】',
    '',
    '',
    '==== FIELDS END ====',
  ].join('\n');
}

/** 数値（円）→ int */
function s2005_yenToInt_(v) {
  var s = String(v == null ? '' : v);
  try {
    s = s.normalize('NFKC');
  } catch (_) {}
  var n = parseInt(s.replace(/[^\d]/g, ''), 10);
  return isFinite(n) ? n : 0;
}

/** 日付 正規化: '2025年10月21日' / '2025-10-21' / '2025/10/21' → 'YYYY-MM-DD' */
function s2005_toYmd_(v) {
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  try {
    s = s.normalize('NFKC');
  } catch (_) {}
  s = s
    .replace(/年/g, '-')
    .replace(/月/g, '-')
    .replace(/日/g, '')
    .replace(/[.\/]/g, '-')
    .replace(/-+/g, '-');
  var m = s.match(/^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?$/);
  if (!m) return '';
  var y = m[1],
    mo = ('0' + m[2]).slice(-2),
    d = m[3] ? ('0' + m[3]).slice(-2) : '01';
  return y + '-' + mo + '-' + d;
}

/** 真偽抽出（調査票・意見の有無 の値から推定） */
function s2005_flagsFromOpinionField_(raw) {
  var s = String(raw == null ? '' : raw).trim();
  return {
    opinion_present: /意見/.test(s),
    investigation_present: /調査票/.test(s),
    raw: s,
  };
}

/** S2005 のフィールドを {email, creditors:[...], totals:{...}} にマッピング */
function mapS2005FieldsToModel_(fields) {
  var email = '';
  var creditorsMap = Object.create(null); // index→rec

  function ensureRec_(idx) {
    var k = String(idx);
    if (!creditorsMap[k]) creditorsMap[k] = { index: idx };
    return creditorsMap[k];
  }

  (fields || []).forEach(function (row) {
    var label = String((row && row.label) || '');
    var value = row && row.value;

    // メール
    if (/メールアドレス/.test(label)) {
      email = String(value || '').trim();
      return;
    }

    // "【債権者N:項目】" 形式のみ対象
    var m = label.match(/^【\s*債権者\s*(\d+)\s*:\s*([^】]+)\s*】$/);
    if (!m) return;

    var idx = parseInt(m[1], 10);
    if (!isFinite(idx)) return;
    var fieldJa = m[2].trim();
    var rec = ensureRec_(idx);

    if (/債権者名/.test(fieldJa)) {
      rec.creditor_name = String(value || '').trim();
    } else if (/現在の債務額/.test(fieldJa)) {
      rec.current_debt_jpy = s2005_yenToInt_(value);
    } else if (/開始日/.test(fieldJa)) {
      rec.date_start = s2005_toYmd_(value);
    } else if (/終了日/.test(fieldJa)) {
      rec.date_end = s2005_toYmd_(value);
    } else if (/使途/.test(fieldJa)) {
      rec.purpose = String(value || '').trim();
    } else if (/備考/.test(fieldJa)) {
      rec.memo = String(value || '').trim();
    } else if (/調査票・意見の有無/.test(fieldJa)) {
      var f = s2005_flagsFromOpinionField_(value);
      rec.opinion_or_sheet_raw = f.raw;
      rec.opinion_present = !!f.opinion_present;
      rec.investigation_present = !!f.investigation_present;
    }
  });

  // map → array & 集計
  var creditors = Object.keys(creditorsMap)
    .map(function (k) {
      return creditorsMap[k];
    })
    .sort(function (a, b) {
      return a.index - b.index;
    })
    .filter(function (r) {
      // 空行は除外（名前 or 金額 or 使途/備考/日付のいずれかがあれば採用）
      return !!(
        (r.creditor_name && r.creditor_name.trim()) ||
        (r.current_debt_jpy && r.current_debt_jpy > 0) ||
        (r.purpose && r.purpose.trim()) ||
        (r.memo && r.memo.trim()) ||
        (r.date_start && r.date_start.trim()) ||
        (r.date_end && r.date_end.trim())
      );
    });

  var totalAmount = 0;
  for (var i = 0; i < creditors.length; i++) {
    totalAmount += creditors[i].current_debt_jpy || 0;
  }

  return {
    email: email,
    creditors: creditors,
    totals: {
      creditor_count: creditors.length,
      amount_sum: totalAmount,
    },
  };
}

// 登録（forms_ingest_core.js のレジストリへ）
//   - S2006 と同様に FACTORIES 経由でも登録しておく
var globalObj = typeof globalThis !== 'undefined' ? globalThis : this;
if (!globalObj.FORM_MAPPER_FACTORIES) globalObj.FORM_MAPPER_FACTORIES = {};
globalObj.FORM_MAPPER_FACTORIES.s2005_creditors = mapS2005FieldsToModel_;
try {
  registerFormMapper('s2005_creditors', mapS2005FieldsToModel_);
} catch (_) {}

/** 件名・本文（通知メール）から解析 → ケース直下に保存 */
function run_SaveS2005JsonByMail(subject, body) {
  return ingestFormMailToCase_(subject, body, { form_key: 's2005_creditors' });
}

/** caseId 指定で保存 */
function run_SaveS2005JsonByCaseId(caseId, subject, body) {
  var normalized =
    typeof normalizeCaseIdString_ === 'function'
      ? normalizeCaseIdString_(caseId)
      : String(caseId || '').trim();
  if (!normalized) throw new Error('caseId is required');
  return ingestFormMailToCase_(subject, body, {
    form_key: 's2005_creditors',
    case_id: normalized,
  });
}

/** caseKey 指定で保存（caseId を推定） */
function run_SaveS2005JsonByCaseKey(caseKey, subject, body) {
  if (!caseKey) throw new Error('caseKey is required');
  var parts = String(caseKey).trim().split('-');
  var caseIdCandidate = parts.length ? parts[parts.length - 1] : '';
  var normalized =
    typeof normalizeCaseIdString_ === 'function'
      ? normalizeCaseIdString_(caseIdCandidate)
      : ('0000' + String(caseIdCandidate || '').replace(/\D/g, '')).slice(-4);
  if (!normalized) throw new Error('Unable to derive case_id from caseKey');
  return ingestFormMailToCase_(subject, body, {
    form_key: 's2005_creditors',
    case_id: normalized,
  });
}

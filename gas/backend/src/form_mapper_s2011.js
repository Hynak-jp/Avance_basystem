/** ===== form_mapper_s2011.js (drop-in replacement) =====
 * S2011（二か月家計）フォーム → 内部モデル 変換（テーブル駆動）
 * - 「その他:金額」は全体一致（誤ヒット防止）
 * - 「雇用保険」は空値を 0 で上書きしない（値があるときのみ反映）
 */

function _nm_(s) {
  return String(s || '')
    .replace(/[ \u3000]/g, '')
    .replace(/[（）]/g, (m) => ({ '（': '(', '）': ')' }[m] || m))
    .replace(/[・、，,]/g, '')
    .toLowerCase();
}
function _pick_(fields, labelLike, exact) {
  const want = _nm_(labelLike);
  for (const f of fields || []) {
    const lab = _nm_(f && f.label);
    const hit = exact ? lab === want : lab.includes(want);
    if (hit) return (f && f.value) != null ? String(f.value) : '';
  }
  return '';
}
function _yen_(v) {
  const s = String(v == null ? '' : v)
    .replace(/[^\d]/g, '')
    .trim();
  if (!s) return null;
  const n = Number(s);
  return isFinite(n) ? n : null;
}
function _set_(obj, path, val) {
  if (val == null || val === '') return;
  const ks = String(path).split('.');
  let o = obj;
  for (let i = 0; i < ks.length - 1; i++) o = o[ks[i]] ?? (o[ks[i]] = {});
  o[ks.at(-1)] = val;
}

const MAP = [
  // 期間
  { src: '家計収支のスタート日', dst: 'start_date', kind: 'text', exact: true },

  // ===== 収入 =====
  { src: '前月からの繰越金額', dst: 'incomes.carryover_prev_month_jpy', kind: 'yen', exact: true },
  {
    src: '給与（申立人分）:受取方法',
    dst: 'incomes.salary_applicant.method',
    kind: 'text',
    exact: true,
  },
  {
    src: '給与（申立人分）:金額',
    dst: 'incomes.salary_applicant.amount_jpy',
    kind: 'yen',
    exact: true,
  },
  { src: '給与（配偶者分）:金額', dst: 'incomes.salary_spouse_jpy', kind: 'yen', exact: true },
  {
    src: '給与（その他同居家族分）',
    dst: 'incomes.salary_family_other.memo',
    kind: 'text',
    exact: true,
  },
  {
    src: '給与（その他同居家族分）:金額',
    dst: 'incomes.salary_family_other.amount_jpy',
    kind: 'yen',
    exact: true,
  },
  {
    src: '自営収入（申立人分）:金額',
    dst: 'incomes.self_employed_applicant_jpy',
    kind: 'yen',
    exact: true,
  },
  {
    src: '自営収入（申立人以外分）:金額',
    dst: 'incomes.self_employed_others_jpy',
    kind: 'yen',
    exact: true,
  },
  { src: '年金（申立人分）:金額', dst: 'incomes.pension_applicant_jpy', kind: 'yen', exact: true },
  { src: '年金（申立人以外分）', dst: 'incomes.pension_others.memo', kind: 'text', exact: true },
  {
    src: '年金（申立人以外分）:金額',
    dst: 'incomes.pension_others.amount_jpy',
    kind: 'yen',
    exact: true,
  },

  // 雇用保険（空なら触らない）
  {
    src: '雇用保険',
    dst: 'incomes.unemployment_insurance_jpy',
    kind: 'yen',
    exact: true,
    writeIfEmpty: false,
  },
  {
    src: '雇用保険:金額',
    dst: 'incomes.unemployment_insurance_jpy',
    kind: 'yen',
    exact: true,
    writeIfEmpty: false,
  },

  { src: '生活保護:受取方法', dst: 'incomes.welfare.method', kind: 'text', exact: true },
  { src: '生活保護:金額', dst: 'incomes.welfare.amount_jpy', kind: 'yen', exact: true },

  { src: '児童（扶養）手当:金額', dst: 'incomes.child_allowance_jpy', kind: 'yen', exact: true },
  { src: '援助:誰からか', dst: 'incomes.support.from', kind: 'text', exact: true },
  { src: '援助:金額', dst: 'incomes.support.amount_jpy', kind: 'yen', exact: true },
  { src: '借入れ:誰からか', dst: 'incomes.borrowing.from', kind: 'text', exact: true },
  { src: '借入れ:金額', dst: 'incomes.borrowing.amount_jpy', kind: 'yen', exact: true },
  { src: 'その他', dst: 'incomes.other_income.memo', kind: 'text', exact: true },
  { src: 'その他:金額', dst: 'incomes.other_income.amount_jpy', kind: 'yen', exact: true }, // 全体一致

  // ===== 支出 =====
  { src: '住居費:種別', dst: 'expenses.housing.type', kind: 'text', exact: true },
  { src: '住居費:金額', dst: 'expenses.housing.amount_jpy', kind: 'yen', exact: true },

  { src: '駐車場代:車名義', dst: 'expenses.parking.owner', kind: 'text', exact: true },
  { src: '駐車場代:金額', dst: 'expenses.parking.amount_jpy', kind: 'yen', exact: true },

  { src: '食費:金額', dst: 'expenses.food_jpy', kind: 'yen', exact: true },
  { src: '嗜好品代:品名', dst: 'expenses.luxury_goods.name', kind: 'text', exact: true },
  { src: '嗜好品代:金額', dst: 'expenses.luxury_goods.amount_jpy', kind: 'yen', exact: true },
  { src: '外食費:金額', dst: 'expenses.eating_out_jpy', kind: 'yen', exact: true },
  { src: '電気代:金額', dst: 'expenses.electricity_jpy', kind: 'yen', exact: true },
  { src: 'ガス代:金額', dst: 'expenses.gas_jpy', kind: 'yen', exact: true },
  { src: '水道代:金額', dst: 'expenses.water_jpy', kind: 'yen', exact: true },

  { src: '携帯電話料金:何人分か', dst: 'expenses.mobile.persons', kind: 'text', exact: true },
  { src: '携帯電話料金:金額', dst: 'expenses.mobile.amount_jpy', kind: 'yen', exact: true },
  {
    src: 'その他通話料・通信料・CATV等:金額',
    dst: 'expenses.telecom_other_jpy',
    kind: 'yen',
    exact: true,
  },

  { src: '日用品代:金額', dst: 'expenses.daily_goods_jpy', kind: 'yen', exact: true },
  { src: '新聞代:金額', dst: 'expenses.newspaper_jpy', kind: 'yen', exact: true },
  {
    src: '国民健康保険料（国民年金）:金額',
    dst: 'expenses.kokuho_or_kokunen_jpy',
    kind: 'yen',
    exact: true,
  },
  {
    src: '保険料（任意加入）:金額',
    dst: 'expenses.optional_insurance_jpy',
    kind: 'yen',
    exact: true,
  },

  { src: 'ガソリン代:車名義', dst: 'expenses.gasoline.owner', kind: 'text', exact: true },
  { src: 'ガソリン代:金額', dst: 'expenses.gasoline.amount_jpy', kind: 'yen', exact: true },
  { src: '交通費:金額', dst: 'expenses.transport_jpy', kind: 'yen', exact: true },
  { src: '医療費:金額', dst: 'expenses.medical_jpy', kind: 'yen', exact: true },
  { src: '被服費:金額', dst: 'expenses.clothing_jpy', kind: 'yen', exact: true },

  { src: '教育費:内訳', dst: 'expenses.education.memo', kind: 'text', exact: true },
  { src: '教育費:金額', dst: 'expenses.education.amount_jpy', kind: 'yen', exact: true },

  { src: '交際費:内訳', dst: 'expenses.social.memo', kind: 'text', exact: true },
  { src: '交際費:金額', dst: 'expenses.social.amount_jpy', kind: 'yen', exact: true },

  { src: '娯楽費:内訳', dst: 'expenses.entertainment.memo', kind: 'text', exact: true },
  { src: '娯楽費:金額', dst: 'expenses.entertainment.amount_jpy', kind: 'yen', exact: true },

  { src: '債務返済実額1:対象者', dst: 'expenses.debt_repayments.0.to', kind: 'text', exact: true },
  {
    src: '債務返済実額1:金額',
    dst: 'expenses.debt_repayments.0.amount_jpy',
    kind: 'yen',
    exact: true,
  },
  { src: '債務返済実額2:対象者', dst: 'expenses.debt_repayments.1.to', kind: 'text', exact: true },
  {
    src: '債務返済実額2:金額',
    dst: 'expenses.debt_repayments.1.amount_jpy',
    kind: 'yen',
    exact: true,
  },

  { src: 'その他1:内容', dst: 'expenses.misc.0.memo', kind: 'text', exact: true },
  { src: 'その他1:金額', dst: 'expenses.misc.0.amount_jpy', kind: 'yen', exact: true },
  { src: 'その他2:内容', dst: 'expenses.misc.1.memo', kind: 'text', exact: true },
  { src: 'その他2:金額', dst: 'expenses.misc.1.amount_jpy', kind: 'yen', exact: true },
];

function mapS2011_TableDriven_(fields, meta) {
  const model = {
    email: (meta && meta.email) || '',
    start_date: _pick_(fields, '家計収支のスタート日', true) || '',
    incomes: {},
    expenses: {},
    totals: {},
    _meta: {},
  };
  for (const r of MAP) {
    const raw = _pick_(fields, r.src, !!r.exact);
    if (r.kind === 'yen') {
      const n = _yen_(raw);
      if (n == null && r.writeIfEmpty === false) continue;
      _set_(model, r.dst, n ?? 0);
    } else {
      if (raw !== '') _set_(model, r.dst, String(raw).trim());
    }
  }
  model._meta.submitted_at = (meta && (meta.submitted_at || meta.received_at)) || '';
  return model;
}

// レジストリ登録（S2011のみ）
try {
  registerFormMapper('s2011_income_m1', mapS2011_TableDriven_);
} catch (_) {}
try {
  registerFormMapper('s2011_income_m2', mapS2011_TableDriven_);
} catch (_) {}

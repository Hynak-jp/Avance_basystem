/** =========================
 * S2010 Part1（経歴等） Intake
 *  - Gmailの通知メール → JSON化 → 案件(caseId)フォルダ直下へ保存
 *  - 既存の S2002 コードのユーティリティ（META/FIELDSパース、Drive保存、cases解決）を再利用
 *    ※ 必須：s2002_draft.js が同プロジェクトに存在すること
 * ========================= */

/** ===== パース＆マッピング（S2010_P1専用） ===== */

function parseFormMail_S2010_P1_(subject, body) {
  const meta = parseMetaBlock_(body);
  const fields = parseFieldsBlock_(body);

  meta.submission_id =
    meta.submission_id ||
    (subject.match(/submission_id:(\d+)/) || subject.match(/提出\s+(\d{6,})/))?.[1] ||
    '';

  const model = mapFieldsToModel_S2010_P1_(fields);
  return { meta, fieldsRaw: fields, model };
}

function mapFieldsToModel_S2010_P1_(fields) {
  const rows = fields.filter(function (f) {
    return !/^【\s*▼/.test(f.label);
  });

  const applicant = {
    email: pickVal_(rows, /^【\s*メールアドレス\s*】$/),
  };

  const jobs = [];
  for (let i = 1; i <= 8; i++) {
    const start = pickIdx_(rows, i, /就業期間-開始日/);
    const end = pickIdx_(rows, i, /就業期間-終了日/);
    const type = pickIdx_(rows, i, /種別/);
    const emp = pickIdx_(rows, i, /就業先/);
    const avg = toIntYen_(pickIdx_(rows, i, /平均月収\(円\)/));
    const sever = pickIdx_(rows, i, /退職金の有無/);
    const role = pickIdx_(rows, i, /地位・業務の内容/);

    if (start || end || type || emp || avg || sever || role) {
      jobs.push({
        index: i,
        start_iso: toIsoYmd_(start),
        end_iso: toIsoYmd_(end),
        type,
        employer: emp,
        avg_month_income: avg,
        severance: (sever || '').trim(),
        role_desc: role || '',
      });
    }
  }

  const marital = [];
  for (let i = 1; i <= 4; i++) {
    const when = pickIdx_(rows, i, /^時期$/);
    const who = pickIdx_(rows, i, /^相手方の氏名$/);
    const why = pickIdx_(rows, i, /^事由$/);
    if (when || who || why) {
      marital.push({
        index: i,
        date_iso: toIsoYmd_(when),
        partner_name: normSpace_(who),
        reason: (why || '').trim(),
      });
    }
  }

  const household = [];
  for (let i = 1; i <= 9; i++) {
    const name = pickIdx_(rows, i, /^氏名$/);
    const rel = pickIdx_(rows, i, /^続柄$/);
    const age = pickIdx_(rows, i, /^年齢$/);
    const job = pickIdx_(rows, i, /^職業・学年$/);
    const live = pickIdx_(rows, i, /^同居・別居$/);
    const inc = toIntYen_(pickIdx_(rows, i, /^平均月収\(円\)$/));
    if (name || rel || age || job || live || inc) {
      household.push({
        index: i,
        name: normSpace_(name),
        relation: (rel || '').trim(),
        age: toInt_(age),
        occupation_or_grade: (job || '').trim(),
        living: (live || '').trim(),
        avg_month_income: inc,
      });
    }
  }

  const inheritances = [];
  for (let i = 1; i <= 4; i++) {
    const who = pickIdx_(rows, i, /被相続人氏名/);
    const rel = pickIdx_(rows, i, /^続柄$/);
    const date = pickIdx_(rows, i, /相続発生日/);
    const stat = pickIdx_(rows, i, /相続状況/);
    if (who || rel || date || stat) {
      inheritances.push({
        index: i,
        decedent_name: normSpace_(who),
        relation: (rel || '').trim(),
        date_iso: toIsoYmd_(date),
        status: (stat || '').trim(),
      });
    }
  }

  const housingRaw = pickVal_(rows, /^【\s*現在の住居の状況（居住する家屋の形態等）\s*】$/);
  const ownedDetail = pickVal_(rows, /^【\s*持ち家\s*】$/);
  const otherName = pickVal_(rows, /^【\s*申立人以外の者の氏名\s*】$/);
  const otherRel = pickVal_(rows, /^【\s*申立人との関係\s*】$/);
  const otherOwnRent = pickVal_(rows, /^【\s*家屋は所有か賃借か\s*】$/);
  const housingEtc = pickVal_(rows, /^【\s*その他\s*】$/);

  const housing = parseHousing_(housingRaw, ownedDetail, {
    otherName,
    otherRel,
    otherOwnRent,
    etc: housingEtc,
  });

  return {
    schema: 's2010_p1_career@v1',
    applicant,
    jobs,
    marital_events: marital,
    household,
    inheritances,
    housing,
  };
}

function pickVal_(rows, re) {
  const r = rows.find(function (f) {
    return re.test(f.label);
  });
  return r ? String(r.value || '').trim() : '';
}

function pickIdx_(rows, idx, nameRe) {
  const re = new RegExp(`^【\\s*${idx}\\s*[：:]\\s*.*${nameRe.source}.*】$`);
  const r = rows.find(function (f) {
    return re.test(f.label);
  });
  return r ? String(r.value || '').trim() : '';
}

function toInt_(v) {
  const n = parseInt(String(v || '').replace(/[^\d]/g, ''), 10);
  return Number.isFinite(n) ? n : '';
}

function toIntYen_(v) {
  return toInt_(v);
}

function toIsoYmd_(ja) {
  if (!ja) return '';
  return toIsoBirth_(ja);
}

function parseHousing_(raw, ownedDetail, extra) {
  const out = { type: '', owner: '', detail: '', notes: '' };

  const s = String(raw || '').trim();
  if (!s) {
    if (ownedDetail) out.detail = ownedDetail.trim();
    if (extra) {
      out.other_name = (extra.otherName || '').trim();
      out.other_relation = (extra.otherRel || '').trim();
      out.other_own_or_rent = (extra.otherOwnRent || '').trim();
      out.notes = (extra.etc || '').trim();
    }
    return out;
  }

  const m = s.match(/^(民間賃貸住宅|公営住宅|持ち家|申立人以外の者|その他)/);
  if (m) out.type = m[1];

  const ownerM = s.match(/（[^）]*所有者[＝=]([^）]+)）/);
  if (ownerM) out.owner = ownerM[1].trim();

  if (ownedDetail) out.detail = ownedDetail.trim();

  if (extra) {
    out.other_name = (extra.otherName || '').trim();
    out.other_relation = (extra.otherRel || '').trim();
    out.other_own_or_rent = (extra.otherOwnRent || '').trim();
    out.notes = (extra.etc || '').trim();
  }

  return out;
}

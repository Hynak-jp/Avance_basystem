/** =========================
 * S2010 Part1（経歴等） Intake
 *  - Gmailの通知メール → JSON化 → 案件(caseId)フォルダ直下へ保存
 *  - 既存の S2002 コードのユーティリティ（META/FIELDSパース、Drive保存、cases解決）を再利用
 *    ※ 必須：s2002_draft.js が同プロジェクトに存在すること
 * ========================= */

/** ===== パース＆マッピング（S2010_P1専用） ===== */

function parseFormMail_S2010_P1_(subject, body) {
  const meta = parseMetaBlock_(body);
  const fields = (parseFieldsBlock_(body) || []).map(function (f, idx) {
    return { label: f.label, value: f.value, _idx: idx };
  });

  meta.submission_id =
    meta.submission_id ||
    (subject.match(/submission_id:(\d+)/) || subject.match(/提出\s+(\d{6,})/))?.[1] ||
    '';

  const model = mapFieldsToModel_S2010_P1_(fields);
  return { meta, fieldsRaw: fields, model };
}

function mapFieldsToModel_S2010_P1_(fields) {
  const rowsWithIndex = (fields || []).map(function (f, idx) {
    return {
      label: f.label,
      value: f.value,
      _idx: typeof f._idx === 'number' ? f._idx : idx,
    };
  });
  const sectioned = s2010p1_buildSectionRows_(rowsWithIndex);
  const rows = sectioned.rows;
  const jobsRows = sectioned.sections.jobs;
  const maritalRows = sectioned.sections.marital;
  const householdRows = sectioned.sections.household;
  const inheritanceRows = sectioned.sections.inheritance;
  const housingRows = sectioned.sections.housing;
  const housingRowsForOther = sectioned.fallback.housing ? [] : housingRows;
  const useJobsFallback = sectioned.fallback.jobs;
  const useMaritalFallback = sectioned.fallback.marital;
  const useHouseholdFallback = sectioned.fallback.household;
  const useInheritanceFallback = sectioned.fallback.inheritance;
  if (useJobsFallback || useMaritalFallback || useHouseholdFallback || useInheritanceFallback || sectioned.fallback.housing) {
    try {
      Logger.log(
        '[S2010_P1] section anchor missing. marital=%s household=%s inherit=%s housing=%s',
        sectioned.anchors.marital,
        sectioned.anchors.household,
        sectioned.anchors.inheritance,
        sectioned.anchors.housing
      );
    } catch (_) {}
  }

  const applicant = {
    email: s2010p1_pickVal_(rows, /^【\s*メールアドレス\s*】$/),
  };

  const jobs = [];
  for (let i = 1; i <= 8; i++) {
    const start = s2010p1_pickInSection_(jobsRows, i, /就業期間-開始日/);
    const end = s2010p1_pickInSection_(jobsRows, i, /就業期間-終了日/);
    const type = s2010p1_pickInSection_(jobsRows, i, /種別/);
    const emp = s2010p1_pickInSection_(jobsRows, i, /就業先/);
    const avg = s2010p1_toIntYen_(s2010p1_pickInSection_(jobsRows, i, /平均月収\(円\)/));
    const sever = s2010p1_pickInSection_(jobsRows, i, /退職金の有無/);
    const role = s2010p1_pickInSection_(jobsRows, i, /地位・業務の内容/);

    if (start || end || type || emp || avg || sever || role) {
      jobs.push({
        index: i,
        start_iso: s2010p1_toIsoYmd_(start),
        end_iso: s2010p1_toIsoYmd_(end),
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
    const when = useMaritalFallback
      ? s2010p1_pickIdxStrict_(maritalRows, i, '時期')
      : s2010p1_pickInSection_(maritalRows, i, /^時期$/);
    const who = useMaritalFallback
      ? s2010p1_pickIdxStrict_(maritalRows, i, '相手方の氏名')
      : s2010p1_pickInSection_(maritalRows, i, /^相手方の氏名$/);
    const why = useMaritalFallback
      ? s2010p1_pickIdxStrict_(maritalRows, i, '事由')
      : s2010p1_pickInSection_(maritalRows, i, /^事由$/);
    if (when || who || why) {
      marital.push({
        index: i,
        date_iso: s2010p1_toIsoYmd_(when),
        partner_name: normSpace_(who),
        reason: (why || '').trim(),
      });
    }
  }

  const household = [];
  for (let i = 1; i <= 9; i++) {
    const name = useHouseholdFallback
      ? s2010p1_pickIdxStrict_(householdRows, i, '氏名')
      : s2010p1_pickInSection_(householdRows, i, /^氏名$/);
    const rel = useHouseholdFallback
      ? s2010p1_pickIdxStrict_(householdRows, i, '続柄')
      : s2010p1_pickInSection_(householdRows, i, /^続柄$/);
    const age = useHouseholdFallback
      ? s2010p1_pickIdxStrict_(householdRows, i, '年齢')
      : s2010p1_pickInSection_(householdRows, i, /^年齢$/);
    const job = useHouseholdFallback
      ? s2010p1_pickIdxStrict_(householdRows, i, '職業・学年')
      : s2010p1_pickInSection_(householdRows, i, /^職業・学年$/);
    const live = useHouseholdFallback
      ? s2010p1_pickIdxStrict_(householdRows, i, '同居・別居')
      : s2010p1_pickInSection_(householdRows, i, /^同居・別居$/);
    const inc = s2010p1_toIntYen_(
      useHouseholdFallback
        ? s2010p1_pickIdxStrict_(householdRows, i, '平均月収(円)')
        : s2010p1_pickInSection_(householdRows, i, /^平均月収\(円\)$/)
    );
    if (name || rel || age || job || live || inc) {
      household.push({
        index: i,
        name: normSpace_(name),
        relation: (rel || '').trim(),
        age: s2010p1_toInt_(age),
        occupation_or_grade: (job || '').trim(),
        living: (live || '').trim(),
        avg_month_income: inc,
      });
    }
  }

  const inheritances = [];
  for (let i = 1; i <= 4; i++) {
    const who = useInheritanceFallback
      ? s2010p1_pickIdxStrict_(inheritanceRows, i, '被相続人氏名')
      : s2010p1_pickInSection_(inheritanceRows, i, /被相続人氏名/);
    const rel = useInheritanceFallback
      ? s2010p1_pickIdxStrict_(inheritanceRows, i, '続柄')
      : s2010p1_pickInSection_(inheritanceRows, i, /^続柄$/);
    const date = useInheritanceFallback
      ? s2010p1_pickIdxStrict_(inheritanceRows, i, '相続発生日')
      : s2010p1_pickInSection_(inheritanceRows, i, /相続発生日/);
    const stat = useInheritanceFallback
      ? s2010p1_pickIdxStrict_(inheritanceRows, i, '相続状況')
      : s2010p1_pickInSection_(inheritanceRows, i, /相続状況/);
    if (who || rel || date || stat) {
      inheritances.push({
        index: i,
        decedent_name: normSpace_(who),
        relation: (rel || '').trim(),
        date_iso: s2010p1_toIsoYmd_(date),
        status: (stat || '').trim(),
      });
    }
  }

  const housingRaw = s2010p1_pickVal_(rows, /^【\s*現在の住居の状況（居住する家屋の形態等）\s*】$/);
  const ownedDetail = s2010p1_pickVal_(housingRows, /^【\s*持ち家\s*】$/);
  const otherName = s2010p1_pickVal_(housingRows, /^【\s*申立人以外の者の氏名\s*】$/);
  const otherRel = s2010p1_pickVal_(housingRows, /^【\s*申立人との関係\s*】$/);
  const otherOwnRent = s2010p1_pickVal_(housingRows, /^【\s*家屋は所有か賃借か\s*】$/);
  const housingEtc = s2010p1_pickVal_(housingRowsForOther, /^【\s*その他\s*】$/);

  const housing = s2010p1_parseHousing_(housingRaw, ownedDetail, {
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

function s2010p1_pickVal_(rows, re) {
  const r = rows.find(function (f) {
    return re.test(f.label);
  });
  return r ? String(r.value || '').trim() : '';
}

function s2010p1_extractInnerLabel_(label) {
  const raw = String(label || '').trim();
  const m = raw.match(/^【\s*(.+?)\s*】$/);
  return m ? m[1] : raw;
}

function s2010p1_isHeadingRow_(row) {
  const label = String((row && row.label) || '');
  const value = String((row && row.value) || '').trim();
  if (/^【\s*▼/.test(label)) return true;
  if (!value && /入力/.test(label)) {
    const inner = s2010p1_extractInnerLabel_(label);
    if (!/^\s*\d+\s*[：:]/.test(inner)) return true;
  }
  return false;
}

function s2010p1_buildSectionRows_(rowsWithIndex) {
  const rows = rowsWithIndex.filter(function (f) {
    return !s2010p1_isHeadingRow_(f);
  });
  const anchors = {
    marital: s2010p1_findHeadingIndex_(rowsWithIndex, /婚姻|離婚|内縁/),
    household: s2010p1_findHeadingIndex_(rowsWithIndex, /配偶者.*同居者|同居者.*配偶者/),
    inheritance: s2010p1_findHeadingIndex_(rowsWithIndex, /相続/),
    housing: s2010p1_findHousingAnchorIndex_(rowsWithIndex),
  };
  const nextAnchor = function (key) {
    return s2010p1_nextAnchorByKey_(anchors, key);
  };
  const sections = {
    jobs:
      anchors.marital < 0
        ? rows
        : s2010p1_sliceRowsFromStart_(rows, anchors.marital),
    marital: anchors.marital < 0 ? rows : s2010p1_sliceRowsByIndex_(rows, anchors.marital, nextAnchor('marital')),
    household:
      anchors.household < 0 ? rows : s2010p1_sliceRowsByIndex_(rows, anchors.household, nextAnchor('household')),
    inheritance:
      anchors.inheritance < 0
        ? rows
        : s2010p1_sliceRowsByIndex_(rows, anchors.inheritance, nextAnchor('inheritance')),
    housing: anchors.housing < 0 ? rows : s2010p1_sliceRowsByIndex_(rows, anchors.housing, -1),
  };
  const fallback = {
    jobs: anchors.marital < 0,
    marital: anchors.marital < 0,
    household: anchors.household < 0,
    inheritance: anchors.inheritance < 0,
    housing: anchors.housing < 0,
  };
  return { rows: rows, sections: sections, anchors: anchors, fallback: fallback };
}

function s2010p1_findHeadingIndex_(rows, re) {
  for (let i = 0; i < rows.length; i++) {
    const label = String(rows[i].label || '');
    if (!label) continue;
    const inner = s2010p1_extractInnerLabel_(label);
    const isHeading = s2010p1_isHeadingRow_(rows[i]) || /▼/.test(inner);
    if (isHeading && re.test(inner)) return rows[i]._idx;
  }
  return -1;
}

function s2010p1_findLabelIndex_(rows, re) {
  for (let i = 0; i < rows.length; i++) {
    const label = String(rows[i].label || '');
    if (!label) continue;
    const inner = s2010p1_extractInnerLabel_(label);
    if (re.test(label) || re.test(inner)) return rows[i]._idx;
  }
  return -1;
}

function s2010p1_findHousingAnchorIndex_(rows) {
  const heading = s2010p1_findHeadingIndex_(rows, /現在の住居|住居の状況/);
  if (heading >= 0) return heading;
  return s2010p1_findLabelIndex_(rows, /現在の住居の状況（居住する家屋の形態等）/);
}

function s2010p1_nextAnchorByKey_(anchors, key) {
  const order = ['marital', 'household', 'inheritance', 'housing'];
  const start = order.indexOf(key);
  if (start < 0) return -1;
  for (let i = start + 1; i < order.length; i++) {
    const v = anchors[order[i]];
    if (v >= 0) return v;
  }
  return -1;
}

function s2010p1_sliceRowsByIndex_(rows, startIdx, endIdx) {
  if (startIdx < 0) return [];
  const end = endIdx > startIdx ? endIdx : -1;
  return rows.filter(function (r) {
    return r._idx > startIdx && (end < 0 || r._idx < end);
  });
}

function s2010p1_sliceRowsFromStart_(rows, endIdx) {
  if (endIdx < 0) return rows;
  return rows.filter(function (r) {
    return r._idx < endIdx;
  });
}

function s2010p1_pickInSection_(rowsSection, idx, nameRe) {
  return s2010p1_pickIdx_(rowsSection, idx, nameRe);
}

function s2010p1_pickIdx_(rows, idx, nameRe) {
  if (!rows || !rows.length) return '';
  const headRe = new RegExp(`^\\s*${idx}\\s*[：:]\\s*(.+?)\\s*$`);
  for (let i = 0; i < rows.length; i++) {
    const inner = s2010p1_extractInnerLabel_(rows[i].label);
    const m = inner.match(headRe);
    if (m && nameRe.test(m[1])) return String(rows[i].value || '').trim();
  }
  return '';
}

function s2010p1_escapeRegExp_(s) {
  return String(s || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function s2010p1_pickIdxStrict_(rows, idx, labelName) {
  if (!rows || !rows.length) return '';
  const name = s2010p1_escapeRegExp_(labelName);
  const reLabel = new RegExp(`^\\s*【\\s*${idx}\\s*[：:]\\s*${name}\\s*】\\s*$`);
  const rePlain = new RegExp(`^\\s*${idx}\\s*[：:]\\s*${name}\\s*$`);
  for (let i = 0; i < rows.length; i++) {
    const label = String(rows[i].label || '').trim();
    if (reLabel.test(label) || rePlain.test(label)) return String(rows[i].value || '').trim();
  }
  return '';
}

function s2010p1_toInt_(v) {
  const n = parseInt(String(v || '').replace(/[^\d]/g, ''), 10);
  return Number.isFinite(n) ? n : '';
}

function s2010p1_toIntYen_(v) {
  return s2010p1_toInt_(v);
}

function s2010p1_toIsoYmd_(ja) {
  if (!ja) return '';
  let raw = String(ja).trim();
  if (!raw) return '';
  raw = raw.replace(/[０-９]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0xfee0));
  const compact = raw.replace(/\s+/g, '');

  const era = compact.match(
    /^(令和|平成|昭和|R|H|S)\s*([0-9]{1,2}|元)(?:年)?(?:[.\/-]?(\d{1,2}))?(?:月)?(?:[.\/-]?(\d{1,2}))?(?:日)?/i
  );
  if (era) {
    const eraRaw = era[1];
    const eraKey = /^[RHS]$/i.test(eraRaw) ? eraRaw.toUpperCase() : eraRaw;
    const base = { 令和: 2018, 平成: 1988, 昭和: 1925, R: 2018, H: 1988, S: 1925 }[eraKey];
    if (!base) return '';
    const eraYear = era[2] === '元' ? 1 : parseInt(era[2], 10);
    if (!Number.isFinite(eraYear)) return '';
    const yyyy = base + eraYear;
    const mm = era[3] ? s2010p1_pad2_(era[3]) : '';
    const dd = era[4] ? s2010p1_pad2_(era[4]) : '';
    if (mm && dd) return `${yyyy}-${mm}-${dd}`;
    if (mm) return `${yyyy}-${mm}`;
    return `${yyyy}`;
  }

  if (/^\d{8}$/.test(compact)) {
    return `${compact.slice(0, 4)}-${compact.slice(4, 6)}-${compact.slice(6, 8)}`;
  }
  if (/^\d{6}$/.test(compact)) {
    return `${compact.slice(0, 4)}-${compact.slice(4, 6)}`;
  }
  if (/^\d{4}$/.test(compact)) {
    return compact;
  }

  const norm = compact
    .replace(/[年月]/g, '-')
    .replace(/[日]/g, '')
    .replace(/[.\/]/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
  const parts = norm.split('-');
  if (parts[0] && /^\d{4}$/.test(parts[0])) {
    const yyyy = parts[0];
    const mm = parts[1] ? s2010p1_pad2_(parts[1]) : '';
    const dd = parts[2] ? s2010p1_pad2_(parts[2]) : '';
    if (mm && dd) return `${yyyy}-${mm}-${dd}`;
    if (mm) return `${yyyy}-${mm}`;
    return `${yyyy}`;
  }
  return '';
}

function s2010p1_pad2_(v) {
  const s = String(v || '').replace(/[^\d]/g, '');
  return s ? s.padStart(2, '0') : '';
}

function s2010p1_parseHousing_(raw, ownedDetail, extra) {
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

// 簡易デバッグ用サンプル（本番導線は変更しない）
function s2010p1_debugSamples_() {
  const mk = function (pairs) {
    return pairs.map(function (p, idx) {
      return { label: p[0], value: p[1], _idx: idx };
    });
  };
  const withJob = mk([
    ['【1:就業期間-開始日】', '2020-01-01'],
    ['【1:就業期間-終了日】', '2021-12-31'],
    ['【1:種別】', '正社員'],
    ['【1:就業先】', 'テスト株式会社'],
    ['【1:平均月収(円)】', '300000'],
    ['【▼ 婚姻】', ''],
    ['【1:時期】', '2022-03'],
    ['【1:相手方の氏名】', '配偶者 花子'],
    ['【1:事由】', '結婚'],
    ['【▼ 配偶者、同居者】', ''],
    ['【1:氏名】', '同居 太郎'],
    ['【1:続柄】', '子'],
    ['【▼ 相続】', ''],
    ['【1:被相続人氏名】', '父'],
    ['【1:続柄】', '子'],
    ['【現在の住居の状況（居住する家屋の形態等）】', 'その他'],
    ['【その他】', '実家に同居中'],
  ]);
  const noJob = mk([
    ['【▼ 婚姻】', ''],
    ['【1:時期】', '2022-03'],
    ['【▼ 配偶者、同居者】', ''],
    ['【1:氏名】', '同居 次郎'],
    ['【▼ 相続】', ''],
    ['【1:被相続人氏名】', '母'],
    ['【現在の住居の状況（居住する家屋の形態等）】', 'その他'],
    ['【その他】', '親族宅に間借り'],
  ]);
  return {
    withJob: mapFieldsToModel_S2010_P1_(withJob),
    noJob: mapFieldsToModel_S2010_P1_(noJob),
  };
}

function s2010p1_debugLogSamples_() {
  const samples = s2010p1_debugSamples_();
  const text = JSON.stringify(samples, null, 2);
  s2010p1_logChunked_('debugSamples', text, 5000);
}

function s2010p1_logChunked_(title, text, chunkSize) {
  const size = Math.max(1, chunkSize || 5000);
  const s = String(text || '');
  if (!s) {
    Logger.log('[S2010_P1][%s] (empty)', title || 'log');
    return;
  }
  const total = Math.ceil(s.length / size);
  for (let i = 0; i < total; i++) {
    const part = s.slice(i * size, (i + 1) * size);
    Logger.log('[S2010_P1][%s %s/%s] %s', title || 'log', i + 1, total, part);
  }
}

/**
 * S2010 Part2（破産申立てに至った事情）メール→JSON 変換
 * - META/FIELDS パースは既存ユーティリティを利用
 * - 自由記述の項目を安全に吸い上げ、よくある見出しをゆるくマッピング
 */

function parseFormMail_S2010_P2_(subject, body) {
  if (typeof parseMetaBlock_ !== 'function' || typeof parseFieldsBlock_ !== 'function') {
    throw new Error('parseMetaBlock_ / parseFieldsBlock_ are required');
  }
  const meta = parseMetaBlock_(body);

  meta.submission_id =
    meta.submission_id ||
    (subject.match(/submission_id:(\d+)/) || subject.match(/提出\s+(\d{6,})/))?.[1] ||
    '';

  const fields = parseFieldsBlock_(body);

  if (typeof mapFieldsToModel_S2010_P2_ !== 'function') {
    throw new Error('mapFieldsToModel_S2010_P2_ is not defined');
  }
  const model = mapFieldsToModel_S2010_P2_(fields);
  return { meta, fieldsRaw: fields, model };
}

function mapFieldsToModel_S2010_P2_(fields) {
  const rows = (fields || []).filter(function (f) {
    return !/^【\s*▼/.test(String(f.label || ''));
  });

  const kv = rows.map(function (row) {
    return {
      label: String(row.label || '').trim(),
      value: String(row.value || '').trim(),
    };
  });

  const summaryKeys = [
    /事情/,
    /経緯/,
    /理由/,
    /背景/,
    /使途/,
    /浪費/,
    /ギャンブル/,
    /投資/,
    /病気|療養/,
    /コロナ|新型.?感染症/,
    /失業|解雇|休業/,
    /売上減|収入減/,
    /家族|扶養/,
  ];

  const longTexts = kv
    .filter(function (entry) {
      return (
        entry.value &&
        entry.value.length >= 10 &&
        summaryKeys.some(function (rx) {
          return rx.test(entry.label);
        })
      );
    })
    .sort(function (a, b) {
      return b.value.length - a.value.length;
    });
  const narrative = longTexts.length ? longTexts[0].value : '';

  const dateLike = /(\d{4}年|\d{4}\/\d{1,2}|令和\d+年)/;
  const timeline = kv
    .filter(function (entry) {
      return entry.value && dateLike.test(entry.value + entry.label);
    })
    .map(function (entry) {
      return {
        hint: entry.label,
        text: entry.value,
      };
    });

  const flags = {
    gambling: kv.some(function (entry) {
      return /ギャンブル|パチンコ|スロット|競馬|競艇|競輪/.test(entry.value + entry.label);
    }),
    wasteful: kv.some(function (entry) {
      return /浪費|衝動買い|ブランド|飲食|交際費/.test(entry.value + entry.label);
    }),
    investment: kv.some(function (entry) {
      return /投資|仮想通貨|FX|株|先物/.test(entry.value + entry.label);
    }),
    medical: kv.some(function (entry) {
      return /病気|入院|療養|手術|障害/.test(entry.value + entry.label);
    }),
    covid: kv.some(function (entry) {
      return /コロナ|新型.?感染症|売上減|休業/.test(entry.value + entry.label);
    }),
    job_loss: kv.some(function (entry) {
      return /失業|解雇|倒産|廃業/.test(entry.value + entry.label);
    }),
  };

  const moneyLike = /[0-9０-９,，．\.]+(円|万円|億)/;
  const amounts = kv
    .filter(function (entry) {
      return moneyLike.test(entry.value);
    })
    .slice(0, 20);

  const reasonsRaw = s2010p2_pickValueByLabel_(kv, /^1\.\s*多額の借金をした理由$/);
  const reasonOther = s2010p2_pickValueByLabel_(kv, /多額の借金をした理由\s*[:：]\s*その他/);
  const reason = s2010p2_buildReason_(reasonsRaw, reasonOther);

  const triggersRaw = s2010p2_pickValueByLabel_(kv, /^2\.\s*返済できなくなったきっかけ$/);
  const triggerOther = s2010p2_pickValueByLabel_(kv, /返済できなくなったきっかけ\s*[:：]\s*その他/);
  const trigger = s2010p2_buildTrigger_(triggersRaw, triggerOther);

  const unableRaw = s2010p2_pickValueByLabel_(kv, /^3\.\s*支払不能になった時期$/);
  const unableYM = s2010p2_toYmpParts_(unableRaw);
  const monthlyRaw = s2010p2_pickValueByLabel_(kv, /当時の約定返済額|約定返済額/);
  const noticeRaw = s2010p2_pickValueByLabel_(kv, /^4\.\s*受任通知発送日$/);
  const notice = s2010p2_toYmdParts_(noticeRaw);
  const circ = s2010p2_buildCirc_(kv);

  return {
    schema: 's2010_p2_cause@v2',
    narrative,
    timeline,
    flags,
    amounts,
    causesRaw: kv,
    reason,
    trigger,
    unable_yyyy: unableYM.yyyy || '',
    unable_mm: unableYM.mm || '',
    unable_monthly_total: s2010p2_toNumberText_(monthlyRaw),
    notice_yyyy: notice.yyyy || '',
    notice_mm: notice.mm || '',
    notice_dd: notice.dd || '',
    circ: circ,
  };
}

function s2010p2_extractInnerLabel_(label) {
  const raw = String(label || '').trim();
  const m = raw.match(/^【\s*(.+?)\s*】$/);
  return m ? m[1] : raw;
}

function s2010p2_pickValueByLabel_(kv, re) {
  const hit = (kv || []).find(function (entry) {
    const label = String(entry.label || '');
    const inner = s2010p2_extractInnerLabel_(label);
    return re.test(label) || re.test(inner);
  });
  return hit ? String(hit.value || '').trim() : '';
}

function s2010p2_splitMulti_(s) {
  return String(s || '')
    .split(/[\r\n、，,・;；]+/)
    .map(function (v) {
      return String(v || '').trim();
    })
    .filter(Boolean);
}

function s2010p2_hasAny_(arr, re) {
  return (arr || []).some(function (v) {
    return re.test(v);
  });
}

function s2010p2_buildReason_(raw, otherText) {
  const reasons = s2010p2_splitMulti_(raw);
  const out = {
    living: s2010p2_hasAny_(reasons, /生活費/),
    mortgage: s2010p2_hasAny_(reasons, /住宅ローン|住宅/),
    education: s2010p2_hasAny_(reasons, /教育/),
    waste: s2010p2_hasAny_(reasons, /浪費|飲食|飲酒|投資|投機|商品購入|ギャンブル/),
    business: s2010p2_hasAny_(reasons, /事業|経営破綻|マルチ|ネットワーク/),
    guarantee: s2010p2_hasAny_(reasons, /保証/),
    other: s2010p2_hasAny_(reasons, /その他/),
    other_text: String(otherText || '').trim(),
  };
  if (out.other_text) out.other = true;
  return out;
}

function s2010p2_buildTrigger_(raw, otherText) {
  const triggers = s2010p2_splitMulti_(raw);
  const out = {
    overpay: s2010p2_hasAny_(triggers, /収入以上|返済金額/),
    dismiss: s2010p2_hasAny_(triggers, /解雇/),
    paycut: s2010p2_hasAny_(triggers, /減額/),
    hospital: s2010p2_hasAny_(triggers, /病気|入院/),
    other: s2010p2_hasAny_(triggers, /その他/),
    other_text: String(otherText || '').trim(),
  };
  if (out.other_text) out.other = true;
  return out;
}

function s2010p2_buildCirc_(kv) {
  const out = [];
  for (let i = 1; i <= 6; i++) {
    const dateRe = new RegExp(`^${i}:日付$`);
    const textRe = new RegExp(`^${i}:内容$`);
    const dateRaw = s2010p2_pickValueByLabel_(kv, dateRe);
    const textRaw = s2010p2_pickValueByLabel_(kv, textRe);
    const ymd = s2010p2_toYmdParts_(dateRaw);
    const yyyy = ymd.yyyy || '';
    const mm = ymd.mm || '';
    const dd = ymd.dd || '';
    let dateIso = '';
    if (yyyy && mm && dd) dateIso = `${yyyy}-${mm}-${dd}`;
    else if (yyyy && mm) dateIso = `${yyyy}-${mm}`;
    else if (yyyy) dateIso = `${yyyy}`;
    out.push({
      date_iso: dateIso,
      yyyy: yyyy,
      mm: mm,
      dd: dd,
      text: String(textRaw || '').trim(),
    });
  }
  return out;
}

function s2010p2_normalizeDigits_(v) {
  return String(v || '').replace(/[０-９]/g, function (d) {
    return String.fromCharCode(d.charCodeAt(0) - 0xfee0);
  });
}

function s2010p2_toNumberText_(v) {
  const norm = s2010p2_normalizeDigits_(v);
  if (typeof s2010_toNumberText_ === 'function') return s2010_toNumberText_(norm);
  const n = String(norm || '').replace(/[^\d]/g, '');
  return n ? n : '';
}

function s2010p2_toYmdParts_(raw) {
  if (!raw) return { yyyy: '', mm: '', dd: '' };
  const normalized = s2010p2_normalizeDigits_(raw);
  if (typeof s2010_toYMD_ === 'function') {
    const d = s2010_toYMD_(normalized) || {};
    return {
      yyyy: d.yyyy || '',
      mm: s2010p2_pad2_(d.mm || ''),
      dd: s2010p2_pad2_(d.dd || ''),
    };
  }
  const m = String(normalized || '')
    .trim()
    .replace(/[年月]/g, '-')
    .replace(/[日]/g, '')
    .replace(/[.\/]/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .split('-');
  return { yyyy: m[0] || '', mm: s2010p2_pad2_(m[1] || ''), dd: s2010p2_pad2_(m[2] || '') };
}

function s2010p2_toYmpParts_(raw) {
  const d = s2010p2_toYmdParts_(raw);
  return { yyyy: d.yyyy || '', mm: d.mm || '' };
}

function s2010p2_pad2_(v) {
  const s = String(v || '').replace(/[^\d]/g, '');
  return s ? s.padStart(2, '0') : '';
}

// デバッグ用: circ 期待形のサンプル
function s2010p2_debugCircSample_() {
  return [
    { date_iso: '2025-12-08', yyyy: '2025', mm: '12', dd: '08', text: '具体的事情の本文1' },
    { date_iso: '2022-06-16', yyyy: '2022', mm: '06', dd: '16', text: '具体的事情の本文2' },
  ];
}

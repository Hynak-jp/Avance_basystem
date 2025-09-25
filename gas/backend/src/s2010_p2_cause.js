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
        entry.value.length >= 40 &&
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

  return {
    schema: 's2010_p2_cause@v1',
    narrative,
    timeline,
    flags,
    amounts,
    causesRaw: kv,
  };
}

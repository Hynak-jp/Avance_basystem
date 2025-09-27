/** header 正規化：大文字小文字/アンダーバー差異を吸収（ex. lineId ⇔ line_id） */
function canonHeaderKey_(k) {
  return String(k || '')
    .trim()
    .replace(/\s+/g, '') // 空白除去
    .replace(/_/g, '') // アンダーバー除去
    .toLowerCase();
}

/** ヘッダー行から "正規化キー → 0-based列Index" のMapを作る */
function buildHeaderIndexMap_(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => {
    const key = canonHeaderKey_(h);
    if (key && map[key] === undefined) map[key] = i;
  });
  return map;
}

/** Map 経由でセル取得（存在しなければ null） */
function getCellByKey_(rowValues, idxMap, keyCanon) {
  const idx = idxMap[keyCanon];
  return idx === undefined ? null : rowValues[idx] ?? null;
}

/** Map 経由でセル設定（Rangeに書く側で +1 する想定：Apps Scriptは1-based） */
function setCellByKey_(sheet, row1Based, idxMap, keyCanon, value) {
  const idx0 = idxMap[keyCanon];
  if (idx0 === undefined) return false;
  sheet.getRange(row1Based, idx0 + 1).setValue(value);
  return true;
}

/** よく使う正規化キーを定義（コード内はコレだけ使う） */
const K = {
  lineId: 'lineid',
  userKey: 'userkey',
  caseId: 'caseid',
  caseKey: 'casekey',
  folderId: 'folderid',
  createdAt: 'createdat',
  lastActivity: 'lastactivity',
  status: 'status',
  // submissions/case_forms系
  formKey: 'formkey',
  seq: 'seq',
  submissionId: 'submissionid',
  receivedAt: 'receivedat',
  jsonPath: 'jsonpath',
  email: 'email',
  emailHash: 'emailhash',
  displayName: 'displayname',
  firstSeenAt: 'firstseenat',
  lastSeenAt: 'lastseenat',
  lastForm: 'lastform',
  lastSubmitYm: 'lastsubmitym',
  notes: 'notes',
  rootFolderId: 'rootfolderid',
  nextCaseSeq: 'nextcaseseq',
  activeCaseId: 'activecaseid',
  updatedAt: 'updatedat',
  intakeAt: 'intakeat',
};

# BAS Case / Intake / Staging Policy (決定版)

## 0. 用語
- **line_id**: LINE ユーザーID（外部キー）
- **user_key**: 連絡先キー（例: `uc13df`）
- **case_id**: 4桁ゼロ埋めの案件ID（例: `"0001"` ※文字列）
- **case_key**: `<user_key>-<case_id>`（例: `uc13df-0001`）
- **staging**: `_email_staging`（および `_staging`）フォルダ
- **knownFolderId**: 直前に作成/確定した案件フォルダID

---

## 1. 不変条件（Invariants）
1. **ケースフォルダは intake 完了時のみ作成**する。`bootstrap` / `status` 経路では新規作成しない。
2. **case_id は常に文字列**で保持・表示する（シートは列単位で `@` 書式固定）。
3. **submissions への追記は1回のみ**。`submission_id + form_key` で重複を抑止する（`submission_id` は半角数字のみ）。
4. **submission_id は常に半角数字**（META 等に不正値が来た場合は `submitted_at` から `yyyyMMddHHmmss` を復元し、それも不可なら現在時刻で再採番）。ACK 行（`ack:...`）は本体着信時に自動削除する。
5. **staging → 案件移送の一致判定は** `case_key` > `case_id` > `line_id` の優先順。
6. **救済の限定 ensure**（status 経路）は「本人の intake が実在」する場合のみ許可する。
7. 署名は **V2（HMAC）を優先**し、**±600秒**の時刻スキュー検証を通過したもののみ処理する。V1はフェイルセーフとして継続。
8. 並行実行を避けるため、**ScriptLock(〜10秒)** ＋（必要に応じて）短命キャッシュを用いる。
9. **通知メール（フォーム取込）は NOTIFY_SECRET を正規化比較で検証**し、ScriptLock + CacheService で競合/二重実行を避けつつ `case_key` 解決 → JSON 保存（同一ファイルは上書き）の流れを踏む。secret 不一致は `FormAttach/Rejected` へ隔離する。

---

## 2. Intake 完了フロー（/api/intake/complete, `doPost`）
1. **認証/署名検証**（V2推奨, ts±600秒）。
2. **user_key 決定 → case_id 決定**
   - 既存が無ければ採番（連番）し、**"0001" 形式の文字列**に正規化。
3. **ケースフォルダ作成**
   - `case_key = <user_key>-<case_id>` を用い **ensure**。戻りは Folder or ID に対応。
4. **payload の正規化＆保存**
   - 保存前に payload（JSON）へ以下を**必ず**書き戻す（欠落時のみ埋める）:
     ```json
     {
       "meta": {
         "line_id": "...",
         "user_key": "uc13df",
         "case_id": "0001",
         "case_key": "uc13df-0001"
       }
     }
     ```
   - ファイル名は `intake__<submission_id>.json`（なければ `Date.now()`）。
   - **案件フォルダ直下**へ保存。
5. **台帳更新**
   - `contacts.active_case_id = "0001"`（列単位で `@` 固定）
   - `cases` を upsert/update
     - `case_id`（文字列）/ `case_key` / `folder_id` / `status='intake'` / `updated_at` ほか
6. **staging 吸い上げ**（すぐ下の §4 の実装を **knownFolderId** 付きで呼ぶ）
7. **submissions 追記**
   - **列名マッピングAPI**で `submissions_appendRow()` を使用
   - 先に `submissions_hasRow_(submission_id,'intake')` で重複ガード
   - `case_id` 列は列単位で `@` 固定、`form_key='intake'`、`referrer` はファイル名

> ポイント: intake 完了直後は **knownFolderId** を渡して吸い上げを実行することで、初回から確実に移送・起票まで終わらせる。

---

## 3. Status ルート（/api/status, `doGet`）
- **新規作成は禁止**。既存フォルダがある場合のみ通常の吸い上げを実行。
- 救済の限定 ensure は **「staging に本人の intake が実在」**かつ**ファイル側に `case_key` か `case_id` がある**場合に限る。
- `getCaseForms_` 等の補助関数は **未ロードでも落ちない**よう `typeof` ガードで空配列/既定値にフォールバック。

---

## 4. 吸い上げ（staging → 案件）アルゴリズム

### 4.1 マッチ判定（優先順位）
```js
function matchesTarget(json, { ukey, cid, ckey, lid }) {
  const m = (json && json.meta) || {};
  const fileCKey = m.case_key || json.case_key || '';
  const fileCID  = m.case_id  || json.case_id  || '';
  const fileLID  = m.line_id  || m.lineId      || '';

  if (fileCKey && fileCKey === ckey) return true;                      // 1) case_key
  if (fileCID  && String(fileCID).padStart(4,'0') === cid) return true;// 2) case_id
  if (fileLID  && String(fileLID) === lid) return true;                // 3) line_id
  return false;
}
```

### 4.2 実施手順
1. **既存フォルダ解決**（作成しない）。intake 完了直後は **knownFolderId** を使用。
2. staging を列挙し、`intake__*.json` を対象に **4.1 の優先順位**で一致判定。
3. 一致したファイルを**案件フォルダへ移送**（作成済フォルダのみ）。
4. `submission_id` は **ファイル名 → JSON → Date.now()** の順で補完し、最終的に半角数字へ正規化（不正フォーマットはその場で再採番）。
5. `submissions` に **列名マッピング＋重複ガード**で1行追記。`cases_forms` を `intake: received` に upsert（任意）。
6. **計測ログ**: `moved`, `appended` 件数を出す。

### 4.3 限定フェイルセーフ
- **knownFolderId があるのに moved=0** の場合、**直近10分で最も新しい intake を1件だけ救済**して移送/起票（不意な過去ファイルを拾わない）。
- **status 救済**でフォルダが無い場合は、**ファイル側が `case_key` を持つ**ときに限って ensure→移送/起票。
- 救済後はその場で `contacts.active_case_id` と `cases.case_key/folder_id` を補正して整合させる。

---

## 5. スプレッドシート・スキーマ/書式
- `contacts.active_case_id` / `cases.case_id` / `submissions.case_id` は **列単位で `@` 書式**に固定。
- `submissions` 必須列: `submission_id`, `form_key`, `case_id`（不足時は **append をスキップ**しログを出す）。`submission_id` は保存時に半角数字へ正規化され、Gmail 取込ログシート（submission_logs）とは区別して管理する。ACK 用の `ack:<caseId>:<formKey>` 行は本体行が入ったタイミングで自動削除し、スイープユーティリティでも残す。
- 旧API（配列 index で `setValues`）は廃止。**列名ベース**関数 `submissions_appendRow` のみ使用。

---

## 6. セキュリティ / 署名
- **V2署名**: `p=<base64(lineId|caseId|ts)>`, `ts`, `sig=HMAC(p)`
  - **±600秒**の時刻スキュー検証を実施（UNIX秒 vs ミリ秒混在に注意）。
- **V1署名**: 互換運用。tsは**ミリ秒へ変換して比較**。
- 秘密鍵は `HMAC_SECRET` を**末尾空白/改行除去**して使用。
- **通知メール（FormMailer）**: META の `secret` と Script Properties `NOTIFY_SECRET`（未設定時は `FM-BAS`）を小文字化＆トリム後に厳密比較。

---

## 7. 競合防止・冪等性
- `statusApi_routeStatus_` など副作用開始点で **ScriptLock(〜10s)** を取得。
- 直近の収集を `CacheService` で `collect:<case_key>`（TTL 60s）にして二重実行を抑止可。
- **全処理は冪等**に設計（同じファイル/同じ submission_id は2回記録されない）。

---

## 8. ログ指針（観測性）
- 入口: `router:in`（keys, action, has_p/ts/sig）
- ルート分岐: `status:route { via: 'v2'|'v1-fallback' }`
- 署名: `sig:p-decoded { has_p, payload_preview }`, `sig:ts-skew`
- 吸い上げ結果: `[collectStaging] lid=... cid=... moved=X appended=Y`
- フェイルセーフ発火: `[collectStaging:FALLBACK] ...`
- submissions スキップ: `[submissions] append skipped (please add headers: submission_id, form_key, case_id)`

---

## 9. テスト項目（抜粋）
- **ログインのみ**でフォルダ/ケースが作成されない。
- **intake 完了**直後に `_email_staging` が空になり、**<user_key>-<case_id>** 直下へ `intake__*.json` が移動。
- `cases.case_id = "0001"`（文字列）, `case_key = <user_key>-0001`, `folder_id` が入る。
- `submissions` に **1 行のみ**（`form_key=intake`, `case_id="0001"`, `referrer`=ファイル名, `status=received`）。
- **meta が line_id を持たない/ case_id を持つ**JSONでも移送できる（`case_key/ case_id` 優先が効く）。
- 署名 ts スキュー境界（±600秒）で許可/拒否の切り替えを確認。
- submissions シートに非数値の `submission_id` 行が残らない（`sheetsRepo_sweepSubmissions_` で掃除可能）。
- `getCaseForms_` 未ロードでも `forms: []` で返る。

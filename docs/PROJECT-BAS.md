# PROJECT-BAS

破産申立自動化システム（BAS）における案件固有の仕様、フォーム設計、Google Drive/GAS 連携ルールをまとめます。
全体のアーキテクチャは [ARCHITECTURE.md](ARCHITECTURE.md) を参照してください。

---

## 1. プロジェクト概要

- 顧客が LINE ログイン後、フォーム入力と添付書類アップロードを行う
- 提出された内容は Google Drive に自動整理され、案件ごとに統合される
- 必須の裁判所提出様式（docx/xls）は GAS により生成される
- ユーザー提出の補足書類（給与明細、通帳コピー等）はそのまま印刷して提出可

---

## 2. 提出書類の分類

### 公式様式（システムで生成）

- S2002 破産手続開始申立書（docx）
  - 現況: GDoc テンプレ差し込みでドラフト生成済み（docx エクスポートは未実装）
- S2005 債権者一覧表（xlsx）
- S2006 債権者一覧表（公租公課用）（xlsx）
- S2009 財産目録（xls）
- S2010 陳述書（doc）
- S2011 家計収支表（xls）
  → 将来的にフォーム入力から自動生成。現状は紙で手書き提出も運用可能。

### 補足書類（顧客提出そのまま提出）

- 給与明細
- 銀行通帳コピー
- 公共料金等の領収書
- その他生活費に関する証憑

---

## 3. キーと役割（line_id / user_key / case_id）

BAS の実運用キーは次の3つです。役割がぶれないように使い分けます。

- line_id（初回受付で確定）
  - LINE のユーザーID。初回受付の導線で署名付きトークンとして渡され、GAS が検証・保存。
  - 以後は表に出さず、裏側の照合にのみ使用する（プライバシー配慮）。
- user_key（恒久キー）
  - `line_id` の先頭6文字の小文字。ユーザー横断での入口や一覧の主キー。
  - 例: `Uc13df94016ee50eb9dd5552bffbe6624` → `uc13df`
- case_id（案件の一次キー）
  - ユーザー単位で採番する4桁ゼロ埋め連番。案件内の全処理（提出物、進捗、権限、請求、遷移）の主キー。
  - 例: `0001`, `0002`, ...。`user_key-case_id` を `case_key` と呼ぶ（例: `uc13df-0001`）。

### 保存ルール

- JSON: `<formKey>__<submissionId>.json`（`meta.case_id` / `meta.case_key` を格納。submissionId は半角数字に正規化し、META が不正値のときは `submitted_at` →14桁タイムスタンプで再採番。Script Properties `ENFORCE_SID_14=1` で常に14桁へ強制可能）
- スタッフ入力: `staff_<caseId>_vN.json`
- 生成物: `S2002_<caseId>.docx`, `S2011_<caseId>.xls` など

### 3.x 管理シート（対応表の明確化）— 追記

**cases シート**を「Drive フォルダ名（=caseKey）」と「表示名（displayName）」の対応表としても使う：

| 列名         | 例               | 説明                                   |
| ------------ | ---------------- | -------------------------------------- |
| lineId       | Uc13df…6624      | ユーザーの LINE ID（完全値）           |
| userKey      | uc13df           | lineId 先頭 6 文字・小文字             |
| caseId       | 0001             | 4 桁連番                               |
| caseKey      | uc13df-0001      | `userKey-caseId`（= Drive フォルダ名） |
| displayName  | テスト 太郎      | 表示用の氏名（Drive には持ち込まない） |
| folderId     | 1AbC…            | `BAS_提出書類/<caseKey>` のフォルダ ID |
| createdAt    | 2025-09-05 10:00 | 生成日時                               |
| status       | draft/ready/...  | 案件ステータス                         |
| lastActivity | 2025-09-05 10:00 | 最終更新                               |

> 既存の **contacts** は `displayName` を保持してよいが、**Drive のフォルダ名には使わない**。

- **contacts**
  - コア: `line_id, user_key, active_case_id, intake_at, updated_at`
  - 任意（将来の追跡/照合用）: `display_name, email, email_hash, last_seen_at, last_form, last_submit_ym, notes`
- **cases**
  - `lineId, caseId, createdAt, status, lastActivity`
- **submissions（任意）**
  - `lineId, formKey, submissionId, fileId, caseId, createdAt`（submissionId は半角数字に正規化）

---

## 4. Google Drive フォルダ構成

> **個人名を Drive 上に出さない方針**のため、ユーザーフォルダ名は
> **`<userKey-caseId>/`（例：`uc13df-0001/`）** に統一する。
> displayName 等の個人情報は Sheets 側の対応表で管理する。

```
BAS_提出書類（ROOT_FOLDER_ID/DRIVE_ROOT_FOLDER_ID: 15QnwkhoXUkh8gVg56R3cSX-fg--PqYdu）/
  └─ <userKey-caseId>/                 # 例: uc13df-0001（個人名は含めない）
       ├─ attachments/                 # 顧客提出ファイル（カテゴリ配下に格納）
       │    ├─ 給与明細/
       │    │    ├─ 202509_PAY.png
       │    │    └─ 202508_PAY_02.jpg
       │    ├─ 銀行通帳/
       │    │    └─ 202509_BANK.jpg
       │    └─ 家計収支表/
       │         └─ 202509_BUDG.jpg
       ├─ staff_inputs/                # スタッフ補足JSON
       │    └─ staff_0001_v1.json
       ├─ drafts/                      # 生成物
       │    ├─ S2002_0001_draft_48029097（Googleドキュメント）
       │    ├─ S2005_0001.xlsx
       │    └─ S2011_0001.xls
       ├─ s2002_userform__48070763.json  # 案件直下（caseKey 直下）に JSON を保存
       └─ intake__48062408.json
```

### 命名ルール

- **userKey**：`lineId` 先頭 6 文字を小文字化（例: `uc13df`）
- **caseId**：ユーザー単位の 4 桁連番（例: `0001`）
- **caseKey**：`userKey-caseId`（例: `uc13df-0001`） ← フォルダ名に使用
- **JSON**：`<formKey>__<submissionId>.json`（`meta.case_id` / `meta.case_key` を保存）
- **添付**：`attachments/<日本語カテゴリ>/YYYYMM_TYPE[_n].ext`（例: `attachments/給与明細/202509_PAY.png`）
- **生成物**：`S20xx_<caseId>.<ext>`（例: `S2002_0001.docx`）
  - S2002（現況）: `drafts/S2002_<caseId>_draft_<submission_id>`（GDoc）。docx は今後対応。
- **管理表**：cases シートに displayName を保持し、caseKey と対応付け

---

## 4.x メール取り込み・保存ポリシー（共通） — 追記

- 安全ガード（すべてのフォーム共通）

  - 件名に `[#FM-BAS]` を含め、META ブロックの `secret` と Script Properties `NOTIFY_SECRET`（未設定時は `FM-BAS`）を突き合わせて検証。
  - 判定は正規化して比較（全角 → 半角、各種ハイフン統一、ゼロ幅スペース除去、空白除去・小文字化）。

- 保存先の決定（添付ファイル・JSON）

  - LINE ID と `caseId` が判明している場合:
    - 添付: `<caseKey>/attachments/<日本語カテゴリ>/...`
  - JSON: `<caseKey>/<formKey>__<submissionId>.json`（submissionId は半角数字）
  - いずれか不明（メールのみ等）の場合:
    - `_email_staging/<YYYY-MM>/<email_hash>/` または `/_staging/<YYYY-MM>/submission_<SID>/` に一時保存
    - **intake は `line_id` が入っても `case_id` が無いので staging 扱い**
  - 後から LINE+caseId が分かった時点で、`_email_staging` から案件フォルダの `attachments/` 配下へ統合（重複は content_hash で回避）。

- カテゴリ判定とファイル名
  - 本文の「【項目名】↔ ファイル名」ペア、およびファイル名ヒューリスティックからカテゴリを決定
  - ファイル名は `YYYYMM_TYPE.ext` とし、カテゴリ毎のフォルダに配置

---

## 5. フロー（正常系）と署名

初回受付で `line_id` を確定・保存し、以後は `case_id` を軸に処理します。

1) LINE起点
- ユーザーが LINE から BAS を開く。サーバ側（Bot/中継）で `line_id` を把握（Messaging API の userId）。

2) 署名トークン生成（V2）
- payload = `line_id|case_id|ts`（初回は `case_id=''`）
- `sig = base64url(HMAC_SHA256(payload, SECRET))`
- `p = base64url(payload)` を付け、formlist/GAS の `/api/bootstrap`（GAS: action=bootstrap）へ誘導

3) GAS で検証（bootstrap）
- p/ts/sig を検証して `line_id` を復元
- `contacts` に `line_id, user_key, display_name` を upsert
- `active_case_id` が無ければ採番し付与（4桁）
- `<user_key>-<case_id>/` を Drive に作成（cases 行を保証・folder_id/status 更新）

4) ケース開始
- 以後の各フォームは `meta.case_id` でひも付け。`line_id` は不要（裏で照合のみ）。

署名方式（兼用）
- V2（推奨・GET）: p/ts/sig を使う。payload=`lineId|caseId|ts` を base64url。
- V1（互換・POST）: `sig = HEX(HMAC_SHA256(
  `${ts}.${lineId}.${caseId}`, SECRET))`（intake_complete など既存互換のため存置）。

Next.js 側の内部 API
- `/api/bootstrap` → GET で GAS の action=bootstrap に V2 署名で到達
- `/api/status` → GET で GAS の action=status に V2 署名で到達（ヘッダ: x-line-id, x-case-id）
- `/api/intake/complete` → POST（V1 互換）

staging 吸い上げ
- `_email_staging` / `_staging` にある `intake__*.json` は、`/api/status`（GAS: action=status）を叩いたタイミングで `<case_key>/` 直下へ移送・重複排除。

---

## 6. フォーム設計（例）

### S2002 破産手続開始申立書フォーム

- 申立人情報（氏名・住所・生年月日）
- 司法書士欄（認定の有無、認定番号、送達場所）

### S2011 家計収支表フォーム（将来対応）

- 対象月（直近 2 か月）
- 収入（給与、家族収入、その他）
- 支出（住居、食費、光熱水道、通信、交通、教育、医療、保険、税金、債務返済、雑費）
- メモ欄
- 自動計算（収入合計、支出合計、差額）

---

## 7. 将来拡張

- OCR は「紙で提出された書類の補助入力」用途に縮小
- フォーム入力から自動生成する公式様式の範囲を拡大
- `caseId` をキーにスタッフ修正・複数回送信をまとめる統合 UI を構築

---

### 7.x 初回ログイン時フロー（bootstrap）

1. Next.js → GAS `/bootstrap`（V2 署名: p/ts/sig）
2. GAS
   - `contacts` を upsert（`user_key` 算出）
   - `active_case_id` を採番 or 既存を使用
   - Drive に `<user_key>-<case_id>/` を保証、`cases` 行を保証・`folder_id`/`status` 更新
3. 以降のフォーム URL には `case_id` を付与（JSON の `meta.case_id` にも保存）

## 8. 最小セット実装（caseId ブートストラップ）

本リリースで導入した「最小セット」の実装要点をまとめます。

- 目的: 初回の `intake_complete` で `activeCaseId` を払い出し、以降のフォーム送信が自動で同一案件に紐づく。

### 8.1 GAS（WebApp）側

- エンドポイント: `doGet(e)` の `action=bootstrap`（V2: p/ts/sig）
- 署名検証: payload=`lineId|caseId|ts` を base64url 署名（p/ts/sig）
- Script Properties: `BAS_MASTER_SPREADSHEET_ID`, `DRIVE_ROOT_FOLDER_ID`（or `ROOT_FOLDER_ID`）
- 処理フロー（冪等）:
  - `action=bootstrap`（GET /exec）
    - 署名検証 → `contacts` を upsert（コア: `line_id, user_key, active_case_id, intake_at, updated_at`。`display_name` などは任意）
    - 既存の `contacts.active_case_id` を読むだけ（未設定なら空のまま）
    - **この時点ではケース採番・フォルダ作成はしない**
    - レスポンス例: `{ ok: true, case_id, caseFolderReady: false }`
  - `action=intake_complete`（POST /exec）
    - `active_case_id` が未設定なら `bs_issueCaseId_()` で **cases の最大値+1** を 4 桁採番して保存
    - `cases` の該当行を保証（`draft` → `intake`）、Drive に `<user_key>-<active_case_id>/` を保証して `cases.folder_id` を更新
    - `contacts.active_case_id` / `contacts.intake_at` を更新
    - 可能なら `submissions` に intake を 1 行 upsert（キーは `submission_id + form_key`、`upsertSubmission_` / `submissions_upsert_` を使用）
    - レスポンス例: `{ ok: true, activeCaseId, caseKey, folderId }`（実装で必要最小に絞ってOK）
- 旧エンドポイント:
  - `doPost_drive(e)` は **deprecated**（明示的に終了レスポンスを返すだけ）。呼び出し元は持たない想定。
- JSON 保存拡張:
  - `saveSubmissionJsonShallow_()` は保存直前に `resolveCaseId_()` を通し、**`meta.case_id` を埋めて保存**する
  - `resolveCaseId_()` の優先順位: `META.case_id` → `contacts.active_case_id`
  - Drive への保存は `<userKey-caseId>/` をルート（個人名は含めない）

### 7.4 既存データの移行ヘルパ — 追記

- 旧ルール（`氏名__LINEID/`）から新ルール（`<userKey-caseId>/`）へ移す補助関数を用意
  - 概要: 旧直下のカテゴリフォルダ配下のファイルを、`<caseKey>/attachments/<カテゴリ>/` へ移送（`content_hash` で重複回避）
  - 実行: GAS エディタから `migrateLegacyUserFolderToCaseKey_(lineId, caseId, displayName)` を呼び出す

### 8.2 Next.js 側

- サーバールート: `/api/bootstrap`（GET → GAS action=bootstrap, V2）
  - 環境変数: `GAS_ENDPOINT`, `BOOTSTRAP_SECRET`（or `TOKEN_SECRET`）
- クライアント: ログイン直後に `/api/bootstrap` を 1 回叩く（冪等）
- フォーム URL 付与: `makeFormUrl(base, lineId, caseId)`（サーバ専用）
  - 付与: `redirect_url[0]` に `lineId/caseId/ts/p/sig`（V2）を載せる
- intake: `makeIntakeUrl(intakeBase, intakeRedirect, lineId, { lineIdQueryKeys: ['line_id[0]'] })`（`caseId=''` で署名）

### 8.3 FormMailer 側（META テンプレート）

- **intake（初回受付）**: `case_id` はメール本文に存在しない前提でOK（GAS が `contacts.active_case_id` から補完して `meta.case_id` に保存する）
  - 代わりに **hidden の `line_id` を META に出力**（URL の `line_id[0]` を受け取る）
- **それ以外のフォーム**: 可能なら `case_id` を含める（例: フォームに hidden `case_id` を持たせ、META に書き出す）

例（intake は `case_id` 無し）

```
==== META START ====
form_name: %%_FORM_NAME_%%
form_key: intake
secret: FM-BAS
line_id: %%line_id%%
submission_id: %%_SUBMISSION_ID_%%
submitted_at: %%_SUBMISSION_CREATED_AT_%%
seq: %%_SEQNUM_%%
referrer: %%_REFERER_%%
client_ip: %%_CLIENT_IP_%%
user_agent: %%_USER_AGENT_%%
redirect_url: https://formlist.vercel.app/done?formId=302516
==== META END ====
```

例（それ以外のフォームは `case_id` を入れるのが望ましい）

```
==== META START ====
form_name: %%_FORM_NAME_%%
form_key: s2002_userform
case_id: <フォームの hidden case_id を差し込み>
==== META END ====
```

`case_id` を含められない場合でも、原則 `contacts.active_case_id` から補完できるようにしておく（ただし「過去案件に紐づけたい」等の例外要件が出たら見直し）。

---

## 9. 命名規約：snake_case / camelCase の運用方針

ゴールデンルール

- データ層（シート・JSON保存・Driveメタ）＝ snake_case を正
- アプリ層（TypeScript/React の変数・プロパティ）＝ camelCase を正
- 双方向の境界では正規化ユーティリティを必ず経由し、snake 優先＋camel 互換で“読む”、snake で“書く”。

使い分けの適用範囲

- snake_case（正）
  - Google Sheets の列名（例：case_id, user_key, folder_id, active_case_id）
  - 永続化される JSON のキー（intake__*.json, submissions 行の JSON 等）
  - Drive 上のメタ情報を JSON 化する際のキー
  - GAS（Apps Script）で台帳に書き戻すときのキー

- camelCase（正）
  - TypeScript/React のローカル変数・props・関数名（例：activeCaseId, folderId）
  - API レスポンスをフロントで一時的に扱う型（ただし保存前に snake へ）

- NG（避ける）
  - シート列に camelCase を新規追加すること（互換維持で“読む”のはOK）
  - フロント→GAS で camel のまま直接書き戻し（正規化関数を噛ませる）

正規化ユーティリティ（境界で必ず通す）

JS/TS（共通ユーティリティ例）

```ts
// snake ⇄ camel を相互変換（浅い構造を想定）
export const toCamel = <T extends Record<string, any>>(row: T) =>
  Object.fromEntries(
    Object.entries(row).map(([k, v]) => [k.replace(/_([a-z])/g, (_,$1)=>$1.toUpperCase()), v])
  ) as any;

export const toSnake = <T extends Record<string, any>>(row: T) =>
  Object.fromEntries(
    Object.entries(row).map(([k, v]) => [k.replace(/[A-Z]/g, c => `_${c.toLowerCase()}`), v])
  ) as any;

// 読み出し時：snake 優先＋camel 互換で拾う
export const pickField = (row: any, snake: string) => {
  const camel = snake.replace(/_([a-z])/g, (_,$1)=>$1.toUpperCase());
  return row[snake] ?? row[camel] ?? '';
};
```

GAS（Apps Script）側のフィールド取得ヘルパ例

```js
function getField_(obj, snake) {
  var camel = snake.replace(/_([a-z])/g, function(_, s){ return s.toUpperCase(); });
  return (obj && (obj[snake] != null ? obj[snake] : obj[camel])) || '';
}

// 例：cases 行
var caseId  = getField_(row, 'case_id');
var userKey = getField_(row, 'user_key');
var folderId= getField_(row, 'folder_id');
```

シート設計ルール

- 既存列は snake に統一（例：active_case_id）。
- 互換期間中は読み取りのみ camel 互換可（getField_/pickField で吸収）。
- 新規列は必ず snake。PR レビューでチェックする。

典型スキーマ

- contacts: `line_id, user_key, display_name, active_case_id, updated_at`
- cases: `case_id, user_key, folder_id, status, created_at, updated_at`
- submissions: `case_id, form_key, submission_id, file_id, saved_at`

API/署名・JSON/保存の扱い

- 署名ペイロードは仕様（V1/V2）に従うが、内部に展開した後のキー名は snake に揃えて JSON 保存。
- フロントの API 型は camel で扱い、保存・書戻しの直前に `toSnake` を通す。

型定義（フロント側例）

```ts
// フロント内部用（camel）
export type CaseRow = {
  caseId: string;
  userKey: string;
  folderId: string;
  status: 'draft'|'intake'|'submitted'|'complete';
  createdAt?: string;
  updatedAt?: string;
};

// 保存時は snake へ
const saveCase = async (row: CaseRow) => {
  const payload = toSnake(row); // { case_id, user_key, ... }
  await fetch('/api/cases/save', { method: 'POST', body: JSON.stringify(payload) });
};
```

マイグレーション方針（既存データに混在がある場合）

- 列名を snake に揃える（例：activeCaseId → active_case_id）。
- GAS 側の読み出しを snake 優先＋camel 互換で暫定運用（getField_ を使用）。
- バッチで既存行の camel キーを snake 列にコピー（手動/スクリプト）。
- 期限を切って camel 列（旧列）を削除。

PR/レビュー・Lint ルール

- サーバ／GAS で書き戻す JSON のキー名が camel のままなら差し戻し。
- 新規シート列に camel があれば差し戻し。
- ESLint：フロントは camel を基本、id 系だけ snake からのマッピングを許容（例：最小限の naming-convention 除外）。

落とし穴と対策

- 表計算の手入力で camel 列が紛れる → スプレッドシートに「注意書き」行を固定表示。
- フォーム→メール→JSON の途中で camel 化 → 取り込み時に snake へ正規化して保存。
- Drive メタからそのまま書戻す → 必ず `toSnake` を通してから保存。

チェックリスト（PR 時に見るポイント）

- [ ] シート列はすべて snake_case
- [ ] 保存系 API は `toSnake` を通している
- [ ] GAS の書戻しは snake キーのみ
- [ ] 読み出しは snake 優先＋camel 互換（getField_/pickField 使用）
- [ ] 新規 JSON（staging/ケースフォルダ）のキーは snake
- [ ] ドキュメントの例コードも命名規約に従っている
### 通知メール（FormMailer → GAS）

- Script Properties `NOTIFY_SECRET` は必須。未設定の場合は取込開始時に即エラーとなる。
- `secret` の照合は **正規化（NFKC/ハイフン統一/空白削除/小文字化）** 後に実施し、件名の `[ #FM-BAS ]` も同ロジックで検証。
- 取込処理は `forms_ingest_core.js` に集約し、ScriptLock + CacheService（`ingest:<messageId>`）で競合・二重実行を抑止。
- 保存前に `case_key`, `received_at`, `received_tz`, `ingest_subject` を補完し、同一 `form_key/submission_id` の JSON は上書き保存。
- `secret` 不一致などの拒否は `FormAttach/Rejected` ラベルへ隔離し、構文エラー等は `FormAttach/Error`、META欠落は `FormAttach/NoMeta`（24h 経過で Archive）へ退避する。

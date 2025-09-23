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

## 3. caseId 仕様

BAS の全データを束ねるためのキーとして **caseId** を導入する。

- **単位**: ユーザーごとの案件
- **userKey**: LINE ID の先頭 6 文字を小文字化
  - 例: `Uc13df94016ee50eb9dd5552bffbe6624` → `uc13df`
- **caseId**: ユーザー単位で採番する 4 桁ゼロ埋め連番
  - 例: `0001`, `0002`, ...
- **caseKey**: `userKey-caseId` で案件を一意に識別
  - 例: `uc13df-0001`

### 保存ルール

- JSON: `<formKey>__<submissionId>.json`（`meta.case_id` に格納）
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
  - `lineId, displayName, userKey, rootFolderId, nextCaseSeq, activeCaseId`
- **cases**
  - `lineId, caseId, createdAt, status, lastActivity`
- **submissions（任意）**
  - `lineId, formKey, submissionId, fileId, caseId, createdAt`

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
- **JSON**：`<formKey>__<submissionId>.json`（`meta.case_id` にも保存）
- **添付**：`attachments/<日本語カテゴリ>/YYYYMM_TYPE[_n].ext`（例: `attachments/給与明細/202509_PAY.png`）
- **生成物**：`S20xx_<caseId>.<ext>`（例: `S2002_0001.docx`）
  - S2002（現況）: `drafts/S2002_<caseId>_draft_<submission_id>`（GDoc）。docx は今後対応。
- **管理表**：cases シートに displayName を保持し、caseKey と対応付け

---

## 4.x メール取り込み・保存ポリシー（共通） — 追記

- 安全ガード（すべてのフォーム共通）

  - 件名に `[#FM-BAS]`、META ブロックに `secret: FM-BAS` を両方含むことを必須とする（`REQUIRE_SECRET=true`）。
  - 判定は正規化して比較（全角 → 半角、各種ハイフン統一、ゼロ幅スペース除去、空白除去・小文字化）。
  - Script Properties: `NOTIFY_SECRET`（未設定時は `FM-BAS`）。

- 保存先の決定（添付ファイル・JSON）

  - LINE ID と `caseId` が判明している場合:
    - 添付: `<caseKey>/attachments/<日本語カテゴリ>/...`
    - JSON: `<caseKey>/<formKey>__<submissionId>.json`
  - いずれか不明（メールのみ等）の場合:
    - `_email_staging/<YYYY-MM>/<email_hash>/` または `/_staging/<YYYY-MM>/submission_<SID>/` に一時保存
  - 後から LINE+caseId が分かった時点で、`_email_staging` から案件フォルダの `attachments/` 配下へ統合（重複は content_hash で回避）。

- カテゴリ判定とファイル名
  - 本文の「【項目名】↔ ファイル名」ペア、およびファイル名ヒューリスティックからカテゴリを決定
  - ファイル名は `YYYYMM_TYPE.ext` とし、カテゴリ毎のフォルダに配置

---

## 5. フォーム設計（例）

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

## 6. 将来拡張

- OCR は「紙で提出された書類の補助入力」用途に縮小
- フォーム入力から自動生成する公式様式の範囲を拡大
- `caseId` をキーにスタッフ修正・複数回送信をまとめる統合 UI を構築

---

### 6.x 初回ログイン時フロー（bootstrap）の要点 — 追記

1. Next.js → GAS `/bootstrap`
   - `lineId`, `displayName`, 署名（HMAC）を送信
2. GAS 側
   - `contacts` を upsert（`userKey` 算出）
   - （アップデート）ログイン時は採番しない。初回の「受付フォーム」送信後に `intake_complete` を GAS に通知し、その時点で初採番・フォルダ作成を行う。
3. 以降のフォーム URL には `caseId` を付与
   - 保存する JSON の `meta.case_id` にも保持

## 7. 最小セット実装（caseId ブートストラップ）

本リリースで導入した「最小セット」の実装要点をまとめます。

- 目的: 初回ログインで `activeCaseId=0001` を払い出し、以降のフォーム送信が自動で同一案件に紐づく。

### 7.1 GAS（WebApp）側

- エンドポイント: `doPost(e)` でブートストラップ受付
- 署名検証: `sig = HMAC_SHA256(base, BOOTSTRAP_SECRET)`
  - `base = lineId + '|' + ts`
  - Script Properties に `BOOTSTRAP_SECRET` を設定
- 処理フロー:
  - `contacts` を upsert（`lineId, displayName, userKey, rootFolderId, nextCaseSeq, activeCaseId`）
  - `activeCaseId` 未設定なら `allocateNextCaseId_()` で 4 桁採番し `setActiveCaseId_()`
  - `cases` シートへ 1 行追加（`draft`）
  - レスポンス: `{ userKey, activeCaseId, rootFolderId }`
- JSON 保存拡張:
  - `saveSubmissionJsonShallow_()` が `meta.case_id` を必ず格納
  - `resolveCaseId_()` が `META.case_id` → `contacts.activeCaseId` の順で決定
  - Drive への保存は `<userKey-caseId>/` をルート（個人名は含めない）

### 7.4 既存データの移行ヘルパ — 追記

- 旧ルール（`氏名__LINEID/`）から新ルール（`<userKey-caseId>/`）へ移す補助関数を用意
  - 概要: 旧直下のカテゴリフォルダ配下のファイルを、`<caseKey>/attachments/<カテゴリ>/` へ移送（`content_hash` で重複回避）
  - 実行: GAS エディタから `migrateLegacyUserFolderToCaseKey_(lineId, caseId, displayName)` を呼び出す

### 7.2 Next.js 側

- サーバールート: `/api/bootstrap`
  - リクエスト: `{ lineId, displayName, ts, sig }`
  - `sig = HMAC_SHA256(lineId + '|' + ts, BOOTSTRAP_SECRET)`
  - 環境変数: `GAS_ENDPOINT`, `BOOTSTRAP_SECRET`
- クライアント: ログイン直後に `/api/bootstrap` を 1 回叩く（冪等）
- フォーム URL 付与: `makeFormUrl(base, lineId, caseId)`
  - 付与パラメータ: `lineId`, `caseId`, `ts`, `sig`
  - `sig = HMAC_SHA256(lineId + '|' + caseId + '|' + ts, BOOTSTRAP_SECRET)`

### 7.3 FormMailer 側（META テンプレート）

通知メールの META ブロックに `case_id` を必ず出す。例:

```
==== META START ====
form_name: %%_FORM_NAME_%%
form_key: s2002_userform
case_id: %%_CASE_ID_%%
==== META END ====
```

これにより GAS 側が `meta.case_id` を JSON へ確実に保存し、マージ処理で案件単位に束ねられます。

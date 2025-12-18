# REPO_STRUCTURE

<!--
Codex へのお願い（常設メモ）：
- このファイルは「ディレクトリ/ファイル構成」と「各コンポーネントの役割・入口・環境変数」を最新に保つためのドキュメントです。
- 新しいディレクトリや重要ファイルを追加/変更/削除したら、必ず本ファイルを更新してください。
- 可能な限り「なぜ必要か（理由）」と「更新日」を残してください。
- セクション構成は維持し、差分が分かるように箇条書き・ツリーと要点説明を併記してください。
-->

最終更新: 2025-10-17（フォーム取込パイプライン共通化・S2006 対応）

このドキュメントは、実際のディレクトリ/ファイル構成と主要コードの役割を反映します。
主なモジュールは **formlist（Next.js）** と **gas（Apps Script）** です。進捗・方針は `docs/PROJECT_STATUS.md` を参照。

---

## 0. 全体像（現状の役割とデータフロー）

- **formlist/**（Next.js, Vercel 想定）
  - LINE で認証（NextAuth）。ログイン後に提出フォーム一覧（外部 FormMailer）を表示
  - 署名付き URL を生成して外部フォームへ遷移（サーバ専用の `formUrl.ts` が V2 署名を生成し、`redirect_url` に `p/ts/sig/lineId/caseId` を付与）
  - 送信完了後は `redirect_url`（本アプリの `/done`）に戻し、ローカル進捗を更新
  - 画像/OCR 抽出の社内 API: `/api/extract` を提供（OpenAI API を利用）
- **gas/**（Google Apps Script）
  - Gmail ラベルを監視して Drive へ「浅い構造」で自動整理、JSON 保存、台帳スプレッドシート更新
  - 添付画像は公開直リンク化 → Next の `/api/extract` に POST → `_model.json` を生成可能
  - 仕上げ用パッチで Google Doc 作成 → DOCX にエクスポート
  - WebApp ルーティング（HMAC 検証）
    - `doGet(action=bootstrap/status)`：V2（GET + p/ts/sig）対応（status は staging 吸い上げの副作用あり）
    - `doPost(action=intake_complete)`：V1（HEX）互換
- **Google Drive / Spreadsheet**
  - `BAS_提出書類/` 配下にユーザー別フォルダを作成・整理
  - JSON: `<formkey>__<submissionId>.json`
  - 添付: `YYYYMM_TYPE[_n].ext`（月フォルダは作らずカテゴリ直下）
  - スプレッドシート: `contacts`, `submissions`, `form_logs` を運用

---

## 1. ルート構成（現況）

.
├─ formlist/ # Next.js アプリ
├─ gas/
│ ├─ backend/ # 本番用 GAS プロジェクト（clasp 連携）
│ └─ hello-world/ # 動作確認用の最小サンプル
├─ docs/ # 補助資料
│ ├─ PROJECT_STATUS.md
│ └─ REPO_STRUCTURE.md（本ファイル）
└─ README.md # リポジトリ案内（雛形）

> **Codex メモ**：新しいトップレベルディレクトリを作成/変更したら、このツリーと下記の各セクションを更新してください。

---

## 2. formlist/（Next.js）

### 2.1 主要パス（実在）

formlist/
├─ src/app/
│ ├─ page.tsx # ルート → /login にリダイレクト
│ ├─ layout.tsx, providers.tsx, globals.scss
│ ├─ login/page.tsx, login/LoginClient.tsx
│ ├─ form/page.tsx # フォーム一覧（外部 FormMailer へのリンク生成）
│ ├─ done/page.tsx, done/DoneClient.tsx # 送信完了で進捗更新して /form へ戻す
│ ├─ api/auth/[...nextauth]/route.ts # NextAuth（LINE）
│ ├─ api/bootstrap/route.ts # GAS bootstrap を叩いて activeCaseId を取得
│ ├─ api/status/route.ts # intake 済みか問い合わせ（副作用あり：staging 吸い上げも起動）
│ ├─ api/intake/complete/route.ts # 受付完了通知 → GAS で初採番
│ ├─ api/avatar/route.ts # 外部アバターを同一オリジン経由でプロキシ
│ └─ api/extract/route.ts # OpenAI を用いた給与明細抽出 API（CORS 対応）
├─ src/components/
│ ├─ FormProgressClient.tsx, FormCard.tsx, ProgressBar.tsx,
│ │ ResetProgressButton.tsx, UserInfo.tsx
│ └─ ui/button.tsx, ui/card.tsx
├─ src/lib/
│ ├─ auth.ts # NextAuth 設定（LINE プロバイダ）
│ ├─ callGas.ts # GAS WebApp を叩くヘルパ（現状は未使用）
│ ├─ progressStore.ts # フォーム進捗（Zustand + localStorage）
│ ├─ formUrl.ts # 署名付きフォームURL生成（サーバ専用 / V2: base64url, payload=lineId|caseId|ts を redirect_url[0] に付与）
│ └─ utils.ts
├─ src/types/next-auth.d.ts # Session/JWT 拡張（lineId 等）
├─ middleware.ts # /form への直接アクセスを /login へ誘導（lineId Cookie 前提）
├─ next.config.ts, tailwind.config.js, postcss.config.js
├─ tsconfig.json, eslint.config.mjs
└─ .env.local（ローカル開発用・秘匿）

### 2.2 このモジュールがすること（現況）

- LINE での認証（NextAuth + LINE）。JWT に `lineId` を格納
- フォーム一覧 `/form` で必要フォームを列挙し、外部 FormMailer へ遷移
- 遷移 URL には `line_id[0]`, `form_id`, `case_id[0]`, `redirect_url[0]=/done?formId=...` を付与
  - `/done` 受信でローカル進捗を `done` に更新 → `/form` へ戻す
- `/api/extract` は GAS からの画像抽出要求を受け、OpenAI Chat Completions で JSON を返す
- `/api/status` と `/api/intake/complete` で intake ゲートを管理（採番は intake 完了時）
- 補足：`middleware.ts` は `lineId` Cookie でガードする設計だが、Cookie の発行処理は未配線（要対応）

### 2.3 環境変数（主要）

- NextAuth: `NEXTAUTH_URL`, `NEXTAUTH_SECRET`, `LINE_CLIENT_ID`, `LINE_CLIENT_SECRET`
- 抽出 API: `OPENAI_API_KEY`, `OPENAI_MODEL`（例: gpt-4oi）, `EXTRACT_SECRET`, `ALLOW_ORIGIN`
- GAS 連携: `GAS_ENDPOINT`, `BOOTSTRAP_SECRET`（bootstrap と署名付きURL用）
- Sheets ステータス API: `BAS_API_ENDPOINT`, `BAS_API_HMAC_SECRET`（GAS status_api.js と共有）
- 参考（旧ヘルパ）: `NEXT_PUBLIC_GAS_URL`, `GAS_SHARED_SECRET`

> 注意：`.env.local` に機微情報が含まれるため、公開リポジトリではコミットしないこと。

> **Codex メモ**：新しい API/ページ/環境変数を追加した場合はここを更新。

---

## 3. gas/（Google Apps Script）

### 3.1 主要ファイル（実在）

gas/backend/src/

- `Gドライブ整理_BAS提出物.js`：Gmail ラベル監視 →Drive 整理・JSON 保存・添付命名・台帳更新。
  - 添付はカテゴリ直下へ `YYYYMM_TYPE[_n].ext` 命名で保存
  - 画像直リンク化 →Next `/api/extract` へ POST（CORS, 秘密鍵ヘッダ対応）
- `仕上げ用パッチ.js`：`_model.json` をもとに Google Doc 作成 →DOCX エクスポート → 台帳追記
- `forms_ingest_core.js`：フォーム通知共通パイプライン。`parseFormMail_` → `ensureSubmissionIdDigits_` → `resolveCaseByCaseId_` → `ensureCaseFolderId_` → `saveSubmissionJson_` を統合し、`NOTIFY_SECRET` 検証と ScriptLock による競合回避、既存 JSON の上書き保存を提供。
- `form_mapper_s2006.js`：S2006（公租公課）専用マッパーとデバッグエントリ。車両番号正規化・`registerFormMapper` での登録・共通パイプライン呼び出しを担当。
- `s2002_draft.js`：S2002 Intake の通知メール取り込み → ケース直下に JSON 保存 → GDoc テンプレ差し込みでドラフト生成（`drafts/` に保存）。`saveSubmissionJson_` は同名 JSON の上書き・重複トラッシュに対応。
- `トリガー（５分おき）.js`：`cron_1min`（既存）と `run_ProcessInbox_S2002` を5分間隔で実行・健全化
- `_model.json が無ければ作ってから finalize まで自動実行.js`：抽出 → モデル化 → 仕上げを一括実行
- `appsscript.json`：スコープ/高度なサービス（Drive, Vision）設定

> Script Properties（抜粋）  
> `NOTIFY_SECRET`（通知検証用シークレット）/ `BAS_MASTER_SPREADSHEET_ID` / `ROOT_FOLDER_ID` / `DRIVE_ROOT_FOLDER_ID` / `ROUTER_MAX_THREADS`（通知処理1バッチ当たりのスレッド上限、既定30）/ `ENFORCE_SID_14`（`1` で submission_id を常に14桁に強制）/ `INGEST_CONFIG_BUST`（値を変えると取込側キャッシュを即時クリア）など。

gas/hello-world/src/

- `Code.js`：最小サンプル

### 3.2 実行形態/トリガ

- 現状は Gmail ラベルベースのバッチ処理（`processLabel` など）で駆動
- S2002 Intake の処理を 5 分間隔（`run_ProcessInbox_S2002`）で実行（メール → JSON → GDoc ドラフト）
  - WebApp ルーティング：`doGet(action=bootstrap/status)`（V2）/ `doPost(action=intake_complete)`（V1）
- `gas/.vscode/tasks.json` に clasp 用タスクあり（push/open/pull）

### 3.3 スプレッドシート（台帳）

- `contacts`：`userKey, lineId, displayName, email, activeCaseId, updatedAt, intakeAt`（運用列。旧列は整理）
- `submissions`：`case_id, case_key, form_key, seq, submission_id, received_at, supersedes_seq, json_path, ...`（status_api / sheets_repo 管理）
- `submission_logs`：Gmail 取込用の簡易ログ（`ts_saved, line_id, form_name, submission_id, ...`）。`submission_id` は半角数字に自動正規化される
- `form_logs`：フォーム原本のログ保管（任意）
- `cases`：`caseId, userKey, lineId, status, folderId, createdAt`（bootstrap で初期行追加）

### 3.4 命名・保存ルール（運用）

- JSON：`<formkey>__<submissionId>.json`（ユーザーフォルダ直下 or staging。`meta.case_id` / `meta.case_key` を必ず格納し、`submissionId` は数字のみのユニークIDに正規化。`submitted_at` から復元できない場合は 14 桁タイムスタンプでフォールバックし、仮ACK（`ack:...`）行は本体行到着時に自動削除）
- 添付：`YYYYMM_TYPE[_n].ext`（カテゴリ直下。例：`202509_PAY.png`）
- 画像直リンク：`https://drive.usercontent.google.com/u/0/uc?id=<ID>&export=download`
- S2002 ドラフト：`drafts/S2002_<caseId>_draft_<submission_id>`（Googleドキュメント）

> **Codex メモ**：Apps Script のファイル増減、WebApp エンドポイントの追加、シート列の変更があればここに反映。

---

## 4. templates/（テンプレート）

- 現在このディレクトリは存在しません（将来作成）。
- 例：`S2002_template.gdoc`（差し込み →DOCX エクスポート）や `S2011_template.xls` など

---

## 5. scripts/（補助・移行）

- 現在このディレクトリは存在しません（将来必要に応じて追加）。

---

## 6. 命名規則（再掲）

- **userKey**：`lineId` 先頭 6 文字を想定（将来運用）。
- **caseId**：ユーザー単位 4 桁連番（運用中）。
- **form JSON**：`<formkey>__<submissionId>.json`（`submissionId` は半角数字。META が不正値の場合は `submitted_at` →14桁タイムスタンプで再採番し、`meta.case_id` / `meta.case_key` を保存）
- **attachments**：`YYYYMM_TYPE[_n].ext`（例：`202509_PAY.png`）
- **drafts**：`S2002_<caseId>.docx` など（テンプレ運用開始後）
  - 現況は GDoc ドラフト（`S2002_<caseId>_draft_<submission_id>`）。docx エクスポートは今後対応。

---

## 7. 開発・デプロイ

- **formlist**
  - Vercel デプロイ想定。`NEXTAUTH_*`, `LINE_*`, `OPENAI_*`, `EXTRACT_SECRET` などを環境に設定
  - GAS から `/api/extract` を叩くため、公開 URL を Script Properties（`NEXT_BASE_URL` 等）に設定
- **gas**
  - clasp で `gas/backend` と紐付け。`appsscript.json` のスコープ/高度サービスを維持
  - Gmail ラベル駆動（`FormAttach/ToProcess` → `FormAttach/Processed`）。時間ベーストリガを設定
- **スプレッドシート**
  - 初回実行で `contacts`, `submissions`, `form_logs` を自動作成（列定義はコード参照）

---

## 8. 未完了/保留（実装と運用）

- [ ] 認証後に `lineId` Cookie を発行し、`middleware.ts` と整合させる
- [x] 2025-09-05 GAS 連携の bootstrap（Next → GAS WebApp）実装
- [x] 2025-09-05 `cases` シートの設計・導入（`caseId` 採番/活性管理）
- [x] 2025-09-09 S2002 Intake → JSON → GDoc ドラフト生成（テンプレ差し込み）
- [ ] S2002 ドラフトの docx エクスポート対応
- [ ] `/api/extract` の堅牢化（レート制限・監査ログ・モデル更新の手順）
- [ ] FormMailer 側の META 埋め込み/命名規則の最終確定
- [ ] docs: `README.md` の雛形 → 実態に合わせて整理

> **Codex メモ**：タスクを進めたら、ここに日付入りで完了印を付けてください（例：`[x] 2025-09-05 完了`）。

---
## 9. 変更履歴

- 2025-10-17: フォーム取込パイプライン（`forms_ingest_core.js`）追加、S2006 マッパー分離、JSON 保存仕様更新
- 2025-09-09: S2002 intake/draft と 5 分トリガ、submissions 列定義を反映
- 2025-09-12: submission_id の数値化・ACK 行クリーンアップ、staging intake の仕様更新を反映
- 2025-09-05: caseId ブートストラップ / 署名付きフォームURL / GAS doPost を反映
- 2025-09-04: 現状コードに同期（formlist/API 実体・GAS 構成・シート定義の整合、未実装箇所明示）

# AGENTS.md（BAS 専用：破産申立支援システム）

> このファイルは**AI コーディングエージェント（Codex 等）への指示書**です。人間向けの README は別途用意しつつ、ここではエージェントが安全・確実に開発を進めるための**実行手順・規約・注意点**を具体的に記載します。

---

## 0. コミュニケーション規約（最重要）

- 以後の応答は**常に日本語（です・ます調）**。
- コード識別子は英語、**コメントは日本語**。
- まず**要点（サマリ）→ 変更計画（箇条書き）→ 実行手順（必ず作業ディレクトリを明記）→ 変更差分 or コード**の順で提示。
- **すべてのコマンドは作業ディレクトリ（cwd）を明記**： 1 行目に `# cwd: <repo/subdir>` を付与。
- 破壊的変更・外部通信・秘密情報の取扱いが絡む場合は**必ず事前に確認**し、プランを提示してから実行。

---

## 1. リポジトリ構成（BAS 想定）

```
<repo-root>/
├─ formlist/                 # Next.js (BASフロント：フォーム送信・署名生成・GAS連携)
├─ gas/
│  ├─ backend/               # GAS Web App（受け口API、署名検証、台帳/Drive連携）
│  └─ tools/                 # GASユーティリティ（Gmail/Drive整備、メンテ用）
├─ contracts/                # API契約・JSON Schema・型共有
├─ scripts/                  # デプロイ/運用スクリプト（clasp/環境変数同期など）
├─ .github/                  # CI（lint/typecheck/test/build）
└─ .vscode/ など
```

---

## 2. 開発の共通原則

- **Node.js**: `>= 20.x` / **pnpm**: `>= 9.x`
- **エディタ**: VS Code（ESLint, Prettier, Codex）。保存時 `eslint --fix` を有効化。
- **型/静的解析**: TypeScript `strict: true`。`pnpm lint` / `pnpm typecheck` を PR 前に必ず実行。
- **テスト**: `pnpm test`（Vitest/Jest）。存在しない場合は**最小テストを自動追加**してから変更。
- **ブランチ/PR**:

  - ブランチ: `feature/<topic>` / `fix/<topic>` / `chore/<topic>`
  - PR タイトル: `[scope] 目的`（例: `[formlist] 署名検証の例外処理を追加`）
  - PR 前チェック: `pnpm lint && pnpm test && pnpm build`（対象スコープのみ可）

- **WSL2/Windows 配慮**: 重い処理・依存導入は WSL 内で実行。Windows/WSL パス混在に注意。

---

## 3. セキュリティ / 秘密情報

- **絶対にコミットしない**: `.env*`, `*.pem`, `service-account*.json`, `*.pfx` 等。
- ログや PR 本文に**機密値を書かない**（`GAS_ENDPOINT`, `BOOTSTRAP_SECRET`, API Key 等）。サンプルは `****` でマスク。
- 自動化スクリプトは**失敗時に副作用を巻き戻す**（DB/ファイル生成をロールバック、リトライは指数バックオフ）。

---

## 4. ルート共通タスク

```bash
# cwd: <repo-root>
pnpm install
pnpm typecheck
pnpm lint
pnpm test
```

---

## 5. サブプロジェクト別ガイド

### 5.1 `formlist/`（Next.js / BAS フロント）

**目的**: 破産申立支援の UI・フォーム送信・HMAC 署名・GAS 連携。

**主要環境変数（`.env.local`）**

- `GAS_ENDPOINT` : GAS Web App のエンドポイント URL
- `BOOTSTRAP_SECRET` : HMAC 署名用シークレット
- （必要に応じて）`ALLOW_DEBUG`, `LOG_LEVEL`, `NEXT_PUBLIC_*`

**初期化**

```bash
# cwd: formlist
cp .env.local.example .env.local   # 必要項目を埋める（コミット禁止）
pnpm install
```

**開発/ビルド/起動/検証**

```bash
# cwd: formlist
pnpm dev          # 開発
pnpm build        # 本番ビルド
pnpm start        # 本番起動
pnpm lint         # ESLint
pnpm typecheck    # tsc --noEmit
pnpm test         # 単体/コンポーネントテスト
```

**品質ゲート**

- 変更時は最低 1 つのユニットテストを追加（フォーム検証/署名ユーティリティなど）。
- 署名（HMAC）は**時刻依存**のため、時刻をモックしてテスト。
- フォーム送信用の\*\*API 契約（JSON スキーマ）\*\*を`/contracts/`に置き、差分は PR で説明。

**危険操作の扱い**

- 外部通信（GAS/Drive 等）はダミー値で**ドライラン**→ 差分プラン提示 → 実行。

---

### 5.2 `gas/backend/`（GAS Web App）

**目的**: 受信 API（doPost）、署名検証、台帳（Spreadsheet）/Drive 連携、Gmail ラベル駆動の処理。

**Script Properties（例）**

- `BAS_MASTER_SPREADSHEET_ID`
- `DRIVE_ROOT_FOLDER_ID`
- `BOOTSTRAP_SECRET`
- （運用する場合）`LABEL_TO_PROCESS`, `LABEL_PROCESSED`, `LABEL_ERROR`

**主なエントリ/関数（例）**

- `doPost(e)`: 署名検証 → 受付 → 台帳追記/フォルダ作成 等
- `setupTriggers()`: `cron_1min`（または `every 5 min`）を登録
- `cron_1min()`: ラベル付きメールを処理（Gmail→Drive 保存/整形）
- `ensureLabels()`: 必要ラベルの存在保証

**デプロイ（clasp）**

```bash
# cwd: gas/backend
npx clasp login
npx clasp push          # スクリプトを反映
npx clasp deploy        # 新バージョンをデプロイ（Web App）
# デプロイ後、発行URLを控えて formlist の GAS_ENDPOINT に反映
```

**トリガー**

- トリガーはコードで**再現可能**に保つ（`setupTriggers()`）。
- 重複作成を避けるため、既存トリガーの削除 → 新規作成の手順を徹底。

**ロギング**

- `Logger.log` / 実行ログ（Apps Script ダッシュボード）。
- 必要に応じて Stackdriver（現在の名称）相当のログ出力を追加。

**スプレッドシート列命名規約**

- `contacts`, `cases`, `cases_forms`, `submissions` など運用シートの列名は **snake_case**（例: `case_id`, `line_id`, `can_edit`）に統一します。
- 互換ユーティリティ（`bs_toIndexMap_`, `sheetsRepo_getValue_` など）が camelCase も吸収しますが、新規列は必ず snake_case で追加してください。
- 列追加・名称変更時は、関連スクリプト（`sheets_repo.js`, `bootstrap.js`, `status_api.js` など）の alias 対応を確認した上で反映します。

**Status API 運用メモ**

- `/exec?action=markReopen` は **POST + JSON (Content-Type: application/json)** で送信します。GET 要求は 405 となるため禁止です。
- `status_api_collectStaging_` が `_staging` 配下の `intake__*.json` を案件フォルダへ吸い上げるため、`status`/`markReopen` を呼び出すとステージングが自動整理されます。
- 冪等性確保のため、同一 `ts/sig` の再送は `nonce_reused` エラーになります。再送時は `ts` を更新して署名を再計算してください。
- intake フォーム通知（`form_key: intake`）には `case_id` は含まれません。GAS で採番した値を JSON へ書き戻してから案件フォルダへ保存することが前提です。

---

### 5.3 `gas/tools/`（GAS ユーティリティ）

**目的**: Gmail 通知の正規化、ファイル命名規則適用、Drive 整備、簡易メンテ等。

**注意点**

- 本番データを扱うため、**ドライランモード**を実装（`DRY_RUN=1` 相当のフラグ）。
- フォルダ名・ファイル名規約の変更時は、**影響範囲の確認**と**マイグレーション手順**を PR に含める。

---

## 6. リリース/デプロイ指針

- **formlist（Vercel 等）**: `pnpm build` が通った PR のみデプロイ。`.env*`は**環境側で設定**。
- **GAS（backend）**: `clasp deploy` で新バージョンを発行。**エンドポイント URL が変わる**場合は、運用手順に沿って `GAS_ENDPOINT` を**確実に更新**し、**旧 URL は無効化**。

---

## 7. 変更提案テンプレ（エージェント用）

> PR 前にこのフォーマットで提示してください：

**要点**

- （1〜3 行）

**変更計画**

- [ ] 影響範囲
- [ ] 実装手順
- [ ] ロールバック手順

**実行手順**

```bash
# cwd: <repo/subdir>
# 1) ...
# 2) ...
```

**テスト**

```bash
# cwd: <repo/subdir>
pnpm test
```

**差分**

- 主要ファイルのパッチ or 生成物

---

## 8. よくある落とし穴（BAS）

- `.env.local` を誤コミット / PR で平文貼付 → **禁止**。必要なら `.env.local.example` を更新。
- **署名エラー `bad_sig`**: （例）改行/空白混入、`ts`の単位（秒/ミリ秒）不一致、HEX の大文字小文字差。
- **トリガー多重**: `setupTriggers()` 実装で既存削除 → 再作成を徹底。
- **Gmail ラベル漏れ**: `ensureLabels()` を先に実行し、ラベル名の不一致を潰す。
- **CORS/通信失敗**: `fetch`のヘッダ・メソッド・`Content-Type` を API 契約通りに。
- **タイムゾーン**: フロント/サーバ/スプレッドシートでの時刻整合をテストする。

---

## 9. 将来の拡張メモ

- **Firestore/Cloud Run** 等への移行検討（スケール時）。
- **LIFF** 連携（将来の個別プッシュ通知・深い連携が必要になったタイミングで）。
- GitHub Actions で `pnpm lint && pnpm typecheck && pnpm test` をワークスペース毎に並列実行。

---

# （雛形）サブプロジェクト用 AGENTS.md（BAS）

## `formlist/AGENTS.md`（雛形）

### コミュニケーション

- 常に日本語。要点 → 計画 → 手順 → コード。
- コマンドは `# cwd: formlist` を付与。

### セットアップ

```bash
# cwd: formlist
cp .env.local.example .env.local
pnpm install
```

### スクリプト

```bash
# cwd: formlist
pnpm dev
pnpm build && pnpm start
pnpm lint && pnpm typecheck && pnpm test
```

### 注意点

- HMAC・タイムスタンプの**境界テスト**を追加。
- GAS 連携は**ドライラン**→ 本実行の順。

---

## `gas/backend/AGENTS.md`（雛形）

### 目的

- `doPost` 受付、署名検証、台帳/Drive 連携、Gmail ラベル処理、トリガー管理。

### デプロイ

```bash
# cwd: gas/backend
npx clasp login
npx clasp push
npx clasp deploy   # Web App URL を控える
```

### プロパティ

- `BAS_MASTER_SPREADSHEET_ID`, `DRIVE_ROOT_FOLDER_ID`, `BOOTSTRAP_SECRET` などを **Script Properties** に設定。

### 注意

- トリガー重複/暴走を避けるため、`setupTriggers()` で一元管理。
- ログ監視とレート制限（Gmail/Drive のクォータ）に注意。

# BAS 自動申立システム

**破産申立書類の自動生成・提出補助システム**
LINE ログイン → フォーム入力 → Drive 整理 → OCR → 裁判所提出様式（docx/xls）生成までを支援します。

---

## 🔗 Quick Links

- **全体アーキテクチャ**：[`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md)
- **BAS 案件の仕様（caseId など）**：[`docs/PROJECT-BAS.md`](docs/PROJECT-BAS.md)
- **リポジトリ構成（どこに何があるか）**：[`docs/REPO_STRUCTURE.md`](docs/REPO_STRUCTURE.md)
- **現在の進捗・方向性**：[`docs/PROJECT_STATUS.md`](docs/PROJECT_STATUS.md)
- **GAS 実装の落とし穴と対処（必読）**：[`docs/OPS_GAS_PITFALLS.md`](docs/OPS_GAS_PITFALLS.md)

---

## ⚠️ 初期セットアップの要点（落とし穴回避）

- **Web アプリのデプロイ ID は 1 本固定**。**新規デプロイは作らず** GAS UI の「デプロイを管理」→ **既存 Web アプリの行を編集 → デプロイ（更新）**で中身だけ差し替える。
- **エントリポイントは doPost を 1 つだけ**（他ファイルの doPost は `doPost_xxx` 等に改名）。
- **Script Properties（ID は“ID のみ”）**
  - `BOOTSTRAP_SECRET`：任意の秘密（formlist と同一）
  - `ROOT_FOLDER_ID`：`https://drive.google.com/drive/folders/{ID}` の `{ID}`
  - `BAS_MASTER_SPREADSHEET_ID`：`https://docs.google.com/spreadsheets/d/{ID}/edit` の `{ID}`
- **Spreadsheet → Sheet は二段取得**（`openById(...).getSheetByName(...)` のチェーン禁止）。
- 署名は **`HMAC_SHA256(\`\${lineId}|\${ts}\`, BOOTSTRAP_SECRET)` を base64url（末尾=除去）**。

> さらに詳しいアンチパターンと対処は **[`docs/OPS_GAS_PITFALLS.md`](docs/OPS_GAS_PITFALLS.md)** を参照。

---

## 🚀 Getting Started（最短）

### 1. formlist (Next.js)

```bash
cd formlist
cp .env.local.example .env.local
# .env.local に GAS_ENDPOINT, BOOTSTRAP_SECRET, LINE_ID, DISPLAY_NAME を設定
npm install
npm run dev
```

→ http://localhost:3000 へ。LINE ログイン後にフォーム一覧が表示。

### 2. gas (Google Apps Script)

- `gas/` のコードを GAS プロジェクトへ反映（clasp/手動どちらでも可）
- **デプロイは UI から既存 Web アプリの行を「編集 → デプロイ（更新）」**
  - 実行ユーザー: 自分 / アクセス: 全員（匿名）
- **Script Properties を設定**
  - `BOOTSTRAP_SECRET`, `ROOT_FOLDER_ID`, `BAS_MASTER_SPREADSHEET_ID`（すべて **ID のみ**）

### 3. ヘルスチェック & 署名スモークテスト

```bash
# GET ヘルス（"ok":true,"VER":... が返ればOK）
node -e 'fetch(new URL(process.env.GAS_ENDPOINT)).then(r=>r.text()).then(t=>console.log(t))'

# 署名POST
node <<'NODE'
const { createHmac } = require('crypto');
const ep=process.env.GAS_ENDPOINT, sec=process.env.BOOTSTRAP_SECRET, id=process.env.LINE_ID, name=process.env.DISPLAY_NAME||'';
(async()=>{
  const ts=Math.floor(Date.now()/1000);
  const sig=createHmac('sha256',sec).update(`${id}|${ts}`,'utf8').digest('base64url');
  const u=new URL(ep); u.searchParams.set('lineId',id); u.searchParams.set('ts',String(ts)); u.searchParams.set('sig',sig);
  const r=await fetch(u,{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({lineId:id,displayName:name,ts,sig})});
  console.log('status:',r.status); console.log('body:',await r.text());
})();
NODE
```

> `{"error":"bad_sig"}` が出たら、`docs/OPS_GAS_PITFALLS.md` の「bad_sig が続く」を参照（SECRET 不一致 / base64url / ts ずれ などを即切り分け）。

---

## 🧪 よくある問題（ショートカット）

- **404**：ライブラリにデプロイしている or `GAS_ENDPOINT` の ID 不一致 → UI で Web アプリの行を更新し、その `/exec` を使う
- **bad_sig**：SECRET 不一致 / 署名対象違い / base64url 未対応 / ts ミス → `debug=1` 応答で provided/expect を突合
- **invalid_spreadsheet_id**：URL やフォルダ ID を入れている → **ID のみ**に修正
- **TypeError: ss.getSheetByName**：1 行チェーンで Sheet を `ss` に代入している → 二段取得に統一

> 詳細手順・スニペットは **[`docs/OPS_GAS_PITFALLS.md`](docs/OPS_GAS_PITFALLS.md)**。

---

## 📂 リポジトリ構成（抜粋）

```
repo-root/
├─ formlist/                 # Next.js アプリ（Vercelデプロイ可）
├─ gas/                      # Google Apps Script (clasp 推奨)
├─ docs/
│  ├─ ARCHITECTURE.md        # 全体像・依存関係・データフロー
│  ├─ REPO_STRUCTURE.md      # ディレクトリとファイル構成の詳細
│  ├─ PROJECT_STATUS.md      # 現在の進捗・方向性
│  ├─ PROJECT-BAS.md         # 案件固有の仕様・フォーム・GAS・Drive連携
│  ├─ WORKSPACE.md           # VSCode/Codex 環境の使い方
│  ├─ ENVIRONMENTS.md        # .env / secrets の扱い
│  ├─ RUNBOOK.md             # 運用手順（デプロイ/ロールバック等）
│  ├─ OPS_GAS_PITFALLS.md   # ← ここに既知の落とし穴を集約
│  ├─ CODING_GUIDE.md        # コーディング規約・ディレクトリ規約
│  └─ GIT_CONVENTIONS.md     # ブランチ戦略・コミット規約
└─ .github/
   ├─ ISSUE_TEMPLATE.md
   └─ PULL_REQUEST_TEMPLATE.md
```

---

## 🤝 コントリビュート

- コーディング規約: `docs/CODING_GUIDE.md`
- Git 戦略/コミット規約: `docs/GIT_CONVENTIONS.md`
- 技術判断は `docs/ADR/` へ（必要に応じて）

---

## 📜 ライセンス

（ここにライセンスを記入）

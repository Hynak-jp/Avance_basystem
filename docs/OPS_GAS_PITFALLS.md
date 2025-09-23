# OPS_GAS_PITFALLS.md
BAS/GAS Known Pitfalls & Quick Fixes (歴史踏まえた再発防止メモ)

> プロジェクトで実際にハマった箇所を、**症状 → 原因 → 即応（一手） → 予防**の順で短く。
> コードとコマンドはコピペで使えるよう最小構成にしてあります。

---

## TL;DR（まずこれだけ守る）
- **Webアプリは 1 本のデプロイIDに固定**し、**“編集→デプロイ(更新)”のみ**。新規デプロイは作らない。
- **`doPost` はプロジェクト内で 1 個だけ**（他は `doPost_xxx` に改名）。
- Script Properties は **キー名・値の形式を厳密に**：
  - `BOOTSTRAP_SECRET` … 任意の秘密（前後空白なし）
  - `ROOT_FOLDER_ID` … `https://drive.google.com/drive/folders/{ID}` の `{ID}`
  - `BAS_MASTER_SPREADSHEET_ID` … `https://docs.google.com/spreadsheets/d/{ID}/edit` の `{ID}`
- **Spreadsheet → Sheet は二段取得**（`openById(...).getSheetByName(...)` のチェーン禁止）。
- 署名方式は**統一**：`HMAC-SHA256( lineId + "|" + ts )` を **base64url**（末尾 `=` 除去）で比較。

---

## 既知の落とし穴（履歴から）

### 1) `/exec` が **404** になる
- **症状**: GET/POST が 404。ログも出ない。
- **原因**: ライブラリにデプロイ（`/macros/library/d/...`）、または `GAS_ENDPOINT` と実際に更新したデプロイIDが不一致。
- **即応**: GAS UIで **Webアプリの行**を開き、編集→**デプロイ（更新）**。その行の `/exec` URLを `GAS_ENDPOINT` に使用。
- **予防**: 新規デプロイ禁止。**同一WebアプリID固定** + UI更新のみ。  
  一致チェック：
  ```bash
  node -e 'const u=process.env.GAS_ENDPOINT||"";const m=u.match(/\/s\/([^/]+)\/exec$/);console.log({equal:(m&&m[1])===process.env.DEPLOYMENT_ID})'
  ```

### 2) `bad_sig` が続く（署名NG）
- **症状**: 常に `{error:"bad_sig"}`。
- **原因**（主）:
  1. `BOOTSTRAP_SECRET` の **不一致/余分な空白**  
  2. **署名対象の不一致**（JSON本文をHMACしていた／正解は `"lineId|ts"`）  
  3. **base64 vs base64url** の取り違え（パディング `=` 付き/無し）  
  4. **`ts` の単位/差**（ms を送っている、body と query の `ts` がズレ）
- **即応**: `debug=1` で **providedSig/expect/secretFP** を返すブロックで切り分け。
  ```js
  if (String(qs.debug ?? body.debug) === '1') {
    const base = lineId + '|' + ts;
    const raw  = Utilities.computeHmacSha256Signature(base, SECRET, Utilities.Charset.UTF_8);
    const expect = Utilities.base64EncodeWebSafe(raw).replace(/=+$/,'');
    const secretFP = Utilities.base64EncodeWebSafe(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SECRET, Utilities.Charset.UTF_8)
    ).replace(/=+$/,'').slice(0,16);
    return json({ ok:true, base, providedSig, expect, secretFP, VER });
  }
  ```
- **予防**:
  - **計算式統一**：`base = \`\${lineId}|\${ts}\`` → `HMAC-SHA256(base, SECRET)` → **base64url**（末尾 `=` 除去）
  - **許容時差**：GAS側で ±300s。body と query の `ts` は**同一値**。

### 3) `doPost` が **複数**あった
- **症状**: 期待のロジックに入らない／ログに到達しない／`bad_sig`連発。
- **原因**: 別ファイルにも `function doPost(e){...}` があり、**そちらが入口を奪取**。
- **即応**: 余計な `doPost` を `doPost_xxx` に改名。
- **予防**: 入口は**1本**に統一（必要ならディスパッチ）。

### 4) スクリプトプロパティの **キー/値** ミス
- **症状**: `... is empty` で落ちる、または後続で型エラー。
- **原因**: キー名の揺れ（`ROOT_FOLDER_ID` vs `DRIVE_ROOT_FOLDER_ID`）、**フルURLを入れている**、余分な空白/改行。
- **即応**: 本プロジェクトは **`ROOT_FOLDER_ID` と `BAS_MASTER_SPREADSHEET_ID` に統一**（IDのみ）。
- **予防**: 入力例  
  - `ROOT_FOLDER_ID` → `https://drive.google.com/drive/folders/{ID}` の `{ID}`  
  - `BAS_MASTER_SPREADSHEET_ID` → `https://docs.google.com/spreadsheets/d/{ID}/edit` の `{ID}`  
  - `BOOTSTRAP_SECRET` → 前後空白なし

### 5) `TypeError: ss.getSheetByName is not a function`
- **症状**: シート取得で型エラー。
- **原因**: `SpreadsheetApp.openById(...).getSheetByName(...)` の **1行チェーン**で、`ss` に **Sheet** を代入後さらに `getSheetByName` を呼んでいる／Sheets API(JSON)を誤用。
- **即応**: **二段取得**に統一：
  ```js
  const ss = SpreadsheetApp.openById(id);        // Spreadsheet
  let sh   = ss.getSheetByName(name) || ss.insertSheet(name); // Sheet
  ```
- **予防**: チェーン**全面禁止**。CI/grepで検知：
  ```bash
  grep -nR --include='*.js' -E 'openById\(.*\)\.getSheetByName\(' src
  ```

### 6) ログが見えない・出ない
- **症状**: Web実行で何も見えない。
- **原因**: 間違ったデプロイ／`console.log` だけ使っている／Log Explorer のフィルタ。
- **即応**: **入口ピン**を常設：
  ```js
  Logger.log(JSON.stringify({ pin:'bootstrap-enter', VER, now: Math.floor(Date.now()/1000) }));
  ```
  Web実行のログは **Cloud Logging（script.googleapis.com/console_logs）** に出る。
- **予防**: `debug=1` 応答を常設。GET `/exec` でヘルス返却（`{"ok":true,"VER":...}`）。

### 7) `403` が返る
- **症状**: Forbidden。
- **原因**: Webアプリの **アクセス権が “全員” ではない**／「実行するユーザー」が不適切。
- **即応**: GAS UI → Webアプリの行 → 編集 →  
  「実行するユーザー：自分」「アクセス：全員」で **デプロイ（更新）**。

### 8) `ts=NaN` など環境変数ハマり
- **症状**: 署名が合わない／`base = ...|NaN`。
- **原因**: Eval用コマンドの書き方ミス、未定義の環境変数。
- **即応**: テストでは **コード内で `ts` を生成**（Envに依存しない）。

### 9) クエリとボディの **不一致**
- **症状**: `bad_ts` / `ts_mismatch`。
- **原因**: クエリの `ts/sig` とボディの `ts/sig` が違う値。
- **即応/予防**: **同一値**を両方へ付与（将来の受け側変更にも強い）。

---

## スニペット置き場

### Node: ヘルスチェック（GET）
```bash
node -e 'fetch(new URL(process.env.GAS_ENDPOINT)).then(r=>r.text()).then(t=>console.log(t))'
```

### Node: 署名POST（本番）
```bash
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

### GAS: 署名比較コア（WebSafe + 末尾`=`除去）
```js
const providedSig = String((body||{}).sig ?? (qs||{}).sig ?? (body||{}).signature ?? (qs||{}).signature ?? '')
  .replace(/=+$/,'');
const raw = Utilities.computeHmacSha256Signature(`${lineId}|${ts}`, SECRET, Utilities.Charset.UTF_8);
const expect = Utilities.base64EncodeWebSafe(raw).replace(/=+$/,'');
if (providedSig !== expect) return json({ error:'bad_sig', VER, providedSig, expect }, 200);
```

---

## プロパティ定義（最終形）
| Key                         | 値の例/形式                                       | 備考 |
|---                          |---                                                |---|
| `BOOTSTRAP_SECRET`          | ランダム文字列（前後空白なし）                    | 署名用 |
| `ROOT_FOLDER_ID`            | `15QnwkhoXUkh8gVg56R3cSX-fg--PqYdu`               | DriveフォルダID |
| `BAS_MASTER_SPREADSHEET_ID` | `1Gy5fqViwwU96fpEZe-xzyoh0XbPXoifpnQaM`           | SpreadsheetのID |

---

## チェックリスト

**Before deploy**
- [ ] `doPost` は1つだけ
- [ ] Script Properties（上表の3キー）を確認
- [ ] 署名デバッグブロックが有効
- [ ] `openById` → `getSheetByName` は二段で統一

**After deploy (更新)**
- [ ] `GET /exec` が `{"ok":true,"VER":...}` を返す
- [ ] 署名POST（debug=1）で `providedSig === expect`
- [ ] 本番POSTで 200 + 業務レス
- [ ] Cloud Logging にピンログが出る

---

最小運用ルール：**WebアプリID固定・UI更新のみ・入口1本・プロパティIDのみ**。  
これで“404地獄 / bad_sig地獄 / シート型エラー”は封じ込められます。

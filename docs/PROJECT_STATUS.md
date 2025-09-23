# BAS 破産自動申立システム - プロジェクトステータス

## 現状（2025-09-09 時点）

### Next.js アプリ

- [x] LINE ログイン実装済み
- [x] フォーム一覧ページ（/form）実装済み
- [x] Vercel にデプロイ成功
- [x] 初回ログイン時に GAS bootstrap API 呼び出し（実装）

### GAS スクリプト

- [x] Gmail → Drive 自動整理（ユーザーフォルダ作成、添付保存）
- [x] JSON 保存ルールをシンプル化（`formKey__submissionId.json`）
- [x] 添付ファイルは `YYYYMM_TYPE[_n].ext` 命名でカテゴリ直下に保存
- [x] OCR 抽出（.ocr.json, .model.json）まで稼働
- [x] caseId 採番（ユーザー単位 4 桁連番）導入・運用開始（bootstrap 経由）
- [x] WebApp `doPost` 実装（HMAC 署名検証）
- [x] S2002 Intake メール取り込み → JSON 保存 → S2002 ドラフト生成（gdoc テンプレ差し込み）
- [ ] S2002 の docx エクスポート（gdoc からのエクスポート処理）
- [x] 時間主導トリガー整備（`cron_1min`, `run_ProcessInbox_S2002` を5分間隔で実行）

### スプレッドシート（BAS_master）

- [x] contacts シート：lineId, displayName, userKey, rootFolderId, nextCaseSeq, activeCaseId
- [x] cases シート：caseId, userKey, lineId, status, folderId, createdAt（lastActivity 等は任意列）
- [x] submissions シート：lineId, formKey, submissionId, fileId（json_id）, caseId（任意）

### データモデル

- userKey: LINE ID 先頭 6 文字（例: `uc13df`）
- caseId: ユーザー単位で採番される 4 桁（例: `0001`）
- caseKey = `userKey-caseId` で一意に識別

---

## 今後の方向性

- [ ] **初回ログイン時 bootstrap**

  - GAS WebApp に POST → caseId 採番 → activeCaseId を返却
  - Next.js 側でフォーム URL に `?caseId=0001&sig=...` を付与
  - [x] 2025-09-05 実装・連携完了

- [ ] **S2002 破産手続開始申立書の docx エクスポート**

  - フォーム入力とスタッフ入力をマージして差し込み
  - チェックボックスは □→☑ に置換

  - [ ] **S2011 家計収支表(司法書士作成用・4.1).xls 生成**

  - ユーザーがフォーム入力
  - JSON → Google Sheets テンプレに差し込み → Excel (.xls) でエクスポート

- [ ] **スタッフ入力フォーム**

  - 司法書士欄、補足項目、補正内容を入力
  - staff\_<caseId>\_vN.json として保存

- [ ] **OCR 結果の転記**

  - 給与明細 → 転記先は未定 ※家計収支表（S2011.xls）には転記しない
  - 通帳 → 財産目録（S2009.xls）
  - 完璧でなくても数値が入るデモを優先

- [ ] **デモ台本**
  1. ユーザーフォーム送信 → 自動整理（Drive）
  2. スタッフフォーム入力 → caseId 紐づけ
  3. S2002.docx 生成
  4. OCR で給与明細から金額抽出 → S2011.xls に転記

---

## リスク・確認ポイント

- [x] ユーザー提出書類（例：家計収支表）は裁判所提出に「公式様式（xls/doc）」への転記は不要
- [ ] LINE ID が変わった場合のユーザー統合方法
- [x] HMAC 署名による GAS WebApp セキュリティ確保（BOOTSTRAP_SECRET）
- [ ] OCR 精度：手書き家計簿は低め、デモは給与明細中心で見せる

---

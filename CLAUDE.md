# tamaki-zaisan-dashboard — 公開可否の札

最終更新: 2026-05-29

## ★★★ このリポの分類（global CLAUDE.md「データ種別×公開可否マトリクス」準拠）

| 項目 | 値 |
|---|---|
| **データ種別** | 会社情報（会計・経営）のみ |
| **GitHub公開設定** | **PUBLIC**（GitHub Pages配信のため） |
| **アクセス制御** | アプリ側 ID/PW ゲート前提（静的JSON単体閲覧を禁じる） |
| **配信先** | https://tamakimotoyuki.github.io/tamaki-zaisan-dashboard/ |

## 入れてよいもの

- ✅ 月次試算表の集計値（大項目.json）= 4法人（医)明和会・MS・メディエンス・社福明和福祉会）の事業所別/全体PL・BS
- ✅ webapp の静的アセット（HTML・CSS・JS）
- ✅ 配色・グラフ設定（`【デザイン】`由来）

## 入れてはいけないもの

- ❌ 職員PII（氏名・住所・電話・個別給与・健診結果・マイナンバー）
- ❌ 患者PII（氏名・カルテ・診療内容・病歴）
- ❌ 認証情報（APIキー・トークン・パスワード・OTP・SSH秘密鍵）
- ❌ 個別役員報酬（職員PII相当）
- ❌ 田蒔個人の経営判断ロジック・税務戦略文書

## アプリ側ゲートの担保

公開しても閲覧不可になるよう、`index.html` または上位 hosting レイヤで以下を担保：
- ID/PW 認証（または Firebase Auth / Cloudflare Access 等）
- 認証なしで `data/*.json` を fetch できない構成

→ ゲートが外れている場合は **緊急対応＝即リポを Private に変更** or **データ削除**。

## 上流リポ（編集の本拠）

- **trial-balance-accounting**（Private）= 集計データ・スクリプト・運用ロジックの本拠
- このリポは **配信専用コピー**：上流の `webapp/data/大項目.json` と `webapp/assets/*` をミラーする
- 編集は上流で行い、本リポへ反映＝ push する流れ

## デプロイ手順

```powershell
# 上流リポでデータ更新後、本リポへ反映:
$src = "C:\Users\tamak\claude for desktop\【試算表・会計・決算】\webapp"
$dst = "C:\Users\tamak\AppData\Local\Temp\tamaki-zaisan-dashboard"
Copy-Item "$src\data\大項目.json" "$dst\data\大項目.json" -Force
# assets/ や index.html を変更した場合は同様にコピー
cd $dst
git add -A
git commit -m "deploy: <変更内容>"
git push
```

## 札の整合性チェック

- このリポへ何かを push する前に、対象ファイルが「入れてよいもの」に該当するか確認
- 「入れてはいけないもの」が混入していないか grep:
  - `git diff --cached | grep -E "(マイナンバー|個人番号|sk-ant|gho_|patient_id|カルテ番号)"`
- 不一致なら **push 中止して上流で除外**

## 過去のしくじり

- 2026-05-29: 上流（trial-balance-accounting）で社福月次データを大項目.jsonに統合後、本リポへの反映時に「Public リポへの会計データpush」を抽象分類で過剰警告した。本札を貼ることで以後は「会計データ・Public可」と即判定可能に。

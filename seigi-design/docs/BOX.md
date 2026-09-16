# Box 配布手順（せいぎデザイン）

配布物は `seigi-design/dist/` にあります。

- `seigi-starter-kit.zip` … まず配るもの
- `seigi-powerpoint-all.zip` … PPT一式
- `seigi-excel-all.zip` … Excel一式

## 推奨フォルダ構成（Box）

```
せいぎデザイン/
  01_starter/
    seigi-starter-kit.zip
  02_powerpoint/
    seigi-powerpoint-all.zip
  03_excel/
    seigi-excel-all.zip
  README.txt（任意）
```

## アップロード手順

1. Box にログイン
2. 上記フォルダを作成（初回のみ）
3. 同名ファイルがある場合は版を上げる／上書き方針を部署で統一
4. フォルダ共有リンクを生産技術室メンバーへ送る

## エージェント／CI から上げる場合

- Box MCP が認証済みなら、フォルダIDを指定してアップロード
- 未認証の環境では本手順の手動アップロードを使う

## 更新タイミング

トークンやテンプレを変えたら:

```bash
cd seigi-design
python3 scripts/generate_pptx.py
python3 scripts/generate_excel.py
python3 scripts/build_dist.py
```

その後、Box の3つのZIPを差し替える。

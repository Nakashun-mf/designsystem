# せいぎデザイン — 開発者向け

## 方針

- `tokens/seigi.json` が正
- PPT / Excel はスクリプト再生成
- 既存リポジトリ直下の soft-cloud パッケージは参照用。破壊的変更をしない
- 公式ロゴはコミットしない

## よく触る場所

| パス | 役割 |
|---|---|
| `tokens/seigi.json` | 色・組織名・フォント |
| `assets/mark/seigi-mark.svg` | 簡易マーク |
| `scripts/generate_pptx.py` | パワポ生成 |
| `scripts/generate_excel.py` | エクセル生成 |
| `scripts/build_dist.py` | 配布ZIP |

## 色を少し変えたいとき

1. `tokens/seigi.json` の HEX を変更
2. `python3 scripts/generate_pptx.py && python3 scripts/generate_excel.py && python3 scripts/build_dist.py`
3. Box の配布フォルダを更新

## Box 配布

1. `dist/*.zip` を生成
2. Box の「せいぎデザイン」フォルダへアップロード（同名は版を上げる）
3. 社内メンバーにはフォルダリンクを共有

Box MCP / CLI が使えない環境では、管理者が手動アップロードする。

## デモ

```bash
cd seigi-design
python3 -m http.server 4174
# http://localhost:4174/demos/
```

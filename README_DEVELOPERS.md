# 開発者向けガイド

## 方針

- `tokens/themes/*.json` を正とする
- CSS（`css/tokens.css`）と CSV / PPTX は JSON に追従
- 公式アセットはコミットしない

## よく触る場所

| パス | 役割 |
|---|---|
| `tokens/themes/` | テーマトークン |
| `css/tokens.css` | CSS 変数 |
| `css/components.css` | UI 部品 |
| `demos/index.html` | スタイルガイド |
| `scripts/generate_all_pptx.py` | パワポ一括生成 |

## パワポ追加手順

1. `scripts/generate_all_pptx.py` に `slide_xxx` を追加
2. `FULL_DECK_BUILDERS` か `packs` に登録
3. `python3 scripts/generate_all_pptx.py`
4. `ppt/POWERPOINT.md` を更新
5. `VERSION` を上げる

## ローカル確認

```bash
python3 -m http.server 4173
# http://localhost:4173/demos/
```

## Docker（任意）

```bash
docker run --rm -p 4173:80 \
  -v "$PWD:/usr/share/nginx/html:ro" \
  nginx:alpine
# http://localhost:4173/demos/
```

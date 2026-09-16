# まるい雲デザインシステム

しぐれういっぽさの主軸を **「丸くて雲みたいな形」** に置いた独立デザインシステムです（v1.3.0）。

参照の中心は [Wishing Umbrella MV](https://www.youtube.com/watch?v=qR2pggPX_Sc) のふわふわ世界観。
暖色ピーチ・青い傘・ぽむぽむ・やわらかいボケ光を先に定義しています。

> 公式の二次配布ではありません。ロゴ・KVは含みません。

## 主軸テーマ

**`soft-cloud`（まるい雲）** … デフォルト

| 要素 | 内容 |
|---|---|
| 形 | 重ね楕円の雲、大きな角丸、ピルボタン |
| 色 | ピンク `#f07a9a` / 空 `#d9efff` / 雲白 / ピーチ |
| フォント | Zen Maru Gothic / M PLUS Rounded 1c |

イベント配色（Wishing Umbrella / Uitopia / masterpiece / 5th）と **Uindows（OS風）** はサブテーマとして残しています。丸みは共通で寄せています。

## すぐ使う

```bash
python3 -m http.server 4173
# http://localhost:4173/demos/
```

パワポは `ppt/templates/soft-cloud-full.pptx` と `pack-soft-cloud-only.pptx` から。

```bash
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python3 scripts/generate_all_pptx.py
```

生成物は `ppt/templates/`（フルデッキ・用途パック・layouts 合計 40 ファイル）。

## ドキュメント

- [README_USERS.md](README_USERS.md) — ユーザー向け
- [README_DEVELOPERS.md](README_DEVELOPERS.md) — 開発者向け
- [docs/DESIGN.md](docs/DESIGN.md)
- [ppt/POWERPOINT.md](ppt/POWERPOINT.md)

## バージョン

- **1.3.0** — Uindows / OS UI・パワポ一式を含む統合パッケージ
- **1.2.0** — Wishing Umbrella MV のふわふわ世界観に寄せ直し
- **1.1.0** — Soft Cloud を主軸に再設計。雲シェイプの Web / PPT を強化
- **1.0.0** — イベントサイト色抽出の初版

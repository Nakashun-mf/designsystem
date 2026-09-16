# まるい雲デザインシステム ＋ せいぎデザイン

このリポジトリには2つのレイヤーがあります。

| パス | 内容 |
|---|---|
| リポジトリ直下 | しぐれうい系 soft-cloud デザインシステム（v1.3.0・参照用） |
| [`seigi-design/`](seigi-design/) | **大成工業 生産技術室向け「せいぎデザイン」**（v1.0.0・本番利用） |

部署の日常資料は **せいぎデザイン** を使ってください。

## せいぎデザイン（おすすめ）

やわらかさを残しつつ、大成ブルー `#00315E` とノーリツレッド `#E8380D` をアクセントにした社内テンプレです。

```bash
cd seigi-design
pip install -r requirements.txt
python3 scripts/generate_pptx.py
python3 scripts/generate_excel.py
python3 scripts/build_dist.py
```

配布ZIP: `seigi-design/dist/`（starter / powerpoint-all / excel-all）  
詳細: [seigi-design/README.md](seigi-design/README.md)  
仕様: [docs/superpowers/specs/2026-09-16-seigi-design-design.md](docs/superpowers/specs/2026-09-16-seigi-design-design.md)

最新の社内配布は **Box**（「せいぎデザイン」フォルダ）を参照。アップロード手順は [seigi-design/docs/BOX.md](seigi-design/docs/BOX.md)。

## soft-cloud（参照）

```bash
python3 -m http.server 4173
# http://localhost:4173/demos/
```

パワポ再生成: `pip install -r requirements.txt && python3 scripts/generate_all_pptx.py`

# せいぎデザイン

大成工業株式会社 **生産技術室**向けの部署デザインシステムです（v1.0.0）。

しぐれうい系 soft-cloud の「やわらかさ・親しみやすさ」を残しつつ、**大成ブルー**と**ノーリツレッド**をアクセントにした社内用フォークです。  
公式ロゴは使わず、雲＋歯車の簡易マークを使います。

> 本パッケージは社内テンプレート用途です。公式ブランドの二次配布ではありません。

## すぐ使う

1. `dist/seigi-starter-kit.zip` を解凍
2. PowerPoint は `seigi-starter.pptx` または `seigi-basic.pptx`
3. ジグDRは `seigi-jig-dr.pptx` / `seigi-jig-dr.xlsx`
4. 文言を差し替えるだけ（色は触らなくてよい）

最新配布は **Box** の「せいぎデザイン」フォルダを参照してください。

## 再生成

```bash
cd seigi-design
python3 -m pip install -r requirements.txt
python3 scripts/generate_pptx.py
python3 scripts/generate_excel.py
python3 scripts/build_dist.py
```

## ドキュメント

- [README_USERS.md](README_USERS.md)
- [README_DEVELOPERS.md](README_DEVELOPERS.md)
- 仕様: `../docs/superpowers/specs/2026-09-16-seigi-design-design.md`

# せいぎデザイン Implementation Plan

> **For agentic workers:** 仕様承認済み。ユーザー指示により実装〜配布まで一気に実行。

**Goal:** 同一リポジトリ内に `seigi-design/` を追加し、PPT/Excel を生成して Box に配布する。

**Architecture:** soft-cloud は維持。`seigi-design/` にトークン・マーク・生成スクリプト・成果物を分離。

**Tech Stack:** Python 3, python-pptx, openpyxl, Box（MCP または手動）

## Global Constraints

- 公式ロゴ不使用（雲＋歯車マーク）
- 大成ブルー `#00315E` / `#3477B6`、ノーリツレッド `#E8380D`
- 名称: せいぎデザイン／生産技術室／大成工業
- 既存 soft-cloud はほぼ変更しない

---

## Tasks

- [x] Task 1: 骨格・トークン・マーク・README
- [x] Task 2: PPT 生成
- [x] Task 3: Excel 生成
- [x] Task 4: dist ZIP・デモ・ルート案内
- [ ] Task 5: Box アップロード（認証不可時は手順明記＋成果物配置）

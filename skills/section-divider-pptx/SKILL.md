---
name: section-divider-pptx
description: >
  中扉（Section Divider）のPowerPointスライドを生成するスキル。
  企業調査レポート・BDD・M&A評価・新規参入調査などのデッキ内で、
  各セクションの開始時に配置するセクション区切りスライド。
  左半分にアクセントカラー背景＋巨大なセクション番号（180pt白）、
  右半分にセクションタイトル・サブタイトル・トピックリストを配置する。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「中扉」「セクション区切り」「Section Divider」「インターページ」という言葉が出た場合
  - 「セクションごとの区切りページ」「タイトルスライド」「セクション開始ページ」という要望
  - 企業調査デッキの各セクション冒頭の中扉作成
  - 「目次のセクションごとに区切りスライドを入れたい」という要望
supported_brands: [stellar_aiz, roleup]

---

# 中扉（Section Divider）PowerPoint ジェネレーター

レポートデッキの各セクション開始時に配置する中扉スライド。
**読み手にセクション転換を視覚的に伝える**ことで、デッキの論理構造が明確になる。

---

## スライド構成

```
┌──────────────┬───────────────────────────────────────┐
│              │                                       │
│  SECTION     │   マクロ・市場環境分析                │
│              │   ━━━ (アクセントライン)              │
│   03         │   業界を取り巻くマクロ環境と...       │
│              │                                       │
│   (大番号)   │   ▍ このセクションで扱う内容           │
│              │   ▸ PEST分析                          │
│              │   ▸ 市場規模・成長率                  │
│              │   ▸ 市場シェア構造                    │
│              │   ▸ ポジショニングマップ              │
└──────────────┴───────────────────────────────────────┘
```

### 左半分（アクセント）
- **背景色**: アクセントカラー（セクション番号に応じて自動）
- **"SECTION" ラベル** (Bold 20pt 白、Arial)
- **巨大なセクション番号** (Bold 180pt 白、Arial、"01"〜"99")

### 右半分（コンテンツ）
- **タイトル** (Bold 36pt 黒)
- **アクセントライン** (1.0in × 0.06in、アクセントカラー)
- **サブタイトル** (18pt サブテキスト色)
- **トピックリスト** (▸マーカー、13pt、アクセント色マーカー)

### セクション色のローテーション（TOCと統一）

| セクション | 色 |
|---|---|
| 01 | 紺 #2E4A6B |
| 02 | 紫 #7B4FB0 |
| 03 | 青 #2E6FBF |
| 04 | 緑 #3D8F5A |
| 05 | オレンジ #DA7A2D |
| 06 | 赤 #C03A3A |
| 07 | グレー #595959 |

`color` フィールドで個別指定も可能。

---

## JSONデータ仕様

```json
{
  "section_number": 3,
  "title": "マクロ・市場環境分析",
  "subtitle": "業界を取り巻くマクロ環境と市場構造を俯瞰する",
  "topics": [
    "PEST分析（政治・経済・社会・技術）による業界外部環境の整理",
    "市場規模・成長率の推移と将来予測",
    "市場シェア構造と上位プレイヤーの動向",
    "ポジショニングマップによる競合の戦略グループ分析"
  ]
}
```

### JSONフィールド仕様

| フィールド | 型 | 必須 | 説明 |
|---|---|---|---|
| `section_number` | int | 必須 | セクション番号（1〜99） |
| `title` | string | 必須 | セクションタイトル |
| `subtitle` | string | 任意 | サブタイトル（補足説明） |
| `topics` | array (string) | 任意 | このセクションで扱うトピック（3〜5個推奨） |
| `color` | string | 任意 | アクセント色（#RRGGBB）。未指定時は section_number から自動 |

---

## スクリプト実行コマンド

```bash
pip install python-pptx -q --break-system-packages

python <SKILL_DIR>/scripts/fill_section_divider.py \
  --data {{WORK_DIR}}/section_divider_data.json \
  --brand stellar_aiz \
  --output {{OUTPUT_DIR}}/SectionDivider_output.pptx
```

`--brand` は `stellar_aiz`（既定、16:9 / Meiryo UI / 7 色ローテ）か `roleup`
（A4 横 / Yu Gothic UI / theme.chart_palette 8 色ローテ）を指定。
`--template` を省略すると brand に応じたテンプレートが自動解決される。

---

## 品質チェックリスト

- [ ] section_numberが目次と整合しているか
- [ ] タイトルが目次のセクション名と一致しているか
- [ ] サブタイトルがそのセクションの目的を簡潔に表しているか
- [ ] トピックがそのセクション内のスライドと対応しているか
- [ ] 色が目次（table-of-contents-pptx）と統一されているか

---

## アセット / スクリプト / 参考

| ファイル | 用途 |
|---|---|
| `assets/stellar_aiz/section-divider-template.pptx` | スライドテンプレート (stellar_aiz, 16:9) |
| `assets/stellar_aiz/layout.json` | stellar_aiz レイアウト座標 (V1 hardcode 値ミラー) |
| `assets/roleup/section-divider-template.pptx` | スライドテンプレート (roleup, A4 横) |
| `assets/roleup/layout.json` | roleup レイアウト座標 (slide_h 8.27 in に合わせ下方シフト) |
| `scripts/fill_section_divider.py` | 生成スクリプト (Pattern A: brand-aware) |
| `references/sample_data.json` | サンプル（Section 03 マクロ・市場環境分析） |

---

## 注意事項

- **テンプレートのプレースホルダーは削除される**: 中扉は専用デザインのため、テンプレートの「Title 1」「Text Placeholder 2」は自動削除して一から組む
- **目次と色を統一**: `table-of-contents-pptx` と同じセクション番号で同じ色になるため、両スキルを使うとデッキ全体の一貫性が保たれる
- **トピック数**: 3〜5個が読みやすい

---

## 企業調査レポートでの使用例

### 配置パターン（拡張版デッキ）

```
1. (表紙)
2. エグゼクティブサマリー
3. 目次
4. 中扉 「Section 01: 対象会社の現状」 ← 本スキル
5. 対象会社プロファイル
6. 事業ポートフォリオ
7. SWOT
8. 中扉 「Section 02: マクロ・市場環境」 ← 本スキル
9. PEST分析
10. 市場規模
...
```

### 配置パターン（基本版デッキ）

スライド数が少ないデッキでは中扉は不要。
標準版以上（10枚超）で初めて中扉の効果が出る。

### 中扉と目次のセットで使う

`table-of-contents-pptx` と `section-divider-pptx` は**ペアで使う**ことを推奨。
目次の各セクションと、各セクション開始時の中扉が**同じセクション番号・同じ色**で
連動することで、読み手の認知負荷を下げる。

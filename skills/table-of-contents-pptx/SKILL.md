---
name: table-of-contents-pptx
description: >
  目次（Table of Contents）のPowerPointスライドを生成するスキル。
  企業調査レポート・BDD・M&A評価・新規参入調査などのデッキ冒頭または
  エグゼクティブサマリー直後に配置する目次スライド。
  3〜7個のセクションを色付き番号バッジ + Boldタイトル + サブ項目リスト + ページ番号で表示する。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「目次」「Table of Contents」「TOC」「アジェンダ」「Agenda」という言葉が出た場合
  - 「セクション一覧」「章立て」「構成スライド」という要望
  - 企業調査レポート・デッキの冒頭目次スライド作成
  - 「ページ番号付きの目次」「サブ項目付き目次」を求められた場合
supported_brands: [stellar_aiz]

---

# 目次（Table of Contents）PowerPoint ジェネレーター

レポートデッキの冒頭に配置する目次スライド。読み手にデッキ全体の構成を俯瞰させ、
**論理的な流れと配置を一目で示す**。

---

## スライド構成

| セクション | 位置 | 内容 |
|---|---|---|
| **タイトル** | 最上部 | 「目次」「Table of Contents」 |
| **セクション一覧** | 中央 | 3〜7個のセクション行 |

### 各セクション行の構成

```
┌──────┐                                                     P. 5
│  02  │  対象会社の現状分析
└──────┘  ▸ 会社プロファイル
          ▸ 事業ポートフォリオ
          ▸ SWOT分析
          ─────────────────────────── (区切り線)
```

- **番号バッジ** (左、四角): セクション色、Arial Bold 28pt白
- **タイトル** (Bold 18pt、セクション色)
- **サブ項目** (11pt、▸マーカー、サブテキスト色)
- **ページ番号** (右端、Bold 11pt、セクション色)
- **区切り線**: 各セクション間（薄グレー）

### セクション色のローテーション

|セクション | 色 |
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
  "main_message": "目次",
  "chart_title": "Table of Contents",
  "sections": [
    {
      "title": "エグゼクティブサマリー",
      "page": "3",
      "subitems": ["調査全体のKey Findings", "推奨アクションの概要"]
    },
    {
      "title": "対象会社の現状分析",
      "page": "5",
      "subitems": ["会社プロファイル", "事業ポートフォリオ", "SWOT分析"]
    },
    {
      "title": "マクロ・市場環境分析",
      "page": "9",
      "subitems": ["PEST分析", "市場規模・シェア", "ポジショニング"]
    }
  ]
}
```

### JSONフィールド仕様

| フィールド | 型 | 必須 | 説明 |
|---|---|---|---|
| `main_message` | string | 任意 | タイトル（デフォルト「目次」） |
| `chart_title` | string | 任意 | サブタイトル（デフォルト「Table of Contents」） |
| `sections` | array | 必須 | セクション配列（3〜7個推奨、最大7） |
| `sections[].title` | string | 必須 | セクション名 |
| `sections[].page` | string | 任意 | ページ番号 |
| `sections[].subitems` | array (string) | 任意 | サブ項目（2〜4個推奨） |
| `sections[].color` | string | 任意 | セクション色（#RRGGBB）。未指定時は自動 |

---

## スクリプト実行コマンド

```bash
pip install python-pptx -q --break-system-packages

python <SKILL_DIR>/scripts/fill_table_of_contents.py \
  --data {{WORK_DIR}}/toc_data.json \
  --template <SKILL_DIR>/assets/table-of-contents-template.pptx \
  --output {{OUTPUT_DIR}}/TableOfContents_output.pptx
```

---

## 品質チェックリスト

- [ ] セクション数が3〜7個の適切な範囲
- [ ] 各セクションのページ番号が実際のデッキと整合しているか
- [ ] サブ項目がそのセクションの実際のスライドと整合しているか
- [ ] セクション色が他のスライド（中扉等）と統一感があるか

---

## アセット / スクリプト / 参考

| ファイル | 用途 |
|---|---|
| `assets/table-of-contents-template.pptx` | スライドテンプレート |
| `scripts/fill_table_of_contents.py` | 生成スクリプト |
| `references/sample_data.json` | サンプル（6セクション構成） |

---

## 注意事項

- **デッキ完成後にページ番号を最終調整**: 目次は最初に作っても、最終的な並び替え後にページ番号を更新
- **サブ項目は省略可**: シンプルにセクションタイトルのみでもOK
- **`section-divider-pptx` と色を統一**: 同じセクション番号で同じ色を使うことで、デッキ全体の一貫性が出る

---

## 企業調査レポートでの使用例

### 配置位置（推奨）

```
1. (表紙)
2. エグゼクティブサマリー
3. 目次 ← 本スキル
4. (中扉) 対象会社の現状分析
5. 対象会社プロファイル
...
```

エグゼクティブサマリー直後に配置することで、読み手が
「結論を見た後で全体構成を把握する」流れになる。

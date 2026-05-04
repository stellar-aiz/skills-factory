# Rollup Standard Format (vF 20250928) 正本仕様書

**版**: vF 20250928 / **確定日**: 2026-05-04 / **出典**: `work/rollup_official_templates/standard_format_vF_20250928.pptx`(slide3「PPT作成時の留意点」)

本ドキュメントは brand `rollup` の納品標準フォーマットの**正本**である。`theme.json` の数値はすべて本書由来であり、矛盾が生じた場合は本書が優先する。

---

## 1. スライド基本仕様

| 項目 | 値 | 出所 |
|---|---|---|
| スライドサイズ | A4 横 = 11.69 × 8.27 inch (29.70 × 21.00 cm) | 公式テンプレ `presentation.xml/sldSz` |
| アスペクト | A4-landscape | — |
| フォント (latin) | Yu Gothic UI | 公式テンプレ `theme1.xml/fontScheme/majorFont` |
| フォント (ea) | Yu Gothic UI | 公式テンプレ `theme1.xml/fontScheme/minorFont` |
| フォントフォールバック (ea) | Yu Gothic / Hiragino Sans / MS PGothic | brand_resolver 既存値 |

## 2. カラーパレット (theme1.xml/clrScheme)

| 用途 | hex | 役割 |
|---|---|---|
| 基本テキスト色 | `#241A17` | 全本文・タイトル・キーメッセージ |
| 出所/注記 | `#3E3A39` | 6pt 出所、12pt 注記 |
| ヘッダー背景 | `#F5EFE5` | テーブルヘッダー行・ラベルバー背景 |
| ハイライト (対象会社) | `#C78624` | テーブルセル強調・棒グラフ強調 |
| ハイライト (その他) | `#CDCECE` | 比較対象の他社、グレー塗り |
| ラベルバー前景 | `#7C4C2C` | object 8 等のセクション区切り線 |
| ラベル背景 | `#F2E8DD` | サブ項目の薄塗り背景 |
| 売上棒 (accent_revenue_bar) | `#7C4C2C` | チャート売上系列 |
| 利益率折れ線 (accent_op_margin_line) | `#604C3F` | チャート利益率系列 |
| CAGR 矢印 | `#241A17` | 注記矢印 |

**chart_palette** (連続使用時の系列色順):
1. `#7C4C2C`(accent1) / 2. `#897141`(accent2) / 3. `#604C3F`(accent3) / 4. `#C78624`(dk2) / 5. `#AF7026`(lt2) / 6. `#3E3A39`(accent5) / 7. `#9C755F` / 8. `#CDCECE`(accent4)

## 3. フォントサイズ (slide_type 別)

| Shape 種別 | サイズ | 用途 |
|---|---|---|
| Title (タイトル 3) | **22pt** | スライドタイトル |
| Subtitle (テキスト プレースホルダー 5) | **12pt** | サブタイトル / 各パネルラベル |
| Key Message (テキスト プレースホルダー 4) | **14pt** | スライド最上部のメインメッセージ |
| **本文・テーブル (通常)** | **10pt** | 全 skill デフォルト |
| **本文・テーブル (executive-summary-pptx のみ)** | **12pt** | F-3 判定: skill 名で自動切替 |
| 出所 / 注記 (テキスト プレースホルダー 3) | **6pt** | 下端の出所ライン |
| CAGR ラベル | **16pt** | チャート上の CAGR 注記 |
| チャートタイトル | **11pt** | (用途限定) |
| データラベル | **11pt** | (用途限定) |
| 凡例 | **12pt** | (用途限定) |
| 軸ラベル | **11pt** | (用途限定) |

### F-3 判定ロジック (executive_summary 12pt 切替)

`fill_*.py` の本文・テーブル描画ロジックは、**実行中の skill 名**(またはオーケストレーターから渡される `skill_id`)で判定する:

```python
# brand_resolver.py 拡張 API
def font_size_body_pt(self, skill_id: str) -> int:
    if self.id == "rollup" and skill_id == "executive-summary-pptx":
        return 12
    return self.defaults.get("font_size_body_pt", 10)
```

JSON input への `slide_type` field 追加は不要(F-3 = (b) 採用)。

## 4. 行高 (Line Spacing)

**規定**: 本文・テーブルセル内の段落で `<a:lnSpc><a:spcPts val="1200"/></a:lnSpc>` を必ず指定する(F-4 = (b) 採用)。

- OOXML 表現: 段落 `<a:pPr>` 配下の `<a:lnSpc><a:spcPts val="1200"/></a:lnSpc>`
- 12pt = `val="1200"`(spcPts 単位は 1/100 pt)
- 適用対象: 本文 10pt 段落・テーブルセル段落
- 適用しない: タイトル 22pt / キーメッセージ 14pt / サブタイトル 12pt(これらは shape 高さに依存して可変)

**テーブル `<a:tr h="...">` (cell row height) は強制しない**。公式テンプレ slide2 のテーブルは見出し 42pt / データ行 36-62pt と内容に応じて拡張可となっており、本書も同方針(行送り 12pt のみ規定)。

## 5. 数値書式

### 推奨 Excel ユーザー定義書式

```
_ * #,##0_ ;_ * (#,##0)_ ;_ "-"_ ;_ @_ 
```

### Python 側変換ルール (`format_helpers.py`)

```python
def format_cell_value(value, theme):
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)):
        if value == 0:
            return theme.zero_text()              # "-"
        if value < 0:
            if theme.negative_format() == "paren":
                return f"({abs(value):,.0f})"     # rollup default
            return f"△{abs(value):,.0f}"          # 許容代替
        return f"{value:,.0f}"
    return str(value)
```

| 規定 | 値 | 備考 |
|---|---|---|
| ゼロ表記 | `-`(ハイフン) | `0` 文字列禁止 |
| マイナス表記デフォルト | `(XXXX)` 括弧 | `theme.negative_format() = "paren"` |
| マイナス表記許容代替 | `△XXXX` | JSON で `"negative_format": "triangle"` 指定時 |
| マイナス表記禁止 | `-XXXX` 直書き / 赤字色 | hard-fail 対象 |
| 千位区切り | `#,##0` | 全数値で必須 |

## 6. 会計期間表記

**規定**: `YY/MM期` 形式(例: `19/10期`, `20/10期`, `21/10期`)

- 西暦 4 桁(`2019/10期`)は **禁止**
- `YY` は和暦ではなく西暦下 2 桁
- `MM` は決算月(2 桁)
- `期` の後ろに半角スペースなし

`format_helpers.py`:

```python
def format_fiscal_period(year: int, month: int) -> str:
    return f"{year % 100:02d}/{month:02d}期"
```

## 7. 左端揃えガイド

**規定**: コンテンツ shape の x 座標 = **0.41 inch**(全 layout 共通)

- 公式テンプレ slide3-13 の content shape (`正方形/長方形 1`) はすべて x=0.41 inch 起点
- ガイド逸脱は `check_brand_compliance.py` の警告対象(±0.02 inch 許容)
- subtitle (`テキスト プレースホルダー 5`) は x=0.42 inch だが、これも 0.41 inch ガイドの誤差範囲とみなす

**1 スライドに 2 つ以上の表がある場合**: 横幅を揃える(slide3 の留意点より)。`layout.json` で `uniform_table_width_in` を必須指定。

## 8. 出所 (Source) 必須化

**規定**: **全 brand で hard-fail**(F-8 = (b) 採用)

```python
# fill_*.py 共通
if not data.get("source"):
    raise ValueError(f"source is required (brand={brand.id})")
```

- 既存 stella 挙動からの **挙動変更**: stella も warning → hard-fail に強化
- 影響範囲: customer-profile / company-history / market-environment の sample_data.json で `source` field の存在確認が必要(Phase 4 で実施)
- 出所表示位置: 下端の `テキスト プレースホルダー 3` (x=0.41, y=7.57, 6pt)

## 9. テーブル運用ガイドライン

slide3 留意点より:

1. **表は Excel で作成し、リンク貼り付けで PPT に貼り付け** — V2 では運用 caveat とし、**ネイティブテーブル + 数値書式遵守で妥協可**(ユーザー判断 2026-05-04)
2. **表のサイズ調整は Excel 上で行う** — V2 適用外(ネイティブテーブル想定)
3. **表は極力ガイド左端 (x=0.41 inch) に合わせる**
4. **1 スライドに 2 つ以上の表を挿入する場合、横幅を揃える** — `layout.json/uniform_table_width_in` で実現

各 SKILL.md の運用 caveat に「最終納品時に Excel リンク貼り付けに置換可」を追記する(Phase 4)。

## 10. 公式テンプレートの 11 レイアウト一覧

`work/rollup_official_templates/standard_format_vF_20250928.pptx` 内のレイアウト見本(Phase 2 で各 skill にどのレイアウトを割り当てるかを決定):

| slide | レイアウト名 | 構造 | 想定 skill 用途(暫定) |
|---|---|---|---|
| 1 | 表紙 | ROLEUP for Succession | (cover-pptx 系、対象外) |
| 2 | カラーパレット見本 | 標準カラー一覧 | (参照用、対象外) |
| 3 | レイアウト① | フルワイド本文(留意点記載 layout) | market-environment? |
| 4 | レイアウト② | 左右 2 分割 (5.23 + 5.23 inch) | customer-profile, company-history |
| 5 | レイアウト③ | 4 分割 (2×2 グリッド) | (KPI dashboard 系) |
| 6 | レイアウト④ | 4 分割 (上 2 + 下 2、上下サイズ違い) | (KPI dashboard 系) |
| 7 | レイアウト⑤ | 4 分割 (バリエーション) | (KPI dashboard 系) |
| 8 | レイアウト⑥ | 左右 2 + 左下 1 (3 セクション) | (応用) |
| 9 | レイアウト⑦ | 左大 (6.47 inch) + 右小 (4.28 inch) | revenue-analysis 系 |
| 10 | レイアウト⑧ | 左小 (4.70 inch) + 右大 (6.06 inch) | (応用) |
| 11 | レイアウト⑨ | 上下 2 段 (フルワイド) | (応用) |
| 12 | レイアウト⑩ | 上小 + 下大 | (応用) |
| 13 | レイアウト⑪ | 上大 + 下小 | (応用) |
| 14 | 目次 | TOC | (table-of-contents-pptx 対象外) |
| 15 | 中扉 | section divider | (section-divider-pptx 対象外) |

**Phase 2 確定タスク**: Pilot 3 (customer-profile / company-history / market-environment) に slide4 / slide11 / slide3 のいずれかを割り当て、`assets/rollup/<skill>-template.pptx` として抽出する。詳細は Phase 2 で決定。

## 11. theme.json schema 2.0 への反映

本書から `skills/_common/brands/rollup/theme.json` に追加する key 一覧(Phase 1 で実施):

```jsonc
{
  "$schema_version": "2.0",
  "defaults": {
    "font_size_title_pt": 22,
    "font_size_subtitle_pt": 12,
    "font_size_key_message_pt": 14,
    "font_size_body_pt": 10,
    "font_size_executive_summary_body_pt": 12,
    "font_size_source_pt": 6,
    "font_size_cagr_pt": 16,
    "font_size_chart_title_pt": 11,
    "font_size_data_label_pt": 11,
    "font_size_legend_pt": 12,
    "font_size_axis_pt": 11,
    "line_height_pt": 12,
    "number_format_excel": "_ * #,##0_ ;_ * (#,##0)_ ;_ \"-\"_ ;_ @_ ",
    "negative_format": "paren",
    "zero_text": "-",
    "fiscal_period_format": "YY/MM期"
  },
  "layout_rules": {
    "left_align_guide_x_in": 0.41,
    "uniform_table_width": true,
    "source_required": true
  },
  "executive_summary_skill_ids": ["executive-summary-pptx"]
}
```

stella 側 `theme.json` も schema 2.0 互換化が必要(後方互換 default で既存挙動を担保、Phase 1 で実施)。

## 12. 確認履歴

| 日付 | 確認者 | 確認事項 |
|---|---|---|
| 2026-05-04 | ユーザー | F-3 = (b) skill 名で自動判定、F-4 = (b) lnSpc/spcPts、F-8 = (b) 全 brand hard-fail |
| 2026-05-04 | ユーザー | F-1 (公式テンプレ本数) は Phase 2 で決定 |
| 2026-05-04 | ユーザー | F-2 / F-5 / F-6 / F-7 は公式テンプレからの機械抽出値で確定 |

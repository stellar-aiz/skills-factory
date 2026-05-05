---
name: market-kbf-pptx
description: >
  対象市場における Key Business Factor（KBF・キービジネスファクター）3つを
  「KBF / 詳細 / 主要プレイヤーのKBF抑え方の例」の3列テーブルで1枚のスライドに整理する
  PowerPoint生成スキル。各KBFには主要プレイヤー（最大5社）が実際にそのKBFをどう実装しているかの
  具体例を併記し、「市場で勝つために必要な要素」を読み手に直感的に示す。
  KBFは3つ固定、テーブルは PPTXネイティブ（編集可能）。
  market-overview-agent オーケストレーターから呼び出されることを主用途とするが、単独起動も可能。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「KBF」「キービジネスファクター」「Key Business Factor」「KSF」「Key Success Factor」という言葉が出た場合
  - 「市場で勝つために必要な要素」「成功要因」「市場の勝ち筋」という要望
  - market-overview-agent / strategy-report-agent オーケストレーターから呼び出された場合
  - 「KBFをスライドにして」「成功要因をパワポで整理」という要望
supported_brands: [stellar_aiz]

---

# 市場KBF（Key Business Factor）PowerPoint ジェネレーター

対象市場で勝つために必要な KBF を3つに絞り、各KBFがなぜ重要か（詳細）と、主要プレイヤーが
実際にどう抑えているか（例）を併記した 3列テーブルのスライドを1枚生成するスキル。

**方式: PPTXネイティブテーブル**
- 全オブジェクトがPPT上で編集可能（画像化しない）
- python-pptx の `add_table()` で動的にテーブルを構築

---

## スライド構成

| セクション | 位置 | 内容 |
|---|---|---|
| **メインメッセージ** | 最上部 | **最大65文字（hard-fail）**。「本市場で勝つには〜が鍵」のように KBF を要約 |
| **チャートタイトル** | メインメッセージ直下 | 「{市場名}：勝ち筋を決める3つのKey Business Factor」等 |
| **KBFテーブル** | コンテンツエリア | 3行×3列のネイティブテーブル（KBF×3） |
| **出典** | 左下 | 情報ソースの一括記載 |

### KBFテーブルの仕様

- **列構成（3列固定）**:
  - 列1: **KBF**（幅20%、紺色背景＋白文字 / KBF名を大きめのフォントで強調）
  - 列2: **詳細**（幅30%、白背景 / なぜこのKBFが重要か 80〜120字）
  - 列3: **プレイヤーの抑え方の例**（幅50%、薄グレー背景 / 主要プレイヤー最大5社の具体例）

- **行構成（4行固定）**:
  - ヘッダー行: 「KBF」「詳細」「プレイヤーの抑え方の例」（高さ0.40"）
  - データ行: KBF×3（残りの高さを等分、約1.50"/行）

- **プレイヤー例セルの書式**:
  - `{プレイヤー名}: {具体例}` を改行区切りで列挙
  - プレイヤー名のみ太字、コロン以降は通常体
  - 1〜5社（推奨3〜5社）

---

## 入力パターンと処理フロー

### パターンA：生情報（業界レポート・コンサルレポート・記事等）が入力された場合

**いきなりPowerPointを作成しない**

1. **Step 1: 市場の特性から KBF 候補を3つ抽出**（成功条件・競争優位の源泉・参入障壁の観点で網羅性チェック）
2. **Step 2: 各KBFについて主要プレイヤーが実装している具体例を整理**
3. **Step 3: Markdownでユーザーに提示**し、確認・修正を求める
4. **Step 4: ユーザーの承認後**、PowerPointを生成する

### パターンB：データが整理済み（JSON）で入力された場合

1. **Step 1: 内容を確認**し、KBF が3つあるか、各KBFに player_examples が1〜5件あるかをバリデーション
2. **Step 2: 確認後、PowerPointを生成する**

### パターンC：market-overview-agent / strategy-report-agent から呼ばれる場合

1. 渡された JSON データを使用し、直接 PowerPoint を生成する

---

## Step 3: Markdownでの確認出力フォーマット（パターンAの場合）

```markdown
## KBF整理結果

**Main Message（※ドラフト、最大70文字）**
本市場で勝つには「データ蓄積」「日本企業へのカスタマイズ力」「PMI実行力」の3つのKBFが鍵となる

**Chart Title**
国内HR Tech市場：勝ち筋を決める3つのKey Business Factor

### KBFテーブル

| # | KBF | 詳細 | プレイヤーの抑え方の例 |
|---|-----|------|------------------------|
| 1 | データ蓄積 | 推薦精度・スコアリング精度はデータ量に強く依存。先行プレイヤーが累積優位を持つ。 | A社: 累計1,000万件の応募データを保有 / B社: 大手事業会社との独占提携でデータ吸い上げ |
| 2 | ... | ... | ... |
| 3 | ... | ... | ... |
```

確認メッセージ例：
> 上記のKBF整理でよろしいでしょうか？修正があればお知らせください。確認後にPowerPointを生成します。

---

## Step 4: PowerPointの生成

### JSONデータ仕様

`{{WORK_DIR}}/market_kbf_data.json` に以下の形式で保存する：

```json
{
  "main_message": "本市場で勝つには「データ蓄積」「日本企業へのカスタマイズ力」「PMI実行力」の3つのKBFが鍵となる",
  "chart_title": "国内HR Tech市場：勝ち筋を決める3つのKey Business Factor",
  "source": "出典：矢野経済研究所「HR Tech市場2025」、各社IR、コンサルレポート",
  "kbf_list": [
    {
      "name": "データ蓄積",
      "description": "推薦精度・スコアリング精度はデータ量に強く依存。先行プレイヤーが累積優位を持ち、後発が逆転しにくい構造。",
      "player_examples": [
        { "player": "A社", "example": "累計1,000万件の応募データを保有、AI推薦エンジンに活用" },
        { "player": "B社", "example": "大手事業会社との独占提携でデータを吸い上げ" },
        { "player": "C社", "example": "業界特化（製造業）で深いデータ層を構築" }
      ]
    },
    {
      "name": "日本企業へのカスタマイズ力",
      "description": "海外SaaSの直輸入では日本企業の組織慣行に合わず、導入後の定着が困難。日本仕様の業務フロー対応が成否を分ける。",
      "player_examples": [
        { "player": "A社", "example": "100名規模の国内エンジニアが顧客個別カスタマイズに対応" },
        { "player": "D社", "example": "テンプレ100種類超を予め用意" }
      ]
    },
    {
      "name": "PMI実行力",
      "description": "M&Aによる機能・顧客拡張がスケールの主要ドライバー。買収後の統合スピードと顧客クロスセル力が成長を左右する。",
      "player_examples": [
        { "player": "A社", "example": "過去3年で5社買収、いずれも12ヶ月以内に統合完了" },
        { "player": "E社", "example": "M&A専担チーム30名を抱え、年2-3件のディールを実行" }
      ]
    }
  ]
}
```

### JSONフィールド仕様

| フィールド | 型 | 必須 | 説明 |
|---|---|---|---|
| `main_message` | string | 必須 | メインメッセージ。**最大65文字（hard-fail）**、3つのKBFを要約 |
| `chart_title` | string | 任意 | チャートタイトル。デフォルト「Key Business Factor」 |
| `source` | string | 任意 | 出典（左下に一括表示） |
| `kbf_list` | array | 必須 | **長さ3固定（hard-fail）**。3つのKBFを配列で記述 |
| `kbf_list[].name` | string | 必須 | KBF名（**最大15文字、hard-fail**） |
| `kbf_list[].description` | string | 必須 | KBFの詳細（**最大120文字、hard-fail**、推奨80〜120字）|
| `kbf_list[].player_examples` | array | 必須 | 主要プレイヤーの抑え方の例。**1〜5要素（hard-fail）**、推奨3〜5 |
| `kbf_list[].player_examples[].player` | string | 必須 | プレイヤー名 |
| `kbf_list[].player_examples[].example` | string | 必須 | 具体的な抑え方の例（**最大80文字、hard-fail**、推奨40〜80字）|

### バリデーション

- `kbf_list.length === 3` を強制（KBFは3つ固定の方針）
- 各 KBF の `player_examples.length` は 1〜5（5社上限）

---

## スクリプト実行コマンド

```bash
pip install python-pptx lxml -q --break-system-packages

python {{SKILL_DIR}}/scripts/fill_market_kbf.py \
  --data {{WORK_DIR}}/market_kbf_data.json \
  --template {{SKILL_DIR}}/assets/market-kbf-template.pptx \
  --output {{OUTPUT_DIR}}/MarketKBF_output.pptx
```

### 出力確認

```bash
python -m markitdown {{OUTPUT_DIR}}/MarketKBF_output.pptx
```

---

## デザイン仕様

### フォント

| 要素 | フォント |
|---|---|
| 日本語 | Meiryo UI |
| ラテン文字 | Arial |

### フォントサイズ

| 要素 | サイズ |
|---|---|
| ヘッダー行 | 12pt Bold |
| KBF名（列1） | 14pt Bold（白文字） |
| 詳細（列2） | 11pt |
| プレイヤー例（列3） | 10pt（プレイヤー名のみ Bold） |

### 色

| 要素 | カラーコード |
|---|---|
| テキスト（基本） | #333333 |
| KBF列背景（強調） | #1F3864（紺）|
| KBF列文字 | #FFFFFF（白）|
| ヘッダー行背景 | #F0F0F0（グレー）|
| 詳細列背景 | #FFFFFF |
| プレイヤー例列背景 | #FAFAFA |

### レイアウト定数

| 要素 | 値 |
|---|---|
| スライドサイズ | 13.33" × 7.50"（ワイドスクリーン） |
| コンテンツ左端 | 0.41" |
| コンテンツ上端 | 1.50" |
| コンテンツ幅 | 12.52" |
| コンテンツ下端 | 6.90"（出典の上） |
| 列1（KBF）幅比 | 20% |
| 列2（詳細）幅比 | 30% |
| 列3（プレイヤー例）幅比 | 50% |
| ヘッダー行高 | 0.40" |
| データ行高 | 残り高さ÷3（約1.50") |

---

## 品質チェックリスト

- [ ] メインメッセージが正しく表示されているか
- [ ] チャートタイトルが正しく表示されているか
- [ ] KBFが**ちょうど3つ**表示されているか
- [ ] 列1（KBF）が紺背景＋白文字で強調されているか
- [ ] 列幅比率が 20% / 30% / 50% で配分されているか
- [ ] プレイヤー例セルでプレイヤー名のみが太字になっているか
- [ ] テーブルが1枚スライド内に収まっているか（オーバーフローなし）
- [ ] 出典が左下に表示されているか
- [ ] セルが PPT 上で編集可能であるか（ネイティブテーブル）

---

## オーケストレーター連携

`market-overview-agent` から呼び出される場合の規約：

| 項目 | 値 |
|---|---|
| 入力JSONファイル名 | `data_NN_market_kbf.json`（NN = デッキ通し番号、例: `data_10_market_kbf.json`）|
| 出力PPTXファイル名 | `slide_NN_market_kbf.pptx` |
| 入力ディレクトリ | `{{WORK_DIR}}/<run_id>/`（オーケストレーター作業領域） |
| 出力ディレクトリ | 同上 |

オーケストレーターは `merge_order.json` の `entries[]` に
`{ "slide_number": NN, "skill_name": "market-kbf-pptx", "data_file": "data_NN_market_kbf.json", "file_name": "slide_NN_market_kbf.pptx" }`
を登録すること。

---

## アセット

| ファイル名 | 用途 |
|---|---|
| `assets/market-kbf-template.pptx` | KBFスライドのテンプレート（competitor-summary-template と同構造） |

## スクリプト

| ファイル名 | 用途 |
|---|---|
| `scripts/fill_market_kbf.py` | JSONデータからKBFテーブル（3行×3列）をネイティブで生成 |

## 参考

| ファイル名 | 内容 |
|---|---|
| `references/sample_data.json` | サンプルJSONデータ（KBF×3） |

---
name: company-overview-pptx-v2
description: >
  BDD（Business Due Diligence）における対象会社の会社概要スライドをPowerPointで生成するスキル（v2）。
  インフォメーション・メモランダムや会社HPの情報をもとに、左側に会社概要テーブル（柔軟な項目数）、
  右側に本社家屋と主要製品/サービスの写真2枚を配置した1枚スライドを作成する。
  PPTXネイティブテーブルオブジェクトで生成するため、ユーザーがPowerPoint上で自由に編集可能。

  **v2の改善点**: 会社HPのURLが入力された場合、web_fetchツールを使って画像を自動取得するため、
  ユーザーが手動で画像をダウンロード・アップロードする手間が不要になった。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「会社概要」「会社概要スライド」「Company Overview」「対象会社概要」という言葉が出た場合
  - 「対象会社の概要をスライドにして」「会社概要をパワポで」「会社プロフィールをスライドに」という要望
  - BDD・デューデリジェンス文脈で「対象会社の基本情報をまとめて」「IM情報をスライドに」という要望
  - 「商号・所在地・設立・代表者をまとめてスライドに」のように会社基本情報の列挙を含む要望
  - ユーザーがインフォメーション・メモランダム（IM）や会社HPのURLを貼り付けて、会社概要のスライド化を求めた場合
  - 「対象会社概要」「ターゲット企業概要」「買収対象の会社概要」をスライドにしたいという要望
supported_brands: [stellar_aiz, roleup]
---

# 会社概要 PowerPoint ジェネレーター v2（BDD用）

BDDにおける対象会社の基本情報を、テーブル＋写真の1枚スライドにまとめるスキル。
**v2ではweb_fetchによる画像自動取得に対応し、ユーザーの手間を大幅に削減した。**

---

## v1からの主な変更点

| 項目 | v1 | v2 |
|------|----|----|
| HP画像の取得 | ユーザーに手動アップロードを依頼 | **web_fetchで自動取得** |
| ユーザーの作業 | URLを開く→保存→アップロード（3ステップ） | **URLを伝えるだけ（0ステップ）** |
| 画像取得失敗時 | N/A | 自動でフォールバック（アップロード依頼） |

---

## 会社概要スライドとは

M&AのBDD（Business Due Diligence）で使用する、対象会社の基本プロフィールを1枚に集約したスライド。
IM（インフォメーション・メモランダム）や会社HPから情報を抽出し、構造化する。

### スライド全体の構成

| 要素 | 配置 | 説明 |
|------|------|------|
| **Main Message** | 上段左寄せ（Title 1） | 対象会社の特徴を一文で要約。**最大65文字**。必ず「〜である」等で締める |
| **Chart Title** | 下段左寄せ（Text Placeholder 2） | スライドタイトル。**10〜20文字**。通常は「対象会社概要：会社概要」 |
| **テーブル** | 左側（約58%） | 会社基本情報。項目数は案件に応じて柔軟に増減可能 |
| **本社家屋 写真** | 右上（約42%） | IMまたはHP画像。未指定時はプレースホルダー枠を表示 |
| **主要製品/サービス 写真** | 右下（約42%） | IMまたはHP画像。未指定時はプレースホルダー枠を表示 |
| **出典** | 左下 | 情報ソースを記載（10pt） |

### 記述ルール

- **Main Message**: 最大65文字。対象会社の業態・強み・特徴を一文で要約する。ユーザーが指定した場合はそのまま使用、指定がない場合はドラフトして確認を取る
- **Chart Title**: 10〜20文字程度。通常は「対象会社概要：会社概要」を使用。会社名を含める場合は「〇〇 会社概要」等

### 標準的なテーブル項目

以下は標準的な項目例。**案件に応じて柔軟に増減する**こと。

| # | 項目 | 記載内容の例 |
|---|------|------------|
| 1 | 商号 | 株式会社〇〇 |
| 2 | 本社所在地 | 都道府県市区町村＋番地 |
| 3 | 設立 | 平成XX年XX月XX日 |
| 4 | 資本金 | XX,XXX千円 |
| 5 | 代表者 | 氏名（年齢） |
| 6 | 事業内容 | 主要事業の概要（改行で複数行可） |
| 7 | 主要販売先 | 顧客名（構成比%）を列挙 |
| 8 | 主要仕入先 | 仕入先企業名を列挙 |
| 9 | 主要外注先 | 外注先企業名を列挙 |
| 10 | 直近売上高 | XXX百万円（20XX年X月期） |
| 11 | 主要取引銀行 | 銀行名 支店名 |
| 12 | 社員数 | XX名（うち有資格者X名） |

業種に応じて「建設許認可」「免許・認可」「拠点数」「グループ会社」等を追加してよい。

### フォントサイズルール

| 要素 | サイズ | 備考 |
|------|--------|------|
| Main Message | 26pt Bold | Meiryo UI / Arial（スクリプトがrunに明示書き込み） |
| Chart Title | 18pt Bold | Meiryo UI / Arial（スクリプトがrunに明示書き込み） |
| テーブル（ラベル・値） | **14pt以上** | 下限厳守 |
| 写真キャプション | **14pt以上** | |
| 出典 | 10pt | |

---

## 入力パターンと処理フロー

### パターンA：IMが入力された場合

**いきなりPowerPointを作成しない**

1. **Step 1: IMから会社概要情報を抽出**し、テーブル項目をドラフトする
2. **Step 2: IMから画像を抽出**する（本社家屋・主要製品/サービス）
3. **Step 3: Markdownでユーザーに提示**し、確認・修正を求める
4. **Step 4: ユーザーの承認後**、PowerPointを生成する

### パターンB：会社HPのURLが入力された場合（v2で大幅改善）

**ユーザーに画像アップロードを依頼しない。web_fetchで自動取得する。**

1. **Step 1: web_fetchでHPを閲覧**し、会社概要情報を抽出。テーブル項目をドラフトする
2. **Step 2: HPのHTMLから画像URLを特定**する（本社家屋・主要製品/サービス）
3. **Step 3: web_fetchで画像データを自動取得**し、ローカルファイルに保存する（詳細は後述の「画像自動取得フロー」参照）
4. **Step 4: Markdownでユーザーに提示**し、確認・修正を求める
5. **Step 5: ユーザーの承認後**、PowerPointを生成する

**フォールバック**: web_fetchで画像取得に失敗した場合のみ、ユーザーに画像アップロードを依頼する

### パターンC：会社概要の詳細が直接入力された場合

1. **Step 1: 内容を確認**し、不足項目があれば提案する
2. **Step 2: 写真の確認**（IMがあればIMから抽出、HPのURLがあればweb_fetchで自動取得、なければアップロードを依頼）
3. **Step 3: 確認後、PowerPointを生成する**

---

## 画像自動取得フロー（v2の核心）

<!-- @if:claude_ai -->
### なぜbash_toolから直接画像をダウンロードできないのか

bash_tool（Linuxコンテナ）にはegressプロキシが設定されており、pypi.org・github.com等のパッケージ管理系ドメイン以外へのアクセスはブロックされる。
一方、`web_fetch`ツールはClaude本体のインフラを経由するため、一般的なWebサイトにもアクセス可能。

```
【bash_tool】  コンテナ → egressプロキシ → ブロック！（一般サイト不可）
【web_fetch】  Claude本体 → Anthropicインフラ → Web全体にアクセス可能
```
<!-- @endif -->

### 画像取得の手順（Claude本体が実行）

以下の手順をClaude本体がツールを組み合わせて実行する。**スクリプト内で完結する処理ではない**ことに注意。

#### Step 3-1: HPページのHTMLを取得

```
web_fetch で会社HPのURL（例：会社概要ページ、事業紹介ページ）を取得
→ HTMLの中から <img> タグのsrc属性を確認し、画像URLを特定
```

**画像URLの選定基準:**
- 本社家屋: 会社概要ページ、アクセスページの社屋外観写真
- 主要製品/サービス: 事業紹介ページ、製品ページの代表製品写真
- 解像度が高い画像を優先（サムネイルではなくオリジナル画像）
- 相対パスの場合はドメインを補完して絶対URLにする

#### Step 3-2: 画像データを取得

```
web_fetch で画像URL（例：https://example.com/images/hq.jpg）を取得
→ Base64エンコードされた画像データが返却される
```

#### Step 3-3: ローカルファイルに保存

```
bash_tool で以下を実行:

# 方法A: ヘルパースクリプトを使用（推奨）
{{PYTHON_BIN}} {{SKILL_DIR}}/scripts/save_image.py \
  --base64-file {{WORK_DIR}}/img_b64.txt \
  --output {{WORK_DIR}}/hq_photo.jpg

# 方法B: Pythonワンライナーで直接保存
{{PYTHON_BIN}} -c "
import base64, sys
data = open('{{WORK_DIR}}/img_b64.txt').read().strip()
if ',' in data and data.startswith('data:'): data = data.split(',',1)[1]
open('{{WORK_DIR}}/hq_photo.jpg','wb').write(base64.b64decode(data))
print('saved')
"
```

※ Base64データが大きい場合は、直接引数で渡すとシェルの制限に引っかかるため、
  **必ずファイル経由（--base64-file）で渡す**こと。

#### Step 3-4: JSONのphotosにローカルパスを指定

```json
"photos": {
  "headquarters": {
    "url": "{{WORK_DIR}}/hq_photo.jpg",
    "caption": "本社家屋"
  },
  "product": {
    "url": "{{WORK_DIR}}/product_photo.jpg",
    "caption": "主要製品/サービス"
  }
}
```

### フォールバック処理

web_fetchで画像取得に失敗するケースがある：
- 画像がJavaScriptで動的にロードされている場合
- CDN等でアクセス制限がかかっている場合
- Base64デコードが失敗した場合

**失敗した場合のみ**、以下のメッセージでユーザーにアップロードを依頼する：

```
HPから画像を自動取得できませんでした。
スライドに使用する写真をアップロードしてください：

1. **本社家屋**: 社屋の外観写真
   → HPの会社概要ページ等にあります: [URL]
2. **主要製品/サービス**: 代表製品・サービスの写真
   → HPの事業・製品ページ等にあります: [URL]
```

---

## Step 1: 画像の取得（優先順位まとめ）

| 優先順位 | 方法 | 条件 |
|---------|------|------|
| 1 | IMから画像を抽出 | IMがPDF等で提供されている場合 |
| 2 | **web_fetchで自動取得（v2新機能）** | HPのURLが提供されている場合 |
| 3 | ユーザーにアップロード依頼 | 上記2つが不可能な場合のみ |

### 優先順位1: IMから画像を抽出する

IMがPDF等で提供されている場合、以下の手順で画像を抽出する：

1. IMをページごとにラスタライズまたは画像抽出ツールで確認
2. 本社建屋の写真、主要製品/サービスの写真を特定
3. 該当画像を抽出し、ローカルファイルとして保存
4. JSONの `photos` セクションにローカルパスを指定

**IMに画像がある場合は、そのまま使用してPowerPointを生成する。**

### 優先順位2: web_fetchで自動取得する（v2新機能）

HPのURLが入力されている場合は「画像自動取得フロー」に従って自動取得する。

### 優先順位3: ユーザーにアップロードを依頼する

優先順位1・2がいずれも不可能な場合のみ、ユーザーに画像のアップロードを依頼する。

### 画像の選定基準

**本社家屋:**
- IMの会社概要セクションに掲載されている社屋外観写真
- HPの場合：会社概要ページ、アクセスページ等

**主要製品/サービス:**
- IMの事業概要セクションに掲載されている製品・サービス写真
- HPの場合：事業紹介、製品ページ等
- 複数ある場合は、対象会社の主力事業に最も関連する画像を選ぶ

---

## Step 2: 会社概要情報を抽出する

IM・HP・その他資料から以下の情報を抽出・整理する：

- 商号（正式名称）
- 本社所在地
- 設立年月日
- 資本金
- 代表者（氏名、可能であれば年齢）
- 事業内容（複数行可）
- 主要販売先（構成比があれば記載）
- 主要仕入先
- 主要外注先
- 直近売上高（決算期を明記）
- 主要取引銀行
- 社員数

業種特有の項目があれば追加する。

---

## Step 3: 会社概要のMarkdown出力フォーマット

ユーザーに確認を求める際は、以下のフォーマットで出力する：

```markdown
## 会社概要 整理結果

**Main Message（※ドラフト、最大65文字）**
〇〇は△△を中核とした□□型の企業であり、〜である

**Chart Title（10〜20文字）**
対象会社概要：会社概要

### テーブル項目
| 項目 | 内容 |
|------|------|
| 商号 | 株式会社〇〇 |
| 本社所在地 | 〇〇県〇〇市... |
| ... | ... |

### 写真
- 本社家屋: ✓ 自動取得済み（[URL]）（または「IMのp.XX より抽出済み ✓」または「取得失敗 → アップロードをお願いします」）
- 主要製品/サービス: ✓ 自動取得済み（[URL]）（または同上）

### 出典
インフォメーション・メモランダム、株式会社〇〇HP
```

確認メッセージ例（画像を自動取得済みの場合）：
> 上記の会社概要整理でよろしいでしょうか？写真はHPから自動取得しました。テーブル項目の修正や写真の変更があればお知らせください。

確認メッセージ例（画像の自動取得に失敗した場合）：
> 上記の会社概要整理でよろしいでしょうか？HPから画像を自動取得できませんでした。
> HPの以下ページ等から写真2枚をダウンロードしてアップロードしてください：
> - 本社家屋: [HPのURL]
> - 主要製品/サービス: [HPのURL]

---

## Step 4: PowerPointの生成

### テンプレートの参照

テンプレートは `assets/company-overview-template.pptx` を使用する。

```bash
TEMPLATE="{{SKILL_DIR}}/assets/company-overview-template.pptx"
```

テンプレートのShape構造は `references/template-mapping.md` を参照。

### 会社概要データのJSON化

会社概要情報を `{{WORK_DIR}}/company_overview_data.json` に以下の形式で保存する：

```json
{
  "title": "対象会社概要：会社概要",
  "main_message": "〇〇は△△を中核とした□□型の企業であり、〜である（最大65文字）",
  "source": "インフォメーション・メモランダム、株式会社〇〇HP",
  "items": [
    {"label": "商号", "value": "株式会社〇〇"},
    {"label": "本社所在地", "value": "〇〇県〇〇市〇〇番地"},
    {"label": "設立", "value": "平成XX年XX月XX日"},
    {"label": "資本金", "value": "XX,XXX千円"},
    {"label": "代表者", "value": "〇〇 太郎（XX歳）"},
    {"label": "事業内容", "value": "〇〇の設計・施工\n△△の製造・販売"},
    {"label": "主要販売先", "value": "A社（40%）、B社（30%）、C社（30%）"},
    {"label": "主要仕入先", "value": "D社、E社、F社"},
    {"label": "主要外注先", "value": "G社、H社"},
    {"label": "直近売上高", "value": "XXX百万円（20XX年X月期）"},
    {"label": "主要取引銀行", "value": "〇〇銀行 〇〇支店"},
    {"label": "社員数", "value": "XX名（うち有資格者X名）"}
  ],
  "photos": {
    "headquarters": {
      "url": "{{WORK_DIR}}/hq_photo.jpg",
      "caption": "本社家屋"
    },
    "product": {
      "url": "{{WORK_DIR}}/product_photo.jpg",
      "caption": "主要製品/サービス"
    }
  }
}
```

**注意**: photos.url にはローカルファイルパスを指定する。web_fetchで取得した画像は事前にローカルに保存済みであること。

### JSON仕様

| フィールド | 必須 | 説明 |
|-----------|------|------|
| `title` | ○ | Chart Title（Text Placeholder 2）。10〜20文字。通常は「対象会社概要：会社概要」 |
| `main_message` | ○ | Main Message（Title 1）。対象会社を一文で要約。**最大65文字** |
| `source` | ○ | 出典テキスト（「出典：」プレフィックスは自動付与） |
| `items` | ○ | テーブル項目の配列。各要素は `{"label": "...", "value": "..."}` |
| `items[].value` | ○ | `\n` で改行可。HTML内で `<br>` に変換される |
| `photos.headquarters.url` | △ | 本社家屋の画像ローカルパス。空文字なら非表示 |
| `photos.headquarters.caption` | △ | 本社写真のキャプション。デフォルト「本社家屋」 |
| `photos.product.url` | △ | 主要製品/サービスの画像ローカルパス。空文字なら非表示 |
| `photos.product.caption` | △ | 製品写真のキャプション。デフォルト「主要製品/サービス」 |

### スクリプト実行コマンド

```bash
pip install python-pptx -q {{PIP_FLAGS}}

# 推奨: brand 指定で起動（template は brand_resolver で自動解決）
{{PYTHON_BIN}} {{SKILL_DIR}}/scripts/fill_company_overview.py \
  --brand stellar_aiz \
  --data {{WORK_DIR}}/company_overview_data.json \
  --output {{OUTPUT_DIR}}/CompanyOverview_output.pptx

# Roleup 出力
{{PYTHON_BIN}} {{SKILL_DIR}}/scripts/fill_company_overview.py \
  --brand roleup \
  --data {{WORK_DIR}}/company_overview_data.json \
  --output {{OUTPUT_DIR}}/CompanyOverview_output.pptx
```

`--template` を明示指定すると brand 解決を上書きできる（任意）。

### 出力確認

<!-- @if:claude_ai -->
```bash
{{PYTHON_BIN}} -m markitdown {{OUTPUT_DIR}}/CompanyOverview_output.pptx
```
<!-- @endif -->
<!-- @if:claude_code -->
```bash
{{OPEN_CMD}} {{OUTPUT_DIR}}/CompanyOverview_output.pptx
```
<!-- @endif -->

内容が正しく反映されているか確認し、ユーザーに提示する。

---

## 品質チェックリスト

PowerPoint生成後、以下を確認：

- [ ] Main Messageが65文字以内で対象会社の特徴を端的に要約しているか
- [ ] Chart Titleが10〜20文字でスライドの文脈を示しているか
- [ ] テーブルの全項目が表示されているか（行が切れていないか）
- [ ] テーブルのフォントが14pt以上になっているか
- [ ] 写真2枚（本社家屋・主要製品/サービス）が正しく表示されているか
- [ ] 出典が左下に10ptで表示されているか
- [ ] PPTXのmarkitdown出力でプレースホルダーが残っていないか

---

## アセット

| ファイル名 | 用途 |
|---|---|
| `assets/stellar_aiz/company-overview-template.pptx` | Stella 16:9 用テンプレート |
| `assets/stellar_aiz/layout.json` | Stella 用座標 (source dynamic textbox 用フォールバック) |
| `assets/roleup/company-overview-template.pptx` | Roleup A4 横用テンプレート（`tools/setup_company_overview_roleup_template.py` で生成: cp roleup template ベース + Overview Table / Photo Area / Photo Caption を追加） |
| `assets/roleup/layout.json` | Roleup 用座標 |

## スクリプト

| ファイル名 | 用途 |
|---|---|
| `scripts/fill_company_overview.py` | JSONデータからPPTXネイティブテーブル＋写真を生成するスクリプト |
| `scripts/save_image.py` | **（v2追加）** web_fetchで取得したBase64画像データをローカルファイルに保存するヘルパー |

## 参考

| ファイル名 | 内容 |
|---|---|
| `references/template-mapping.md` | テンプレートのShape名とスライド各セクションのマッピング表 |

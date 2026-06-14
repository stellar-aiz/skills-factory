---
name: positioning-map-pptx
description: >
  ポジショニングマップ（Positioning Map / 競合マッピング）のPowerPointスライドを生成するスキル。
  2軸の平面上に競合各社をバブル（円）で配置し、市場内のポジション関係を可視化する。
  X軸・Y軸はユーザー定義（例：価格帯×品質、地域範囲×製品領域、成長性×利益率 等）。
  バブルサイズを第3変数（売上・従業員数など）に対応可能。
  4象限ラベルで市場の各領域の性質を表現できる。対象会社は目立つ色＋太枠＋赤太字ラベルで強調。
  右側に「ポジショニングからの示唆」ブレットリストを配置。
  Web情報（各社HP・IR資料・業界レポート）から作成可能。
  company-history-pptx / customer-profile-pptx と同じテンプレート・フォントを使用。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「ポジショニングマップ」「Positioning Map」「2軸マッピング」「競合マッピング」という言葉が出た場合
  - 「価格×品質」「地域×製品」など2軸での競合配置を求められた場合
  - 「業界の勢力図」「競合の位置関係」「バブルチャート」をスライド化したい要望
  - BDD（ビジネスDD）や競合分析の文脈でポジショニング分析のスライド化を求められた場合
  - 「対象会社はどこにいるか」「差別化の余地」を可視化したい要望
  - ユーザーが各社の特徴データを貼り付けて、2軸マッピングのスライド化を求めた場合
  - 「市場マップ」「競合地図」「戦略グループ分析」をスライドにしたいという要望
supported_brands: [stellar_aiz, roleup]

---

# ポジショニングマップ PowerPoint ジェネレーター

2軸の平面上に競合各社をバブル（円）で配置し、市場内のポジション関係と戦略グループを可視化するスキル。
BDD（ビジネスデュー・ディリジェンス）・競合分析・市場参入戦略・差別化検討の場面で、
「対象会社はどこに位置するか」「誰と競合密集しているか」「空いている象限はあるか」を一目で示す。

---

## スライド構成

| セクション | 位置 | 内容 |
|---|---|---|
| **メインメッセージ** | 最上部 | 最大65文字（hard-fail）。ポジショニングからの示唆を「〜すべき」で締める |
| **チャートタイトル** | メインメッセージ直下 | 「競合ポジショニングマップ：○○市場」 |
| **ポジショニングマップ** | 左側（約8.0×5.2in） | 2軸バブルチャート（手動描画） |
| **示唆パネル** | 右側（約4.5in幅） | ポジショニングからの示唆3〜5項目 |
| **出典** | 左下 | 情報ソースの記載 |

### テンプレート構造（v2: 固定枠方式）

骨組み（下記 1〜7・9 のうち**バブル以外**）は `assets/<brand>/positioning-map-template.pptx` に
**名前付きシェイプ** として焼き込み済み。`scripts/fill_positioning_map.py` の実行時の責務は
**(1) 名前付き枠へのテキスト流し込み（空データの枠は非表示）** と **(2) バブル＋バブルラベルの描画** の
2 つだけ。バブル配置の幾何基準は `MAP_FRAME` シェイプの座標を single source of truth として読む。

| シェイプ名 | 役割 | fill の扱い |
|---|---|---|
| `Title 1` / `Text Placeholder 2` | main_message / chart_title | テキスト流し込み |
| `SECTION_LEFT_TITLE` / `IMPL_TITLE` | 左右セクションタイトル | テキスト流し込み |
| `MAP_FRAME` | マップ外枠（**幾何アンカー**） | バブル配置の基準座標を読むのみ |
| `MAP_GUIDE_V` / `MAP_GUIDE_H` | 十字ガイドライン（破線） | 触らない |
| `AXIS_X_LABEL` / `AXIS_X_LOW` / `AXIS_X_HIGH` | X 軸ラベル | テキスト流し込み（矢印は fill が付与） |
| `AXIS_Y_LABEL` / `AXIS_Y_LOW` / `AXIS_Y_HIGH` | Y 軸ラベル（回転） | テキスト流し込み |
| `QUAD_TL` / `QUAD_TR` / `QUAD_BL` / `QUAD_BR` | 4 象限ラベル（固定 4 枠） | 流し込み or 非表示 |
| `IMPL_1` … `IMPL_5` | 示唆（固定 5 枠） | 流し込み or 非表示 |
| `Source 3` | 出典 | テキスト流し込み |

排他ルール: `quadrants` が定義されていれば `AXIS_Y_LOW`/`AXIS_Y_HIGH` を非表示にし、
定義されていなければ `QUAD_*` を非表示にする（同じ位置の重複表示を回避）。

**テンプレを変更したいとき**: 生成後のテンプレを **PowerPoint で直接微調整して正本コミット**するのが基本。
ゼロから作り直す場合のみ `tools/build_positioning_map_template.py`（リポジトリ側の bootstrap・install 非同梱）を
使う（**再実行は人の手調整を上書きするので注意**）。

> **PowerPoint 微調整時の注意**: シェイプの**名前（`MAP_FRAME` 等）は変えない**こと（fill が名前で探すため）。
> マップ要素を**グループ化・移動・リサイズするのは可**（fill はグループ内を再帰探索し、`MAP_FRAME` の
> 絶対座標をグループ変換から解決してバブルを正しく載せる）。

### マップの構成要素

1. **マップ外枠**: 薄いグレー背景の矩形、黒線の枠
2. **十字ガイドライン**: マップを4象限に分ける破線（中央縦線・中央横線）
3. **X軸ラベル**: 下部中央にメインラベル（例：「製品専門性」）
4. **X軸 low/high**: 左右両端に「← 低側」「高側 →」（グレー、小サイズ）
5. **Y軸ラベル**: 左側中央に縦書きメインラベル（例：「価格帯」）
6. **Y軸 low/high**: 縦書き（※象限ラベルが定義されている場合は重複回避のため非表示）
7. **4象限ラベル**: 各象限の性質（例：「ハイエンド・専門」「マス・コモディティ」）
8. **バブル**: 各プレイヤーを円で表現（色・サイズ・位置で情報表現）
9. **バブルラベル**: 企業名（位置カスタマイズ可能）

### バブル仕様

- **形状**: 楕円（MSO_SHAPE.OVAL）
- **直径**: `size` フィールドで指定（最小0.40in、最大0.95inで自動スケーリング）
- **色**: `color` フィールド（#RRGGBB）。未指定時は自動割り当て
- **透明度**: 通常30%透明、対象会社のみ10%透明（視認性重視）
- **枠線**: 通常1.0pt、対象会社は2.5pt（目立つ）

### 対象会社の強調

- バブル塗りつぶし: 指定色（デフォルト赤系 #E15759）
- バブル枠線: 濃赤 #8B2C2E、2.5pt（通常の2.5倍）
- バブル透明度: 10%（ほぼ不透明で目立つ）
- ラベル: 12pt Bold（通常は11pt Regular）、濃赤色

### バブルラベルの配置

`label_position` フィールドで位置を指定：
- `"bottom"` (デフォルト): バブルの下
- `"top"`: バブルの上
- `"left"`: バブルの左
- `"right"`: バブルの右

**自動境界調整**: ラベルがマップ境界をはみ出す場合、自動的に反対側へ調整される。

---

## 軸の設計ガイドライン

### よく使われる2軸の組み合わせ

| X軸 | Y軸 | 用途 |
|---|---|---|
| 価格帯 | 品質・機能 | 一般的な製品ポジショニング |
| 地域カバレッジ | 製品ラインアップ | 事業規模ポジショニング |
| 市場成長率 | 市場シェア | BCGマトリクス風 |
| 技術成熟度 | 市場規模 | 技術ポジショニング |
| 専門特化度 | 顧客規模 | ターゲット顧客ポジショニング |
| デジタル化度 | 人材規模 | DX成熟度ポジショニング |

### 4象限の命名例

- **ハイエンド・専門** ↔ **マス・コモディティ**
- **プレミアム** ↔ **バリュー**
- **ニッチ・特化** ↔ **広域・総合**
- **スター** ↔ **問題児** ↔ **金のなる木** ↔ **負け犬**（BCGマトリクス）

### バブルサイズ（第3変数）の使い方

- **売上高** (最も一般的)
- **従業員数**
- **販売台数・販売数量**
- **市場シェア**
- **時価総額**

---

## 軸選択ゲート（必須・スキップ不可）

**2軸はスキルが勝手に確定しない。必ずユーザーに選ばせてから生成する。**
ポジショニングマップの価値は「どの2軸で切るか」に集約されるため、ここを自動化すると
「勝手に作られた」結果になる。以下のゲートを **JSON 生成・スクリプト実行の前** に必ず通す。

1. **Step 0（ゲート）: 軸候補を 3 案、根拠付きで提示する**
   - 各案について `X軸 / Y軸` の名称、それぞれの `low / high` の意味、そして
     **1〜2 行の根拠**（なぜこの市場でこの軸を取るとプレイヤーが分散するか）を示す。
   - 「軸の設計ガイドライン」の組み合わせ例表を候補ソースに使ってよい。
   - **ユーザーの選択（または代案の提示）を待つ**。選択されるまで先へ進まない。
2. **Step 1: 選択された軸で各社の特徴を抽出**し、x/y/size 値を推計。
   - **選んだ軸の根拠を provenance に記録**: `$.x_axis` / `$.y_axis` に「なぜこの市場でこの軸か」を `rationale`、出典を `source_name`/`source` に。
   - **各社の x/y/size を決めたその場で provenance を記録**: ソースを読んで配置を判断した瞬間に、`$.players[i].x` / `.y` / `.size` の `rationale`（なぜこの値か）・`source_name`（出典名）・`source`（URL or 社内ファイル名）・`confidence`（実測/推定/仮置）を埋める。**硬い数値（売上・シェア・海外比率 等）は `actual_value`/`actual_unit`/`actual_metric` にも構造化する**（WEB裏取りの入力になる）。**後でまとめて書こうとせず取得時に書く**（→「出典・根拠トレース」節）。
3. **Step 2: 推計結果を Markdown でユーザーに提示**（プレビュー）。
4. **Step 3: ユーザーの承認後**、PowerPoint を生成し、続けて provenance トレース表を生成する。

> **市場概要オーケストレーター（market-overview-agent）経由でも本ゲートは省略しない。**
> オーケストレーターは軸を想像で `data_07_positioning.json` に書き込まず、必ず軸候補 3 案を
> ユーザーに提示・選択させてから軸を確定すること。

データが既に整理済み（軸も x/y も確定済み）で入力された場合のみ、Step 0〜2 を省略し
内容確認 → 生成に進んでよい。ただし軸がユーザー意図と一致しているかは必ず一度確認する。

---

## JSONデータ仕様

`{{WORK_DIR}}/positioning_map_data.json` に以下の形式で保存する：

```json
{
  "source": "出典：各社公式HP・プレスリリース・IR資料、MarkLines、業界レポート",
  "main_message": "対象会社は中価格・汎用領域で競合が密集しており、高価格・専門特化にポジションを移すことで差別化と利益率向上を図るべき",
  "chart_title": "競合ポジショニングマップ：EV市場における価格帯×専門性",
  "target_company": "対象会社",
  "section_title": "ポジショニングマップ",
  "x_axis": {
    "label": "製品専門性",
    "low": "汎用モデル中心",
    "high": "専門・高機能モデル",
    "min": 0,
    "max": 10
  },
  "y_axis": {
    "label": "価格帯",
    "low": "低価格帯",
    "high": "高価格帯",
    "min": 0,
    "max": 10
  },
  "quadrants": {
    "top_left": "高価格・汎用",
    "top_right": "ハイエンド・専門",
    "bottom_left": "マス・コモディティ",
    "bottom_right": "ニッチ・専門"
  },
  "players": [
    {"name": "対象会社", "x": 4.5, "y": 5.0, "size": 50, "color": "#E15759", "label_position": "bottom"},
    {"name": "テスラ", "x": 8.0, "y": 8.0, "size": 90, "label_position": "left"},
    {"name": "BYD", "x": 5.5, "y": 3.0, "size": 95, "label_position": "bottom"}
  ],
  "implications_title": "ポジショニングからの示唆",
  "implications": [
    "対象会社は中価格・汎用領域（マップ中央）に位置し、VW・現代自動車と競合密集",
    "テスラは高価格・高専門性でポジション確立、圧倒的な売上規模",
    "対象会社は『高価格・専門特化』象限への移行で差別化余地あり"
  ]
}
```

### JSONフィールド仕様

| フィールド | 型 | 必須 | 説明 |
|---|---|---|---|
| `source` | string | 任意 | 出典テキスト |
| `main_message` | string | 必須 | メインメッセージ（最大65文字（hard-fail）） |
| `chart_title` | string | 任意 | チャートタイトル |
| `target_company` | string | 任意 | 対象会社名（強調ハイライト対象） |
| `section_title` | string | 任意 | 左セクションタイトル |
| `x_axis.label` | string | 推奨 | X軸のメインラベル |
| `x_axis.low/high` | string | 推奨 | X軸の両端表記 |
| `x_axis.min/max` | number | 任意 | X軸の数値範囲（デフォルト 0〜10） |
| `y_axis.*` | (同上) | | Y軸の設定 |
| `quadrants.top_left/top_right/bottom_left/bottom_right` | string | 任意 | 4象限のラベル（定義しない場合はY軸low/highが表示される） |
| `players` | array | 必須 | プレイヤー配列 |
| `players[].name` | string | 必須 | 企業名（target_companyと一致させる対象会社は強調） |
| `players[].x` | number | 必須 | X座標（x_axis の min〜max の範囲内） |
| `players[].y` | number | 必須 | Y座標（y_axis の min〜max の範囲内） |
| `players[].size` | number | 任意 | バブルサイズ（相対値。全プレイヤー間でスケーリング） |
| `players[].color` | string | 任意 | 色（#RRGGBB）。未指定時は自動割り当て |
| `players[].label_position` | string | 任意 | ラベル位置: "bottom"(デフォルト) / "top" / "left" / "right" |
| `implications_title` | string | 任意 | 右パネルのタイトル。デフォルト「ポジショニングからの示唆」 |
| `implications` | array | 任意 | 示唆のブレット項目配列。**固定 5 枠**（`IMPL_1`〜`IMPL_5`）。6 個以上は先頭 5 件のみ表示し stderr に WARN。3〜5 個推奨。1 項目 ≤ 約45全角を目安（固定枠なので長文は折返しで枠を圧迫する） |

> **出典・根拠は本体 JSON には書かない。** 各社 x/y/size と軸選択の出典・根拠は**別ファイル**
> `positioning_map_provenance.json`（json_path をキーにしたサイドカー）に持たせる。本体 JSON は値のみ。
> 詳細は後述の「出典・根拠トレース（provenance サイドカー）」節。

---

## 出典・根拠トレース（provenance サイドカー）

「どの値をどのソースから取ってきたか」を人が細部まで辿れるよう、本体データ（値のみ）とは
**別ファイル** `positioning_map_provenance.json` に出典・根拠を構造として持たせる。

### 設計思想（3層モデル）

- **① 調査・記録層（あなた＝LLM）**: Web を読み x/y/軸を判断し、**取得・判断したその場で provenance を記録する**。URL は WebFetch/WebSearch で取得した瞬間に得られる最も具体的な出典。**硬い数値（売上・シェア・海外比率 等）は散文の rationale だけでなく `actual_value`/`actual_unit`/`actual_metric` に構造化する**（WEB裏取りの入力になる）。
- **② レンダリング・検証層（決定論的・Web なし）**: `fill_positioning_map.py`（data→PPTX）と
  `build_provenance_trace.py`（provenance の検証＋表生成）。①が記録した内容を検証・可視化するだけ。
- **②.5 WEB 検証層（provenance を入力に裏取り）**: `build_provenance_trace.py --mode verify-prep` が
  「裏取りすべき硬い事実」の worklist を決定論的に抽出 → **あなた（LLM）が記録 source URL を WebFetch して
  `actual_value` を照合**（裏付けない/到達不可なら WebSearch にフォールバック）→ `render --verification` が
  結果をトレース表へ統合し、誤帰属・食い違いがあれば非ゼロ終了する。
- `fill_positioning_map.py` は provenance を**一切参照しない**（純粋レンダラー）。トレースは必ず別コマンドで実行する。

### サイドカーのスキーマ

`json_path` をキー、値は `{value, actual_value, actual_unit, actual_metric, rationale, source_name, source, confidence}`。

| フィールド | 必須 | 説明 |
|---|---|---|
| `rationale` | 必須（空欄NG） | なぜこの値・この軸か |
| `source_name` | 必須（空欄NG） | 人が読む出典名（例「Tesla IR 2024 決算説明資料」） |
| `source` | 必須（空欄NG） | 実際に辿れるロケータ＝**Web URL** または**社内資料ファイル名**（例 `競合価格_2024.xlsx`） |
| `confidence` | 必須（空欄NG） | `実測`（出典の硬い数値）/ `推定`（代理指標から推計）/ `仮置`（定性判断のみ） |
| `value` | 任意 | 本体値との drift 検知用。**size は正規化相対値（例 100）なので検証には使えない**点に注意 |
| `actual_value` | 任意（推奨） | **検証対象の生数値**（例 `"3382"`）。これが WEB検証 worklist の採録キー。`confidence=実測` なら必ず埋める |
| `actual_unit` | 任意 | 生数値の単位（例 `"億円"` / `"%"`） |
| `actual_metric` | 任意 | 生数値が何の指標か（例 `"連結売上(2024年12月期)"` / `"海外売上比率"`） |

**対象 path（事実・配置系のみ）**: `$.players[i].x` / `.y` / `.size`（本体に size があるとき）＋ `$.x_axis` / `$.y_axis`。
main_message・示唆（結論文）は対象外。記入例は `references/sample_provenance.json`。

> **actual_\* の使いどころ**: x/y の**配置値そのもの**（x=9 等）は定性的な見立てで検証不能だが、その配置の
> **根拠としてLLMが引いた硬い事実**（例: パイロット y の「海外比率70%超」、三菱 y の「海外59.7%」）は事実なので、
> その path の `actual_value`（=`"59.7"`, unit=`"%"`）に構造化すれば size と同じく WEB裏取りの対象になる。

### ワークフロー（記録 → 検証表 → WEB裏取り・いずれも fill とは独立）

```bash
# ① 雛形生成: 対象 path を列挙した空欄 provenance を自動生成（漏れを構造的に防ぐ）
python <SKILL_DIR>/scripts/build_provenance_trace.py --mode skeleton \
  --data {{WORK_DIR}}/positioning_map_data.json \
  --out  {{WORK_DIR}}/positioning_map_provenance.json
#   既存ファイルには欠けた path / actual_* だけ追補（手書き分は温存）。--force で全上書き。

# ② あなたが調査の過程で rationale/source_name/source/confidence を逐次埋める（取得時記録）。
#    硬い数値は actual_value/actual_unit/actual_metric にも構造化する。

# ③ トレース表の生成＋検証（render）
python <SKILL_DIR>/scripts/build_provenance_trace.py --mode render \
  --data       {{WORK_DIR}}/positioning_map_data.json \
  --provenance {{WORK_DIR}}/positioning_map_provenance.json \
  --out-md     {{OUTPUT_DIR}}/PositioningMap_output_provenance.md
```

- `render` は `項目(json_path) | 値 | 根拠 | 出典名 | 出典(ロケータ) | 自信度 | 状態` の Markdown 表を出す。
- **空欄はNG**: 対象 path のエントリ欠落、または `rationale/source_name/source/confidence` のいずれかが空だと
  `⚠NG` 表示＋NG一覧＋**非ゼロ終了**（表自体は何が NG か見せるため出力する）。空欄を埋めて再実行する。
- `value` が本体値とズレると `⚠値ズレ`、`confidence` が語彙外だと WARN。

### WEB検証フロー（②.5・任意だが硬い数値があるとき推奨）

provenance に記録した**硬い事実**を、LLM が記録 source URL を fetch して照合する。fact-check-reviewer の
盲目再検索と違い、**LLMが主張した出典そのものを当たる**ので、出典の誤帰属・捏造（`source_mismatch`）まで検出できる。

```bash
# ④ verify-prep: 裏取りすべき硬い事実の worklist を抽出（決定論・WEBなし）
python <SKILL_DIR>/scripts/build_provenance_trace.py --mode verify-prep \
  --data       {{WORK_DIR}}/positioning_map_data.json \
  --provenance {{WORK_DIR}}/positioning_map_provenance.json \
  --out        {{WORK_DIR}}/positioning_map_verification.json
#   採録条件: confidence ∈ {実測, 推定} かつ actual_value 非空。仮置・純定性配置は対象外。
#   既存 worklist の記入済み検証欄は温存（追補）。--force で全上書き。

# ⑤ あなた（LLM）が各エントリの source を WebFetch → actual_value を照合 → 検証欄を埋める。
#    記録URLが裏付けない/到達不可なら WebSearch にフォールバック。
#    verification_result は次の語彙から1つ:
#      confirmed / source_mismatch / source_unreachable / discrepancy / not_found / stale
#    併せて verified_value（ソースが示した値）・method（fetch_source / web_search_fallback）・
#    verification_note を記入する。

# ⑥ render に --verification を渡して検証結果をトレース表へ統合
python <SKILL_DIR>/scripts/build_provenance_trace.py --mode render \
  --data         {{WORK_DIR}}/positioning_map_data.json \
  --provenance   {{WORK_DIR}}/positioning_map_provenance.json \
  --verification {{WORK_DIR}}/positioning_map_verification.json \
  --out-md       {{OUTPUT_DIR}}/PositioningMap_output_provenance.md
```

- `--verification` 指定時は「## WEB検証結果（記録URL照合）」節が追加される。`--verification` 省略時は従来どおり（後方互換）。
- **非ゼロ終了**: `verification_result` が空（未検証）、または `source_mismatch`/`discrepancy`/`not_found` のいずれかがあると `1` を返す。
  `source_unreachable`/`stale` は WARN 止まり、`confidence=実測` なのに `actual_value` 未記入も WARN。
- **これは出典照合であって万能のファクトチェックではない**。x/y の**定性配置値そのもの**は依然検証対象外。
  複数の独立ソースでのクロス検証が要る場合のみ `fact-check-reviewer` を併用する。

---

## スクリプト実行コマンド

```bash
pip install python-pptx -q --break-system-packages

python <SKILL_DIR>/scripts/fill_positioning_map.py \
  --data {{WORK_DIR}}/positioning_map_data.json \
  --brand stellar_aiz \
  --output {{OUTPUT_DIR}}/PositioningMap_output.pptx
```

`--brand` を `roleup` に切り替えると Roleup ブランド (A4 横、Yu Gothic UI、褐色アクセント) で生成。
roleup ではマップが狭くなるため、バブルラベル幅を縮小し、上部 quadrant ラベルを枠外に配置する自動調整が入る。
`--template` は省略可（brand から `assets/<brand>/positioning-map-template.pptx` を自動解決）。

スライド生成後、上記「出典・根拠トレース」の `build_provenance_trace.py --mode render` を実行して
`*_provenance.md` を併せて納品する。硬い数値（売上・シェア・海外比率 等）を含む場合は、続けて
`verify-prep` → 記録URLの WebFetch 照合 → `render --verification` の WEB検証フローまで通す。

---

## デザイン仕様

### フォントサイズ

| 要素 | サイズ |
|---|---|
| メインメッセージ | 28pt (Bold) |
| チャートタイトル | 11pt |
| セクションタイトル | 14pt (Bold、下線付き) |
| 軸メインラベル (X/Y) | 12pt (Bold) |
| 軸 low/high ラベル | 10pt (グレー) |
| 象限ラベル | 10pt (Italic、薄グレー) |
| バブルラベル（通常） | 11pt |
| バブルラベル（対象会社） | 12pt (Bold、濃赤) |
| 示唆項目 | 12pt |
| 出典 | 10pt (#666666) |

### 色

| 要素 | カラーコード |
|---|---|
| テキスト（本文） | #333333 |
| マップ背景 | #FAFAFA |
| マップ枠線 | #333333 |
| 十字ガイドライン（破線） | #C0C0C0 |
| 軸 low/high ラベル | #666666 |
| 象限ラベル | #999999 |
| 対象会社バブル色（デフォルト） | #E15759 |
| 対象会社バブル枠・ラベル色 | #8B2C2E |
| 示唆のブレットマーカー | #2E4A6B |

### レイアウト定数

| 要素 | 値 |
|---|---|
| パネル開始Y | 1.55in |
| パネル高さ | 5.35in |
| 左パネル（マップ）開始X / 幅 | 0.41in / 7.70in |
| 右パネル（示唆）開始X / 幅 | 8.30in / 4.65in |
| マップ内マージン（左/右/上/下） | 0.75 / 0.20 / 0.70 / 0.80 in |
| バブル直径（最小〜最大） | 0.40 〜 0.95 in |
| 出典Y | 6.93in |

---

## 品質チェックリスト

- [ ] メインメッセージが最大65文字（hard-fail）以内で「〜すべき」で締められているか
- [ ] チャートタイトルが表示されているか
- [ ] マップ外枠・十字ガイドラインが表示されているか
- [ ] X軸・Y軸のメインラベルが表示されているか
- [ ] X軸の low/high ラベル（「← 低」「高 →」）が両端に表示されているか
- [ ] 4象限ラベル（定義されている場合）が各象限の隅に表示されているか
- [ ] 全プレイヤーのバブルが正しい座標に配置されているか
- [ ] バブルラベルが重ならず読みやすいか（label_positionで調整）
- [ ] 対象会社バブルが目立つ色・太枠・Bold ラベルで強調されているか
- [ ] 右側に示唆のブレット項目が3〜5個表示されているか
- [ ] 出典が左下にグレー表示、©マークと重なっていないか

---

## アセット / スクリプト / 参考

**ランタイム必須ファイル（install/配布 zip に同梱）:**

| ファイル | 用途 |
|---|---|
| `assets/<brand>/positioning-map-template.pptx` | 名前付きシェイプを焼き込んだ固定枠テンプレート（正本） |
| `scripts/fill_positioning_map.py` | 生成スクリプト（テキスト流し込み＋バブル描画のみ。provenance は参照しない） |
| `scripts/build_provenance_trace.py` | 出典・根拠トレース（skeleton 雛形生成＋verify-prep worklist 抽出＋render 表生成・検証）。Web アクセスなし |
| `references/sample_data.json` | サンプルデータ（EV市場の架空ポジショニング例） |
| `references/sample_provenance.json` | provenance サイドカーの記入済みサンプル（actual_* 含む。json_path × 出典・根拠） |
| `references/sample_verification.json` | WEB検証 worklist の記入済みサンプル（confirmed / source_mismatch / source_unreachable 例） |

**参考/開発用ファイル（スキル外・install 非同梱）:**

| ファイル | 用途 |
|---|---|
| `tools/build_positioning_map_template.py` | テンプレを 2 ブランド分生成する bootstrap（ゼロ作り直し時のみ。再実行は手調整を上書き） |
| `docs/positioning-map-pptx-architecture.md` | 設計理解ドキュメント（どの情報がどこに書かれるか） |

---

## 注意事項

- **座標の主観性**: x/y値は定性判断で配置することが多い。できるだけ客観指標（価格実績、製品数等）を基準にし、**各社の x/y/size と軸選択の根拠・出典を provenance サイドカー（`positioning_map_provenance.json`）に必ず記録する**（空欄はNG。「出典・根拠トレース」節参照）
- **バブルサイズ統一**: 同じ指標（売上・従業員数など）で統一することを推奨。異なる指標を混ぜると誤解を招く
- **象限ラベルの命名**: 中立的な表現を使う。「負け組」「勝ち組」など主観的な表現は避ける
- **ラベル重なり**: プレイヤーが密集している場合、label_position を明示的に指定して重なりを回避する
- **プレイヤー数**: **既定5社・上限5社**（`target_company` を含む）。`fill_positioning_map.py` が 2〜5 の範囲で hard-fail 検証する。市場概要オーケストレーター（market-overview-agent）の `max_competitors`（既定5・上限5）と一致させ、market-share / competitor-summary / market-kbf と同じ5社で揃える。多すぎると視認性が低下するため上限5を維持する
- **対象会社の表記**: `target_company` と `players[].name` を完全一致させる

---

## 競合分析デッキでの使用例

`competitor-analyst-agent` オーケストレータースキルから呼び出される位置づけ:

```
1. 対象会社プロファイル (customer-profile-pptx)
2. 対象会社 事業ポートフォリオ (business-portfolio-pptx)
3. 対象会社 SWOT (swot-pptx)
4. 市場環境 - 市場規模推移 (market-environment-pptx)
5. 市場環境 - 市場シェア分析 (market-share-pptx)
6. 市場環境 - ポジショニングマップ ← 本スキル
7. 競合比較サマリー (competitor-summary-pptx)
8. 各競合プロファイル (customer-profile-pptx × N)
```

市場環境セクションの締めとしてポジショニングマップを配置することで、
「市場の大きさ（4枚目）」「主要プレイヤーのシェア（5枚目）」「各社の戦略ポジション（6枚目）」
の3段階で市場環境の全体像を示せる。

---

## オーケストレーター連携

`market-overview-agent` から呼び出される場合の規約：

| 項目 | 値 |
|---|---|
| 入力JSONファイル名 | `data_07_positioning.json` |
| 出力PPTXファイル名 | `slide_07_positioning.pptx` |
| 入力ディレクトリ | `{{WORK_DIR}}/<run_id>/`（オーケストレーター作業領域） |
| 出力ディレクトリ | 同上 |
| 競合5社上限 | デッキ全体で同じ5社（market-overview-agent が一貫性を担保）|

オーケストレーターは `merge_order.json` の `entries[]` に
`{ "slide_number": 7, "skill_name": "positioning-map-pptx", "data_file": "data_07_positioning.json", "file_name": "slide_07_positioning.pptx" }`
を登録する。

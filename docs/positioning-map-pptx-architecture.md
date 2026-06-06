# positioning-map-pptx 設計理解ドキュメント

> 「どの情報が、どのファイルに、なぜ書かれているか」の全体像。
> **設計理解用**であり、実行に必須のファイルではない（読まなくてもスキルは動く）。
> v2（テンプレ固定化リファクタ, 2026-06）時点の構造を反映。

---

# Part 1. 概要（まず全体像を3分で）

## 1-1. いちばん大事な考え方：情報を「変わらないもの」と「変わるもの」に分ける

ポジショニングマップ1枚は、性質の違う情報の組み合わせでできています。
v2 では **「毎回同じもの＝固定」「調査ごとに変わるもの＝可変」** で置き場所を分けました。

```
┌─────────────────────────────────────────────────────────────┐
│  変わらない（固定）                  変わる（可変）            │
│  ───────────────                  ─────────────            │
│  ・枠線、十字線                     ・軸の名前（価格×品質…）   │
│  ・軸ラベルの「位置」                ・各社の位置（バブル座標） │
│  ・象限ラベルの「位置」(4枠)         ・各社の名前・色・大きさ   │
│  ・示唆の「位置」(5枠)               ・象限/示唆の「文言」      │
│  ・色、フォント、紙面サイズ          ・出典文                  │
│        ↓                                  ↓                  │
│   テンプレ .pptx に焼き込む          データ JSON で渡す        │
│   （人が PowerPoint で微調整可）     （調査のたびに書く）       │
└─────────────────────────────────────────────────────────────┘
```

## 1-2. 登場ファイルと役割（5つだけ覚えればよい）

| ファイル | ひとことで言うと | 何を決める |
|---|---|---|
| **テンプレ `.pptx`** | スライドの「型」 | 枠・軸・象限・示唆の **位置／色／フォント**（=見た目の骨組み） |
| **データ JSON** | 調査結果 | 軸名・各社の位置・文言・出典（=毎回変わる中身） |
| **`fill_positioning_map.py`** | 組立工 | JSON の文字を型に流し込み、バブルだけ描く |
| **`build_..._template.py`** | 型を作る工場 | テンプレ `.pptx` を生成（座標の元データを持つ） |
| **`theme.json`**（_common） | ブランド設定 | 色・フォント・紙面サイズ（stellar / roleup） |

> **ファイルの置き場所**: ランタイム必須なのは `scripts/fill_positioning_map.py` / `assets/<brand>/*.pptx` /
> `references/sample_data.json`（これらだけが install/配布 zip に同梱される）。
> 型を作る bootstrap は **`tools/build_positioning_map_template.py`**、本ドキュメントは
> **`docs/positioning-map-pptx-architecture.md`** とスキル外に置き、install には含めない。
> `layout.json` は v2 で不要になったため **削除済み**（詳細は Part 2-7）。

## 1-3. データの流れ（生成の順番）

```
        ┌──── 一度だけ（型作り）────┐
        │                          │
  build_template.py ──生成──▶ テンプレ .pptx  ◀──人が PowerPoint で微調整──┐
        ▲                          │                                      │
   theme.json ────色/font/紙面────┘                              （正本としてコミット）

        ┌──── 調査のたびに（中身入れ）────┐
        │                                │
   データ JSON ──┐                       │
                 ├──▶ fill_positioning_map.py ──▶ 完成 .pptx
   テンプレ .pptx ─┘         │
   theme.json ──バブル色/font─┘
```

- **テンプレ作り**は基本一度きり（or レイアウト変更時のみ）。
- **中身入れ**は調査のたび。fill は「文字を流す」＋「バブルを描く」だけ。

## 1-4. スライアウト上の位置と置き場所（ざっくり地図）

```
┌───────────────────────────────────────────────────────────┐
│  Title 1            ← main_message（JSON）                  │ ← テンプレが位置/font
│  Text Placeholder 2 ← chart_title（JSON）                   │ ← テンプレが位置/font
├──────────────────────────────┬────────────────────────────┤
│ SECTION_LEFT_TITLE(JSON文字)  │ IMPL_TITLE (JSON文字)       │
│ ┌──────────────────────────┐ │  ● IMPL_1  ← implications[0]│
│ │ QUAD_TL        QUAD_TR    │ │  ● IMPL_2  ← implications[1]│
│ │   AXIS_Y_LABEL            │ │  ● IMPL_3  ← implications[2]│
│ │      ○ ← バブル(players[])│ │  ● IMPL_4  ← implications[3]│
│ │   MAP_FRAME(=バブル基準)  │ │  ● IMPL_5  ← implications[4]│
│ │ QUAD_BL        QUAD_BR    │ │   (空枠は自動で非表示)      │
│ │      AXIS_X_LABEL         │ │                            │
│ └──────────────────────────┘ │                            │
│   AXIS_X_LOW    AXIS_X_HIGH   │                            │
├──────────────────────────────┴────────────────────────────┤
│  Source 3  ← source（JSON）                                 │
└───────────────────────────────────────────────────────────┘
   ※ 枠の「位置」はテンプレ固定／枠の「中の文字」は JSON
   ※ バブルだけは実行時に MAP_FRAME を基準に描画
```

---

# Part 2. 詳細

## 2-1. 「どの情報がどこに書かれるか」完全マッピング表

スライド上の各要素について「位置・色・フォントを決めるもの」と「文字内容を決めるもの」を分けて示す。

| スライド上の要素 | シェイプ名 | 位置/色/font を決める | 文字内容を決める（JSON フィールド） |
|---|---|---|---|
| 上部メッセージ | `Title 1` | テンプレ（既存PH） | `main_message`（stella）/ `chart_title`（roleup）※2-4 |
| 副題 | `Text Placeholder 2` | テンプレ（既存PH） | `chart_title`（stella）/ `main_message`（roleup）※2-4 |
| 左タイトル | `SECTION_LEFT_TITLE` | テンプレ | `section_title`（既定「ポジショニングマップ」） |
| 右タイトル | `IMPL_TITLE` | テンプレ | `implications_title`（既定「ポジショニングからの示唆」） |
| マップ外枠 | `MAP_FRAME` | テンプレ | （文字なし。**バブル配置の基準座標**として読まれる） |
| 十字線 | `MAP_GUIDE_V` / `MAP_GUIDE_H` | テンプレ | （fill は一切触らない） |
| X軸ラベル | `AXIS_X_LABEL` | テンプレ | `x_axis.label` |
| X軸 左端 | `AXIS_X_LOW` | テンプレ | `"← " + x_axis.low` |
| X軸 右端 | `AXIS_X_HIGH` | テンプレ | `x_axis.high + " →"` |
| Y軸ラベル | `AXIS_Y_LABEL`（回転） | テンプレ | `y_axis.label` |
| Y軸 上端 | `AXIS_Y_HIGH`（回転） | テンプレ | `y_axis.high + " →"` ※象限なし時のみ |
| Y軸 下端 | `AXIS_Y_LOW`（回転） | テンプレ | `"← " + y_axis.low` ※象限なし時のみ |
| 象限ラベル×4 | `QUAD_TL/TR/BL/BR` | テンプレ | `quadrants.top_left/top_right/bottom_left/bottom_right` ※象限あり時のみ |
| 示唆×5 | `IMPL_1`〜`IMPL_5` | テンプレ（bullet付） | `implications[0]`〜`[4]`（6件目以降は切り捨て＋WARN） |
| 出典 | `Source 3` | テンプレ | `source` |
| 各社バブル | （実行時に動的生成） | **fill が実行時に描画**：位置=`MAP_FRAME`座標＋`players[].x/y`、色=`players[].color` or theme palette、大きさ=`players[].size` | `players[].name`（ラベル） |

### 排他ルール（重複表示の回避）
- `quadrants` が**ある** → `QUAD_*` を表示し、`AXIS_Y_LOW/HIGH` を**非表示**。
- `quadrants` が**ない** → `AXIS_Y_LOW/HIGH` を表示し、`QUAD_*` を**非表示**。
- 空文字の枠（例：示唆3件しかない時の `IMPL_4/5`）は**シェイプごと削除**（空箱が残らない）。

### シェイプ名（`MAP_FRAME` 等）はどこにある？— PowerPoint 内に定義され、選択ウィンドウで確認できる

上表の `MAP_FRAME` / `AXIS_X_LABEL` / `QUAD_TL` / `IMPL_1` … といった名前は、**テンプレ `.pptx`（PowerPoint ファイル）の中に保存されているシェイプ名**。`fill_positioning_map.py` はこの名前を頼りにシェイプを探してテキストを流し込む。

- **実体（XML）**: pptx は ZIP で、スライドは `ppt/slides/slide1.xml`。各シェイプ先頭の
  `<p:cNvPr id="74" name="MAP_FRAME"/>`（cNvPr = common Non-Visual Properties）の **`name` 属性**が名前。
  種類でタグは変わる（図形=`p:sp` / 線=`p:cxnSp` / グループ=`p:grpSp`）が、名前は必ず `cNvPr@name`。
- **PowerPoint での確認・変更場所**: **［ホーム］→［配置（整列）］→［選択ウィンドウ］**（Windows: `Alt+F10`）。
  全シェイプ名が一覧表示され、名前をダブルクリックでリネームできる。グループ化すると、グループ名の下に
  子シェイプ（`MAP_FRAME` 等）が階層表示される。
- **付与しているのは誰か**: 初期生成時に `build_template.py` が `shape.name = "MAP_FRAME"`（→ `cNvPr@name`）で設定。
  fill 側は読み取るだけ。

> **重要な約束**: 選択ウィンドウでの**リネームは禁止**（fill が名前で探すため、変えると流し込み先を見失う）。
> 位置・サイズ・グループ化は自由に変えてよい（2-5-3 参照）。

## 2-2. ファイル別の責務（誰が何を持つか）

| ファイル | 持っている情報 | 持っていない情報 |
|---|---|---|
| テンプレ `.pptx` | 全シェイプの**位置・サイズ・色・フォント・bullet・回転** | 文字内容、バブル |
| `build_template.py` | テンプレを作る**座標テーブル**（`_brand_spec` の STELLAR/ROLEUP dict）＋配色/フォントの解決ロジック | 調査データ |
| `fill_positioning_map.py` | **流し込みロジック**（どの JSON フィールドをどのシェイプへ）＋**バブル描画**＋バリデーション | 座標（テンプレから読む）、色定義（theme から読む） |
| `theme.json`（_common/brands/） | ブランドの**色・フォント・紙面サイズ・許容フォントサイズ** | レイアウト座標、文字内容 |
| データ JSON | **調査結果の全文字＋各社座標** | 見た目（位置・色・font） |
| `layout.json` | （v2 で削除済み。2-7参照） | — |

## 2-3. データ JSON スキーマ（中身を書く人向け）

```jsonc
{
  "main_message": "…すべき",          // 必須・65字以内（hard-fail）
  "chart_title": "競合ポジショニングマップ：○○市場",  // 任意
  "section_title": "ポジショニングマップ",            // 任意
  "implications_title": "ポジショニングからの示唆",    // 任意
  "target_company": "対象会社",        // 任意（指定社のバブルを強調）
  "x_axis": {                          // 必須（label/low/high 必須）
    "label": "製品専門性", "low": "汎用", "high": "専門",
    "min": 0, "max": 10                // 任意（既定 0〜10）
  },
  "y_axis": { "label": "価格帯", "low": "低", "high": "高" },  // 必須
  "quadrants": {                       // 任意（あれば Y軸low/high と排他）
    "top_left": "…", "top_right": "…",
    "bottom_left": "…", "bottom_right": "…"
  },
  "players": [                         // 必須・2〜5件（既定5・上限5。target含む）
    {"name": "対象会社", "x": 4.5, "y": 5.0, "size": 50,
     "color": "#E15759", "label_position": "bottom"}
    // x/y/name 必須、size/color/label_position 任意
  ],
  "implications": ["…", "…"],          // 任意・最大5件（超過は切り捨て＋WARN）
  "source": "出典：…"                  // roleup は必須 / stella は任意
}
```

- **必須キー**（欠けると hard-fail）: `main_message`, `players`, `x_axis`, `y_axis`
- `x_axis`/`y_axis` のネスト必須: `label`, `low`, `high`
- `players[]` の各要素必須: `name`, `x`, `y`
- **社数: 2〜5 件（既定5・上限5、`target_company` 含む）**。単独実行・market-overview-agent 経由とも同一ポリシー。上限超過は hard-fail。`fill_positioning_map.py` の `PLAYERS_MIN/MAX` で検証。
- 検証は `_common/lib/validate_fill_input.py` が実施（想定外キーは stderr WARN）。

## 2-3-1. 軸（x_axis / y_axis）の読み方

`x_axis` は横軸の設定一式。中の4キーは **「表示用の文字3つ」** と **「計算用の数値2つ」** で役割が全く違う。
（`y_axis` も同構造で、縦方向に効く。`low`=下端、`high`=上端）

```
      │              マップ（MAP_FRAME）              │
      │         ○汎用な会社        ○専門な会社        │
      └──────────────────────────────────────────────┘
   ← 汎用                                        専門 →     ← AXIS_X_LOW / AXIS_X_HIGH
                      製品専門性                            ← AXIS_X_LABEL
```

### ① 表示用の文字（label / low / high）— 読み手に軸の意味を伝えるラベル

| JSONキー | 値の例 | 入る枠 | 実際の表示 |
|---|---|---|---|
| `label` | 製品専門性 | `AXIS_X_LABEL`（軸の真下・中央） | 製品専門性 |
| `low` | 汎用 | `AXIS_X_LOW`（左端） | **← 汎用** |
| `high` | 専門 | `AXIS_X_HIGH`（右端） | **専門 →** |

**AXIS_X_LOW の使われ方**：fill は `x_axis.low` を取り出し、**頭に「← 」を自動付与**して `AXIS_X_LOW` 枠へ流す。

```python
_fill_or_hide(slide, "AXIS_X_LOW", f"← {x_low}" if x_low else "")
#  x_low="汎用" → 表示は "← 汎用"
```
- 矢印（←/→）は JSON に書かない（fill が付ける）。意味は「左に行くほど汎用」という注記。
- `low` が空文字ならその枠ごと**非表示**。
- 高い側 `AXIS_X_HIGH` は逆に末尾へ「 →」を付ける（`"専門 →"`）。

### ② 計算用の数値（min / max）— 画面には出ない目盛り

`min`/`max` は**表示されない**。各社バブルの**横位置を計算するためだけ**の目盛り範囲（2-5 の方眼紙の「0〜10」がこれ）。

```python
x_ratio = (player.x - x_min) / (x_max - x_min)   # 0.0〜1.0 の割合に変換
#  min:0,max:10 で x:8 → (8-0)/(10-0)=0.8 → マップ左から80%
```
- 省略時は既定 `0〜10`（多くの調査はこのままでよい）。
- `players[].x` の値はこの `min〜max` の範囲内で付ける（範囲を 0〜100 にすれば各社 x も 0〜100 で書く）。

### まとめ

| | low / high / label | min / max |
|---|---|---|
| 性質 | 見せる文字（ラベル） | 計算用の数値（目盛り） |
| 画面表示 | される | されない |
| 役割 | 軸の意味を読み手に伝える | バブルの横位置を決める |

## 2-4. ブランド差はどう吸収されるか

2ブランド（`stellar_aiz` 16:9 / `roleup` A4横）で**コードは共通**。差分は2か所に閉じ込める。

| 差分の種類 | どこで吸収 |
|---|---|
| 紙面サイズ・パネル幅・各シェイプ座標 | テンプレ `.pptx`（`build_template.py` の `_brand_spec`）|
| 色・フォント・許容フォントサイズ | `theme.json`（roleup）/ V1ハードコード（stella）|
| 上部PHに何を出すか（結論文 or タイトル） | `theme.json` の `placeholder_role_mapping` → `resolve_top_text/subtitle_text` |

**上部プレースホルダの入れ替え**（混乱しやすい点）:
- stella: `Title 1`=`main_message`（結論文を最上部）／`Text Placeholder 2`=`chart_title`
- roleup: `Title 1`=`chart_title`（タイトルを最上部）／`Text Placeholder 2`=`main_message`
- この出し分けは `resolve_top_text()`/`resolve_subtitle_text()`（`format_helpers.py`）が theme から解決。

## 2-5. なぜバブルだけ実行時描画なのか

バブルは「何社・どこに・どの大きさ・何色」が**調査ごとに完全に変わる**ため、テンプレに固定枠を置けない（5枠固定の示唆とは性質が違う）。したがってバブルだけは `fill` が `add_shape(MSO_SHAPE.OVAL)` で実行時に描画する。
その配置基準が次の `MAP_FRAME`。

## 2-5-1. MAP_FRAME とは何か（バブルの土俵）

`MAP_FRAME` は、スライドに置いてある **「グレーの長方形（マップの枠）」そのもの**。役割は2つ。

1. **見た目（背景の箱）**：バブルが乗る、薄いグレー＋黒枠の四角形。テンプレに置いてある普通の図形。
2. **バブルの座標の物差し（本質）**：fill はバブルを「スライドのどこに置くか」を直接は知らない。
   代わりに **MAP_FRAME の位置と大きさを読み、その四角形の中の相対位置で計算**する。

### イメージは「方眼紙」

MAP_FRAME という紙の上で：
- 左端 = x:0 / 右端 = x:10（既定レンジ。`x_axis.min/max` で変更可）
- 下端 = y:0 / 上端 = y:10

データ JSON に `{"name":"テスラ","x":8,"y":8}` とあれば：

```
横 = 8/10 = 80% → 枠の左から 80% の位置
縦 = 8/10 = 80% → 枠の下から 80% の位置（＝上の方）
```

### 数字で見る

MAP_FRAME が「左 1.16in・幅 6.75in」に置いてあるとき、x=8 のバブルの横位置は：

```
1.16in + 6.75in × 0.8 = 6.56in
```

これを式にしたものが：
- `cx = map_x + map_w × x比率`
- `cy = map_y + map_h × (1 − y比率)`  ← 縦は上下反転（y が大きいほど上）

## 2-5-2. MAP_FRAME の座標は「固定値」ではなく毎回テンプレから読む

ここが重要。**コードに `map_x = 1.16` のような固定値は書かれていない。**
fill は実行のたびにテンプレ `.pptx` を開き、`MAP_FRAME` シェイプの実座標を読み直す。

```python
frame = find_shape(slide, "MAP_FRAME")   # テンプレから探す
map_x, map_y = frame.left, frame.top      # 実座標を読む
map_w, map_h = frame.width, frame.height  # → この4値だけがバブル配置の基準
```

| もし固定値だったら | 実際（テンプレを毎回読む） |
|---|---|
| PowerPoint で枠を動かすとコードとズレてバブルが枠外に出る | 枠を動かしても次回実行時に新座標を読むのでバブルが追従 |
| 座標変更にコード修正が必要 | テンプレを直すだけ（コード不変） |

`build_template.py` が持つ座標値は**「テンプレを最初に作るとき」だけ**に使う。いったんテンプレができれば、その後の fill は**毎回テンプレを正として読み直す**。これが「`MAP_FRAME` = single source of truth（座標の唯一の正）」の実体。

> まとめ：**MAP_FRAME は「バブルを描く土俵」であり、fill はその土俵の四隅を見て各社の位置を割り出す。**
> だから「マップの大きさ・位置を変えたい」＝ PowerPoint で `MAP_FRAME` を動かすだけ。再生成もコード修正も不要でバブルが追従する。

## 2-5-3. 人が PowerPoint でグループ化しても動く（グループ対応）

「人が微調整」する際、マップ要素（枠・ガイド・軸・象限）を**まとめて選択してグループ化**し、
一体で移動・リサイズするのは自然な操作。そのため fill は**グループ対応**にしてある。

- **再帰探索**：`find_shape`/`_silent_remove_shape` はグループの中まで再帰的に探す
  （グループ内の `AXIS_*`/`QUAD_*` にもテキストを流せる・空枠を消せる）。
- **絶対座標の解決**：グループ内の `MAP_FRAME` の座標は「グループ内のローカル座標」になるため、
  `_abs_geometry()` がグループの変換（`off/ext/chOff/chExt`）を解いて**スライド上の絶対座標**に直す。
  グループごと動かしても・拡大しても、バブルは正しく枠に追従する。
- バブル自体は（グループの中ではなく）**スライド直下に絶対座標で**描かれる。

> 注意：`MAP_FRAME` という**名前は変えない**こと（fill が名前で探すため）。グループ化・移動・リサイズは自由。

## 2-6. 「○○を変えたい」逆引き表

| やりたいこと | 触る場所 |
|---|---|
| 軸名・象限文言・示唆文・出典を変える | **データ JSON** |
| 各社の位置・大きさ・色を変える | データ JSON の `players[]` |
| 軸自体（何×何で切るか）を変える | データ JSON の `x_axis`/`y_axis`（※軸選択ゲートでユーザー選択） |
| マップ枠や示唆枠の**位置・大きさ**を変える | テンプレ `.pptx` を PowerPoint で動かす（or `build_template.py` 再生成） |
| 色・フォントをブランド全体で変える | `_common/brands/<brand>/theme.json` |
| 示唆の枠数（5→N）を変える | `build_template.py`（枠生成）＋ `fill` の `IMPL_SLOTS` |
| バブルの色ロジック・強調ルールを変える | `fill_positioning_map.py` の `_apply_theme`/`draw_bubbles` |

## 2-7. layout.json は削除済み（一般機構としての解説）

- **本来の役割**：fill が**実行時に図形を描く**スキルで、その座標を**ブランド別に・コード分岐なしで**持たせるための JSON（`brand_resolver` が `theme.layout()` / `theme.layout_in()` で読む）。今も `market-environment-pptx` 等の動的描画スキルでは現役。
- **positioning-map の現状**：v2 で骨組みをテンプレ固定＋バブルは `MAP_FRAME` 参照に切り替え、**fill は layout.json を一切読まなくなった**ため、`assets/<brand>/layout.json` は **削除済み**（座標の正本はテンプレ本体＋`tools/build_positioning_map_template.py` の dict）。`brand_resolver` は layout.json 不在でも `{}` を返すので動作に影響なし。
- 座標を変えたいときは **テンプレの直接微調整**（基本）か `tools/build_positioning_map_template.py` の dict 編集（ゼロ作り直し）。

### layout.json が「正しい道具」になる3条件

次の3つが揃ったときに layout.json を使う：**①実行時に図形を描く（テンプレ固定枠にできない）／②座標がブランドで変わる／③コードを分岐させたくない**。

```python
# ✗ これを避けたい（ブランド分岐がコードに散る）
if brand == "roleup": left_w = 6.80
else:                 left_w = 7.70

# ◎ layout.json に逃がす（コードは共通、値は JSON 差分）
left_w = theme.layout_in("left_w_in")
```

| シナリオ | layout.json が要る理由 |
|---|---|
| バブル/チャート/テーブル等、**個数がデータ依存で固定枠にできない**ものを描く | 描き始める領域の座標をブランド別に持たせたい |
| 同一スキルを**紙面サイズ違いの複数ブランド**で出し分ける | パネル幅・位置の差をコード分岐なしで吸収 |
| 座標を**コードを触らず**チューニングしたい | JSON 1行で位置を動かせる |

### 逆に layout.json を「使わないべき」ケース（＝今回 positioning-map が選んだ方）

枠・軸・象限・タイトルのように **個数も位置も毎回同じ固定の骨組み**は、layout.json で座標を持つより**テンプレ `.pptx` に名前付きシェイプとして焼き込む方が優れる**（WYSIWYG で直せる／fill は文字を流すだけ／ブランド差はテンプレ自身が持つ）。

### positioning-map が layout.json を「再び使う」としたら

| 将来こうするとき | layout.json が復活する |
|---|---|
| 凡例・サイズ目盛り等、**データ依存で動的に描く要素**を追加 | テンプレ固定枠にできずブランド別座標が要る |
| バブル描画領域を `MAP_FRAME` と切り離し**コードから動かす** | テンプレ再生成せず JSON で領域を渡す |
| 新ブランド追加で**動的描画部分の座標だけ**変えたい | テンプレ全再生成より JSON 差分が軽い |

## 2-8. 軸選択ゲート（運用ルール）

ポジショニングの価値は「どの2軸で切るか」に集約されるため、**スキルが軸を勝手に決めない**。
- JSON 生成前に**軸候補3案を根拠付きで提示 → ユーザー選択を待つ**（SKILL.md「軸選択ゲート」）。
- `market-overview-agent` 経由でも省略しない（オーケストレーターが軸を想像で確定しない）。

---

# 付録：処理ステップ（fill_positioning_map.py の流れ）

1. `--brand` から `theme` を解決 → `_apply_theme`（バブル用の色・フォントだけ設定）
2. テンプレ `.pptx` を開く
3. `validate_fill_input` で JSON スキーマ検証（必須キー・想定外キー）
4. `main_message` 65字・`players` 2〜10件・`implications` 5件上限をチェック
5. `Title 1`/`Text Placeholder 2` にテキスト流し込み（ブランドで入替）
6. `fill_text_shapes`：軸・象限・示唆・タイトル・出典を**名前付き枠へ流し込み**、空枠は削除、排他ルール適用
7. `draw_bubbles`：`MAP_FRAME` 座標を基準に各社バブル＋ラベルを**実行時描画**
8. 保存 → `_finalize_pptx`（LibreOffice 整形で修復ダイアログ回避）

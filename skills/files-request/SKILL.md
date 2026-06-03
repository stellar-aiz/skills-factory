---
name: files-request
description: >
  BDD（ビジネスデューデリジェンス）の起点であるデータリクエスト（IRL = Information Request
  List）の Excel を生成するスキル。ビジネスモデル分類（stores / assets / labors / products /
  flows）を引数で受け取り、PJ ルート配下のナレッジフォルダにあるビジネスモデル別ベース IRL
  をコピーして、PJ 用のリクエストファイル（ファイル名と Excel 内の PJ 名称を書き換えたもの）を
  作成する。

  やることはシンプル：ベース Excel をコピーし、ファイル名と中の PJ 名を書き換えるだけ。
  項目の再整形・取捨選択・管理列の付与などは行わない（忠実コピー + PJ 名置換のみ）。

  以下のいずれかのトリガーで必ずこのスキルを使うこと：
  - 「IRL」「データリクエスト」「リクエストファイル」「資料依頼リスト」という言葉が出た場合
  - 「BDD のデータ依頼」「依頼資料リストを作って」という要望
  - 「ビジネスモデル（stores / assets / labors / products / flows）の IRL を作って」という要望

  実行前提：このスキルは PJ のルートフォルダ内で実行される（cwd = PJ_root）。
  入出力はすべて cwd 相対の `3_KnowledgeSource/` / `4_Task/` を使う。
---

# files-request: ビジネスモデル別データリクエスト（IRL）生成

## 役割・ミッション

世界最高の戦略コンサルタントとして、ビジネスデューデリジェンス（BDD）の最初に行う
**データリクエスト（IRL）を抜け漏れなく** 行うベテランのスキルを担う。

このスキルの実体は単純で、ビジネスモデルごとに定義されたベース IRL から、当該 PJ の IRL を
**コピー + 書き換え** で作成するだけである。

## 引数とビジネスモデル定義

引数 `business-model` は以下の 5 分類に**限定**される（いずれか必須）:

- `stores`（店舗・拠点型）
- `assets`（資産型）
- `labors`（労働型）
- `products`（製品型）
- `flows`（仲介・流通型）

各分類の価値源泉・代表業種・ベース IRL ファイル名、および未定義時の扱い（禁止事項）は、
**唯一の正本である `references/business_model_taxonomy.md`** を参照。

## 入出力仕様

すべて cwd（= PJ ルート）相対。

- **入力**: `./3_KnowledgeSource/questions-for-<business-model>.xlsx`（PJ 側に既存。読むだけ）
- **出力**: `./4_Task/<yyyymmdd>-RequestFilesv<version>/<PJ名>_RequestFiles_v<version>.xlsx`
  - `<yyyymmdd>`: 実行日（既定は本日）
  - `<version>`: 既定 1
  - PJ 名は既定で cwd フォルダ名。相違があればユーザーに確認する。

## 処理の流れ

1. **cwd 確認** — スキルは PJ ルートフォルダ内で実行されている前提。`pwd` で確認する。
2. **business-model の検証** — 引数を取得し 5 分類に含まれるか確認。含まれなければ、有効な
   5 つを提示して停止（禁止事項）。
3. **PJ 名の確定** — 既定は cwd フォルダ名。ユーザーに確認し、違えば指定された名前を使う。
4. **ベース IRL の存在確認** — `./3_KnowledgeSource/questions-for-<bm>.xlsx` を確認。
   無ければ報告して停止。
5. **（任意）置換語の検出** — ベース Excel をざっと開き、PJ 名が入るプレースホルダ
   （例 `[PJ名]` / `〇〇社` 等のテンプレ表記）が何かを把握し、`--placeholder` に渡す。
   置換対象が無ければ `--placeholder` を省略して純粋コピーにする。
6. **スクリプト実行** — `make_request_files.py` でコピー + 置換 + リネーム。
7. **報告** — 出力パスと置換セル数をユーザーに伝える。

## スクリプト実行コマンド

```bash
pip install openpyxl {{PIP_FLAGS}}

{{PYTHON_BIN}} {{SKILL_DIR}}/scripts/make_request_files.py \
  --business-model stores \
  --pj-name "対象会社名" \
  --placeholder "[PJ名]" \
  --version 1
```

オプション:
- `--placeholder "<文字列>"`: Excel 内で PJ 名に置換するプレースホルダ。省略時は置換せず純粋コピー。
- `--version N`: リクエストファイルのバージョン（既定 1）。
- `--date YYYYMMDD`: 出力フォルダの日付（既定は本日）。
- `--root <path>`: PJ ルート（既定は cwd）。

スクリプトの動作:
- business-model を 5 分類で検証（不正なら stderr + exit 1）。
- 入力 `./3_KnowledgeSource/questions-for-<bm>.xlsx` の存在確認（無ければ exit 1）。
- 出力フォルダ `./4_Task/<yyyymmdd>-RequestFilesv<version>/` を作成。
- `shutil.copy2` でベース Excel をコピー（書式・シート構成をそのまま保持）。
- コピー先を openpyxl で開き、全シートの全セルから `--placeholder` を `--pj-name` に置換して保存。
- コピー先を `<PJ名>_RequestFiles_v<version>.xlsx` にリネーム。

## アセット・参考

| パス | 内容 |
|------|------|
| `scripts/make_request_files.py` | コピー + PJ 名置換 + リネーム |
| `references/business_model_taxonomy.md` | 5 分類定義・ファイル名規約・未定義時の扱い |

## 設計思想

1. ベース IRL は読むだけ。コピー + PJ 名置換に徹し、再整形・取捨選択はしない。
2. ビジネスモデルは 5 分類限定。未定義 / ベース欠如は捏造せず明示停止（禁止事項）。
3. `bdd-` プレフィックスを付けない（誤起動抑止ルール対象外、通常 install 可）。

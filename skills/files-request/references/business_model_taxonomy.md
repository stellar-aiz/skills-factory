# ビジネスモデル分類（files-request 正本）

`files-request` スキルが扱うビジネスモデルは以下の **5 分類に限定**する。
引数 `business-model` はこのいずれかでなければならない。

| id | 名称 | 価値源泉（一言で） | 代表業種 | ベース IRL ファイル名 |
|----|------|--------------------|----------|------------------------|
| `stores` | 店舗・拠点型 | 立地と店舗網が価値源泉 | 小売、飲食、サービス店舗 | `questions-for-stores.xlsx` |
| `assets` | 資産型 | 設備・資本が価値源泉 | 製造、不動産、レンタル | `questions-for-assets.xlsx` |
| `labors` | 労働型 | 人の稼働・専門性が価値源泉 | 建設、人材、受託サービス | `questions-for-labors.xlsx` |
| `products` | 製品型 | 製品・知財が価値源泉 | メーカー、ブランド、ソフトウェア | `questions-for-products.xlsx` |
| `flows` | 仲介・流通型 | 仕入れて売るスプレッドが価値源泉 | 卸、商社、代理店、EC 物販 | `questions-for-flows.xlsx` |

## ファイル名規約

ベース IRL は PJ ルート配下の `3_KnowledgeSource/questions-for-<id>.xlsx` に置かれている前提。
スキルはこれを **読むだけ**（中身の設計・メンテはスキルの責務外）。

## 未定義時の扱い（禁止事項）

- 引数 `business-model` が上記 5 つ以外の場合 → **定義されていない旨を明示して停止**する。
  あたかも定義されているかのように振る舞わない。有効な 5 つを提示する。
- 対応するベース IRL（`questions-for-<id>.xlsx`）が `3_KnowledgeSource/` に存在しない場合 →
  **ファイルが無い旨を報告して停止**する。空ファイルや代替の生成はしない。

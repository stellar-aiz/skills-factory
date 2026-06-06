# Step 1: 論点別 Web 検索キーワード テンプレ

各論点について、以下のキーワードパターンで WebSearch を実行する。`{market_name}` は scope.json から、`{geography}` も同様に展開する。

各論点 5〜8 コール程度を目安（v0.1の既定値・深度議論は v0.2 以降）。

---

## 論点1: 市場規模・成長率

```
"{market_name} 市場規模 {YYYY}"
"{market_name} 市場規模 推移 過去{N}年"
"{market_name} 成長率 CAGR"
"{market_name} 矢野経済"
"{market_name} 富士経済"
"{market_name} 市場予測 {YYYY}"
"{market_name} market size {YYYY}"   # geography=グローバル/北米欧州 の場合
"{market_name} market forecast"
```

優先ソース: 矢野経済研究所 / 富士経済 / IDC / Gartner / 業界団体 / 政府統計（経済産業省・総務省）

出力: `data_04_market_environment.json`

---

## 論点2: KBF（Key Business Factor）

```
"{market_name} 成功要因"
"{market_name} 競争優位"
"{market_name} 重要成功要因"
"{market_name} key success factor"
"{market_name} prevailing strategy"
"{market_name} 業界レポート 動向"
"{market_name} デロイト" "{market_name} PwC" "{market_name} BCG"
"{market_name} 専門家 インタビュー"
```

優先ソース: 業界紙 / コンサルレポート / 専門家インタビュー記事 / 各社IR の中計

抽出物: KBF×3 候補と、その根拠（なぜ重要か）と、各プレイヤーが実装している例

出力: `data_10_market_kbf.json`

---

## 論点3: 各社の市場シェア

```
"{market_name} シェア {YYYY}"
"{market_name} シェアランキング"
"{market_name} 競合 売上"
"{プレイヤー名} 売上高 セグメント別"
"{プレイヤー名} 統合報告書 {YYYY}"
"{プレイヤー名} 中期経営計画"
```

優先ソース: 各社IR（年次報告書・決算短信）/ 業界統計 / 経済紙

シェア計算は「明示されていれば公開数値、なければ売上高から推計」と明示。推計の場合は `notes` に記載。

出力: `data_06_market_share.json`

---

## 論点4: ポジショニング（プレイヤー位置付け）

```
"{プレイヤー名} 強み 戦略"
"{プレイヤー名} 価格帯 サービス領域"
"{プレイヤー名} ターゲット顧客"
"{プレイヤー名} 製品カテゴリ"
"{market_name} 競合マップ"
"{market_name} ポジショニング"
```

軸の決め方（**軸を勝手に確定しない＝必須ゲート**）:
- ポジショニングマップの価値は「どの 2 軸で切るか」に集約される。**オーケストレーターが軸を
  想像で確定してはならない。** positioning-map-pptx の「軸選択ゲート（必須・スキップ不可）」を必ず通す。
- 手順: (1) 軸候補を **3 案** 根拠付きで生成しユーザーに提示（X軸/Y軸 名称・low/high の意味・
  なぜプレイヤーが分散するかの根拠 1〜2 行）→ (2) **ユーザーの選択を待つ** → (3) 選択された軸でのみ
  `data_07_positioning.json` の `x_axis`/`y_axis` を確定する。
- 候補ソース例: X軸=顧客セグメント・地域範囲・価格帯・ターゲットセグメント / Y軸=機能スコープ・
  成長性・利益率・ブランドポジション 等（プレイヤーが分散する軸を選ぶ）。

優先ソース: 各社HP / IR / プレスリリース / 業界レポート

出力: `data_07_positioning.json`（軸はユーザー選択済みのもののみ）

---

## 論点5: 各社の戦略比較

```
"{プレイヤー名} 中期経営計画"
"{プレイヤー名} 統合報告書"
"{プレイヤー名} 決算説明会 {YYYY}"
"{プレイヤー名} 戦略 重点施策"
"{プレイヤー名} M&A 投資"
"{プレイヤー名} 海外展開"
```

比較軸（competitor-summary-pptx の comparison_items）:
- 事業内容 / 本社所在地 / 設立年 / 従業員数 / 売上高 / 上場区分 / 強み・差別化（標準7項目）
- 市場分析特化なら以下に置き換え可: 主力製品 / 主要顧客セグメント / 売上高 / 直近成長率 / 戦略の重点 / グローバル展開度 / 強み

出力: `data_08_competitor_summary.json`

---

## 論点（追加）: PEST環境

```
"{market_name} 規制 動向"
"{market_name} 政府 政策"
"{market_name} 法改正"
"{market_name} 経済 マクロ環境"
"{market_name} 社会 トレンド"
"{market_name} 技術 革新"
"{market_name} {DX/AI/IoT 等の関連キーワード}"
```

各象限で 3〜5 項目を抽出。影響度（▲追い風 / ▬中立 / ▼逆風）も付与する。

優先ソース: 政府統計 / シンクタンク / メディア / 業界団体

出力: `data_11_pest.json`

---

## 検索結果の品質基準

各論点で取得した情報には、以下を必ず JSON 内に記録する：

```json
{
  "source": "矢野経済研究所「HR Tech市場に関する調査2025」",
  "source_url": "https://www.yano.co.jp/...",
  "fetched_at": "2026-04-26",
  "confidence": "high|medium|low",
  "notes": "シェアは売上高ベースの推計。原典は競合上位5社の合算売上÷市場規模"
}
```

confidence の判定:
- **high**: 一次ソース（IR・政府統計）または複数ソースで一致
- **medium**: 単一の信頼できる二次ソース（業界レポート・経済紙）
- **low**: 推計値・古い情報・単一ソース

`confidence=low` の項目は **Step 2 Data Availability で △ ステータス** とし、Step 9 の
FactCheck_Report.md でも `data_gaps` に明示する。

---

## ユーザーアップロード情報の優先

ユーザーが業界レポート・有報・HP・社内資料などをアップロードした場合は、Web検索より**先**に
そのコンテンツを読み取り、以下のように扱う：

- ユーザー提供は `confidence=high` 起点（出典が明確なため）
- 数値や年度は Step 2.5 で fact-check-reviewer が裏取り
- 矛盾発生時は ユーザー提供を優先するが、Web 検索で確認した数値も併記し、severity=high として
  Step 3 でユーザー判断を仰ぐ

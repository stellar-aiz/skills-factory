# skills_factory プロジェクト懸案管理

このファイルは skills_factory プロジェクトで判断保留した事項・将来検討事項をセッション横断で蓄積するためのものです。

**運用ルール**:
- セッション開始時に必ず本ファイルを読み込み、未解決事項を確認する（プロジェクト直下 CLAUDE.md からも参照）
- 新規イシューは末尾に追記。Status は `保留` / `進行中` / `クローズ` のいずれか
- クローズしたイシューも履歴として残す（削除しない）

---

## ISSUE-001: build_skill.py への `@import` 機構導入

**Status**: 保留 / **Priority**: P3 / **Decided**: 2026-04-27 / **Updated**: 2026-05-02 (Phase B-4/B-5 で `_common/` ファイル数が 9 に到達し**起票トリガー条件 1 達成**。ただし同期漏れ 0/3 のため起票判断はユーザーへ持ち込み)

### 背景
v0.2 Phase D で 3 orchestrator（market-overview-agent / strategy-report-agent / smallcap-strategy-research）の重複ブロックを `skills/_common/` に集約する作業を、**手動コピペ運用（D1）** で開始した。`build_skill.py` への `@import` 自動化（D2）は、今セッション（2026-04-27）で **判断保留** に決定。

### D2 を保留した理由（後で変更しづらい設計判断 5 つ）
1. **パスの基準** — `@import` パスをソースファイル相対 / `skills/` ルート / リポジトリルートのどれを基準に解決するか。最初に決めた基準を後で変えると、既に書いた `@import` 文を全部書き換えになる。
2. **import と {{VAR}} 置換の合成順序** — imported ファイル内の `{{VAR}}` を import 前/後どちらで解決するか。imported ファイル内の `{{VAR}}` が新たな `@import` を生成するケースの可否。
3. **循環インポート検出** — `a → b → a` を検出して停止するロジックの追加。現状の build_skill.py には依存関係追跡機構がない。
4. **キャッシュ無効化** — `_common/*` を変更したとき、被参照スキルを自動再ビルドする仕組み。現状の `build_skill.py install <name>` は単一スキルのみ対象なので、`_common/` 変更時に手動で全部 install しないと反映されない。
5. **エラーメッセージの追跡可能性** — ネスト import で未解決変数が出たときに、どのファイル経由で持ち込まれたかを辿れるか。

### v0.3 で D2 着手するトリガー（どちらかを満たしたら起票）
- ✅ `skills/_common/` 配下のファイル数が **8 ファイル以上** に膨らむ → **2026-05-02 達成**(9 ファイル: cross_topic_consistency_check / main_message_principles / step_final_visual_review_loop / step_state_tracking / step0_scope_clarification / step2_5_factcheck_invocation / harness_levers / orchestrator_contract / chart_palette)
- ⏳ 手動コピペ運用で **3 回以上の同期漏れインシデント** → 0/3 のまま

### 2026-05-02 の判断
ファイル数トリガーは達成したが、同期漏れインシデントが発生していないため運用上の痛みは未顕在化。harness_check ブランチでも「`<!-- source: skills/_common/... (manual sync until D2) -->` コメント + 手動コピペ」が機能しており、9 ファイル目（`harness_levers.md` / `step_state_tracking.md`）の追加でも同期漏れは生じなかった。**起票するか継続保留かはユーザー判断**(harness_check のレビュー時に合わせて確定)。

### 2026-05-03 の判断（β + γ E2E 完了後）

β（business-deepdive-agent / 二幸産業）+ γ（market-overview-agent / 国内タクシー市場 Step 1 限定）E2E でも、**`_common/` 手動コピペ運用は同期漏れインシデント 0 件** で安定運用されている。ファイル数 9 のままで Phase B クローズ。

**判断: 継続保留**(`Status: 保留`)。トリガー条件 2「同期漏れ 3 回」は依然 0/3 のまま。v0.4 以降で `_common/` ファイル数が 12 を超える、もしくは新規 orchestrator 追加時に同期漏れが発生したら起票検討。

### 参考ファイル
- `tools/build_skill.py`(現状: {{VAR}} 3パス置換 + @if/@endif 実装)
- `skills/_common/`(v0.2 D1 で新設、9 ファイル / 手動運用)
- `/Users/nakamaru/.claude/plans/4-market-overview-v0.2.md` P3-9 セクション
- `/Users/nakamaru/.claude/plans/fancy-cooking-walrus.md` Phase D1 セクション
- `/Users/nakamaru/.claude/plans/md-llm-melodic-twilight.md` harness_check 計画書（2026-05-02 で Phase B-6 まで完了、検証フェーズ残）

---

## ISSUE-002: Web 検索深度の動的制御

**Status**: 保留 / **Priority**: P3 / **Decided**: 2026-04-27 / **Updated**: 2026-04-29 (再確認、保留継続: 現状の Step 2.5 ユーザー判断フローで十分機能)

### 背景
v0.1 で「次フェーズで議論」と保留した項目。現状は market-overview-agent / strategy-report-agent で論点別に Web 検索コール数を 5〜8 で固定している。

### 検討内容
fact-check-reviewer の severity が `high` の論点については追加で Web 検索コールを発射する動的拡張を入れるか。コスト・所要時間とのトレードオフをどう設計するか。

### 参考ファイル
- `skills/market-overview-agent/SKILL.md` Step 2 の Web 検索セクション
- `skills/fact-check-reviewer/SKILL.md`
- `/Users/nakamaru/.claude/plans/4-market-overview-v0.2.md` 「v0.2 で扱わないこと」セクション

---

## ISSUE-003: AI による自動 main_message 短縮

**Status**: 保留 / **Priority**: P3 / **Decided**: 2026-04-27 / **Updated**: 2026-04-29 (再確認、保留継続: ルール強制で十分機能)

### 背景
v0.2 Phase B で `skills/_common/prompts/main_message_principles.md` を整備し、LLM 出力をルール強制（4 原則）で 65字以内に収める方針を採用。一方で、超過時に AI が自動で短縮するヘルパースキルを別途作る選択肢もある。

### 検討内容
- 現状: ルール強制（プロンプトで 4 原則を厳守させる + fill_*.py 入口で hard-fail）
- 代替: 短縮専用の補助スキル（main_message を入力 → 65字以内の候補を 3 案返す）
- どちらが運用負荷が低いか、LLM 呼び出しコストとの兼ね合いで再検討

### 参考ファイル
- `skills/_common/prompts/main_message_principles.md`(v0.2 Phase B で作成予定)
- `/Users/nakamaru/.claude/plans/4-market-overview-v0.2.md` 「v0.2 で扱わないこと」セクション

---

## ISSUE-004: 会社・事業 深掘りエージェント群の新規実装

**Status**: クローズ（Phase 4 完了、Phase 5 は不要判断で実装保留） / **Priority**: P2 / **Decided**: 2026-04-27 / **Re-scoped**: 2026-04-29 / **Closed**: 2026-04-29

### 背景
当初は「`strategy-report-agent v5.1` を Company Overview Agent としてリネーム」案が有力とされていたが、ユーザーとの議論の結果、本タスクは**4 つのスキルを新規実装する**方向に再スコープされた。

ユーザーの実用想定は「市場 → 各社の戦略を横並び比較」のフロー:
1. `market-overview-agent` でプレイヤーを把握
2. 各社（A/B/C 社）について深掘り（**会社レベル + 事業セグメント単位**で別エージェント）
3. 全社・全事業を横並びで比較

「戦略が透けて見える」深さを実現するため、論点を**会社レベル 5 つ・事業レベル 5 つ**に整理。

### 新規作成スキル（4 件）
1. ✅ `business-overview-pptx` — 事業の概要 1 枚スライド（PPTX 単体）— **Phase 2 完了**
2. ✅ `company-deepdive-agent` — 会社レベル 5 論点 + セグメント検出 + 事業深掘り N 回呼ぶ + merge — **Phase 4 完了**
3. ✅ `business-deepdive-agent` — 事業セグメント単位 5 論点 — **Phase 3 完了**
4. ⏸ `comparison-synthesis-agent` — 全社比較 + 検証論点統合 — **skeleton 温存、実装保留**（2026-04-29 判断）

### Phase 進捗

| Phase | 内容 | 状態 | 検証 |
|---|---|---|---|
| Phase 1 | 4 skeleton 作成 | ✅ 完了 | 7f517a8 |
| Phase 2 | business-overview-pptx 実装 | ✅ 完了 | a3a96f4、二幸産業タクシー事業 1 枚生成成功 |
| Phase 3 | business-deepdive-agent E2E | ✅ 完了 | a3a96f4、二幸産業 施設運営事業で 5 PPTX 生成成功 |
| Phase 4 | company-deepdive-agent E2E | ✅ 完了 | 91237ab、二幸産業 N=1 で 15 枚デッキ生成成功（merge_warnings 0 件、visual-review pass）|
| **Phase 4 後追補** | **論点間整合性ルール導入** | ✅ 完了 | 671dc95、`_common/cross_topic_consistency_check.md` 新設 |
| Phase 5 | comparison-synthesis-agent 実装 | ⏸ 実装保留 | 2026-04-29 判断、後述 |

### Phase 4 後追補（2026-04-29）の経緯
Phase 3 E2E で「P14 顧客市場の取り違え」を発見。論点 5（顧客市場）に二幸産業自身のビルメンテ市場を入れたが、設計上は顧客（不動産業）が属する市場を扱うべきだった。既存の品質ゲート（main_message ルール / fill_*.py validate / visual-quality-reviewer）はすべて単一スライドの自己整合性しか見ず、論点間の意味的整合性検査機構が欠如していたため、**共通プロンプト `skills/_common/prompts/cross_topic_consistency_check.md`** を新設して構造的に防止する体制を導入。`business-deepdive-agent` のみ適用、`company-deepdive-agent` は別形のため対象外（ユーザー確定）。

### スコープ外（明示的に切り離し）
- `smallcap-strategy-research` および 5 つの `smallcap-*-pptx` スキル（4/27 retrospective でユーザー評価「うまくいっていない」、別物として実装）
- `strategy-report-agent` v5.1（並走、本タスクでは触らない）

### 上場・非上場の境界
v0.3 では境界判定を入れず、**取れる公開情報から最大限作成**。取れなかった項目は `data-availability-pptx` で「✗未取得」明示。Phase 3-4 の二幸産業（非上場）E2E で実証済。

### Phase 5 不要判断（2026-04-29）

Phase 4 完了時点で本スキル（`comparison-synthesis-agent`）の独自価値を再評価した結果、**実装保留**で合意。

#### 既存スキルでカバー済の領域（80% 以上）

| 観点 | 既存スキルでの代替 |
|---|---|
| 全社の財務横並び比較 | `financial-benchmark-pptx`（売上・営業利益率・ROE 等の小型バーチャート群、2×3 グリッド） |
| 全社の事業内容・強み比較 | `competitor-summary-pptx`（横型比較テーブル、3-10 社、対象会社ハイライト可） |
| 市場全体での各社ポジション | `market-overview-agent`（market-share-pptx + positioning-map-pptx + market-kbf-pptx の組み合わせ） |
| 業界 KBF と各社の押さえ方 | `market-kbf-pptx`（KBF 3 つ × 主要プレイヤーの実装例） |
| 事業セグメント単位の深掘り | `business-deepdive-agent`（Phase 3 で完成） |
| 会社単位のフルデッキ | `company-deepdive-agent`（Phase 4 で完成、15-30 枚） |

#### Phase 5 の独自価値（残り 20%）

- 複数社 deepdive デッキの論点間統合（A 社 vs B 社 vs C 社の事業 deepdive 比較）
- 全社共通の検証論点（`open_questions`）の業界横断集約

#### 判断

実利用シーンでは「BDD（1 社評価）」「市場分析」が中心で、3-5 社統合 deepdive の需要が顕在化しないため、skeleton のまま温存して**需要発生時に実装着手**する方針に変更。

skeleton ファイル（`skills/comparison-synthesis-agent/SKILL.md` 237 行）は main にコミット済で、冒頭に「実装保留中」注記を追加（671dc95 以降のコミットで反映）。`competitor-summary-pptx/SKILL.md` の `competitor-analysis-pptx` 参照書き換えも保留。

復活時の参照先:
- `skills/comparison-synthesis-agent/SKILL.md` の冒頭注記
- `skills/_common/prompts/cross_topic_consistency_check.md` の「適用範囲」表（保留中マーク）
- `~/.claude/plans/tidy-soaring-elephant.md`（Phase 5 設計）

### クローズ理由

ISSUE-004 の当初目標「会社・事業 深掘りエージェント群」は Phase 1〜4 と Phase 4 後追補で本質的に達成済。Phase 5（横並び比較統合）は既存スキルの組み合わせで 80% 代替可能であり、独自実装の優先度は低いと判断。

### 参考ファイル
- 計画書: `/Users/nakamaru/.claude/plans/tidy-soaring-elephant.md`（v0.3 メイン）
- Phase 4 後追補 plan: `/Users/nakamaru/.claude/plans/phase3-fancy-dove.md`（最新版に上書き）
- 4/27 smallcap retrospective（参考、使わない）: `/Users/nakamaru/.claude/plans/claude-plans-smallcap-strategy-research-sleepy-koala.md`
- 仕様メモ: `docs/issue-004-spec.md`（Phase 1 のユーザー指示録）

---

## ISSUE-005: 市場スコープの事業モデル境界確認（Phase F 主タスク）

**Status**: クローズ（F-1〜F-5 完了） / **Priority**: P1 / **Decided**: 2026-04-27 / **Updated**: 2026-04-28 / **Closed**: 2026-04-28

### 背景
v0.2 Phase E（国内タクシー市場 E2E）で、handoff の想定プレイヤー欄に従い「タクシー事業者（5社）」と「配車アプリ事業者（4社）」を混在させてレポート化したところ、ユーザーから「タクシー事業者の市場を見たかった、アプリは除外」との指摘を受けた。シェア表で第一交通産業 4.7% と GO 1.2% を同列に並べる構図は、収益構造（営業収入 vs 配車手数料）が異なるため誤解を招く。

現状の `market-overview-agent` Step 0 は地理スコープ・セグメント粒度・分析年数・max_competitors・kbf_count のみを聞き、**同一業界内の異なる事業モデルを含めるかの確認質問が存在しない**。

### 採用アプローチ（B+C ハイブリッド、ユーザー承認済 2026-04-27）

**B**: Step 0.5（事前スコーピング Web 検索）追加
- `market_name` 確定後 Step 1 の前に「市場構造ザックリ把握」用 Web 検索 1-2 件を走らせ、事業モデルの heterogeneity を検知したらユーザーに再確認

**C**: `scope.json` schema 拡張
- `included_business_models[]` / `excluded_segments[]` を必須フィールドとして追加
- `step0_scope_clarification.md` に必須質問を追加し永続化
- 後続スライドはこの境界を尊重する責務をオーケストレーターに置く

### Phase F 実装タスク
1. ✅ `step0_scope_clarification.md` に Step 0.5 節 + 新フィールド追加（2026-04-28 完了）
2. ✅ `market-overview-agent/SKILL.md` Step 0 を更新（2026-04-28 完了）
3. ✅ `strategy-report-agent/SKILL.md` Step 0 を同様更新（2026-04-28 完了）
4. ✅ `smallcap-strategy-research/SKILL.md` には適用範囲注記のみ追加（2026-04-28 完了）
5. ✅ `orchestrator_contract.md` にセクション 4「scope.json の責務分担」追記、チェックリストに 4 項目追加（2026-04-28 完了）
6. ✅ E2E リラン（国内タクシー市場・事業者のみ 5 社）完了（2026-04-28）。シェア表・ポジショニング・competitor summary すべて事業者ベースで構成され、配車アプリは scope.json で明示的除外。検証レポート: `outputs/taxi_industry_operators_2026-04-28/v0.3_e2e_report.md`

### 参考ファイル
- `skills/_common/prompts/step0_scope_clarification.md`
- `skills/market-overview-agent/SKILL.md` Step 0
- `skills/_common/references/orchestrator_contract.md`
- 関連 memory: `feedback_market_scope_business_model_boundary.md`

---

## ISSUE-006: render_pptx.py CLI と SKILL.md の引数不整合

**Status**: クローズ（2026-04-29 修正） / **Priority**: P2 / **Decided**: 2026-04-27 / **Closed**: 2026-04-29

### 背景
v0.2 Phase E E2E で `visual-quality-reviewer/scripts/render_pptx.py` を起動した際、SKILL.md の説明（`--merge-order` / `--data-dir` を受ける）と実際の CLI（`--pptx --out-dir --dpi` のみ）が一致しないことが判明。

### 検討事項
- SKILL.md と実装のどちらを正とするか
- 自動修正ループ（visual-quality-reviewer が下流で merge_order / data_dir を読む）の実現可否

### 修正方針候補
- (a) `render_pptx.py` に `--merge-order` / `--data-dir` を追加し、`context.json` を別途出力する
- (b) SKILL.md を `--pptx --out-dir` のみ受ける仕様に修正し、merge_order/data_dir はオーケストレーターが LLM 経由で渡す運用とする

### 参考ファイル
- `skills/visual-quality-reviewer/scripts/render_pptx.py`
- `skills/visual-quality-reviewer/SKILL.md`

### 修正内容（2026-04-29）
方針 (b) を採用: ドキュメント側を実装に合わせて修正、`render_pptx.py` / `collect_context.py` の実装は変更しない。

- `skills/visual-quality-reviewer/SKILL.md` の入力仕様表を更新: `merge_order` / `data_dir` の渡し先スクリプトを `collect_context.py` と明記、`render_pptx.py` は `--pptx --out-dir --dpi` のみ受け取る旨を明示
- スクリプト構成（render_pptx.py = PNG 化専用、collect_context.py = context.json 構築専用）を入力仕様表の下に追記
- オーケストレーター側（market-overview-agent / strategy-report-agent / company-deepdive-agent）は既に正しく実装されているため変更なし

---

## ISSUE-007: market-environment-pptx の bars/line Y軸スケール乖離

**Status**: クローズ（2026-04-28 修正） / **Priority**: P3 / **Decided**: 2026-04-27 / **Closed**: 2026-04-28

### 背景
v0.2 Phase E E2E のスライド 4（市場規模推移）で、棒グラフ（営業収入 1.45-1.93兆円）と折れ線（2019年比回復率 75-99%）が同じ Y 軸を共有するため、棒の差分が視認しづらい。Y軸 0-120 で線は明瞭だが棒は底辺に張り付く。

### 検討事項
- データ単位を 兆円→千億円 に変換（推奨、データ側で対応可能）
- テンプレート側で dual-axis（左軸=兆円、右軸=%）対応を入れるか
- `unit_label` / `total_label` の自動推奨ロジックを fill_market_environment.py に入れるか

### 参考ファイル
- `skills/market-environment-pptx/scripts/fill_market_environment.py`
- `skills/market-environment-pptx/assets/market-environment-template.pptx`

### 修正内容（2026-04-28）
F-5 E2E でユーザー要望を受けて根本対応。`fill_market_environment.py` に以下 3 点を導入:

1. **Y 軸自動レンジ計算ロジック** (`_calc_primary_axis_max` / `_calc_secondary_axis_max` / `_round_up_nice`)
   - 棒グラフの左 Y 軸: 積み上げ合計 ×1.15 を nice round（1.0, 1.5, 2.0, 2.5, ...）で自動決定
   - 折れ線の右 Y 軸: 値域が 0-100 内なら 120、それ以外は最大値 ×1.15 を自動丸め
   - JSON で `chart.primary_y_axis_max` / `chart.secondary_y_axis_max` を任意指定可能

2. **二次軸の可視化**: `c:delete val="0"` で右側に折れ線用の Y 軸を表示。`chart.line.num_format` で書式指定可

3. **`plotArea` 子要素順序の OOXML schema 正規化** (`_reorder_plot_area`):
   chart 要素 (barChart/lineChart) → axes (catAx/valAx) の順に並び替え。
   従来は混在順序で LibreOffice が dual-axis を解釈失敗、結果として bars が同一軸に潰されて非表示になっていた。

検証: F-5 E2E のスライド 4 で棒グラフ（左軸 0-2.5 兆円）と折れ線（右軸 0-120%）の両方が明瞭に可視化されることを目視確認。

---

## ISSUE-008: competitor-summary 30字 cell 制限が 9 列構成で厳しい

**Status**: クローズ（2026-04-29 修正、`max_competitors` を 5 に戻す根本対応） / **Priority**: P3 / **Decided**: 2026-04-27 / **Closed**: 2026-04-29

### 背景
v0.2 Phase A で `max_competitors` を 5→10 に拡張した結果、9 列構成（target+8）でフォントが 9pt まで自動縮小されるが、`事業内容` / `強み・差別化` の cell 30字制限が運用上厳しい。E2E で「タクシー配車アプリ（DiDi Global＋ソフトバンク合弁）」（31字）等が hard-fail で何度も書き直しが発生。

### 当初の検討事項（採用せず）
- cell 制限を競合数に応じて動的化（5社=40字、8-10社=30字 等）
- フォント自動縮小と cell 文字数の連動を見直し
- 事業内容と強み・差別化を改行可とし 2 段表示にする

### 修正内容（2026-04-29、ユーザー判断）
当初の動的化案ではなく、**根本対応として `max_competitors` を 10 から 5 に戻す**方向で確定。v0.2 Phase A で拡張した範囲を撤回し、9 列構成の運用ストレスを解消。

修正対象:
- `skills/market-overview-agent/references/deck_skeleton_standard.json`: `limits.max_competitors.max` を 10→5
- `skills/competitor-summary-pptx/scripts/fill_competitor_summary.py`:
  - `COMPETITORS_MAX = 10` → `5`
  - `get_font_sizes()` の 7-10 社向け elif ブロック 4 つを削除し 6 社対応（target+5）まで簡素化、フォントサイズも 14/13/12pt 系に再調整
- `skills/competitor-summary-pptx/SKILL.md`: 「競合 3〜10 社」→「競合 3〜5 社」、フォント表を新しい get_font_sizes に合わせて 14/13/12pt 系で更新、撤回経緯を注記
- `skills/market-overview-agent/SKILL.md`: Step 0 の質問選択肢「3 / 5 / 7 / 10」→「3 / 5」、注意事項の「既定 5、最大 10」→「既定 5、上限 5」

### 参考ファイル
- `skills/competitor-summary-pptx/scripts/fill_competitor_summary.py`
- `skills/competitor-summary-pptx/SKILL.md`
- `skills/market-overview-agent/references/deck_skeleton_standard.json`

---

## ISSUE-009: research-subagent return value の規約遵守不完全（wrapper bloat）

**Status**: クローズ（改善策 (B) 実装完了、E2E 再計測はユーザー判断で省略） / **Priority**: P2 / **Decided**: 2026-05-03 / **Updated**: 2026-05-05 / **Closed**: 2026-05-05 / **Discovery**: γ E2E (market-overview-agent / 国内タクシー市場)

### 背景
β 検証（2026-05-02）で、`research-subagent` の return value に「Based on my comprehensive web research...」等の前置き説明文・マークダウン code fence・末尾 `Sources:` トレーリングが混入する事象を 5/5 件で観測。

これを踏まえ Phase B 総括前に対策 (A) として `.claude/agents/research-subagent.md` を強化（コミット f828879）:
- 「最重要ルール」セクション新設、良い例 / 悪い例 3 パターンを対比形式で明示
- アンチパターン欄に 3 件追加（前置き混入 / code fence / Sources 二重記述）
- 「最終出力チェックリスト」5 項目を新設

### γ E2E (2026-05-03) での実測結果

(A) 強化後でも、**6 subagent すべてが混入を再発**:

| 混入パターン | 件数 |
|---|---|
| 前置き説明文 | 6/6 |
| マークダウン code fence | 6/6 |
| JSON 後の `Sources:` トレーリング | 4/6 |
| 二重 JSON 出力（最深刻） | 1/6 (data_06) |

結果、context 削減効果は理論値 -50% に対し実測 **-20.4%**（measurements.md 参照）。残り 30% が wrapper bloat に食われている。

### 原因仮説
- subagent モデル `haiku` の指示遵守率が弱い
- 「最終出力チェックリスト」の self-check 5 項目が「json.loads シミュレーション」の手順だけを促しても、実出力を変える行動につながらない
- LLM の「親切心」(出典明示しないと親が困るという推測) が制約より優先される

### 改善策候補

| 改善策 | 期待効果 | 工数 | 優先度 |
|---|---|---|---|
| **(B) 親側 JSON 抽出 helper**(regex で `{...}` ブロック抽出 + `data.*` フィールド取り出し) | 削減効果 +10-15%（理論値 -40% まで改善見込み） | 1-2h | **高** |
| (A2) research-subagent.md のさらなる強化（末尾に絶叫トーン制約 / 出典は data.sources[] のみ再強調） | +5% 程度 | 30 分 | 中 |
| subagent モデル変更（haiku → sonnet） | +10-20% | コスト 4-5x | 低 |
| 二重 JSON 防止の hard guard（最初の `}` で打ち切り） | edge case 防止 | 1h | 中 |

### 推奨: v0.4 で (B) の着手検討

(B) は `_common/` に汎用 helper を 1 本追加し、3 orchestrator の SKILL.md で「parsed = parse_subagent_return(result)」のように呼ぶ規約を導入する形が想定。

### 改善策 (B) の実装内容（2026-05-05）

format_add ブランチ内で先行実装。E2E 再計測は次セッション以降に持ち越す。

**実装ファイル**:
- `skills/_common/lib/parse_subagent_return.py`(新設) — 4 段抽出ロジック
  1. そのまま `json.loads`(規約遵守ケース、最速)
  2. マークダウン code fence (` ```json ... ``` ` / ` ``` ... ``` `) を除去して再試行
  3. 最初に出現する均衡の取れた `{...}` ブロックを抽出して再試行（二重 JSON 出力時は最初の 1 つを優先）
  4. 全失敗で `ValueError`(原文 head/tail 200 字付き)
- `skills/_common/lib/test_parse_subagent_return.py`(新設) — pytest 12 ケース全通過
  - clean / 前置きあり / fence (json 付/なし) / 末尾 Sources / 二重 JSON / 文字列リテラル内の `{}` / エスケープ済み `"` / 複合 / 不正値 / 非 str 型 / エラー msg 内の head/tail 検証

**規約反映**:
- `skills/_common/references/harness_levers.md` レバー②に「subagent return value のパース規約」セクション追加
- 3 orchestrator の Step 1 code 例を `parsed = json.loads(result)` → `parsed = parse_subagent_return(result)` に書き換え（market-overview-agent / company-deepdive-agent / business-deepdive-agent）

**E2E 再計測について**:
γ E2E 再実行による削減効果の実測（現状 -20.4% → 理論値 -40% への到達確認）はユーザー判断で省略。helper の単体動作は pytest 12 ケースで担保済。次回 γ E2E 実行時に実測値を `outputs/harness_check/measurements.md` へ追記すれば ISSUE-009 の効果検証は完結する（追加 ISSUE 起票は不要）。

### 参考ファイル
- `.claude/agents/research-subagent.md`(2026-05-03 強化済)
- `skills/_common/lib/parse_subagent_return.py`(2026-05-05 新設)
- `outputs/harness_check/measurements.md`(γ 実測値)
- `docs/harness_check/handoff.md` Section 5「2026-05-03 γ E2E 観察」
- `docs/harness_check/lever_mapping.md`「Phase B 完了時の **最終実 Status**」

---

## ISSUE-011: format_add ブランチ V2 前倒し完結（roleup 仕様正本化と Pilot 3 改修）

**Status**: クローズ / **Priority**: P1 / **Decided**: 2026-05-04 / **Closed**: 2026-05-05

### 背景
V1（commit `4d752b1` 〜 `0a811c8`）で Pilot 3 を brand-aware 化したが、ユーザーが `--brand roleup` 出力を確認し
「キーメッセージ・タイトル位置とフォント、コンテンツ配置、フォントサイズなど、ロールアップ仕様の根幹がほぼ反映できていない」と判定。
構造的未到達領域（curated roleup template 不在、theme.json 数値が stella 流用、行高/数値書式/条件付きフォントサイズ schema 未定義）が判明し、
ISSUE-010 で V2 に分離していたタスクを **format_add ブランチ内で前倒し完結** する方針で合意。

### 計画書
`/Users/nakamaru/.claude/plans/lovely-skipping-aho.md`(Option A、format_add 内完結)

### 進捗

| Phase | 内容 | 状態 |
|---|---|---|
| Phase 0 | 仕様正本化 + 公式テンプレ受領 | ✅ 完了（2026-05-04, a2c0d88） |
| Phase 1 | theme.json schema 2.0 拡張 + brand_resolver accessor 追加 | ✅ 完了（2026-05-04, 20f323e） |
| Phase 2 | curated roleup template 配置 + check_template_invariants.py | ✅ 完了（2026-05-04, e386e89） |
| Phase 3 | layout.json A4 横座標再計算 | ✅ 完了（2026-05-04, 0b11d12） |
| Phase 4a | format_helpers.py 新設 + customer-profile rollup 改修 | ✅ 完了（2026-05-04, a3cbf70） |
| Phase 4 fix-1 | rollup → roleup 全リネーム + shape mapping 公式整合 | ✅ 完了（2026-05-04, 2d3057c） |
| Phase 4 fix-2 | スライドタイトル/メインメッセージ/サブタイトル概念整理 | ✅ 完了（2026-05-04, 6b85a3c） |
| Phase 4 fix-3 | roleup 微修正 7 件（cp 1-1〜1-4 + me 2-1〜2-3） | ✅ 完了（2026-05-04, 25071be） |
| Phase 4 fix-4 | ユーザー視覚調整値に正確に追従（cp/me 微差解消） | ✅ 完了（2026-05-04, fbd02d7） |
| Phase 4 fix-5 | roleup 視覚レビュー指摘 3 件対応 (茶色ガイド除去 / 10pt 統一 / 凡例被り) | ✅ 完了（2026-05-05, f82521c） |
| Phase 5 | check_brand_compliance.py + ch ガイド除去 + ch 出典 | ✅ 完了（2026-05-05, 577739d） |
| 最終ユーザー視覚レビュー | pilot3 v5 出力 (cp/me/ch × roleup/stella) | ✅ 完了（2026-05-05、ユーザー OK） |

### クローズ理由

Pilot 3 (customer-profile / company-history / market-environment) の roleup brand 化を、
公式 vF 20250928 仕様に整合させる目標を達成。最終ユーザー視覚レビューで合格。
- Pilot 3 × roleup/stella の 6 ファイル全てが期待どおりの出力
- 全 stella regression-zero (shape 構成・座標が旧版と完全一致)
- Phase 5 静的検査で全 28 checks PASS (`tools/check_brand_compliance.py`)

残り 25 fill scripts への brand 展開は ISSUE-010 (保留中) で別途対応する。

### 次セッション着手点（2026-05-05 引き継ぎ）

**最新 commit**: `577739d` (format_add ブランチ)
**最新出力**: `outputs/v2_phase4_align_v5/{cp,me,ch}_{roleup,stella}.pptx` (全 6 ファイル)
**Phase 5 静的検査**: 全 28 checks PASS (`tools/check_brand_compliance.py`)

#### Phase 4 fix-5 の解決事項（2026-05-05）
1. **茶色ガイド矩形除去**: cp/me の `正方形/長方形 1` `正方形/長方形 8` (accent2 茶色) を出力から silently 除去。fill 内に `_silent_remove_shape` helper 新設(stella では no-op で regression-zero)
2. **本文・表 10pt 統一**: cp/me に `_body_pt(stella_fallback)` / `_body_sz(stella_fallback_sz_str)` helper 新設。roleup は `theme.font_size_body_pt`(=10pt) を返し、stella は旧 hardcode 値を維持。chart 軸 (sz='1100'→'1000')・data label・unit_label・凡例・CAGR・KPI 注釈の Pt(11/12/14/15/16) hardcode を全置換
3. **cp 凡例被り解消**: `add_custom_legend` に `max_right_emu` 引数を追加し、unit_label 左端 - 0.10 in を上限に凡例右端を制限。stella は None で従来挙動維持
4. **Phase 4c 同梱**: me Source placeholder のフォールバック検索 (Source/Source 3) を追加し warning 解消

#### Phase 5 の解決事項（2026-05-05, 577739d）
1. **brand compliance checker 新設**:
   - `skills/_common/lib/brand_compliance_rules.py` (ルール群と profile mapping)
   - `tools/check_brand_compliance.py` (CLI、--pptx/--skill ペア複数対応、text/json 出力)
   - pilot3 × roleup プロファイル全 10 ルール: C1/C2/C4/C5/C6/C7/C8/C10/C11/C12
     (cp/me 各 10、ch はチャート無いため C10/C12 除外で 8)
   - stella 版 profile は ISSUE-010 で stella 仕様確定後に追加 (現状 skeleton)
2. **ch fill 改修**:
   - `_silent_remove_shape` ヘルパー新設 + 茶色ガイド矩形除去
   - `require_source` 呼び出し追加 (roleup 出所必須)
   - `SHAPE_SOURCE = "Source 3"` 定数 + main 内 source 書き込みロジック
3. **sample_data.json 更新**: ch に `"source"` フィールド追加
4. **検証結果**: pilot3 全 28 checks PASS、ch_stella 含む全 stella regression-zero

#### 残タスク
- **最終ユーザー視覚レビュー** (cp/me/ch × roleup/stella の 6 ファイル): 問題なければ ISSUE-011 クローズ
- (任意) ISSUE-010 着手時に stella 用 compliance profile を populate

#### 設計済みの仕組み（次セッションでそのまま使える）
- **brand 別の shape 用途切替**: theme.json の `placeholder_role_mapping` (top/subtitle field)
- **stella regression-zero 担保パターン**: fill 内 `if theme.id == "stellar_aiz":` で旧コード分岐、roleup 経路に新ロジック
- **shape 生成順序**: stella では旧順序維持(cNvPr id 連番が旧と同一になり diff 0 を保証)
- **format_helpers**: `format_cell_value` / `format_fiscal_period` / `apply_line_spacing` / `require_source` / `resolve_top_text` / `resolve_subtitle_text` 完成済
- **brand_resolver accessor**: `top_placeholder_field()` / `subtitle_placeholder_field()` / `font_size_body_pt(skill_id)` / `line_height_pt()` / `number_format_excel()` / `zero_text()` / `negative_format()` / `fiscal_period_format()` / `layout_rule(key, default)` / `is_source_required()` / `is_executive_summary_skill(skill_id)` 完成済
- **検証ツール**: `tools/check_template_invariants.py`(全 6 テンプレ PASS), `tools/extract_roleup_template.py`(再抽出可能)

#### 現在の skills_factory 共通理解（2026-05-04 ユーザー定義）

| 用語 | 定義 | 例 | sample_data フィールド | フォントサイズ(roleup) |
|---|---|---|---|---|
| スライドタイトル | 短い見出し | 「○○市場の動向」 | `chart_title` | 22pt |
| メインメッセージ | 結論文 | 「○○は CAGR x.x% で成長」 | `main_message` | 14pt |
| サブタイトル | 各セクション小見出し | 「企業の概要」「業績」 | `section_title` | 12pt 左寄せ #897141 |

**brand 別の最上部 placeholder 意味**:
- stella: 最上部の最大フォント = メインメッセージ (`main_message`) ← 既存運用
- roleup: 最上部の最大フォント = スライドタイトル (`chart_title`) ← 公式 vF 整合

### Phase 0 確定値（2026-05-04 ユーザー確認済）

| Phase 0 項目 | 確定値 |
|---|---|
| F-1 公式テンプレ本数 | Phase 2 で決定（保留） |
| F-2 本文 10pt 適用範囲 | 本文・表すべて 10pt（タイトル 22pt / サブタイトル 12pt / キーメッセージ 14pt / 出所 6pt は別格） |
| F-3 executive_summary 12pt 判定 | (b) skill 名で自動判定（`executive-summary-pptx` のみ 12pt） |
| F-4 行高 12pt OOXML 表現 | (b) 段落の `<a:lnSpc><a:spcPts val="1200"/>` 指定 |
| F-5 マイナス表記デフォルト | `(XXXX)` 括弧表記（`△XXXX` も許容） |
| F-6 会計期間表記 | `YY/MM期`（`19/10期` 等） |
| F-7 左端揃えガイド | x = 0.41 inch（全 layout 共通） |
| F-8 出所必須化レベル | (b) 全 brand で hard-fail |

### 成果物
- ✅ `skills/_common/brands/roleup/format_spec.md` 新規作成（vF 20250928 正本）
- ✅ 公式テンプレ `work/roleup_official_templates/standard_format_vF_20250928.pptx` 受領

### 関連ファイル
- 計画書: `/Users/nakamaru/.claude/plans/lovely-skipping-aho.md`
- 仕様正本: `skills/_common/brands/roleup/format_spec.md`
- 公式テンプレ: `work/roleup_official_templates/standard_format_vF_20250928.pptx`

---

## ISSUE-010: 残り 25 fill scripts への brand 展開（V2）

**Status**: 進行中（Phase 0 完了、Phase 1 着手待ち） / **Priority**: P2 / **Decided**: 2026-05-04 / **Updated**: 2026-05-05

### 背景
V1（format_add ブランチ、commit `4d752b1` 〜 `128fa15`）で Pilot 3 スキル
（customer-profile-pptx / company-history-pptx / market-environment-pptx）を brand-aware 化し、
`stellar_aiz` / `roleup` の出力切替機構を確立した。残り 25 fill scripts は brand-aware 化未対応。

未対応スキル群（V2 対象）:
- BDD 系: company-overview-pptx-v2, revenue-analysis, sga-breakdown, financial-benchmark, shareholder-structure, customer-sales-detail, sales-by-customer, workforce-composition, cost-breakdown, scenario-forecast, current-period-forecast, valuation-summary, executive-summary, data-availability
- 市場分析系: market-share, positioning-map, market-kbf, pest-analysis, competitor-summary
- 戦略フレームワーク系: swot, five-forces, value-chain, value-chain-matrix, business-portfolio, business-overview, business-model, growth-driver, gate-process, comparison, conceptual, process-arrow, process-flow, kpi-dashboard, table-chart, table-of-contents, section-divider, gantt-chart, project-team-structure, issue-risk-list, logic-tree, pyramid-structure, smallcap-* 系

### V2 トリガー条件（どちらかを満たしたら着手）
- Roleup でのクライアント納品需要が顕在化（pilot 3 以外のスキルで Roleup 出力が必要になる）
- 追加クライアント（C 社・D 社…）の brand 追加要望が発生

### 改修パターン（pilot 3 で確立済）
詳細は `skills/_common/references/brand_migration_guide.md` 参照。
- Pattern A（hardcode 駆動）: 90 分 〜 3 時間 / スキル
- Pattern B（テンプレ rPr/tcPr 駆動）: 30 分 / スキル
- Pattern C（HTML→Playwright 駆動）: 60 〜 90 分 / スキル

### V2 で curated roleup template の導入も併せて検討
V1 では Roleup 用テンプレ pptx は配置せず、`brand_resolver.template_path()` のフォールバックで stella テンプレを流用。
V2 で各スキルに A4 横（11.69×8.27）+ Yu Gothic UI + 褐色アクセントの curated roleup テンプレを順次導入。

### 参考ファイル
- `skills/_common/lib/brand_resolver.py`
- `skills/_common/brands/{stellar_aiz,roleup}/theme.json`
- `skills/_common/references/brand_migration_guide.md`
- pilot 3 commits: `b767ee3`, `c199f03`, `128fa15`
- 計画書: `~/.claude/plans/1-pc-harness-check-2-20250928-vf-pptx-agile-kite.md`(V1 期、参考)
- v0.4 計画書: `~/.claude/plans/dreamy-greeting-pizza.md`(Phase 0、2026-05-05)

### Phase 0 完了内容（2026-05-05）

ユーザー要件「ブランド対応を全スキルに拡充 + agent 系の brand 確定 UX 作り込み」を受けて、N 社 agnostic 設計の規約と土台を確定。fill 改修は無し（Phase 2）。

**ユーザー確定事項**:
- Q1: agent Step 0 で `AskUserQuestion` 都度確定（env / config 固定はしない）
- Q2: 将来 N 社追加前提（agnostic 設計）
- Q3: 未対応 fill は warning + stella fallback（hard-fail せず）
- Q4: fill 展開優先順は BDD 系から（Phase 2）

**実装したもの**:
- `skills/_common/lib/brand_resolver.py` agnostic 化:
  - `_discover_brands(brands_dir=None)`(`os.listdir` で `theme.json` 存在 dir を動的検出、D1 命名規則で filter)
  - `_validate_brand_id(brand)`(正規表現 `^[a-z][a-z0-9_]{1,23}$`)
  - `is_brand_supported_by_skill(skill_dir, brand)`(SKILL.md frontmatter `supported_brands` を簡易 regex で読む、未指定は `[stellar_aiz]`)
  - `resolve_brand_with_fallback(brand, skill_dir)`(unsupported なら warnings.warn + stella を返す one-call helper)
  - `VALID_BRANDS` ハードコード廃止（外部参照ゼロ確認済）
- `skills/_common/lib/test_brand_resolver.py`(新設): pytest 27 ケース全 pass
- `skills/_common/prompts/step0_brand_clarification.md`(新設): agent 共通プロンプト、AskUserQuestion テンプレ + scope.json 保存例 + warning fallback フロー
- `skills/_common/references/orchestrator_contract.md` §4: agnostic 化対応で §4.1〜§4.6 に再構成、`brand_fallback` warning スキーマを §2 に追記
- `skills/_common/references/brand_migration_guide.md`: §7 チェックリストに `supported_brands` frontmatter 追記項目を追加、新章 §9「Phase 0 で確定した agnostic 規約サマリ」

**検証結果**:
- pytest: 27（brand_resolver） + 12（parse_subagent_return）= 39 件全 pass
- Pilot 3 regression-zero: cp / me を `--brand stellar_aiz` / `--brand roleup` で起動、両方完走
- check_brand_compliance: cp + me roleup で 20/20 PASS
- VALID_BRANDS 外部参照: brand_resolver.py 以外でゼロ確認

### Phase 1 進捗

- ✅ **(i) frontmatter 一括追加** — 全 fill SKILL.md に supported_brands を機械的追加（commit `5837223`、48 件 / pilot 3 = `[stellar_aiz, roleup]` / その他 = `[stellar_aiz]` / nttdata-pptx は除外で default `(stellar_aiz,)` 扱い、merge-pptxv2 は glob 不一致で自然除外）
- ✅ **(ii) 共通プロンプトのコピペ展開** — agent 7 件（market-overview / strategy-report / company-deepdive / business-deepdive / smallcap-strategy-research / bdd-init / comparison-synthesis-agent）の Step 0 に `step0_brand_clarification.md` の sync コメント + 短縮抜粋（AskUserQuestion 擬似コード + scope 保存仕様）を埋め込み（commit `42e5377`、grep で 7 件全 hit）
- ✅ **(iii) orchestrator に warning fallback 実装** — `_common/lib/orchestrator_helpers.py` に `resolve_fill_brand_with_warning()` / `append_brand_warnings_to_merge_file()` を新設し pytest 10 ケース全 pass（commit `c07ac24`）。bdd-init を除く 6 agent の fill 起動ループに warning fallback 擬似コード差し込み（commit `a70304a`）。warning 蓄積先は `merge_warnings.json` 流用（orchestrator が merge 後に append、merge-pptxv2 の `"w"` 上書きとの干渉を回避）
- ✅ **(iv) E2E 1 本** — market-overview-agent × `--brand roleup` 疎通確認完了（2026-05-05）
- ✅ **(v) Phase 1.5 — 42 非 pilot fill に passive `--brand` 導入** — Phase 1 (iv) で炙り出した「非 pilot fill が `--brand` を argparse 拒否」問題を mechanical に解消（2026-05-05、commit `28ed496`、smallcap-* 5 件除外）

### Phase 1 (ii)+(iii) 完了内容（2026-05-05）

**ユーザー確定事項**（本セッション）:
- Q-A 共通プロンプト埋め込み方式: **短縮抜粋 + 参照リンク**（既存 `step0_scope_clarification.md` の sync 慣習に準拠）
- Q-B warning fallback 蓄積先: **`merge_warnings.json` 流用**（§4.4 確定済スキーマと整合、別ファイル分割回避）
- Q-C warning fallback ロジック集約: **`_common/lib/orchestrator_helpers.py` 新設**（pytest 担保 + Phase 2 以降の保守性）

**重要発見**: `merge-pptxv2/scripts/merge_pptx_v2.py:321-328` の `write_merge_warnings()` は `"w"` モードで `merge_warnings.json` を**新規上書き**する。orchestrator が fill 起動前に追記すると merge 時に消えるため、**「fill ループ中はメモリバッファ → merge 完了後に read+append+write」フロー**で干渉を回避する設計を採用。

**実装ファイル**:
- 新規 `skills/_common/lib/orchestrator_helpers.py`（60 行、2 関数）
- 新規 `skills/_common/lib/test_orchestrator_helpers.py`（pytest 10 ケース、全 pass）
- 改修 7 agent SKILL.md（market-overview / strategy-report / company-deepdive / business-deepdive / smallcap-strategy-research / bdd-init / comparison-synthesis-agent、+450 行 / -31 行）

**business-deepdive-agent の特殊対応**: 内部呼び出し JSON schema に `"brand"` フィールドを追加（親 `company-deepdive-agent` の `scope.json.brand` を子に転写、子側で AskUserQuestion 再質問しない）。`segment_summary.json` schema に `brand` / `brand_warnings` を追加し、子は `merge_warnings.json` に直接書かず親に返却する責務分離。

**bdd-init の特殊対応**: 「実行手順」の冒頭に新 Step 0「ブランド確認」を新設（既存は Step 1 mkdir から開始）。Step 2 で `meta.json.brand` / `meta.json.brand_label` を保存し、後続の `bdd-report` 等が読む設計（fill を直接起動しないため (iii) は対象外）。

**検証結果**: pytest 49 件全 pass（27 brand_resolver + 12 parse_subagent_return + 10 orchestrator_helpers）、sync コメント 7 件 grep ヒット、helper 単体動作確認 OK（pilot3 + roleup → roleup / swot + roleup → stellar_aiz fallback + buffer 1 件）。

### Phase 1 (iv) E2E 完了内容（2026-05-05）

**E2E スコープ**: `market-overview-agent` × `scope.json.brand="roleup"` の Step 5（fill 起動ループ + brand fallback）+ Step 7（merge-pptxv2 + `append_brand_warnings_to_merge_file`）の疎通確認に集中。Web 検索（Step 1）/ fact-check（Step 2.5）/ visual review（Step 8）はスコープ外（brand wiring と直交）。

**実行手順**: 既存 γ E2E（2026-04-28 国内タクシー市場、事業者のみ）の data 12 件を `work/market-overview-agent/2026-05-05_taxi_e2e_roleup/` にコピー → `scope.json.brand="roleup"` 設定 → Python ドライバー `run_e2e.py` で Step 5-7 を実行。

**検証結果**:

| 検証項目 | 期待値 | 実測値 | 結果 |
|---|---|---|---|
| 12 fill 完走 | 12/12 success | 12/12 | ✅ |
| pilot 3 (market-environment) が `--brand roleup` で起動 | True | True | ✅ |
| 非 pilot 9 skill が `--brand stellar_aiz` フォールバック | True (11 invocations: section-divider × 3 + 8 unique skills) | True | ✅ |
| brand_fallback warning 件数 | 11 (section-divider 3 件 + 他 8 件) | 11 | ✅ |
| `merge_warnings.json` のスキーマ整合 | §4.4 (`slide_index=-1`, `type="brand_fallback"`, `message`) | 完全一致 | ✅ |
| merged PPTX 完走 | 12 slide | 12 slide / 842KB | ✅ |
| `RuntimeWarning` 発火 | 11 件（fallback ごとに 1 回） | 11 件 | ✅ |

**成果物**:
- `outputs/E2E_MarketOverview_taxi_roleup_2026-05-05.pptx`(12 スライド、842KB)
- `outputs/merge_warnings.json`(brand_fallback × 11)
- `work/market-overview-agent/2026-05-05_taxi_e2e_roleup/{run_e2e.py, fill_results.json, merge_order.json, scope.json, data_*.json, slide_*.pptx}`

**Phase 1 (iv) で炙り出した SKILL.md 整備課題（Phase 2 で対応）**:

1. **`--brand` 引数の条件付き付与**: SKILL.md の擬似コード（market-overview-agent L567-L577 / strategy-report-agent / 他 5 agent）は `subprocess.run(["python", fill_script, "--brand", fill_brand, ...])` と無条件に `--brand` を渡すが、現状 `--brand` を受け付けるのは pilot 3（customer-profile / company-history / market-environment）のみ。**非 pilot 9 fill は plain `argparse.ArgumentParser` で `--brand` を未定義引数として SystemExit する**。E2E ドライバーでは `add_brand_arg` import を grep する `fill_supports_brand_flag()` ヘルパーで条件分岐を実装したが、orchestrator 側にも同等の判定が必要。Phase 2 で全 fill が `--brand` を passively 受けるように `add_brand_arg` を導入するのが本筋（その間は擬似コード側に注記）。
2. **テンプレートパス命名の不統一**: `table-of-contents-pptx-template.pptx` / `section-divider-pptx-template.pptx` は `<skill_name>-template.pptx` 慣習から外れる（"-pptx-template" になっている）。`market-kbf-pptx` の fill script 名も `fill_kbf.py`(他は `fill_<skill_id>.py`)。orchestrator 側に explicit override が必要。Phase 2 で命名を揃えるか、orchestrator helper にマッピング dict を追加するかの判断を要する。
3. **pilot 3 の roleup データバリデーション**: `market-environment` は roleup の `fiscal_period_format()` が ON のとき `int(d["year"])` を要求するため、`"2025E"`/`"2026E"` 等の estimated suffix 入りデータは ValueError で落ちる。Phase 2 では fill 内で suffix を strip するか、データ生成側（subagent prompt）に整数化を強制するかの判断を要する。今 E2E では手動で suffix を除去して回避。

### Phase 1 (v) — Phase 1.5 完了内容（2026-05-05、commit `28ed496`）

**実装内容**: AST migration helper（`tools/add_passive_brand_arg.py`、commit には含めず使い捨て）で 42 非 pilot fill に以下 6 行を mechanical に挿入:

```python
# brand_resolver bootstrap (passive --brand acceptance until brand-aware migration)
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(SKILL_DIR, "..", "_common", "lib"))
from brand_resolver import add_brand_arg  # noqa: E402

# in main() before parser.parse_args():
add_brand_arg(<parser_var>)  # passive: accepted but ignored until brand migration
```

**スコープ**: 42 fill（smallcap-* 5 件は ISSUE-011 retrospective に従い除外）

**検証結果**:
- pytest 49 件全 pass（regression-zero）
- pilot 3 × roleup で `check_brand_compliance.py` 28/28 PASS（Pilot 3 出力に変化なし）
- Phase 1 (iv) E2E driver から `fill_supports_brand_flag` 条件分岐削除 + 全 fill に `--brand` 無条件付与 → 12/12 完走、warning 11 件
- 5 重点 fill（exec-summary / market-share / positioning-map / competitor-summary / data-availability）の `--brand roleup` smoke test 全完走

**意義**: Phase 2 以降、orchestrator は SKILL.md 擬似コード通りに `--brand` を fill に無条件付与すれば動作する。Phase 2 で個別に brand-aware 化する際は、`add_brand_arg` の呼び出しは既に置いてあるので、`args.brand` を `resolve_brand` に渡すだけで完成。

### Phase 2 進捗

| 着手順 | スキル | 状態 | Pattern | コミット |
|---|---|---|---|---|
| 1 | `executive-summary-pptx` | ✅ 完了 | A (hardcode 駆動) | `a96f53f`(2026-05-05) |
| C-1 | `market-share-pptx` | ✅ 完了 | A | `0f6dffd`(2026-05-05) |
| C-2 | `positioning-map-pptx` | ✅ 完了 | A | `64599c8`(2026-05-05) |
| C-3 | `competitor-summary-pptx` | ✅ 完了 | A | `fc6e5e7`(2026-05-05) |
| C-4 | `market-kbf-pptx` | ✅ 完了 | A | `9c02a67`(2026-05-05) |
| C-5 | `pest-analysis-pptx` | ✅ 完了 | A | `f12f738`(2026-05-05) |
| C-6 | `section-divider-pptx` | ✅ 完了 | A (装飾系、C4 除外 profile) | `63e2773`(2026-05-05) |
| C-7 | `table-of-contents-pptx` | ✅ 完了 | A | `23abee4`(2026-05-05) |
| C-8 | `data-availability-pptx` | ✅ 完了 | A | `6f3cc31`(2026-05-05) |
| 2 | `revenue-analysis-pptx` | ✅ 完了 | A (hardcode 駆動) | (next commit)(2026-05-06) |
| 3 | `financial-benchmark-pptx` | ✅ 完了 | A (hardcode 駆動) | (next commit)(2026-05-06) |
| 4 | `company-overview-pptx-v2` | ✅ 完了 | A (hardcode 駆動 + 専用テンプレ生成スクリプト) | `693837b`(2026-05-06) |
| 5 | `shareholder-structure-pptx` | ✅ 完了 | A | `8828c1d`(2026-05-06) |
| 6 | `business-portfolio-pptx` | ✅ 完了 | A | `9e1c320`(2026-05-06) |
| 7 | `sga-breakdown-pptx` | ✅ 完了 | A (複合チャート + secondary axis line + trend arrow) | `b3ac7e8`(2026-05-06) |
| 8 | `cost-breakdown-pptx` | ✅ 完了 | A (1〜2チャートモード) | `acb4dfe`(2026-05-06) |
| 9 | `workforce-composition-pptx` | ✅ 完了 | A (棒チャート + テーブル) | `dff331f`(2026-05-06) |
| 10 | `business-overview-pptx` | ✅ 完了 | A (revenue_chart + kpi_cards 両モード) | `928a6f3`(2026-05-06) |
| 11 | `sales-by-customer-pptx` | ✅ 完了 | A (テーブル N 個横並び、期数別動的フォント vs roleup 固定 10pt) | (next commit)(2026-05-06) |
| 12 | `valuation-summary-pptx` | ✅ 完了 | A (3 chart_type: football_field / equity_bridge / financial_summary) | (next commit)(2026-05-06) |
| 13 | `scenario-forecast-pptx` | ✅ 完了 | A (Base/Up/Down 折れ線 ×2 並列、シリーズ + 期間種別凡例) | (next commit)(2026-05-06) |
| 14 | `business-model-pptx` | ✅ 完了 | C (HTML→Playwright、CSS 変数注入 + curated roleup テンプレ生成) | (next commit)(2026-05-06) |
| 15 | `customer-sales-detail-pptx` | ✅ 完了 | C (HTML→Playwright、CSS 変数注入 + curated roleup テンプレ生成) | (next commit)(2026-05-06) |
| 16 | `current-period-forecast-pptx` | ✅ 完了 | C (HTML→Playwright、CSS 変数注入 + curated roleup テンプレ生成、テンプレ命名統一) | (next commit)(2026-05-06) |

### market-overview-agent × roleup フルネイティブ達成（2026-05-05）

C-6〜C-8 (section-divider / table-of-contents / data-availability) 完了で、
market-overview-agent デッキ 12 スライド全てが roleup native 化。

**E2E 再実行結果**:
- 既存 work dir `work/market-overview-agent/2026-05-05_taxi_e2e_roleup/` を流用、
  `run_e2e.py` 内のテンプレ命名 (`<skill>-pptx-template.pptx` → `<skill>-template.pptx`)
  と fill script 命名 override (`fill_kbf.py` → `fill_market_kbf.py`) を統一規則に修正
- 12/12 fill 完走 (success rate 100%)
- `brand_fallback` warning 件数 = **0** (前回 11 件 → 0 件に減少、目標達成)
- `merge_warnings.json` = `[]`(warnings 完全に 0)
- 出力: `outputs/E2E_MarketOverview_taxi_roleup_2026-05-05.pptx` (12 slide / 1.43MB / A4 横)
- 全 RuntimeWarning 発火 0 件

**累積 compliance check**: pilot 3 + market 系 5 + sd/toc/da の 10 PPTX × 80 checks 全 PASS
(executive-summary と pest-analysis は sample_data 文字数違反でスキップ。後述)

### 既知の sample_data 不備（Phase 2 残課題、優先度低）

~~セッション中に累積 compliance check 走行中、以下 2 件の sample_data が hard-fail:~~
~~- `executive-summary-pptx/references/sample_data.json`: findings[1].detail = 103 chars (上限 100)~~
~~- `pest-analysis-pptx/references/sample_data.json`: main_message = 66 chars (上限 65)~~

→ **完了 (2026-05-06、commit `50f38d2`)**:
- executive-summary findings[1].detail 103 → 91 chars
- pest-analysis main_message 66 → 63 chars
両 fill は累積 compliance check 対象に復帰した。

### BDD トリオ完結 (2026-05-06)

revenue-analysis / financial-benchmark / company-overview-pptx-v2 の 3 fill 完了で、
`company-deepdive-agent` デッキ (会社レベル + 事業セグメント) の core 構成が
roleup ネイティブで生成可能に。

**累積 compliance check (2026-05-06)**: 15 PPTX × 122 checks 全 PASS
- pilot 3 (cp/me/ch) + market 系 5 (market-share/positioning-map/competitor-summary/market-kbf/pest-analysis)
- 装飾系 3 (section-divider/TOC/data-availability) + executive-summary
- BDD 3 (revenue-analysis/financial-benchmark/company-overview-pptx-v2)

### 設計判断ログ (2026-05-06、ユーザー判断)

| 件 | 選択肢 | 採用 | 理由 |
|---|---|---|---|
| revenue-analysis EBITDA 棒色 | chart_palette[3] / chart_palette[1] / accent_ebitda_bar 新設 | **chart_palette[1] = #897141 (ベージュ)** | 茶系隣接トーン、Revenue (#7C4C2C) + Margin Line (#604C3F) と 3 色で brand 一貫性高 |
| financial-benchmark マイナス値バー | 純赤維持 / accent_op_margin_line / negative_bar 新設 | **accent_op_margin_line (#604C3F)** | 財務文化の純赤を捨てて brand 一貫性優先。対象会社 negative も同色 + bold で識別 |

### Phase 2 着手点（次セッション以降）

1. ~~**BDD 系 fill 3 件 brand-aware 化**~~ → **完了 (2026-05-06)**: revenue-analysis / financial-benchmark / company-overview-pptx-v2
2. ~~**市場系 fill 5 件**~~ → **C-1〜C-5 完了 (2026-05-05)**
3. ~~**Phase 1 (iv) 残り課題**~~ → **完了 (2026-05-05、commit 7dec7b6)**

**残 BDD 系 (15 件)**: ~~shareholder-structure / business-portfolio / sga-breakdown / cost-breakdown / workforce-composition~~ (5 件完了 2026-05-06) / ~~business-model~~ / ~~sales-by-customer~~ (完了 2026-05-06) / ~~customer-sales-detail~~ / ~~scenario-forecast~~ (完了 2026-05-06) / ~~current-period-forecast~~ / ~~valuation-summary~~ (完了 2026-05-06) / ~~business-overview~~ (完了 2026-05-06) / その他戦略フレームワーク系。

**現セッション完了分 (2026-05-06)**:
- 累積 24 PPTX × 206 checks 全 PASS (前回 20 PPTX × 162 → +4 PPTX × 44 checks)
- Pattern A BDD 系 4 件追加完了 (business-overview / sales-by-customer / valuation-summary / scenario-forecast)
- valuation-summary は副次的に Phase 1.5 AST migration で混入していた `import os` 抜けバグも修正
- 残 Pattern C 系 3 件 (business-model / customer-sales-detail / current-period-forecast)

**Pattern C 3 件完了 (2026-05-06、後続セッション)**:
- 累積 27 PPTX × 230 checks 全 PASS (24 → +3 PPTX × +24 checks)
- 各スキルに `scripts/build_roleup_template.py` (one-shot generator) を新設し、
  stella テンプレ → A4 横 + Yu Gothic UI + 茶系 + Source 3 placeholder の curated roleup テンプレを生成
- HTML CSS 変数を theme.json から注入する Pattern C 規約を確立（`_apply_theme(theme)` で
  module-level *_HEX 文字列を上書き → CSS f-string で参照）
- Title/Subtitle/Source の placeholder text には `_make_brand_run` でフォントサイズと色を明示注入
  (LibreOffice roundtrip 後の placeholder rename "Title 1"→"PlaceHolder 1" にも fallback で対応)
- Content Area 装飾矩形は picture 挿入後 `_silent_remove_shape` で削除し C1 PASS を担保
- current-period-forecast はテンプレ命名 `forecast-template.pptx` →
  `current-period-forecast-template.pptx` へ統一 (`brand_resolver.template_path()` 規約に揃える)
- 全 fill が `--brand stellar_aiz / roleup` で完走、SKILL.md `supported_brands: [stellar_aiz, roleup]`

**company-deepdive-agent × roleup E2E**: 3 fill 完了したので次々セッションで実施可能。
N=1 (二幸産業) work dir があれば流用、なければ簡易 E2E driver を整備して回す。

### executive-summary-pptx 完了内容（2026-05-05、commit `a96f53f`）

**手順**:
- A1 調査: Pattern A 判定（hardcode 駆動、5 findings 縦積み layout、SHAPE_MAIN_MESSAGE/CHART_TITLE/SOURCE）
- A2 template: cp roleup template から派生（object 8 装飾削除、Title 1/Text Placeholder 2/Source 3 + 茶色ガイド矩形を保持）。check_template_invariants 両 brand PASS
- A3 fill: resolve_brand + _apply_theme(theme) + resolve_top_text/subtitle_text + require_source + silent_remove ガイド矩形 + Source 3 placeholder 利用
- A4 検証: check_brand_compliance 8/8 PASS / pilot 3 regression-zero (36/36 PASS) / stella byte-stable
- A5 SKILL.md: supported_brands [stellar_aiz, roleup]、`--brand` ドキュメント更新

**重要技術発見**:
- `add_text_box` の `font_name=FONT_NAME_JP` デフォルト引数は **module load 時に評価される**（Python late-binding）。`_apply_theme(theme)` が module global を更新しても、デフォルト引数は古い値を保持。次の Pattern A 移行（revenue-analysis 等）でも同じ落とし穴を踏む可能性が高い。**修正パターン**: `font_name=None` + 関数内 `if font_name is None: font_name = FONT_NAME_JP` で late binding を強制する。
- exec-summary は font_size_executive_summary_body_pt（roleup 12pt）を使う特殊スキルで、`theme.font_size_body_pt(skill_id="executive-summary-pptx")` 経由でアクセス。`executive_summary_skill_ids` に skill_id が登録されているため自動切替される。

**Phase 2（5-10 セッション）の前提**: Phase 1 (iv) E2E で wiring 自体は健全と確認済 + Phase 1.5 で fallback flow が SKILL.md 擬似コードのまま本番動作可能になった。fallback で運用可能なので、急いで brand-aware 化しなくても roleup ユーザーには「pilot 3 + exec-summary が roleup native、他は stella で fallback」が透明性高く伝わる状態。

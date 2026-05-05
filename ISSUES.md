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
- ⏳ **(ii) 共通プロンプトのコピペ展開** — agent 7 件（market-overview / strategy-report / company-deepdive / business-deepdive / smallcap-strategy-research / bdd-init / comparison-synthesis-agent）の Step 0 に `step0_brand_clarification.md` を埋め込み
- ⏳ **(iii) orchestrator に warning fallback 実装** — fill 起動前の `is_brand_supported_by_skill` 呼び出しと `merge_warnings.json`(または別ファイル) への brand_fallback 追記
- ⏳ **(iv) E2E 1 本** — market-overview-agent × roleup で疎通確認

**次セッションは (ii)+(iii) をまとめて実施**(密接に連動)。実装ヒント:
- agent SKILL.md の Step 0 冒頭に `<!-- source: skills/_common/prompts/step0_brand_clarification.md (manual sync until D2) -->` コメント付きでコピペ
- orchestrator は scope.json の brand 確定後、各 fill 起動 loop で `is_brand_supported_by_skill(skill_dir, scope_brand)` を呼ぶ
- warning 蓄積先（既存 `merge_warnings.json` を流用するか、別ファイル `brand_warnings.json` を切るか）は (iii) 着手時に判断

### Phase 2（5-10 セッション）

BDD 系 fill 14 件を Pattern A/B/C で順次 brand-aware 化。優先順: executive-summary → revenue-analysis → data-availability → financial-benchmark → company-overview-pptx-v2。

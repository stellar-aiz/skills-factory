# Step 進捗トラッキング（共通パターン）

> **このファイルは `skills/_common/prompts/step_state_tracking.md` です。**
> オーケストレータースキル（market-overview-agent / company-deepdive-agent / business-deepdive-agent 等）の SKILL.md の各 Step 冒頭から、`<!-- source: skills/_common/prompts/step_state_tracking.md (manual sync until D2) -->` コメント付きで**手動コピペ**してください。
> このファイルを変更したら `grep -rn "source: skills/_common/prompts/step_state_tracking.md" skills/*/SKILL.md` で被参照スキルを全て検出し、コピペし直すこと（ISSUE-001 D2 で自動化検討中）。

オーケストレーターが各 Step を進める際の **TaskCreate / TaskUpdate 利用規約** と、**`task_state.json` ファイルベース状態共有** の標準パターンを定義する。

目的:
1. **#5 状態統合**: ログを掘らなくても TaskList で「いま何 Step まで終わってるか」が分かる
2. **#6 開始停止再開**: 中断後の resume 地点を機械的に判定できる
3. **#8 制御フロー**: hooks（`tools/hooks/check_task_progression.py`）が前 Step 完了を assert できる
4. **#12 ステートレス**: LLM の会話履歴に頼らず、外部ファイルが state machine の真実源

---

## 必須運用：3 つの操作セット

各 Step 開始時 / 終了時 / 失敗時に、以下 3 操作セットを **必ず** 実行する。

### Step 開始時

```
1. TaskCreate(
     subject="<orchestrator>: Step <N> - <topic>",
     description="<入力 / 出力 / 完了条件>",
     activeForm="<...ing>"
   )
   → task_id を取得

2. TaskUpdate(taskId=<id>, status="in_progress")

3. task_state.json に entry を追加（後述スキーマ）
```

### Step 終了時

```
1. TaskUpdate(taskId=<id>, status="completed")
2. task_state.json の該当 entry の status を "completed" に、completed_at を現在時刻に
```

### Step 失敗時（再試行する場合）

```
1. TaskUpdate は触らない（status="in_progress" のまま）
2. task_state.json の該当 entry に retry_count をインクリメント
3. 再試行が 2 回失敗したら、ユーザーに判断を仰ぐ
```

---

## TaskCreate の subject 命名規約

### フォーマット

```
<orchestrator-slug>: Step <N> - <topic>
```

| 部分 | 説明 | 例 |
|---|---|---|
| `<orchestrator-slug>` | オーケストレーター名（拡張子なし） | `market-overview` / `company-deepdive` / `business-deepdive` |
| `<N>` | Step 番号（数字または `0.5` 等） | `1` / `2.5` / `8-b` |
| `<topic>` | Step の主目的（5〜15 字程度） | `Web検索` / `スライド生成` / `Visual Review` |

### 例

| ✅ Good | ❌ Bad |
|---|---|
| `market-overview: Step 1 - Web検索` | `Step 1` （オーケストレーター不明） |
| `company-deepdive: Step 5 - 会社レベルPPTX生成` | `スライド作る` （Step 番号なし） |
| `business-deepdive: Step 4 - 5論点 PPTX 生成` | `business-deepdive: スライド` （Step 番号なし） |

理由: hooks が subject 文字列から `Step <N>` パターンを正規表現で抽出するため、フォーマット統一が必要。

---

## TaskCreate の description フォーマット

3 行の構造化テキスト:

```
入力: <この Step が消費する成果物 / ファイル>
出力: <この Step が生成する成果物 / ファイル>
完了条件: <"completed" にしてよい客観的判定基準>
```

### 例（market-overview-agent Step 5）

```
入力: scope.json, data_01..12_*.json (Step 1-4 の成果物)
出力: slide_01..12_*.pptx (12 ファイル)
完了条件: 全 12 PPTX が生成され、各 fill_*.py が exit 0 で終了している
```

理由: LLM が「Step 5 を completed にしてよいか」を客観判定するため。曖昧な完了条件は端折りの温床。

---

## `task_state.json` のスキーマ

### 配置

`{{WORK_DIR}}/<run_id>/task_state.json`(scope.json と同じディレクトリ)

### スキーマ

```json
{
  "run_id": "2026-04-30_taxi_industry",
  "orchestrator": "market-overview-agent",
  "started_at": "2026-04-30T10:00:00+09:00",
  "steps": [
    {
      "step_id": "step_1",
      "name": "Web検索による論点別情報収集",
      "status": "completed",
      "task_id": "12",
      "started_at": "2026-04-30T10:01:23+09:00",
      "completed_at": "2026-04-30T10:15:42+09:00",
      "retry_count": 0
    },
    {
      "step_id": "step_2",
      "name": "データアベイラビリティ整理",
      "status": "in_progress",
      "task_id": "13",
      "started_at": "2026-04-30T10:15:42+09:00",
      "retry_count": 0
    }
  ]
}
```

### フィールド定義

| フィールド | 型 | 必須 | 説明 |
|---|---|---|---|
| `run_id` | str | ✓ | scope.json と一致 |
| `orchestrator` | str | ✓ | "market-overview-agent" 等のスキル名（拡張子なし） |
| `started_at` | str (ISO 8601) | ✓ | run_id の開始時刻 |
| `steps[]` | array | ✓ | Step 進捗の配列。Step 開始時に append、終了時に該当 entry を更新 |
| `steps[].step_id` | str | ✓ | `step_1` / `step_2_5` / `step_8_b` 等。subject の Step 番号と対応 |
| `steps[].name` | str | ✓ | TaskCreate の `<topic>` と同じ |
| `steps[].status` | str | ✓ | `pending` / `in_progress` / `completed` / `failed` |
| `steps[].task_id` | str | 任意 | TaskCreate が返した ID（参考情報、hooks は使わない） |
| `steps[].started_at` | str | ✓ | Step を `in_progress` にした時刻 |
| `steps[].completed_at` | str | 任意 | Step を `completed` にした時刻 |
| `steps[].retry_count` | int | 任意 | 失敗 → 再試行のカウント。デフォルト 0 |

### hooks による参照

`tools/hooks/check_task_progression.py` が `{{FACTORY_ROOT}}/work/*/*/task_state.json` から本ファイルを探し、`fill_*.py` / `merge_pptx_v2.py` 起動前に **Step ordering inversion**（`steps[:-1]` の中に `status != "completed"` の entry がある状態）を検出する。違反は exit 2 でブロックされ、stderr に「どの step が違反か」が表示される。task_state.json が存在しない orchestrator は素通り（backward compat）。

### 後方互換

`task_state.json` が存在しないオーケストレーターは hooks で警告のみ（block しない）。新規・改修オーケストレーターから順次本規約に乗せる。

---

## SKILL.md にコピペするテンプレート

各 Step の冒頭に以下を貼る（**Step 番号と name は書き換えること**）:

````markdown
### Step N: <Step の名前>

<!-- source: skills/_common/prompts/step_state_tracking.md (manual sync until D2) -->

**進捗トラッキング（必須）**:

開始時:
```
TaskCreate(
  subject="<orchestrator-slug>: Step N - <topic>",
  description="入力: <...> / 出力: <...> / 完了条件: <...>",
  activeForm="<...ing>"
)
TaskUpdate(taskId=<id>, status="in_progress")
# task_state.json の steps[] に append
```

終了時:
```
TaskUpdate(taskId=<id>, status="completed")
# task_state.json の該当 entry を completed に更新
```

<!-- 以下、本 Step の本処理を記載 -->
````

---

## アンチパターン

- ❌ Step 開始時に TaskCreate を呼ばない（前 Step 完了の hooks チェックが効かない）
- ❌ Step 終了時に TaskUpdate(completed) を呼ばずに次の Step に進む（同上）
- ❌ subject に Step 番号が無い（hooks の正規表現でマッチしない、進捗判定不能）
- ❌ `task_state.json` を更新せず TaskCreate / TaskUpdate だけ呼ぶ（hooks は task_state.json しか読めない）
- ❌ `task_state.json` を作業ディレクトリ外に置く（resume 時に検出できない）
- ❌ description の完了条件が主観的（「いい感じになったら」「だいたい揃ったら」）
- ❌ 失敗時に TaskUpdate(completed) を呼んでしまう（hooks が誤って次 Step を許可、データ不整合の温床）

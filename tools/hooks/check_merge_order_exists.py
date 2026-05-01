#!/usr/bin/env python3
"""PreToolUse hook (Bash matcher) — TODO: implement in Phase B-2.

予定: stdin の JSON を読み、tool_input.command が merge_pptx_v2.py 起動なら
--merge-order 引数を抽出して、指定された merge_order.json の存在を確認する。
存在しなければ exit 2 でブロックし、stderr にエラー理由を書く。

今はスタブで全 Bash 呼び出しを素通りさせる（exit 0）。

設計仕様: tools/hooks/README.md を参照。
"""
import sys

sys.exit(0)

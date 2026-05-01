#!/usr/bin/env python3
"""PostToolUse hook (Bash matcher) — TODO: implement in Phase B-2.

予定: stdin の JSON を読み、tool_input.command が fill_*.py 起動なら --output 引数を
抽出し、生成された PPTX に対して tools/validate_pptx.py を自動実行する。
PPTX が壊れていれば exit 2 で警告し、stderr に validate_pptx の出力を書く。

今はスタブで全 Bash 呼び出しを素通りさせる（exit 0）。

設計仕様: tools/hooks/README.md を参照。
"""
import sys

sys.exit(0)

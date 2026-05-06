#!/usr/bin/env python3
"""
fill_pptx.py — イシューツリーの内容をPowerPointテンプレートに書き込むスクリプト

Usage:
    python scripts/fill_pptx.py --data issue_data.json --output output.pptx

JSON形式:
{
  "main_message": "〜すべき（最大70文字）",
  "chart_title": "〜のテーマ（10〜20文字）",
  "root": "ルートイシューのテキスト",
  "sub1": {
    "label": "ラベルA",
    "title": "詳細のイシュー文（〜すべきか？）",
    "hypothesis": "仮説テキスト",
    "children": [
      "サブイシュー1-1のイシュー文（〜すべきか？）",
      "サブイシュー1-2のイシュー文（〜すべきか？）",
      "サブイシュー1-3のイシュー文（〜すべきか？）"
    ]
  },
  "sub2": { ... },
  "sub3": { ... }
}
"""

import argparse
import os
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# brand_resolver bootstrap (passive --brand acceptance until brand-aware migration)
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(SKILL_DIR, "..", "_common", "lib"))
from brand_resolver import add_brand_arg  # noqa: E402

SKILL_DIR = Path(__file__).parent.parent
TEMPLATE_PATH = SKILL_DIR / "assets" / "IssueTree.pptx"
SCRIPTS_DIR = Path("/mnt/skills/public/pptx/scripts")


def xml_escape(text):
    """XML特殊文字をエスケープする"""
    return (text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&apos;"))


def extract_end_para_rpr(para):
    """
    endParaRPr要素を取得する（自己閉じタグと子要素付きの両方に対応）
    """
    # 子要素付き: <a:endParaRPr ...>...</a:endParaRPr>
    m = re.search(r'<a:endParaRPr\b.*?</a:endParaRPr>', para, re.DOTALL)
    if m:
        return m.group(0)
    # 自己閉じ: <a:endParaRPr ... />
    m = re.search(r'<a:endParaRPr\b[^>]*/>', para, re.DOTALL)
    if m:
        return m.group(0)
    return ''


def replace_paragraph_plain(content, old_text, new_text):
    """段落内テキストをプレーンテキストで置換する（単一ラン）"""
    escaped_new = xml_escape(new_text)
    paragraphs = re.findall(r'(<a:p>.*?</a:p>)', content, re.DOTALL)
    for para in paragraphs:
        texts = re.findall(r'<a:t[^>]*>(.*?)</a:t>', para)
        combined = ''.join(texts)
        if combined.strip() == old_text:
            runs = re.findall(r'(<a:r>.*?</a:r>)', para, re.DOTALL)
            if not runs:
                continue
            first_run = runs[0]
            new_run = re.sub(r'<a:t[^>]*>.*?</a:t>',
                             '<a:t>' + escaped_new + '</a:t>',
                             first_run, count=1)
            pre_runs = re.split(r'<a:r>.*?</a:r>', para, flags=re.DOTALL)[0]
            end_rpr = extract_end_para_rpr(para)
            new_para = pre_runs + new_run + end_rpr + '</a:p>'
            content = content.replace(para, new_para, 1)
    return content


def replace_paragraph_with_bold_label(content, old_text, label, detail):
    """
    段落内テキストを「ボールドラベル：」＋「通常テキスト」の2ラン構成で置換する。
    2階層目のサブイシュー（サブイシュー1/2/3）に使用。
    """
    escaped_label = xml_escape(label + "：")
    escaped_detail = xml_escape(detail)

    paragraphs = re.findall(r'(<a:p>.*?</a:p>)', content, re.DOTALL)
    for para in paragraphs:
        texts = re.findall(r'<a:t[^>]*>(.*?)</a:t>', para)
        combined = ''.join(texts)
        if combined.strip() == old_text:
            runs = re.findall(r'(<a:r>.*?</a:r>)', para, re.DOTALL)
            if not runs:
                continue

            # 1つ目のランのrPr属性を取得（フォント・サイズ等のベース）
            first_rpr_match = re.search(r'<a:rPr([^>]*?)>', runs[0], re.DOTALL)
            rpr_attrs = first_rpr_match.group(1) if first_rpr_match else \
                ' lang="ja-JP" altLang="en-US" sz="1400" dirty="0"'
            # b="..." 属性を除去してからbold用に付与
            rpr_attrs_clean = re.sub(r'\s*b="[^"]*"', '', rpr_attrs)

            # solidFill要素を取得
            fill_match = re.search(r'(<a:solidFill>.*?</a:solidFill>)', runs[0], re.DOTALL)
            fill_xml = (fill_match.group(1)
                        if fill_match
                        else '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>')

            # ボールドラン（ラベル＋読点）
            bold_run = (
                '<a:r>'
                '<a:rPr' + rpr_attrs_clean + ' b="1">'
                + fill_xml +
                '</a:rPr>'
                '<a:t>' + escaped_label + '</a:t>'
                '</a:r>'
            )

            # 通常ラン（詳細テキスト）
            normal_run = (
                '<a:r>'
                '<a:rPr' + rpr_attrs_clean + '>'
                + fill_xml +
                '</a:rPr>'
                '<a:t>' + escaped_detail + '</a:t>'
                '</a:r>'
            )

            # endParaRPrを保持（自己閉じ・子要素付きの両方に対応）
            end_rpr = extract_end_para_rpr(para)
            pre_runs = re.split(r'<a:r>.*?</a:r>', para, flags=re.DOTALL)[0]
            new_para = pre_runs + bold_run + normal_run + end_rpr + '</a:p>'
            content = content.replace(para, new_para, 1)
    return content


def fill_template(data, output_path):
    """テンプレートにデータを書き込む"""
    with tempfile.TemporaryDirectory() as tmpdir_str:
        tmpdir = Path(tmpdir_str)
        tmp_pptx = tmpdir / "work.pptx"
        shutil.copy(TEMPLATE_PATH, tmp_pptx)
        unpacked = tmpdir / "unpacked"

        result = subprocess.run(
            ["python", str(SCRIPTS_DIR / "office" / "unpack.py"), str(tmp_pptx), str(unpacked)],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            sys.stderr.write("Error unpacking: " + result.stderr + "\n")
            sys.exit(1)

        slide_path = unpacked / "ppt" / "slides" / "slide1.xml"
        content = slide_path.read_text(encoding="utf-8")

        # --- プレーン置換（ルートイシュー・3階層目・仮説・Main Message・Chart Title）---
        plain_replacements = [
            ("Main Message",       data.get("main_message", "")),
            ("Chart Title",        data.get("chart_title", "")),
            ("ルートイシュー",     data["root"]),
            # 3階層目はラベルなし・プレーンテキスト
            ("サブイシュー1-1",    data["sub1"]["children"][0]),
            ("サブイシュー1-2",    data["sub1"]["children"][1]),
            ("サブイシュー1-3",    data["sub1"]["children"][2]),
            ("サブイシュー2-1",    data["sub2"]["children"][0]),
            ("サブイシュー2-2",    data["sub2"]["children"][1]),
            ("サブイシュー2-3",    data["sub2"]["children"][2]),
            ("サブイシュー3-1",    data["sub3"]["children"][0]),
            ("サブイシュー3-2",    data["sub3"]["children"][1]),
            ("サブイシュー3-3",    data["sub3"]["children"][2]),
            ("サブイシュー1仮説：", "仮説：" + data["sub1"]["hypothesis"]),
            ("サブイシュー2仮説：", "仮説：" + data["sub2"]["hypothesis"]),
            ("サブイシュー3仮説：", "仮説：" + data["sub3"]["hypothesis"]),
        ]
        for old, new in plain_replacements:
            content = replace_paragraph_plain(content, old, new)

        # --- ボールドラベル置換（2階層目のサブイシュー）---
        for key in ("sub1", "sub2", "sub3"):
            placeholder = "サブイシュー" + key[-1]  # "サブイシュー1" / "2" / "3"
            content = replace_paragraph_with_bold_label(
                content,
                placeholder,
                data[key]["label"],
                data[key]["title"]
            )

        slide_path.write_text(content, encoding="utf-8")

        result = subprocess.run(
            ["python", str(SCRIPTS_DIR / "office" / "pack.py"),
             str(unpacked), output_path, "--original", str(TEMPLATE_PATH)],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            sys.stderr.write("Error packing: " + result.stderr + "\n")
            sys.exit(1)

        print("✅ PowerPoint出力完了: " + output_path)


def main():
    parser = argparse.ArgumentParser(description="イシューツリーPPTX生成ツール")
    parser.add_argument("--data", required=True, help="JSONデータファイルのパス")
    parser.add_argument("--output", required=True, help="出力PPTXファイルのパス")
    add_brand_arg(parser)  # passive: accepted but ignored until brand migration
    args = parser.parse_args()

    with open(args.data, encoding="utf-8") as f:
        data = json.load(f)

    fill_template(data, args.output)


if __name__ == "__main__":
    main()

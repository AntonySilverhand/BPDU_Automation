#!/usr/bin/env python3
"""
BP Debate Union — Event Preview Excel Generator

Generates a weekly activity preview Excel report.

Usage:
    python generate_preview.py --week 10
    # Or edit INPUT DATA below

All generation scripts write to ../output/ (relative to this script's location).
"""

import os
import argparse
from datetime import datetime

import pandas as pd

# ===================== INPUT DATA =====================
WEEK_NUMBER = 10
ACTIVITIES = [
    {"日期": "2025年3月30日", "时间": "18:20-20:20", "内容": "常规活动", "地点": "博闻楼B-602"},
    {"日期": "2025年4月2日",  "时间": "18:20-20:20", "内容": "常规活动", "地点": "博闻楼B-602"},
    {"日期": "2025年4月4日",  "时间": "18:20-20:20", "内容": "常规活动", "地点": "博闻楼B-602"},
]
# =====================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(SCRIPT_DIR)
OUTPUT_DIR = os.path.join(SKILL_DIR, "output")


def parse_args():
    parser = argparse.ArgumentParser(description="Generate BP DU Event Preview Excel.")
    parser.add_argument("--week", type=int, default=WEEK_NUMBER,
                        help="Week number (e.g. 7 for 第七周)")
    return parser.parse_args()


def make_unique_path(filepath):
    """Append timestamp if file already exists to avoid overwrite."""
    if not os.path.exists(filepath):
        return filepath
    base, ext = os.path.splitext(filepath)
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    return f"{base}_{timestamp}{ext}"


def generate(args):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = []
    for a in ACTIVITIES:
        data.append({
            "社团名称": "BP Debate Union",
            "活动内容": a["内容"],
            "活动地点": a["地点"],
            "开展时间": f"{a['日期']} {a['时间']}"
        })

    df = pd.DataFrame(data)

    # Title row
    title = f"温州商学院2025-2026学年第一学期第{args.week}周社团活动预告"
    df_title = pd.DataFrame([[title, "", "", ""]], columns=df.columns)
    df_final = pd.concat([df_title, df], ignore_index=True)

    filename = f"BP_Debate_Union_第{args.week}周活动预告汇总.xlsx"
    filepath = make_unique_path(os.path.join(OUTPUT_DIR, filename))
    df_final.to_excel(filepath, index=False, engine='openpyxl')

    print(f"Generated: {filepath}")
    print("\nFile contents:")
    print(df_final.to_string(index=False))
    return filepath


if __name__ == "__main__":
    args = parse_args()
    generate(args)

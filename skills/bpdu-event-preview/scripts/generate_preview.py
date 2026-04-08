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
    {"日期": "2025年4月2日",  "时间": "18:20-20:20", "内容": "日常训练", "地点": "博闻楼B-602"},
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

    # Title row text
    title = f"温州商学院2025-2026学年第一学期第{args.week}周社团活动预告"

    filename = f"BP_Debate_Union_第{args.week}周活动预告汇总.xlsx"
    filepath = make_unique_path(os.path.join(OUTPUT_DIR, filename))

    # Write df starting from row 2 (startrow=1 is 0-indexed, so Row 2)
    # This puts headers on Row 2 and data on Row 3+
    df.to_excel(filepath, index=False, engine='openpyxl', startrow=1)

    # Post-processing with openpyxl
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment
    wb = load_workbook(filepath)
    ws = wb.active

    # 1. Add title to Row 1 and merge
    ws.cell(row=1, column=1, value=title)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)

    # 2. Center all cells and 3. Auto-fit columns
    center_align = Alignment(horizontal='center', vertical='center')

    # Calculate widths while centering
    column_widths = {}

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center_align
            if cell.value:
                # Estimate width: Chinese characters count as 2, others as 1
                val_str = str(cell.value)
                length = sum(2 if ord(c) > 127 else 1 for c in val_str)
                col_letter = cell.column_letter
                if length > column_widths.get(col_letter, 0):
                    column_widths[col_letter] = length

    # Apply widths
    for col_letter, width in column_widths.items():
        # Add a little padding
        ws.column_dimensions[col_letter].width = width + 2

    wb.save(filepath)

    print(f"Generated: {filepath}")
    print("\nFile contents (data only):")
    print(df.to_string(index=False))
    return filepath


if __name__ == "__main__":
    args = parse_args()
    generate(args)

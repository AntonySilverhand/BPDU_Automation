#!/usr/bin/env python3
"""
BP Debate Union — Event Preview Excel Validator

Validates an event preview Excel file against school standards.
Identifies problems and proposes fixes. Fixes are NOT applied
automatically — user is asked to approve each fix.

Usage:
    python validate_preview.py "path/to/file.xlsx" --week 10

Exit codes:
    0 = all pass
    1 = problems found (fix proposals printed)
"""

import argparse
import os
import re
import sys

import pandas as pd

VALID_ACTIVITY_TYPES = {"文化沙龙", "日常训练", "常规活动"}
EXPECTED_CLUB_NAME = "BP Debate Union"
TITLE_PATTERN = re.compile(
    r"温州商学院\d{4}-\d{4}学年第[一二三四五六七八九十百\d]+学期第[一二三四五六七八九十百\d]+周社团活动预告"
)
TIME_PATTERN = re.compile(
    r"\d{4}年\d{1,2}月\d{1,2}日 \d{1,2}:\d{2}-\d{1,2}:\d{2}"
)
LOCATION_PATTERN = re.compile(r"博闻楼[B研]-[A-Za-z0-9]+")


def validate(filepath, expected_week=None):
    """Returns (passed: bool, problems: list[dict])."""
    problems = []

    if not os.path.exists(filepath):
        return False, [{"type": "file", "message": f"File not found: {filepath}"}]

    expected_cols = {"社团名称", "活动内容", "活动地点", "开展时间"}

    try:
        # First, read raw (no header) to detect file structure
        df_raw = pd.read_excel(filepath, header=None)

        # Try header=0: row 0 = title, row 1 = column names, row 2+ = data
        df = pd.read_excel(filepath, header=0)
        actual_cols = set(df.columns)
        title_row_idx = 0  # With header=0, title is at df index 0 (raw row 0)

        if actual_cols != expected_cols:
            # Try header=1: row 0 = title (merged), row 1 = column names, row 2+ = data
            df = pd.read_excel(filepath, header=1)
            actual_cols = set(df.columns)
            title_row_idx = 0  # With header=1, title is still at raw row 0

        if actual_cols != expected_cols:
            return False, [{"type": "structure", "message": f"Column headers do not match expected: {expected_cols}. Found: {actual_cols}. File may have an unexpected structure."}]

        # Extract title text from the detected title row
        title_text = str(df_raw.iloc[title_row_idx, 0]) if title_row_idx < len(df_raw) else ""

    except Exception as e:
        return False, [{"type": "file", "message": f"Cannot read Excel file: {e}"}]

    # Ensure we have at least 4 columns
    if df.shape[1] < 4:
        problems.append({
            "type": "structure",
            "field": "columns",
            "severity": "error",
            "message": f"Expected at least 4 columns, found {df.shape[1]}",
            "fix": "Add missing columns: 社团名称, 活动内容, 活动地点, 开展时间"
        })
        return False, problems

    # ---- 1. Title row ----
    if not TITLE_PATTERN.search(title_text):
        problems.append({
            "type": "title",
            "field": "社团名称 (row 1)",
            "severity": "error",
            "message": f"Title row format incorrect: '{title_text}'",
            "fix": f"Row 1 should be: 温州商学院2025-2026学年第一学期第X周社团活动预告"
        })

    # ---- 2. Data rows (df index 1+) ----
    data_rows = df.iloc[1:].reset_index(drop=True)

    if data_rows.empty:
        problems.append({
            "type": "data",
            "field": "rows",
            "severity": "error",
            "message": "No activity data rows found",
            "fix": "Add at least one activity row"
        })

    # ---- 4. Validate each data row ----
    week_mismatches = []
    for i, row in data_rows.iterrows():
        row_num = i + 2  # Excel row number (1-indexed, title is row 1)

        # 社团名称
        if str(row["社团名称"]).strip() != EXPECTED_CLUB_NAME:
            problems.append({
                "type": "data",
                "field": f"社团名称 (row {row_num})",
                "severity": "error",
                "message": f"Expected '{EXPECTED_CLUB_NAME}', found '{row['社团名称']}'",
                "fix": f"Change 社团名称 to '{EXPECTED_CLUB_NAME}' in row {row_num}"
            })

        # 活动内容
        content = str(row["活动内容"]).strip()
        if content not in VALID_ACTIVITY_TYPES:
            problems.append({
                "type": "data",
                "field": f"活动内容 (row {row_num})",
                "severity": "warning",
                "message": f"Activity type '{content}' not in standard list {VALID_ACTIVITY_TYPES}",
                "fix": f"Change 活动内容 to one of: {', '.join(VALID_ACTIVITY_TYPES)}"
            })

        # 活动地点
        location = str(row["活动地点"]).strip()
        # Only validate building+room format when location looks like one.
        # If it contains building-related keywords, warn if spaces are present or format is wrong.
        # Arbitrary locations (e.g., "线上", "博闻楼门口集合") are always accepted.
        if any(kw in location for kw in ("博闻楼", "研", "楼", "室", "馆")):
            if " " in location or not LOCATION_PATTERN.search(location):
                problems.append({
                    "type": "data",
                    "field": f"活动地点 (row {row_num})",
                    "severity": "warning",
                    "message": f"Location '{location}' appears to be building+room but has a space or wrong format",
                    "fix": f"Remove spaces in 活动地点. Format should be: 博闻楼B-606 (no space between building and room)"
                })

        # 开展时间 format
        time_str = str(row["开展时间"]).strip()
        if not TIME_PATTERN.match(time_str):
            problems.append({
                "type": "data",
                "field": f"开展时间 (row {row_num})",
                "severity": "error",
                "message": f"Time format incorrect: '{time_str}'",
                "fix": f"Use format: YYYY年MM月DD日 HH:MM-HH:MM (e.g., 2025年11月12日 18:20-21:00)"
            })

        # Week number match (if expected_week given)
        if expected_week is not None:
            time_match = TIME_PATTERN.search(time_str)
            if time_match:
                date_str = time_match.group()
                # Extract week from title
                week_match = re.search(r"第([一二三四五六七八九十百\d]+)周", title_text)
                if week_match:
                    cn_nums = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
                               "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
                    week_cn = week_match.group(1)
                    try:
                        actual_week = int(week_cn) if week_cn.isdigit() else cn_nums.get(week_cn[0], None)
                        if actual_week != expected_week:
                            week_mismatches.append((row_num, actual_week, expected_week))
                    except (ValueError, KeyError):
                        pass

    if week_mismatches:
        for row_num, actual, expected in week_mismatches:
            problems.append({
                "type": "week_mismatch",
                "field": f"title / row {row_num}",
                "severity": "error",
                "message": f"Title declares week {actual}, but --week={expected} was expected",
                "fix": f"Update title row week number to 第{expected}周, or run with --week {actual}"
            })

    passed = len(problems) == 0
    return passed, problems


def print_report(passed, problems):
    print("\n=== Validation Results ===")
    if passed:
        print("All checks PASSED.")
        return

    errors = [p for p in problems if p.get("severity") == "error"]
    warnings = [p for p in problems if p.get("severity") == "warning"]

    for p in errors + warnings:
        label = "[ERROR]" if p["severity"] == "error" else "[WARNING]"
        print(f"\n{label} {p['message']}")
        print(f"  Location: {p.get('field', 'N/A')}")
        print(f"  Fix: {p.get('fix', 'N/A')}")

    print(f"\n--- Summary ---")
    print(f"Errors: {len(errors)}  |  Warnings: {len(warnings)}")


def apply_fix(problem, filepath):
    """Apply a single fix to the Excel file. Returns description of what was done."""
    import openpyxl
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    field = problem.get("field", "")
    fix = problem.get("fix", "")

    # Parse row number from field string
    row_match = re.search(r"row (\d+)", field)
    if not row_match:
        return "Could not determine row number"

    row_num = int(row_match.group(1))

    # Parse column from fix string
    col_name = None
    new_val = None
    if "社团名称" in fix and "BP Debate Union" in fix:
        col_name = "社团名称"
        new_val = "BP Debate Union"
    elif "活动内容" in fix:
        col_name = "活动内容"
        # Determine which valid type to use based on the fix message context
        # We cycle through valid types; for a smarter fix we'd inspect the current value
        # For now, pick 文化沙龙 as the default safe type
        new_val = "文化沙龙"
    elif "开展时间" in fix and "YYYY年" in fix:
        col_name = "开展时间"
        # Can't auto-determine correct time — flag for user
        return "Cannot auto-fix time — please edit manually"

    if col_name and new_val:
        # Map column name to column index
        col_map = {"社团名称": 1, "活动内容": 2, "活动地点": 3, "开展时间": 4}
        col_idx = col_map.get(col_name, 1)
        ws.cell(row=row_num, column=col_idx, value=new_val)
        wb.save(filepath)
        return f"Set {col_name} in row {row_num} to '{new_val}'"

    return "Fix not implemented for this case"


def main():
    parser = argparse.ArgumentParser(description="Validate BP DU Event Preview Excel.")
    parser.add_argument("xlsx_path", help="Path to the event preview .xlsx file")
    parser.add_argument("--week", type=int, default=None,
                        help="Expected week number (e.g. 10 for 第十周)")
    parser.add_argument("--fix", action="store_true",
                        help="Prompt to apply each proposed fix")
    args = parser.parse_args()

    passed, problems = validate(args.xlsx_path, expected_week=args.week)
    print_report(passed, problems)

    if not passed and args.fix:
        print("\n=== Apply Fixes ===")
        for i, p in enumerate(problems, 1):
            print(f"\n{i}. {p['message']}")
            print(f"   Fix: {p['fix']}")
            response = input(f"   Apply this fix? [y/n/q (quit)]: ").strip().lower()
            if response == "y":
                result = apply_fix(p, args.xlsx_path)
                print(f"   → {result}")
            elif response == "q":
                break

        # Re-validate after fixes
        passed, problems = validate(args.xlsx_path, expected_week=args.week)
        print("\n=== Re-validation ===")
        print_report(passed, problems)

    sys.exit(0 if passed else 1)


if __name__ == "__main__":
    main()

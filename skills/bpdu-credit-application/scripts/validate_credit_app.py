#!/usr/bin/env python3
"""
BP Debate Union — Credit Application Excel Validator

Validates two types of credit application files:
1. 社团学分认证材料上交名单.xlsx  (member submission list)
2. xx级社团学分申请认证表.xlsx     (individual certification form)

Identifies problems and proposes fixes. Fixes are NOT applied
automatically — user is asked to approve each fix.

Usage:
    python validate_credit_app.py "path/to/file.xlsx" [--fix]

Exit codes:
    0 = all pass
    1 = problems found (fix proposals printed)
"""

import argparse
import os
import re
import sys

import pandas as pd


# ===== Required columns =====
MEMBER_LIST_COLS = ["序号", "姓名", "班级", "联系方式"]
AUTH_FORM_COLS   = ["姓名", "班级", "学号", "联系方式", "学分数量", "备注", "活动认证情况"]
VALID_CREDITS    = {0.5, 1.0}
EXPECTED_CLUB    = "BP Debate Union"
PHONE_DIGITS_MIN = 7


def detect_file_type(df):
    """Detect whether this is a member list or auth form based on columns."""
    cols = set(df.columns.astype(str).str.strip())
    if all(c in cols for c in MEMBER_LIST_COLS):
        return "member_list"
    if all(c in cols for c in AUTH_FORM_COLS):
        return "auth_form"
    return "unknown"


def validate_member_list(df, problems):
    """Validate 社团学分认证材料上交名单.xlsx"""
    # 1. All required columns present
    for col in MEMBER_LIST_COLS:
        if col not in df.columns:
            problems.append({
                "type": "missing_column",
                "field": col,
                "severity": "error",
                "message": f"Missing required column: '{col}'",
                "fix": f"Add column '{col}' with appropriate data"
            })

    # 2. No empty cells in required columns
    for col in MEMBER_LIST_COLS:
        if col in df.columns:
            empty_rows = df[df[col].isna() | (df[col].astype(str).str.strip() == "")]
            if not empty_rows.empty:
                row_nums = list(empty_rows.index + 1)
                problems.append({
                    "type": "empty_cell",
                    "field": f"{col} (rows {row_nums})",
                    "severity": "error",
                    "message": f"Empty cells found in column '{col}' at rows: {row_nums}",
                    "fix": f"Fill in the {col} for rows: {row_nums}"
                })

    # 3. Phone numbers — plausibility check
    if "联系方式" in df.columns:
        for idx, val in df["联系方式"].items():
            val_str = str(val).strip()
            digits = re.sub(r"\D", "", val_str)
            if len(digits) < PHONE_DIGITS_MIN:
                problems.append({
                    "type": "invalid_data",
                    "field": f"联系方式 (row {idx+1})",
                    "severity": "error",
                    "message": f"Phone number too short: '{val}' (min {PHONE_DIGITS_MIN} digits)",
                    "fix": f"Correct phone number in row {idx+1}"
                })


def validate_auth_form(df, problems):
    """Validate xx级社团学分申请认证表.xlsx"""
    # 1. All required columns present
    for col in AUTH_FORM_COLS:
        if col not in df.columns:
            problems.append({
                "type": "missing_column",
                "field": col,
                "severity": "error",
                "message": f"Missing required column: '{col}'",
                "fix": f"Add column '{col}' with appropriate data"
            })

    # 2. No empty cells in required columns (except 活动认证情况 which can be blank)
    for col in AUTH_FORM_COLS:
        if col == "活动认证情况":
            continue
        if col in df.columns:
            empty_rows = df[df[col].isna() | (df[col].astype(str).str.strip() == "")]
            if not empty_rows.empty:
                row_nums = list(empty_rows.index + 1)
                problems.append({
                    "type": "empty_cell",
                    "field": f"{col} (rows {row_nums})",
                    "severity": "error",
                    "message": f"Empty cells in required column '{col}' at rows: {row_nums}",
                    "fix": f"Fill in the {col} for rows: {row_nums}"
                })

    # 3. 学分数量 must be 0.5 or 1
    if "学分数量" in df.columns:
        invalid = df[~df["学分数量"].isin({0.5, 1})]
        if not invalid.empty:
            for idx, row in invalid.iterrows():
                problems.append({
                    "type": "invalid_data",
                    "field": f"学分数量 (row {idx+1})",
                    "severity": "error",
                    "message": f"Invalid credit value: '{row['学分数量']}' (must be 0.5 or 1)",
                    "fix": f"Change 学分数量 to 0.5 or 1 in row {idx+1}"
                })

    # 4. 备注 must be "BP Debate Union"
    if "备注" in df.columns:
        for idx, val in df["备注"].items():
            if str(val).strip() != EXPECTED_CLUB:
                problems.append({
                    "type": "invalid_data",
                    "field": f"备注 (row {idx+1})",
                    "severity": "error",
                    "message": f"备注 should be '{EXPECTED_CLUB}', found: '{val}'",
                    "fix": f"Change 备注 to '{EXPECTED_CLUB}' in row {idx+1}"
                })

    # 5. 活动认证情况 must be blank (students do not fill this)
    if "活动认证情况" in df.columns:
        filled = df[df["活动认证情况"].notna() & (df["活动认证情况"].astype(str).str.strip() != "")]
        if not filled.empty:
            for idx, row in filled.iterrows():
                problems.append({
                    "type": "invalid_data",
                    "field": f"活动认证情况 (row {idx+1})",
                    "severity": "error",
                    "message": f"活动认证情况 should be blank (school fills this), found: '{row['活动认证情况']}'",
                    "fix": f"Clear 活动认证情况 cell in row {idx+1} (leave blank)"
                })

    # 6. Phone plausibility
    if "联系方式" in df.columns:
        for idx, val in df["联系方式"].items():
            val_str = str(val).strip()
            digits = re.sub(r"\D", "", val_str)
            if len(digits) < PHONE_DIGITS_MIN:
                problems.append({
                    "type": "invalid_data",
                    "field": f"联系方式 (row {idx+1})",
                    "severity": "error",
                    "message": f"Phone number too short: '{val}'",
                    "fix": f"Correct phone number in row {idx+1}"
                })


def validate(filepath, apply_fix=False):
    """Main validation entry point."""
    problems = []

    if not os.path.exists(filepath):
        return False, [{"type": "file", "message": f"File not found: {filepath}"}]

    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        return False, [{"type": "file", "message": f"Cannot read Excel: {e}"}]

    # Normalize column names (strip whitespace)
    df.columns = df.columns.astype(str).str.strip()

    file_type = detect_file_type(df)
    if file_type == "unknown":
        return False, [{
            "type": "structure",
            "message": "Unrecognized file format. Expected either:\n"
                       "  - 社团学分认证材料上交名单.xlsx (columns: 序号, 姓名, 班级, 联系方式)\n"
                       "  - xx级社团学分申请认证表.xlsx (columns: 姓名, 班级, 学号, 联系方式, 学分数量, 备注, 活动认证情况)",
            "fix": "Check that the correct template was used"
        }]

    if file_type == "member_list":
        validate_member_list(df, problems)
    else:
        validate_auth_form(df, problems)

    passed = len(problems) == 0 and all(p.get("severity") != "error" for p in problems)
    return passed, problems


def print_report(passed, problems):
    print("\n=== Validation Results ===")
    if passed:
        print("All checks PASSED.")
        return

    errors   = [p for p in problems if p.get("severity") == "error"]
    warnings = [p for p in problems if p.get("severity") == "warning"]

    for p in errors + warnings:
        label = "[ERROR]" if p["severity"] == "error" else "[WARNING]"
        print(f"\n{label} {p['message']}")
        print(f"  Field: {p.get('field', 'N/A')}")
        print(f"  Fix:   {p.get('fix', 'N/A')}")

    print(f"\n--- Summary ---")
    print(f"Errors: {len(errors)}  |  Warnings: {len(warnings)}")


def apply_single_fix(problem, filepath):
    """Apply a single fix to an Excel file. Returns description."""
    import openpyxl
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    field = problem.get("field", "")

    # Parse row from field string like "姓名 (row 3)" or "columns (rows [1, 2])"
    row_match = re.search(r"row[s]?\s*(\d+)", str(field))
    if not row_match:
        return "Could not determine row number to apply fix"

    row_num = int(row_match.group(1))
    fix = problem.get("fix", "")

    # Map fix instruction to column
    col_name = None
    new_val = ""

    if "序号" in fix:
        col_name = "序号"
        # Can't auto-calculate row number
        new_val = row_num
    elif "联系方式" in fix and "clear" in fix.lower():
        col_name = "联系方式"
        new_val = ""
    elif "联系方式" in fix:
        col_name = "联系方式"
        new_val = ""
    elif "学分数量" in fix:
        col_name = "学分数量"
        new_val = ""
    elif "备注" in fix and "BP Debate Union" in fix:
        col_name = "备注"
        new_val = "BP Debate Union"
    elif "活动认证情况" in fix:
        col_name = "活动认证情况"
        new_val = ""  # blank

    if col_name:
        col_map = {name: i+1 for i, name in enumerate(AUTH_FORM_COLS)}
        if col_name not in col_map:
            # Member list columns
            col_map = {name: i+1 for i, name in enumerate(MEMBER_LIST_COLS)}

        col_idx = col_map.get(col_name)
        if col_idx:
            ws.cell(row=row_num, column=col_idx, value=new_val)
            wb.save(filepath)
            return f"Set {col_name} in row {row_num} to '{new_val}'"

    return f"Auto-fix not implemented for: {fix}"


def main():
    parser = argparse.ArgumentParser(description="Validate BP DU Credit Application Excel.")
    parser.add_argument("xlsx_path", help="Path to the credit application .xlsx file")
    parser.add_argument("--fix", action="store_true",
                        help="Prompt to apply each proposed fix")
    args = parser.parse_args()

    passed, problems = validate(args.xlsx_path, apply_fix=args.fix)
    print_report(passed, problems)

    if not passed and args.fix:
        print("\n=== Apply Fixes ===")
        for i, p in enumerate(problems, 1):
            print(f"\n{i}. {p['message']}")
            print(f"   Fix: {p['fix']}")
            response = input(f"   Apply this fix? [y/n/q (quit)]: ").strip().lower()
            if response == "y":
                result = apply_single_fix(p, args.xlsx_path)
                print(f"   → {result}")
            elif response == "q":
                break

        # Re-validate
        passed, problems = validate(args.xlsx_path)
        print("\n=== Re-validation ===")
        print_report(passed, problems)

    sys.exit(0 if passed else 1)


if __name__ == "__main__":
    main()

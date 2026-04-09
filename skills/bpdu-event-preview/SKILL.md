---
name: bpdu-event-preview
description: Use when generating or validating weekly club activity preview Excel files for BP Debate Union. Generates properly formatted .xlsx with title row and activity schedule. Also validates existing files for column structure, title format, activity types, time format, and filename. Examples: "generate event preview for week 10", "validate this event preview Excel"
env:
  dependencies:
    - openpyxl
    - pandas
---

# BP DU Event Preview

## Overview

Generates and validates weekly activity preview Excel reports for BP Debate Union. Produces `.xlsx` files matching school format, and validates existing files against submission standards.

## When to Use

- Generating `BP_Debate_Union_第X周活动预告汇总.xlsx` for the upcoming week
- Validating a user-provided event preview file before submission
- Deadline: Friday 12:00 every week

## Exact Validation Standards

See `references/submission_guide.md` for full details.

| Check | Standard |
|-------|----------|
| Columns | 4 required: 社团名称, 活动内容, 活动地点, 开展时间 |
| Row 1 | Title: `温州商学院YYYY-YYYY学年第一学期第X周社团活动预告` |
| 社团名称 | Must be `BP Debate Union` on all data rows |
| 活动内容 | Must be one of: 文化沙龙, 日常训练, 常规活动 |
| 活动地点 | Must include building + room (e.g., 博闻楼B-606) |
| 开展时间 | Format: `YYYY年MM月DD日 HH:MM-HH:MM` |
| Filename | `BP_Debate_Union_第X周活动预告汇总.xlsx` — week in filename must match content |

## Generate a Document

**Script:** `scripts/generate_preview.py`

```bash
# Basic usage (uses defaults in script)
python scripts/generate_preview.py --week 10

# Full CLI usage (no script editing required)
python scripts/generate_preview.py --week 10 \
    --activity "2026年4月14日" "18:20-20:20" "常规活动" "博闻楼B-602" \
    --activity "2026年4月15日" "18:20-20:20" "常规活动" "博闻楼B-604"
```

The `--activity` flag is repeatable and takes 4 arguments: `DATE`, `TIME`, `CONTENT`, and `LOCATION`.

**Output:** `skills/bpdu-event-preview/output/BP_Debate_Union_第X周活动预告汇总.xlsx`

## Validate a Document

**File formats:** `.xlsx` (Excel), `.xls` (legacy Excel)

**Script:** `scripts/validate_preview.py`

```bash
python scripts/validate_preview.py "path/to/file.xlsx" --week 10
```

**Options:**
- `--week X`: Check if the file's content matches the specified week.

**Checks performed:**
1. Column count — reads `df.shape[1]`, must be ≥ 4
2. Title row — regex against `温州商学院\d{4}-\d{4}学年第X学期第X周社团活动预告`
3. 社团名称 — must be exactly `BP Debate Union` on all data rows
4. 活动内容 — must be one of: 文化沙龙, 日常训练, 常规活动
5. 活动地点 — regex `博闻楼[B研]-[A-Za-z0-9]+` must match
6. 开展时间 — regex `\d{4}年\d{1,2}月\d{1,2}日 \d{1,2}:\d{2}-\d{1,2}:\d{2}` must match
7. Week number match — extracts week from title and dates, compares to `--week` argument

**Agent double-check policy:**
- **Creation**: No sub-agent manual scan is needed after running a create script.
- **Validation**: A sub-agent manual scan is ONLY needed when the user explicitly asks to check if there is something wrong with an existing file. In such cases, spawn a subagent (Agent tool, general-purpose type) to independently read the Excel file (e.g., pandas `read_excel`) and verify each result. The subagent should report findings without running the validation script. The main agent then synthesizes both results.

## Submission

- **Email:** wzbcgjxystfwb@163.com
- **Subject:** BP Debate Union 第X周活动预告汇总
- **Deadline:** Friday 12:00
- **Location:** 博闻楼B-606 (standard venue)

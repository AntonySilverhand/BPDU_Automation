---
name: bpdu-event-preview
description: Use when generating or validating weekly club activity preview Excel files for BP Debate Union. Generates properly formatted .xlsx with title row and activity schedule. Also validates existing files for column structure, title format, activity types, time format, and filename. Examples: "generate event preview for week 10", "validate this event preview Excel"
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
python scripts/generate_preview.py --week 10
```

Or edit `INPUT DATA` in the script directly:

```python
WEEK_NUMBER = 10
ACTIVITIES = [
    {"日期": "2025年3月30日", "时间": "18:20-20:20", "内容": "常规活动", "地点": "博闻楼B-602"},
    ...
]
```

**Output:** `skills/bpdu-event-preview/output/BP_Debate_Union_第X周活动预告汇总.xlsx`

Filename collision is handled automatically (timestamp suffix added if file exists).

## Validate a Document

**File formats:** `.xlsx` (Excel), `.xls` (legacy Excel) — standard Office formats, not compressed archives.

**Script:** `scripts/validate_preview.py`

```bash
python scripts/validate_preview.py "path/to/file.xlsx" --week 10
```

**Checks performed:**
1. Column count — reads `df.shape[1]`, must be ≥ 4
2. Title row — regex against `温州商学院\d{4}-\d{4}学年第X学期第X周社团活动预告`
3. 社团名称 — must be exactly `BP Debate Union` on all data rows
4. 活动内容 — must be one of: 文化沙龙, 日常训练, 常规活动
5. 活动地点 — regex `博闻楼[B研]-[A-Za-z0-9]+` must match
6. 开展时间 — regex `\d{4}年\d{1,2}月\d{1,2}日 \d{1,2}:\d{2}-\d{1,2}:\d{2}` must match
7. Week number match — extracts week from title and dates, compares to `--week` argument

**Agent double-check (required after script runs):**
After the script exits, the agent MUST independently read the Excel file in read-only mode (e.g., pandas `read_excel`) and verify each result:
- Visually inspect the 4 columns and confirm they match expected names
- Read each row's values and check them manually
- Confirm the title row format by reading row 1 directly
- Check time format by reading each date string
- Judge whether each PASS/FAIL from the script is actually correct — override the script if it made a wrong call
- The script is a tool, not an authority — the agent's judgment prevails

## Fix with Approval

Use `--fix` to prompt for each proposed fix:

```bash
python scripts/validate_preview.py "path/to/file.xlsx" --week 10 --fix
```

For each problem found, you will be asked:
- `y` — apply the fix
- `n` — skip this fix
- `q` — quit

Fixes are applied one at a time, and the file is re-validated after all changes.

## Submission

- **Email:** wzbcgjxystfwb@163.com
- **Subject:** BP Debate Union 第X周活动预告汇总
- **Deadline:** Friday 12:00
- **Location:** 博闻楼B-606 (standard venue)

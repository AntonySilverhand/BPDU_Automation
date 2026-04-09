---
name: bpdu-credit-application
description: Use when validating student club credit application Excel files for BP Debate Union. Validates both the member submission list and individual certification forms against school requirements. Examples: "validate credit application files", "check if this student qualifies for 1 credit"
env:
  dependencies:
    - openpyxl
    - pandas
---

# BP DU Credit Application Validation

## Overview

Validates student credit application Excel files for BP Debate Union. Checks two file types against school eligibility rules and format requirements. Proposes fixes for problems found.

## When to Use

- Verifying submitted credit application files before school deadline
- Checking if students meet eligibility requirements
- Validating Excel file format and required columns
- Fixing identified problems before resubmission

## Exact Validation Standards

See `references/eligibility_rules.md` for full details.

### Eligibility Requirements

| Student Type | Activities Required | Hours Required | Credits |
|--------------|---------------------|----------------|---------|
| 本科 (Undergrad) | 36 | 90 | 1 |
| 专科 0.5 credit | 8 | 16 | 0.5 |
| 专科 1 credit | 16 | 32 | 1 |

Activity counts can combine across: 1 club, 2 clubs, club+dissolved club, etc.

### File 1 — 社团学分认证材料上交名单.xlsx

| Column | Requirement |
|--------|-------------|
| 序号 | Required — sequential row numbers |
| 姓名 | Required — no blanks |
| 班级 | Required — no blanks |
| 联系方式 | Required — phone number, 7+ digits, no blanks |

### File 2 — xx级社团学分申请认证表.xlsx

| Column | Requirement |
|--------|-------------|
| 姓名 | Required — no blanks |
| 班级 | Required — no blanks |
| 学号 | Required — no blanks |
| 联系方式 | Required — phone, 7+ digits |
| 学分数量 | Must be exactly 0.5 or 1 |
| 备注 | Must contain `BP Debate Union` |
| 活动认证情况 | Must be **blank** (school fills this upon approval) |

### Filename Conventions

- List: `BP_Debate_Union_社团学分认证材料上交名单.xlsx`
- Individual: `24级BP_Debate_Union_学分申请认证表_姓名.xlsx`

## Validate a File

**File formats:** `.xlsx` (Excel), `.xls` (legacy Excel)

**Script:** `scripts/validate_credit_app.py`

```bash
python scripts/validate_credit_app.py "path/to/file.xlsx"
```

**Checks performed:**
1. File type detection — matches columns against known schema (member list vs auth form)
2. Missing columns — checks each required column is present
3. Empty required cells — checks `isna()` or blank string on all required columns
4. 学分数量 — value must be exactly 0.5 or 1
5. 备注 — cell must contain `BP Debate Union`
6. 活动认证情况 — cell must be blank (students do not fill this — school does)
7. Phone plausibility — strips non-digits, checks remaining digits ≥ 7

**Agent double-check policy:**
- **Creation**: No sub-agent manual scan is needed after running a create script.
- **Validation**: A sub-agent manual scan is ONLY needed when the user explicitly asks to check if there is something wrong with an existing file. In such cases, spawn a subagent (Agent tool, general-purpose type) to independently read the Excel file (e.g., pandas `read_excel`) and verify each result. The subagent should report findings without running the validation script. The main agent then synthesizes both results.

## Submission

**Email:** wzbcgjxystfwb@163.com
**Subject:** BP Debate Union 学分认证材料
**Hard Copy:** 博闻楼B-109
**Deadline:** Check school announcement (typically within 48h of notice)

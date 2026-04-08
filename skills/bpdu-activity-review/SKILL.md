---
name: bpdu-activity-review
description: Use when generating or validating post-event activity review documents for BP Debate Union. Generates .docx with 2-3 embedded photos and 50-100 word third-person description. Also validates existing documents for photo count, word count, Chinese ratio, time reference, and perspective. Examples: "generate activity review for week 7", "validate this activity review document"
env:
  dependencies:
    - python-docx
    - tabulate
---

# BP DU Activity Review

## Overview

Generates and validates weekly club activity review Word documents for BP Debate Union. Produces `.docx` files with embedded photos and third-person descriptions, and validates existing files against school standards.

## When to Use

- Generating `国际学院BP_Debate_Union_第X周活动剪影.docx` after club activities
- Validating a user-provided activity review document before submission
- Deadline: Sunday 22:00 each week

## Exact Validation Standards

See `references/requirements.md` for full details.

| Check | Standard |
|-------|----------|
| Photo count | 2–3 embedded images |
| Word count | 50–100 words (WPS-style: Chinese chars = 1 word, English words = 1 word) |
| Chinese ratio | >50% of characters are Chinese |
| Time reference | Must use last-week format (上周三, 上周五, etc.) |
| Perspective | Third-person only (no 我们/我社/笔者/本人) |

## Generate a Document

**Script:** `scripts/generate_review.py`

```bash
# Recommended: Let the script generate the description automatically
python scripts/generate_review.py \
    --week 7 \
    --date "11月12日" \
    --location "博闻楼B-606" \
    --activity "苏格拉底式研讨会" \
    --participants 20 \
    --topic "社交媒体的continuing the spark现象" \
    --photos "photo1.jpg" "photo2.jpg"

# Or provide a custom description
python scripts/generate_review.py --week 7 --description "..." --photos "p1.jpg" "p2.jpg"
```

**Output:** `skills/bpdu-activity-review/output/国际学院BP_Debate_Union_第X周活动剪影.docx`

Notes:
- Provide 2–3 photo paths. Minimum 2 required.
- If `--description` is omitted, one is automatically generated using the other arguments.
- Filename collision is handled automatically.

## Validate a Document

**File formats:** `.docx` (Word), `.doc` (legacy Word)

**Script:** `scripts/validate_review.py`

```bash
python scripts/validate_review.py "path/to/document.docx"
```

**Checks performed:**
1. Photo count — counts embedded image parts in the docx, must be 2–3
2. Word count — WPS-style: Chinese chars = 1 word, English tokens = 1 word, must be 50–100
3. Chinese ratio — Chinese chars / total chars, must be >50%
4. Time reference — regex search for `上周[一二三四五六日]` pattern in text
5. Third-person perspective — scans for 我们/我社/笔者/本人

Exit code 0 = all pass, 1 = one or more failures.

**Agent double-check policy:**
- **Creation**: No sub-agent manual scan is needed after running a create script.
- **Validation**: A sub-agent manual scan is ONLY needed when the user explicitly asks to check if there is something wrong with an existing file. In such cases, spawn a subagent (Agent tool, general-purpose type) to independently read the file (e.g., using `python-docx` in read-only mode) and verify each result. The subagent should report findings without running the validation script. The main agent then synthesizes both results.

## Submission

- **To:** Assigned student affairs 干事 (varies by week)
- **Deadline:** Sunday 22:00
- **Format:** Word document (.docx)

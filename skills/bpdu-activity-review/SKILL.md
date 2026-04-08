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
python scripts/generate_review.py \
    --week 7 \
    --date "11月12日" \
    --location "博闻楼B-606" \
    --activity "苏格拉底式研讨会" \
    --participants 20 \
    --topic "社交媒体的continuing the spark现象" \
    --description "11月12日，BP Debate Union在博闻楼B-606开展了..." \
    --photos "photo1.jpg" "photo2.jpg" "photo3.jpg"
```

Or edit the `INPUT DATA` constants at the top of the script directly, then run:

```bash
python scripts/generate_review.py
```

**Output:** `skills/bpdu-activity-review/output/国际学院BP_Debate_Union_第X周活动剪影.docx`

Notes:
- Provide 2–3 photo paths. Minimum 2 required.
- The description should already contain a last-week time reference (e.g., 上周三) matching the event date.
- Filename collision is handled automatically (timestamp suffix added if file exists).

## Validate a Document

**File formats:** `.docx` (Word), `.doc` (legacy Word) — standard Office formats, not compressed archives.

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

**Agent double-check (required after script runs):**
After the script exits, the agent MUST independently read the file in read-only mode and verify each result:
- Count photos by inspecting the document inline
- Read the description text and manually count Chinese chars and English tokens to verify word count
- Confirm no first-person language by scanning the text
- Judge whether each PASS/FAIL from the script is actually correct — override the script if it made a wrong call
- The script is a tool, not an authority — the agent's judgment prevails

## Submission

- **To:** Assigned student affairs 干事 (varies by week)
- **Deadline:** Sunday 22:00
- **Format:** Word document (.docx)

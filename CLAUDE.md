# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

BPDU Automation automates weekly administrative documents for a Chinese university BP Debate Union. Three document types:

- **Event Preview (活动预告)**: Weekly Excel schedule submitted Friday 12:00
- **Activity Review (活动剪影)**: Weekly Word document with photos submitted Sunday 22:00
- **Credit Application (学分申请)**: Per-deadline Excel certification forms

Email destination for all submissions: `wzbcgjxystfwb@163.com`

## Python Environment

```bash
source venv/bin/activate
```

Dependencies are in `venv/`: `openpyxl`, `pandas`, `python-docx`. No `requirements.txt` or `pyproject.toml` exists.

## Skills Architecture

Each skill lives in `skills/<skill-name>/` with a consistent layout:
- `SKILL.md` — skill definition (YAML frontmatter `name` + `description` registers it with Claude Code)
- `scripts/` — generation and validation Python scripts
- `references/` — submission guides and requirements
- `examples/` — example outputs
- `output/` — generated files (gitignored)

### Skill Names (use with `/<skill-name>` or `Skill` tool)

| Skill | Purpose |
|-------|---------|
| `bpdu-event-preview` | Generate/validate weekly activity preview Excel |
| `bpdu-activity-review` | Generate/validate weekly activity review Word doc |
| `bpdu-credit-application` | Validate credit application Excel forms |

### Document Generation Quality
- **Uniqueness**: Never rely on the script's default templates for Activity Reviews. Always generate a unique, context-aware description (50-100 words) that reflects the specific topic of the week to avoid repetitive submissions.
- **Language**: Use professional, third-person Chinese. Use terms like "批判性思维" (critical thinking) and "逻辑推演" (logical deduction).

## Common Commands

```bash
# Event Preview
python skills/bpdu-event-preview/scripts/generate_preview.py --week 10
python skills/bpdu-event-preview/scripts/validate_preview.py "file.xlsx" --week 10

# Activity Review
python skills/bpdu-activity-review/scripts/generate_review.py --week 7 --photos "p1.jpg" "p2.jpg"
python skills/bpdu-activity-review/scripts/validate_review.py "document.docx"

# Credit Application
python skills/bpdu-credit-application/scripts/validate_credit_app.py "file.xlsx"
```

## Important Notes

- `Club Maintenance.zip` and `Manage Events.zip` are source reference materials — do not modify
- Generated output files land in `skills/<skill>/output/`
- For activity review validation, photo count is checked via `doc.part.rels`; word count uses WPS-style counting (Chinese chars=1, English words=1)

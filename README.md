# BPDU Automation

> 中文版：[README_CN.md](./README_CN.md)

A Claude Code skill set for automating weekly administrative documents for the BP Debate Union.

## Document Workflows

### Event Preview
Weekly activity schedule Excel file.

- **Due**: Friday 12:00 every week
- **Send to**: `wzbcgjxystfwb@163.com`

### Activity Review
Weekly Word document with embedded photos and third-person descriptions.

- **Due**: Sunday 22:00 every week
- **Send to**: Student affairs staff (干事)

### Credit Application
Club credit certification Excel forms.

- **Due**: Per announcement
- **Send to**: `wzbcgjxystfwb@163.com`

## Quick Start

```bash
# Activate Python environment
source venv/bin/activate

# Generate event preview
python skills/bpdu-event-preview/scripts/generate_preview.py --week 10

# Generate activity review
python skills/bpdu-activity-review/scripts/generate_review.py --week 7 --photos "p1.jpg" "p2.jpg"

# Validate outputs
python skills/bpdu-event-preview/scripts/validate_preview.py "file.xlsx" --week 10 --fix
python skills/bpdu-activity-review/scripts/validate_review.py "document.docx"
python skills/bpdu-credit-application/scripts/validate_credit_app.py "file.xlsx" --fix
```

## Skills

Invoke skills via Claude Code's `/` command:

| Skill | Description |
|-------|-------------|
| `/bpdu-event-preview` | Generate or validate weekly event preview Excel |
| `/bpdu-activity-review` | Generate or validate weekly activity review Word doc |
| `/bpdu-credit-application` | Validate credit application Excel forms |

## Project Structure

```
skills/
├── bpdu-event-preview/       # Event preview Excel generation/validation
├── bpdu-activity-review/       # Activity review Word doc generation/validation
└── bpdu-credit-application/  # Credit form validation
    ├── SKILL.md               # Skill definition
    ├── scripts/               # Python scripts
    ├── references/            # Submission guides
    ├── examples/             # Example outputs
    └── output/               # Generated files
```

## Note

Always independently verify script outputs — scripts are tools, not authorities.

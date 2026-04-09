# Event Preview Examples

## Example Output

| 社团名称 | 活动内容 | 活动地点 | 开展时间 |
|----------|----------|----------|----------|
| 温州商学院2025-2026学年第一学期第九周社团活动预告 | | | |
| BP Debate Union | 常规活动 | 博闻楼B-606 | 2025年11月12日 18:20-21:00 |
| BP Debate Union | 常规活动 | 博闻楼B-606 | 2025年11月14日 18:20-21:00 |

## Example Input Data

- Week: 9
- Activities:
  - 2025年11月12日 18:20-21:00, 常规活动, 博闻楼B-606
  - 2025年11月14日 18:20-21:00, 常规活动, 博闻楼B-606

## Output Filename

```
BP_Debate_Union_第九周活动预告汇总.xlsx
```

## Excel Cell Properties

The generated `.xlsx` applies the following formatting:

| Area | Property | Value |
|------|----------|-------|
| Row 1 (title) | Merged cells | A1:D1 |
| Row 1 (title) | Font | 等线, size 22, not bold |
| Row 1 (title) | Alignment | Center horizontal, center vertical |
| Row 1 (title) | Row height | 30 pt |
| Row 2 (headers) | Font | 等线, size 11, not bold |
| Row 2 (headers) | Alignment | Center horizontal, center vertical |
| Row 3+ (data) | Font | 等线, size 11 |
| Row 3+ (data) | Alignment | Center horizontal, center vertical |
| All cells | Border | Thin border on all four sides |
| All columns | Width | Auto-fit (Chinese chars × 2 + padding) |

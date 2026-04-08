# BPDU 自动化工具

> English version: [README.md](./README.md)

用于自动化生成 BP Debate Union 每周行政文档的 Claude Code 技能集。

## 文档工作流

### 活动预告
每周活动安排 Excel 表。

- **提交时间**：每周五 12:00
- **发送至**：`wzbcgjxystfwb@163.com`

### 活动剪影
每周活动 Word 文档，包含照片和第三人称描述。

- **提交时间**：每周日 22:00
- **发送至**：学生干事

### 学分申请
学生社团学分认证 Excel 表。

- **提交时间**：按通知
- **发送至**：`wzbcgjxystfwb@163.com`

## 快速开始

```bash
# 激活 Python 环境
source venv/bin/activate

# 生成活动预告
python skills/bpdu-event-preview/scripts/generate_preview.py --week 10

# 生成活动剪影
python skills/bpdu-activity-review/scripts/generate_review.py --week 7 --photos "p1.jpg" "p2.jpg"

# 验证输出
python skills/bpdu-event-preview/scripts/validate_preview.py "file.xlsx" --week 10 --fix
python skills/bpdu-activity-review/scripts/validate_review.py "document.docx"
python skills/bpdu-credit-application/scripts/validate_credit_app.py "file.xlsx" --fix
```

## 技能

使用 Claude Code 的 `/` 命令调用技能：

| 技能 | 说明 |
|------|------|
| `/bpdu-event-preview` | 生成或验证每周活动预告 Excel |
| `/bpdu-activity-review` | 生成或验证每周活动剪影 Word 文档 |
| `/bpdu-credit-application` | 验证学分申请 Excel 表 |

## 项目结构

```
skills/
├── bpdu-event-preview/       # 活动预告 Excel 生成/验证
├── bpdu-activity-review/      # 活动剪影 Word 文档生成/验证
└── bpdu-credit-application/  # 学分申请表验证
    ├── SKILL.md               # 技能定义
    ├── scripts/               # Python 脚本
    ├── references/           # 提交指南
    ├── examples/             # 示例输出
    └── output/               # 生成的文件
```

## 注意

请始终独立验证脚本输出——脚本是工具，不是权威。

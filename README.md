# 毕业论文规范word文档

中文学术格式文档排版技能（Skill），AI Agent 或其他工具可直接调用，产出统一规范的学术排版成品。适用于毕业论文、学位论文、课程作业、读书报告、实验报告、结课论文等场景。

## 排版规范

- **页面**：A4，四边 2.5cm；页脚居中阿拉伯页码
- **题名**：黑体小二（18pt）加粗居中
- **作者行**：楷体小四（12pt）居中
- **摘要/关键词**：引题黑体小五，内容仿宋小五
- **一级标题**：黑体小四（12pt）加粗；**二级标题**：黑体五号（10.5pt）加粗；标题黑色、无下划线
- **正文**：宋体五号（10.5pt），两端对齐、1.5 倍行距、首行缩进 2 字符
- **表格**：三线表（顶底线 1pt、表头下线 0.5pt、无竖线），表头加粗居中、跨页自动重复、行不跨页
- **参考文献**：[1][2]… 原生自动编号，悬挂缩进 2 字符
- 完整规范见 [references/format-spec.md](references/format-spec.md)

## 使用办法

### 方式一：AI Agent 直接调用

在支持 Skill 的 AI 助手（如豆包办公）中直接说：

> 用毕业论文规范格式写一篇《题目》……

技能会自动唤起并按上述规范排版输出。

### 方式二：脚本执行（供其他工具 / 程序调用）

依赖：`python-docx`（`pip install python-docx`）

```bash
python scripts/build_academic_docx.py --spec spec.json --out 输出.docx
```

`spec.json` 示例：

```json
{
  "title": "论文题目",
  "author": "姓名　学号　第二次作业",
  "abstract": "摘要文本（可选）",
  "keywords": "关键词1；关键词2（可选）",
  "sections": [
    {"level": 1, "text": "1　引言"},
    {"level": 0, "text": "正文段落……"},
    {"table": 0},
    {"level": 2, "text": "2.1　小节"}
  ],
  "tables": [
    {
      "caption": "表1　标题",
      "headers": ["列1", "列2", "列3"],
      "rows": [["a", "b", "c"]],
      "align": ["center", "left", "left"],
      "widths_cm": [2.8, 6.6, 6.6]
    }
  ],
  "references": ["文献1……", "文献2……"]
}
```

字段说明：

- `sections` 为有序内容块：`level` 0=正文段、1=一级标题、2=二级标题；`{"table": n}` 在当前位置插入 `tables[n]`。
- `tables` 自动渲染为三线表；`align` / `widths_cm` 可省略，缺省时自动处理。
- `references` 自动生成 `[1] [2] …` 编号，无需手工编号。
- 省略的字段留空即可。

### 安装到其他环境

将本仓库目录（`SKILL.md`、`scripts/`、`references/`）放入 AI 环境的技能目录（如 `workspace/.user_skills/`）即可被识别调用。

# academic-word-format

中文学术格式 Word 文档生成 Skill（AI Agent 可调用）。

- 排版规范：A4、2.5cm 边距；题名黑体小二、作者楷体小四、摘要/关键词仿宋小五；一级标题黑体小四、二级标题黑体五号；正文宋体五号、1.5 倍行距、首行缩进 2 字符；三线表；页脚页码；参考文献 [n] 自动编号。
- 用法：`python scripts/build_academic_docx.py --spec spec.json --out 输出.docx`
- spec 结构见 [SKILL.md](SKILL.md)，完整格式规范见 [references/format-spec.md](references/format-spec.md)。

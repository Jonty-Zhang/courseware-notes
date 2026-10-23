# courseware-notes

一个 Claude Code skill：把大学课堂的 PDF 课件（PPT 导出的讲义）整理成知识点与解题方法，输出两份排版规范的 Word：

- **含例题版**：知识点 + 解题方法卡 + 每种方法 1～2 道关键例题（附课件截图）
- **纯方法版**：知识点 + 解题方法卡，不含例题，用语句写清做法和关键点

两版来自同一份 Markdown 源稿（例题写在 `::: example` 块里，纯方法版构建时删除），所以知识点和方法卡完全一致。

## 设计要点

- **看页面图，不只读文字层。** PPT 导出 PDF 的文字层里，Σ、矢量、上下标常变成乱码，受力图、电路图都是图片。脚本把每页渲染成图并拼成 4 页一张的拼图，逐页阅读。
- **逐页清单保证覆盖。** 先写 `inventory.md`（知识点 K / 方法 M / 例题 E，每条带页码，跳过页也登记），再写正文，每条结论都能回查课件页码。
- **多用图。** 概念示意、受力图、分类对比直接截课件配图放在知识点旁边（两个版本都有），几张小图可以拼成带标签的对比图；例题图只放在含例题版。
- **核对课件里的计算。** 例题数字用 Python 重算，课件算错时按正确值写并加注说明。
- **可编辑公式与规范排版。** pandoc 把 LaTeX 转成 Word 公式，再用 python-docx 套中文学术排版（A4、2.5 cm 页边距、宋体五号、1.5 倍行距、宋体小三/黑体小四标题、楷体图表题、原生表格、页码、自动目录）。
- **不依赖 Office 的结构检查。** `check_docx.py` 检查页面、字体字号、表格段距、公式数是否与源稿一致、图题、重点占比和过程性文字。

调研时参考过的同类项目：[ExamPass-Assistant](https://github.com/wubing2023/exampass-assistant)（页码引用、知识点骨架）、[study-assistant-skills](https://github.com/2362094903-ops/study-assistant-skills)（deep / speedrun 两种讲解模式、方法"识别信号→步骤→检查"写法、章节覆盖审计）、Obsidian 插件 Slide Autosummarize（页面渲染给视觉模型、按坐标裁剪配图）。

## 安装

```bash
git clone https://github.com/Jonty-Zhang/courseware-notes ~/.claude/skills/courseware-notes
brew install poppler pandoc        # 或用系统包管理器安装
pip install -r ~/.claude/skills/courseware-notes/requirements.txt
```

## 使用

在 Claude Code 里把课件 PDF 放进工作目录，然后说"把这些课件整理成知识点和解题方法的 Word"。也可以单独用脚本：

```bash
python3 scripts/render_pages.py lecture.pdf --out work/ch1
python3 scripts/crop_figure.py lecture.pdf crops.txt --out work/figs
python3 scripts/build_docx.py ch1.md ch2.md -o notes.docx --toc --title "课程名 知识点与解题方法" --subtitle "（含例题版）"
python3 scripts/build_docx.py ch1.md ch2.md -o notes_methods.docx --toc --drop-examples
python3 scripts/check_docx.py notes.docx --source ch1.md ch2.md
```

源稿写法见 [references/markdown-conventions.md](references/markdown-conventions.md)，结构模板见 [references/output-template.md](references/output-template.md)。

## License

MIT

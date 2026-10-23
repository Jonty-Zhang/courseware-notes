# courseware-notes

两个 Claude Code skill：把大学课堂的 PDF 课件（PPT 导出的讲义）整理成知识点与解题方法，输出排版规范的 Word。按需手动选择：

| skill | 调用 | 内容 |
|---|---|---|
| `lecture-with-examples` | `/lecture-with-examples` | 知识点 + 解题方法卡 + 每种方法 1～2 道关键例题（带题图） |
| `lecture-no-examples` | `/lecture-no-examples` | 知识点 + 解题方法卡，不含例题，用语句写清做法和关键点 |

两个版本的知识点都配有课件截图，首页都有标题和自动目录。两者共用 `shared/scripts` 和 `shared/references`（各 skill 目录里的 `scripts`、`references` 是指向 `shared/` 的符号链接）。

## 设计要点

- **看页面图，不只读文字层。** PPT 导出的 PDF，文字层里的 Σ、矢量、上下标常变成乱码，受力图、电路图都是图片。脚本把每页渲染成图，并拼成每张 4 页的拼图逐页阅读。
- **逐页清单保证覆盖。** 先写 `inventory.md`（知识点 K / 方法 M / 例题 E，每条带页码，跳过页也登记）再写正文，每条结论都能回查课件页码。
- **多用图。** 概念示意、受力图、分类对比直接截课件配图，放在知识点旁边；几张小图可以拼成带标签的对比图。
- **核对课件里的计算。** 例题数字用 Python 重算，课件算错时按正确值写并加注说明。
- **可编辑公式与规范排版。** pandoc 把 LaTeX 转成 Word 公式，再用 python-docx 套中文学术排版：A4、2.5 cm 页边距、宋体五号、1.5 倍行距、宋体小三/黑体小四标题、楷体图表题、原生表格、页码、标题与自动目录。
- **不依赖 Office 的结构检查。** `check_docx.py` 检查页面、字体字号、表格段距、公式数是否与源稿一致、图题、重点占比和过程性文字。

调研时参考过的同类项目：[ExamPass-Assistant](https://github.com/wubing2023/exampass-assistant)（页码引用、知识点骨架）、[study-assistant-skills](https://github.com/2362094903-ops/study-assistant-skills)（deep / speedrun 两种讲解模式、"识别信号→步骤→检查"的方法写法、章节覆盖审计）、Obsidian 插件 Slide Autosummarize（页面渲染给视觉模型、按坐标裁剪配图）。

## 安装

```bash
git clone https://github.com/Jonty-Zhang/courseware-notes ~/.claude/skill-repos/courseware-notes
ln -s ~/.claude/skill-repos/courseware-notes/lecture-with-examples ~/.claude/skills/lecture-with-examples
ln -s ~/.claude/skill-repos/courseware-notes/lecture-no-examples ~/.claude/skills/lecture-no-examples
brew install poppler pandoc        # 或用系统包管理器安装
pip install -r ~/.claude/skill-repos/courseware-notes/requirements.txt
```

不能用符号链接的环境（如 Windows），可把 `shared/scripts` 和 `shared/references` 复制进两个 skill 目录。

## 使用

在 Claude Code 里把课件 PDF 放进工作目录，输入 `/lecture-with-examples` 或 `/lecture-no-examples`。也可以单独用脚本：

```bash
python3 shared/scripts/render_pages.py lecture.pdf --out work/ch1
python3 shared/scripts/crop_figure.py lecture.pdf crops.txt --out work/figs
python3 shared/scripts/compose_figures.py work/figs/对比.jpg --cols 3 a.jpg:标签A b.jpg:标签B c.jpg:标签C
python3 shared/scripts/build_docx.py ch1.md ch2.md -o notes.docx --toc --title "课程名 知识点与解题方法" --subtitle "（含例题版）"
python3 shared/scripts/build_docx.py ch1.md ch2.md -o notes_methods.docx --toc --title "课程名 知识点与解题方法" --subtitle "（纯方法版）" --drop-examples
python3 shared/scripts/check_docx.py notes.docx --source ch1.md ch2.md
```

源稿写法见 [shared/references/markdown-conventions.md](shared/references/markdown-conventions.md)，结构模板见 [shared/references/output-template.md](shared/references/output-template.md)。

## License

MIT

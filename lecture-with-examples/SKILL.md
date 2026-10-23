---
name: lecture-with-examples
description: 课件整理·含例题版。从大学课堂 PDF 课件（PPT 导出的讲义）中提取知识点和主要解题方法，每种方法配 1～2 道课件里的关键例题（带题图和解题过程），生成排版规范、带标题和目录的 Word。用户输入 /lecture-with-examples，或明确要求"带例题""含例题"的课件知识点/解题方法整理时使用。只要纯方法、不要例题时用 lecture-no-examples；用户没说要哪种时先问。
---

# 课件知识点与解题方法（含例题版）

把一个或多个课件 PDF 整理成一份 Word：**知识点 + 解题方法卡 + 每种方法 1～2 道关键例题**，知识点和例题都配课件截图。

姊妹 skill `lecture-no-examples` 产出不含例题的纯方法版。两者共用 `scripts/` 和 `references/`。用户想两版都要时，先按本 skill 写好源稿（例题放在 `::: example` 块里），纯方法版用同一份源稿加 `--drop-examples` 构建即可，两版的知识点和方法卡逐字一致。

## 为什么必须看页面图，不能只读文字层

PPT 导出的 PDF 文字层通常不可靠：公式里的 Σ、矢量箭头、上下标会变成私用区乱码（如 ``）或被拆成碎片，受力图、电路图、曲线全是图片。只读文字层写出来的公式和结论会出错。所以本流程把每页渲染成图片逐页看，文字层只用来快速定位。

## 工作流程

### 0. 确认范围

确认要处理哪些 PDF、输出到哪个文件夹（默认放在课件旁边）。所有课件合成一份 Word，每份课件对应一个一级标题。用户另有学校或老师的排版要求时，以其为准。

### 1. 渲染课件

```bash
python3 scripts/render_pages.py "课件.pdf" --out work/ch1
```

输出 `pages/`（逐页 PNG）、`sheets/`（每张 4 页的拼图，带页码标签）、`text/`（文字层，仅作索引）和 `manifest.json`。依赖 poppler（`pdftoppm`/`pdftotext`）和 Pillow。

### 2. 逐页阅读，整理清单

按顺序读完每一张拼图，每读完一章就把清单写进 `work/inventory.md`，防止上下文过长被压缩后丢失。格式见 `references/inventory-format.md`：跳过页、知识点 K、方法 M、例题 E，每条带页码，能用图说明的知识点标上"图 pNN"。

覆盖要求：每一页要么被某条 K/M/E 引用，要么列在跳过页里。公式或小字看不清时，单独读 `pages/pNNN.png` 放大确认，不要猜。

### 3. 选例题并核对数字

每种方法选 1～2 道最能体现"怎么想"的例题：取哪个研究对象、为什么对这一点取矩、临界状态怎么判断。每章一般 2～5 道，优先选有题图、步骤完整的。

把选中例题的关键计算用 Python 重算一遍。课件本身也会算错或印错符号：若发现不一致，按正确值写，并在例题末尾加一句"注：课件中写作……，按……计算为……"。课件省略的推导步骤不要自己编具体结果，写"同理由……求出"。

### 4. 截配图

能用图说明的地方就放图，比文字更直观。配图分两类：

- **知识点配图**：概念示意、受力图、几何关系、分类对比（力的平移、各类约束的约束力画法、摩擦角……），放在对应段落后面，例题块外。每章一般 3～9 张。
- **例题配图**：题目的结构简图，放在例题块里。

几张小图需要对照看时，用 `compose_figures.py` 拼成一张带标签的对比图。照片、动画截图、装饰图不放。

```bash
python3 scripts/crop_figure.py "课件.pdf" crops.txt --out work/figs
python3 scripts/compose_figures.py work/figs/约束对比.jpg --cols 3 a.jpg:光滑面 b.jpg:圆柱铰链 …
```

裁剪清单每行 `名字 页码 x0 y0 x1 y1`，坐标是页面宽高的比例，从拼图或单页图上读。脚本按 220 dpi 重新渲染裁剪，默认存 JPEG，并生成 `_review.png`；看一遍审查图，把带进标题文字、边框线或截掉标注的框调好再裁。

### 5. 写 Markdown 源稿

每章一个文件，结构见 `references/output-template.md`，写法约定见 `references/markdown-conventions.md`。每章是"（一）知识点 +（二）解题方法"，方法卡后面紧跟它的例题块：

```markdown
::: example
#### 例 4-2 雨蓬支撑杆（三矩式）〔4-p26〕

题目叙述……

![雨蓬结构](c4_p26.jpg){width=48%}

$$\sum M_A=0:\ \sqrt2F_1-5F=0\ \Rightarrow\ F_1=\frac{5F}{\sqrt2}$$

**要点：** 一句话点出这道题体现的方法。
:::
```

最后加一章"题型与方法总表"，按题型汇总识别信号、首选做法和关键点。

写作要求：

- 方法卡写清"在什么情况下、怎么做、要注意什么"，例题是对方法的示范，不替代方法卡。
- 每个知识点、方法、例题都标课件页码 `〔章-p页〕`。
- 重点用 `==…==` 标成浅蓝色，只标核心结论和关键条件，全文 3%～8%，不超过 20%。
- 源稿里不写文档标题，不写"以下是整理的内容""根据要求"之类的过程性文字。

### 6. 构建 Word

```bash
python3 scripts/build_docx.py ch1.md ch2.md … 总表.md \
    -o "课程名_知识点与解题方法（含例题版）.docx" \
    --toc --title "课程名 知识点与解题方法" --subtitle "（含例题版）"
```

图片路径相对于第一个 Markdown 文件所在目录（或用 `--resource-path` 指定）。首页是居中标题（宋体二号加粗）和版本副标题（楷体小三），下面接自动目录；用户明确不要标题时去掉 `--title`。正文按 word-format 规范排版：A4、2.5 cm 页边距、宋体五号、1.5 倍行距、宋体小三/黑体小四标题、楷体图表题、原生表格、可编辑公式、页码，每章另起一页。

### 7. 检查

```bash
python3 scripts/check_docx.py out.docx --source ch1.md …
```

检查 A4、页边距、标题与正文字体字号、表格单元格段距为 0、公式数与源稿一致（没有残留 `$`）、每张图都有图题、浅蓝重点占比、过程性文字。全部 PASS 才交付。机器装有 LibreOffice 时，可再 `soffice --headless --convert-to pdf` 后逐页看图；没有渲染器时，在交付说明里请用户用 Word/WPS 打开确认，并允许更新域，让目录页码刷新。

### 8. 交付

docx 放到课件文件夹。回复里说明：覆盖了哪些课件，知识点、方法、例题、配图各有多少，发现的课件错误及处理，以及没能验证的地方。

## 文件

- `scripts/render_pages.py`：渲染页面、拼图、文字层。
- `scripts/crop_figure.py`：按比例框批量裁剪配图并生成审查图。
- `scripts/compose_figures.py`：把几张小图拼成一张带标签的对比图。
- `scripts/build_docx.py`：Markdown → 排版好的 docx。
- `scripts/check_docx.py`：不依赖渲染器的结构检查。
- `references/output-template.md`：章节结构、方法卡、例题块模板。
- `references/markdown-conventions.md`：源稿 Markdown 写法。
- `references/inventory-format.md`：逐页清单格式。

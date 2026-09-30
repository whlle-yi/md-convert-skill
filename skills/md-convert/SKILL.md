---
name: md-convert
description: 高质量把 Markdown 转换为 Word(docx)、PDF、HTML。当用户要求"md 转 word / docx / pdf / html"、"导出成 Word"、"把报告/论文/文档转成 XX 格式"、"生成可提交的正式文档"，或任何以 .md 为源、以常见文档格式为目标的转换时使用。中文宋体 + Times New Roman、思源宋体 Heavy 强调（无伪粗体）、三线表、原生公式、自动目录。
---

# md-convert — Markdown 高质量转换

把 Markdown 转成**排版达到交付标准**的 Word / PDF / HTML：受控样式（非 pandoc 默认外观）、中文排版规范（宋体 + Times New Roman、Heavy 字重强调、首行缩进、三线表）、原生公式、可刷新目录。

**不适用**：目标是电子表格 → 用 xlsx 技能；目标是指南/PPT → 用 pptx 技能；只是读取 Office 文档内容 → 用 anydoc-convert 技能。

## 执行规程

按"**源处理 → 确认产出组合 → 按格式路由执行**"三段走，之后验收、交付：

**① 源处理与预检**（只做一次，与产出几种格式无关）：读一遍 Markdown——图片相对路径是否存在、公式 `$`/`$$` 是否成对、YAML 头是否完整、是否学术论文等特定场景；发现硬伤先与用户确认再转。

**② 确认产出组合**：
- 用户已明确要哪些格式（"转 word"、"给我 pdf 和 html"）→ 按说的办，不再追问；
- 未明确 → 询问一次：Word / PDF / HTML / 三件套，并给出建议（正式存档推荐 Word，分发推荐 PDF）；
- 识别为学术论文场景（期刊/学位论文/竞赛论文）→ 直接定为 **XeLaTeX PDF**，不产 Word；
- 用户表达过不想被询问 → 默认 docx 并在交付时说明可补产其他格式。

**③ 按格式路由执行**：每种格式有自己的处理方式和失败分支——
- docx → docx 管线（脚本内自动：生成模板 → pandoc → 修补细节 → 字体内嵌）；
- pdf → 引擎选择：Windows + Word 自动走 Word 引擎；学术论文、公式密集 → `--pdf-engine latex`；兜底 LibreOffice；
- html → html 管线；
- 多种格式用 `--to docx,pdf,html`（或 `all`）一次产出，`-o` 给输出目录。

**④ 验收**：对照下方「转换后验收清单」逐项检查每种产物；失败 → 查 `references/troubleshooting.md` 定位 → 调参数重转，同一问题最多重试两次，仍失败则带着报错信息询问用户。

**⑤ 交付**：报告每个产物的路径与大小，说明需要用户动作的点（如 Word 打开带目录的文档时点一次"更新域"）。

## 快速开始

所有转换通过一个入口脚本完成（在技能目录下）：

```bash
# 转 Word（默认格式）
python scripts/md_convert.py 报告.md -o 报告.docx

# 转 PDF（Windows + Word 自动走 Word COM 引擎，保真度最高）
python scripts/md_convert.py 报告.md -o 报告.pdf

# 转 PDF（学术/公式优先，走 XeLaTeX）
python scripts/md_convert.py 论文.md -o 论文.pdf --pdf-engine latex --number-sections

# 转独立单文件 HTML（CSS 已内嵌，可直接分发）
python scripts/md_convert.py 报告.md -o 报告.html --toc
```

依赖：**pandoc**（必需）+ **python-docx**（必需，`pip install python-docx`）。PDF 的 Word 引擎需要 Windows + Microsoft Word；LaTeX 引擎需要 TeX Live / MiKTeX。缺依赖时脚本会给出可读的安装指引，不要绕过脚本手写 pandoc 命令（会丢失样式体系）。

## 格式决策

| 用户目标 | 命令 | 引擎说明 |
|---|---|---|
| 交付/存档的 Word 文档 | `-o out.docx` 或 `--to docx` | pandoc + 生成式 reference.docx |
| 打印/提交的 PDF（一般文档） | `-o out.pdf` | auto → Word COM（Windows） |
| 学术论文、公式密集文档 | `--to pdf --pdf-engine latex` | XeLaTeX；**论文场景只出 PDF，不产 Word** |
| 网页/在线传阅 | `-o out.html` | 单文件，无外部依赖 |
| 一次要多种 | `--to docx,pdf,html -o 目录/` | 源处理一次，按需扇出 |

Word 引擎 = 先按 docx 管线产出（含全部后处理），再用 Word 更新目录域并导出 PDF，样式与 docx 版完全一致——**需要"Word 版和 PDF 版长一样"时用 auto/word**。

## 排版规范（默认值）

| 项目 | 默认 | 覆盖参数 |
|---|---|---|
| 中文正文 | 宋体 SimSun | `--cjk-font` |
| 西文/数字 | Times New Roman | `--latin-font` |
| 标题/加粗中文 | 思源宋体 Heavy X，**不伪粗** | `--heading-cjk-font` / `--no-heavy` |
| 中文斜体（HTML/PDF） | 楷体，禁用合成斜体 | — |
| 代码 | Consolas + 浅灰底纹，语法高亮默认关（黑白正式文档风格） | `--mono-font`；`--highlight` 开启高亮 |
| 正文 | 12pt（小四）、1.5 倍行距、首行缩进 2 字符 | `--font-size` `--line-spacing` `--no-body-indent` |
| 表格 | 三线表 | `--table-style grid` 换全框线 |
| 目录 | 默认无 | `--toc --toc-depth 3` |
| 页面 | A4，上下 2.54cm、左右 3.17cm | — |

实现要点（改代码前先理解）：
- docx 的样式来自 `scripts/make_reference_docx.py` 程序化生成的 reference.docx，样式名与 pandoc 约定一一对应，不可改名。
- "中文加粗无伪粗体"靠 `md_convert.py::fix_cjk_bold` 后处理：把加粗 run 按中西文拆分，中文段设 Heavy 字体并显式 `w:b=0`。本机没有 Heavy 字体时用 `--no-heavy` 回退。
- 输入格式默认 `markdown+east_asian_line_breaks`（中文硬换行不产生多余空格）；来源是严格 GFM 时加 `--reader gfm`。

## 转换后验收清单

执行规程第 4 步的检查项，交付前逐项确认：

1. **文件能正常打开**，大小合理（>5KB 说明内容非空）。
2. **公式**：docx 中是可编辑的 Word 公式（OMML），不是乱码或图片；PDF 中公式无缺字。
3. **表格**：三线表边框正确（顶/底粗线、表头下细线、无竖线）。
4. **目录**（若有）：Word 打开时会提示"是否更新域"→ 点是后页码正确；PDF（Word 引擎）页码已自动刷新。
5. **字体**：中文无豆腐块；强调文字是 Heavy 字重而非拉伸加粗。
6. **图片**：显示且未变形（超宽图受页面宽度约束）。

发现问题先查 `references/troubleshooting.md`，再查 `references/styles.md` 调样式。

## 样式定制

- 调字体/字号/缩进/行距/表格样式 → 命令行参数（见上表），见 `references/styles.md` 的完整参数说明。
- 需要全新版式（页眉页脚、封面、页码样式）→ 用 `python scripts/make_reference_docx.py -o my-ref.docx ...` 生成基线，在 Word 里手动微调后，用 `pandoc --reference-doc my-ref.docx` 直接使用；把修改固化进 `make_reference_docx.py` 才能进入版本管理。

## 已知边界

- `--number-sections` 在 docx 中依赖 pandoc 的标题编号实现，老版本 pandoc 可能不生效 → 升级 pandoc（≥2.10）。
- Word 打开带目录的 docx 会弹一次"更新域"确认，属预期行为（PDF 引擎已自动处理）。
- 本机未安装"思源宋体 Heavy X"时，docx 中该字体名会回退为 Word 默认字体 → 装字体或用 `--no-heavy`。
- 输出文件被 PDF 阅读器/预览窗格占用时报错退出（码 3）→ 关闭占用程序重试。

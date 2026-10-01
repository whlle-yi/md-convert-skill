# md-convert-skill

[![CI](https://github.com/whlle-yi/md-convert-skill/actions/workflows/ci.yml/badge.svg)](https://github.com/whlle-yi/md-convert-skill/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![pandoc](https://img.shields.io/badge/pandoc-%E2%89%A52.10-black)

> 一个 [ZCode 技能（Skill）](https://code.claude.com/docs/en/skills)：把 Markdown **高质量**转换为 Word（.docx）、PDF、HTML——不是"能转换"，而是排版达到可直接交付的正式文档标准。

**English**: A ZCode skill for high-fidelity Markdown → Word / PDF / HTML conversion, built on pandoc + programmatically generated reference.docx, with CJK typography rules (SimSun + Times New Roman, Heavy-weight emphasis without synthetic bold, three-line academic tables, native equations, refreshable TOC).

## 特性

- **三个独立入口脚本**：`to_docx` / `to_pdf` / `to_html` 按产出格式各司其职，职责单一、可单独调用；
- **受控样式体系**：docx 外观由程序化生成的 reference.docx 决定，每次转换结果一致可复现，不依赖操作系统的 Word 模板。
- **中文排版规范**：正文宋体、西文/数字 Times New Roman；标题与加粗使用思源宋体 Heavy，**通过 run 拆分实现"中文 Heavy 字重 + 西文 TNR Bold"，杜绝宋体伪粗体**；首行缩进 2 字符按字符单位自适应字号。
- **学术级表格**：默认三线表（顶/底 1.5pt、表头下 0.75pt、无竖线），可切换全框线。
- **原生公式**：LaTeX 数学转 Word 原生公式（OMML），可继续编辑；PDF 走 XeLaTeX 或 Word 两条引擎。
- **目录即所得**：docx 的目录域打开时自动提示刷新；PDF 目录由 LaTeX 编译生成，页码即所得。
- **XeLaTeX 学术排版**：PDF 唯一引擎，公式、三线表（booktabs）、字体嵌入，输出跨机器一致。
- **单文件 HTML**：CSS 内嵌、图片内联，直接分发，打印样式对齐纸面排版。
- **AI 友好**：SKILL.md 提供决策表与验收清单，供智能体按规程执行转换。

## 安装

### 依赖

| 依赖 | 必需性 | 说明 |
|---|---|---|
| Python ≥ 3.10 | 必需 | |
| [pandoc](https://pandoc.org/installing.html) ≥ 2.10 | 必需 | `winget install JohnMacFarlane.Pandoc` |
| python-docx | 必需 | `pip install python-docx` |
| Microsoft Word | 可选 | Windows 上 PDF 默认引擎 |
| TeX Live / MiKTeX（XeLaTeX） | PDF 必需 | `to_pdf.py` 的唯一引擎 |

字体：默认方案使用 **宋体（SimSun）、Times New Roman、Consolas**（Windows 自带）与 **思源宋体 Heavy**（`Noto Serif SC Heavy X`，推荐安装；未安装时自动回退，见 [references/styles.md](references/styles.md)）。

### 安装为 ZCode 技能

技能本体在仓库的 `skills/md-convert/` 子目录，整体复制或链接该目录即可：

```bash
# 方式一：junction（推荐，随仓库更新）
git clone https://github.com/whlle-yi/md-convert-skill.git D:/tools/md-convert-skill
mklink /J "%USERPROFILE%\.agents\skills\md-convert" "D:\tools\md-convert-skill\skills\md-convert"

# 方式二：直接复制
xcopy /E /I md-convert-skill\skills\md-convert "%USERPROFILE%\.agents\skills\md-convert"
```

安装后在 ZCode 中说"把这个 md 转成 word"即可自动触发；也可 `/md-convert` 显式调用。

## 快速开始

```bash
pip install python-docx

# 按产出格式选择入口脚本（在技能目录内执行）
python skills/md-convert/scripts/to_docx.py examples/demo.md -o demo.docx
python skills/md-convert/scripts/to_pdf.py  examples/demo.md -o demo.pdf
python skills/md-convert/scripts/to_html.py examples/demo.md -o demo.html

# 学术论文：XeLaTeX + 编号 + 文献（不产 Word，--pdf-engine 已移除，to_pdf.py 即 XeLaTeX）
python skills/md-convert/scripts/to_pdf.py paper.md -o paper.pdf \
    --number-sections --citeproc --bibliography refs.bib
```

完整参数见各脚本的 `--help`（如 `python scripts/to_docx.py --help`）；示例输入与产出在 [examples/](examples/)。

## 转换质量对照

| 排版项 | 一般转换工具 | md-convert |
|---|---|---|
| 中文字体 | 主题默认（等线/雅黑） | 宋体，西文 Times New Roman |
| 加粗中文 | 宋体伪粗体 | 思源宋体 Heavy，无伪粗体 |
| 表格 | 全框线网格 | 三线表（学术规范） |
| 正文段落 | 无缩进 | 首行缩进 2 字符 |
| 公式 | 图片或乱码 | Word 原生 OMML / XeLaTeX |
| 目录 | 静态文本或没有 | Word 目录域，可刷新 |
| 复现性 | 依赖本机模板 | 样式由代码生成，跨机器一致 |

## 项目结构

```text
md-convert-skill/
├── skills/                        # ① 技能区：整体复制/链接即可安装使用
│   └── md-convert/                #    目录名 = 技能名
│       ├── SKILL.md               #   技能定义（触发条件 + 决策表 + 验收清单）
│       ├── scripts/               #   入口 to_docx / to_pdf / to_html
│       │                          #   共享 common + docx_lib；make_reference_docx 模板生成器
│       ├── assets/                #   html-style.css / latex-header.tex
│       └── references/            #   styles.md / troubleshooting.md
├── examples/                      # ② 项目示例与展示样张
├── docs/architecture.md           #   设计决策（面向维护者）
├── tests/                         #   pytest 测试
├── .github/workflows/             #   CI
├── README.md  CHANGELOG.md  LICENSE
└── temp/                          # ③ 本地过程产物（gitignore，不入库）
    ├── output/                    #   手动转换产物
    ├── render/                    #   视觉验收渲染图
    └── pytest/                    #   测试缓存与临时目录
```

## 文档

- [样式体系详解](skills/md-convert/references/styles.md) —— 样式链路、pandoc 样式名对照、伪粗体规避原理、固定 reference.docx
- [故障排查](skills/md-convert/references/troubleshooting.md) —— 按症状索引的排查手册
- [架构与设计决策](docs/architecture.md) —— 为什么是 pandoc + 生成式 reference.docx + 后处理（面向维护者，存于项目说明区）
- [更新日志](CHANGELOG.md)

## 测试

```bash
pip install -r requirements.txt
pytest -v
```

测试覆盖样式生成、三种格式转换的端到端产物校验；PDF 测试需要本机 XeLaTeX，CI 上自动跳过。CI 在 Ubuntu / Windows 双平台运行。

## 开发约定

- 样式改动只落在 `make_reference_docx.py` / `docx_lib.py`（可代码复现，不提交二进制模板）；
  新参数先写进 `references/params.md` 再改脚本。
- 新功能必须带测试（产物断言写入 `temp/pytest/`）；`pytest -v` 全绿再提交。
- 行为变化同步三处：`SKILL.md`（路由/规范）、`references/troubleshooting.md`（故障）、`CHANGELOG.md`。

## 安全注意

不要转换不受信任来源的 Markdown 后直接打开产物：pandoc 默认保留 raw HTML/LaTeX，
如需隔离，加 `--reader markdown-raw_html-raw_tex` 禁用原始内容透传。
转换全程本地完成，文档内容不经任何网络服务（网络图片为用户显式行为）。

## Roadmap

- [ ] PPTX 输出（配合大纲结构）
- [ ] 自定义样式档案（配置文件式 `--style academic/thesis/...`）
- [ ] Word 页眉页脚与封面模板
- [ ] HTML 中文本斜体的字体子集化

## License

[MIT](LICENSE) © 2026 whlle-yi

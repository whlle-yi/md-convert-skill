# 架构与设计决策

本文记录 md-convert 的技术选型理由与实现结构，供维护者与二次开发者阅读。

## 总体管线

```text
                      ┌────────────────────────────────────────────┐
                      │                md_convert.py               │
 input.md ───────────►│  参 数 解 析 / 引 擎 调 度 / 后 处 理        │
                      └──────┬───────────────┬───────────────┬─────┘
                             │               │               │
                     docx    │           pdf │           html │
                             ▼               ▼               ▼
                    pandoc → docx    ┌─ word: docx→COM→pdf   pandoc → html5
                    (reference.docx) ├─ latex: pandoc→xelatex  (+内嵌 CSS)
                             │       └─ libreoffice: docx→pdf
                             ▼
                    python-docx 后处理
                    ├─ style_tables()   三线表/全框线 + 表头加粗
                    ├─ fix_cjk_bold()   中西文加粗拆分（Heavy 无伪粗）
                    └─ enable_update_fields()  目录域自动刷新
```

## 为什么以 pandoc 为核心

Markdown 方言、公式（LaTeX → OMML）、表格、脚注、题注等解析与转换是重活，
pandoc 是该领域事实标准且质量最高：

- **公式**：`$...$` 直接转 Word 原生公式（OMML），可继续编辑；XeLaTeX 路径天然无损。
- **样式契约**：docx 写入器全部外观取自 reference.docx 样式 → 把"排版质量"问题
  完全转化为"样式表质量"问题，可程序化解决。
- 自研解析器（markdown-it → python-docx）的公式/表格保真度长期无法企及，放弃该路线。

## 为什么程序化生成 reference.docx

pandoc 官方做法是让用户手动做一个 reference.docx。其缺陷：不可复现、不可评审、
跨机器漂移。改为**每次转换前从 pandoc 默认种子程序化改写生成**：

1. `pandoc --print-default-data-file reference.docx` 提供种子（保证全部样式存在）；
2. python-docx 改写字体/字号/行距/缩进/颜色 → 与仓库代码同源，版本可控；
3. 需要页面级定制（页眉页脚、封面）时允许"生成基线 + Word 微调 + 固定使用"逃生通道。

## 为什么需要后处理

pandoc 样式体系覆盖不到的三个"质量命门"：

| 问题 | 为什么 pandoc 做不到 | 后处理方案 |
|---|---|---|
| 中文加粗伪粗体 | `Strong` 样式的 `w:b` 对中西文同时生效，宋体无粗体字形 | `fix_cjk_bold`：加粗 run 按 CJK/非 CJK 拆分，中文段设 Heavy 字体并**显式 `w:b=0`**（run 级关闭覆盖样式级），西文段保留 TNR Bold |
| 三线表 | `Table` 样式无法表达"仅首行下边线" | `style_tables`：表级边框 + 首行单元格 `tcBorders` |
| 目录页码 | Word 需排版后才能算页码 | `enable_update_fields`：打开时刷新；PDF 走 Word 引擎时由 COM 先行更新 |

## PDF 引擎调度

| 引擎 | 适用 | 权衡 |
|---|---|---|
| Word COM | Windows + Word，一般文档 | 保真度最高（与 docx 完全同源）；依赖 Office；速度中等 |
| XeLaTeX | 学术、公式密集、跨平台 | 公式/断行质量最好；字体配置复杂；HTML 语义的表格样式受限 |
| LibreOffice | Linux 服务器兜底 | 无 Office 环境的替代；部分样式还原度略低 |

Word COM 用 **PowerShell 脚本**而非 pywin32：Windows 自带 PowerShell，避免引入
`pywin32` 依赖与版本问题；`docx2pdf` 包仅作为兜底存在。

## HTML 路径

pandoc `--embed-resources` 产出真正的单文件；`assets/html-style.css` 用
`@font-face + local() + unicode-range` 实现**按字符**的字体分层——只有 CJK 字符
命中中文特殊字体（强调 Heavy、斜体楷体），西文保持 Times New Roman；
`font-synthesis: weight` 禁止楷体被合成斜体。打印样式（`@page A4`）与纸面排版对齐。

## 依赖策略

运行时仅 `python-docx` + 外部二进制 `pandoc`。刻意**不引入**：
- `pywin32`（改用 PowerShell COM）；
- `weasyprint`/`playwright`（HTML→PDF 重依赖，由 Word/LaTeX 引擎覆盖）；
- LaTeX 模板框架（动态拼接导言即可，保持透明）。

## 已知限制与演进方向

- docx 标题编号依赖 pandoc 实现（≥2.10），个别版本行为漂移 → 必要时自建 numbering.xml；
- 松散列表项会继承正文首行缩进（pandoc 用 Body Text 承载松散列表）→ 可在后处理中按 pStyle 过滤；
- PPTX / 样式档案配置文件在 Roadmap。

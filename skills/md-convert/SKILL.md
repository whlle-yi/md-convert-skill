---
name: md-convert
description: 高质量把 Markdown 转换为 Word、PDF、HTML。用户要求"md 转 word/pdf/html"、"导出成 Word/PDF"、"生成可提交的正式文档/论文"，或任何以 .md 为源转成常见文档格式的场景都使用。中文排版规范：宋体 + Times New Roman、Heavy 强调无伪粗体、三线表、原生公式。
---

# md-convert — Markdown 高质量转换

把 Markdown 转成**排版达到交付标准**的 Word / PDF / HTML。电子表格 → xlsx 技能；演示文稿 → pptx 技能；读取 Office 内容 → anydoc-convert 技能。

## 执行规程

**① 源处理与预检**（只做一次）：读一遍 Markdown，检查图片相对路径、公式 `$`/`$$` 成对、YAML 头完整；发现硬伤先与用户确认再转。

**② 确认产出组合**：
- 用户已明确格式 → 照办，不追问；
- 未明确 → 询问一次，**默认推荐 PDF**；
- 识别为学术论文场景 → 直接定为 PDF（XeLaTeX 引擎，不产 Word），**执行前先读 `references/scenario-academic.md`**；
- 用户不想被询问 → 默认 PDF，交付时说明可补产其他格式。

**③ 按格式路由执行**（多格式用 `--to docx,pdf,html` 一次产出，`-o` 为输出目录）：
- docx → docx 管线（脚本内自动：生成模板 → pandoc → 修补细节 → 字体内嵌）；
- pdf → Windows + Word 自动走 Word 引擎（与 docx 版样式一致）；兜底 LibreOffice；
- html → html 管线；
- 带参考文献 → 追加 `--citeproc --bibliography 文献.bib`（引用键 `[@id]`，支持 BibTeX/CSL YAML）；
- 用户要调整排版 → **先读 `references/params.md`**；来源是严格 GFM 时加 `--reader gfm`。

**④ 验收**：对照下方清单逐项检查每种产物；失败 → 查 `references/troubleshooting.md` 定位，调参重转，同一问题最多两次，仍失败则带报错询问用户。

**⑤ 交付**：报告每个产物的路径与大小，以及需要用户动作的点。

## 转换后验收清单

1. **文件能正常打开**，大小合理（>5KB 说明非空）。
2. **公式**：docx 中是可编辑的 Word 公式（OMML）；PDF 中公式无缺字。
3. **表格**：三线表边框正确（顶/底粗线、表头下细线、无竖线），表内文字单倍行距。
4. **目录**（若有）：Word 打开时点一次"更新域"后页码正确；PDF（Word 引擎）页码已自动刷新。
5. **字体**：中文无豆腐块；强调是 Heavy 字重而非拉伸加粗。
6. **图片**：显示且未变形。

## 文件路由

| 何时读 | 文件 |
|---|---|
| 调排版参数 / 查默认值 | `references/params.md` |
| 学术论文场景 | `references/scenario-academic.md` |
| 转换报错 / 产物异常 | `references/troubleshooting.md` |
| 深度样式定制（固化模板） | `references/styles.md` |

# 更新日志

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 格式，
版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

### Added

- `--citeproc` / `--bibliography` / `--csl`：参考文献处理透传（pandoc citeproc），
  docx / pdf / html 三条管线均可用，服务学术论文路线。

### Changed

- SKILL.md 参数表与学术论文场景封装为 references 文件（params.md、
  scenario-academic.md），执行规程内联路由指针 + 路由总表双层导航；
  SKILL.md 减至 45 行，只保留每次转换必读的规程骨架与验收清单。
- SKILL.md 瘦身：删除与脚本自指引/排障手册/架构文档重复的小节（快速开始、
  格式决策、实现要点、样式定制、已知边界），格式路由表并入执行规程第 ③ 步，
  正文 110 行减至 70 行以内，信息由 references/ 与 docs/ 承接。
- 默认产出格式改为 **PDF**（原为 docx）：产出组合不明确时推荐 PDF，
  `-o` 给目录时也默认落 PDF；Word 改为用户点名时才产出。
- CLI 支持 `--to docx,pdf,html|all` 一次产出多种格式（`-o` 为输出目录）；
  SKILL.md 执行规程改为"源处理 → 确认产出组合 → 按格式路由"三段式，
  产出组合不明确时询问用户，学术论文场景固定走 XeLaTeX PDF。
- SKILL.md 增加「执行规程」：确认输入与格式 → 预检源文件 → 组装命令 →
  验收清单 → 交付，明确执行顺序与失败分支（默认格式、引擎不可用、重试上限）。
- 目录重构：技能本体归入 `skills/md-convert/`（自包含，整体复制/链接即可安装）；
  项目说明、示例、测试与 CI 留在仓库根；新增 `temp/`（gitignore）收纳本地
  过程产物（转换输出、渲染图、pytest 缓存）。
- CI：runner 镜像不再预装 pandoc，改为显式安装；转换产物输出至 `temp/output/`。

## [0.1.0] - 2026-09-30

### Added

- 核心转换管线 `scripts/md_convert.py`：Markdown → docx / pdf / html。
- 受控样式生成器 `scripts/make_reference_docx.py`：程序化生成 pandoc reference.docx
  （宋体 + Times New Roman、思源宋体 Heavy 标题与强调、三线表基准、A4 页面）。
- docx 后处理：中西文加粗拆分（中文 Heavy 无伪粗体）、三线表/全框线边框、
  目录域打开时自动刷新。
- PDF 三引擎调度：Word COM（`scripts/docx2pdf.ps1`，零第三方依赖）→
  XeLaTeX（动态导言，含 `\key` 强调命令）→ LibreOffice 兜底。
- 单文件 HTML 输出：CSS 内嵌、中文楷体斜体映射（禁合成斜体）、打印样式对齐 A4。
- XeLaTeX 便携导言模板 `assets/latex-header.tex`。
- ZCode 技能定义 `SKILL.md`（触发条件、决策表、验收清单）。
- pytest 测试套件（样式生成 + 三格式端到端校验，PDF 引擎缺失自动跳过）。
- GitHub Actions CI（Ubuntu / Windows 双平台）。
- 示例文档与产出 `examples/`，深入文档 `references/`、`docs/`。

[Unreleased]: https://github.com/whlle-yi/md-convert-skill/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/whlle-yi/md-convert-skill/releases/tag/v0.1.0

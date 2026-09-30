# 更新日志

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 格式，
版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

### Changed

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

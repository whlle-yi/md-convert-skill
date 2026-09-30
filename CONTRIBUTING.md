# 参与贡献

感谢关注本项目！提交 Issue 或 PR 前请先阅读本文。

## 开发环境

```bash
git clone https://github.com/whlle-yi/md-convert-skill.git
cd md-convert-skill
pip install -r requirements-dev.txt   # python-docx、pytest
pandoc --version                      # ≥ 2.10
```

PDF 相关用例在无 Word/XeLaTeX 的环境会自动跳过，属正常现象。

## 分支与提交规范

- 分支：`feat/xxx`（功能）、`fix/xxx`（修复）、`docs/xxx`（文档）。
- 提交信息：[Conventional Commits](https://www.conventionalcommits.org/zh-hans/)，
  如 `fix(docx): 三线表在合并单元格时边框缺失`。
- 一个 PR 聚焦一件事，附带转换前后对比（截图或文件）会显著加快评审。

## 改动约定

### 样式相关改动

- 样式修改一律落在 `skills/md-convert/scripts/make_reference_docx.py` 或
  `skills/md-convert/scripts/docx_lib.py` 的后处理中，**不要**提交二进制
  docx 模板——样式必须可由代码复现。
- 技能目录 `skills/md-convert/` 必须保持自包含：技能文件不得引用仓库内
  技能区之外的路径（examples/tests 属项目层，技能运行不依赖）。
- pandoc 样式名（`Body Text`、`Strong`、`Source Code` 等）是契约，不可改名。
- 涉及排版观感的改动请附 `examples/demo.md` 的新产出文件，便于直观评审。

### 代码风格

- Python：类型注解、`from __future__ import annotations`、中文 docstring；
  不引入重型依赖（当前仅 python-docx）。

### 测试

- 新功能须带测试；测试产物经 `tmp_path` 落在 `temp/pytest/`（已 gitignore），
  不得污染仓库其他目录。
- 运行：`pytest -v`；提交前确保全绿（或按预期跳过）。

## 文档

- 用户可见的行为变化须同步更新：`SKILL.md`（决策表/验收清单）、
  `references/troubleshooting.md`（新故障模式）、`CHANGELOG.md`（Unreleased 段）。
- 文档排版遵循项目规范：中文宋体语境、西文 Times New Roman、关键结论加粗。

## Issue

- Bug 报告请附：完整命令、`--verbose` 输出、`pandoc --version`、操作系统、
  最小复现 md 片段。
- 功能建议请说明使用场景与期望产物格式。

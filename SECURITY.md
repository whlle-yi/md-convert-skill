# 安全策略

## 支持版本

| 版本 | 状态 |
|---|---|
| 0.1.x | ✅ 支持 |

## 报告漏洞

请勿以公开 Issue 形式报告安全漏洞。通过 GitHub 私密渠道联系仓库所有者
（whlle-yi），或在具备条件时使用 [私密安全报告](https://github.com/whlle-yi/md-convert-skill/security/advisories/new)。

请附带：影响版本、复现步骤、影响范围评估。收到报告后 7 天内确认，修复周期视严重程度另行沟通。

## 安全设计说明

- 转换在本地完成，**文档内容不经过任何网络服务**；网络图片由 pandoc 拉取，属用户显式行为。
- 信任边界：Markdown 输入方与 docx/pdf 打开方之间。pandoc 转换会保留 raw HTML/LaTeX（`--reader` 默认含 `raw_html`/`raw_tex`），**不要转换不受信任来源的 Markdown 后直接打开产物**；如需隔离，加 `--reader markdown-raw_html-raw_tex` 禁用原始内容透传。
- Word COM 自动化（docx2pdf.ps1）以只读方式打开中间文档，不执行宏（转换不涉及 .docm）。

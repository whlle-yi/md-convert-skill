# 排版参数速查

组装命令前查此表；默认值即规范值，只叠加用户明确要求的覆盖项。
深度定制（固定 reference.docx、样式机制）见 `styles.md`；转换故障见 `troubleshooting.md`。

## 默认值与覆盖参数

| 项目 | 默认 | 覆盖参数 |
|---|---|---|
| 中文正文 | 宋体 SimSun | `--cjk-font` |
| 西文/数字 | Times New Roman | `--latin-font` |
| 标题/加粗中文 | 思源宋体 Heavy X，**不伪粗** | `--heading-cjk-font` / `--no-heavy` |
| 中文斜体（HTML/PDF） | 楷体，禁用合成斜体 | — |
| 代码 | Consolas + 浅灰底纹，语法高亮默认关 | `--mono-font`；`--highlight` 开启 |
| 正文 | 12pt（小四）、1.5 倍行距、首行缩进 2 字符 | `--font-size` `--line-spacing` `--no-body-indent` |
| 表格 | 三线表；表内文字单倍行距 | `--table-style grid` 换全框线 |
| 目录 | 默认无 | `--toc --toc-depth 3` |
| 页面 | A4，上下 2.54cm、左右 3.17cm | — |
| 图片分辨率 | 96 dpi | `--dpi` |
| 输入格式 | `markdown+east_asian_line_breaks`（中文硬换行不产生多余空格） | `--reader gfm`（严格 GFM 源） |
| 参考文献 | 默认关闭 | `--citeproc` + `--bibliography 文献.bib`（BibTeX / BibLaTeX / CSL JSON / YAML）；`--csl` 指定样式；引用键写 `[@id]` |
| 多格式产出 | 单格式 | `--to docx,pdf,html`（或 `all`）一次扇出，`-o` 为输出目录 |

## 常用组合

```bash
# 标准交付（默认即规范值）
md_convert.py 文档.md --to docx

# 学术论文（编号 + 文献）
md_convert.py 论文.md --to pdf --pdf-engine latex --number-sections --citeproc --bibliography refs.bib

# 无目录的紧凑报告，多格式一次产出
md_convert.py 报告.md --to docx,pdf -o 输出目录/
```

## 注意

- `--font-size` 同步影响标题字号（正文 +4 / +2 / +1 / 0 pt 阶梯）。
- `--dpi` 影响位图显示尺寸（显示宽度 = 像素 ÷ dpi）。
- 字体名须与系统已安装字体精确匹配，缺失时回退（见 `troubleshooting.md`）。

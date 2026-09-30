#!/usr/bin/env python3
"""to_pdf.py — Markdown → PDF，三引擎调度。

- word（Windows + Word 默认）：复用 docx_lib 四步管线产临时 docx，再经 Word COM 导出
- latex：动态生成 XeLaTeX 导言（宋体 + Times New Roman + Heavy 强调）
- libreoffice：无 Word 环境兜底

命令形态与参数见 references/params.md；学术论文场景见 references/scenario-academic.md。
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import common
import docx_lib


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="to_pdf.py",
        description="高质量 Markdown → PDF（md-convert）",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    common.add_core_args(p, common.__version__)
    common.add_style_args(p)
    p.add_argument("--pdf-engine", choices=("auto", "word", "latex", "libreoffice"),
                   default="auto", help="PDF 引擎")
    return p


def build_latex_header(args) -> str:
    """动态生成 XeLaTeX 导言（与 docx 排版规范一致）。"""
    lines: list[str] = []
    lines.append(f"\\setmainfont{{{args.latin_font}}}")
    lines.append("\\usepackage{xeCJK}")
    emph_file = Path(args.emph_font_file) if args.emph_font_file else None
    if args.no_heavy or emph_file is None or not emph_file.is_file():
        lines.append(f"\\setCJKmainfont{{{args.cjk_font}}}[AutoFakeBold=2.5]")
        lines.append(
            f"\\IfFontExistsTF{{{args.heading_cjk_font}}}"
            f"{{\\setCJKfamilyfont{{emphsong}}{{{args.heading_cjk_font}}}}}"
            f"{{\\setCJKfamilyfont{{emphsong}}{{{args.cjk_font}}}[AutoFakeBold=2.5]}}"
        )
    else:
        font_dir = emph_file.parent.as_posix()
        font_name = emph_file.name
        lines.append(
            f"\\setCJKmainfont{{{args.cjk_font}}}"
            f"[Path={font_dir}/, BoldFont={font_name}, AutoFakeBold=2.5]"
        )
        lines.append(
            f"\\setCJKfamilyfont{{emphsong}}{{{font_name}}}[Path={font_dir}/, BoldFont={font_name}]"
        )
    lines.append("\\newcommand{\\key}[1]{{\\CJKfamily{emphsong}\\bfseries #1}}")
    return "\n".join(lines) + "\n"


def _convert_word(src: Path, out: Path, args, tmpdir: Path) -> Path:
    tmp_docx = tmpdir / (src.stem + ".docx")
    docx_lib.convert_docx(src, tmp_docx, args, tmpdir)
    ps1 = Path(__file__).resolve().parent / "docx2pdf.ps1"
    if ps1.is_file():
        cmd = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
               "-File", str(ps1), "-InPath", str(tmp_docx.resolve()),
               "-OutPath", str(out.resolve())]
    else:
        try:
            import docx2pdf  # noqa: F401
        except ImportError:
            raise common.ConvertError("缺少 docx2pdf.ps1 且未安装 docx2pdf，无法用 Word 转 PDF。")
        cmd = [sys.executable, "-c",
               f"import docx2pdf; docx2pdf.convert(r'{tmp_docx}', r'{out}')"]
    if args.verbose:
        common.log("执行: " + " ".join(cmd))
    proc = subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", timeout=300)
    if proc.returncode != 0 or not out.exists():
        detail = (proc.stderr or proc.stdout or "").strip()
        raise common.ConvertError(f"Word COM 转 PDF 失败：{detail}")
    common.log(f"pdf 完成（Word 引擎）：{out.name}")
    return out


def _convert_latex(src: Path, out: Path, args, tmpdir: Path) -> Path:
    if not shutil.which("xelatex"):
        raise common.ConvertError("未找到 xelatex。请安装 TeX Live / MiKTeX，或改用 --pdf-engine word。")
    header = tmpdir / "header.tex"
    header.write_text(build_latex_header(args), encoding="utf-8")
    pandoc_args = [
        str(src), "-f", args.reader,
        "--pdf-engine", "xelatex",
        "-H", str(header),
        "-V", "geometry:top=2.54cm", "-V", "geometry:bottom=2.54cm",
        "-V", "geometry:left=3.17cm", "-V", "geometry:right=3.17cm",
        "-V", f"linestretch={args.line_spacing}",
        "-V", "colorlinks=true", "-V", "linkcolor=black",
        "-V", "urlcolor=black", "-V", "citecolor=black",
        "--resource-path", str(src.parent),
        "--dpi", str(args.dpi),
        "-M", f"lang={args.lang}",
        "-o", str(out),
    ]
    if args.font_size in (10.0, 11.0, 12.0):
        pandoc_args += ["-V", f"fontsize={int(args.font_size)}pt"]
    if not args.highlight:
        pandoc_args.append("--no-highlight")
    pandoc_args += common.citeproc_args(args)
    if args.toc:
        pandoc_args += ["--toc", "--toc-depth", str(args.toc_depth),
                        "-M", "toc-title=目录"]
    if args.number_sections:
        pandoc_args += ["--number-sections"]
    common.run_pandoc(pandoc_args, args.verbose)
    common.log(f"pdf 完成（XeLaTeX 引擎）：{out.name}")
    return out


def _convert_libreoffice(src: Path, out: Path, args, tmpdir: Path) -> Path:
    soffice = shutil.which("soffice")
    if not soffice:
        raise common.ConvertError("未找到 soffice（LibreOffice）。")
    tmp_docx = tmpdir / (src.stem + ".docx")
    docx_lib.convert_docx(src, tmp_docx, args, tmpdir)
    cmd = [soffice, "--headless", "--convert-to", "pdf",
           "--outdir", str(out.parent.resolve()), str(tmp_docx.resolve())]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    produced = out.parent / (tmp_docx.stem + ".pdf")
    if proc.returncode != 0 or not produced.exists():
        raise common.ConvertError(f"LibreOffice 转 PDF 失败：{(proc.stderr or '').strip()}")
    shutil.move(str(produced), str(out))
    common.log(f"pdf 完成（LibreOffice 引擎）：{out.name}")
    return out


def convert_pdf(src: Path, out: Path, args, tmpdir: Path) -> Path:
    engine = args.pdf_engine
    if engine == "auto":
        if sys.platform == "win32":
            engine = "word"
        elif shutil.which("xelatex"):
            engine = "latex"
        elif shutil.which("soffice"):
            engine = "libreoffice"
        else:
            raise common.ConvertError(
                "未找到可用的 PDF 引擎。三选一：\n"
                "  1. Windows + Microsoft Word（推荐，自动使用）\n"
                "  2. XeLaTeX（TeX Live/MiKTeX）\n"
                "  3. LibreOffice（soffice 加入 PATH）\n"
                "或用 --pdf-engine 手动指定。"
            )
    if engine == "word":
        return _convert_word(src, out, args, tmpdir)
    if engine == "latex":
        return _convert_latex(src, out, args, tmpdir)
    if engine == "libreoffice":
        return _convert_libreoffice(src, out, args, tmpdir)
    raise common.ConvertError(f"未知 PDF 引擎：{engine}")


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    return common.standard_main(args, convert_pdf,
                                default_ext="pdf", version=common.__version__)


if __name__ == "__main__":
    sys.exit(main())

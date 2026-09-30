"""md-convert 端到端与样式测试。

- 样式测试直接校验 make_reference_docx 的产物；
- 转换测试运行 md_convert.main()，对实际产物做断言；
- PDF 用例在无可用引擎（无 Word 且无 XeLaTeX）的环境自动跳过。
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn

import md_convert
from make_reference_docx import DEFAULT_HEAVY_CJK_FONT, build_reference

ROOT = Path(__file__).resolve().parents[1]
DEMO = ROOT / "examples" / "demo.md"


# ---------------------------------------------------------------------------
# PDF 引擎探测（模块级计算一次）
# ---------------------------------------------------------------------------

def _word_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        proc = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command",
             "try { $w = New-Object -ComObject Word.Application; $w.Quit(); 'yes' } "
             "catch { 'no' }"],
            capture_output=True, text=True, timeout=90,
        )
        return "yes" in (proc.stdout or "")
    except Exception:
        return False


# 仅探测 Word 引擎：LaTeX 引擎依赖本机中/西文字体（SimSun、Times New Roman），
# CI 的 Linux 镜像两者皆无，自动跳过；本地开发环境可完整覆盖。
PDF_ENGINE = "word" if _word_available() else None


# ---------------------------------------------------------------------------
# reference.docx 样式
# ---------------------------------------------------------------------------

class TestReferenceDocx:
    @pytest.fixture(scope="class")
    @classmethod
    def ref(cls, tmp_path_factory):
        out = tmp_path_factory.mktemp("ref") / "reference.docx"
        return build_reference(out)

    def test_core_styles_exist(self, ref):
        doc = Document(str(ref))
        for name in ("Normal", "Body Text", "First Paragraph", "Compact",
                     "Heading 1", "Heading 2", "Title", "Strong", "Emphasis",
                     "Source Code", "Verbatim Char", "Block Text",
                     "Table Caption", "Image Caption", "Footnote Text"):
            assert name in [s.name for s in doc.styles], f"缺少样式 {name}"

    def test_normal_fonts(self, ref):
        doc = Document(str(ref))
        normal = doc.styles["Normal"]
        assert normal.font.name == "Times New Roman"
        east_asia = normal.element.rPr.rFonts.get(qn("w:eastAsia"))
        assert east_asia == "SimSun"
        assert normal.font.size.pt == 12.0

    def test_body_first_line_indent(self, ref):
        doc = Document(str(ref))
        for name in ("Body Text", "First Paragraph"):
            ind = doc.styles[name].element.pPr.find(qn("w:ind"))
            assert ind is not None and ind.get(qn("w:firstLineChars")) == "200"

    def test_heading_heavy_font(self, ref):
        doc = Document(str(ref))
        h1 = doc.styles["Heading 1"]
        east_asia = h1.element.rPr.rFonts.get(qn("w:eastAsia"))
        assert east_asia == DEFAULT_HEAVY_CJK_FONT

    def test_page_is_a4(self, ref):
        doc = Document(str(ref))
        sec = doc.sections[0]
        assert round(sec.page_width.mm) == 210
        assert round(sec.page_height.mm) == 297

    def test_compact_single_spacing(self, ref):
        """表格/列表文字：单倍行距、无首行缩进（Compact 样式）。"""
        doc = Document(str(ref))
        compact = doc.styles["Compact"]
        assert compact.paragraph_format.line_spacing == 1.0
        ind = compact.element.pPr.find(qn("w:ind"))
        assert ind is not None and ind.get(qn("w:firstLineChars")) == "0"

    def test_no_theme_font_override(self, ref):
        """主题属性优先于显式字体，受控样式必须清除；主题本身也应改为规范字体。"""
        doc = Document(str(ref))
        for name in ("Normal", "Heading 1", "Title", "Strong", "Author", "Date"):
            rfonts = doc.styles[name].element.rPr.rFonts
            for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
                assert rfonts.get(qn(f"w:{attr}")) is None, f"{name} 残留 {attr}"
        theme = [p for p in doc.part.package.iter_parts()
                 if str(p.partname) == "/word/theme/theme1.xml"][0]
        xml = theme.blob.decode("utf-8")
        assert "Times New Roman" in xml and "Aptos" not in xml


# ---------------------------------------------------------------------------
# docx 转换
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def docx_out(tmp_path_factory):
    out = tmp_path_factory.mktemp("docx") / "demo.docx"
    code = md_convert.main([str(DEMO), "-o", str(out), "--toc", "--verbose"])
    assert code == 0
    return out


class TestConvertDocx:
    def test_output_valid(self, docx_out):
        assert docx_out.exists() and docx_out.stat().st_size > 10_000
        doc = Document(str(docx_out))  # 能正常解析

    def test_tables_threeline(self, docx_out):
        doc = Document(str(docx_out))
        assert len(doc.tables) == 2
        # 单元格段落必须用 Compact 样式（单倍行距），不得回退到 1.5 倍的 Body Text
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        assert p.style.name == "Compact", \
                            f"单元格段落样式异常：{p.style.name}"
        borders = doc.tables[0]._tbl.tblPr.find(qn("w:tblBorders"))
        assert borders is not None
        assert borders.find(qn("w:top")).get(qn("w:val")) == "single"
        assert borders.find(qn("w:top")).get(qn("w:sz")) == "12"
        assert borders.find(qn("w:insideV")).get(qn("w:val")) == "none"
        # 表头下 0.75pt 细线
        tc = doc.tables[0].rows[0].cells[0]._tc
        bottom = tc.tcPr.find(qn("w:tcBorders")).find(qn("w:bottom"))
        assert bottom.get(qn("w:sz")) == "6"

    def test_header_row_font_weight(self, docx_out):
        """表头：西文 run 加粗；中文 run 为 Heavy 字体（显式关闭伪粗体）。"""
        doc = Document(str(docx_out))
        runs = [r for c in doc.tables[0].rows[0].cells for p in c.paragraphs for r in p.runs]
        assert runs
        for r in runs:
            rpr = r._r.rPr
            east_asia = rpr.rFonts.get(qn("w:eastAsia")) if rpr is not None and rpr.rFonts is not None else None
            ok_bold = r.font.bold is True
            ok_heavy = east_asia == DEFAULT_HEAVY_CJK_FONT and (
                rpr.find(qn("w:b")) is not None and rpr.find(qn("w:b")).get(qn("w:val")) in ("0", "false")
            )
            assert ok_bold or ok_heavy, f"表头 run 未规范化：text={r.text!r}"

    def test_cjk_bold_no_fake_bold(self, docx_out):
        """中文加粗 run 应为 Heavy 字体且显式 w:b=0。"""
        doc = Document(str(docx_out))
        body = doc.element.body
        hits = 0
        for rpr in body.iter(qn("w:rPr")):
            rfonts = rpr.find(qn("w:rFonts"))
            b = rpr.find(qn("w:b"))
            if (rfonts is not None
                    and rfonts.get(qn("w:eastAsia")) == DEFAULT_HEAVY_CJK_FONT
                    and b is not None and b.get(qn("w:val")) in ("0", "false")):
                hits += 1
        assert hits > 0, "未找到 Heavy + w:b=0 的中文强调 run"

    def test_update_fields_when_toc(self, docx_out):
        doc = Document(str(docx_out))
        uf = doc.settings.element.find(qn("w:updateFields"))
        assert uf is not None
        assert uf.get(qn("w:val")) == "true"

    def test_document_contains_key_content(self, docx_out):
        doc = Document(str(docx_out))
        text = "\n".join(p.text for p in doc.paragraphs)
        assert "转换质量演示" in text
        assert "麦克斯韦" in text  # 公式小节标题


# ---------------------------------------------------------------------------
# html 转换
# ---------------------------------------------------------------------------

class TestConvertHtml:
    def test_html_self_contained(self, tmp_path):
        out = tmp_path / "demo.html"
        assert md_convert.main([str(DEMO), "-o", str(out), "--toc"]) == 0
        html = out.read_text(encoding="utf-8")
        assert "<table" in html
        assert "font-family" in html          # CSS 已内嵌
        assert "http-equiv" in html or "<title>" in html
        assert "<math" in html or "mathml" in html.lower()  # MathML 公式

    def test_title_extracted_from_yaml(self, tmp_path):
        out = tmp_path / "demo.html"
        md_convert.main([str(DEMO), "-o", str(out)])
        assert "<title>md-convert 转换质量演示</title>" in out.read_text(encoding="utf-8")


# ---------------------------------------------------------------------------
# pdf 转换（引擎可用时）
# ---------------------------------------------------------------------------

@pytest.mark.skipif(PDF_ENGINE is None, reason="无可用 PDF 引擎（需要 Word 或 XeLaTeX）")
class TestConvertPdf:
    def test_pdf_generated(self, tmp_path):
        out = tmp_path / "demo.pdf"
        code = md_convert.main([str(DEMO), "-o", str(out),
                                "--pdf-engine", PDF_ENGINE, "--toc"])
        assert code == 0
        data = out.read_bytes()
        assert data[:5] == b"%PDF-"
        assert len(data) > 20_000


# ---------------------------------------------------------------------------
# CLI 行为
# ---------------------------------------------------------------------------

class TestCli:
    def test_missing_input(self, capsys):
        assert md_convert.main(["no-such-file.md"]) == 2

    def test_unknown_format(self, tmp_path, capsys):
        md = tmp_path / "a.md"
        md.write_text("# hi", encoding="utf-8")
        assert md_convert.main([str(md), "-o", str(tmp_path / "a.xyz")]) == 2

    def test_format_from_output_suffix(self, tmp_path):
        md = tmp_path / "b.md"
        md.write_text("# 标题\n\n正文**加粗**测试。", encoding="utf-8")
        assert md_convert.main([str(md), "-o", str(tmp_path / "b.html")]) == 0
        assert (tmp_path / "b.html").exists()

    def test_multi_output(self, tmp_path):
        md = tmp_path / "c.md"
        md.write_text("# 多产出\n\n正文**加粗**测试。", encoding="utf-8")
        outdir = tmp_path / "out"
        assert md_convert.main([str(md), "--to", "docx,html", "-o", str(outdir)]) == 0
        assert (outdir / "c.docx").exists()
        assert (outdir / "c.html").exists()
        assert not (outdir / "c.pdf").exists()  # 未要求的格式不产出

    def test_to_unknown_format(self, tmp_path):
        md = tmp_path / "d.md"
        md.write_text("# x", encoding="utf-8")
        assert md_convert.main([str(md), "--to", "docx,xyz"]) == 2

    def test_to_conflicts_with_format(self, tmp_path):
        md = tmp_path / "e.md"
        md.write_text("# x", encoding="utf-8")
        assert md_convert.main([str(md), "--to", "docx", "-f", "pdf"]) == 2

    def test_citeproc_metadata(self, tmp_path):
        md = tmp_path / "cite.md"
        md.write_text(
            "---\ntitle: 引用测试\n"
            "references:\n- id: n2008\n"
            "  title: Bitcoin: A Peer-to-Peer Electronic Cash System\n"
            "  author:\n  - family: Nakamoto\n    given: Satoshi\n"
            "  issued:\n    year: 2008\n"
            "---\n\n去中心化电子现金系统 [@n2008]。\n",
            encoding="utf-8",
        )
        out = tmp_path / "cite.docx"
        assert md_convert.main([str(md), "-o", str(out), "--citeproc"]) == 0
        text = "\n".join(p.text for p in Document(str(out)).paragraphs)
        assert "Nakamoto" in text  # 引文与文末参考文献均已渲染

    def test_citeproc_bibliography_file(self, tmp_path):
        md = tmp_path / "cite2.md"
        md.write_text("去中心化电子现金系统 [@n2008]。\n", encoding="utf-8")
        bib = tmp_path / "refs.yaml"
        bib.write_text(
            "references:\n- id: n2008\n  title: Bitcoin\n"
            "  author:\n  - family: Nakamoto\n    given: Satoshi\n"
            "  issued:\n    year: 2008\n",
            encoding="utf-8",
        )
        out = tmp_path / "cite2.docx"
        assert md_convert.main([str(md), "-o", str(out), "--citeproc",
                                "--bibliography", str(bib)]) == 0
        text = "\n".join(p.text for p in Document(str(out)).paragraphs)
        assert "Nakamoto" in text

"""按《党政机关公文格式》(GB/T 9704-2012) 生成 Word 文档。"""

from __future__ import annotations

import io
from collections.abc import Sequence

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BODY = Pt(16)          # 三号
TITLE_SIZE = Pt(22)    # 二号
LINE = Pt(28)          # 正文固定行距 28 磅，一页 22 行
INDENT = Pt(32)        # 首行缩进 2 字符（三号 16pt × 2）
RED = RGBColor(0xC0, 0x00, 0x00)


def _font(
    run, east_asia: str, size, bold: bool = False, color=None, ascii_font: str = "Times New Roman"
) -> None:
    """同时设置中西文字体：中文用公文字体，数字与字母默认用 Times New Roman。"""
    run.font.size = size
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color

    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:eastAsia"), east_asia)
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)


def _para(
    doc,
    text: str,
    east_asia: str,
    size,
    *,
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    indent=0,
    right=0,
    line=LINE,
    after=0,
    bold=False,
    color=None,
):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = align
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = line
    pf.space_before = Pt(0)
    pf.space_after = Pt(after)
    if indent:
        pf.first_line_indent = Pt(indent)
    if right:
        pf.right_indent = Pt(right)
    _font(p.add_run(text), east_asia, size, bold, color)
    return p


def _page_number(section) -> None:
    """页脚居中页码，四号宋体，数字两侧加一字线。"""
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    _font(p.add_run("— "), "宋体", Pt(14), ascii_font="宋体")

    holder = p.add_run()
    _font(holder, "宋体", Pt(14), ascii_font="宋体")
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    holder._r.append(begin)
    holder._r.append(instr)
    holder._r.append(end)

    _font(p.add_run(" —"), "宋体", Pt(14), ascii_font="宋体")


def build_docx(
    blocks: Sequence[dict], body_font: str = "仿宋", title_font: str = "黑体"
) -> bytes:
    doc = Document()

    normal = doc.styles["Normal"]
    normal.font.name = body_font
    normal.font.size = BODY
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), body_font)

    # A4，上 3.7cm / 下 3.5cm / 左 2.8cm / 右 2.6cm
    sec = doc.sections[0]
    sec.page_width = Cm(21)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(3.7)
    sec.bottom_margin = Cm(3.5)
    sec.left_margin = Cm(2.8)
    sec.right_margin = Cm(2.6)
    _page_number(sec)

    for block in blocks:
        role = block.get("role", "body")
        text = block.get("text", "")

        if role == "pagebreak":
            # 分页符挂在上一段末尾，避免多出一个空段落占位
            if doc.paragraphs:
                doc.paragraphs[-1].add_run().add_break(WD_BREAK.PAGE)
            else:
                doc.add_page_break()
        elif role == "docorg":
            _para(doc, text, title_font, Pt(36),
                  align=WD_ALIGN_PARAGRAPH.CENTER, line=Pt(52), after=28, color=RED)
        elif role == "title":
            _para(doc, text, title_font, TITLE_SIZE,
                  align=WD_ALIGN_PARAGRAPH.CENTER, line=Pt(32), after=28)
        elif role == "recipient":
            _para(doc, text, body_font, BODY, align=WD_ALIGN_PARAGRAPH.LEFT)
        elif role == "heading1":
            _para(doc, text, "黑体", BODY, indent=INDENT)
        elif role == "heading2":
            _para(doc, text, "楷体", BODY, indent=INDENT)
        elif role == "heading3":
            _para(doc, text, body_font, BODY, indent=INDENT, bold=True)
        elif role == "signature":
            _para(doc, text, body_font, BODY,
                  align=WD_ALIGN_PARAGRAPH.RIGHT, right=32)
        elif role == "date":
            _para(doc, text, body_font, BODY,
                  align=WD_ALIGN_PARAGRAPH.RIGHT, right=64)
        elif role == "blank":
            _para(doc, "", body_font, BODY)
        else:
            _para(doc, text, body_font, BODY, indent=INDENT)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

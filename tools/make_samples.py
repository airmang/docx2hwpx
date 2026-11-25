# -*- coding: utf-8 -*-
"""Five synthetic Word documents for the DOCX->HWPX bench. All text is made up.

    venv-py python -X utf8 make_samples.py <out_dir>
"""
import copy
import io
import re
import sys
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor
from PIL import Image, ImageDraw

OUT = Path(sys.argv[1])
OUT.mkdir(parents=True, exist_ok=True)

FILLER = ("가상의 안내 문장입니다. 이 문장은 줄바꿈과 쪽 나눔을 시험하려고 일부러 길게 썼으며, "
          "실제 기관이나 사람과는 아무 관련이 없습니다. 같은 내용을 몇 번 되풀이해 문단 길이를 늘립니다. ")


def set_east_asia(run, name):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    fonts.set(qn("w:eastAsia"), name)
    fonts.set(qn("w:ascii"), name)
    fonts.set(qn("w:hAnsi"), name)


def base_doc():
    d = Document()
    st = d.styles["Normal"]
    st.font.size = Pt(11)
    rpr = st.element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    for k in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        fonts.set(qn(k), "맑은 고딕")
    s = d.sections[0]
    s.page_width, s.page_height = Mm(210), Mm(297)
    s.top_margin, s.bottom_margin = Mm(20), Mm(20)
    s.left_margin, s.right_margin = Mm(25), Mm(25)
    return d


def shade(cell, hex_fill):
    tcpr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcpr.append(shd)


def page_field(paragraph, prefix="- ", suffix=" -"):
    paragraph.add_run(prefix)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    fld.append(r)
    paragraph._p.append(fld)
    paragraph.add_run(suffix)


def hyperlink(paragraph, url, text):
    part = paragraph.part
    rid = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                         is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), rid)
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    c = OxmlElement("w:color")
    c.set(qn("w:val"), "0563C1")
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rpr.append(c)
    rpr.append(u)
    r.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    h.append(r)
    paragraph._p.append(h)


def footnote_ref(paragraph, fid):
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    va = OxmlElement("w:vertAlign")
    va.set(qn("w:val"), "superscript")
    rpr.append(va)
    r.append(rpr)
    ref = OxmlElement("w:footnoteReference")
    ref.set(qn("w:id"), str(fid))
    r.append(ref)
    paragraph._p.append(r)


def picture_png(path, w=600, h=300, label="그림"):
    im = Image.new("RGB", (w, h), (235, 242, 250))
    dr = ImageDraw.Draw(im)
    dr.rectangle([10, 10, w - 10, h - 10], outline=(40, 90, 160), width=6)
    for i in range(5):
        x = 60 + i * 100
        hh = 40 + i * 35
        dr.rectangle([x, h - 40 - hh, x + 60, h - 40], fill=(60 + i * 30, 120, 200 - i * 20))
    im.save(path)


FOOTNOTES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
{notes}
</w:footnotes>"""
NOTE = ('<w:footnote w:id="{i}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>'
        '<w:r><w:t xml:space="preserve"> {text}</w:t></w:r></w:p></w:footnote>')


def add_footnotes_part(docx_path, texts):
    """python-docx cannot write footnotes: add word/footnotes.xml, its relationship and content type."""
    src = docx_path.read_bytes()
    zin = zipfile.ZipFile(io.BytesIO(src))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                s = data.decode("utf-8")
                s = s.replace("</Types>", '<Override PartName="/word/footnotes.xml" ContentType="application/'
                              'vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>')
                data = s.encode("utf-8")
            elif item.filename == "word/_rels/document.xml.rels":
                s = data.decode("utf-8")
                s = s.replace("</Relationships>", '<Relationship Id="rIdFn1" Type="http://schemas.openxmlformats.org/'
                              'officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>')
                data = s.encode("utf-8")
            zout.writestr(copy.copy(item), data)
        notes = "\n".join(NOTE.format(i=i + 1, text=t) for i, t in enumerate(texts))
        zout.writestr("word/footnotes.xml", FOOTNOTES_XML.format(notes=notes))
    docx_path.write_bytes(buf.getvalue())


# ---------------------------------------------------------------- s1 text formatting
def s1():
    d = base_doc()
    d.add_heading("1. 글자와 문단 서식 시험", level=1)
    d.add_heading("1.1 글자 모양", level=2)
    p = d.add_paragraph("보통 글자, ")
    p.add_run("굵게").bold = True
    p.add_run(", ")
    p.add_run("기울임").italic = True
    p.add_run(", ")
    p.add_run("밑줄").underline = True
    p.add_run(", ")
    r = p.add_run("취소선")
    r.font.strike = True
    p.add_run(", ")
    r = p.add_run("빨간 글자")
    r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    p.add_run(", ")
    r = p.add_run("16pt 글자")
    r.font.size = Pt(16)
    p.add_run(", ")
    r = p.add_run("형광펜")
    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    p.add_run(", 위첨자 x")
    r = p.add_run("2")
    r.font.superscript = True
    p.add_run(", 아래첨자 H")
    r = p.add_run("2")
    r.font.subscript = True
    p.add_run("O.")
    p = d.add_paragraph()
    r = p.add_run("바탕 글꼴로 쓴 문장입니다. ")
    set_east_asia(r, "바탕")
    r = p.add_run("돋움 글꼴로 쓴 문장입니다.")
    set_east_asia(r, "돋움")
    d.add_heading("1.2 문단 정렬과 들여쓰기", level=2)
    for text, al in (("왼쪽 정렬 문단입니다.", WD_ALIGN_PARAGRAPH.LEFT), ("가운데 정렬 문단입니다.", WD_ALIGN_PARAGRAPH.CENTER),
                     ("오른쪽 정렬 문단입니다.", WD_ALIGN_PARAGRAPH.RIGHT), ("양쪽 정렬 문단입니다. " + FILLER * 2, WD_ALIGN_PARAGRAPH.JUSTIFY)):
        d.add_paragraph(text).alignment = al
    p = d.add_paragraph("첫 줄 들여쓰기 10mm 문단입니다. " + FILLER * 2)
    p.paragraph_format.first_line_indent = Mm(10)
    p = d.add_paragraph("내어쓰기 문단입니다. 첫 줄이 왼쪽으로 나오고 다음 줄들이 들어갑니다. " + FILLER * 2)
    p.paragraph_format.left_indent = Mm(15)
    p.paragraph_format.first_line_indent = Mm(-15)
    p = d.add_paragraph("왼쪽 20mm, 오른쪽 20mm 들여 쓴 문단입니다. " + FILLER)
    p.paragraph_format.left_indent = Mm(20)
    p.paragraph_format.right_indent = Mm(20)
    d.add_heading("1.3 줄 간격과 문단 간격", level=2)
    p = d.add_paragraph("줄 간격 1.5줄 문단입니다. " + FILLER * 2)
    p.paragraph_format.line_spacing = 1.5
    p = d.add_paragraph("줄 간격 고정 24pt 문단입니다. " + FILLER * 2)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(24)
    p = d.add_paragraph("문단 앞 18pt, 뒤 12pt 간격 문단입니다.")
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(12)
    p = d.add_paragraph("이 문단 뒤에서 쪽을 나눕니다.")
    p.add_run().add_break(WD_BREAK.PAGE)
    d.add_paragraph("둘째 쪽 첫 문단입니다. " + FILLER)
    d.save(OUT / "s1_text.docx")


# ---------------------------------------------------------------- s2 lists
def s2():
    d = base_doc()
    d.add_heading("2. 목록 시험", level=1)
    d.add_paragraph("글머리표 목록:")
    for t, lv in (("첫째 항목", 1), ("둘째 항목", 1), ("둘째 항목의 하위 항목 가", 2), ("둘째 항목의 하위 항목 나", 2), ("셋째 항목 " + FILLER, 1)):
        d.add_paragraph(t, style="List Bullet" if lv == 1 else "List Bullet 2")
    d.add_paragraph("번호 목록:")
    for t, lv in (("준비물을 챙깁니다.", 1), ("신청서를 냅니다.", 1), ("서류 하나", 2), ("서류 둘", 2), ("결과를 확인합니다. " + FILLER, 1)):
        d.add_paragraph(t, style="List Number" if lv == 1 else "List Number 2")
    d.add_paragraph("목록 뒤의 보통 문단입니다.")
    d.save(OUT / "s2_list.docx")


# ---------------------------------------------------------------- s3 tables
def s3():
    d = base_doc()
    d.add_heading("3. 표 시험", level=1)
    d.add_paragraph("표 1. 머리 줄 음영, 가로·세로 병합, 열 너비가 다른 표")
    t = d.add_table(rows=5, cols=4)
    t.style = "Table Grid"
    widths = (Mm(30), Mm(60), Mm(35), Mm(35))
    heads = ("구분", "내용", "수량", "금액(원)")
    for i, h in enumerate(heads):
        c = t.cell(0, i)
        c.text = h
        c.paragraphs[0].runs[0].bold = True
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        shade(c, "D9E2F3")
    rows = (("가", "연필", "12", "6,000"), ("가", "공책", "5", "7,500"), ("나", "물감", "3", "21,000"))
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row):
            cell = t.cell(r, c)
            cell.text = v
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if c >= 2:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    t.cell(1, 0).merge(t.cell(2, 0)).text = "가"
    m = t.cell(4, 0).merge(t.cell(4, 2))
    m.text = "합계"
    m.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    t.cell(4, 3).text = "34,500"
    t.cell(4, 3).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = w
    d.add_paragraph("")
    d.add_paragraph("표 2. 칸 안에 표가 든 표")
    t2 = d.add_table(rows=2, cols=2)
    t2.style = "Table Grid"
    t2.cell(0, 0).text = "바깥 표 칸 1"
    t2.cell(0, 1).text = "바깥 표 칸 2"
    t2.cell(1, 0).text = "아래 칸에는 작은 표가 있습니다."
    inner = t2.cell(1, 1).add_table(rows=2, cols=2)
    inner.style = "Table Grid"
    for i, v in enumerate(("안1", "안2", "안3", "안4")):
        inner.cell(i // 2, i % 2).text = v
    d.add_paragraph("")
    d.add_paragraph("표 3. 테두리 없는 표")
    t3 = d.add_table(rows=2, cols=3)
    for i, v in enumerate(("이름", "학년", "반", "홍길동", "3", "2")):
        t3.cell(i // 3, i % 3).text = v
    d.add_paragraph("표 뒤의 보통 문단입니다.")
    d.save(OUT / "s3_table.docx")


# ---------------------------------------------------------------- s4 page features
def s4(png):
    d = base_doc()
    sec = d.sections[0]
    sec.header.paragraphs[0].text = "머리말: 가상 기관 안내 문서"
    sec.header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    page_field(fp)
    d.add_heading("4. 쪽 기능 시험", level=1)
    p = d.add_paragraph("각주가 달린 문장입니다")
    footnote_ref(p, 1)
    p.add_run(". 두 번째 각주도 있습니다")
    footnote_ref(p, 2)
    p.add_run(".")
    p = d.add_paragraph("링크: ")
    hyperlink(p, "https://example.com/", "예시 누리집")
    p.add_run(" 를 누르면 열립니다.")
    d.add_paragraph("아래는 본문에 넣은 그림입니다.")
    d.add_picture(str(png), width=Cm(8))
    d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    d.add_paragraph("그림 1. 막대 그림").alignment = WD_ALIGN_PARAGRAPH.CENTER
    for _ in range(2):
        d.add_paragraph(FILLER * 3)
    # section 2: landscape
    s2 = d.add_section(WD_SECTION.NEW_PAGE)
    s2.orientation = WD_ORIENT.LANDSCAPE
    s2.page_width, s2.page_height = Mm(297), Mm(210)
    d.add_heading("가로 방향 구역", level=2)
    d.add_paragraph("이 구역은 가로 방향입니다. " + FILLER * 2)
    # section 3: portrait, two columns
    s3_ = d.add_section(WD_SECTION.NEW_PAGE)
    s3_.orientation = WD_ORIENT.PORTRAIT
    s3_.page_width, s3_.page_height = Mm(210), Mm(297)
    cols = s3_._sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        s3_._sectPr.append(cols)
    cols.set(qn("w:num"), "2")
    cols.set(qn("w:space"), "425")
    d.add_heading("두 단 구역", level=2)
    for _ in range(6):
        d.add_paragraph(FILLER * 2)
    path = OUT / "s4_page.docx"
    d.save(path)
    add_footnotes_part(path, ["첫째 각주의 설명 글입니다.", "둘째 각주의 설명 글입니다."])


# ---------------------------------------------------------------- s5 mixed notice (multi-page)
def s5(png):
    d = base_doc()
    sec = d.sections[0]
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    page_field(fp, "", "")
    t = d.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("가상 학교 여름 체험 활동 안내")
    r.bold = True
    r.font.size = Pt(20)
    d.add_paragraph("작성: 가상 담당 부서").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    d.add_heading("1. 목적", level=2)
    d.add_paragraph(FILLER * 3)
    d.add_heading("2. 일정", level=2)
    tb = d.add_table(rows=4, cols=3)
    tb.style = "Table Grid"
    for i, v in enumerate(("날짜", "활동", "장소", "7월 1일", "숲 체험", "가상 공원", "7월 2일", "과학 실험", "실험실",
                           "7월 3일", "발표회", "강당")):
        c = tb.cell(i // 3, i % 3)
        c.text = v
        if i < 3:
            c.paragraphs[0].runs[0].bold = True
            shade(c, "E2EFDA")
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    d.add_heading("3. 준비물", level=2)
    for v in ("물병", "모자", "필기도구", "갈아입을 옷"):
        d.add_paragraph(v, style="List Bullet")
    d.add_heading("4. 신청 방법", level=2)
    for v in ("신청서를 작성합니다.", "담임 선생님께 냅니다.", "확인 문자를 받습니다."):
        d.add_paragraph(v, style="List Number")
    d.add_picture(str(png), width=Cm(10))
    d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    d.add_heading("5. 유의 사항", level=2)
    for _ in range(8):
        p = d.add_paragraph(FILLER * 2)
        p.paragraph_format.first_line_indent = Mm(5)
    p = d.add_paragraph("※ 문의는 가상 담당 부서로 해 주십시오.")
    p.runs[0].font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    d.save(OUT / "s5_mixed.docx")


png = OUT / "_bar.png"
picture_png(png)
s1()
s2()
s3()
s4(png)
s5(png)
png.unlink()
for f in sorted(OUT.glob("*.docx")):
    print(f.name, f.stat().st_size)

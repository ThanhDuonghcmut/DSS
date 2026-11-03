# -*- coding: utf-8 -*-
"""Chuyển PhanB_BaoCao.md thành PhanB_BaoCao.docx để dán vào báo cáo chung của nhóm.

Chạy (từ thư mục PhanB, trong venv dựng từ requirements.txt):
    python3 build_docx.py
Hỗ trợ đúng tập Markdown mà báo cáo dùng: tiêu đề #/##/###, đoạn văn, **đậm**, *nghiêng*,
`mã`, gạch đầu dòng (hai cấp), danh sách đánh số, bảng, ảnh. Kiểu chữ theo PhanA/build_docx.py
(Calibri 11, tiêu đề xanh navy, đầu bảng tô nền).
"""
import os
import re

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "PhanB_BaoCao.md")
OUT = os.path.join(HERE, "PhanB_BaoCao.docx")
NAVY = RGBColor(0x1F, 0x4E, 0x79)
MAX_W, MAX_H = 6.3, 6.0  # inch: vừa khổ A4/Letter lề 1", ảnh dọc không tràn trang

INLINE = re.compile(r"(\*\*.+?\*\*|\*.+?\*|`.+?`)")

doc = Document()
doc.styles["Normal"].font.name = "Calibri"
doc.styles["Normal"].font.size = Pt(11)
# tiếng Việt cần khai báo font cho cả vùng eastAsia để Word không đổi font dấu
doc.styles["Normal"].element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")


def add_runs(p, text, size=None):
    """Thêm các đoạn chữ vào đoạn văn p, xử lý **đậm**, *nghiêng*, `mã`."""
    for part in INLINE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            r = p.add_run(part[2:-2])
            r.bold = True
        elif part.startswith("`") and part.endswith("`") and len(part) > 2:
            r = p.add_run(part[1:-1])
            r.font.name = "Consolas"
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            r = p.add_run(part[1:-1])
            r.italic = True
        else:
            r = p.add_run(part)
        if size:
            r.font.size = Pt(size)


def heading(text, level):
    p = doc.add_heading(level=level)
    add_runs(p, text)
    for r in p.runs:
        r.font.color.rgb = NAVY
        r.font.name = "Calibri"
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    return p


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def table(rows):
    ncol = len(rows[0])
    t = doc.add_table(rows=len(rows), cols=ncol)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j in range(ncol):
            cell = t.cell(i, j)
            cell.text = ""
            p = cell.paragraphs[0]
            add_runs(p, row[j] if j < len(row) else "", size=10)
            if i == 0:
                shade(cell, "DCE6F1")
                for r in p.runs:
                    r.bold = True
    doc.add_paragraph()


def picture(path):
    full = os.path.join(HERE, path)
    with Image.open(full) as im:
        w, h = im.size
    width = MAX_W
    if h / w * width > MAX_H:  # ảnh dọc (hộp thoại): giới hạn theo chiều cao
        width = MAX_H * w / h
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(full, width=Inches(width))
    p.paragraph_format.keep_with_next = True  # giữ ảnh cùng trang với chú thích


def main():
    lines = open(SRC, encoding="utf-8").read().split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]
        if not line.strip():
            i += 1
            continue
        m = re.match(r"^(#{1,3}) (.*)$", line)
        if m:
            heading(m.group(2), len(m.group(1)))
        elif line.startswith("|"):
            block = []
            while i < len(lines) and lines[i].startswith("|"):
                block.append(lines[i])
                i += 1
            rows = [[c.strip() for c in b.strip().strip("|").split("|")]
                    for b in block if not re.match(r"^\|[\s:|-]+\|$", b)]
            table(rows)
            continue
        elif re.match(r"^!\[.*\]\((.*)\)$", line):
            picture(re.match(r"^!\[.*\]\((.*)\)$", line).group(1))
        elif re.match(r"^\*Hình \d+:.*\*$", line):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_runs(p, line, size=9)
        elif re.match(r"^ {3}- ", line):
            p = doc.add_paragraph(style="List Bullet 2")
            add_runs(p, line.strip()[2:])
        elif line.startswith("- "):
            p = doc.add_paragraph(style="List Bullet")
            add_runs(p, line[2:])
        elif re.match(r"^\d+\. ", line):
            # đánh số tay để mỗi danh sách bắt đầu lại từ 1 (List Number của Word đánh liên tục)
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.3)
            p.paragraph_format.first_line_indent = Inches(-0.25)
            add_runs(p, line)
        else:
            p = doc.add_paragraph()
            add_runs(p, line)
        i += 1
    doc.save(OUT)
    print("Đã ghi", OUT)


if __name__ == "__main__":
    main()

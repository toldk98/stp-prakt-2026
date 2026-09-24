#!/usr/bin/env python3
"""Конвертер md → docx для роздаткової бази ПР1.

Читає «ПР1. Аналіз та декомпозиція задачі.md» і будує Word-документ:
заголовки H1/H2, звичайні абзаци, марковані списки, таблиці.

Запуск:  python3 md_to_docx_pr1.py
"""

import re
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

SRC = "/home/toldk98/Документи/ХНТУ/ОК/СТП/Практичні роботи/ПР1. Аналіз та декомпозиція задачі.md"
OUT = "/home/toldk98/Документи/ХНТУ/ОК/СТП/Практичні роботи/ПР1. Аналіз та декомпозиція задачі.docx"

DARK = RGBColor(0x0F, 0x1E, 0x2E)
TEAL = RGBColor(0x14, 0xB8, 0xA6)
ORANGE = RGBColor(0xE0, 0x74, 0x22)
GRAY = RGBColor(0x46, 0x54, 0x60)

TEAM = "СТП • ПР1 • 2026"


def para(doc, text, bold=False, size=11, color=None, align=None,
         space_after=6, italic=False):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    r.font.name = "Calibri"
    if color:
        r.font.color.rgb = color
    return p


def table(doc, rows):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Light Grid Accent 1"
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = t.cell(i, j)
            cell.text = ""
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(10.5)
            r.font.name = "Calibri"
            if i == 0:
                r.bold = True
                r.font.color.rgb = DARK


def make():
    doc = Document()
    # базовий стиль
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(11)
    st.font.color.rgb = GRAY

    lines = []
    m = 0
    for raw in Path(SRC).read_text(encoding="utf-8").splitlines():
        lines.append(raw)

    # титул
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ПР1. Аналіз та декомпозиція задачі")
    r.bold = True
    r.font.size = Pt(20)
    r.font.color.rgb = DARK
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Коротка роздатка для першого практичного заняття")
    r.font.size = Pt(12)
    r.font.color.rgb = TEAL
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(TEAM)
    r.font.size = Pt(10)
    r.font.color.rgb = GRAY
    doc.add_paragraph()

    i = 0
    in_table = False
    buf = []
    rows = []
    while i < len(lines):
        raw = lines[i]
        line = raw.rstrip()
        i += 1

        # таблиця: рядок типу "| a | b |"
        if line.startswith("|"):
            c = [x.strip() for x in line.strip().strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", x) for x in c):
                continue
            rows.append(c)
            in_table = True
            continue
        if in_table and rows:
            table(doc, rows)
            rows = []
            in_table = False
            doc.add_paragraph()

        # заголовки
        if line.startswith("### "):
            para(doc, line[4:], bold=True, size=13, color=DARK, space_after=4)
            continue
        if line.startswith("## "):
            para(doc, line[3:], bold=True, size=15, color=TEAL, space_after=2)
            continue
        if line.startswith("# "):
            para(doc, line[2:], bold=True, size=17, color=DARK, space_after=6)
            continue

        # список
        if re.match(r"^\s*[-*•] ", line):
            para(doc, line, size=11, space_after=2)
            continue
        if re.match(r"^\s*\d+\.\s", line):
            para(doc, line, size=11, space_after=2)
            continue

        # блок коду — як моноширинний абзац
        if line.strip().startswith("```") or line.strip().startswith("~~~"):
            continue

        # порожній рядок
        if not line.strip():
            continue

        # звичайний абзац
        para(doc, line, size=11, space_after=PanhandleNone if False else 4)

    if rows:
        table(doc, rows)

    doc.save(OUT)
    print("docx:", OUT)


if __name__ == "__main__":
    make()

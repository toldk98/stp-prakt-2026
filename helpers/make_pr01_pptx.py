#!/usr/bin/env python3
"""ПР1 → коротка презентація (pptx), дані з бази .md."""
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

BASE = Path("/home/toldk98/Документи/ХНТУ/ОК/СТП/Практичні роботи/ПР1. Аналіз та декомпозиція задачі.md")
OUT = Path("/home/toldk98/Документи/ХНТУ/ОК/СТП/Практичні роботи/ПР1. Аналіз та декомпозиція задачі.pptx")

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x0F, 0x1E, 0x2E)
TEAL = RGBColor(0x14, 0xB8, 0xA6)
ORANGE = RGBColor(0xE0, 0x74, 0x22)
GRAY = RGBColor(0x46, 0x54, 0x60)

BLANK = 6
prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)


def add_slide():
    return prs.slides.add_slide(prs.slide_layouts[BLANK])


def add_textbox(s, x, y, w, h):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    return tf


def para(tf, text, size, color, bold=False, align=PP_ALIGN.LEFT, first=False, space=6):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space)
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.bold = bold
    r.font.name = "Calibri"


bg_hex = "0F1E2E"
def bg(s, hexcolor):
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = RGBColor.from_string(hexcolor)


def header(s, title, subtitle=None):
    bg(s, bg_hex)
    tf = add_textbox(s, 0.9, 0.9, 11.5, 1.2)
    para(tf, title, 34, WHITE, bold=True, first=True)
    if subtitle:
        tf2 = add_textbox(s, 0.9, 2.1, 11.5, 0.8)
        para(tf2, subtitle, 16, RGBColor(0x9C, 0xB3, 0xC5), first=True, align=PP_ALIGN.RIGHT)


def body(s, lines, size=20, color=DARK, start_x=0.9, start_y=2.6, gap=18):
    tf = add_textbox(s, start_x, start_y, 11.5, 4.5)
    for i, (txt, fmt) in enumerate(lines):
        mark = fmt if fmt else None
        para(tf, txt, size, color, bold=(mark == "b"), space=gap, first=(i == 0))


s = add_slide()
bg(s, bg_hex)
tf = add_textbox(s, 0.9, 2.2, 11.5, 2.4)
para(tf, "ПР1", 60, TEAL, bold=True, first=True)
para(tf, "Аналіз та декомпозиція задачі", 40, WHITE, bold=True)
para(tf, "Практична робота 1 • 2 год • 2,5 бала", 16, RGBColor(0x9C, 0xB3, 0xC5))

s = add_slide()
header(s, "Крок 1", "Що робимо")
body(s, [
    ("Відкрий свій варіант: tasks/task_vNN.md", "b"),
    ("Випиши доменні сутності (Прилад, Профіль…) та їх поля.", None),
    ("Сформулюй постановку: вхід, результат, обмеження, критерій готовності.", None),
    ("Код не пишемо.", "b"),
])

s = add_slide()
header(s, "Крок 2", "Декомпозиція")
body(s, [
    ("Розбий задачу на функції/модулі.", "b"),
    ("Для кожної — сигнатура: ім'я, аргументи, що повертає, винятки.", None),
    ("Декомпозиція по відповідальностях, а не по файлах.", None),
])

s = add_slide()
header(s, "Крок 3", "Структури даних та алгоритм")
body(s, [
    ("Для кожної сутності обери list / dict / set / dataclass.", "b"),
    ("Обґрунтуй вибір (доступ, порядок, пошук).", None),
    ("Напиши псевдокод ключової операції + оцінку O(...).", None),
])

s = add_slide()
header(s, "Крок 4", "Огляд еталона")
body(s, [
    ("Відкрий ядро power_calc_edu.", "b"),
    ("Знайди, як реалізовано ці сутності: models/ та services/.", None),
    ("Порівняй зі своїм планом — що збіглося, що зробив би інакше.", None),
])

s = add_slide()
header(s, "Що здаємо", "Артефакт")
tf = add_textbox(s, 0.9, 2.5, 11.5, 0.9)
para(tf, "students/<прізвище>/variant-NN/pr1/design.md", 22, TEAL, bold=True, first=True)
body(s, [
    ("постановка задачі;", None),
    ("декомпозиція (таблиця модулів);", None),
    ("обґрунтування структур даних;", None),
    ("псевдокод і складність;", None),
    ("порівняння з еталоном.", None),
], start_y=3.5, size=18, gap=8)

s = add_slide()
header(s, "Оцінювання", "2,5 бала")
body(s, [
    ("Постановка та сутності — 0,5", "b"),
    ("Декомпозиція — 0,5", "b"),
    ("Структури даних — 0,5", "b"),
    ("Псевдокод і складність — 0,5", "b"),
    ("Участь і захист — 0,5", "b"),
], size=20, gap=14)

s = add_slide()
header(s, "Типові помилки", "Тримайся критерію готовності")
body(s, [
    ("декомпозиція «по файлах», а не по відповідальностях;", None),
    ("структура обрана без обґрунтування;", None),
    ("складність без аналізу циклів/пошуку;", None),
    ("постановка без критерію готовності.", None),
], size=18, gap=10)

prs.save(OUT)
print("pptx збережено:", OUT)
print("слайдів:", len(prs.slides._sldIdLst))

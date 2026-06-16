# -*- coding: utf-8 -*-
"""
Генератор презентации для защиты ВКР «Тропа».

Структура — по шаблону «преза для диплома» (17 слайдов).
Дизайн — в стиле презентации Чмелева: светло-серый фон, крупные жирные
чёрные заголовки-гротеск, рукописный курсивный акцент-слово на каждом
слайде, монохром, много воздуха, асимметричная вёрстка.

Запуск:  python3 generate_pptx.py
Результат: Презентация_ВКР_Тропа.pptx
"""

import os

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# Логотип университета в левом верхнем углу слайдов.
# Положите файл сюда (PNG с прозрачным фоном — лучше всего).
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "assets", "tsu_logo.png")
LOGO_HEIGHT_IN = 0.62   # высота логотипа, дюймы

# --------------------------------------------------------------------------
# Палитра и шрифты
# --------------------------------------------------------------------------
BG      = RGBColor(0xEC, 0xEB, 0xE8)   # тёплый светло-серый фон
INK     = RGBColor(0x14, 0x14, 0x14)   # почти чёрный — заголовки
BODY    = RGBColor(0x2B, 0x2B, 0x2B)   # основной текст
MUTED   = RGBColor(0x8C, 0x89, 0x84)   # серые надписи-кикеры
LINE    = RGBColor(0xCB, 0xC9, 0xC4)   # тонкие линии
ACCENT  = RGBColor(0x2F, 0x6B, 0x53)   # приглушённый зелёный — только цифры
PANEL   = RGBColor(0xF6, 0xF5, 0xF2)   # светлая плашка таблиц
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)

F_HEAD   = "Arial Black"     # тяжёлый гротеск для заголовков
F_SCRIPT = "Segoe Script"    # рукописный акцент
F_BODY   = "Arial"           # основной текст

EMU_IN = 914400
SW, SH = 13.333, 7.5         # размеры слайда 16:9, дюймы


# --------------------------------------------------------------------------
# Низкоуровневые помощники
# --------------------------------------------------------------------------
def _set_spacing(run, centipoints):
    """Межбуквенный интервал (centipoints = 1/100 pt)."""
    rPr = run._r.get_or_add_rPr()
    rPr.set("spc", str(int(centipoints)))


def textbox(slide, l, t, w, h, wrap=True, anchor=None):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    if anchor is not None:
        tf.vertical_anchor = anchor
    return tb, tf


def _style_run(r, font, size, color, bold, italic, spacing):
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    if spacing is not None:
        _set_spacing(r, spacing)
    # кириллица: задаём шрифт и для cs/ea
    rPr = r._r.get_or_add_rPr()
    for tag in ("a:latin", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", font)


def para(tf, content, font=F_BODY, size=16, color=BODY, bold=False,
         italic=False, align=PP_ALIGN.LEFT, before=0, after=0, line=1.0,
         spacing=None, first=False):
    """
    content: строка ИЛИ список ранов [(text, {overrides}), ...].
    """
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(before)
    p.space_after = Pt(after)
    try:
        p.line_spacing = line
    except Exception:
        pass

    runs = content if isinstance(content, list) else [(content, {})]
    for text, ov in runs:
        r = p.add_run()
        r.text = text
        _style_run(
            r,
            ov.get("font", font),
            ov.get("size", size),
            ov.get("color", color),
            ov.get("bold", bold),
            ov.get("italic", italic),
            ov.get("spacing", spacing),
        )
    return p


def bg(slide):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = BG


def line(slide, l, t, w, h, color=LINE, weight=1.0):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t),
                                 Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def logo(slide, left=0.55, top=0.3, height=LOGO_HEIGHT_IN):
    """Эмблема университета в левом верхнем углу."""
    if os.path.exists(LOGO_PATH):
        slide.shapes.add_picture(LOGO_PATH, Inches(left), Inches(top),
                                 height=Inches(height))


def kicker(slide, left_text=None, right_text=None):
    """Левый угол — логотип; правый угол — увеличенная подпись-раздел."""
    logo(slide)
    if right_text:
        _, tf2 = textbox(slide, SW - 7.55, 0.34, 7.0, 0.5)
        para(tf2, right_text.upper(), font=F_BODY, size=15, color=MUTED,
             bold=True, align=PP_ALIGN.RIGHT, spacing=180, first=True)


def heading(slide, lines, top=0.7, size=54, left=0.5, width=12.55,
            align=PP_ALIGN.LEFT, color=INK, line_sp=0.94):
    if isinstance(lines, str):
        lines = [lines]
    h = 0.25 + size / 46.0 * len(lines)
    _, tf = textbox(slide, left, top, width, h, wrap=True)
    for i, ln in enumerate(lines):
        para(tf, ln, font=F_HEAD, size=size, color=color, align=align,
             line=line_sp, first=(i == 0))
    return tf


def script(slide, text, left, top, size=34, color=INK, width=6.5,
           align=PP_ALIGN.LEFT):
    _, tf = textbox(slide, left, top, width, size / 36.0 + 0.5, wrap=False)
    para(tf, text, font=F_SCRIPT, size=size, color=color, italic=True,
         align=align, first=True)
    return tf


def script_r(slide, text, top, size=30, color=BODY, right=12.6, width=9.0):
    """Рукописный акцент, прижатый к правому полю (не обрезается)."""
    _, tf = textbox(slide, right - width, top, width, size / 30.0 + 0.5,
                    wrap=False)
    para(tf, text, font=F_SCRIPT, size=size, color=color, italic=True,
         align=PP_ALIGN.RIGHT, first=True)
    return tf


def bullets(tf, items, size=16, color=BODY, gap=8, line_sp=1.06,
            bold_lead=False, first=True):
    """items: список строк или (lead, tail)."""
    for i, it in enumerate(items):
        if isinstance(it, tuple):
            lead, tail = it
            runs = [("•  ", {"color": INK, "bold": True}),
                    (lead, {"bold": True, "color": INK}),
                    (tail, {})]
        else:
            runs = [("•  ", {"color": INK, "bold": True}), (it, {})]
        para(tf, runs, size=size, color=color, line=line_sp,
             after=gap, first=(first and i == 0))


def labeled_block(tf, label, value, first=False, gap=14, vsize=17):
    para(tf, label.upper(), font=F_BODY, size=11, color=MUTED, bold=True,
         spacing=160, first=first, after=2)
    para(tf, value, font=F_BODY, size=vsize, color=INK, line=1.08, after=gap)


def stat(slide, number, caption, left, top, num_size=48, cap_w=3.9):
    _, tf = textbox(slide, left, top, cap_w, 1.7)
    para(tf, number, font=F_HEAD, size=num_size, color=ACCENT, first=True,
         line=0.95)
    for i, ln in enumerate(caption.split("\n")):
        para(tf, ln, font=F_BODY, size=12.5, color=BODY, line=1.08,
             before=6 if i == 0 else 0)


# --------------------------------------------------------------------------
# Таблицы
# --------------------------------------------------------------------------
def _cell_border(cell, edges=("bottom",), color=LINE, w_emu=9525):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tagmap = {"left": "a:lnL", "right": "a:lnR",
              "top": "a:lnT", "bottom": "a:lnB"}
    for edge in edges:
        tag = tagmap[edge]
        for old in tcPr.findall(qn(tag)):
            tcPr.remove(old)
        ln = tcPr.makeelement(qn(tag), {"w": str(w_emu), "cap": "flat"})
        fill = ln.makeelement(qn("a:solidFill"), {})
        clr = fill.makeelement(qn("a:srgbClr"),
                               {"val": "%02X%02X%02X" % (color[0], color[1], color[2])})
        fill.append(clr)
        ln.append(fill)
        tcPr.append(ln)


def set_cell(cell, text, bold=False, size=12.5, color=BODY, fill=None,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE):
    cell.vertical_anchor = anchor
    cell.margin_left = Inches(0.12)
    cell.margin_right = Inches(0.12)
    cell.margin_top = Inches(0.05)
    cell.margin_bottom = Inches(0.05)
    if fill is None:
        cell.fill.background()
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    _style_run(r, F_BODY, size, color, bold, False, None)


def make_table(slide, rows, col_widths, left, top, row_h=0.4,
               header=True, head_fill=INK, head_color=WHITE,
               body_size=12.5, head_size=12.5, accent_last=False):
    nrows = len(rows)
    ncols = len(col_widths)
    width = sum(col_widths)
    gr = slide.shapes.add_table(nrows, ncols, Inches(left), Inches(top),
                                Inches(width), Inches(row_h * nrows))
    tbl = gr.table
    # отключаем стандартный стиль (полосы/рамки темы)
    tbl.first_row = False
    tbl.horz_banding = False
    for j, cw in enumerate(col_widths):
        tbl.columns[j].width = Inches(cw)
    for i in range(nrows):
        tbl.rows[i].height = Inches(row_h)
        for j in range(ncols):
            val = rows[i][j]
            is_head = header and i == 0
            if is_head:
                set_cell(tbl.cell(i, j), val, bold=True, size=head_size,
                         color=head_color, fill=head_fill,
                         align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
            else:
                fill = PANEL if (i % 2 == 0) else BG
                col = INK if j == 0 else BODY
                bold = (j == 0)
                if accent_last and j == ncols - 1:
                    col = ACCENT
                    bold = True
                set_cell(tbl.cell(i, j), val, bold=bold, size=body_size,
                         color=col, fill=fill,
                         align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
                _cell_border(tbl.cell(i, j), ("bottom",), LINE)
    return tbl


# --------------------------------------------------------------------------
# Сборка презентации
# --------------------------------------------------------------------------
prs = Presentation()
prs.slide_width = Emu(int(SW * EMU_IN))
prs.slide_height = Emu(int(SH * EMU_IN))
BLANK = prs.slide_layouts[6]


def new_slide(left_kick=None, right_kick=None):
    s = prs.slides.add_slide(BLANK)
    bg(s)
    if left_kick is not None:
        kicker(s, left_kick, right_kick)
    return s


KICK = "ИС «Тропа» · защита ВКР"


def foot(slide, n):
    _, tf = textbox(slide, SW - 2.6, SH - 0.62, 2.15, 0.45, wrap=False)
    para(tf, f"{n:02d} / 17", font=F_BODY, size=13.5, color=MUTED, bold=True,
         align=PP_ALIGN.RIGHT, spacing=140, first=True)


# ---- Слайд 1. Титульный --------------------------------------------------
s = new_slide()
kicker(s, "Бакалаврская работа", "ТулГУ · Тула, 2026")
line(s, 0.55, 0.78, SW - 1.1, 0.014, LINE)
script(s, "доступный туризм", 0.6, 2.02, size=30, color=BODY)
heading(s, "Тропа", top=2.35, size=150, left=0.45)
_, tf = textbox(s, 0.6, 4.65, 9.2, 1.4)
para(tf, "Веб-сервис построения туристических маршрутов по Тульской "
         "области для людей с ограниченными возможностями здоровья",
     font=F_BODY, size=18, color=BODY, line=1.18, first=True)
line(s, 0.55, 6.25, SW - 1.1, 0.012, LINE)
_, tf = textbox(s, 0.6, 6.45, 8.0, 0.9)
para(tf, [("Студент гр. 221321  —  ", {"color": MUTED}),
          ("Шулепова Д. А.", {"bold": True, "color": INK})],
     size=14, first=True, after=4)
para(tf, [("Руководитель  —  ", {"color": MUTED}),
          ("Сафронова М. А.", {"bold": True, "color": INK})], size=14)
_, tf = textbox(s, SW - 4.6, 6.55, 4.0, 0.6, wrap=False)
para(tf, "09.03.03 Прикладная информатика", font=F_BODY, size=12.5,
     color=MUTED, align=PP_ALIGN.RIGHT, first=True)

# ---- Слайд 2. О работе ---------------------------------------------------
s = new_slide(KICK, "о работе")
heading(s, "О работе", top=0.95, size=66)
script_r(s, "коротко", 1.32, size=30)
line(s, 0.55, 2.35, SW - 1.1, 0.012, LINE)
_, tf = textbox(s, 0.6, 2.7, 6.0, 4.3)
labeled_block(tf, "Объект исследования",
              "Процессы информационной поддержки планирования "
              "туристических поездок по объектам Тульской области для "
              "граждан с ОВЗ", first=True)
labeled_block(tf, "Предмет исследования",
              "Модели, алгоритмы и программные средства построения "
              "инклюзивных маршрутов с учётом доступности среды")
labeled_block(tf, "Цель работы",
              "Разработка веб-сервиса (ИС «Тропа») построения "
              "туристических маршрутов с учётом потребностей "
              "пользователя в доступности среды")
_, tf = textbox(s, 7.1, 2.7, 5.7, 4.3)
para(tf, "ЗАДАЧИ", font=F_BODY, size=11, color=MUTED, bold=True,
     spacing=160, first=True, after=8)
bullets(tf, [
    "анализ предметной области и обзор систем-аналогов",
    "модель оценки доступности по категориям ОВЗ",
    "архитектура системы и структура базы данных",
    "алгоритм построения инклюзивного маршрута",
    "методика автосбора сведений (обработка текста)",
    "реализация сервера и клиента, тестирование",
    "меры информационной безопасности и защиты ПДн",
], size=14.5, gap=7, first=False)
foot(s, 2)

# ---- Слайд 3. Актуальность ----------------------------------------------
s = new_slide(KICK, "актуальность")
heading(s, "Актуальность", top=0.95, size=64)
script_r(s, "проблема", 1.32, size=30)
line(s, 0.55, 2.35, SW - 1.1, 0.012, LINE)
stat(s, "10 млн+", "граждан с инвалидностью\nв РФ (2024)", 0.6, 2.75, num_size=46)
stat(s, "~30 %", "населения нужна адаптированная\nтуристическая информация", 4.85, 2.75, num_size=46)
stat(s, "500 тыс.+", "туристов в Тульской\nобласти ежегодно", 9.1, 2.75, num_size=46)
_, tf = textbox(s, 0.6, 4.75, 6.0, 2.4)
para(tf, "ПРОБЛЕМА", font=F_BODY, size=11, color=MUTED, bold=True,
     spacing=160, first=True, after=6)
bullets(tf, [
    "сведения о доступности объектов фрагментарны",
    "сервисы (Tripadvisor, Яндекс Карты, 2GIS) не "
    "учитывают потребности людей с ОВЗ",
    "планирование поездки трудоёмко и часто срывается",
], size=14.5, gap=7, first=False)
_, tf = textbox(s, 7.1, 4.75, 5.7, 2.4)
para(tf, "РЕШЕНИЕ", font=F_BODY, size=11, color=MUTED, bold=True,
     spacing=160, first=True, after=6)
para(tf, [("ИС «Тропа»", {"bold": True, "color": INK}),
          (" — специализированный веб-сервис с многопрофильной оценкой "
           "доступности и построением инклюзивных маршрутов по региону.",
           {})], size=15, line=1.16)
foot(s, 3)

# ---- Слайд 4. Системы-аналоги -------------------------------------------
s = new_slide(KICK, "анализ")
heading(s, "Системы-аналоги", top=0.95, size=58)
script_r(s, "сравнение", 1.32, size=28)
line(s, 0.55, 2.3, SW - 1.1, 0.012, LINE)
rows = [
    ["Критерий", "Tripadvisor", "Яндекс Карты", "Wheelmap", "Визиттула", "«Тропа»"],
    ["Каталог объектов", "да", "да", "част.", "да", "да"],
    ["Сведения о доступности", "миним.", "миним.", "только мобильн.", "нет", "развёрнутые"],
    ["Учёт категорий ОВЗ", "нет", "нет", "нет", "нет", "9 категорий"],
    ["Построение маршрутов", "нет", "без ОВЗ", "нет", "нет", "с учётом ОВЗ"],
    ["Автосбор данных (LLM)", "нет", "нет", "нет", "нет", "да"],
    ["Специализация на регионе", "нет", "нет", "нет", "Тула", "Тула"],
]
make_table(s, rows, [3.3, 1.6, 1.45, 1.55, 1.5, 2.0], 0.6, 2.65,
           row_h=0.58, body_size=12, head_size=12, accent_last=True)
foot(s, 4)

# ---- Слайд 5. Структура системы -----------------------------------------
s = new_slide(KICK, "структура")
heading(s, "Структура системы", top=0.95, size=56)
script_r(s, "функции", 1.32, size=28)
line(s, 0.55, 2.3, SW - 1.1, 0.012, LINE)
_, tf = textbox(s, 0.6, 2.65, 6.1, 4.5)
para(tf, "9 ФУНКЦИОНАЛЬНЫХ ПОДСИСТЕМ · 25 ФУНКЦИЙ", font=F_BODY, size=11,
     color=MUTED, bold=True, spacing=140, first=True, after=8)
bullets(tf, [
    "аутентификация и управление профилем",
    "каталог объектов и паспорта доступности",
    "картографическое отображение (Яндекс.Карты)",
    "построение инклюзивных маршрутов",
    "автоматизированный сбор сведений (LLM)",
    "модерация, отзывы, администрирование",
], size=14, gap=6, first=False)
_, tf = textbox(s, 7.1, 2.65, 5.7, 4.5)
para(tf, "РОЛИ ПОЛЬЗОВАТЕЛЕЙ (RBAC)", font=F_BODY, size=11, color=MUTED,
     bold=True, spacing=140, first=True, after=8)
bullets(tf, [
    ("user", " — каталог, маршруты, отзывы"),
    ("expert", " — шаблоны маршрутов"),
    ("moderator", " — каталог, модерация автосбора"),
    ("admin", " — управление и конфигурация"),
], size=14, gap=6, first=False)
para(tf, "ИНКЛЮЗИВНЫЙ ПРОФИЛЬ", font=F_BODY, size=11, color=MUTED,
     bold=True, spacing=140, before=10, after=6)
para(tf, [("9 категорий ограничений здоровья", {"bold": True, "color": INK}),
          (" — независимая оценка по каждой; учёт комбинаций потребностей "
           "пользователя.", {})], size=14, line=1.12)
foot(s, 5)

# ---- Слайд 6. Архитектура и математическая модель -----------------------
s = new_slide(KICK, "ядро")
heading(s, ["Архитектура", "и модель"], top=0.72, size=50, line_sp=0.95)
script_r(s, "ядро", 1.85, size=30)
line(s, 0.55, 2.55, SW - 1.1, 0.012, LINE)
_, tf = textbox(s, 0.6, 2.85, 5.9, 4.3)
para(tf, "АРХИТЕКТУРА", font=F_BODY, size=11, color=MUTED, bold=True,
     spacing=160, first=True, after=8)
bullets(tf, [
    "клиент-серверная, 4 уровня (представление, логика, данные, интеграции)",
    "сервер — Python 3.12 + FastAPI (REST/JSON)",
    "клиент — TypeScript + React + Vite (SPA)",
    "СУБД — PostgreSQL 16 + PostGIS, кеш Redis",
    "развёртывание в Docker",
], size=13.8, gap=7, first=False)
_, tf = textbox(s, 6.85, 2.85, 6.0, 4.3)
para(tf, "МАТЕМАТИЧЕСКАЯ МОДЕЛЬ", font=F_BODY, size=11, color=MUTED,
     bold=True, spacing=160, first=True, after=8)
para(tf, "Модифицированный алгоритм Дейкстры — вес ребра:",
     size=13.8, color=BODY, after=6, line=1.1)
para(tf, "w(vᵢ, vⱼ, π) = α·d(vᵢ,vⱼ) + β·(1 − f)·D",
     font=F_BODY, size=15.5, color=INK, bold=True, after=6)
para(tf, [("d", {"bold": True}), (" — расстояние (гаверсинус),  ", {}),
          ("f∈[0,1]", {"bold": True}),
          (" — оценка доступности под профиль.", {})],
     size=13.5, line=1.12, after=10)
para(tf, "Оценка маршрута — по «слабому звену»:", size=13.8, after=6, line=1.1)
para(tf, "F(S, π) = min { f(p(sᵢ), π) }", font=F_BODY, size=15.5,
     color=INK, bold=True, after=6)
para(tf, "Сложность O((|V|+|E|)·log|V|); маршрут до 5 точек < 1,5 с.",
     size=13, color=BODY, line=1.12)
foot(s, 6)

# ---- Слайд 7. Оценка доступности (реализация модели) --------------------
s = new_slide(KICK, "реализация")
heading(s, "Оценка доступности", top=0.95, size=54)
script_r(s, "9 профилей", 1.32, size=28)
line(s, 0.55, 2.3, SW - 1.1, 0.012, LINE)
_, tf = textbox(s, 0.6, 2.65, 6.1, 4.4)
para(tf, "МОДЕЛЬ ПРАВИЛ", font=F_BODY, size=11, color=MUTED, bold=True,
     spacing=160, first=True, after=8)
bullets(tf, [
    ("блокирующие правила", " — нарушение делает объект "
     "недоступным (f = 0)"),
    ("желательные правила", " — снижают оценку, но не блокируют"),
    ("несколько профилей", " — берётся минимум (консервативно)"),
], size=14.5, gap=8, first=False)
para(tf, "ПАСПОРТ ДОСТУПНОСТИ", font=F_BODY, size=11, color=MUTED,
     bold=True, spacing=160, before=12, after=6)
para(tf, "51 атрибут в 7 категориях; показатель заполненности c(p) для "
         "приоритизации работы модераторов.", size=14, line=1.14)
_, tf = textbox(s, 7.1, 2.65, 5.7, 4.4)
para(tf, "ШКАЛА ОЦЕНКИ  f ∈ [0, 1]", font=F_BODY, size=11, color=MUTED,
     bold=True, spacing=160, first=True, after=10)
para(tf, [("f ≥ 0.7  ", {"font": F_HEAD, "size": 18, "color": ACCENT}),
          ("— доступен", {"size": 15})], after=8)
para(tf, [("0.3 ≤ f < 0.7  ", {"font": F_HEAD, "size": 18, "color": INK}),
          ("— частично", {"size": 15})], after=8)
para(tf, [("f < 0.3  ", {"font": F_HEAD, "size": 18, "color": INK}),
          ("— недоступен", {"size": 15})], after=14)
para(tf, "На карте сегменты маршрута подсвечиваются: зелёный — доступно, "
         "жёлтый — частично, красный — недоступно.", size=13.5,
     color=BODY, line=1.16)
foot(s, 7)

# ---- Слайд 8. Целевая аудитория -----------------------------------------
s = new_slide(KICK, "аудитория")
heading(s, "Целевая аудитория", top=0.95, size=56)
script_r(s, "сегменты", 1.32, size=28)
line(s, 0.55, 2.3, SW - 1.1, 0.012, LINE)
rows = [
    ["Категория", "Особые потребности", "Доля, %"],
    ["Маломобильные граждане", "пандусы, лифты, отсутствие ступеней", "8–10"],
    ["Нарушения зрения", "тактильная плитка, аудиогид, контраст", "1–2"],
    ["Нарушения слуха", "субтитры, индукционные петли, жесты", "2–3"],
    ["Когнитивные особенности", "предсказуемость, ясный язык, тишина", "1–2"],
    ["Пожилые граждане", "места отдыха, низкая нагрузка", "15–18"],
    ["Семьи с детьми", "доступность для колясок, комнаты МиР", "10–12"],
]
make_table(s, rows, [3.6, 6.4, 2.1], 0.6, 2.6, row_h=0.6,
           body_size=12.5, accent_last=True)
foot(s, 8)

# ---- Слайд 9. Бизнес-модель ---------------------------------------------
s = new_slide(KICK, "бизнес-модель")
heading(s, "Бизнес-модель", top=0.95, size=60)
script_r(s, "канвас", 1.32, size=28)
line(s, 0.55, 2.3, SW - 1.1, 0.012, LINE)
rows = [
    ["Компонент", "Содержание"],
    ["Ценностное предложение", "Самостоятельное планирование инклюзивных поездок; достоверные паспорта доступности"],
    ["Сегменты клиентов", "Минтуризма региона; турфирмы инклюзивных туров; ассоциации людей с ОВЗ"],
    ["Каналы", "Госконтракты, профильные мероприятия, партнёрство с НКО"],
    ["Потоки дохода", "Контракт на разработку и сопровождение; тиражирование на регионы"],
    ["Ключевые ресурсы", "Команда, база паспортов доступности, технологии (LLM, ГИС)"],
    ["Ключевые партнёры", "Яндекс (Карты, GPT), организации инвалидов, музеи региона"],
]
make_table(s, rows, [3.6, 8.5], 0.6, 2.6, row_h=0.62, body_size=12.5)
foot(s, 9)

# ---- Слайд 10. Финансовая модель ----------------------------------------
s = new_slide(KICK, "финансы")
heading(s, "Финансовая модель", top=0.95, size=56)
script_r(s, "3 года", 1.32, size=28)
line(s, 0.55, 2.3, SW - 1.1, 0.012, LINE)
rows = [
    ["Показатель, тыс. руб.", "1-й год", "2-й год", "3-й год"],
    ["Выручка", "3 600", "9 000", "16 200"],
    ["Расходы (с налогом)", "3 513", "5 460", "8 940"],
    ["Чистая прибыль", "87", "3 540", "7 260"],
    ["Новых регионов за год", "—", "+2", "+3"],
]
make_table(s, rows, [4.6, 2.5, 2.5, 2.5], 0.6, 2.6, row_h=0.6, body_size=13.5)
_, tf = textbox(s, 0.6, 5.7, 12.1, 1.3)
para(tf, [("Старт — пилот для Минтуризма Тульской области, далее "
           "тиражирование на субъекты РФ.  ", {"color": BODY}),
          ("Суммарная чистая прибыль за 3 года — 10 887 тыс. руб.",
           {"bold": True, "color": INK})], size=14.5, line=1.2, first=True)
foot(s, 10)

# ---- Слайд 11. Интеграция и автосбор ------------------------------------
s = new_slide(KICK, "интеграция")
heading(s, ["Интеграция", "и автосбор"], top=0.72, size=50, line_sp=0.95)
script_r(s, "ИИ в контуре", 1.85, size=28)
line(s, 0.55, 2.55, SW - 1.1, 0.012, LINE)
_, tf = textbox(s, 0.6, 2.85, 6.1, 4.3)
para(tf, "ВНЕШНИЕ СЕРВИСЫ", font=F_BODY, size=11, color=MUTED, bold=True,
     spacing=160, first=True, after=8)
bullets(tf, [
    ("YandexGPT", " — извлечение сведений о доступности из текста"),
    ("Яндекс.Карты JS API 3.0", " — карта и геометрия маршрута"),
    ("OpenRouteService", " — построение пешеходных сегментов"),
    ("Redis", " — кеширование справочников"),
], size=14, gap=8, first=False)
_, tf = textbox(s, 7.1, 2.85, 5.7, 4.3)
para(tf, "АВТОСБОР: ЧЕЛОВЕК В КОНТУРЕ", font=F_BODY, size=11, color=MUTED,
     bold=True, spacing=140, first=True, after=8)
para(tf, "парсинг страницы → LLM (JSON) → очередь модерации → "
         "подтверждение модератором → паспорт объекта",
     size=14, color=INK, line=1.18, after=12)
para(tf, [("Качество извлечения  ", {"color": BODY, "size": 14}),
          ("F1 = 0,81", {"font": F_HEAD, "size": 26, "color": ACCENT})],
     after=4)
para(tf, "микро-усреднение на контрольной выборке; данные не "
         "применяются без модерации.", size=12.5, color=BODY, line=1.14)
foot(s, 11)

# ---- Слайд 12. Информационная безопасность ------------------------------
s = new_slide(KICK, "безопасность")
heading(s, "Безопасность", top=0.95, size=60)
script_r(s, "152-ФЗ", 1.32, size=28)
line(s, 0.55, 2.3, SW - 1.1, 0.012, LINE)
_, tf = textbox(s, 0.6, 2.65, 6.1, 4.4)
para(tf, "ЗАЩИТА ПЕРСОНАЛЬНЫХ ДАННЫХ", font=F_BODY, size=11, color=MUTED,
     bold=True, spacing=140, first=True, after=8)
bullets(tf, [
    ("правовые", " — согласие на обработку ПДн, политика "
     "конфиденциальности, соответствие 152-ФЗ"),
    ("технические", " — bcrypt, JWT, HTTPS/TLS 1.2+, RBAC"),
    ("организационные", " — разделение полномочий ролей"),
], size=14, gap=8, first=False)
_, tf = textbox(s, 7.1, 2.65, 5.7, 4.4)
para(tf, "МОДЕЛЬ УГРОЗ", font=F_BODY, size=11, color=MUTED, bold=True,
     spacing=160, first=True, after=10)
para(tf, [("10", {"font": F_HEAD, "size": 30, "color": ACCENT}),
          ("  актуальных угроз (по банку ФСТЭК)", {"size": 14})], after=8)
para(tf, [("4", {"font": F_HEAD, "size": 30, "color": INK}),
          ("  категории нарушителя", {"size": 14})], after=8)
para(tf, [("7", {"font": F_HEAD, "size": 30, "color": INK}),
          ("  направлений мер защиты", {"size": 14})], after=8)
para(tf, "Комплекс мер перекрывает все выявленные актуальные угрозы.",
     size=13, color=BODY, line=1.16)
foot(s, 12)

# ---- Слайд 13. Результаты и технология работы ---------------------------
s = new_slide(KICK, "результаты")
heading(s, "Результаты", top=0.95, size=62)
script_r(s, "контрольный пример", 1.4, size=26)
line(s, 0.55, 2.35, SW - 1.1, 0.012, LINE)
_, tf = textbox(s, 0.6, 2.7, 6.0, 4.3)
para(tf, "КОНТРОЛЬНЫЙ ПРИМЕР · РУЧНАЯ КОЛЯСКА", font=F_BODY, size=11,
     color=MUTED, bold=True, spacing=130, first=True, after=8)
bullets(tf, [
    ("время входа — 350 мс", "  (требование ≤ 1000 мс)"),
    ("построение маршрута — 1,2 с", "  (≤ 3 с)"),
    ("оценка 0,72 — accessible", "  ·  длина 2,1 км, 85 мин"),
    ("1 предупреждение", "  (брусчатка у входа в музей)"),
], size=14, gap=8, first=False)
_, tf = textbox(s, 7.1, 2.7, 5.7, 4.3)
para(tf, "ГОТОВНОСТЬ И ТЕСТИРОВАНИЕ", font=F_BODY, size=11, color=MUTED,
     bold=True, spacing=140, first=True, after=8)
bullets(tf, [
    "20 объектов с верифицированными паспортами",
    "18 маршрутов-шаблонов по 6 категориям ОВЗ",
    "тестирование: модульное, интеграционное, приёмочное",
    "руководства пользователя и модератора",
], size=14, gap=8, first=False)
foot(s, 13)

# ---- Слайд 14. Экономическая эффективность ------------------------------
s = new_slide(KICK, "экономика")
heading(s, "Экономика", top=0.95, size=64)
script_r(s, "инвестиции", 1.4, size=28)
line(s, 0.55, 2.35, SW - 1.1, 0.012, LINE)
stat(s, "500 тыс.", "стартовый капитал, руб.", 0.6, 2.85, num_size=46)
stat(s, "350 тыс.", "ежемесячные расходы, руб.", 4.85, 2.85, num_size=46)
stat(s, "80 %", "расходов — фонд\nоплаты труда", 9.1, 2.85, num_size=46)
_, tf = textbox(s, 0.6, 4.85, 12.1, 2.0)
para(tf, "МОДЕЛЬ ДОХОДА", font=F_BODY, size=11, color=MUTED, bold=True,
     spacing=160, first=True, after=8)
bullets(tf, [
    ("контракт 3 600 тыс. руб./год", " в первый год (разработка + "
     "сопровождение), далее 1 800 тыс. руб./год"),
    ("команда из 4 человек", "; форма — ИП на УСН «Доходы» (6 %)"),
    ("тиражирование на регионы", " — основной драйвер роста выручки"),
], size=14, gap=7, first=False)
foot(s, 14)

# ---- Слайд 15. Окупаемость и прибыль ------------------------------------
s = new_slide(KICK, "окупаемость")
heading(s, "Окупаемость", top=0.95, size=60)
script_r(s, "ROI", 1.32, size=30)
line(s, 0.55, 2.35, SW - 1.1, 0.012, LINE)
stat(s, "15 мес.", "срок окупаемости\nстартовых инвестиций", 0.6, 2.85, num_size=46)
stat(s, "2077 %", "ROI за 3 года\n(≈178 % годовых)", 4.85, 2.85, num_size=46)
stat(s, "7,25 млн", "NPV за 3 года\nпри ставке r = 15 %", 9.1, 2.85, num_size=44)
_, tf = textbox(s, 0.6, 5.0, 12.1, 1.8)
para(tf, "ЧИСТАЯ ПРИБЫЛЬ ПО ГОДАМ", font=F_BODY, size=11, color=MUTED,
     bold=True, spacing=160, first=True, after=8)
para(tf, [("1-й год — ", {"color": BODY}), ("87 тыс.", {"bold": True, "color": INK}),
          ("     2-й год — ", {"color": BODY}), ("3 540 тыс.", {"bold": True, "color": INK}),
          ("     3-й год — ", {"color": BODY}), ("7 260 тыс. руб.", {"bold": True, "color": INK})],
     size=16)
para(tf, "Положительный NPV подтверждает экономическую целесообразность "
         "проекта.", size=13.5, color=BODY, before=8, line=1.16)
foot(s, 15)

# ---- Слайд 16. Заключение -----------------------------------------------
s = new_slide(KICK, "заключение")
heading(s, "Заключение", top=0.95, size=62)
script_r(s, "итоги", 1.32, size=30)
line(s, 0.55, 2.35, SW - 1.1, 0.012, LINE)
_, tf = textbox(s, 0.6, 2.7, 6.0, 4.3)
para(tf, "ЦЕЛЬ ДОСТИГНУТА, ЗАДАЧИ ВЫПОЛНЕНЫ", font=F_BODY, size=11,
     color=MUTED, bold=True, spacing=130, first=True, after=8)
bullets(tf, [
    "разработан работающий веб-сервис «Тропа»",
    "20 объектов, 18 маршрутов-шаблонов, ~14 000 строк кода",
    "система готова к опытной эксплуатации",
], size=14, gap=8, first=False)
para(tf, "ПРАКТИЧЕСКАЯ ЗНАЧИМОСТЬ", font=F_BODY, size=11, color=MUTED,
     bold=True, spacing=140, before=12, after=6)
para(tf, "Самостоятельное планирование поездок гражданами с ОВЗ; "
         "тиражируемость на другие субъекты РФ.", size=14, line=1.16)
_, tf = textbox(s, 7.1, 2.7, 5.7, 4.3)
para(tf, "НАУЧНАЯ НОВИЗНА", font=F_BODY, size=11, color=MUTED, bold=True,
     spacing=140, first=True, after=8)
bullets(tf, [
    "многопрофильная модель доступности (9 категорий ОВЗ, учёт комбинаций через минимум)",
    "модификация алгоритма Дейкстры со штрафом за недоступность объекта",
    "методика автосбора сведений LLM с обязательной модерацией (human-in-the-loop)",
], size=14, gap=9, first=False)
foot(s, 16)

# ---- Слайд 17. Спасибо ---------------------------------------------------
s = new_slide()
kicker(s, "ИС «Тропа» · защита ВКР", "ТулГУ · 2026")
line(s, 0.55, 0.78, SW - 1.1, 0.014, LINE)
heading(s, "Спасибо", top=2.2, size=120)
script_r(s, "за внимание!", 4.35, size=52, color=INK)
_, tf = textbox(s, 0.6, 6.1, 11.0, 1.0)
para(tf, "Шулепова Д. А., гр. 221321  ·  Руководитель — Сафронова М. А.",
     size=14, color=MUTED, first=True)
para(tf, "Веб-сервис построения инклюзивных туристических маршрутов "
         "по Тульской области", size=13, color=BODY, before=4)

# --------------------------------------------------------------------------
out = "Презентация_ВКР_Тропа.pptx"
prs.save(out)
print("saved", out, "slides:", len(prs.slides._sldIdLst))

#!/usr/bin/env python3
"""Презентация для CEO по приоритетам реинжиниринга процессов ДУП.

Короткий доклад на одно решение: в каком порядке ДУП берётся за процессы
и почему именно в таком. Все цифры, объекты, волны и правила очерёдности
импортируются из build_dup_priority.py — того же модуля, который собирает
«ДУП_приоритеты_реинжиниринга.xlsx», поэтому слайды не могут разойтись
с реестром.

Фирменный стиль — из корпоративного шаблона «Общие слайды_ver 12.12.24.pptx»
через brand.py: мастера, макеты, логотип, футер и нумерация наследуются.
"""

from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.packuri import PackURI
from pptx.util import Emu, Pt

from brand import (
    BLUE,
    BLUE_LIGHT,
    CARD_BG,
    CONTENT_TOP,
    CONTENT_W,
    DARK,
    FONT,
    FONT_BOLD,
    FONT_MED,
    GRAY,
    GREEN,
    LAYOUT_TITLE,
    MARGIN_L,
    MIST,
    RED,
    STEEL,
    TEXT,
    WHITE,
    YELLOW,
    accent_bar,
    card,
    divider,
    ink,
    on,
    rect,
    set_text,
    stat,
    table_grid,
    textbox,
    title,
)

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE.parent / "01_Процессы"))

import build_dup_priority as P  # noqa: E402

TEMPLATE = BASE.parent / "brandbook" / "Общие слайды_ver 12.12.24.pptx"
OUTPUT = BASE / "ДУП_Приоритеты_реинжиниринга_презентация.pptx"

DECK_TITLE = "Приоритеты реинжиниринга"
DECK_SUBTITLE = "С чего начинаем и почему именно с этого"
DECK_DATE = "Сентябрь 2026"

RELS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

# Цвет волны на слайдах: та же логика, что у заливок в Excel — от синего
# «фундамента» к тёплому «потом».
WAVE_COLORS = {
    P.В1: BLUE,
    P.В2: GREEN,
    P.В3: YELLOW,
    P.В4: BLUE_LIGHT,
}


# --------------------------------------------------------------------------
# Данные
# --------------------------------------------------------------------------
ROWS = P.objects()
SCORED = [r for r in ROWS if r["total"] is not None]
BY_WAVE = Counter(r["wave"] for r in ROWS)
BY_QUADRANT = Counter(r["quadrant"] for r in SCORED)
NO_OWNER = [r for r in ROWS if r["role"] == "Не определён"]


def wave_1_breakdown():
    """Разбивка первой волны на три группы.

    Название волны — «Владение и базовая линия», и её размер легко принять
    за число «ничьих» зон. На деле зон без владельца меньше: волна включает
    ещё объекты, где владелец формально есть, но роль A внутри ДУП не выделена,
    и подготовку замера. Группы считаются по данным, а не вписываются руками.
    """
    wave = [r for r in ROWS if r["wave"] == P.В1]
    groups = [
        ([r for r in wave if r["role"] == "Не определён"],
         "зон без владельца",
         "Владелец не закреплён ни в должностных инструкциях, ни в кадровой матрице"),
        ([r for r in wave if r["role"] != "Не определён" and r["kind"] == P.НАЗНАЧЕНИЕ],
         "с размытым владением",
         "Процесс принадлежит ДУП, но единственная роль A внутри ДУП не выделена"),
        ([r for r in wave if r["role"] != "Не определён" and r["kind"] != P.НАЗНАЧЕНИЕ],
         "замер и стык с ЦО",
         "Пилотный аудит, контроль качества данных, матрица кураторов, срок приказов HR"),
    ]
    if sum(len(rows) for rows, *_ in groups) != len(wave):
        raise SystemExit("Разбивка волны 1 не покрывает все её объекты")
    return [(len(rows), label, note) for rows, label, note in groups]


def wave_no(wave: str) -> str:
    """«Волна 2. Быстрые победы…» -> «2»."""
    return wave.split(".")[0].replace("Волна", "").strip()


def wave_name(wave: str) -> str:
    """Часть после номера: «Быстрые победы внутри ДУП»."""
    return wave.split(". ", 1)[1] if ". " in wave else wave


def clip(text: str, limit: int) -> str:
    """Обрезает по границе слова — в ячейку таблицы слайда влезает не всё."""
    if len(text) <= limit:
        return text
    cut = text[:limit].rsplit(" ", 1)[0]
    return cut.rstrip(" ,;·—-") + "…"


# --------------------------------------------------------------------------
# Работа со слайдами шаблона
# --------------------------------------------------------------------------
def drop_slide(prs: Presentation, index: int) -> None:
    id_list = prs.slides._sldIdLst
    items = list(id_list)
    prs.part.drop_rel(items[index].get(RELS))
    id_list.remove(items[index])


def move_slide_to_end(prs: Presentation, index: int) -> None:
    id_list = prs.slides._sldIdLst
    items = list(id_list)
    id_list.remove(items[index])
    id_list.append(items[index])


def reserve_partname(slide, number: int) -> None:
    """Сдвигает имя части слайда, чтобы новые слайды не заняли тот же путь."""
    slide.part.partname = PackURI(f"/ppt/slides/slide{number}.xml")


def find_shape(slide, needle: str):
    for sh in slide.shapes:
        if sh.has_text_frame and needle.lower() in sh.text_frame.text.lower():
            return sh
    return None


def replace_text(shape, lines, *, size, font=FONT_BOLD, color=DARK, align=PP_ALIGN.LEFT):
    tf = shape.text_frame
    tf.clear()
    set_text(tf, lines, size=size, font=font, color=color, align=align, line_spacing=1.15)


def new_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[LAYOUT_TITLE])


def numbered_row(slide, x, y, w, h, number, heading, body, *,
                 accent=BLUE, heading_size=11, body_size=9):
    """Строка «номер в кружке + заголовок + пояснение» — основной приём деka."""
    badge = rect(slide, x, y, 400000, 400000, fill=accent, radius=0.5)
    tf = badge.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    set_text(tf, str(number), size=12, font=FONT_BOLD, color=on(accent), align=PP_ALIGN.CENTER)

    text_x = x + 540000
    text_w = w - 540000
    textbox(slide, text_x, y - 10000, text_w, 250000, heading,
            size=heading_size, font=FONT_BOLD, color=DARK)
    if body:
        textbox(slide, text_x, y + 250000, text_w, h - 250000, body, size=body_size, color=GRAY)


# --------------------------------------------------------------------------
# Слайды
# --------------------------------------------------------------------------
def build_cover(prs: Presentation) -> None:
    slide = prs.slides[0]
    text_x = 373227

    note = find_shape(slide, "Выберите титулку")
    if note is not None:
        note._element.getparent().remove(note._element)

    name_box = find_shape(slide, "Название презентации")
    if name_box is not None:
        name_box.left = Emu(text_x)
        name_box.top = Emu(1320000)
        name_box.width = Emu(4100000)
        name_box.height = Emu(900000)
        replace_text(name_box, DECK_TITLE, size=25, color=WHITE)

    date_box = find_shape(slide, "Дата")
    if date_box is not None:
        date_box.left = Emu(text_x)
        date_box.top = Emu(3420000)
        replace_text(date_box, DECK_DATE, size=10, font=FONT, color=STEEL)

    rect(slide, text_x + 60000, 2440000, 620000, 30000, fill=GREEN, shape=MSO_SHAPE.RECTANGLE)
    textbox(slide, text_x + 60000, 2600000, 4100000, 300000, DECK_SUBTITLE,
            size=12.5, font=FONT_MED, color=GREEN)
    textbox(slide, text_x + 60000, 2990000, 4100000, 280000,
            "Департамент управления проектами · ТОО «Көркем Телеком»",
            size=9.5, color=MIST)


def build_headline(prs: Presentation) -> None:
    """Слайд решения: одна мысль и четыре цифры под ней."""
    slide = new_slide(prs)
    title(slide, "Очередь реинжиниринга построена")

    accent_bar(
        slide, MARGIN_L, CONTENT_TOP - 30000, CONTENT_W, 700000,
        ["Первыми берём не самые выгодные процессы, а те, что разблокируют остальные.",
         "Пока у процесса нет владельца и замера «до», выгоду от его оптимизации нечем доказать."],
        fill=BLUE, size=12, align=PP_ALIGN.LEFT,
    )

    stats = [
        (str(len(ROWS)), "объектов реестра оценены и поставлены в очередь", GREEN),
        (str(len(NO_OWNER)), "зон без владельца — они блокируют старт", RED),
        (str(BY_QUADRANT["Быстрые победы"]), "быстрых побед: высокая ценность, лёгкая реализация", GREEN),
        (str(len(P.FIRST_STEPS)), "первых шагов — без бюджета и ИТ-разработки", BLUE),
    ]
    col_w = (CONTENT_W - 3 * 160000) // 4
    y = CONTENT_TOP + 910000
    card_h = 1700000
    for i, (number, caption, accent) in enumerate(stats):
        x = MARGIN_L + i * (col_w + 160000)
        rect(slide, x, y, col_w, card_h, fill=CARD_BG)
        textbox(slide, x + 160000, y + 200000, col_w - 320000, 620000, number,
                size=44, font=FONT_BOLD, color=accent, line_spacing=1.0)
        textbox(slide, x + 160000, y + 940000, col_w - 320000, 620000,
                caption, size=9.5, color=GRAY, line_spacing=1.3)

    divider(slide, MARGIN_L, y + card_h + 200000, CONTENT_W)
    textbox(
        slide, MARGIN_L, y + card_h + 280000, CONTENT_W, 300000,
        "Источник: реестр из 62 процессов, подпроцессов и этапов; "
        "оценка по 5 критериям на основании Методики KT_METHOD v1.",
        size=8.5, color=STEEL,
    )


def build_criteria(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Пять критериев оценки",
          "Балл по каждому от 1 до 5, умножается на вес, приводится к шкале 0–100")

    rows = [("Критерий", "Вес", "На какой вопрос отвечает")]
    rows += [(name, str(weight), question) for name, weight, question, *_ in P.CRITERIA]
    end_y = table_grid(
        slide, MARGIN_L, CONTENT_TOP + 260000, CONTENT_W,
        [30, 10, 60], rows, row_h=370000, size=10,
    )

    textbox(
        slide, MARGIN_L, end_y + 260000, CONTENT_W, 420000,
        ["Вес отражает управленческий приоритет: то, что даёт результат, стоит дороже того, "
         "что даёт удобство.",
         "Объекта на 100 баллов в реестре нет — шкала нужна для сравнения объектов между "
         "собой, а не для абсолютной оценки."],
        size=9, color=GRAY, line_spacing=1.3, space_after=4,
    )


def build_rules(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Почему очередь не совпадает с рейтингом",
          "Пять правил очерёдности, каждое опирается на пункт Методики")

    y = CONTENT_TOP + 180000
    row_h = 660000
    for i, (name, _, short) in enumerate(P.ORDER_RULES, 1):
        numbered_row(slide, MARGIN_L, y, CONTENT_W, row_h - 80000, i, name, short,
                     body_size=9.5)
        if i < len(P.ORDER_RULES):
            divider(slide, MARGIN_L, y + row_h - 90000, CONTENT_W)
        y += row_h


def build_waves(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Четыре волны работ",
          "Волна задаёт зависимости, балл — порядок внутри волны")

    waves = [w for w in P.WAVES if w[0] != P.ВН]
    col_w = (CONTENT_W - 3 * 140000) // 4
    y = CONTENT_TOP + 120000
    for i, (wave, phase, goal, *_rest) in enumerate(waves):
        accent = WAVE_COLORS[wave]
        x = MARGIN_L + i * (col_w + 140000)

        head = rect(slide, x, y, col_w, 420000, fill=accent, radius=0.10)
        tf = head.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_top = tf.margin_bottom = 0
        set_text(tf, f"Волна {wave_no(wave)}", size=12, font=FONT_BOLD,
                 color=on(accent), align=PP_ALIGN.CENTER)

        body = rect(slide, x, y + 440000, col_w, 1180000, fill=CARD_BG)
        tf = body.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        set_text(tf, wave_name(wave), size=10.5, font=FONT_BOLD, color=ink(accent),
                 line_spacing=1.2)
        para = tf.add_paragraph()
        para.space_before = Pt(8)
        para.line_spacing = 1.25
        run = para.add_run()
        run.text = goal
        run.font.name = FONT
        run.font.size = Pt(9)
        run.font.color.rgb = TEXT

        foot = rect(slide, x, y + 1680000, col_w, 360000, fill=WHITE, line=MIST)
        tf = foot.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_top = tf.margin_bottom = 0
        set_text(tf, f"{BY_WAVE[wave]} объектов  ·  {phase.replace(' дорожной карты', '')}",
                 size=9, font=FONT_MED, color=GRAY, align=PP_ALIGN.CENTER)

    textbox(
        slide, MARGIN_L, y + 2120000, CONTENT_W, 250000,
        f"Отдельно: {BY_WAVE[P.ВН]} объекта вне очереди — исполняются другим подразделением "
        f"целиком, срок контролируется через SLA по родительскому объекту.",
        size=8.5, color=STEEL,
    )

    # Размер первой волны больше числа «ничьих» зон, и это первое, о чём спрашивают.
    breakdown_y = y + 2450000
    divider(slide, MARGIN_L, breakdown_y - 30000, CONTENT_W)
    textbox(
        slide, MARGIN_L, breakdown_y, CONTENT_W, 250000,
        f"Из чего состоит волна 1: {BY_WAVE[P.В1]} объектов",
        size=10, font=FONT_MED, color=DARK,
    )

    groups = wave_1_breakdown()
    seg_w = (CONTENT_W - 2 * 140000) // 3
    seg_y = breakdown_y + 290000
    for i, (count, label, note) in enumerate(groups):
        accent = RED if i == 0 else BLUE
        x = MARGIN_L + i * (seg_w + 140000)
        seg = rect(slide, x, seg_y, seg_w, 810000, fill=CARD_BG)
        tf = seg.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_top = tf.margin_bottom = Emu(80000)

        p = tf.paragraphs[0]
        p.line_spacing = 1.0
        run = p.add_run()
        run.text = str(count)
        run.font.name = FONT_BOLD
        run.font.size = Pt(16)
        run.font.color.rgb = accent
        run = p.add_run()
        run.text = f"   {label}"
        run.font.name = FONT_BOLD
        run.font.size = Pt(10)
        run.font.color.rgb = accent

        para = tf.add_paragraph()
        para.space_before = Pt(5)
        para.line_spacing = 1.2
        run = para.add_run()
        run.text = note
        run.font.name = FONT
        run.font.size = Pt(8.5)
        run.font.color.rgb = TEXT


def build_matrix(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Ценность и лёгкость реализации",
          "Ценность — четыре критерия без учёта лёгкости; граница высокой ценности — 60 из 100")

    quad_accent = {
        P.QUADRANTS[0][0]: GREEN,
        P.QUADRANTS[1][0]: BLUE,
        P.QUADRANTS[2][0]: BLUE_LIGHT,
        P.QUADRANTS[3][0]: GRAY,
    }
    layout = [
        (["Ценность", "высокая"], P.QUADRANTS[0], P.QUADRANTS[1]),
        (["Ценность", "умеренная"], P.QUADRANTS[2], P.QUADRANTS[3]),
    ]

    label_w = 940000
    grid_x = MARGIN_L + label_w + 140000
    grid_w = CONTENT_W - label_w - 140000
    col_w = (grid_w - 160000) // 2
    head_h = 320000
    cell_h = 1330000
    y0 = CONTENT_TOP + 140000

    for j, head in enumerate(["Реализация лёгкая (4–5 из 5)", "Реализация тяжёлая (1–3 из 5)"]):
        x = grid_x + j * (col_w + 160000)
        textbox(slide, x, y0, col_w, head_h, head, size=10, font=FONT_MED,
                color=GRAY, align=PP_ALIGN.CENTER)

    for i, (label, left, right) in enumerate(layout):
        y = y0 + head_h + i * (cell_h + 160000)

        band = rect(slide, MARGIN_L, y, label_w, cell_h, fill=BLUE, radius=0.10)
        tf = band.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Emu(40000)
        set_text(tf, label, size=10, font=FONT_MED,
                 color=WHITE, align=PP_ALIGN.CENTER, line_spacing=1.25)

        for j, quad in enumerate((left, right)):
            name, _definition, decision = quad
            accent = quad_accent[name]
            x = grid_x + j * (col_w + 160000)
            cell = rect(slide, x, y, col_w, cell_h, fill=CARD_BG)
            tf = cell.text_frame
            tf.vertical_anchor = MSO_ANCHOR.TOP

            p = tf.paragraphs[0]
            p.line_spacing = 1.0
            run = p.add_run()
            run.text = str(BY_QUADRANT[name])
            run.font.name = FONT_BOLD
            run.font.size = Pt(24)
            run.font.color.rgb = ink(accent)
            run = p.add_run()
            run.text = f"   {name}"
            run.font.name = FONT_BOLD
            run.font.size = Pt(11.5)
            run.font.color.rgb = ink(accent)

            para = tf.add_paragraph()
            para.space_before = Pt(8)
            para.line_spacing = 1.25
            run = para.add_run()
            run.text = decision
            run.font.name = FONT
            run.font.size = Pt(9)
            run.font.color.rgb = TEXT


def build_queue(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Начало очереди: первые десять объектов",
          "Волна 1 целиком — вопрос ответственности и базовая линия")

    rows = [("№", "Код", "Объект реестра", "Балл", "Что делаем")]
    for r in ROWS[:10]:
        rows.append((
            str(r["rank"]), r["code"], clip(r["name"], 52),
            f'{r["total"]:.0f}', r["kind"],
        ))
    table_grid(
        slide, MARGIN_L, CONTENT_TOP + 160000, CONTENT_W,
        [6, 10, 48, 8, 28], rows, row_h=282000, size=9,
    )


def build_top_value(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Самые ценные объекты идут не первыми",
          "И это осознанное решение, а не недосмотр")

    top = sorted(SCORED, key=lambda r: (-r["total"], r["code"]))[:6]
    rows = [("Балл", "Код", "Объект реестра", "Когда берём")]
    for r in top:
        rows.append((
            f'{r["total"]:.0f}', r["code"], clip(r["name"], 46),
            f'Волна {wave_no(r["wave"])}. {wave_name(r["wave"])}',
        ))
    end_y = table_grid(
        slide, MARGIN_L, CONTENT_TOP + 160000, CONTENT_W,
        [8, 10, 48, 34], rows, row_h=300000, size=9.5,
    )

    waves_used = sorted({wave_no(r["wave"]) for r in top})
    accent_bar(
        slide, MARGIN_L, end_y + 300000, CONTENT_W, 620000,
        ["Это вход проекта: приём пакета от ДРБ, приоритизация платежей, закупки "
         "и готовность продукта.",
         f"Они ждут волн {' и '.join(waves_used)}: сначала владельцы и замер, "
         "иначе эффект будет не доказан, а объявлен."],
        fill=GREEN, size=10, align=PP_ALIGN.LEFT,
    )


def build_first_steps(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Программа старта: пять первых действий",
          "Шаги 1–5 закрывают ответственность и снимают базовую линию")

    y = CONTENT_TOP + 140000
    row_h = 620000
    owner_w = 2200000
    owner_x = MARGIN_L + CONTENT_W - owner_w
    for i, step in enumerate(P.FIRST_STEPS[:5], 1):
        action, objects, owner, phase, result, _why = step
        numbered_row(
            slide, MARGIN_L, y, CONTENT_W - owner_w - 200000, row_h - 80000, i,
            action, f"Результат: {result}", accent=BLUE, heading_size=10,
        )
        textbox(slide, owner_x, y - 10000, owner_w, 330000, owner,
                size=9, font=FONT_MED, color=ink(GREEN), line_spacing=1.2)
        textbox(slide, owner_x, y + 340000, owner_w, 220000,
                f"{phase}  ·  {objects}", size=8, color=STEEL)
        if i < 5:
            divider(slide, MARGIN_L, y + row_h - 90000, CONTENT_W)
        y += row_h

    textbox(
        slide, MARGIN_L, y + 40000, CONTENT_W, 280000,
        "Шаги 6–10 — восстановление исполнения внутри ДУП: приёмка пакета от ДРБ, "
        "платежи, статус-встречи, прогноз, переносы оборудования.",
        size=8.5, color=STEEL,
    )


def build_asks(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Что требуется от руководства",
          "Решения, без которых очередь не двигается")

    asks = [step for step in P.FIRST_STEPS if "директор дуп" in step[2].lower()]
    y = CONTENT_TOP + 120000
    row_h = 560000
    for i, step in enumerate(asks, 1):
        action, _objects, _owner, _phase, result, _why = step
        numbered_row(slide, MARGIN_L, y, CONTENT_W, row_h - 80000, i,
                     clip(action, 76), f"Закрывается документом: {result}", accent=GREEN)
        if i < len(asks):
            divider(slide, MARGIN_L, y + row_h - 90000, CONTENT_W)
        y += row_h

    accent_bar(
        slide, MARGIN_L, y + 40000, CONTENT_W, 420000,
        "Ни одно из решений не требует бюджета, ИТ-разработки и согласия смежных блоков.",
        fill=BLUE, size=10.5, align=PP_ALIGN.LEFT,
    )


def build_closing(prs: Presentation) -> None:
    slide = prs.slides[-1]
    msg = find_shape(slide, "Присоединяйтесь")
    if msg is not None:
        replace_text(
            msg,
            "Начинаем с ответственности и замера — дальше очередь считается сама",
            size=16, font=FONT_MED, color=WHITE, align=PP_ALIGN.CENTER,
        )


# --------------------------------------------------------------------------
def main() -> None:
    if not TEMPLATE.exists():
        raise SystemExit(f"Не найден шаблон брендбука: {TEMPLATE}")

    prs = Presentation(str(TEMPLATE))
    for idx in range(10, 0, -1):
        drop_slide(prs, idx)

    reserve_partname(prs.slides[1], 90)
    build_cover(prs)

    for builder in (
        build_headline,
        build_criteria,
        build_rules,
        build_waves,
        build_matrix,
        build_queue,
        build_top_value,
        build_first_steps,
        build_asks,
    ):
        builder(prs)

    move_slide_to_end(prs, 1)
    build_closing(prs)

    prs.save(str(OUTPUT))
    print(f"Сохранено: {OUTPUT}")
    print(f"Слайдов: {len(prs.slides._sldIdLst)}")


if __name__ == "__main__":
    main()

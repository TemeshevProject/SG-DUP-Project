#!/usr/bin/env python3
"""Генерация презентации проекта трансформации ДУП в фирменном стиле Sergek.

Использует корпоративный шаблон «Общие слайды_ver 12.12.24.pptx» как основу:
из него наследуются мастера, макеты, логотипы, футер и нумерация страниц.
"""

from __future__ import annotations

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
    CONTENT_BOTTOM,
    CONTENT_TOP,
    CONTENT_W,
    DARK,
    FONT,
    FONT_BOLD,
    FONT_MED,
    GRAY,
    GREEN,
    LAYOUT_BLACK,
    LAYOUT_TITLE,
    MARGIN_L,
    MIST,
    RED,
    SLIDE_W,
    STEEL,
    TEAL_DARK,
    TEXT,
    WHITE,
    YELLOW,
    accent_bar,
    bullets,
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
TEMPLATE = BASE.parent / "brandbook" / "Общие слайды_ver 12.12.24.pptx"
OUTPUT = BASE / "ДУП_Трансформация_процессов_презентация.pptx"

DECK_TITLE = "Трансформация процессов ДУП"
DECK_SUBTITLE = "Дорожная карта проекта «под ключ»"
DECK_DATE = "Сентябрь 2026"

RELS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


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
    """Сдвигает имя части слайда, чтобы новые слайды не заняли тот же путь.

    После удаления слайдов шаблона python-pptx нумерует новые части с первого
    свободного индекса и может перезаписать сохранённый слайд.
    """
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


def new_slide(prs: Presentation, layout_index: int = LAYOUT_TITLE):
    return prs.slides.add_slide(prs.slide_layouts[layout_index])


# --------------------------------------------------------------------------
# Слайды
# --------------------------------------------------------------------------
def build_cover(prs: Presentation) -> None:
    """Титульный слайд шаблона (синий фон): подставляем название и дату."""
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
    textbox(
        slide,
        text_x + 60000,
        2600000,
        4100000,
        300000,
        DECK_SUBTITLE,
        size=12.5,
        font=FONT_MED,
        color=GREEN,
    )
    textbox(
        slide,
        text_x + 60000,
        2990000,
        4100000,
        280000,
        "Департамент управления проектами · ТОО «Көркем Телеком»",
        size=9.5,
        color=MIST,
    )


def build_agenda(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Содержание")

    items = [
        ("01", "Итоги диагностики", "Что проанализировали и главный вывод"),
        ("02", "Карта процессов и диагноз", "4 блока, 15 узких мест, 3 разрыва"),
        ("03", "Направления работы", "Оптимизация · Трансформация · Цифровизация"),
        ("04", "Дорожная карта", "5 фаз, 12 месяцев, gate-критерии"),
        ("05", "Управление и результат", "KPI, governance, риски, первые шаги"),
    ]
    y = CONTENT_TOP + 60000
    row_h = 620000
    gap = 60000
    for num, name, desc in items:
        badge = rect(slide, MARGIN_L, y, 560000, 540000, fill=BLUE, radius=0.14)
        tf = badge.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_top = tf.margin_bottom = 0
        set_text(tf, num, size=15, font=FONT_BOLD, color=WHITE, align=PP_ALIGN.CENTER)

        textbox(slide, MARGIN_L + 720000, y + 60000, 5000000, 260000, name, size=13, font=FONT_BOLD, color=DARK)
        textbox(slide, MARGIN_L + 720000, y + 320000, 6200000, 240000, desc, size=9.5, color=GRAY)
        divider(slide, MARGIN_L, y + row_h - 10000, CONTENT_W)
        y += row_h + gap


def build_executive_summary(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Главный вывод диагностики")

    banner = rect(slide, MARGIN_L, CONTENT_TOP - 40000, CONTENT_W, 700000, fill=TEAL_DARK, radius=0.1)
    tf = banner.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Emu(240000)
    set_text(
        tf,
        [
            "Проблема ДУП — не в отсутствии процессов, а в их исполнении",
            "и организационной модели. Методология уже описывает почти всё.",
        ],
        size=13,
        font=FONT_BOLD,
        color=WHITE,
        line_spacing=1.25,
    )

    y = CONTENT_TOP + 830000
    col_w = (CONTENT_W - 2 * 180000) // 3
    tiles = [
        ("12 из 15", ["узких мест имеют описанную", "процедуру, но она не исполняется"], GREEN),
        ("2", ["подлинных пробела — требуют", "нового регламента"], YELLOW),
        ("1", ["новый риск — фрагментированный", "ИТ-ландшафт"], RED),
    ]
    for i, (num, cap, accent) in enumerate(tiles):
        x = MARGIN_L + i * (col_w + 180000)
        box = rect(slide, x, y, col_w, 1330000)
        tfb = box.text_frame
        tfb.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tfb.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.0
        r = p.add_run()
        r.text = num
        r.font.name = FONT_BOLD
        r.font.size = Pt(30)
        r.font.color.rgb = ink(accent)
        bullets(box, cap, size=9.5, spacing=4, line_spacing=1.25)

    y2 = y + 1500000
    concl = rect(slide, MARGIN_L, y2, CONTENT_W, 830000, fill=CARD_BG)
    tf2 = concl.text_frame
    tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf2.margin_left = Emu(240000)
    tf2.margin_right = Emu(240000)
    p = tf2.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = 1.15
    r = p.add_run()
    r.text = "Следствие для стратегии проекта"
    r.font.name = FONT_BOLD
    r.font.size = Pt(11)
    r.font.color.rgb = BLUE
    bullets(
        concl,
        [
            "Не переписываем методологию заново — восстанавливаем исполнение, закрываем два пробела "
            "и меняем организационную модель. Это кратно быстрее и дешевле разработки процессов с нуля.",
        ],
        size=10,
        spacing=6,
        line_spacing=1.3,
    )


def build_sources(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Что проанализировали", "Два слоя документации — текущий анализ и наследие методологии")

    y = CONTENT_TOP + 60000
    col_w = (CONTENT_W - 200000) // 2
    card(
        slide,
        MARGIN_L,
        y,
        col_w,
        1730000,
        "Слой А — текущий анализ (2026)",
        [
            "Интервью с руководством ДУП — 15 узких мест",
            "Дерево процессов v4 — 4 блока + внешние",
            "RACI-матрица — 25 подпроцессов",
            "Кадровая и компетентностная матрица",
            "4 проекта должностных инструкций",
        ],
        accent=GREEN,
    )
    card(
        slide,
        MARGIN_L + col_w + 200000,
        y,
        col_w,
        1730000,
        "Слой B — методология «Сергек» (2020–2024)",
        [
            "1 313 страниц, приказ Правления от 15.01.2021",
            "417 страниц Confluence + 13 локальных файлов",
            "5 стадий жизненного цикла × 5 треков",
            "Чек-листы аудита и 9 ролевых инструкций",
            "Навигатор по 21 роли",
        ],
        accent=BLUE,
    )

    y2 = y + 1900000
    textbox(slide, MARGIN_L, y2, CONTENT_W, 250000, "Что это дало", size=11, font=FONT_BOLD, color=DARK)
    rows = [
        ["Артефакт", "Объём", "Роль в проекте"],
        ["Реестр документации", "434 записи", "Инвентаризация ВНД, база для актуализации"],
        ["Реестр бизнес-процессов", "51 запись", "Процессы, подпроцессы, этапы и владельцы"],
        ["Опросник директоров филиалов", "46 вопросов / 8 блоков", "Масштабирование диагностики на сеть"],
    ]
    table_grid(slide, MARGIN_L, y2 + 280000, CONTENT_W, [24, 18, 58], rows, row_h=300000, size=9)


def build_process_map(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Карта процессов ДУП", "23 подпроцесса в четырёх блоках + интерфейсы с центральным офисом")

    y = CONTENT_TOP + 120000
    gap = 130000
    col_w = (CONTENT_W - 3 * gap) // 4
    blocks = [
        (
            "1. Реализация проекта",
            BLUE,
            "9 подпроцессов",
            ["Приём пакета от ДРБ", "Проектная документация", "ФЭМ и бюджет", "Закупки", "Реестр платежей"],
        ),
        (
            "2. Сопровождение",
            GREEN,
            "8 подпроцессов",
            ["Поверка оборудования", "Канал связи", "Прогноз поступлений", "Переносы", "Отчётность"],
        ),
        (
            "3. Филиальная сеть",
            BLUE_LIGHT,
            "3 подпроцесса",
            ["Назначение куратора", "Открытие филиала", "Отношения с филиалом"],
        ),
        (
            "4. Методология / ПМО",
            YELLOW,
            "3 подпроцесса",
            ["Методология и шаблоны", "Контроль качества данных", "Портфельная отчётность"],
        ),
    ]
    for i, (name, accent, count, items) in enumerate(blocks):
        x = MARGIN_L + i * (col_w + gap)
        accent_bar(slide, x, y, col_w, 480000, name, fill=accent, color=on(accent), size=10.5)
        body = rect(slide, x, y + 500000, col_w, 1750000)
        tf = body.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = count
        r.font.name = FONT_BOLD
        r.font.size = Pt(9)
        r.font.color.rgb = ink(accent)
        bullets(body, items, size=8.5, prefix="· ", spacing=6, line_spacing=1.2)

    y2 = y + 2420000
    ext = rect(slide, MARGIN_L, y2, CONTENT_W, 700000, fill=CARD_BG)
    tf = ext.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Emu(200000)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = 1.25
    r = p.add_run()
    r.text = "Раздел Б — процессы других подразделений, где ДУП является участником:  "
    r.font.name = FONT_BOLD
    r.font.size = Pt(9.5)
    r.font.color.rgb = DARK
    r2 = p.add_run()
    r2.text = "ДРБ (передача проекта) · Финблок (ФЭМ, платежи, закрытие) · Закупки · Юр блок · СТ (монтаж)"
    r2.font.name = FONT
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = TEXT


def build_diagnosis(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Диагноз: 15 узких мест", "Ревизия по методологии изменила картину проблемы")

    y = CONTENT_TOP + 120000
    left_w = 2500000
    box = rect(slide, MARGIN_L, y, left_w, 1450000, fill=TEAL_DARK)
    tf = box.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.line_spacing = 1.0
    r = p.add_run()
    r.text = "15"
    r.font.name = FONT_BOLD
    r.font.size = Pt(44)
    r.font.color.rgb = WHITE
    for line in ["узких мест", "выявлено в диагностике"]:
        para = tf.add_paragraph()
        para.alignment = PP_ALIGN.CENTER
        para.line_spacing = 1.2
        run = para.add_run()
        run.text = line
        run.font.name = FONT
        run.font.size = Pt(9.5)
        run.font.color.rgb = MIST

    rows = [
        (GREEN, "12", "Норматив есть — не исполняется", "Приём пакета, прогнозы, закупки, статус-встречи, аудит, ФЭМ, переносы, дефекты"),
        (YELLOW, "2", "Подлинный пробел", "Подтверждение готовности продукта · Передача «реализация → сопровождение»"),
        (RED, "1", "Новый риск", "Фрагментированный ИТ-ландшафт и незавершённая миграция платформы"),
    ]
    x = MARGIN_L + left_w + 200000
    w = CONTENT_W - left_w - 200000
    row_h = 440000
    ry = y
    for accent, num, head, desc in rows:
        rect(slide, x, ry, 22000, row_h - 40000, fill=accent, shape=MSO_SHAPE.RECTANGLE)
        textbox(slide, x + 120000, ry + 10000, 400000, 300000, num, size=18, font=FONT_BOLD, color=ink(accent))
        textbox(slide, x + 560000, ry + 20000, w - 600000, 200000, head, size=10.5, font=FONT_BOLD, color=DARK)
        textbox(slide, x + 560000, ry + 210000, w - 600000, 220000, desc, size=8.5, color=GRAY, line_spacing=1.15)
        ry += row_h + 50000

    y2 = y + 1620000
    textbox(
        slide,
        MARGIN_L,
        y2,
        CONTENT_W,
        250000,
        "Как изменилась оценка после сверки с методологией",
        size=11,
        font=FONT_BOLD,
        color=DARK,
    )
    rows2 = [
        ["Классификация", "До сверки", "После сверки", "Что это меняет"],
        ["Покрыто нормативно", "3", "12", "Фокус смещается на исполнение и контроль"],
        ["Подлинный пробел", "5", "2", "Нужен только 1 новый регламент + чек-лист передачи"],
        ["Новый риск", "2", "1", "ИТ-миграция выносится в отдельный поток"],
    ]
    table_grid(slide, MARGIN_L, y2 + 280000, CONTENT_W, [26, 14, 14, 46], rows2, row_h=290000, size=9)


def build_gaps(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Три ключевых разрыва", "Что именно мешает управляемости портфеля сегодня")

    y = CONTENT_TOP + 130000
    gap = 190000
    col_w = (CONTENT_W - 2 * gap) // 3
    items = [
        (
            "Организационный",
            RED,
            [
                "Блоки «Реализация» и «Сопровождение» ведёт один человек",
                "Роль ПМО не выделена, методология числится за директором",
                "Открытие филиала — «ничья зона» в RACI и ДИ",
                "Отчётность дублируется в трёх ролях",
            ],
        ),
        (
            "Исполнительский",
            YELLOW,
            [
                "12 процедур описаны, но не выполняются",
                "Из 6 обязанностей менеджера ЦА исполняется около одной",
                "Контур аудита остановлен после 2020 года",
                "Прогнозы и платежи ведутся в чатах и ad-hoc запросах",
            ],
        ),
        (
            "Цифровой",
            BLUE,
            [
                "Платформа портфеля работает не полностью",
                "Нет единой карточки проекта и блокировки без пакета",
                "Данные распределены между Confluence, таблицами, мессенджерами",
                "Отсутствуют автоуведомления по срокам и платежам",
            ],
        ),
    ]
    for i, (name, accent, lines) in enumerate(items):
        x = MARGIN_L + i * (col_w + gap)
        accent_bar(slide, x, y, col_w, 440000, name, fill=accent, color=on(accent), size=11)
        body = rect(slide, x, y + 460000, col_w, 2450000)
        body.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        bullets(body, lines, size=9, prefix="— ", spacing=11, line_spacing=1.3)


def build_directions(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Три направления работы", "19 процессов и инициатив в портфеле трансформации")

    y = CONTENT_TOP + 140000
    gap = 190000
    col_w = (CONTENT_W - 2 * gap) // 3
    blocks = [
        ("Оптимизация", GREEN, "10", "процессов", "Восстановление исполнения существующих процедур методологии"),
        ("Трансформация", BLUE, "5", "процессов", "Новое проектирование: пробелы и организационная модель"),
        ("Цифровизация", BLUE_LIGHT, "4", "инициативы", "Единая экосистема данных и автоматизация рутины"),
    ]
    for i, (name, accent, num, unit, desc) in enumerate(blocks):
        x = MARGIN_L + i * (col_w + gap)
        box = rect(slide, x, y, col_w, 1620000)
        tf = box.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.0
        r = p.add_run()
        r.text = name
        r.font.name = FONT_BOLD
        r.font.size = Pt(13)
        r.font.color.rgb = ink(accent)

        pn = tf.add_paragraph()
        pn.alignment = PP_ALIGN.LEFT
        pn.space_before = Pt(10)
        pn.line_spacing = 1.0
        rn = pn.add_run()
        rn.text = num + "  "
        rn.font.name = FONT_BOLD
        rn.font.size = Pt(26)
        rn.font.color.rgb = DARK
        ru = pn.add_run()
        ru.text = unit
        ru.font.name = FONT
        ru.font.size = Pt(10)
        ru.font.color.rgb = GRAY

        bullets(box, [desc], size=9, spacing=12, line_spacing=1.3)

    y2 = y + 1800000
    textbox(slide, MARGIN_L, y2, CONTENT_W, 250000, "Логика приоритизации", size=11, font=FONT_BOLD, color=DARK)
    seq = [
        ("P1", "Быстрый эффект без ИТ", "Восстановление процедур — результат в пределах фазы 2"),
        ("P2", "Организационная готовность", "Разделение ролей и новые регламенты — фаза 3"),
        ("P3", "Системное закрепление", "Цифровизация и актуализация ВНД — фаза 4"),
    ]
    sy = y2 + 330000
    for code, head, desc in seq:
        badge = rect(slide, MARGIN_L, sy, 380000, 310000, fill=BLUE, radius=0.18)
        tfb = badge.text_frame
        tfb.vertical_anchor = MSO_ANCHOR.MIDDLE
        tfb.margin_left = tfb.margin_right = 0
        set_text(tfb, code, size=9.5, font=FONT_BOLD, color=WHITE, align=PP_ALIGN.CENTER)
        textbox(slide, MARGIN_L + 500000, sy + 30000, 2700000, 250000, head, size=10, font=FONT_BOLD, color=DARK)
        textbox(slide, MARGIN_L + 3300000, sy + 35000, CONTENT_W - 3300000, 250000, desc, size=9, color=GRAY)
        sy += 400000


def build_optimization(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Что оптимизируем", "Процессы, где норматив есть — восстанавливаем исполнение")

    rows = [
        ["Код", "Процесс", "Что делаем", "Приоритет"],
        ["1.1", "Приём пакета от ДРБ", "Чек-лист, SLA 5 дней, блокировка старта без пакета", "P1"],
        ["1.5", "ФЭМ и бюджет проекта", "Актуализация нормативов: поверка, переносы, цены", "P1"],
        ["1.7", "Закупки", "Нормативы сроков по шагам и трекер эскалации", "P1"],
        ["1.8", "Статус-встречи", "Регламент: частота, повестка, протокол", "P1"],
        ["1.9", "Реестр платежей", "Единая инструкция трёх потоков (ВУ / КИЗ / КОЗ)", "P1"],
        ["2.3", "Прогноз поступлений и расходов", "Единый реестр и календарь вместо чата", "P1"],
        ["2.5", "Переносы оборудования", "Учёт «факт против ФЭМ», процедура согласования", "P2"],
        ["2.7", "Обратная связь по дефектам", "Регулярный сбор и маршрут в ТЗ и закупки", "P1"],
        ["4.2", "Аудит проектных данных", "Восстановление контура контроля (остановлен с 2020)", "P1"],
        ["2.1", "Поверка оборудования", "Оценка собственной аккредитации, ~6 000 комплексов", "P2"],
    ]
    table_grid(slide, MARGIN_L, CONTENT_TOP + 60000, CONTENT_W, [8, 30, 52, 10], rows, row_h=310000, size=9)


def build_transformation(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Что трансформируем", "Новое проектирование — пробелы и организационная модель")

    y = CONTENT_TOP + 100000
    gap = 160000
    col_w = (CONTENT_W - 2 * gap) // 3
    top_cards = [
        ("1.2 · Готовность продукта", "Новый регламент gate «пилот → реализация»: критерии, роли ДРБ и ДУП, подпись, блокировка старта", YELLOW),
        ("2.0 · Передача 1 → 2", "Чек-лист handover «реализация → сопровождение»: состав, ответственные, дата, архив", YELLOW),
        ("3.2 · Открытие филиала", "Назначение владельца и сквозной маршрут: Юр → HR → Финблок → ДУП → филиал", YELLOW),
    ]
    for i, (head, body, accent) in enumerate(top_cards):
        x = MARGIN_L + i * (col_w + gap)
        card(slide, x, y, col_w, 1350000, head, [body], accent=ink(accent), heading_size=10.5, body_size=9)

    y2 = y + 1520000
    bottom = [
        ("Разделение блоков 1 и 2", "Два заместителя вместо одной совмещённой роли — после утверждения чек-листа передачи", BLUE),
        ("Роль ПМО (главный менеджер проектов)", "Методология, контроль качества данных и портфельная отчётность в одних руках", BLUE),
        ("Разведение отчётности 1.8 / 2.8 / 4.3", "Матрица «кто, что, кому и когда» — устранение дублирования в трёх ролях", BLUE),
    ]
    for i, (head, body, accent) in enumerate(bottom):
        x = MARGIN_L + i * (col_w + gap)
        card(slide, x, y2, col_w, 1350000, head, [body], accent=ink(accent), heading_size=10.5, body_size=9)

    textbox(
        slide,
        MARGIN_L,
        y2 + 1500000,
        CONTENT_W,
        240000,
        "Критическая последовательность: чек-лист передачи утверждается до разделения заместителей — "
        "иначе разрыв ответственности переносится в новую структуру.",
        size=9,
        font=FONT_MED,
        color=RED,
    )


def build_digital(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Что цифровизуем", "Четыре инициативы единой экосистемы данных")

    y = CONTENT_TOP + 120000
    gap = 150000
    col_w = (CONTENT_W - 3 * gap) // 4
    items = [
        ("Единая карточка проекта", ["Обязательные поля", "Блокировка передачи без пакета", "Процессы 1.1, 1.3, 1.8"]),
        ("Автопрогноз", ["Шаблон и календарь сбора", "Напоминания и дашборд", "Процесс 2.3"]),
        ("Автоуведомления", ["Сроки контрактов", "Платежи и просрочки", "Процессы 1.7, 1.9, 2.4"]),
        ("Интеграции", ["Закупки — 1С — портфель", "Сохранение Zabbix и Documentolog", "Процессы 1.5, 1.7"]),
    ]
    for i, (head, lines) in enumerate(items):
        x = MARGIN_L + i * (col_w + gap)
        accent_bar(slide, x, y, col_w, 460000, head, fill=BLUE_LIGHT, size=10)
        body = rect(slide, x, y + 480000, col_w, 1620000)
        body.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        bullets(body, lines, size=9, prefix="· ", spacing=11, line_spacing=1.3)

    y2 = y + 2290000
    note = rect(slide, MARGIN_L, y2, CONTENT_W, 850000, fill=CARD_BG)
    tf = note.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Emu(200000)
    tf.margin_right = Emu(200000)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = 1.3
    r = p.add_run()
    r.text = "Принцип: "
    r.font.name = FONT_BOLD
    r.font.size = Pt(10)
    r.font.color.rgb = BLUE
    r2 = p.add_run()
    r2.text = (
        "цифровизация запускается после восстановления процедур. Автоматизировать неработающий процесс — "
        "значит закрепить проблему в системе. Все ИТ-инициативы стартуют в фазе 4 на подготовленной основе."
    )
    r2.font.name = FONT
    r2.font.size = Pt(10)
    r2.font.color.rgb = TEXT


def build_roadmap(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Дорожная карта", "5 фаз · 12 месяцев · переход между фазами по gate-критериям")

    phases = [
        ("Фаза 0", "Диагностика", "Месяц 1", TEAL_DARK, ["Аудит 1 филиала", "Опросник филиалов", "Интервью с ЦО"]),
        ("Фаза 1", "Фундамент", "Месяцы 2–3", BLUE, ["RACI v2", "Целевая оргсхема", "Пакет ДРБ → ДУП"]),
        ("Фаза 2", "Исполнение", "Месяцы 3–6", GREEN, ["12 процедур", "Реестр прогнозов", "Аудит на 3 филиалах"]),
        ("Фаза 3", "TO-BE", "Месяцы 5–8", YELLOW, ["3 регламента", "Разделение замов", "Матрица отчётности"]),
        ("Фаза 4", "Цифра", "Месяцы 7–12", BLUE_LIGHT, ["Карточка проекта", "Автопрогноз", "ВНД v2"]),
    ]
    y = CONTENT_TOP + 120000
    gap = 100000
    col_w = (CONTENT_W - 4 * gap) // 5
    for i, (code, name, period, accent, items) in enumerate(phases):
        x = MARGIN_L + i * (col_w + gap)
        head = rect(slide, x, y, col_w, 620000, fill=accent, radius=0.1)
        tf = head.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Emu(60000)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.line_spacing = 1.05
        r = p.add_run()
        r.text = code
        r.font.name = FONT_BOLD
        r.font.size = Pt(11)
        r.font.color.rgb = on(accent)
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.line_spacing = 1.05
        r2 = p2.add_run()
        r2.text = name
        r2.font.name = FONT
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = on(accent)

        textbox(
            slide,
            x,
            y + 660000,
            col_w,
            220000,
            period,
            size=8.5,
            font=FONT_MED,
            color=ink(accent),
            align=PP_ALIGN.CENTER,
        )
        body = rect(slide, x, y + 900000, col_w, 1450000)
        body.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        bullets(body, items, size=8.5, prefix="· ", spacing=8, line_spacing=1.25)

    y2 = y + 2480000
    textbox(slide, MARGIN_L, y2, CONTENT_W, 240000, "Критический путь", size=11, font=FONT_BOLD, color=DARK)
    textbox(
        slide,
        MARGIN_L,
        y2 + 270000,
        CONTENT_W,
        400000,
        "Диагностика → RACI и оргмодель → восстановление 12 процедур → чек-лист передачи и gate продукта → "
        "цифровизация. Фазы 2 и 3 частично идут параллельно; фаза 4 стартует только на подготовленной основе.",
        size=9.5,
        color=TEXT,
        line_spacing=1.3,
    )


def build_phase_detail_01(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Фазы 0–1: диагностика и фундамент", "Месяцы 1–3 · подготовка к изменениям")

    y = CONTENT_TOP + 100000
    col_w = (CONTENT_W - 200000) // 2

    accent_bar(slide, MARGIN_L, y, col_w, 420000, "Фаза 0 · Диагностика (месяц 1)", fill=TEAL_DARK, size=10.5, align=PP_ALIGN.LEFT)
    body = rect(slide, MARGIN_L, y + 440000, col_w, 2380000)
    tf = body.text_frame
    f0 = [
        ("Аудит пилотного филиала", "по трём чек-листам методологии"),
        ("Опросник директоров филиалов", "46 вопросов, форма готова к запуску"),
        ("Интервью с ДРБ", "передача пакета и готовность продукта"),
        ("Интервью с Финблоком", "ФЭМ, прогнозы, приоритизация платежей"),
        ("Интервью с Юр блоком", "договоры и открытие филиала"),
        ("Статус методологии", "проверка приказа от 15.01.2021"),
    ]
    for j, (head, desc) in enumerate(f0):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.2
        p.space_before = Pt(0 if j == 0 else 9)
        r = p.add_run()
        r.text = head + " — "
        r.font.name = FONT_BOLD
        r.font.size = Pt(9)
        r.font.color.rgb = DARK
        r2 = p.add_run()
        r2.text = desc
        r2.font.name = FONT
        r2.font.size = Pt(9)
        r2.font.color.rgb = GRAY

    x2 = MARGIN_L + col_w + 200000
    accent_bar(slide, x2, y, col_w, 420000, "Фаза 1 · Фундамент (месяцы 2–3)", fill=BLUE, size=10.5, align=PP_ALIGN.LEFT)
    body2 = rect(slide, x2, y + 440000, col_w, 2380000)
    tf2 = body2.text_frame
    f1 = [
        ("Steering Committee", "состав, ритм, полномочия"),
        ("Целевая оргсхема", "два заместителя, ПМО, менеджеры проектов"),
        ("RACI v2", "с владельцами процессов центрального офиса"),
        ("Таблица ролей", "21 роль методологии → новая структура"),
        ("Пауза утверждения ДИ", "до сверки с пунктом 3.5 методологии"),
        ("Минимальный пакет ДРБ → ДУП", "письменно зафиксирован приказом"),
    ]
    for j, (head, desc) in enumerate(f1):
        p = tf2.paragraphs[0] if j == 0 else tf2.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.2
        p.space_before = Pt(0 if j == 0 else 9)
        r = p.add_run()
        r.text = head + " — "
        r.font.name = FONT_BOLD
        r.font.size = Pt(9)
        r.font.color.rgb = DARK
        r2 = p.add_run()
        r2.text = desc
        r2.font.name = FONT
        r2.font.size = Pt(9)
        r2.font.color.rgb = GRAY

    y3 = y + 2900000
    gate = rect(slide, MARGIN_L, y3, CONTENT_W, 560000, fill=CARD_BG)
    tfg = gate.text_frame
    tfg.vertical_anchor = MSO_ANCHOR.MIDDLE
    tfg.margin_left = Emu(200000)
    tfg.margin_right = Emu(200000)
    p = tfg.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = 1.25
    r = p.add_run()
    r.text = "Gate перехода:  "
    r.font.name = FONT_BOLD
    r.font.size = Pt(9.5)
    r.font.color.rgb = GREEN
    r2 = p.add_run()
    r2.text = (
        "аудит завершён · не менее 70% директоров ответили на опросник · проведено минимум два интервью с ЦО · "
        "RACI v2 согласована с ДРБ и Финблоком"
    )
    r2.font.name = FONT
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = TEXT


def build_phase_detail_234(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Фазы 2–4: исполнение, TO-BE и цифра", "Месяцы 3–12 · основной объём изменений")

    y = CONTENT_TOP + 100000
    gap = 170000
    col_w = (CONTENT_W - 2 * gap) // 3
    phases = [
        (
            "Фаза 2 · Исполнение",
            "Месяцы 3–6",
            GREEN,
            [
                "10 инициатив восстановления процедур",
                "Реестр прогнозов ведётся по календарю",
                "Трекер закупок с нормативами сроков",
                "Регламент статус-встреч",
                "Аудит масштабирован на 3 филиала",
                "Единый канал претензий партнёров",
            ],
            "≥ 8 из 10 инициатив в эксплуатации",
        ),
        (
            "Фаза 3 · TO-BE",
            "Месяцы 5–8",
            YELLOW,
            [
                "Регламент gate готовности продукта",
                "Чек-лист передачи «реализация → сопровождение»",
                "Сквозной регламент открытия филиала",
                "Приказ о разделении заместителей",
                "Утверждение должностных инструкций",
                "Матрица отчётности без дублей",
            ],
            "3 регламента утверждены, handover протестирован",
        ),
        (
            "Фаза 4 · Цифра",
            "Месяцы 7–12",
            BLUE_LIGHT,
            [
                "MVP единой карточки проекта",
                "Автоматизация реестра прогнозов",
                "Уведомления по срокам и платежам",
                "Интеграции закупок и 1С",
                "Актуализация ВНД: роли, глоссарий, архив",
                "Дерево процессов v5 (TO-BE)",
            ],
            "KPI достигнуты, проект передан в ПМО",
        ),
    ]
    for i, (name, period, accent, items, gate) in enumerate(phases):
        x = MARGIN_L + i * (col_w + gap)
        head = rect(slide, x, y, col_w, 500000, fill=accent, radius=0.1)
        tf = head.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.line_spacing = 1.05
        r = p.add_run()
        r.text = name
        r.font.name = FONT_BOLD
        r.font.size = Pt(10.5)
        r.font.color.rgb = on(accent)
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.line_spacing = 1.05
        r2 = p2.add_run()
        r2.text = period
        r2.font.name = FONT
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = on(accent)

        body = rect(slide, x, y + 520000, col_w, 2200000)
        body.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        bullets(body, items, size=8.5, prefix="· ", spacing=10, line_spacing=1.25)

        gate_box = rect(slide, x, y + 2770000, col_w, 560000, fill=CARD_BG)
        tfg = gate_box.text_frame
        tfg.vertical_anchor = MSO_ANCHOR.MIDDLE
        pg = tfg.paragraphs[0]
        pg.alignment = PP_ALIGN.LEFT
        pg.line_spacing = 1.15
        rg = pg.add_run()
        rg.text = "Gate: "
        rg.font.name = FONT_BOLD
        rg.font.size = Pt(8.5)
        rg.font.color.rgb = ink(accent)
        rg2 = pg.add_run()
        rg2.text = gate
        rg2.font.name = FONT
        rg2.font.size = Pt(8.5)
        rg2.font.color.rgb = TEXT


def build_registry(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Реестр бизнес-процессов", "Рабочий инструмент управления изменениями — 51 запись")

    y = CONTENT_TOP + 120000
    gap = 180000
    col_w = (CONTENT_W - 2 * gap) // 3
    sheets = [
        (
            "Реестр процессов",
            GREEN,
            ["Процессы, подпроцессы и этапы", "Владелец (A) и исполнитель (R)", "Статус AS-IS и приоритет", "Ссылка на норматив и узкое место"],
        ),
        (
            "Справочник владельцев",
            BLUE,
            ["12 ролей с зонами ответственности", "Внутренние роли ДУП", "Владельцы процессов ЦО", "Основа для RACI v2"],
        ),
        (
            "Портфель трансформации",
            BLUE_LIGHT,
            ["15 инициатив изменений", "Привязка к процессу и фазе", "Владелец инициативы", "Ожидаемый эффект"],
        ),
    ]
    for i, (name, accent, lines) in enumerate(sheets):
        x = MARGIN_L + i * (col_w + gap)
        accent_bar(slide, x, y, col_w, 420000, name, fill=accent, size=10.5)
        body = rect(slide, x, y + 440000, col_w, 1300000)
        body.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        bullets(body, lines, size=9, prefix="· ", spacing=8, line_spacing=1.3)

    y2 = y + 1850000
    textbox(slide, MARGIN_L, y2, CONTENT_W, 240000, "Структура записи реестра", size=11, font=FONT_BOLD, color=DARK)
    rows = [
        ["Код", "Наименование", "Владелец (A)", "Статус AS-IS", "Приоритет"],
        ["1.1", "Приём и валидация пакета документов от ДРБ", "Зам. по реализации", "Не исполняется норматив", "P1"],
        ["1.2", "Подтверждение готовности продукта (пилот)", "Зам. по реализации + ДРБ", "Пробел", "P1"],
        ["2.3", "Прогноз поступлений и расходов", "Зам. по операционной деятельности", "Не исполняется", "P1"],
    ]
    table_grid(slide, MARGIN_L, y2 + 280000, CONTENT_W, [8, 32, 28, 22, 10], rows, row_h=330000, size=8.5)


def build_kpi(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "KPI проекта", "Целевые показатели на горизонте 12 месяцев")

    rows = [
        ["Показатель", "Базовая линия", "Цель через 12 месяцев"],
        ["Проекты с полным пакетом документов на старте", "не измеряется", "≥ 90%"],
        ["Среднее время старта проекта после передачи от ДРБ", "не измеряется", "−30%"],
        ["Заявки на оплату, проведённые в срок", "ad-hoc", "≥ 85%"],
        ["Регулярность прогноза поступлений", "чат и разовые запросы", "100% по календарю"],
        ["Проекты с регулярными статус-встречами", "нерегулярно", "≥ 90% ежемесячно"],
        ["Покрытие портфеля аудитом качества данных", "0% с 2020 года", "100% в год"],
        ["Удовлетворённость филиалов взаимодействием", "базовый замер в фазе 0", "+20 п.п."],
        ["Доля проектных данных в единой системе", "фрагментировано", "≥ 80%"],
    ]
    table_grid(slide, MARGIN_L, CONTENT_TOP + 60000, CONTENT_W, [52, 24, 24], rows, row_h=370000, size=9)


def build_governance(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Управление проектом", "Governance, роли и ритм принятия решений")

    y = CONTENT_TOP + 120000
    col_w = (CONTENT_W - 200000) // 2

    textbox(slide, MARGIN_L, y, col_w, 240000, "Структура управления", size=11, font=FONT_BOLD, color=DARK)
    roles = [
        ("Steering Committee", "Директор ДУП, заместители, представители ДРБ и Финблока", "Приоритеты, решения, эскалации"),
        ("Руководитель проекта", "Назначается из ДУП или ПМО", "Дорожная карта, риски, ежедневное управление"),
        ("Владельцы процессов", "Владельцы блоков 1–4", "Ответственность за целевое состояние"),
        ("Рабочие группы", "ДУП + центральный офис + филиалы", "Регламенты и интерфейсы"),
    ]
    ry = y + 300000
    for name, who, what in roles:
        box = rect(slide, MARGIN_L, ry, col_w, 620000)
        tf = box.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.15
        r = p.add_run()
        r.text = name
        r.font.name = FONT_BOLD
        r.font.size = Pt(9.5)
        r.font.color.rgb = BLUE
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.LEFT
        p2.line_spacing = 1.15
        r2 = p2.add_run()
        r2.text = who
        r2.font.name = FONT
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = TEXT
        p3 = tf.add_paragraph()
        p3.alignment = PP_ALIGN.LEFT
        p3.line_spacing = 1.15
        r3 = p3.add_run()
        r3.text = what
        r3.font.name = FONT
        r3.font.size = Pt(8.5)
        r3.font.color.rgb = GRAY
        ry += 680000

    x2 = MARGIN_L + col_w + 200000
    textbox(slide, x2, y, col_w, 240000, "Ритм управления", size=11, font=FONT_BOLD, color=DARK)
    rhythm = [
        ("Еженедельно", "Статус проекта: руководитель проекта → директор ДУП"),
        ("Раз в две недели", "Рабочая группа по процессам и регламентам"),
        ("Ежемесячно", "Steering Committee: решения и эскалации"),
        ("Ежеквартально", "Пересмотр дорожной карты и KPI, отчёт Правлению"),
    ]
    ry2 = y + 300000
    for period, desc in rhythm:
        badge = rect(slide, x2, ry2, 1180000, 300000, fill=GREEN, radius=0.16)
        tfb = badge.text_frame
        tfb.vertical_anchor = MSO_ANCHOR.MIDDLE
        tfb.margin_left = tfb.margin_right = 0
        set_text(tfb, period, size=8.5, font=FONT_BOLD, color=WHITE, align=PP_ALIGN.CENTER)
        textbox(slide, x2 + 1300000, ry2 + 30000, col_w - 1300000, 300000, desc, size=9, color=TEXT, line_spacing=1.2)
        ry2 += 420000

    note = rect(slide, x2, ry2 + 120000, col_w, 900000, fill=CARD_BG)
    tfn = note.text_frame
    tfn.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tfn.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = 1.3
    r = p.add_run()
    r.text = "Условие успеха: "
    r.font.name = FONT_BOLD
    r.font.size = Pt(9)
    r.font.color.rgb = BLUE
    r2 = p.add_run()
    r2.text = (
        "спонсор проекта присутствует на Steering Committee. Изменения затрагивают процессы "
        "смежных подразделений, и решения такого уровня принимаются только с участием руководства."
    )
    r2.font.name = FONT
    r2.font.size = Pt(9)
    r2.font.color.rgb = TEXT


def build_risks(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Ключевые риски и меры", "Управление рисками ведётся с фазы 0")

    rows = [
        ["Риск", "Оценка", "Мера снижения"],
        ["Разделение заместителей без чек-листа передачи", "Высокий", "Чек-лист handover (2.0) утверждается до реструктуризации"],
        ["Перегрузка единственного заместителя", "Высокий", "Жёсткая приоритизация P1 и быстрые победы в фазе 2"],
        ["Сопротивление изменениям в центральном офисе", "Средний", "Интервью, RACI v2, спонсор на Steering Committee"],
        ["Недоступность ИТ-ресурсов", "Средний", "Поэтапная автоматизация, MVP вместо полной платформы"],
        ["Низкий отклик филиалов на опросник", "Средний", "Служебная записка от руководства, контроль сроков"],
        ["Методология формально отменена или изменена", "Низкий", "Проверка статуса приказа в фазе 0 до начала работ"],
    ]
    row_h = 390000
    table_grid(slide, MARGIN_L, CONTENT_TOP + 60000, CONTENT_W, [42, 12, 46], rows, row_h=row_h, size=9)

    textbox(
        slide,
        MARGIN_L,
        CONTENT_TOP + 60000 + row_h * len(rows) + 150000,
        CONTENT_W,
        260000,
        "Первые два риска связаны с одной причиной — совмещением блоков 1 и 2 в одной роли. "
        "Их снятие является приоритетом фаз 1 и 3.",
        size=9,
        font=FONT_MED,
        color=GRAY,
    )


def build_next_steps(prs: Presentation) -> None:
    slide = new_slide(prs)
    title(slide, "Первые шаги", "Шесть решений для запуска проекта")

    steps = [
        ("01", "Утвердить паспорт проекта", "Назначить руководителя проекта трансформации", GREEN),
        ("02", "Созвать Steering Committee", "ДУП, ДРБ, Финблок — утвердить приоритеты", GREEN),
        ("03", "Запустить опросник филиалов", "Форма готова, требуется служебная записка", BLUE),
        ("04", "Назначить пилотный филиал", "Аудит по чек-листам методологии", BLUE),
        ("05", "Поставить паузу на утверждение ДИ", "До сверки с пунктом 3.5 методологии", YELLOW),
        ("06", "Зафиксировать пакет ДРБ → ДУП", "Письменно, с блокировкой старта без пакета", YELLOW),
    ]
    y = CONTENT_TOP + 140000
    gap_x = 200000
    gap_y = 210000
    row_h = 940000
    col_w = (CONTENT_W - gap_x) // 2
    for i, (num, head, desc, accent) in enumerate(steps):
        col = i % 2
        row = i // 2
        x = MARGIN_L + col * (col_w + gap_x)
        yy = y + row * (row_h + gap_y)
        box = rect(slide, x, yy, col_w, row_h)
        tf = box.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Emu(700000)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.15
        r = p.add_run()
        r.text = head
        r.font.name = FONT_BOLD
        r.font.size = Pt(10.5)
        r.font.color.rgb = DARK
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.LEFT
        p2.line_spacing = 1.2
        p2.space_before = Pt(3)
        r2 = p2.add_run()
        r2.text = desc
        r2.font.name = FONT
        r2.font.size = Pt(9)
        r2.font.color.rgb = GRAY

        badge = rect(slide, x + 150000, yy + 270000, 420000, 400000, fill=accent, radius=0.2)
        tfb = badge.text_frame
        tfb.vertical_anchor = MSO_ANCHOR.MIDDLE
        tfb.margin_left = tfb.margin_right = 0
        set_text(tfb, num, size=12, font=FONT_BOLD, color=on(accent), align=PP_ALIGN.CENTER)


def build_closing(prs: Presentation) -> None:
    """Финальный слайд шаблона «Спасибо!» — адаптируем сообщение."""
    slide = prs.slides[-1]
    msg = find_shape(slide, "Присоединяйтесь")
    if msg is not None:
        replace_text(
            msg,
            "Управляемые процессы — предсказуемые проекты в каждом филиале",
            size=16,
            font=FONT_MED,
            color=WHITE,
            align=PP_ALIGN.CENTER,
        )


# --------------------------------------------------------------------------
def main() -> None:
    if not TEMPLATE.exists():
        raise SystemExit(f"Не найден шаблон брендбука: {TEMPLATE}")

    prs = Presentation(str(TEMPLATE))

    # Оставляем титульный слайд (0) и финальный «Спасибо!» (11), остальные удаляем.
    for idx in range(10, 0, -1):
        drop_slide(prs, idx)

    reserve_partname(prs.slides[1], 90)
    build_cover(prs)

    builders = [
        build_agenda,
        build_executive_summary,
        build_sources,
        build_process_map,
        build_diagnosis,
        build_gaps,
        build_directions,
        build_optimization,
        build_transformation,
        build_digital,
        build_roadmap,
        build_phase_detail_01,
        build_phase_detail_234,
        build_registry,
        build_kpi,
        build_governance,
        build_risks,
        build_next_steps,
    ]
    for builder in builders:
        builder(prs)

    move_slide_to_end(prs, 1)  # «Спасибо!» в конец
    build_closing(prs)

    prs.save(str(OUTPUT))
    print(f"Сохранено: {OUTPUT}")
    print(f"Слайдов: {len(prs.slides.__iter__.__self__._sldIdLst)}")


if __name__ == "__main__":
    main()

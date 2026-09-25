#!/usr/bin/env python3
"""Доклад по реинжинирингу процессов ДУП: от предпосылок до очереди работ.

Колода ведёт слушателя по одной линии: зачем взялись за проект → что уже
сделано → что такое система управления процессами и как она измеряет процесс
→ как построена приоритизация → где мы сейчас и куда идём → очередь работ
и решения, которые нужны от руководства.

Все цифры, объекты, волны и правила очерёдности импортируются из
build_dup_priority.py — того же модуля, который собирает
«ДУП_приоритеты_реинжиниринга.xlsx», поэтому слайды не могут разойтись
с реестром. Состояние AS-IS и состав участников тоже считаются по реестру,
а не набираются вручную.

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
    bullets,
    divider,
    ink,
    on,
    rect,
    set_text,
    table_grid,
    textbox,
    title,
)

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE.parent / "01_Процессы"))

import build_dup_metrics as metrics  # noqa: E402
import build_dup_priority as P  # noqa: E402
from build_dup_role_map import DECISIONS_AWAITING_DOCUMENT  # noqa: E402

TEMPLATE = BASE.parent / "brandbook" / "Общие слайды_ver 12.12.24.pptx"
OUTPUT = BASE / "ДУП_Приоритеты_реинжиниринга_презентация.pptx"

DECK_TITLE = "Реинжиниринг процессов ДУП"
DECK_SUBTITLE = "От предпосылок до очереди работ"
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
AWAITING_ORDER = DECISIONS_AWAITING_DOCUMENT

# Группы первой волны: набор типов вмешательства, подпись и пояснение.
# Размер волны — первое, о чём спрашивают, поэтому слайд показывает,
# из чего она складывается. Состав считается по данным, а не вписывается руками.
WAVE_1_GROUPS = [
    ({P.ЗАКРЕПЛЕНИЕ, P.ПРОЕКТ}, "доводим владение до приказа",
     "Чек-лист передачи между блоками 1 и 2, разделение заместителей, полномочия ПМО"),
    ({P.БАЗА}, "снимаем базовую линию",
     "Пилотный аудит одного филиала и регулярный контроль качества данных"),
    ({P.SLA}, "нормируем срок стыка с ЦО",
     "Решения волны вступают в силу в день приказа HR — нужен норматив срока"),
]


def wave_1_breakdown():
    wave = [r for r in ROWS if r["wave"] == P.В1]
    groups = [(len([r for r in wave if r["kind"] in kinds]), label, note)
              for kinds, label, note in WAVE_1_GROUPS]
    if sum(count for count, *_ in groups) != len(wave):
        raise SystemExit("Разбивка волны 1 не покрывает все её объекты")
    return groups


# --------------------------------------------------------------------------
# Содержательная часть доклада
# --------------------------------------------------------------------------
# Предпосылки и целевое состояние взяты из паспорта проекта
# (00_Управление_проектом/01_Паспорт_проекта.md), определения и цикл
# управления — из Методики KT_METHOD v1. Это повествование, а не данные
# реестра: цифры в него подставляются из модулей, текст живёт здесь.
PREMISES = [
    ("Портфель вырос, управление — нет",
     "ДУП ведёт доходные проекты компании по ГЧП и госзакупу: фото-видеофиксация, каналы "
     "связи, эко-мониторинг. Портфель и филиальная сеть выросли, а управление осталось "
     "на устных договорённостях."),
    ("Методика компании обязательна, но не применяется",
     "В компании действует Методика системы управления бизнес-процессами KT_METHOD v1. "
     "Она обязательна для всех работников (п. 4), однако процессы ДУП по ней не описаны, "
     "владельцы не закреплены по критериям, метрики не установлены."),
    ("Спорить о сроке приходится словами",
     "На старте проекта часть процессов не была закреплена ни за одной должностью, "
     "а фактические значения не измерялись. Ни доказать эффект изменения, ни назвать "
     "виновника срыва было нечем — только мнения."),
]

PROJECT_GOAL = (
    "Цель проекта: выстроить управляемую, прозрачную и масштабируемую модель управления "
    "проектами — предсказуемый запуск во всех филиалах, единые правила стыка филиала, "
    "ДУП и центрального офиса, восстановленное исполнение методологии."
)

ANALYSED = [
    "Слой А — текущий анализ: интервью, дерево процессов v4, RACI-матрица, кадровая матрица",
    "Слой B — наследие методологии «Сергек»: 434 документа, реестр Confluence",
    "Методика системы управления бизнес-процессами KT_METHOD v1 — 52 страницы",
    "Матрица стыков ДУП со смежными подразделениями центрального офиса",
]

# Построенные документы: подпись и функция, считающая объём по данным.
PRODUCED = [
    ("Реестр процессной модели",
     lambda: f"{objects_word(len(P.PROCESSES))}: процессы, подпроцессы и этапы"),
    ("Карта ролей ДУП",
     lambda: "роль ДУП и бизнес-владелец у каждого объекта"),
    ("Реестр метрик",
     lambda: f"{CRITERIA_COUNT} {plural(CRITERIA_COUNT, 'критерий', 'критерия', 'критериев')} "
             "оценки с формулой и источником данных"),
    ("Приоритеты реинжиниринга",
     lambda: f"{objects_word(len(SCORED))} — балл, место в очереди и обоснование"),
]

# Определение процесса — глоссарий Методики; цель — п. 3; цикл — Таблица 20.
PROCESS_DEFINITION = [
    "Процесс — сквозная последовательность действий, преобразующая входы в результат, "
    "ценный для клиента. У процесса ровно один бизнес-владелец, отвечающий за результат целиком.",
    "Система управления процессами нужна, чтобы описывать, измерять и улучшать работу "
    "одинаково во всех подразделениях, не спорить о том, кто владелец, и не терять "
    "ответственность при передаче процесса между блоками (п. 3 Методики).",
]

PROCESS_CYCLE = [
    ("Проектирование", "Схема, паспорт процесса, карта SIPOC, KPI и SLA, назначение владельца"),
    ("Внедрение", "Автоматизация шагов, интеграция систем, обучение, ввод в эксплуатацию"),
    ("Анализ и оценка", "Сбор фактических значений метрик, контроль SLA, отчётность"),
    ("Улучшение", "Поиск узких мест, оценка зрелости, приоритизация улучшений"),
    ("Оптимизация", "Устранение узких мест, рост зрелости, тиражирование решений"),
]

CYCLE_NOTE = (
    "Проект ДУП стоит на четвёртом этапе для описанных процессов — мы ищем узкие места "
    "и расставляем улучшения по очереди — и возвращается на первый там, где процедуры "
    "не существует вовсе."
)

# Метрики для слайда о формулах: по две-три из каждой группы Методики.
# Наименования, формулы и целевые значения берутся из справочника, а не набираются.
FORMULA_KEYS = ["CT", "OT", "WT", "FPY", "DR", "COPQ", "BV", "BNI", "SLA", "AR"]

# Что предлагаем — четыре направления из паспорта проекта.
TO_BE_DIRECTIONS = [
    ("Восстановить исполнение",
     "Там, где процедура уже описана в методологии, — вернуть её в работу"),
    ("Закрыть подлинные пробелы",
     "Спроектировать с нуля три процедуры, которых нет ни в одном документе"),
    ("Закрепить организацию",
     "Разделение заместителей, полномочия процессного офиса, RACI v2 и паспорта процессов"),
    ("Оцифровать управляемое",
     "Единая карточка проекта, автопрогноз и интеграции — после того, как процесс исполняется"),
]

# Показатели из паспорта проекта: «сегодня» — фактическое состояние, «цель» — горизонт 12 месяцев.
TO_BE_KPI = [
    ("Полный пакет документов на старте проекта", "не измеряется", "≥ 90% проектов"),
    ("Срок старта проекта после передачи от ДРБ", "не измеряется", "−30% к базовой линии"),
    ("Заявки на оплату, поданные в срок", "по обстоятельствам", "≥ 85%"),
    ("Покрытие портфеля аудитом по чек-листам", "0% с 2020 года", "100% портфеля за год"),
    ("Регулярность прогноза поступлений", "чаты и разовые запросы", "100% по календарю"),
    ("Доля данных в единой системе", "фрагментировано", "≥ 80%"),
]

# Участники: роль в проекте, кто её исполняет, за что отвечает.
PARTICIPANTS = [
    ("Заказчик и спонсор", "Директор ДУП",
     "Приоритеты, приказы о владении и полномочиях, эскалации"),
    ("Руководитель проекта", "Назначается из ДУП / ГМП",
     "Дорожная карта, риски, еженедельный статус"),
    ("Владельцы процессов", "Зам. по реализации, зам. по сопровождению, ГМП",
     "Результат своих блоков, паспорта процессов, целевые значения метрик"),
    ("Процессный офис (ПМО)", "Главный менеджер проектов",
     "Методология, аудит качества данных, реестр процессов и метрик"),
    ("Рабочие группы", "ДУП, центральный офис, филиалы",
     "Регламенты стыков, SLA, проверка решений на практике"),
]

# Фазы дорожной карты: номер, название, срок, ключевой результат.
ROADMAP = [
    ("Фаза 0", "Диагностика", "месяц 1", "Пилотный аудит филиала и базовая линия"),
    ("Фаза 1", "Фундамент", "месяцы 2–3", "Приказы о владении, RACI v2, целевая оргсхема"),
    ("Фаза 2", "Исполнение", "месяцы 3–6", "Двенадцать процедур возвращены в работу"),
    ("Фаза 3", "Проектирование", "месяцы 5–8", "Три новых регламента и разделение замов"),
    ("Фаза 4", "Цифра и устойчивость", "месяцы 7–12", "Единая карточка проекта, ВНД v2, KPI"),
]


# --------------------------------------------------------------------------
# Состояние AS-IS: считается по реестру, а не вписывается в слайд
# --------------------------------------------------------------------------
# Колонка «Статус AS-IS» карты ролей — свободный текст, но начало строки
# устойчиво: «Не исполняется …», «Пробел …», «Вне периметра …». Группируем по
# началу строки, чтобы слайд нельзя было рассинхронизировать с картой ролей.
AS_IS_FAMILIES = {
    "не исполняется": "не исполняется",
    "пробел": "пробел",
    "вне периметра": "вне периметра",
}


def as_is_count(family: str) -> int:
    return sum(1 for p in P.PROCESSES if p[14].lower().startswith(family))


# Уровень зрелости: в Методике итог процесса равен наименьшему подтверждённому
# уровню среди семи критериев (п. 179), поэтому берём минимум по всей таблице.
MATURITY_NOW = min(row[2] for row in metrics.MATURITY).split("—")[0].strip()
MATURITY_GOAL = metrics.MATURITY_TARGET.split("—")[0].strip()
MATURITY_PROCESSES = len({row[0] for row in metrics.MATURITY})

CRITERIA_COUNT = len(metrics.criteria_rows())


def as_is_rows():
    """Строки слайда AS-IS: цифра, факт и то, чем он оборачивается.

    Цифры считаются по реестру, формулировки живут здесь. Пятая строка про
    аудит — из паспорта проекта: база «0% с 2020 года» в реестре не хранится.
    """
    dormant = as_is_count("не исполняется")
    gaps = as_is_count("пробел")
    return [
        (str(dormant),
         f"{plural(dormant, 'объект', 'объекта', 'объектов')} из {len(P.PROCESSES)} "
         "описаны, но не исполняются",
         "Проект стартует по договорённости, а не по правилу: срок зависит от того, "
         "кто вспомнил о шаге и до кого дошли руки."),
        (str(gaps),
         f"{plural(gaps, 'процедуры', 'процедуры', 'процедур')} не существует вовсе",
         "Готовность продукта, передача в сопровождение и открытие филиала решаются "
         "каждый раз заново и доходят до директора."),
        ("0",
         f"критериев из {CRITERIA_COUNT} измеряется сегодня",
         "Ни срок, ни стоимость, ни качество подтвердить нечем: спор о результате "
         "выигрывает тот, кто увереннее."),
        (MATURITY_NOW,
         "из 5 — уровень зрелости каждого сквозного процесса ДУП",
         f"По Методике это «Начальный»: результат держится на конкретном человеке, "
         f"а не на системе. Целевой уровень — {MATURITY_GOAL}."),
        ("0%",
         "портфеля проверено аудитом по чек-листам с 2020 года",
         "Чек-листы написаны и лежат в методологии. Ошибка в проектных данных всплывает "
         "не на аудите, а на приёмке у госпартнёра."),
    ]


def wave_no(wave: str) -> str:
    """«Волна 2. Быстрые победы…» -> «2»."""
    return wave.split(".")[0].replace("Волна", "").strip()


def wave_name(wave: str) -> str:
    """Часть после номера: «Быстрые победы внутри ДУП»."""
    return wave.split(". ", 1)[1] if ". " in wave else wave


def plural(n: int, one: str, few: str, many: str) -> str:
    """Форма существительного при числе: 1 объект, 2 объекта, 5 объектов.

    Счёт на слайдах берётся из реестра и меняется вместе с ним, поэтому
    окончание нельзя вписать в текст руками.
    """
    if n % 100 in range(11, 15):
        return many
    return {1: one, 2: few, 3: few, 4: few}.get(n % 10, many)


def objects_word(n: int) -> str:
    return f"{n} {plural(n, 'объект', 'объекта', 'объектов')}"


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


def build_premises(prs: Presentation) -> None:
    """Предпосылки: зачем компания вообще взялась за этот проект."""
    slide = new_slide(prs)
    title(slide, "Почему мы взялись за этот проект",
          "Три предпосылки, с которых начался проект трансформации процессов ДУП")

    col_w = (CONTENT_W - 2 * 180000) // 3
    y = CONTENT_TOP + 140000
    card_h = 2150000
    for i, (heading, body) in enumerate(PREMISES):
        x = MARGIN_L + i * (col_w + 180000)
        rect(slide, x, y, col_w, card_h, fill=CARD_BG)
        rect(slide, x, y, 420000, 26000, fill=GREEN, shape=MSO_SHAPE.RECTANGLE)
        textbox(slide, x + 160000, y + 240000, col_w - 320000, 560000, heading,
                size=12, font=FONT_BOLD, color=BLUE, line_spacing=1.15)
        textbox(slide, x + 160000, y + 860000, col_w - 320000, card_h - 1000000,
                body, size=9, color=TEXT, line_spacing=1.3)

    accent_bar(slide, MARGIN_L, y + card_h + 240000, CONTENT_W, 640000,
               PROJECT_GOAL, fill=BLUE, size=10.5, align=PP_ALIGN.LEFT)


def build_work_done(prs: Presentation) -> None:
    """Что уже сделано: что разобрали и что из этого построили."""
    slide = new_slide(prs)
    title(slide, "Что уже сделано в проекте",
          "Диагностика завершена: процессная модель описана, владение закреплено, "
          "метрики и очередь работ определены")

    col_w = (CONTENT_W - 220000) // 2
    y = CONTENT_TOP + 140000
    card_h = 2860000

    left = rect(slide, MARGIN_L, y, col_w, card_h)
    textbox(slide, MARGIN_L + 160000, y + 180000, col_w - 320000, 320000,
            "Что проанализировали", size=12, font=FONT_BOLD, color=BLUE)
    bullets(left, ANALYSED, size=9.5, spacing=11, line_spacing=1.3)
    left.text_frame.margin_top = Emu(620000)

    right_x = MARGIN_L + col_w + 220000
    rect(slide, right_x, y, col_w, card_h)
    textbox(slide, right_x + 160000, y + 180000, col_w - 320000, 320000,
            "Что построено", size=12, font=FONT_BOLD, color=ink(GREEN))
    item_y = y + 620000
    for name, measure in PRODUCED:
        textbox(slide, right_x + 160000, item_y, col_w - 320000, 230000, name,
                size=10, font=FONT_BOLD, color=DARK)
        textbox(slide, right_x + 160000, item_y + 230000, col_w - 320000, 300000,
                measure(), size=9, color=GRAY, line_spacing=1.25)
        item_y += 540000

    textbox(
        slide, MARGIN_L, y + card_h + 190000, CONTENT_W, 300000,
        "Четыре файла собираются из одного источника: правка в карте ролей сама проходит "
        "в метрики, приоритеты и эту презентацию — разойтись между собой они не могут.",
        size=9, color=STEEL,
    )


def build_process_system(prs: Presentation) -> None:
    """Что такое система управления процессами и зачем она нужна."""
    slide = new_slide(prs)
    title(slide, "Система управления процессами",
          "Определение, цель и цикл управления — из Методики KT_METHOD v1, "
          "обязательной для всех подразделений компании")

    bar_h = 800000
    accent_bar(slide, MARGIN_L, CONTENT_TOP + 60000, CONTENT_W, bar_h,
               PROCESS_DEFINITION, fill=BLUE, size=10, align=PP_ALIGN.LEFT)

    label_y = CONTENT_TOP + bar_h + 300000
    textbox(slide, MARGIN_L, label_y, CONTENT_W, 260000,
            "Цикл управления процессом — пять последовательных этапов (Таблица 20 Методики)",
            size=10.5, font=FONT_MED, color=DARK)

    col_w = (CONTENT_W - 4 * 120000) // 5
    y = label_y + 340000
    for i, (name, action) in enumerate(PROCESS_CYCLE, 1):
        x = MARGIN_L + (i - 1) * (col_w + 120000)
        accent = GREEN if i == 4 else BLUE_LIGHT
        head = rect(slide, x, y, col_w, 380000, fill=accent, radius=0.12)
        tf = head.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_top = tf.margin_bottom = 0
        set_text(tf, f"{i}. {name}", size=9.5, font=FONT_BOLD, color=on(accent),
                 align=PP_ALIGN.CENTER)
        body = rect(slide, x, y + 400000, col_w, 900000, fill=CARD_BG)
        tf = body.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        set_text(tf, action, size=8.5, color=TEXT, line_spacing=1.25)

    textbox(slide, MARGIN_L, y + 1400000, CONTENT_W, 300000, CYCLE_NOTE,
            size=9, color=STEEL, line_spacing=1.25)


def build_formulas(prs: Presentation) -> None:
    """Как измеряется процесс: формулы метрик из Методики."""
    slide = new_slide(prs)
    title(slide, "Как измеряется процесс",
          "Формулы и целевые значения взяты из Глав 9–13 Методики, а не придуманы "
          "под задачу")

    rows = [("Группа", "Метрика", "Формула расчёта", "Целевое значение")]
    for key in FORMULA_KEYS:
        name, group, _purpose, formula, _unit, target, _chapter = metrics.METRICS[key]
        rows.append((group, name, formula, target))
    end_y = table_grid(
        slide, MARGIN_L, CONTENT_TOP + 210000, CONTENT_W,
        [16, 30, 34, 20], rows, row_h=262000, size=8.5,
    )

    textbox(
        slide, MARGIN_L, end_y + 220000, CONTENT_W, 420000,
        [f"В справочнике Методики {len(metrics.METRICS)} метрик; по объектам реестра из них "
         f"разложено {CRITERIA_COUNT} "
         f"{plural(CRITERIA_COUNT, 'критерий', 'критерия', 'критериев')} оценки.",
         "Набор метрик зависит от роли ДУП: владелец отвечает за объект целиком, "
         "участник — за свой шаг, клиент — за требования к входу и приёмку."],
        size=9, color=GRAY, line_spacing=1.3, space_after=4,
    )


def build_as_is(prs: Presentation) -> None:
    """AS-IS: состояние, зафиксированное по каждому объекту реестра."""
    slide = new_slide(prs)
    title(slide, "Где мы сейчас",
          "Состояние зафиксировано по каждому из объектов реестра — это диагноз системы, "
          "а не оценка людей")

    y = CONTENT_TOP + 120000
    row_h = 620000
    rows = as_is_rows()
    for i, (number, fact, consequence) in enumerate(rows):
        badge = rect(slide, MARGIN_L, y, 700000, 420000, fill=RED, radius=0.12)
        tf = badge.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        set_text(tf, number, size=15, font=FONT_BOLD, color=WHITE, align=PP_ALIGN.CENTER)

        text_x = MARGIN_L + 840000
        text_w = CONTENT_W - 840000
        textbox(slide, text_x, y - 10000, text_w, 250000, fact,
                size=11, font=FONT_BOLD, color=DARK)
        textbox(slide, text_x, y + 240000, text_w, 320000, consequence,
                size=9, color=GRAY, line_spacing=1.25)
        if i < len(rows) - 1:
            divider(slide, MARGIN_L, y + row_h - 80000, CONTENT_W)
        y += row_h

    textbox(
        slide, MARGIN_L, y + 30000, CONTENT_W, 300000,
        "Важная оговорка: двенадцать из пятнадцати узких мест уже описаны в методологии "
        "компании. Проблема не в том, что никто не придумал правил, а в том, что "
        "правила перестали исполняться.",
        size=9, color=STEEL, line_spacing=1.25,
    )


def build_to_be(prs: Presentation) -> None:
    """TO-BE: что предлагаем сделать и что компания получит."""
    slide = new_slide(prs)
    title(slide, "Что предлагаем и что получим",
          "Целевое состояние: процессы описаны, измеряются и улучшаются по циклу Методики")

    left_w = 3180000
    y = CONTENT_TOP + 160000
    textbox(slide, MARGIN_L, y - 60000, left_w, 260000, "Что предлагаем",
            size=11, font=FONT_BOLD, color=BLUE)
    row_y = y + 300000
    for i, (heading, body) in enumerate(TO_BE_DIRECTIONS, 1):
        numbered_row(slide, MARGIN_L, row_y, left_w, 640000, i, heading, body,
                     accent=GREEN, heading_size=10, body_size=8.5)
        row_y += 700000

    right_x = MARGIN_L + left_w + 280000
    right_w = CONTENT_W - left_w - 280000
    textbox(slide, right_x, y - 60000, right_w, 260000, "Что получим за 12 месяцев",
            size=11, font=FONT_BOLD, color=ink(GREEN))
    rows = [("Показатель", "Сегодня", "Цель")]
    rows += TO_BE_KPI
    table_grid(slide, right_x, y + 300000, right_w, [52, 24, 24], rows,
               row_h=340000, size=8.5)

    textbox(
        slide, MARGIN_L, row_y + 140000, CONTENT_W, 300000,
        f"Уровень зрелости процессов поднимается с {MATURITY_NOW} до {MATURITY_GOAL} из 5: "
        "паспорт процесса, измеряемые метрики и регулярный цикл улучшений.",
        size=9, color=STEEL, line_spacing=1.25,
    )


def build_headline(prs: Presentation) -> None:
    """Слайд решения: одна мысль и четыре цифры под ней."""
    slide = new_slide(prs)
    title(slide, "Очередь реинжиниринга построена")

    accent_bar(
        slide, MARGIN_L, CONTENT_TOP - 30000, CONTENT_W, 700000,
        ["Первыми берём не самые выгодные процессы, а те, что разблокируют остальные.",
         "Владелец закреплён у каждого объекта. Пока решение не доведено до приказа "
         "и не снят замер «до», выгоду от оптимизации нечем доказать."],
        fill=BLUE, size=12, align=PP_ALIGN.LEFT,
    )

    stats = [
        (str(len(ROWS)),
         f"{plural(len(ROWS), 'объект', 'объекта', 'объектов')} реестра оценены "
         "и поставлены в очередь", GREEN),
        (str(len(AWAITING_ORDER)),
         f"{plural(len(AWAITING_ORDER), 'решение', 'решения', 'решений')} о владении "
         "ждут документа — они держат очередь", RED),
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
    title(slide, "Как считается приоритет",
          "Пять критериев, шкала от 1 до 5 и вес критерия — одинаково для всех объектов реестра")

    rows = [("Критерий", "Вес", "На какой вопрос отвечает")]
    rows += [(name, str(weight), question) for name, weight, question, *_ in P.CRITERIA]
    end_y = table_grid(
        slide, MARGIN_L, CONTENT_TOP + 240000, CONTENT_W,
        [30, 10, 60], rows, row_h=356000, size=10,
    )

    accent_bar(
        slide, MARGIN_L, end_y + 200000, CONTENT_W, 460000,
        f"Балл объекта = сумма (оценка от 1 до 5 × вес критерия) ÷ {P.MAX_SCORE} × 100",
        fill=BLUE, size=11, align=PP_ALIGN.LEFT,
    )

    textbox(
        slide, MARGIN_L, end_y + 740000, CONTENT_W, 420000,
        ["Вес отражает управленческий приоритет: то, что даёт результат, стоит дороже того, "
         "что даёт удобство.",
         "В файле приоритетов расписаны все пять ступеней каждой шкалы — оценка 2 или 4 "
         "не ставится на глаз и защищается на комитете."],
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
        set_text(tf, f"{objects_word(BY_WAVE[wave])}  ·  {phase.replace(' дорожной карты', '')}",
                 size=9, font=FONT_MED, color=GRAY, align=PP_ALIGN.CENTER)

    textbox(
        slide, MARGIN_L, y + 2120000, CONTENT_W, 250000,
        f"Отдельно: {objects_word(BY_WAVE[P.ВН])} вне очереди — исполняются другим "
        f"подразделением целиком, срок контролируется через SLA по родительскому объекту.",
        size=8.5, color=STEEL,
    )

    breakdown_y = y + 2450000
    divider(slide, MARGIN_L, breakdown_y - 30000, CONTENT_W)
    textbox(
        slide, MARGIN_L, breakdown_y, CONTENT_W, 250000,
        f"Из чего состоит волна 1: {objects_word(BY_WAVE[P.В1])}",
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
          "Волна 1 целиком — закрепление владения и базовая линия")

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
         f"Они ждут волн {' и '.join(waves_used)}: сначала приказы и замер, "
         "иначе эффект будет не доказан, а объявлен."],
        fill=GREEN, size=10, align=PP_ALIGN.LEFT,
    )


def build_first_steps(prs: Presentation) -> None:
    slide = new_slide(prs)
    shown = 5
    title(slide, "Программа старта: пять первых действий",
          "Шаги 1–5 доводят владение до приказа и снимают базовую линию")

    y = CONTENT_TOP + 140000
    row_h = 620000
    owner_w = 2200000
    owner_x = MARGIN_L + CONTENT_W - owner_w
    for i, step in enumerate(P.FIRST_STEPS[:shown], 1):
        action, objects, owner, phase, result, _why = step
        numbered_row(
            slide, MARGIN_L, y, CONTENT_W - owner_w - 200000, row_h - 80000, i,
            action, f"Результат: {result}", accent=BLUE, heading_size=10,
        )
        textbox(slide, owner_x, y - 10000, owner_w, 330000, owner,
                size=9, font=FONT_MED, color=ink(GREEN), line_spacing=1.2)
        textbox(slide, owner_x, y + 340000, owner_w, 220000,
                f"{phase}  ·  {objects}", size=8, color=STEEL)
        if i < shown:
            divider(slide, MARGIN_L, y + row_h - 90000, CONTENT_W)
        y += row_h

    textbox(
        slide, MARGIN_L, y + 40000, CONTENT_W, 280000,
        f"Шаги {shown + 1}–{len(P.FIRST_STEPS)} — восстановление исполнения внутри ДУП: "
        "матрица кураторов, приёмка пакета от ДРБ, платежи, статус-встречи, прогноз, "
        "переносы оборудования.",
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


def build_team_roadmap(prs: Presentation) -> None:
    """Кто делает проект и в какие фазы — последний содержательный слайд."""
    slide = new_slide(prs)
    title(slide, "Участники проекта и дорожная карта",
          "Кто принимает решения, кто исполняет и в каком порядке идут фазы")

    col_w = (CONTENT_W - 4 * 110000) // 5
    y = CONTENT_TOP + 60000
    for i, (phase, name, term, result) in enumerate(ROADMAP):
        x = MARGIN_L + i * (col_w + 110000)
        accent = GREEN if i <= 1 else BLUE_LIGHT
        head = rect(slide, x, y, col_w, 340000, fill=accent, radius=0.12)
        tf = head.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_top = tf.margin_bottom = 0
        set_text(tf, f"{phase} · {term}", size=8.5, font=FONT_BOLD, color=on(accent),
                 align=PP_ALIGN.CENTER)
        body = rect(slide, x, y + 360000, col_w, 800000, fill=CARD_BG)
        tf = body.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        set_text(tf, name, size=10, font=FONT_BOLD, color=ink(accent), line_spacing=1.15)
        para = tf.add_paragraph()
        para.space_before = Pt(5)
        para.line_spacing = 1.25
        run = para.add_run()
        run.text = result
        run.font.name = FONT
        run.font.size = Pt(8)
        run.font.color.rgb = TEXT

    table_y = y + 1400000
    divider(slide, MARGIN_L, table_y - 90000, CONTENT_W)
    rows = [("Роль в проекте", "Кто", "За что отвечает")]
    rows += PARTICIPANTS
    end_y = table_grid(slide, MARGIN_L, table_y + 60000, CONTENT_W,
                       [22, 30, 48], rows, row_h=286000, size=8.5)

    externals = ", ".join(sorted({
        metrics.short_owner(p[6]) for p in P.PROCESSES
        if not p[6].startswith("ДУП")
    }))
    textbox(
        slide, MARGIN_L, end_y + 150000, CONTENT_W, 300000,
        f"Смежные подразделения, с которыми согласуются стыки: {externals}, филиалы "
        "и госпартнёр.",
        size=8.5, color=STEEL, line_spacing=1.25,
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

    # Порядок доклада: предпосылки → сделанная работа → как устроена система
    # управления процессами → как считался приоритет → диагноз и цель →
    # очередь работ → что нужно от руководства → кто делает и когда.
    for builder in (
        build_premises,
        build_work_done,
        build_process_system,
        build_formulas,
        build_criteria,
        build_rules,
        build_as_is,
        build_to_be,
        build_headline,
        build_waves,
        build_matrix,
        build_queue,
        build_top_value,
        build_first_steps,
        build_asks,
        build_team_roadmap,
    ):
        builder(prs)

    move_slide_to_end(prs, 1)
    build_closing(prs)

    prs.save(str(OUTPUT))
    print(f"Сохранено: {OUTPUT}")
    print(f"Слайдов: {len(prs.slides._sldIdLst)}")


if __name__ == "__main__":
    main()

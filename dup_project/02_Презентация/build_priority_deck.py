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
    CONTENT_BOTTOM,
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
    TEAL_DARK,
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
from build_dup_role_map_simple import (  # noqa: E402
    simple_role,
    owner_text,
    ВЛАДЕЛЕЦ,
    УЧАСТНИК,
    КЛИЕНТ,
    НЕ_ОПРЕДЕЛЁН,
)

TEMPLATE = BASE.parent / "brandbook" / "Общие слайды_ver 12.12.24.pptx"
OUTPUT = BASE / "ДУП_Приоритеты_реинжиниринга_презентация.pptx"

DECK_TITLE = "Реинжиниринг процессов ДУП"
DECK_SUBTITLE = "От предпосылок до очереди работ"
DECK_DATE = "Сентябрь 2026"

RELS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

# Цвет волны на слайдах: та же логика, что у заливок в Excel.
WAVE_COLORS = {
    P.В1: GREEN,
    P.В2: YELLOW,
    P.В3: BLUE_LIGHT,
}


# --------------------------------------------------------------------------
# Данные
# --------------------------------------------------------------------------
ROWS = P.objects()
SCORED = [r for r in ROWS if r["total"] is not None]
BY_WAVE = Counter(r["wave"] for r in ROWS)
BY_QUADRANT = Counter(r["quadrant"] for r in SCORED)
if DECISIONS_AWAITING_DOCUMENT:
    raise SystemExit(
        "Доклад исходит из того, что стоп-факторов нет, но решения ждут документа: "
        f"{DECISIONS_AWAITING_DOCUMENT}"
    )

# Группы первой волны: набор типов вмешательства, подпись и пояснение.
# Размер волны — первое, о чём спрашивают, поэтому слайд показывает,
# из чего она складывается. Состав считается по данным, а не вписывается руками.
WAVE_1_GROUPS = [
    ({P.РЕГЛАМЕНТ, P.SLA, P.УПРОЩЕНИЕ}, "восстанавливаем исполнение внутри ДУП",
     "Чек-листы, регламенты и упрощение отчётности — без согласования со смежником"),
    ({P.ПАСПОРТ, P.ПРОЕКТ}, "описываем границы процессов",
     "Паспорта блоков 1, 2 и 4 и чек-лист передачи между реализацией и сопровождением"),
    ({P.БАЗА}, "снимаем базовую линию",
     "Пилотный аудит одного филиала и регулярный контроль качества данных"),
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
     "Она обязательна для всех работников (п. 4). Владение по должностям уже закреплено, "
     "но паспортов процессов нет и метрики не установлены."),
    ("Спорить о сроке приходится словами",
     "Фактические значения не измерялись. Ни доказать эффект изменения, ни назвать "
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

# Инструменты системы управления, которые проект реально применяет.
# Список собран по Методике и по файлам комплекта, а не по общему словарю.
PROJECT_TOOLS = [
    ("Реестр процессов", "Процесс, подпроцесс и этап в одном дереве"),
    ("Роль в процессе", "Владелец, участник или клиент у каждого объекта"),
    ("Паспорт процесса", "Границы, вход, выход, владелец, показатели"),
    ("SIPOC", "Поставщик, вход, шаги, выход, клиент"),
    ("Метрики", "Формула, единица и источник данных"),
    ("KPI", "Целевое значение и факт по объекту"),
    ("SLA", "Норматив срока на стыке с другим подразделением"),
    ("Базовая линия", "Замер «до» и аудит по чек-листам"),
    ("Зрелость", "Семь критериев Методики, шкала 1–5"),
    ("Очередь работ", "Балл, волна и порядок реинжиниринга"),
    ("Регламент и чек-лист", "Письменное правило, по которому работают"),
    ("Мониторинговый отчёт", "Проект, операция и портфель на одном контуре"),
]

# Норматив в методологии уже есть — возвращаем исполнение.
# Коды сверены с диагнозом «12 из 15» и с текущим реестром после схлопывания строк.
APPLY_EXISTING = [
    ("1.1", "Приём пакета от ДРБ", "Чек-лист минимального пакета и возврат неполного комплекта"),
    ("1.5", "ФЭМ и бюджет", "Ревизия нормативов модели: поверка, переносы, цены"),
    ("1.7", "Закупки", "Нормативы сроков по шагам и точка эскалации"),
    ("1.9", "Реестр платежей", "Инструкция трёх потоков ВУ / КИЗ / КОЗ"),
    ("2.3", "Прогноз поступлений", "Реестр и календарь вместо чата"),
    ("2.5", "Переносы оборудования", "Учёт «факт против ФЭМ»"),
    ("2.7", "Дефекты компонентов", "Маршрут из эксплуатации в ТЗ и закупки"),
    ("4.2", "Аудит данных", "Чек-листы методологии, контур остановлен с 2020"),
    ("4.3", "Статус и отчётность", "Один регламент: проект, операция, портфель"),
]

# Норматива нет — пишем правило и сразу начинаем по нему работать.
CREATE_NEW = [
    ("1.2", "Готовность продукта", "Gate «пилот → реализация»: критерии, подпись, право остановить старт"),
    ("2.0", "Передача в сопровождение", "Чек-лист handover: состав, дата, ответственный"),
    ("Б.5", "Открытие филиала", "Сквозной маршрут Юр → HR → Финблок → ДУП и сроки по этапам"),
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
     "Приоритеты, утверждение паспортов процессов, эскалации"),
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
    ("Волна 1", "Запуск", "месяцы 1–4",
     "Быстрые победы, паспорта и пилотный аудит идут параллельно"),
    ("Волна 2", "Стыки", "месяцы 3–8",
     "SLA со смежниками, gate продукта и маршрут открытия филиала"),
    ("Волна 3", "Масштаб", "месяцы 7–12",
     "Экономика поверки, единая карточка проекта и автоматизация"),
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
         "Подтверждение готовности продукта и проверка завершения пилота решаются "
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
    """Что такое система управления процессами и какими инструментами её ведём."""
    slide = new_slide(prs)
    title(slide, "Система управления процессами",
          "Цикл из Методики KT_METHOD v1 и инструменты, которыми этот цикл ведём в проекте")

    accent_bar(
        slide, MARGIN_L, CONTENT_TOP + 20000, CONTENT_W, 520000,
        PROCESS_DEFINITION, fill=BLUE, size=9.5, align=PP_ALIGN.LEFT,
    )

    label_y = CONTENT_TOP + 580000
    textbox(slide, MARGIN_L, label_y, CONTENT_W, 220000,
            "Цикл управления — пять этапов (Таблица 20 Методики)",
            size=10, font=FONT_MED, color=DARK)

    col_w = (CONTENT_W - 4 * 100000) // 5
    y = label_y + 240000
    for i, (name, action) in enumerate(PROCESS_CYCLE, 1):
        x = MARGIN_L + (i - 1) * (col_w + 100000)
        accent = GREEN if i == 4 else BLUE_LIGHT
        head = rect(slide, x, y, col_w, 280000, fill=accent, radius=0.12)
        tf = head.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_top = tf.margin_bottom = 0
        set_text(tf, f"{i}. {name}", size=9, font=FONT_BOLD, color=on(accent),
                 align=PP_ALIGN.CENTER)
        body = rect(slide, x, y + 300000, col_w, 520000, fill=CARD_BG)
        tf = body.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        set_text(tf, action, size=8, color=TEXT, line_spacing=1.15)

    tools_y = y + 900000
    textbox(slide, MARGIN_L, tools_y, CONTENT_W, 220000,
            "Инструменты, которые применяем",
            size=10, font=FONT_MED, color=DARK)

    gap = 80000
    tool_w = (CONTENT_W - 5 * gap) // 6
    tool_h = 620000
    base_y = tools_y + 250000
    for i, (name, note) in enumerate(PROJECT_TOOLS):
        col, row = i % 6, i // 6
        x = MARGIN_L + col * (tool_w + gap)
        ty = base_y + row * (tool_h + 70000)
        card = rect(slide, x, ty, tool_w, tool_h, fill=CARD_BG)
        tf = card.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = tf.margin_right = Emu(70000)
        tf.margin_top = Emu(60000)
        set_text(tf, name, size=9, font=FONT_BOLD, color=BLUE, line_spacing=1.05)
        para = tf.add_paragraph()
        para.space_before = Pt(3)
        para.line_spacing = 1.1
        run = para.add_run()
        run.text = note
        run.font.name = FONT
        run.font.size = Pt(8)
        run.font.color.rgb = TEXT


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
    """Как работали до того, как собрали систему управления процессами."""
    slide = new_slide(prs)
    title(slide, "Где мы были",
          "До системы управления: правила либо не исполнялись, либо их не было — "
          "это диагноз работы, а не оценка людей")

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


def build_as_is_now(prs: Presentation) -> None:
    """Что уже дала система управления и что в исполнении пока не изменилось."""
    slide = new_slide(prs)
    title(slide, "Где мы сейчас",
          "Систему управления собрали. Исполнение процессов ещё прежнее — "
          "очередь как раз про то, чтобы это изменить")

    roles = Counter(simple_role(p[0], p[8], p[6]) for p in P.PROCESSES)
    left_w = (CONTENT_W - 180000) // 2
    y = CONTENT_TOP + 80000
    card_h = 2480000

    left = rect(slide, MARGIN_L, y, left_w, card_h, fill=CARD_BG)
    textbox(slide, MARGIN_L + 160000, y + 120000, left_w - 320000, 280000,
            "Система уже работает", size=13, font=FONT_BOLD, color=ink(GREEN))
    applied = [
        f"{objects_word(len(P.PROCESSES))} сведены в один реестр: процесс, подпроцесс, этап",
        f"Роль ДУП названа у каждого: {roles[ВЛАДЕЛЕЦ]} — владелец, "
        f"{roles[УЧАСТНИК]} — участник, {roles[КЛИЕНТ]} — клиент. "
        f"Неназначенных — {roles[НЕ_ОПРЕДЕЛЁН]}",
        f"{CRITERIA_COUNT} критериев оценки расписаны формулой и источником данных",
        f"Очередь из {len([w for w in P.WAVES if w[0] != P.ВН])} волн: "
        "место объекта задаёт балл, паспорт очередь не держит",
    ]
    bullets(left, applied, size=10, prefix="— ", spacing=10, line_spacing=1.25)
    left.text_frame.margin_top = Emu(480000)
    left.text_frame.margin_left = Emu(140000)
    left.text_frame.margin_right = Emu(140000)

    right_x = MARGIN_L + left_w + 180000
    right = rect(slide, right_x, y, left_w, card_h, fill=CARD_BG)
    textbox(slide, right_x + 160000, y + 120000, left_w - 320000, 280000,
            "В работе процессов пока ничего не сдвинулось",
            size=13, font=FONT_BOLD, color=RED)
    still = [
        f"{as_is_count('не исполняется')} объектов по-прежнему не исполняются так, как написано",
        f"{as_is_count('пробел')} процедуры по-прежнему нет: готовность продукта",
        "Ни один критерий ещё не измеряется — факта «до» нет",
        f"Зрелость каждого сквозного процесса — {MATURITY_NOW} из 5. "
        "Паспорта блоков 1, 2 и 4 не утверждены",
    ]
    bullets(right, still, size=10, prefix="— ", spacing=10, line_spacing=1.25)
    right.text_frame.margin_top = Emu(480000)
    right.text_frame.margin_left = Emu(140000)
    right.text_frame.margin_right = Emu(140000)

    accent_bar(
        slide, MARGIN_L, y + card_h + 160000, CONTENT_W, 620000,
        ["Система ответила на вопросы «кто владелец», «чем измеряем» и «что берём первым».",
         "Срок, деньги и качество сдвинутся, когда заработают регламенты волны запуска, "
         "а не в момент, когда реестр собран."],
        fill=BLUE, size=11, align=PP_ALIGN.LEFT,
    )


def build_proposal(prs: Presentation) -> None:
    """Что делаем с методологией: исполняем написанное или пишем недостающее правило."""
    slide = new_slide(prs)
    title(slide, "Что предлагаем",
          "Норматив есть — возвращаем в работу. Норматива нет — пишем правило и применяем")

    gap = 160000
    col_w = (CONTENT_W - gap) // 2
    y = CONTENT_TOP + 140000
    head_h = 360000
    body_h = 2100000

    left_head = rect(slide, MARGIN_L, y, col_w, head_h, fill=GREEN, radius=0.08)
    tf = left_head.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(tf, f"Вернуть в работу — {len(APPLY_EXISTING)} объектов",
             size=12, font=FONT_BOLD, color=WHITE, align=PP_ALIGN.CENTER)
    left = rect(slide, MARGIN_L, y + head_h, col_w, body_h, fill=CARD_BG)
    lines = [f"{code}  {name}. {action}" for code, name, action in APPLY_EXISTING]
    bullets(left, lines, size=8.5, prefix="", spacing=3, line_spacing=1.12)
    left.text_frame.margin_top = Emu(80000)
    left.text_frame.margin_left = Emu(120000)
    left.text_frame.margin_right = Emu(100000)

    rx = MARGIN_L + col_w + gap
    right_head = rect(slide, rx, y, col_w, head_h, fill=YELLOW, radius=0.08)
    tf = right_head.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(tf, "Написать правило и начать применять",
             size=12, font=FONT_BOLD, color=DARK, align=PP_ALIGN.CENTER)
    right = rect(slide, rx, y + head_h, col_w, body_h, fill=CARD_BG)
    tf = right.text_frame
    tf.word_wrap = True
    tf.margin_top = Emu(100000)
    tf.margin_left = Emu(140000)
    tf.margin_right = Emu(120000)
    intro = tf.paragraphs[0]
    intro.line_spacing = 1.15
    run = intro.add_run()
    run.text = "В методологии этих процедур нет. Пока их не напишем, каждый случай решается заново."
    run.font.name = FONT
    run.font.size = Pt(10)
    run.font.color.rgb = TEXT
    for code, name, action in CREATE_NEW:
        para = tf.add_paragraph()
        para.space_before = Pt(8)
        para.line_spacing = 1.15
        head = para.add_run()
        head.text = f"{code}  {name}"
        head.font.name = FONT_BOLD
        head.font.size = Pt(11)
        head.font.color.rgb = DARK
        body = tf.add_paragraph()
        body.line_spacing = 1.1
        detail = body.add_run()
        detail.text = action
        detail.font.name = FONT
        detail.font.size = Pt(10)
        detail.font.color.rgb = GRAY

    foot_y = y + head_h + body_h + 120000
    foot_w = (CONTENT_W - gap) // 2
    for x, heading, body in (
        (MARGIN_L, "Описать то, чем владеем",
         "Паспорта блоков 1, 2 и 4, чек-лист передачи и положение о контроле качества данных"),
        (MARGIN_L + foot_w + gap, "Оцифровать то, что уже исполняется",
         "Единая карточка и автопрогноз — после регламента, иначе в системе закрепится беспорядок"),
    ):
        box = rect(slide, x, foot_y, foot_w, 520000, fill=WHITE, line=MIST)
        tf = box.text_frame
        tf.margin_left = tf.margin_right = Emu(120000)
        tf.margin_top = Emu(60000)
        set_text(tf, heading, size=11, font=FONT_BOLD, color=BLUE)
        para = tf.add_paragraph()
        para.space_before = Pt(4)
        para.line_spacing = 1.1
        run = para.add_run()
        run.text = body
        run.font.name = FONT
        run.font.size = Pt(9)
        run.font.color.rgb = TEXT


def build_outcomes(prs: Presentation) -> None:
    """Два контура эффекта: управляемость системы и результат для компании."""
    slide = new_slide(prs)
    title(slide, "Что получит компания",
          "Сначала процессом можно управлять. Затем это видно в сроке, деньгах и прогнозе")

    gap = 160000
    col_w = (CONTENT_W - gap) // 2
    y = CONTENT_TOP + 20000
    head_h = 320000

    manage = [
        ("Владелец и паспорт", "Границы, вход, выход и один ответственный"),
        ("Прозрачность", "Показатель с источником данных, а не мнение"),
        ("Мониторинговый отчёт", "Каждый месяц: проект, операция, портфель"),
        ("Одинаковый старт", "Старт по чек-листу, а не по памяти"),
    ]
    business = [
        ("Срок старта проекта", "Сегодня не измеряется", "−30% к базовой линии"),
        ("Полный пакет на старте", "Не измеряется", "≥ 90% проектов"),
        ("Оплаты в срок", "По обстоятельствам", "≥ 85% заявок"),
        ("Прогноз поступлений", "Чаты и разовые запросы", "100% по календарю"),
        ("Аудит портфеля", "0% с 2020 года", "100% за год"),
    ]

    left_head = rect(slide, MARGIN_L, y, col_w, head_h, fill=BLUE, radius=0.08)
    tf = left_head.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(tf, "Управляемость и прозрачность",
             size=13, font=FONT_BOLD, color=WHITE, align=PP_ALIGN.CENTER)
    row_y = y + head_h + 80000
    for heading, body in manage:
        textbox(slide, MARGIN_L, row_y, col_w - 60000, 220000, heading,
                size=12, font=FONT_BOLD, color=DARK)
        textbox(slide, MARGIN_L, row_y + 240000, col_w - 80000, 240000, body,
                size=10, color=GRAY, line_spacing=1.05)
        row_y += 520000

    rx = MARGIN_L + col_w + gap
    right_head = rect(slide, rx, y, col_w, head_h, fill=GREEN, radius=0.08)
    tf = right_head.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(tf, "Срок, деньги, регулярность",
             size=13, font=FONT_BOLD, color=WHITE, align=PP_ALIGN.CENTER)
    rows = [("Показатель", "Сейчас", "За 12 месяцев")] + [
        (name, now, target) for name, now, target in business
    ]
    table_grid(slide, rx, y + head_h + 80000, col_w, [40, 32, 28], rows,
               row_h=300000, size=9)

    box = rect(slide, MARGIN_L, CONTENT_TOP + 2480000, CONTENT_W, 1120000,
               fill=TEAL_DARK, radius=0.1)
    tf = box.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Emu(160000)
    tf.margin_top = tf.margin_bottom = Emu(70000)
    set_text(
        tf,
        [f"Зрелость с {MATURITY_NOW} до {MATURITY_GOAL}: портфель больше не держится на одном человеке.",
         "На уровне 1 уход заместителя обнуляет срок, пакет и прогноз: срыв виден, когда его уже принёс госпартнёр.",
         "На уровне 3 отклонение денег и срока видно в ежемесячном отчёте, пока его можно поправить, "
         "а смежник отвечает по подписанному SLA."],
        size=12, font=FONT_BOLD, color=WHITE, line_spacing=1.08,
    )


def build_headline(prs: Presentation) -> None:
    """Слайд решения: одна мысль и четыре цифры под ней."""
    slide = new_slide(prs)
    title(slide, "Очередь реинжиниринга построена")

    accent_bar(
        slide, MARGIN_L, CONTENT_TOP - 30000, CONTENT_W, 700000,
        ["Первыми берём объекты с высшим баллом, которые ДУП начинает сам.",
         "Паспорт и замер «до» идут в той же волне и очередь не держат. "
         "Договорённости со смежниками — следующая волна."],
        fill=BLUE, size=12, align=PP_ALIGN.LEFT,
    )

    stats = [
        (str(len(SCORED)),
         f"{plural(len(SCORED), 'объект', 'объекта', 'объектов')} с баллом "
         "поставлены в очередь", GREEN),
        (str(BY_WAVE[P.В1]),
         f"{plural(BY_WAVE[P.В1], 'объект', 'объекта', 'объектов')} волны 1: паспорта, "
         "замер и быстрые победы внутри ДУП", GREEN),
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
        f"Источник: реестр из {objects_word(len(P.PROCESSES))}; "
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
          "Волна 1 идёт по баллу. Дальше место в очереди задают пять правил")

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
    title(slide, "Три волны работ",
          "Волна 1 — по убыванию балла; волны 2 и 3 ждут договорённости со смежником")

    waves = [w for w in P.WAVES if w[0] != P.ВН]
    gap = 140000
    col_w = (CONTENT_W - (len(waves) - 1) * gap) // len(waves)
    y = CONTENT_TOP + 120000
    for i, (wave, phase, goal, *_rest) in enumerate(waves):
        accent = WAVE_COLORS[wave]
        x = MARGIN_L + i * (col_w + gap)

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
        accent = BLUE
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
          "Волна запуска, строго по убыванию балла")

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
    title(slide, "Высший балл начинается сразу",
          "Волна 2 держит только то, что нельзя решить внутри ДУП")

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

    accent_bar(
        slide, MARGIN_L, end_y + 300000, CONTENT_W, 620000,
        ["Приём пакета от ДРБ, платежи и заявка на закупку стоят в начале очереди.",
         "Закупки целиком, ФЭМ и gate готовности продукта ждут волну 2: "
         "норматив ставит другое подразделение, а не паспорт процесса."],
        fill=GREEN, size=10, align=PP_ALIGN.LEFT,
    )


def build_first_steps(prs: Presentation) -> None:
    slide = new_slide(prs)
    shown = 5
    title(slide, "Программа старта: пять первых действий",
          "Шаги волны 1 стоят по убыванию балла и паспорта не ждут")

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
        f"Шаги {shown + 1}–{len(P.FIRST_STEPS)} — остальная волна 1 по убыванию балла: "
        "реестр переносов, паспорта блоков 1 и 4, чек-лист передачи, пилотный аудит "
        "и матрица кураторов.",
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

    n = len(ROADMAP)
    gap = 140000
    col_w = (CONTENT_W - (n - 1) * gap) // n
    y = CONTENT_TOP + 60000
    for i, (phase, name, term, result) in enumerate(ROADMAP):
        x = MARGIN_L + i * (col_w + gap)
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


def role_counter():
    return Counter(simple_role(p[0], p[8], p[6]) for p in P.PROCESSES)


def level_counter():
    return Counter(p[1] for p in P.PROCESSES)


def processes_by_role():
    """Имена процессов верхнего уровня в порядке дерева, сгруппированные по роли ДУП."""
    order = {code: i for i, code in enumerate(P.TREE_ORDER)}
    groups = {ВЛАДЕЛЕЦ: [], УЧАСТНИК: [], КЛИЕНТ: []}
    for p in P.PROCESSES:
        if p[1] != "Процесс":
            continue
        role = simple_role(p[0], p[8], p[6])
        groups[role].append((order.get(p[0], 99), p[2]))
    return {role: [name for _, name in sorted(items)] for role, items in groups.items()}


def build_perimeter(prs: Presentation) -> None:
    """Периметр — процессы, а не строки реестра. Чужие процессы показаны с ролью ДУП."""
    slide = new_slide(prs)
    title(slide, "Периметр проекта",
          "Управляем процессами. Строки реестра — это их подпроцессы и этапы")

    levels = level_counter()
    roles = role_counter()
    groups = processes_by_role()
    owned = groups[ВЛАДЕЛЕЦ]
    participants = groups[УЧАСТНИК]
    clients = groups[КЛИЕНТ]
    n_rest = len(participants) + len(clients)

    gap = 140000
    col_w = (CONTENT_W - 2 * gap) // 3
    y = CONTENT_TOP + 10000
    card_h = 640000
    stats = [
        (str(levels["Процесс"]), "процессов в периметре", GREEN),
        (str(len(owned)), "ДУП — владелец результата", GREEN),
        (str(n_rest), "ДУП — участник или клиент", BLUE),
    ]
    for i, (number, caption, accent) in enumerate(stats):
        x = MARGIN_L + i * (col_w + gap)
        rect(slide, x, y, col_w, card_h, fill=CARD_BG)
        textbox(slide, x + 120000, y + 40000, col_w - 240000, 360000, number,
                size=28, font=FONT_BOLD, color=accent, line_spacing=1.0)
        textbox(slide, x + 120000, y + 400000, col_w - 240000, 200000, caption,
                size=12, font=FONT_MED, color=DARK, line_spacing=1.05)

    panel_y = y + card_h + 80000
    panel_h = CONTENT_BOTTOM - panel_y - 480000
    left_w = 2680000
    left = rect(slide, MARGIN_L, panel_y, left_w, panel_h, fill=CARD_BG)
    rect(slide, MARGIN_L, panel_y, left_w, 70000, fill=GREEN, shape=MSO_SHAPE.RECTANGLE)
    textbox(slide, MARGIN_L + 120000, panel_y + 90000, left_w - 240000, 220000,
            "Владелец", size=13, font=FONT_BOLD, color=ink(GREEN))
    name_y = panel_y + 320000
    name_h = (panel_h - 360000) // max(len(owned), 1)
    for i, name in enumerate(owned):
        textbox(slide, MARGIN_L + 140000, name_y + i * name_h, left_w - 280000, name_h - 40000,
                name, size=11, color=DARK, line_spacing=1.0)

    right_x = MARGIN_L + left_w + gap
    right_w = CONTENT_W - left_w - gap
    inner_gap = 100000
    inner_w = (right_w - inner_gap) // 2
    columns = [
        (participants, "Участник", BLUE),
        (clients, "Клиент", YELLOW),
    ]
    for col_i, (names, label, accent) in enumerate(columns):
        x = right_x + col_i * (inner_w + inner_gap)
        rect(slide, x, panel_y, inner_w, panel_h, fill=CARD_BG)
        rect(slide, x, panel_y, inner_w, 70000, fill=accent, shape=MSO_SHAPE.RECTANGLE)
        textbox(slide, x + 100000, panel_y + 90000, inner_w - 200000, 220000,
                label, size=13, font=FONT_BOLD, color=ink(accent))
        row_h = (panel_h - 360000) // max(len(names), 1)
        for i, name in enumerate(names):
            textbox(slide, x + 100000, name_y + i * row_h, inner_w - 200000, row_h - 36000,
                    name, size=10, color=DARK, line_spacing=1.0)

    textbox(
        slide, MARGIN_L, CONTENT_BOTTOM - 400000, CONTENT_W, 360000,
        f"В реестре {objects_word(len(P.PROCESSES))}: "
        f"{levels['Подпроцесс']} {plural(levels['Подпроцесс'], 'подпроцесс', 'подпроцесса', 'подпроцессов')} "
        f"и {levels['Этап']} {plural(levels['Этап'], 'этап', 'этапа', 'этапов')}. "
        f"Роль ДУП по этим строкам: {roles[ВЛАДЕЛЕЦ]} — владелец, "
        f"{roles[УЧАСТНИК]} — участник, {roles[КЛИЕНТ]} — клиент.",
        size=11, color=GRAY, line_spacing=1.15,
    )


def wave_short(wave: str) -> str:
    if wave == P.АГ:
        return "—"
    if wave == P.ВН:
        return "вне"
    return wave_no(wave)


# Бизнес-владелец — та же строка, что в карте ролей, а не сокращение подразделения.
BUSINESS_OWNER = {
    p[0]: owner_text(p[8], p[6], p[7]) for p in P.PROCESSES
}
_ROLE_WORD = {
    ВЛАДЕЛЕЦ: "владелец",
    УЧАСТНИК: "участник",
    КЛИЕНТ: "клиент",
    НЕ_ОПРЕДЕЛЁН: "",
}


def level_role(level: str, role: str, target: str) -> str:
    """Роль ДУП пишется только в столбце уровня этой строки."""
    if level != target:
        return ""
    return _ROLE_WORD.get(role, role.lower())


def build_catalog(prs: Presentation) -> None:
    """Перечень: объект, бизнес-владелец, роль ДУП по уровню, действие, волна и балл.

    Роль ДУП стоит в одном столбце — «Процесс», «Подпроцесс» или «Этап» —
    по уровню строки. Текст действия и имя владельца не обрезаются: строка
    рассчитана на три строки при 7 pt.
    """
    page_size = 8
    items = []
    for r in P.tree_rows(ROWS):
        items.append((
            "—" if r["rank"] is None else str(r["rank"]),
            r["name"],
            BUSINESS_OWNER[r["code"]],
            level_role(r["level"], r["role"], "Процесс"),
            level_role(r["level"], r["role"], "Подпроцесс"),
            level_role(r["level"], r["role"], "Этап"),
            r["first"],
            wave_short(r["wave"]),
            "—" if r["total"] is None else f'{r["total"]:.0f}',
        ))
    pages = [items[i:i + page_size] for i in range(0, len(items), page_size)]
    for index, page in enumerate(pages, 1):
        slide = new_slide(prs)
        title(
            slide,
            "Перечень процессов",
            f"Все {len(items)} объектов реестра по процессам. "
            f"Лист {index} из {len(pages)}",
        )
        rows = [(
            "№", "Объект", "Бизнес-владелец", "Процесс", "Подпроцесс", "Этап",
            "Что делаем", "Волна", "Балл",
        )] + page
        table_grid(
            slide, MARGIN_L, CONTENT_TOP + 20000, CONTENT_W,
            [3, 17, 20, 7, 9, 7, 26, 6, 5], rows,
            row_h=370000, size=7, margin_x=45000, margin_y=12000,
        )


def build_gantt(prs: Presentation) -> None:
    """Дорожная карта на 12 месяцев: три волны с перекрытием."""
    slide = new_slide(prs)
    title(slide, "Дорожная карта",
          "12 месяцев. Волны перекрываются: стыки начинаются, пока ещё закрывается запуск")

    months = 12
    label_w = 1680000
    chart_x = MARGIN_L + label_w
    chart_w = CONTENT_W - label_w
    month_w = chart_w // months
    head_y = CONTENT_TOP + 40000
    textbox(slide, MARGIN_L, head_y, label_w - 80000, 220000, "Волна",
            size=9, font=FONT_BOLD, color=GRAY)
    for m in range(months):
        textbox(slide, chart_x + m * month_w, head_y, month_w, 220000, str(m + 1),
                size=9, font=FONT_MED, color=GRAY, align=PP_ALIGN.CENTER)

    # start — номер месяца с нуля, span — длина в месяцах
    bars = [
        (P.В1, GREEN, 0, 4, "Победы, паспорта, аудит"),
        (P.В2, YELLOW, 2, 6, "SLA, gate, маршрут филиала"),
        (P.В3, BLUE_LIGHT, 6, 6, "Экономика, карточка, автоматизация"),
    ]
    row_h = 460000
    row_gap = 80000
    y = head_y + 280000
    for wave, accent, start, span, caption in bars:
        rect(slide, MARGIN_L, y, CONTENT_W, row_h, fill=CARD_BG, shape=MSO_SHAPE.RECTANGLE)
        textbox(slide, MARGIN_L + 80000, y + 60000, label_w - 160000, 180000,
                f"Волна {wave_no(wave)}", size=12, font=FONT_BOLD, color=DARK)
        textbox(slide, MARGIN_L + 80000, y + 240000, label_w - 160000, 180000,
                f"мес. {start + 1}–{start + span}", size=10, color=GRAY)
        for m in range(months):
            divider(slide, chart_x + m * month_w, y, 8000, color=MIST, thickness=row_h / 9525)
        bar = rect(
            slide,
            chart_x + start * month_w + 20000,
            y + 70000,
            span * month_w - 40000,
            row_h - 140000,
            fill=accent,
            radius=0.15,
        )
        tf = bar.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Emu(80000)
        set_text(tf, caption, size=10, font=FONT_BOLD, color=on(accent), align=PP_ALIGN.CENTER)
        y += row_h + row_gap

    textbox(slide, MARGIN_L, y + 40000, CONTENT_W, 200000,
            "Месяцы по горизонтали. Ключевые результаты на стыках волн",
            size=11, font=FONT_BOLD, color=DARK)
    marks = [
        ("1", "Старт: чек-лист приёмки от ДРБ и инструкция платежей"),
        ("4", "Паспорта блоков 1, 2 и 4, восемь регламентов работают"),
        ("8", "Подписаны SLA, утверждены gate и маршрут открытия филиала"),
        ("12", "Зрелость 3, мониторинговый отчёт и карточка проекта"),
    ]
    mark_w = (CONTENT_W - 3 * 100000) // 4
    my = y + 280000
    for i, (month, text) in enumerate(marks):
        x = MARGIN_L + i * (mark_w + 100000)
        badge = rect(slide, x, my, 360000, 280000, fill=BLUE, radius=0.2)
        tf = badge.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = 0
        set_text(tf, month, size=12, font=FONT_BOLD, color=WHITE, align=PP_ALIGN.CENTER)
        textbox(slide, x + 400000, my, mark_w - 420000, 320000, text,
                size=9, color=TEXT, line_spacing=1.1)


def build_closing(prs: Presentation) -> None:
    slide = prs.slides[-1]
    msg = find_shape(slide, "Присоединяйтесь")
    if msg is not None:
        replace_text(
            msg,
            "Начинаем с высшего балла внутри ДУП — паспорт очередь не держит",
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

    # Периметр и полный перечень — в начале, чтобы масштаб был виден
    # до рассказа о том, зачем взялись за проект.
    for builder in (
        build_perimeter,
        build_catalog,
        build_premises,
        build_work_done,
        build_process_system,
        build_formulas,
        build_criteria,
        build_rules,
        build_as_is,
        build_as_is_now,
        build_proposal,
        build_outcomes,
        build_headline,
        build_waves,
        build_matrix,
        build_queue,
        build_top_value,
        build_first_steps,
        build_asks,
        build_team_roadmap,
        build_gantt,
    ):
        builder(prs)

    move_slide_to_end(prs, 1)
    build_closing(prs)

    prs.save(str(OUTPUT))
    print(f"Сохранено: {OUTPUT}")
    print(f"Слайдов: {len(prs.slides._sldIdLst)}")


if __name__ == "__main__":
    main()

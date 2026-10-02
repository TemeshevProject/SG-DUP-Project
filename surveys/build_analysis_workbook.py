#!/usr/bin/env python3
"""Build analysis workbook for branch director survey responses."""

import importlib.util
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SURVEYS_DIR = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("build_survey", SURVEYS_DIR / "build_survey.py")
build_survey = importlib.util.module_from_spec(spec)
spec.loader.exec_module(build_survey)
BLOCKS = build_survey.BLOCKS

OUTPUT = SURVEYS_DIR / "DUP_analiz_otvetov_oprosnika.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
BLOCK_FILL = PatternFill("solid", fgColor="D6E4F0")
WARN_FILL = PatternFill("solid", fgColor="FFF4D9")
GOOD_FILL = PatternFill("solid", fgColor="E8F5E9")
THIN = Side(style="thin", color="B4B4B4")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")


def flat_questions():
    items = []
    for _, _, questions in BLOCKS:
        for q in questions:
            if q["type"] == "матрица":
                for sub in q["options"]:
                    items.append(
                        {
                            "id": f"{q['id']}_{sub[:15].replace(' ', '_')}",
                            "parent_id": q["id"],
                            "text": f"{q['text']} — {sub}",
                            "type": "шкала 1–5",
                            "block": q.get("block"),
                        }
                    )
            else:
                items.append(
                    {
                        "id": q["id"],
                        "parent_id": q["id"],
                        "text": q["text"],
                        "type": q["type"],
                        "block": None,
                    }
                )
    return items


SCALE_IDS = [
    "B1", "B3", "B6", "C1", "C3", "C6", "D2", "D6", "E1", "E2", "E3",
    "F3", "F4", "G2",
]
F1_IDS = [
    ("ДРБ", "F1_ДРБ"),
    ("Финблок", "F1_Финблок"),
    ("Юр блок", "F1_Юр_блок"),
    ("Закупки (ОЛЖА)", "F1_Закупки_(ОЛЖА)"),
    ("СТ (тех. блок)", "F1_СТ_(тех._блок)"),
    ("Зерек / Zerek Road tech", "F1_Зерек_/_Zerek_R"),
    ("HR", "F1_HR"),
    ("Лаборатория (поверка)", "F1_Лаборатория_(по"),
]
OPEN_IDS = ["B11", "C8", "D8", "E7", "E8", "F5", "H1", "H2", "H3"]


def style_header(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def build_instruction_sheet(wb):
    ws = wb.active
    ws.title = "Инструкция"
    lines = [
        ("Анализ ответов опросника директоров филиалов — ДУП", True, 14),
        ("", False, 11),
        ("1. Сбор ответов в Google Forms", True, 12),
        ("   • Создайте форму по файлу DUP_oprosnik_Google_Forms_инструкция.md", False, 11),
        ("   • В Google Forms: Ответы → Создать таблицу (Google Таблицы) ИЛИ скачать CSV (.csv)", False, 11),
        ("", False, 11),
        ("2. Импорт в этот файл", True, 12),
        ("   • Вариант A: Скопируйте все строки из Google Таблицы на лист «Сырые ответы» (начиная со строки 2, без заголовков — они уже есть)", False, 11),
        ("   • Вариант B: Запустите: python3 import_survey_responses.py путь/к/ответы.csv", False, 11),
        ("   • Вариант C: В Excel — Данные → Из текста/CSV → лист «Сырые ответы»", False, 11),
        ("", False, 11),
        ("3. Автообновление сводки", True, 12),
        ("   • Листы «Дашборд», «Шкалы», «Взаимодействие ЦО», «ТОП боли», «Интервью» обновляются по формулам при заполнении «Сырые ответы»", False, 11),
        ("   • Красная зона: средняя оценка по шкале ≤ 2,5", False, 11),
        ("", False, 11),
        ("4. Столбцы на листе «Сырые ответы»", True, 12),
        ("   • A–D: метаданные (дата, филиал, должность, ФИО)", False, 11),
        ("   • E+: ID вопроса в заголовке — вставляйте ответы в соответствующие столбцы", False, 11),
        ("   • Для шкал указывайте число 1–5", False, 11),
    ]
    for r, (text, bold, size) in enumerate(lines, 1):
        cell = ws.cell(row=r, column=1, value=text)
        cell.font = Font(bold=bold, size=size, color="1F4E79" if bold and size > 11 else "000000")
        cell.alignment = WRAP
    ws.column_dimensions["A"].width = 95


def build_raw_sheet(wb, questions):
    ws = wb.create_sheet("Сырые ответы")
    meta = ["Дата ответа", "Регион/филиал", "Должность", "ФИО (опц.)"]
    headers = meta + [f"{q['id']}: {q['text'][:50]}" for q in questions]
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, len(headers))
    ws.freeze_panes = "E2"
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16 if col <= 4 else 14
    # sample row hint
    ws.cell(row=2, column=1, value="(удалите пример после импорта)")
    return ws, len(meta) + 1


def build_mapping_sheet(wb, questions):
    ws = wb.create_sheet("Карта вопросов")
    headers = ["ID", "Текст вопроса", "Тип", "Столбец в «Сырые ответы»", "Участвует в сводке"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, len(headers))
    col_start = 5  # column E in raw sheet
    for r, q in enumerate(questions, 2):
        participates = "Шкала" if q["id"] in SCALE_IDS or q["id"].startswith("F1_") else (
            "Открытый" if q["id"] in OPEN_IDS else "Распределение" if q["type"] in ("один из списка", "да/нет") else "—"
        )
        ws.cell(row=r, column=1, value=q["id"])
        ws.cell(row=r, column=2, value=q["text"]).alignment = WRAP
        ws.cell(row=r, column=3, value=q["type"])
        ws.cell(row=r, column=4, value=get_column_letter(col_start + r - 2))
        ws.cell(row=r, column=5, value=participates)
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = BORDER
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 16


def col_letter_for_id(questions, qid, meta_cols=4):
    for i, q in enumerate(questions):
        if q["id"] == qid:
            return get_column_letter(meta_cols + 1 + i)
    return None


def build_dashboard(wb, questions):
    ws = wb.create_sheet("Дашборд")
    ws["A1"] = "Сводный дашборд — опрос директоров филиалов"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws.merge_cells("A1:D1")

    raw = "'Сырые ответы'"
    # count responses (non-empty region column B)
    ws["A3"] = "Количество ответов"
    ws["B3"] = f'=COUNTA({raw}!B:B)-1'
    ws["A4"] = "Дата последнего обновления"
    ws["B4"] = "=TODAY()"
    ws["B4"].number_format = "DD.MM.YYYY"

    ws["A6"] = "Показатель"
    ws["B6"] = "Среднее"
    ws["C6"] = "Зона"
    ws["D6"] = "Комментарий"
    style_header(ws, 6, 4)

    scale_labels = {
        "B1": "Понятность процесса реализации",
        "B3": "Полнота пакета документов при старте",
        "B6": "Согласование бюджета с Финблоком",
        "C1": "Понятность сопровождения",
        "C3": "Скорость реакции ДУП на проблемы",
        "C6": "Обратная связь по дефектам компонентов",
        "D2": "Понятность приоритизации платежей",
        "D6": "Удобство прогноза поступлений/расходов",
        "E1": "Разграничение ответственности филиал/ДУП",
        "E2": "Понятность, к кому обращаться в ДУП",
        "E3": "Качество коммуникации с куратором",
        "F3": "Своевременность передачи от ДРБ",
        "F4": "Оперативность Юр блока",
        "G2": "Работоспособность платформы проектов",
    }

    row = 7
    for qid, label in scale_labels.items():
        col = col_letter_for_id(questions, qid)
        if not col:
            continue
        ws.cell(row=row, column=1, value=label)
        avg_formula = f'=IFERROR(AVERAGE({raw}!{col}2:{col}500),"")'
        ws.cell(row=row, column=2, value=avg_formula)
        ws.cell(row=row, column=2).number_format = "0.00"
        zone = f'=IF(B{row}="","—",IF(B{row}<=2.5,"🔴 Критично",IF(B{row}<=3.5,"🟡 Внимание","🟢 Норма")))'
        ws.cell(row=row, column=3, value=zone)
        ws.cell(row=row, column=4, value=f"Вопрос {qid}")
        for c in range(1, 5):
            ws.cell(row=row, column=c).border = BORDER
        row += 1

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 14

    # conditional formatting on averages
    ws.conditional_formatting.add(
        f"B7:B{row-1}",
        ColorScaleRule(start_type="num", start_value=1, start_color="F8696B",
                       mid_type="num", mid_value=3, mid_color="FFEB84",
                       end_type="num", end_value=5, end_color="63BE7B"),
    )


def build_scales_sheet(wb, questions):
    ws = wb.create_sheet("Шкалы")
    ws["A1"] = "Детализация по шкалам 1–5 (распределение ответов)"
    ws["A1"].font = Font(bold=True, size=12, color="1F4E79")
    headers = ["ID", "Вопрос", "1", "2", "3", "4", "5", "Среднее", "N"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 3, len(headers))

    raw = "'Сырые ответы'"
    row = 4
    for qid in SCALE_IDS:
        col = col_letter_for_id(questions, qid)
        if not col:
            continue
        qtext = next((q["text"] for q in questions if q["id"] == qid), qid)
        ws.cell(row=row, column=1, value=qid)
        ws.cell(row=row, column=2, value=qtext).alignment = WRAP
        for score, c in enumerate(range(3, 8), 1):
            ws.cell(row=row, column=c, value=f'=COUNTIF({raw}!{col}2:{col}500,{score})')
        ws.cell(row=row, column=8, value=f'=IFERROR(AVERAGE({raw}!{col}2:{col}500),"")')
        ws.cell(row=row, column=9, value=f'=COUNT({raw}!{col}2:{col}500)')
        row += 1

    for c in range(1, 10):
        ws.column_dimensions[get_column_letter(c)].width = 22 if c == 2 else 10


def build_co_sheet(wb, questions):
    ws = wb.create_sheet("Взаимодействие ЦО")
    ws["A1"] = "Оценки взаимодействия с подразделениями ЦО (вопрос F1)"
    ws["A1"].font = Font(bold=True, size=12, color="1F4E79")
    headers = ["Подразделение", "Средняя оценка", "N ответов", "Зона"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 3, 4)

    raw = "'Сырые ответы'"
    row = 4
    for sub, qid in F1_IDS:
        col = col_letter_for_id(questions, qid)
        ws.cell(row=row, column=1, value=sub)
        if col:
            ws.cell(row=row, column=2, value=f'=IFERROR(AVERAGE({raw}!{col}2:{col}500),"")')
            ws.cell(row=row, column=3, value=f'=COUNT({raw}!{col}2:{col}500)')
            ws.cell(row=row, column=4, value=f'=IF(B{row}="","—",IF(B{row}<=2.5,"🔴",IF(B{row}<=3.5,"🟡","🟢")))')
        row += 1

    chart = BarChart()
    chart.type = "bar"
    chart.title = "Средние оценки взаимодействия с ЦО"
    chart.y_axis.title = "Подразделение"
    chart.x_axis.title = "Средняя оценка (1–5)"
    data = Reference(ws, min_col=2, min_row=3, max_row=row - 1)
    cats = Reference(ws, min_col=1, min_row=4, max_row=row - 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.width = 16
    chart.height = 10
    ws.add_chart(chart, "F3")

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 14


def build_pain_sheet(wb, questions):
    ws = wb.create_sheet("ТОП боли")
    ws["A1"] = "Открытые ответы — боли, предложения, блокеры (для ручной кластеризации)"
    ws["A1"].font = Font(bold=True, size=12, color="1F4E79")
    headers = ["Филиал", "ID вопроса", "Тема", "Цитата / ответ", "Категория (заполнить)", "Приоритет"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 3, len(headers))

    themes = {
        "B11": "Блокеры реализации",
        "C8": "Боль сопровождения",
        "D8": "Перерасход бюджета",
        "E7": "Открытие филиала",
        "E8": "Улучшение взаимодействия с ДУП",
        "F5": "Проблемы с ЦО",
        "H1": "ТОП-3 проблемы",
        "H2": "ТОП-3 автоматизации",
        "H3": "Одно главное изменение",
    }

    raw = "'Сырые ответы'"
    row = 4
    for qid, theme in themes.items():
        col = col_letter_for_id(questions, qid)
        if not col:
            continue
        # Pull up to 30 non-empty answers via formula row references - use dynamic approach:
        # For each data row 2..50 in raw, if answer not empty, show on pain sheet
        for data_row in range(2, 52):
            ws.cell(row=row, column=1, value=f"={raw}!B{data_row}")
            ws.cell(row=row, column=2, value=qid)
            ws.cell(row=row, column=3, value=theme)
            ws.cell(row=row, column=4, value=f'=IF({raw}!{col}{data_row}="","",{raw}!{col}{data_row})')
            ws.cell(row=row, column=5, value="")
            ws.cell(row=row, column=6, value="")
            row += 1

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 55
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 12
    # Hide empty pain rows with filter instruction in A2
    ws["A2"] = "Отфильтруйте столбец D (скрыть пустые) для просмотра только заполненных ответов"


def build_interview_sheet(wb, questions):
    ws = wb.create_sheet("Интервью")
    ws["A1"] = "Кандидаты на углублённое интервью (H4 = Да / Возможно)"
    ws["A1"].font = Font(bold=True, size=12, color="1F4E79")
    headers = ["Филиал", "Должность", "Готовность к интервью", "Контакт", "Средняя оценка по шкалам", "Примечание"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 3, len(headers))

    raw = "'Сырые ответы'"
    col_h4 = col_letter_for_id(questions, "H4")
    col_h5 = col_letter_for_id(questions, "H5")
    row = 4
    for data_row in range(2, 52):
        ws.cell(row=row, column=1, value=f"={raw}!B{data_row}")
        ws.cell(row=row, column=2, value=f"={raw}!C{data_row}")
        if col_h4:
            ws.cell(row=row, column=3, value=f"={raw}!{col_h4}{data_row}")
        if col_h5:
            ws.cell(row=row, column=4, value=f"={raw}!{col_h5}{data_row}")
        # avg of key scales for prioritization
        scale_cols = [col_letter_for_id(questions, x) for x in ["E1", "E3", "C3", "D2", "G2"]]
        scale_cols = [c for c in scale_cols if c]
        if scale_cols:
            parts = [f"{raw}!{c}{data_row}" for c in scale_cols]
            ws.cell(row=row, column=5, value=f'=IF(COUNT({",".join(parts)})=0,"",AVERAGE({",".join(parts)}))')
        row += 1

    ws["A2"] = "Отфильтруйте столбец C: оставить «Да» и «Возможно, по согласованию»; сортировать по столбцу E (низкие оценки — приоритет)"
    for c, w in enumerate([18, 22, 22, 28, 14, 20], 1):
        ws.column_dimensions[get_column_letter(c)].width = w


def build_distribution_sheet(wb, questions):
    ws = wb.create_sheet("Распределения")
    ws["A1"] = "Ключевые категориальные вопросы — частота ответов"
    ws["A1"].font = Font(bold=True, size=12, color="1F4E79")

    categorical = [
        ("B7", "Задержки закупок"),
        ("D3", "Задержки платежей"),
        ("D5", "Участие в прогнозе поступлений"),
        ("E5", "Эскалации в обход ДУП"),
        ("C4", "Продукт без пилота"),
        ("G4", "Часы на ручную отчётность"),
    ]

    raw = "'Сырые ответы'"
    row = 3
    for qid, title in categorical:
        col = col_letter_for_id(questions, qid)
        q = next((x for x in questions if x["id"] == qid), None)
        if not col or not q:
            continue
        ws.cell(row=row, column=1, value=f"{qid}: {title}").font = Font(bold=True, color="1F4E79")
        row += 1
        ws.cell(row=row, column=1, value="Вариант ответа")
        ws.cell(row=row, column=2, value="Кол-во")
        ws.cell(row=row, column=3, value="%")
        style_header(ws, row, 3)
        row += 1
        opts = q.get("options") or []
        if q["type"] == "шкала 1–5":
            opts = ["1", "2", "3", "4", "5"]
        first_data_row = row
        for opt in opts:
            ws.cell(row=row, column=1, value=opt)
            ws.cell(row=row, column=2, value=f'=COUNTIF({raw}!{col}2:{col}500,"{opt}")')
            row += 1
        last_data_row = row - 1
        for r in range(first_data_row, last_data_row + 1):
            ws.cell(row=r, column=3, value=f'=IF(SUM($B${first_data_row}:$B${last_data_row})=0,"",B{r}/SUM($B${first_data_row}:$B${last_data_row}))')
            ws.cell(row=r, column=3).number_format = "0%"
        row += 2

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 8


def main():
    questions = flat_questions()
    wb = Workbook()
    build_instruction_sheet(wb)
    build_raw_sheet(wb, questions)
    build_mapping_sheet(wb, questions)
    build_dashboard(wb, questions)
    build_scales_sheet(wb, questions)
    build_co_sheet(wb, questions)
    build_distribution_sheet(wb, questions)
    build_pain_sheet(wb, questions)
    build_interview_sheet(wb, questions)
    wb.save(OUTPUT)
    print(f"Saved: {OUTPUT}")


if __name__ == "__main__":
    main()

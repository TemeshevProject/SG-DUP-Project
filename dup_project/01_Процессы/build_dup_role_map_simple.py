#!/usr/bin/env python3
"""Упрощённая карта ролей ДУП — одна вкладка, четыре колонки.

Данные берутся из подробной карты (build_dup_role_map.py), чтобы два файла
не расходились: правка вносится один раз, оба файла пересобираются.

Колонки: уровень, наименование, роль ДУП (владелец / участник), что делает ДУП.
Роль сворачивается до трёх значений по владельцу сквозного процесса:

    владелец — ДУП (в том числе разделённое владение)  → Владелец
    владелец — другое подразделение                     → Участник
    владелец не закреплён                               → Не определён

Свернуть «ничью» зону в «владельца» или «участника» было бы неверно: там
обязанность не закреплена ни в одной должностной инструкции.
"""

from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from build_dup_role_map import НЗ, PROCESSES

OUTPUT = Path(__file__).resolve().parent / "ДУП_роли_в_процессах_кратко.xlsx"

ВЛАДЕЛЕЦ = "Владелец"
УЧАСТНИК = "Участник"
НЕ_ОПРЕДЕЛЁН = "Не определён"

ROLE_ORDER = [ВЛАДЕЛЕЦ, УЧАСТНИК, НЕ_ОПРЕДЕЛЁН]
ROLE_FILLS = {
    ВЛАДЕЛЕЦ: PatternFill("solid", fgColor="C6E9D5"),
    УЧАСТНИК: PatternFill("solid", fgColor="CFE0F7"),
    НЕ_ОПРЕДЕЛЁН: PatternFill("solid", fgColor="F8CBC4"),
}

LEVEL_INDENT = {"Процесс": 0, "Подпроцесс": 1, "Этап": 2}
LEVEL_FILL = PatternFill("solid", fgColor="EDF1F7")  # подсветка строк уровня «Процесс»

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

HEADERS = ["Уровень", "Наименование процесса / подпроцесса / этапа", "Роль ДУП", "Что делает ДУП"]
WIDTHS = [14, 62, 16, 88]


def simple_role(dup_role: str, owner_dept: str) -> str:
    """Сворачивает роль до трёх значений по владельцу сквозного процесса.

    Решающим является владелец процесса целиком, а не роль ДУП внутри своего
    дерева. Подтверждение готовности продукта (1.2) — подпроцесс блока 1, но
    решение принимает ДРБ, поэтому ДУП здесь участник, а не владелец.
    """
    if dup_role == НЗ:
        return НЕ_ОПРЕДЕЛЁН
    return ВЛАДЕЛЕЦ if owner_dept.startswith("ДУП") else УЧАСТНИК


def action_text(role: str, contribution: str, owner_dept: str, owner_role: str) -> str:
    """Дополняет описание работы ДУП указанием владельца процесса.

    В кратком файле нет колонки владельца, поэтому для чужих процессов имя
    владельца переносится в текст — иначе «участник» повисает без адресата.
    """
    text = contribution.rstrip(" .")
    lowered = text.lower()
    if role == УЧАСТНИК and "владелец процесса" not in lowered:
        owner = owner_dept if owner_dept != "не определён" else owner_role
        text += f". Владелец процесса — {owner}"
    elif role == НЕ_ОПРЕДЕЛЁН and "не назначен" not in lowered:
        text += ". Владелец не назначен"
    return text


def build_rows():
    rows = []
    for p in PROCESSES:
        code, level, name = p[0], p[1], p[2]
        owner_dept, owner_role, dup_role, contribution = p[6], p[7], p[8], p[10]
        role = simple_role(dup_role, owner_dept)
        rows.append((
            level,
            f"{code}  {name}" if code else name,
            role,
            action_text(role, contribution, owner_dept, owner_role),
        ))
    return rows


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Роли ДУП"

    ws.append(HEADERS)
    for c in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

    rows = build_rows()
    for r, (level, name, role, action) in enumerate(rows, 2):
        indent = LEVEL_INDENT.get(level, 0)
        for c, val in enumerate((level, name, role, action), 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = BORDER
            cell.alignment = Alignment(
                wrap_text=True, vertical="top", indent=indent if c == 2 else 0
            )
        if level == "Процесс":
            ws.cell(row=r, column=1).fill = LEVEL_FILL
            ws.cell(row=r, column=2).fill = LEVEL_FILL
            ws.cell(row=r, column=2).font = Font(bold=True, size=10)
        role_cell = ws.cell(row=r, column=3)
        role_cell.fill = ROLE_FILLS[role]
        role_cell.font = Font(bold=True, size=10)

    for i, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{len(rows) + 1}"

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"

    wb.save(OUTPUT)

    counts = Counter(r[2] for r in rows)
    print(f"Сохранено: {OUTPUT}")
    print(f"Записей: {len(rows)}")
    for role in ROLE_ORDER:
        print(f"  {role}: {counts.get(role, 0)}")


if __name__ == "__main__":
    main()

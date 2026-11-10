#!/usr/bin/env python3
"""Два слайда «Цифровой ПМ»: график закупок и исполнение бюджета."""

import json
import subprocess
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Цифровой_ПМ_план_закупок.pptx"

W = 13.333333
H = 7.5
SIDE = 2.18
PAD = 0.14
FONT = "Inter"

C = {
    "side": "0C3D86",
    "pill": "3C86F0",
    "bg": "E8EEF5",
    "card": "FFFFFF",
    "ink": "1B2836",
    "muted": "5D6E80",
    "line": "D5DFEA",
    "navy": "1A4F8C",
    "green": "1E8C4A",
    "red": "D9342C",
    "gold": "E0A106",
    "today": "013D85",
    "late": "FFF5F4",
    "miss": "F3C1BE",
    "track": "E7EEF6",
    "gray": "9AAFC4",
    "off": "9AA8B5",
    "nav": "E7F0FF",
}


def rgb(value):
    return RGBColor.from_string(value)


def load():
    raw = subprocess.check_output(
        [
            "node",
            "-e",
            """
            const M = require('./model.js');
            const rows = M.rows().map((r) => ({
              id: r.id, name: r.name, who: r.who, state: r.state,
              late: r.late, slip: r.slip, stage: r.stage,
              contract: r.contract, paid: r.paid, delivered: r.delivered,
              start: M.pos(r.charterStart), need: M.pos(r.need),
              req: M.pos(r.requiredIso), forecast: M.pos(r.forecastIso),
              needDate: M.shortDate(r.need), forecastDate: M.shortDate(r.forecastIso)
            }));
            const months = M.months();
            console.log(JSON.stringify({
              today: M.shortDate(M.TODAY),
              commission: M.shortDate(M.COMMISSION),
              cost: M.PROJECT_COST,
              todayPos: M.pos(M.TODAY),
              commPos: M.pos(M.COMMISSION),
              months: months.map((m, i) => ({
                label: m.label, year: m.year,
                pos: M.pos(m.start),
                end: i + 1 < months.length ? M.pos(months[i + 1].start) : 100
              })),
              rows,
              summary: M.summary()
            }));
            """,
        ],
        cwd=ROOT,
        text=True,
    )
    return json.loads(raw)


def rect(slide, x, y, w, h, fill, line=None, radius=None):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(line)
        shape.line.width = Pt(1)
    if radius is not None:
        shape.adjustments[0] = radius
    return shape


def _anchor(tf, anchor):
    tf._txBody.bodyPr.set("anchor", {"ctr": "ctr", "t": "t", "b": "b"}[anchor])


def write(shape, paragraphs, size=12, bold=False, color="ink", align="left", anchor="ctr", margin=0.08):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    _anchor(tf, anchor)
    align_enum = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[align]
    if isinstance(paragraphs, str):
        paragraphs = [paragraphs]
    for i, item in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align_enum
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        runs = item if isinstance(item, list) else [(item, {})]
        for text, opt in runs:
            run = p.add_run()
            run.text = text
            run.font.name = FONT
            run.font.size = Pt(opt.get("size", size))
            run.font.bold = opt.get("bold", bold)
            run.font.color.rgb = rgb(C[opt["color"]] if opt.get("color") in C else opt.get("color", C[color]))
            # Force latin and cs typeface so LibreOffice keeps Inter for Cyrillic.
            rpr = run._r.get_or_add_rPr()
            for tag in ("a:latin", "a:cs", "a:ea"):
                node = rpr.find(qn(tag))
                if node is None:
                    node = etree.SubElement(rpr, qn(tag))
                node.set("typeface", FONT)
    return shape


def textbox(slide, x, y, w, h, paragraphs, **kwargs):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    write(shape, paragraphs, **kwargs)
    return shape


def shell(slide, data, view):
    rect(slide, 0, 0, W, H, C["bg"])
    rect(slide, 0, 0, SIDE, H, C["side"])
    textbox(slide, 0.16, 0.22, SIDE - 0.28, 0.4, "Цифровой ПМ", size=16, bold=True, color="card", margin=0.04)
    items = ["План проекта", "План закупок", "Расходные договоры", "Платежи", "Уведомления"]
    y = 0.82
    for name in items:
        if name == "План закупок":
            rect(slide, 0.12, y, SIDE - 0.24, 0.4, C["pill"], radius=0.5)
            textbox(slide, 0.12, y, SIDE - 0.24, 0.4, name, size=13, bold=True, color="card", margin=0.12)
        else:
            textbox(slide, 0.12, y, SIDE - 0.24, 0.4, name, size=13, color="nav", margin=0.12)
        y += 0.46
    textbox(
        slide, 0.16, 6.55, SIDE - 0.32, 0.78,
        ["Учебный пример.", "Данные вымышлены."],
        size=10, color="nav", anchor="b", margin=0.02,
    )

    content_x = SIDE + PAD
    content_r = W - PAD
    switch_w, switch_h = 4.15, 0.4
    switch_x = content_r - switch_w
    textbox(
        slide, content_x, 0.08, switch_x - content_x - 0.12, 0.46,
        "Оснащение учебного кампуса «Норд» · вымышленный пример",
        size=13, bold=True, anchor="ctr", margin=0,
    )
    textbox(
        slide, content_x, 0.56, content_r - content_x, 0.22,
        "Срез " + data["today"] + "  ·  ввод " + data["commission"] + "  ·  учебный проект " + data["cost"],
        size=11, color="muted", anchor="ctr", margin=0,
    )
    rect(slide, switch_x, 0.12, switch_w, switch_h, C["card"], line=C["line"], radius=0.4)
    labels = [("График закупок", "gantt"), ("Исполнение бюджета", "exec")]
    inner = 0.04
    btn_w = (switch_w - inner * 2) / 2
    for i, (label, key) in enumerate(labels):
        bx = switch_x + inner + i * btn_w
        if key == view:
            rect(slide, bx, 0.12 + inner, btn_w, switch_h - inner * 2, C["side"], radius=0.4)
            textbox(slide, bx, 0.12 + inner, btn_w, switch_h - inner * 2, label, size=12, bold=True, color="card", align="center", margin=0.04)
        else:
            textbox(slide, bx, 0.12 + inner, btn_w, switch_h - inner * 2, label, size=12, color="muted", align="center", margin=0.04)

    s = data["summary"]
    if view == "gantt":
        kpis = [
            (str(s["total"]), "строк в плане закупок", "ink"),
            (str(s["total"] - s["late"]), "к дате потребности", "ink"),
            (str(s["late"]), "с отклонением", "red"),
            (data["commission"], "ввод в ПЭ", "ink"),
        ]
    else:
        kpis = [
            (str(s["contract"]), "договоров из " + str(s["total"]), "ink"),
            (str(s["paid"]), "строк оплачено", "ink"),
            (str(s["delivered"]), "поставлено — исполнение", "ink"),
            (str(s["late"]), "прогноз позже даты", "red"),
        ]
    span = content_r - content_x
    gap = 0.1
    kw = (span - gap * 3) / 4
    ky, kh = 0.8, 0.7
    for i, (value, label, tone) in enumerate(kpis):
        x = content_x + i * (kw + gap)
        rect(slide, x, ky, kw, kh, C["card"], line=C["line"], radius=0.12)
        textbox(slide, x + 0.08, ky + 0.04, kw - 0.16, 0.36, value, size=22, bold=True, color=tone, margin=0.04, anchor="b")
        textbox(slide, x + 0.08, ky + 0.38, kw - 0.16, 0.28, label, size=10, color="muted", margin=0.04, anchor="t")

    notes_w = 2.58
    panel_x = content_x
    panel_y = 1.64
    panel_h = H - panel_y - 0.14
    panel_w = span - notes_w - gap
    notes_x = panel_x + panel_w + gap
    return {
        "panel": (panel_x, panel_y, panel_w, panel_h),
        "notes": (notes_x, panel_y, notes_w),
    }


def cards(slide, geo, data, view):
    x, y, w = geo["notes"]
    video = next(r for r in data["rows"] if r["id"] == "video")
    verify = next(r for r in data["rows"] if r["id"] == "verify")
    items = [
        ("red", "Отклонение от плана", "Поверка комплекта: финиш " + verify["forecastDate"] + " вместо " + verify["needDate"] + ". Под угрозой запуск сервиса."),
        ("gold", "НордСнаб · план закупок", "Видеоаналитика: заявку не подали. Старт " + data["today"] + " даёт приёмку " + video["forecastDate"] + ", веха " + video["needDate"] + "."),
        ("green", "Выполнено", "Датчики и камеры приняты 01.02.25. Мобильный комплект принят 28.04.25, до старта поверки."),
    ]
    if view == "exec":
        items.append(("pill", "Как считается исполнение", "Исполнение — только поставленные и принятые ТРУ. Договор и оплата видны отдельно. Оплата после приёмки не двигает дату каналов и поверки."))
    gap = 0.08
    h = 1.28 if view == "gantt" else 1.22
    for i, (tone, tag, body) in enumerate(items):
        cy = y + i * (h + gap)
        rect(slide, x, cy, w, h, C["card"], line=C["line"], radius=0.1)
        rect(slide, x + 0.08, cy + 0.12, 0.06, h - 0.24, C[tone])
        textbox(slide, x + 0.22, cy + 0.08, w - 0.32, 0.24, tag.upper(), size=9, bold=True, color="muted", margin=0, anchor="b")
        textbox(slide, x + 0.22, cy + 0.32, w - 0.32, h - 0.4, body, size=11, margin=0, anchor="t")


def chart_x(pct, origin, width):
    return origin + width * pct / 100.0


def gantt(slide, data, geo):
    px, py, pw, ph = geo["panel"]
    rect(slide, px, py, pw, ph, C["card"], line=C["line"], radius=0.08)
    rows = data["rows"]
    head_h = 0.42
    legend_h = 0.58
    top = py + 0.08
    name_w, who_w = 2.15, 1.02
    chart_origin = px + 0.1 + name_w + who_w
    chart_w = pw - 0.18 - name_w - who_w
    body_top = top + head_h
    body_h = ph - 0.12 - head_h - legend_h
    row_h = body_h / len(rows)

    textbox(slide, px + 0.1, top, name_w, head_h, "ЗАДАЧА", size=10, bold=True, color="muted", margin=0.02, anchor="b")
    textbox(slide, px + 0.1 + name_w, top, who_w, head_h, "КТО", size=10, bold=True, color="muted", margin=0.02, anchor="b")

    months = data["months"]
    year_cuts = []
    current = None
    for m in months:
        if m["year"] != current:
            year_cuts.append([m["year"], m["pos"], m["end"]])
            current = m["year"]
        else:
            year_cuts[-1][2] = m["end"]
    for year, a, b in year_cuts:
        x0 = chart_x(a, chart_origin, chart_w)
        x1 = chart_x(b, chart_origin, chart_w)
        textbox(slide, x0, top, x1 - x0, 0.18, str(year), size=9, color="muted", align="center", margin=0, anchor="ctr")
    for m in months:
        x0 = chart_x(m["pos"], chart_origin, chart_w)
        x1 = chart_x(m["end"], chart_origin, chart_w)
        textbox(slide, x0, top + 0.18, max(0.2, x1 - x0), 0.24, m["label"], size=9, color="muted", align="center", margin=0, anchor="t")
        rect(slide, x0, body_top, 0.01, body_h, "E6EDF4")

    for i, row in enumerate(rows):
        ry = body_top + i * row_h
        if row["late"]:
            rect(slide, px + 0.06, ry + 0.03, pw - 0.12, row_h - 0.04, C["late"], radius=0.08)
        textbox(slide, px + 0.1, ry, name_w - 0.04, row_h, row["name"], size=11, margin=0.02)
        textbox(slide, px + 0.1 + name_w, ry, who_w - 0.04, row_h, row["who"], size=11, color="muted", margin=0.02)
        by = ry + (row_h - 0.16) / 2
        if row["state"] == "notStarted":
            a = chart_x(row["req"], chart_origin, chart_w)
            b = chart_x(data["todayPos"], chart_origin, chart_w)
            c = chart_x(row["forecast"], chart_origin, chart_w)
            rect(slide, a, by, max(0.04, b - a), 0.16, C["miss"])
            rect(slide, b, by, max(0.04, c - b), 0.16, C["red"])
        else:
            a = chart_x(row["start"], chart_origin, chart_w)
            b = chart_x(row["need"], chart_origin, chart_w)
            color = C["green"] if row["delivered"] else C["navy"]
            rect(slide, a, by, max(0.04, b - a), 0.16, color)
            if row["late"]:
                c = chart_x(row["forecast"], chart_origin, chart_w)
                rect(slide, b, by, max(0.04, c - b), 0.16, C["red"])
        tick = chart_x(row["need"], chart_origin, chart_w)
        rect(slide, tick, ry + 0.1, 0.018, row_h - 0.2, C["ink"])

    def vline(pos, color, dashed=False):
        x = chart_x(pos, chart_origin, chart_w)
        if not dashed:
            rect(slide, x, body_top, 0.018, body_h, color)
            return
        y = body_top
        while y < body_top + body_h:
            seg = min(0.08, body_top + body_h - y)
            rect(slide, x, y, 0.018, seg, color)
            y += 0.14

    vline(data["todayPos"], C["today"])
    vline(data["commPos"], C["green"], dashed=True)

    legend_y = py + ph - legend_h
    items = [("navy", "в графике"), ("green", "принято"), ("red", "позже даты"), ("miss", "не подали")]
    lx = px + 0.12
    for color, label in items:
        rect(slide, lx, legend_y + 0.06, 0.18, 0.1, C[color])
        textbox(slide, lx + 0.22, legend_y, 1.05, 0.24, label, size=10, color="muted", margin=0)
        lx += 1.28
    textbox(
        slide, px + 0.12, legend_y + 0.26, pw - 0.24, 0.28,
        "Чёрная черта — дата потребности.  Пунктир — ввод " + data["commission"] + ".  Синяя черта — срез " + data["today"] + ".",
        size=10, color="muted", margin=0, anchor="t",
    )


def execution(slide, data, geo):
    px, py, pw, ph = geo["panel"]
    rect(slide, px, py, pw, ph, C["card"], line=C["line"], radius=0.08)
    textbox(
        slide, px + 0.12, py + 0.08, pw - 0.24, 0.46,
        "Исполнение — поставка. Договор и оплата видны отдельно и в исполнение не входят. Суммы строк в примере не заданы, шкала — по числу строк.",
        size=12, color="muted", anchor="t", margin=0,
    )
    s = data["summary"]
    funnel = [
        ("Строк закупки", s["total"], "gray"),
        ("Договор заключён", s["contract"], "navy"),
        ("Оплачено", s["paid"], "gold"),
        ("Поставлено", s["delivered"], "green"),
    ]
    fy = py + 0.58
    label_w = 1.55
    num_w = 0.36
    track_x = px + 0.14 + label_w
    track_w = pw - 0.28 - label_w - num_w
    for i, (label, value, color) in enumerate(funnel):
        y = fy + i * 0.28
        textbox(slide, px + 0.12, y, label_w, 0.26, label, size=11, margin=0, anchor="ctr")
        rect(slide, track_x, y + 0.07, track_w, 0.12, C["track"], radius=0.5)
        width = track_w * (value / s["total"])
        if width > 0:
            rect(slide, track_x, y + 0.07, max(0.08, width), 0.12, C[color], radius=0.5)
        textbox(slide, track_x + track_w + 0.04, y, num_w, 0.26, str(value), size=12, bold=True, margin=0, anchor="ctr")

    headers = ["ТРУ", "Кто", "Договор", "Оплата", "Поставлено", "Прогноз", "К дате"]
    widths = [2.15, 1.05, 0.78, 0.72, 0.98, 0.82, 0.78]
    table_y = fy + 1.18
    head_h = 0.28
    x = px + 0.1
    for title, width in zip(headers, widths):
        textbox(slide, x, table_y, width, head_h, title.upper(), size=9, bold=True, color="muted", margin=0.02, anchor="b")
        x += width
    rect(slide, px + 0.1, table_y + head_h, pw - 0.2, 0.01, C["line"])
    row_h = (py + ph - 0.08 - (table_y + head_h)) / len(data["rows"])
    for i, row in enumerate(data["rows"]):
        ry = table_y + head_h + i * row_h
        if row["late"]:
            rect(slide, px + 0.06, ry + 0.02, pw - 0.12, row_h - 0.03, C["late"], radius=0.06)
        slip = "+" + str(row["slip"]) + " дн" if row["slip"] > 0 else "в срок"
        cells = [
            (row["name"], "ink", False),
            (row["who"], "muted", False),
            ("да" if row["contract"] else "нет", "green" if row["contract"] else "off", True),
            ("да" if row["paid"] else "нет", "green" if row["paid"] else "off", True),
            ("да" if row["delivered"] else "нет", "green" if row["delivered"] else "off", True),
            (row["forecastDate"], "ink", False),
            (slip, "red" if row["slip"] > 0 else "green", True),
        ]
        x = px + 0.1
        for (text, color, strong), width in zip(cells, widths):
            textbox(slide, x, ry, width - 0.04, row_h, text, size=11, bold=strong, color=color, margin=0.02)
            x += width


def build():
    data = load()
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    prs.core_properties.title = "Цифровой ПМ — план закупок и исполнение бюджета"
    blank = prs.slide_layouts[6]
    for view, draw in (("gantt", gantt), ("exec", execution)):
        slide = prs.slides.add_slide(blank)
        geo = shell(slide, data, view)
        draw(slide, data, geo)
        cards(slide, geo, data, view)
    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()

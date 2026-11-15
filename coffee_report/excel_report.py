"""Сборка оформленного report.xlsx из очищенных данных."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

COMPANY = "Сеть кофеен «Зерно»"

MONEY = '#,##0 "₽"'
INT = "#,##0"
PCT = "0.0%"
DATE = "dd.mm.yyyy"

BRAND = "6F4E37"       # кофейный
BRAND_LIGHT = "F3ECE6"
FONT = "Calibri"

TITLE_FONT = Font(name=FONT, size=16, bold=True, color=BRAND)
SUBTITLE_FONT = Font(name=FONT, size=10, italic=True, color="7F7F7F")
SECTION_FONT = Font(name=FONT, size=12, bold=True, color=BRAND)
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor=BRAND)
TOTAL_FONT = Font(name=FONT, bold=True)
TOTAL_FILL = PatternFill("solid", fgColor=BRAND_LIGHT)
KPI_LABEL_FONT = Font(name=FONT, size=9, color="7F7F7F")
KPI_VALUE_FONT = Font(name=FONT, size=14, bold=True, color=BRAND)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONTHS = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль",
          "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]


def month_title(month: str) -> str:
    period = pd.Period(month, "M")
    return f"{MONTHS[period.month - 1].capitalize()} {period.year}"


def summarize(data: pd.DataFrame) -> dict:
    """Все агрегаты, которые нужны для отчёта и сводки в Telegram."""
    by_branch = (data.groupby("branch")
                 .agg(revenue=("amount", "sum"), receipts=("receipt", "nunique"))
                 .sort_values("revenue", ascending=False))
    by_branch["avg_check"] = by_branch["revenue"] / by_branch["receipts"]
    by_branch["share"] = by_branch["revenue"] / by_branch["revenue"].sum()

    by_product = (data.groupby("product")
                  .agg(qty=("qty", "sum"), revenue=("amount", "sum"))
                  .sort_values("revenue", ascending=False))
    by_product["share"] = by_product["revenue"] / by_product["revenue"].sum()
    by_product["avg_price"] = by_product["revenue"] / by_product["qty"]

    by_day = data.groupby("date")["amount"].sum()
    receipts = int(by_branch["receipts"].sum())
    revenue = float(data["amount"].sum())
    return {
        "revenue": revenue,
        "receipts": receipts,
        "items": int(data["qty"].sum()),
        "avg_check": revenue / receipts if receipts else 0.0,
        "avg_day": float(by_day.mean()),
        "by_branch": by_branch,
        "by_product": by_product,
        "by_day": by_day,
    }


# --- Помощники оформления ----------------------------------------------------

def _header(ws: Worksheet, row: int, col: int, titles: list[str]) -> None:
    for i, title in enumerate(titles):
        cell = ws.cell(row=row, column=col + i, value=title)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 30


def _rows(ws: Worksheet, row: int, col: int, rows: list[list], formats: list[str | None],
          total: list | None = None) -> int:
    """Пишет строки таблицы (и итог), возвращает номер последней строки."""
    for r, values in enumerate(rows + ([total] if total else [])):
        is_total = total is not None and r == len(rows)
        for c, (value, fmt) in enumerate(zip(values, formats)):
            cell = ws.cell(row=row + r, column=col + c, value=value)
            cell.border = BORDER
            cell.font = TOTAL_FONT if is_total else Font(name=FONT)
            if is_total:
                cell.fill = TOTAL_FILL
            if fmt and value is not None:
                cell.number_format = fmt
    return row + len(rows) + (1 if total else 0) - 1


def _section(ws: Worksheet, row: int, col: int, title: str) -> None:
    ws.cell(row=row, column=col, value=title).font = SECTION_FONT


def _widths(ws: Worksheet, widths: dict[str, float]) -> None:
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width


def _color_scale(ws: Worksheet, ref: str) -> None:
    ws.conditional_formatting.add(ref, ColorScaleRule(
        start_type="min", start_color="F8696B",
        mid_type="percentile", mid_value=50, mid_color="FFEB84",
        end_type="max", end_color="63BE7B"))


def _title(ws: Worksheet, title: str, month: str, last_col: int) -> None:
    ws["A1"] = f"{COMPANY} — {title}"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Период: {month_title(month)} · сформировано {date.today():%d.%m.%Y}"
    ws["A2"].font = SUBTITLE_FONT
    ws.row_dimensions[1].height = 24
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)


# --- Листы -------------------------------------------------------------------

def _summary_sheet(ws: Worksheet, s: dict, month: str) -> None:
    _title(ws, "итоги продаж", month, 11)

    kpis = [("Выручка", s["revenue"], MONEY), ("Чеков", s["receipts"], INT),
            ("Средний чек", s["avg_check"], MONEY), ("Продано позиций", s["items"], INT),
            ("Выручка в день", s["avg_day"], MONEY)]
    for i, (label, value, fmt) in enumerate(kpis):
        col = 1 + i * 2
        ws.cell(row=4, column=col, value=label).font = KPI_LABEL_FONT
        cell = ws.cell(row=5, column=col, value=value)
        cell.font = KPI_VALUE_FONT
        cell.number_format = fmt
        cell.alignment = Alignment(horizontal="left")

    # Выручка по филиалам: A–E
    top = 7
    _section(ws, top, 1, "Выручка по филиалам")
    _header(ws, top + 1, 1, ["Филиал", "Выручка", "Чеков", "Средний чек", "Доля"])
    b = s["by_branch"]
    rows = [[name, r.revenue, r.receipts, r.avg_check, r.share] for name, r in b.iterrows()]
    total = ["Итого", s["revenue"], s["receipts"], s["avg_check"], 1.0]
    last = _rows(ws, top + 2, 1, rows, [None, MONEY, INT, MONEY, PCT], total)
    _color_scale(ws, f"B{top + 2}:B{last - 1}")
    branch_last = last

    # Выручка по товарам: G–I
    _section(ws, top, 7, "Выручка по товарам")
    _header(ws, top + 1, 7, ["Товар", "Продано, шт", "Выручка"])
    p = s["by_product"]
    rows = [[name, r.qty, r.revenue] for name, r in p.iterrows()]
    total = ["Итого", s["items"], s["revenue"]]
    last = _rows(ws, top + 2, 7, rows, [None, INT, MONEY], total)
    _color_scale(ws, f"I{top + 2}:I{last - 1}")

    # Выручка по дням: K–L
    _section(ws, top, 11, "Выручка по дням")
    _header(ws, top + 1, 11, ["Дата", "Выручка"])
    rows = [[d.date(), v] for d, v in s["by_day"].items()]
    day_first = top + 2
    day_last = _rows(ws, day_first, 11, rows, [DATE, MONEY], ["Итого", s["revenue"]]) - 1
    _color_scale(ws, f"L{day_first}:L{day_last}")

    # Нативный линейный график выручки по дням
    line = LineChart()
    line.title = "Выручка по дням"
    line.style = 12
    line.height, line.width = 8, 24
    line.y_axis.numFmt = "#,##0"
    line.y_axis.majorGridlines = None
    line.x_axis.number_format = "dd.mm"
    line.legend = None
    line.add_data(Reference(ws, min_col=12, min_row=day_first - 1, max_row=day_last),
                  titles_from_data=True)
    line.set_categories(Reference(ws, min_col=11, min_row=day_first, max_row=day_last))
    line.series[0].graphicalProperties.line.solidFill = BRAND
    line.series[0].graphicalProperties.line.width = 28000
    line.series[0].smooth = False
    line.x_axis.delete = False
    line.y_axis.delete = False
    ws.add_chart(line, f"A{branch_last + 18}")

    # Столбчатая диаграмма выручки по филиалам
    bar = BarChart()
    bar.type = "bar"
    bar.title = "Выручка по филиалам"
    bar.height, bar.width = 7, 12
    bar.legend = None
    bar.y_axis.numFmt = "#,##0"
    bar.y_axis.majorGridlines = None
    bar.add_data(Reference(ws, min_col=2, min_row=top + 1, max_row=branch_last - 1),
                 titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=top + 2, max_row=branch_last - 1))
    bar.series[0].graphicalProperties.solidFill = BRAND
    bar.x_axis.delete = False
    bar.y_axis.delete = False
    bar.x_axis.scaling.orientation = "maxMin"  # лидер сверху
    bar.y_axis.crosses = "max"  # а шкала значений при этом остаётся внизу
    ws.add_chart(bar, f"A{branch_last + 2}")  # под таблицей филиалов

    _widths(ws, {"A": 16, "B": 15, "C": 10, "D": 14, "E": 9, "F": 3,
                 "G": 16, "H": 13, "I": 15, "J": 3, "K": 12, "L": 15})
    ws.freeze_panes = "A4"


def _top_sheet(ws: Worksheet, s: dict, month: str) -> None:
    _title(ws, "топ-10 товаров", month, 6)
    _header(ws, 4, 1, ["Место", "Товар", "Продано, шт", "Выручка", "Доля выручки", "Средняя цена"])
    top = s["by_product"].head(10)
    rows = [[i, name, r.qty, r.revenue, r.share, r.avg_price]
            for i, (name, r) in enumerate(top.iterrows(), start=1)]
    last = _rows(ws, 5, 1, rows, [None, None, INT, MONEY, PCT, MONEY])
    for row in ws.iter_rows(min_row=5, max_row=last, max_col=1):
        row[0].alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=5, max_row=7, min_col=2, max_col=2):
        row[0].font = Font(name=FONT, bold=True)  # тройка лидеров
    ws.conditional_formatting.add(f"D5:D{last}", DataBarRule(
        start_type="num", start_value=0, end_type="max", color="C8A27C"))
    ws.conditional_formatting.add(f"E5:E{last}", DataBarRule(
        start_type="num", start_value=0, end_type="max", color="9DC3E6"))
    _widths(ws, {"A": 8, "B": 18, "C": 13, "D": 18, "E": 16, "F": 14})
    ws.freeze_panes = "A5"


def _data_sheet(ws: Worksheet, data: pd.DataFrame) -> None:
    titles = ["Дата", "Филиал", "№ чека", "Товар", "Кол-во", "Цена", "Сумма"]
    _header(ws, 1, 1, titles)
    formats = [DATE, None, None, None, INT, MONEY, MONEY]
    cols = ["date", "branch", "receipt", "product", "qty", "price", "amount"]
    for r, values in enumerate(data[cols].itertuples(index=False), start=2):
        for c, (value, fmt) in enumerate(zip(values, formats), start=1):
            if c == 1:
                value = value.date()
            cell = ws.cell(row=r, column=c, value=value)
            if fmt:
                cell.number_format = fmt
    last = len(data) + 1
    ws.auto_filter.ref = f"A1:{get_column_letter(len(titles))}{last}"
    ws.freeze_panes = "A2"
    _widths(ws, {"A": 12, "B": 14, "C": 12, "D": 16, "E": 9, "F": 11, "G": 12})


def build_report(data: pd.DataFrame, month: str, output: Path) -> dict:
    """Создаёт report.xlsx и возвращает агрегаты (пригодятся для сводки)."""
    s = summarize(data)
    wb = Workbook()
    wb.active.title = "Итоги"  # до графиков: они ссылаются на лист по имени
    _summary_sheet(wb.active, s, month)
    _top_sheet(wb.create_sheet("Топ-10 товаров"), s, month)
    _data_sheet(wb.create_sheet("Данные"), data)
    for ws in wb:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb["Итоги"].page_setup.fitToHeight = 1  # итоги — ровно одна страница
    wb["Данные"].print_title_rows = "1:1"

    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return s

"""Генерирует 4 «грязных» Excel-файла с продажами филиалов вымышленной сети кофеен.

Каждый филиал выгружает данные по-своему: свой порядок и названия колонок,
свой формат дат, шапки-заголовки над таблицей, пустые строки, опечатки
в названиях товаров, числа строками, дубли, пропущенные суммы и строки
за соседние месяцы.

Запуск:
    python generate_sample_data.py                 # октябрь 2026 в input/
    python generate_sample_data.py --month 2026-11 --output data/
"""

from __future__ import annotations

import argparse
import calendar
import random
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook

SEED = 42

# Каноническое название -> (цена, варианты написания, как их вводят кассиры)
PRODUCTS = {
    "Эспрессо": (150, ["эспрессо", "Эспресо", "ЭСПРЕССО", "Экспрессо"]),
    "Американо": (180, ["американо", "Амеракано", "Американо ", "AMERICANO"]),
    "Капучино": (220, ["капучино", "Капуччино", "КАПУЧИНО ", "  Капучино"]),
    "Латте": (240, ["латте", "LATTE", "Латтте", "Латте  "]),
    "Раф": (260, ["раф", "РАФ", "Раф кофе", "раф "]),
    "Флэт уайт": (250, ["флэт уайт", "Флет уайт", "Flat white", "Флэт  уайт"]),
    "Какао": (200, ["какао", "Какао ", "КАКАО", "Какава"]),
    "Чай чёрный": (150, ["чай черный", "Чай  чёрный", "Чай черный", "ЧАЙ ЧЁРНЫЙ"]),
    "Чай зелёный": (150, ["чай зеленый", "Чай зеленый", "Чай зелёный ", "чай зелённый"]),
    "Круассан": (190, ["круассан", "Круасан", "КРУАССАН", "Круассаан"]),
    "Чизкейк": (290, ["чизкейк", "Чискейк", "ЧИЗКЕЙК", "Чизкеик"]),
    "Маффин": (170, ["маффин", "Мафин", "МАФФИН", "Маффин "]),
}

# Популярность товаров (веса для случайного выбора)
WEIGHTS = [8, 10, 14, 13, 7, 6, 5, 6, 4, 9, 5, 6]

BRANCHES = {
    "Центр": {
        "columns": {"receipt": "№ чека", "date": "Дата", "product": "Товар",
                    "qty": "Кол-во", "price": "Цена", "amount": "Сумма"},
        "order": ["receipt", "date", "product", "qty", "price", "amount"],
        "date_style": "iso",
        "title": ["Отчёт о продажах", "Филиал: Центр"],
        "receipts_per_day": (24, 34),
        "markup": 1.0,
        "total_row": True,
    },
    "Северный": {
        "columns": {"product": "Наименование", "date": "Дата продажи",
                    "qty": "Количество", "amount": "Сумма", "price": "Цена за шт",
                    "receipt": "Номер чека"},
        "order": ["product", "date", "qty", "amount", "price", "receipt"],
        "date_style": "dmy_dot",
        "title": ["Продажи филиала «Северный»"],
        "receipts_per_day": (12, 20),
        "markup": 0.95,
        "money_as_text": True,
    },
    "Набережная": {
        "columns": {"date": "date", "product": "product", "price": "price",
                    "qty": "qty", "amount": "total", "receipt": "check_no",
                    "comment": "comment"},
        "order": ["date", "receipt", "product", "price", "qty", "amount", "comment"],
        "date_style": "mixed_excel",
        "title": [],
        "receipts_per_day": (18, 28),
        "markup": 1.05,
        "missing_amounts": 0.04,
    },
    "Вокзал": {
        "columns": {"product": "ТОВАР", "date": "ДАТА", "price": "Цена, руб.",
                    "qty": "Кол.", "amount": "Итого", "receipt": "Чек"},
        "order": ["date", "receipt", "product", "price", "qty", "amount"],
        "date_style": "mixed_dot",
        "title": ["Выгрузка из кассы", "Вокзал", "Сформировано автоматически"],
        "receipts_per_day": (26, 38),
        "markup": 1.1,
        "total_row": True,
        "money_as_text": True,
    },
}

COMMENTS = ["", "", "", "", "скидка сотруднику", "возврат тары", "без сахара"]


def format_date(d: date, style: str, rng: random.Random):
    """Возвращает дату в том виде, в каком её «выгрузила бы» касса филиала."""
    if style == "iso":
        return d.strftime("%Y-%m-%d")
    if style == "dmy_dot":
        return d.strftime("%d.%m.%Y")
    if style == "mixed_excel":
        # Обычно настоящая дата Excel, но иногда ручной ввод через слэш
        if rng.random() < 0.2:
            return f"{d.day}/{d.month}/{d.strftime('%y')}"
        return datetime(d.year, d.month, d.day)
    if style == "mixed_dot":
        r = rng.random()
        if r < 0.1:
            return (d - date(1899, 12, 30)).days  # серийный номер Excel
        if r < 0.25:
            return d.strftime("%Y/%m/%d")
        return d.strftime("%d.%m.%y")
    raise ValueError(style)


def format_money(value: float, as_text: bool, rng: random.Random):
    if not as_text or rng.random() < 0.5:
        return value
    if rng.random() < 0.5:
        return f"{value:,.2f}".replace(",", " ").replace(".", ",")
    return f"{value:.0f} ₽"


def product_name(name: str, rng: random.Random) -> str:
    variants = PRODUCTS[name][1]
    return rng.choice(variants) if rng.random() < 0.35 else name


def generate_branch_rows(cfg: dict, year: int, month: int,
                         rng: random.Random) -> list[dict]:
    names = list(PRODUCTS)
    days = calendar.monthrange(year, month)[1]
    first = date(year, month, 1)
    dates = [first + timedelta(days=i) for i in range(days)]
    # Пара чеков за соседние месяцы — типичная ошибка выгрузки
    stray = [first - timedelta(days=1), first + timedelta(days=days)]

    rows = []
    receipt_no = rng.randint(1000, 5000)
    for d in dates + stray:
        lo, hi = cfg["receipts_per_day"]
        count = rng.randint(lo, hi) if d in dates else 2
        for _ in range(count):
            receipt_no += 1
            items = rng.choices(names, weights=WEIGHTS, k=rng.randint(1, 3))
            for name in dict.fromkeys(items):  # без повторов внутри чека
                price = round(PRODUCTS[name][0] * cfg["markup"] / 10) * 10
                qty = rng.choices([1, 2, 3], weights=[75, 20, 5])[0]
                rows.append({
                    "receipt": receipt_no,
                    "date": d,
                    "product": name,
                    "qty": qty,
                    "price": price,
                    "amount": price * qty,
                })
    return rows


def write_branch_file(path: Path, branch: str, cfg: dict, rows: list[dict],
                      rng: random.Random) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист1"

    for line in cfg["title"]:
        ws.append([line])
    if cfg["title"]:
        ws.append([])

    order = cfg["order"]
    ws.append([cfg["columns"][key] for key in order])

    total = 0
    for row in rows:
        total += row["amount"]
        out = dict(row)
        out["date"] = format_date(row["date"], cfg["date_style"], rng)
        out["product"] = product_name(row["product"], rng)
        out["price"] = format_money(row["price"], cfg.get("money_as_text", False), rng)
        out["amount"] = format_money(row["amount"], cfg.get("money_as_text", False), rng)
        out["comment"] = rng.choice(COMMENTS) or None
        if branch == "Набережная":
            out["receipt"] = row["receipt"]
        elif branch == "Северный":
            out["receipt"] = f"  {row['receipt']}"
        else:
            out["receipt"] = f"{branch[0]}-{row['receipt']:06d}"
        if rng.random() < cfg.get("missing_amounts", 0):
            out["amount"] = None

        ws.append([out[key] for key in order])
        if rng.random() < 0.02:
            ws.append([])  # пустая строка посреди таблицы
        if rng.random() < 0.01:
            ws.append([out[key] for key in order])  # повторная выгрузка строки

    if cfg.get("total_row"):
        ws.append([])
        total_row = {key: None for key in order}
        total_row["product"] = "ИТОГО"
        total_row["amount"] = total
        ws.append([total_row[key] for key in order])

    for col in ws.iter_cols(min_row=1, max_row=1):
        ws.column_dimensions[col[0].column_letter].width = 14
    wb.save(path)


def generate(output_dir: Path, year: int, month: int, seed: int = SEED) -> list[Path]:
    rng = random.Random(seed)
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    for branch, cfg in BRANCHES.items():
        rows = generate_branch_rows(cfg, year, month, rng)
        path = output_dir / f"{branch}.xlsx"
        write_branch_file(path, branch, cfg, rows, rng)
        paths.append(path)
        print(f"  {path}  ({len(rows)} строк продаж)")
    return paths


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--output", default="input", help="папка для файлов (по умолчанию input/)")
    parser.add_argument("--month", default="2026-10", help="месяц в формате ГГГГ-ММ")
    parser.add_argument("--seed", type=int, default=SEED)
    args = parser.parse_args()

    year, month = map(int, args.month.split("-"))
    print(f"Генерирую данные за {args.month}:")
    generate(Path(args.output), year, month, args.seed)


if __name__ == "__main__":
    main()

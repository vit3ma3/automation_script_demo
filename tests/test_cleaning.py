import math
from datetime import datetime

import pandas as pd
import pytest
from openpyxl import Workbook

from coffee_report.cleaning import (
    clean_file,
    detect_month,
    drop_empty_rows,
    drop_total_rows,
    fill_missing_amount,
    filter_month,
    find_header_row,
    load_all,
    normalize_header,
    normalize_product,
    parse_date,
    remove_duplicates,
    to_number,
)


@pytest.mark.parametrize("raw, expected", [
    ("Дата", "date"),
    ("  ДАТА ", "date"),
    ("Дата продажи", "date"),
    ("Наименование", "product"),
    ("product", "product"),
    ("Кол.", "qty"),
    ("Цена, руб.", "price"),
    ("Цена за шт", "price"),
    ("Итого", "amount"),
    ("№ чека", "receipt"),
    ("comment", None),
    (None, None),
])
def test_normalize_header(raw, expected):
    assert normalize_header(raw) == expected


def test_find_header_row_skips_title_lines():
    raw = pd.DataFrame([
        ["Отчёт о продажах", None, None],
        [None, None, None],
        ["Дата", "Товар", "Сумма"],
        ["2026-10-01", "Латте", 240],
    ])
    assert find_header_row(raw) == 2


def test_find_header_row_raises_without_header():
    with pytest.raises(ValueError):
        find_header_row(pd.DataFrame([["a", "b", "c"], [1, 2, 3]]))


@pytest.mark.parametrize("raw, expected", [
    ("2026-10-05", "2026-10-05"),
    ("2026/10/05", "2026-10-05"),
    ("05.10.2026", "2026-10-05"),
    ("05.10.26", "2026-10-05"),
    ("5/10/26", "2026-10-05"),       # день/месяц, а не месяц/день
    (" 31.10.2026 ", "2026-10-31"),
    (datetime(2026, 10, 5, 14, 30), "2026-10-05"),
    (46300, "2026-10-05"),           # серийный номер Excel
])
def test_parse_date(raw, expected):
    assert parse_date(raw) == pd.Timestamp(expected)


@pytest.mark.parametrize("raw", [None, "", "   ", "вчера", "32.13.2026"])
def test_parse_date_invalid(raw):
    assert pd.isna(parse_date(raw))


@pytest.mark.parametrize("raw, expected", [
    (240, 240.0),
    (240.5, 240.5),
    ("240", 240.0),
    ("1 250,00", 1250.0),
    ("1\xa0250,50", 1250.5),
    ("240 ₽", 240.0),
    ("190 руб.", 190.0),
])
def test_to_number(raw, expected):
    assert to_number(raw) == expected


@pytest.mark.parametrize("raw", [None, "", "н/д"])
def test_to_number_invalid(raw):
    assert math.isnan(to_number(raw))


@pytest.mark.parametrize("raw, expected", [
    ("Капучино", "Капучино"),
    ("капучино", "Капучино"),
    ("  КАПУЧИНО ", "Капучино"),
    ("Капуччино", "Капучино"),      # опечатка
    ("Эспресо", "Эспрессо"),
    ("Экспрессо", "Эспрессо"),
    ("чай черный", "Чай чёрный"),   # е вместо ё
    ("Чай  чёрный", "Чай чёрный"),  # двойной пробел
    ("чай зеленый", "Чай зелёный"),
    ("LATTE", "Латте"),             # по справочнику синонимов
    ("Flat white", "Флэт уайт"),
    ("Раф кофе", "Раф"),
])
def test_normalize_product(raw, expected):
    assert normalize_product(raw) == expected


def test_normalize_product_keeps_unknown_items_tidy():
    assert normalize_product("  сэндвич С ВЕТЧИНОЙ ") == "Сэндвич с ветчиной"


def test_normalize_product_does_not_confuse_teas():
    assert normalize_product("Чай зелёный") != normalize_product("Чай чёрный")


def test_normalize_product_blank():
    assert normalize_product(None) is None
    assert normalize_product("  ") is None


def test_drop_empty_rows():
    df = pd.DataFrame({"a": [1, None, "  ", 2], "b": ["x", None, None, "y"]})
    cleaned, removed = drop_empty_rows(df)
    assert removed == 2
    assert list(cleaned["a"]) == [1, 2]


def test_drop_total_rows():
    df = pd.DataFrame({"product": ["Латте", "ИТОГО", "Всего за месяц", None], "amount": [1, 2, 3, 4]})
    cleaned, removed = drop_total_rows(df)
    assert removed == 2
    assert list(cleaned["amount"]) == [1, 4]


def test_fill_missing_amount():
    df = pd.DataFrame({"qty": [2.0, 1.0, None], "price": [100.0, 50.0, 80.0],
                       "amount": [200.0, None, None]})
    filled, count = fill_missing_amount(df)
    assert count == 1
    assert filled["amount"].tolist()[:2] == [200.0, 50.0]
    assert pd.isna(filled["amount"].iloc[2])  # без количества не угадать


def test_remove_duplicates():
    df = pd.DataFrame({"receipt": ["1", "1", "1"], "product": ["Латте", "Латте", "Раф"]})
    cleaned, removed = remove_duplicates(df)
    assert removed == 1
    assert len(cleaned) == 2


def test_filter_month_and_detect_month():
    df = pd.DataFrame({"date": pd.to_datetime(["2026-09-30", "2026-10-01", "2026-10-31", "2026-11-01"])})
    assert detect_month(df) == "2026-10"
    kept, removed = filter_month(df, "2026-10")
    assert removed == 2
    assert kept["date"].dt.month.unique().tolist() == [10]


def _write_xlsx(path, rows):
    wb = Workbook()
    for row in rows:
        wb.active.append(row)
    wb.save(path)


def test_clean_file_end_to_end(tmp_path):
    path = tmp_path / "Тест.xlsx"
    _write_xlsx(path, [
        ["Выгрузка из кассы"],
        [],
        ["Товар", "comment", "ДАТА", "Кол.", "Цена, руб.", "Итого", "Чек"],
        ["капуччино", "", "01.10.26", 2, "220 ₽", "440,00", "A-1"],
        [],
        ["LATTE", "без сахара", "2026/10/01", 1, 240, None, "A-2"],
        ["LATTE", "без сахара", "2026/10/01", 1, 240, None, "A-2"],
        ["ИТОГО", None, None, None, None, 680],
    ])

    df, report = clean_file(path)

    assert report.rows_clean == 2
    assert df["product"].tolist() == ["Капучино", "Латте"]
    assert df["amount"].tolist() == [440.0, 240.0]
    assert df["branch"].unique().tolist() == ["Тест"]
    assert (df["date"] == pd.Timestamp("2026-10-01")).all()
    assert report.unknown_columns == ["comment"]
    assert report.issues["служебных строк над таблицей"] == 2
    assert report.issues["пустых строк"] == 1
    assert report.issues["строк «Итого»"] == 1
    assert report.issues["дублей удалено"] == 1
    assert report.issues["восстановленных сумм (кол-во × цена)"] == 2


def test_load_all_combines_branches(tmp_path):
    header = ["Дата", "Товар", "Кол-во", "Цена", "Сумма", "№ чека"]
    _write_xlsx(tmp_path / "Центр.xlsx", [header,
                                         ["2026-10-01", "Латте", 1, 240, 240, "1"],
                                         ["2026-09-30", "Латте", 1, 240, 240, "2"]])
    _write_xlsx(tmp_path / "Вокзал.xlsx", [header, ["02.10.2026", "раф", 2, 260, 520, "1"]])

    data, reports, month, outside = load_all(tmp_path)

    assert month == "2026-10"
    assert outside == 1
    assert len(reports) == 2
    assert sorted(data["branch"]) == ["Вокзал", "Центр"]
    assert data["amount"].sum() == 760


def test_load_all_empty_dir(tmp_path):
    with pytest.raises(FileNotFoundError):
        load_all(tmp_path)

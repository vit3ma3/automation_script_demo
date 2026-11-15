"""Чтение и очистка «грязных» выгрузок продаж филиалов.

Каждая функция делает одну вещь и легко тестируется отдельно.
`clean_file` собирает их в конвейер и попутно считает, что было исправлено.
"""

from __future__ import annotations

import difflib
import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

COLUMNS = ["receipt", "date", "branch", "product", "qty", "price", "amount"]

# Синонимы заголовков (в нижнем регистре, без «руб.» и лишних пробелов)
HEADER_SYNONYMS = {
    "receipt": {"№ чека", "номер чека", "чек", "check_no", "receipt"},
    "date": {"дата", "дата продажи", "date"},
    "product": {"товар", "наименование", "позиция", "product"},
    "qty": {"кол-во", "количество", "кол.", "qty", "шт"},
    "price": {"цена", "цена за шт", "price"},
    "amount": {"сумма", "итого", "total", "выручка"},
}

CANONICAL_PRODUCTS = [
    "Эспрессо", "Американо", "Капучино", "Латте", "Раф", "Флэт уайт", "Какао",
    "Чай чёрный", "Чай зелёный", "Круассан", "Чизкейк", "Маффин",
]

# То, что не исправить по похожести написания: другой язык или лишние слова
PRODUCT_ALIASES = {
    "americano": "Американо",
    "latte": "Латте",
    "flat white": "Флэт уайт",
    "раф кофе": "Раф",
    "какава": "Какао",
}

DATE_FORMATS = [
    "%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S",
    "%d.%m.%Y", "%d.%m.%y", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y",
]

EXCEL_EPOCH = date(1899, 12, 30)
TOTAL_MARKERS = ("итого", "всего", "total")


@dataclass
class FileReport:
    """Что было найдено и исправлено в одном файле."""

    file: str
    rows_read: int = 0
    rows_clean: int = 0
    issues: Counter = field(default_factory=Counter)
    product_fixes: Counter = field(default_factory=Counter)
    unknown_columns: list[str] = field(default_factory=list)


# --- Заголовки ---------------------------------------------------------------

def _key(text) -> str:
    """Нижний регистр, ё→е, без лишних пробелов."""
    return re.sub(r"\s+", " ", str(text)).strip().lower().replace("ё", "е")


def normalize_header(name) -> str | None:
    """Название колонки из файла -> каноническое имя или None, если не знаем."""
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return None
    key = _key(name)
    key = re.sub(r",?\s*(руб\.?|₽|р\.)$", "", key).strip()
    for canonical, synonyms in HEADER_SYNONYMS.items():
        if key in {_key(s) for s in synonyms}:
            return canonical
    return None


def find_header_row(raw: pd.DataFrame, min_matches: int = 3) -> int:
    """Индекс первой строки, похожей на шапку таблицы (над ней — заголовки отчёта)."""
    for idx, row in raw.iterrows():
        if sum(normalize_header(v) is not None for v in row) >= min_matches:
            return idx
    raise ValueError("не найдена строка с заголовками колонок")


# --- Строки ------------------------------------------------------------------

def _is_blank(value) -> bool:
    return value is None or (isinstance(value, float) and pd.isna(value)) \
        or (isinstance(value, str) and not value.strip())


def drop_empty_rows(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Удаляет строки, где все ячейки пустые или из пробелов."""
    mask = df.apply(lambda row: all(_is_blank(v) for v in row), axis=1)
    return df[~mask], int(mask.sum())


def drop_total_rows(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Удаляет строки «Итого»/«Всего», которые кассы добавляют в конец выгрузки."""
    mask = df["product"].map(lambda v: isinstance(v, str) and _key(v).startswith(TOTAL_MARKERS))
    return df[~mask], int(mask.sum())


# --- Значения ----------------------------------------------------------------

def parse_date(value) -> pd.Timestamp:
    """Дата из строки любого известного формата, datetime или серийного номера Excel.

    Неоднозначные даты вида 05/10/26 читаем как день/месяц/год (российский формат).
    """
    if _is_blank(value):
        return pd.NaT
    if isinstance(value, (datetime, date)):
        return pd.Timestamp(value).normalize()
    if isinstance(value, (int, float)):
        return pd.Timestamp(EXCEL_EPOCH + timedelta(days=int(value)))
    text = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return pd.Timestamp(datetime.strptime(text, fmt))
        except ValueError:
            continue
    return pd.NaT


def to_number(value) -> float:
    """Число из «1 250,00», «240 ₽», «190 руб.» и т.п.; NaN, если не получилось."""
    if _is_blank(value):
        return float("nan")
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("\xa0", "").replace(" ", "")
    text = re.sub(r"(₽|руб\.?|р\.)$", "", text, flags=re.IGNORECASE)
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return float("nan")


def normalize_product(name, cutoff: float = 0.75) -> str | None:
    """Приводит название товара к справочнику: регистр, пробелы, опечатки."""
    if _is_blank(name):
        return None
    key = _key(name)
    canonical_by_key = {_key(p): p for p in CANONICAL_PRODUCTS}
    if key in canonical_by_key:
        return canonical_by_key[key]
    if key in PRODUCT_ALIASES:
        return PRODUCT_ALIASES[key]
    match = difflib.get_close_matches(key, canonical_by_key, n=1, cutoff=cutoff)
    if match:
        return canonical_by_key[match[0]]
    return re.sub(r"\s+", " ", str(name)).strip().capitalize()


def normalize_receipt(value) -> str | None:
    if _is_blank(value):
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def fill_missing_amount(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Восстанавливает пустую сумму как количество × цена."""
    mask = df["amount"].isna() & df["qty"].notna() & df["price"].notna()
    df = df.copy()
    df.loc[mask, "amount"] = df.loc[mask, "qty"] * df.loc[mask, "price"]
    return df, int(mask.sum())


def remove_duplicates(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Удаляет полностью совпадающие строки (повторная выгрузка одного чека)."""
    before = len(df)
    df = df.drop_duplicates()
    return df, before - len(df)


def filter_month(df: pd.DataFrame, month: str) -> tuple[pd.DataFrame, int]:
    """Оставляет только строки за месяц вида '2026-10'."""
    mask = df["date"].dt.to_period("M") == pd.Period(month, "M")
    return df[mask], int((~mask).sum())


def detect_month(df: pd.DataFrame) -> str:
    """Месяц, за который в данных больше всего строк."""
    return str(df["date"].dt.to_period("M").mode()[0])


# --- Конвейер ----------------------------------------------------------------

def clean_file(path: Path, branch: str | None = None) -> tuple[pd.DataFrame, FileReport]:
    """Читает один файл филиала и возвращает очищенную таблицу и отчёт об исправлениях."""
    path = Path(path)
    branch = branch or path.stem
    report = FileReport(file=path.name)

    raw = pd.read_excel(path, header=None, dtype=object)
    header_idx = find_header_row(raw)
    if header_idx:
        report.issues["служебных строк над таблицей"] = header_idx

    df = raw.iloc[header_idx + 1:].copy()
    names = [normalize_header(v) for v in raw.iloc[header_idx]]
    report.unknown_columns = [str(v) for v, n in zip(raw.iloc[header_idx], names)
                              if n is None and not _is_blank(v)]
    df.columns = [n or f"_drop{i}" for i, n in enumerate(names)]
    df = df[[n for n in names if n]]
    report.rows_read = len(df)

    df, n = drop_empty_rows(df)
    report.issues["пустых строк"] += n
    df, n = drop_total_rows(df)
    report.issues["строк «Итого»"] += n

    non_standard = df["date"].map(lambda v: not isinstance(v, (datetime, date)) and not _is_blank(v))
    report.issues["дат в текстовом/числовом формате"] += int(non_standard.sum())
    df["date"] = df["date"].map(parse_date)

    for col in ("qty", "price", "amount"):
        as_text = df[col].map(lambda v: isinstance(v, str))
        report.issues["чисел, записанных текстом"] += int(as_text.sum())
        df[col] = df[col].map(to_number)

    df, n = fill_missing_amount(df)
    report.issues["восстановленных сумм (кол-во × цена)"] += n

    raw_names = df["product"]
    df["product"] = raw_names.map(normalize_product)
    for before, after in zip(raw_names, df["product"]):
        if isinstance(before, str) and before != after:
            report.product_fixes[(before, after)] += 1
    report.issues["исправленных названий товаров"] += sum(report.product_fixes.values())

    df["receipt"] = df["receipt"].map(normalize_receipt) if "receipt" in df else None
    df["branch"] = branch

    invalid = df["date"].isna() | df["product"].isna() | df["amount"].isna()
    report.issues["строк без даты/товара/суммы удалено"] += int(invalid.sum())
    df = df[~invalid]

    df, n = remove_duplicates(df[COLUMNS])
    report.issues["дублей удалено"] += n

    report.rows_clean = len(df)
    return df, report


def load_all(input_dir: Path, month: str | None = None
             ) -> tuple[pd.DataFrame, list[FileReport], str, int]:
    """Чистит все .xlsx в папке. Возвращает данные, отчёты, месяц и число строк вне месяца."""
    files = sorted(p for p in Path(input_dir).glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        raise FileNotFoundError(f"в папке {input_dir} нет файлов .xlsx")

    frames, reports = [], []
    for path in files:
        df, report = clean_file(path)
        frames.append(df)
        reports.append(report)

    data = pd.concat(frames, ignore_index=True)
    month = month or detect_month(data)
    data, outside = filter_month(data, month)
    for col in ("qty", "price", "amount"):
        data[col] = data[col].astype(float)
    data = data.sort_values(["date", "branch", "receipt"]).reset_index(drop=True)
    return data, reports, month, outside

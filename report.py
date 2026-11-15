"""Сводный отчёт о продажах филиалов сети кофеен из «грязных» Excel-выгрузок.

Примеры:
    python report.py
    python report.py --input input --output output/report.xlsx --month 2026-10
    python report.py --telegram
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from coffee_report.cleaning import FileReport, load_all
from coffee_report.excel_report import build_report, month_title


def print_cleaning_log(reports: list[FileReport], outside: int, month: str) -> None:
    print("Очистка данных:")
    for r in reports:
        print(f"\n  {r.file}: прочитано {r.rows_read} строк, после очистки {r.rows_clean}")
        for issue, count in r.issues.items():
            if count:
                print(f"    • {issue}: {count}")
        if r.unknown_columns:
            print(f"    • пропущены лишние колонки: {', '.join(r.unknown_columns)}")
        if r.product_fixes:
            # Сначала опечатки, потом исправления регистра — они нагляднее
            fixes = [(b.strip(), a) for (b, a), _ in r.product_fixes.most_common() if b.strip() != a]
            fixes.sort(key=lambda fix: fix[0].lower() == fix[1].lower())
            examples = [f"«{b}» → {a}" for b, a in fixes[:4]]
            if examples:
                print(f"      например: {'; '.join(examples)}")

    total_read = sum(r.rows_read for r in reports)
    total_clean = sum(r.rows_clean for r in reports) - outside
    print(f"\n  Строк за пределами месяца {month} отброшено: {outside}")
    print(f"  Итого: {total_read} строк прочитано → {total_clean} чистых строк продаж\n")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--input", default="input", help="папка с файлами филиалов (по умолчанию input/)")
    parser.add_argument("--output", default="output/report.xlsx", help="путь к итоговому отчёту")
    parser.add_argument("--month", help="месяц ГГГГ-ММ; по умолчанию определяется по данным")
    parser.add_argument("--telegram", action="store_true", help="отправить сводку и файл в Telegram")
    args = parser.parse_args(argv)
    if args.month and not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", args.month):
        parser.error("--month должен быть в формате ГГГГ-ММ, например 2026-10")
    return args


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        data, reports, month, outside = load_all(Path(args.input), args.month)
    except (FileNotFoundError, ValueError) as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        return 1

    print_cleaning_log(reports, outside, month)
    if data.empty:
        print(f"Ошибка: нет данных за {month}", file=sys.stderr)
        return 1

    summary = build_report(data, month, Path(args.output))
    print(f"Отчёт за {month_title(month).lower()} сохранён: {args.output}")
    revenue = f"{summary['revenue']:,.0f}".replace(",", " ")
    print(f"Выручка: {revenue} ₽, чеков: {summary['receipts']}")

    if args.telegram:
        from coffee_report.telegram import format_summary, send_report
        try:
            send_report(format_summary(summary, month_title(month)), Path(args.output))
        except Exception as e:  # сеть, неверный токен и т.п. — отчёт уже сохранён
            print(f"Не удалось отправить в Telegram: {e}", file=sys.stderr)
            return 2
        print("Сводка и файл отправлены в Telegram")
    return 0


if __name__ == "__main__":
    sys.exit(main())

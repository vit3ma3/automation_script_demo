"""Отправка сводки и файла отчёта в Telegram через Bot API."""

from __future__ import annotations

import os
from pathlib import Path

import requests
from dotenv import load_dotenv

API = "https://api.telegram.org/bot{token}/{method}"


def _money(value: float) -> str:
    return f"{value:,.0f} ₽".replace(",", " ")


def format_summary(s: dict, period: str) -> str:
    leader = s["by_branch"].index[0]
    top_product = s["by_product"].index[0]
    lines = [
        f"☕ Отчёт о продажах — {period}",
        "",
        f"Выручка: {_money(s['revenue'])}",
        f"Чеков: {s['receipts']:,}".replace(",", " "),
        f"Средний чек: {_money(s['avg_check'])}",
        "",
        "По филиалам:",
        *(f"• {name}: {_money(r.revenue)} ({r.share:.0%})" for name, r in s["by_branch"].iterrows()),
        "",
        f"Лидер: {leader}. Хит продаж: {top_product}.",
    ]
    return "\n".join(lines)


def send_report(text: str, file_path: Path) -> None:
    """Отправляет текст и файл. Токен и chat_id берутся из .env."""
    load_dotenv()
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    chat_id = os.getenv("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        raise RuntimeError("задайте TELEGRAM_BOT_TOKEN и TELEGRAM_CHAT_ID в файле .env")

    _check(requests.post(API.format(token=token, method="sendMessage"),
                         data={"chat_id": chat_id, "text": text}, timeout=30))
    with open(file_path, "rb") as f:
        _check(requests.post(API.format(token=token, method="sendDocument"),
                             data={"chat_id": chat_id},
                             files={"document": (Path(file_path).name, f)}, timeout=60))


def _check(response: requests.Response) -> None:
    """Ошибка без URL запроса: в нём токен бота, ему не место в логах."""
    if not response.ok:
        try:
            reason = response.json().get("description", "")
        except ValueError:
            reason = response.reason
        raise RuntimeError(f"Telegram API {response.status_code}: {reason}")

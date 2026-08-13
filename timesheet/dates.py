"""Calculo de dias laborables y agregaciones del mes."""

from __future__ import annotations

from calendar import monthrange
from datetime import date, timedelta
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .models import State


def working_days(year: int, month: int) -> list[date]:
    """Devuelve los dias Lun-Vie del mes, ordenados."""
    if not (1 <= month <= 12):
        raise ValueError(f"mes fuera de rango: {month}")
    _, last = monthrange(year, month)
    out: list[date] = []
    for d in range(1, last + 1):
        cur = date(year, month, d)
        if cur.weekday() < 5:
            out.append(cur)
    return out


def is_working_day(d: date) -> bool:
    return d.weekday() < 5


def per_day_totals(state: "State") -> list[dict]:
    """Totales por dia laborable del mes activo."""
    wd = working_days(state.year, state.month)
    out: list[dict] = []
    for d in wd:
        iso = d.isoformat()
        day = state.days.get(iso)
        items = day.items if day else []
        items_total = round(sum(it.hours for it in items), 2)
        out.append({
            "date": iso,
            "items_total": items_total,
            "items_count": len(items),
            "entry": day.entry.strftime("%H:%M") if day and day.entry else None,
            "exit": day.exit.strftime("%H:%M") if day and day.exit else None,
        })
    return out


def month_total_hours(state: "State") -> float:
    total = 0.0
    for d in state.days.values():
        total += sum(it.hours for it in d.items)
    return round(total, 2)


def month_total_amount(state: "State") -> float:
    return round(month_total_hours(state) * state.professional.hourly_rate, 2)


def first_data_row() -> int:
    return 10


def base_data_row_count() -> int:
    return 22


def last_base_data_row() -> int:
    return first_data_row() + base_data_row_count() - 1  # 31


__all__ = [
    "working_days",
    "is_working_day",
    "per_day_totals",
    "month_total_hours",
    "month_total_amount",
    "first_data_row",
    "base_data_row_count",
    "last_base_data_row",
]

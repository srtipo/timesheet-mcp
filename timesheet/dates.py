"""Calculo de dias registrables y agregaciones por mes."""

from __future__ import annotations

from calendar import monthrange
from datetime import date, time
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .models import Day, State


DEFAULT_ENTRY: time = time(8, 0)


def _compute_exit(entry: time, total_hours: float) -> time:
    """Devuelve entry + total_hours, con wrap a 24h."""
    base_minutes = entry.hour * 60 + entry.minute
    extra = int(round(total_hours * 60))
    minutes = (base_minutes + extra) % (24 * 60)
    return time(minutes // 60, minutes % 60)


def effective_entry(day: "Day | None") -> time | None:
    """Entry efectiva: la guardada o DEFAULT_ENTRY si None."""
    if day is None:
        return None
    return day.entry if day.entry is not None else DEFAULT_ENTRY


def effective_exit(day: "Day | None") -> time | None:
    """Exit efectiva: la guardada, o entry + sum(hours) si None.

    Devuelve None si el dia no tiene items ni exit guardada.
    """
    if day is None:
        return None
    if day.exit is not None:
        return day.exit
    if not day.items:
        return None
    entry = day.entry if day.entry is not None else DEFAULT_ENTRY
    return _compute_exit(entry, sum(it.hours for it in day.items))


def today_month() -> tuple[int, int]:
    """Mes actual del sistema (today.year, today.month)."""
    t = date.today()
    return (t.year, t.month)


def month_of(iso_date: str) -> tuple[int, int]:
    """Mes calendario de una fecha ISO. Valida formato."""
    try:
        d = date.fromisoformat(iso_date)
    except ValueError as exc:
        raise ValueError(f"fecha invalida {iso_date!r}; formato esperado YYYY-MM-DD") from exc
    return (d.year, d.month)


def months_present(state: "State") -> list[tuple[int, int]]:
    """Meses (year, month) que tienen al menos un dia en state.days, ordenados."""
    seen: set[tuple[int, int]] = set()
    for iso in state.days.keys():
        try:
            seen.add(month_of(iso))
        except ValueError:
            continue
    return sorted(seen)


def trackable_days(year: int, month: int) -> list[date]:
    """Devuelve todos los dias del mes, ordenados.

    La verificacion de mes/anio se hace en callers (state.py / server.py).
    """
    if not (1 <= month <= 12):
        raise ValueError(f"mes fuera de rango: {month}")
    if not (2000 <= year <= 2100):
        raise ValueError(f"year fuera de rango: {year}")
    _, last = monthrange(year, month)
    return [date(year, month, d) for d in range(1, last + 1)]


def is_trackable_day(d: date) -> bool:
    """Indica si el dia es registrable. Cualquier dia del calendario es valido."""
    return True


def per_day_totals(state: "State", year: int, month: int) -> list[dict]:
    """Totales por dia de un mes especifico.

    Para `entry`/`exit` se aplica la politica de defaults:
    - Si el dia tiene items, entry=08:00 por default y exit=entry+sum(hours).
    - Si el dia no tiene items, entry/exit son los guardados (o None si no hay).
    - Overrides explicitos via set_day_metadata siempre prevalecen.
    """
    td = trackable_days(year, month)
    out: list[dict] = []
    for d in td:
        iso = d.isoformat()
        day = state.days.get(iso)
        items = day.items if day else []
        items_total = round(sum(it.hours for it in items), 2)
        if day and items:
            entry_eff = effective_entry(day)
            exit_eff = effective_exit(day)
        else:
            entry_eff = day.entry if day else None
            exit_eff = day.exit if day else None
        out.append({
            "date": iso,
            "items_total": items_total,
            "items_count": len(items),
            "entry": entry_eff.strftime("%H:%M") if entry_eff else None,
            "exit": exit_eff.strftime("%H:%M") if exit_eff else None,
        })
    return out


def month_total_hours(state: "State", year: int, month: int) -> float:
    """Suma de horas del mes (year, month)."""
    total = 0.0
    td = trackable_days(year, month)
    for d in td:
        day = state.days.get(d.isoformat())
        if day is None:
            continue
        total += sum(it.hours for it in day.items)
    return round(total, 2)


def month_total_amount(state: "State", year: int, month: int) -> float:
    return round(month_total_hours(state, year, month) * state.professional.hourly_rate, 2)


def first_data_row() -> int:
    return 10


def base_data_row_count() -> int:
    return 22


def last_base_data_row() -> int:
    return first_data_row() + base_data_row_count() - 1  # 31


working_days = trackable_days
is_working_day = is_trackable_day


__all__ = [
    "trackable_days",
    "is_trackable_day",
    "working_days",
    "is_working_day",
    "per_day_totals",
    "month_total_hours",
    "month_total_amount",
    "months_present",
    "today_month",
    "month_of",
    "first_data_row",
    "base_data_row_count",
    "last_base_data_row",
    "DEFAULT_ENTRY",
    "_compute_exit",
    "effective_entry",
    "effective_exit",
]

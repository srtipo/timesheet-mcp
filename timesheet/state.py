"""Persistencia y mutadores del estado del timesheet.

A partir de v0.3 el estado NO tiene un mes activo. Cada entrada vive por su
fecha ISO en `state.days`; los meses se calculan a partir de esas fechas.
Como consecuencia, registrar o borrar entradas nunca puede borrar data de
otro mes: no existe la operacion "cambiar de mes".
"""

from __future__ import annotations

import json
import os
import re
import uuid
from datetime import date, datetime, time
from pathlib import Path
from typing import Any, Optional

from .dates import DEFAULT_ENTRY, _compute_exit
from .models import Day, DayItem, Professional, State, Supervisor


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_ROOT / "data"
STATE_FILE = DATA_DIR / "state.json"
TEMPLATE_DIR = PROJECT_ROOT / "template"
TEMPLATE_FILE = TEMPLATE_DIR / "Timesheet Modelo Mayo 2025 PRIS.xlsx"
EXPORTS_DIR = PROJECT_ROOT / "exports"


class ValidationError(ValueError):
    """Error de validacion que se reporta al cliente como isError."""


_UNSET: Any = object()


def _ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)


def default_state() -> State:
    return State(
        schema_version=1,
        professional=Professional(
            name="Victor Acosta",
            specialty="Desarrollador Front-End",
            hourly_rate=3.5,
        ),
        supervisor=Supervisor(name="Raúl D. Olivero Carrucini"),
        days={},
    )


def _migrate_assign_ids(state: State) -> bool:
    """Asigna UUID4 a items sin id. Devuelve True si hubo cambios."""
    changed = False
    for day in state.days.values():
        for item in day.items:
            if not item.id:
                item.id = str(uuid.uuid4())
                changed = True
    return changed


def _migrate_drop_legacy_fields(payload: dict) -> tuple[dict, bool]:
    """Quita campos legacy (year, month) si aparecen en el JSON guardado.

    Esto permite cargar archivos state.json de v0.2 sin error.
    Los dias se conservan porque viven en su propia fecha.

    Devuelve (payload, changed). Si changed=True el archivo debe reescribirse.
    """
    changed = False
    for k in ("year", "month"):
        if k in payload:
            del payload[k]
            changed = True
    return payload, changed


def load_state() -> State:
    """Carga el estado desde data/state.json; crea defaults si no existe.

    Aplica migraciones on-load:
    - Si hay campos legacy `year`/`month` se descartan silenciosamente.
    - Si algun item carece de id, se le asigna UUID4 y se reescribe atomicamente.
    """
    _ensure_dirs()
    if not STATE_FILE.exists():
        st = default_state()
        save_state(st)
        return st
    raw = STATE_FILE.read_text(encoding="utf-8")
    data: dict[str, Any] = json.loads(raw)
    data, legacy_dropped = _migrate_drop_legacy_fields(data)
    state = State.model_validate(data)
    dirty = legacy_dropped
    if _migrate_assign_ids(state):
        dirty = True
    if dirty:
        save_state(state)
    return state


def save_state(state: State) -> None:
    """Escribe el estado a disco de forma atomica (.tmp + os.replace)."""
    _ensure_dirs()
    payload = state.model_dump(mode="json", by_alias=True)
    tmp = STATE_FILE.with_suffix(".json.tmp")
    tmp.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    os.replace(tmp, STATE_FILE)


# ---------------------------------------------------------------------------
# Validadores / helpers puros
# ---------------------------------------------------------------------------

def _ensure_day_defaults(day: Day) -> None:
    """Aplica defaults in-place: entry=08:00 si None, exit=entry+sum(hours).

    Entry: solo se setea si es None (override explicito del usuario preservado).
    Exit: se recalcula como entry+sum(hours) en cada llamada, salvo que
    exit_locked=True (caso en que el usuario fijo la salida via set_day_metadata
    y futuras mutaciones de items no la deben tocar).
    """
    if day.entry is None:
        day.entry = DEFAULT_ENTRY
    if not day.exit_locked:
        total = sum(it.hours for it in day.items)
        day.exit = _compute_exit(day.entry, total)


_HHMM_RE = re.compile(r"^([01]?\d|2[0-3]):([0-5]\d)$")


def normalize_hhmm(value: Any) -> Optional[str]:
    """Acepta '8:00' o '08:00' y devuelve '08:00' con cero-padding. None -> None."""
    if value is None:
        return None
    if not isinstance(value, str):
        raise ValidationError("hora debe ser string 'HH:MM' o null")
    s = value.strip()
    if not s:
        return None
    m = _HHMM_RE.match(s)
    if not m:
        raise ValidationError(f"hora invalida {value!r}; formato esperado HH:MM")
    hh, mm = m.group(1), m.group(2)
    return f"{int(hh):02d}:{mm}"


def parse_hhmm(value: Any) -> Optional[time]:
    s = normalize_hhmm(value)
    if s is None:
        return None
    return datetime.strptime(s, "%H:%M").time()


def assert_iso_date(s: str) -> date:
    """Valida que `s` sea una fecha ISO valida (YYYY-MM-DD).

    No valida contra ningun mes activo: el estado es multi-mes.
    """
    if not isinstance(s, str):
        raise ValidationError("date debe ser string 'YYYY-MM-DD'")
    try:
        return date.fromisoformat(s)
    except ValueError as exc:
        raise ValidationError(
            f"fecha invalida {s!r}; formato esperado YYYY-MM-DD"
        ) from exc


def assert_stripped_nonempty(value: Any, field: str) -> str:
    if not isinstance(value, str):
        raise ValidationError(f"{field} debe ser string")
    s = value.strip()
    if not s:
        raise ValidationError(f"{field} no puede quedar vacio tras strip")
    return s


def assert_positive_number(value: Any, field: str, *, ge: float = 0) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValidationError(f"{field} debe ser numero")
    f = float(value)
    if f < ge:
        raise ValidationError(f"{field} debe ser >= {ge}")
    return round(f, 2)


def assert_year_month(year: int, month: int) -> None:
    if not (2000 <= year <= 2100):
        raise ValidationError(f"year fuera de rango [2000, 2100]: {year}")
    if not (1 <= month <= 12):
        raise ValidationError(f"month fuera de rango [1, 12]: {month}")


def ensure_day(state: State, iso_date: str) -> Day:
    day = state.days.get(iso_date)
    if day is None:
        day = Day()
        state.days[iso_date] = day
    return day


# ---------------------------------------------------------------------------
# Mutadores de alto nivel (devuelven dicts de respuesta, sin tocar persistencia)
# ---------------------------------------------------------------------------

def apply_set_professional_info(
    state: State,
    *,
    name: Any = _UNSET,
    specialty: Any = _UNSET,
    hourly_rate: Any = _UNSET,
) -> dict:
    if name is not _UNSET:
        state.professional.name = assert_stripped_nonempty(name, "name")
    if specialty is not _UNSET:
        state.professional.specialty = assert_stripped_nonempty(specialty, "specialty")
    if hourly_rate is not _UNSET:
        rate = assert_positive_number(hourly_rate, "hourly_rate", ge=0)
        state.professional.hourly_rate = rate
    return {}


def apply_set_supervisor(
    state: State,
    *,
    name: Any = _UNSET,
    sup_date: Any = _UNSET,
) -> dict:
    if name is not _UNSET:
        state.supervisor.name = assert_stripped_nonempty(name, "supervisor.name")
    if sup_date is not _UNSET:
        if sup_date is None:
            state.supervisor.sup_date = None
        elif isinstance(sup_date, str):
            try:
                state.supervisor.sup_date = date.fromisoformat(sup_date)
            except ValueError as exc:
                raise ValidationError(
                    f"supervisor.date invalido {sup_date!r}; formato esperado YYYY-MM-DD o null"
                ) from exc
        elif isinstance(sup_date, date):
            state.supervisor.sup_date = sup_date
        else:
            raise ValidationError("supervisor.date debe ser string 'YYYY-MM-DD' o null")
    return {}


def apply_set_day_metadata(
    state: State,
    *,
    iso_date: str,
    entry: Any = _UNSET,
    exit_: Any = _UNSET,
) -> dict:
    assert_iso_date(iso_date)
    day = ensure_day(state, iso_date)
    if entry is not _UNSET:
        day.entry = parse_hhmm(entry)
    if exit_ is not _UNSET:
        day.exit = parse_hhmm(exit_)
        day.exit_locked = True
    return {
        "date": iso_date,
        "entry": day.entry.strftime("%H:%M") if day.entry else None,
        "exit": day.exit.strftime("%H:%M") if day.exit else None,
    }


def apply_set_day_item(
    state: State,
    *,
    iso_date: str,
    description: str,
    hours: Any = _UNSET,
) -> dict:
    assert_iso_date(iso_date)
    desc = assert_stripped_nonempty(description, "description")
    day = ensure_day(state, iso_date)

    existing_idx = next(
        (i for i, it in enumerate(day.items) if it.description == desc), None
    )

    if existing_idx is not None and hours is _UNSET:
        existing = day.items[existing_idx]
        _ensure_day_defaults(day)
        return {
            "warning": "horas omitidas, item existente preservado",
            "item": existing.model_dump(),
            "date": iso_date,
            "items_total": round(sum(it.hours for it in day.items), 2),
        }

    if existing_idx is not None:
        h = assert_positive_number(hours, "hours", ge=0)
        existing = day.items[existing_idx]
        existing.hours = h
        _ensure_day_defaults(day)
        return {
            "item": existing.model_dump(),
            "date": iso_date,
            "items_total": round(sum(it.hours for it in day.items), 2),
        }

    if hours is _UNSET:
        other_sum = sum(it.hours for it in day.items)
        new_hours = max(0.0, 8.0 - other_sum)
        new_item = DayItem(description=desc, hours=round(new_hours, 2))
    else:
        h = assert_positive_number(hours, "hours", ge=0)
        new_item = DayItem(description=desc, hours=h)
    day.items.append(new_item)
    _ensure_day_defaults(day)
    return {
        "item": new_item.model_dump(),
        "date": iso_date,
        "items_total": round(sum(it.hours for it in day.items), 2),
    }


def apply_set_day_items(
    state: State,
    *,
    iso_date: str,
    items: list,
) -> dict:
    assert_iso_date(iso_date)
    day = ensure_day(state, iso_date)

    new_items: list[DayItem] = []
    for raw in items:
        if not isinstance(raw, dict):
            raise ValidationError("cada item debe ser un objeto")
        desc = assert_stripped_nonempty(raw.get("description", ""), "items[].description")
        h = raw.get("hours", 0)
        h_val = assert_positive_number(h, "items[].hours", ge=0)
        item_id = raw.get("id")
        if item_id is None or item_id == "":
            item_id = str(uuid.uuid4())
        elif not isinstance(item_id, str):
            raise ValidationError("items[].id debe ser string")
        new_items.append(DayItem(id=item_id, description=desc, hours=h_val))

    entry = day.entry
    exit_ = day.exit
    day.items = new_items
    day.entry = entry
    day.exit = exit_
    _ensure_day_defaults(day)

    return {
        "date": iso_date,
        "items_count": len(new_items),
        "items_total": round(sum(it.hours for it in new_items), 2),
    }


def apply_delete_day_item(state: State, *, iso_date: str, description: str) -> dict:
    assert_iso_date(iso_date)
    desc = assert_stripped_nonempty(description, "description")
    day = state.days.get(iso_date)
    if day is None:
        return {
            "warning": "item no encontrado",
            "data": {"date": iso_date, "items": []},
        }
    idx = next((i for i, it in enumerate(day.items) if it.description == desc), None)
    if idx is None:
        return {
            "warning": "item no encontrado",
            "data": {
                "date": iso_date,
                "items": [it.model_dump() for it in day.items],
            },
        }
    removed = day.items.pop(idx)
    return {
        "removed": removed.model_dump(),
        "date": iso_date,
        "items_total": round(sum(it.hours for it in day.items), 2),
    }


def apply_delete_day(state: State, *, iso_date: str) -> dict:
    assert_iso_date(iso_date)
    existed = state.days.pop(iso_date, None)
    if existed is None:
        return {"warning": "dia inexistente", "date": iso_date}
    return {"date": iso_date, "deleted": True}


def apply_move_item(
    state: State,
    *,
    from_date: str,
    to_date: str,
    item_id: str,
) -> dict:
    assert_iso_date(from_date)
    assert_iso_date(to_date)
    if not isinstance(item_id, str) or not item_id.strip():
        raise ValidationError("item_id debe ser string no vacio")
    src = state.days.get(from_date)
    if src is None:
        return {
            "warning": "item_id no encontrado",
            "data": {"date": from_date, "items": []},
        }
    idx = next((i for i, it in enumerate(src.items) if it.id == item_id), None)
    if idx is None:
        return {
            "warning": "item_id no encontrado",
            "data": {
                "date": from_date,
                "items": [it.model_dump() for it in src.items],
            },
        }
    moving = src.items.pop(idx)
    dst = ensure_day(state, to_date)
    dst.items.append(moving)
    return {
        "moved": moving.model_dump(),
        "from_date": from_date,
        "to_date": to_date,
        "from_items_total": round(sum(it.hours for it in src.items), 2),
        "to_items_total": round(sum(it.hours for it in dst.items), 2),
    }


__all__ = [
    "PROJECT_ROOT",
    "DATA_DIR",
    "STATE_FILE",
    "TEMPLATE_DIR",
    "TEMPLATE_FILE",
    "EXPORTS_DIR",
    "ValidationError",
    "DEFAULT_ENTRY",
    "default_state",
    "load_state",
    "save_state",
    "normalize_hhmm",
    "parse_hhmm",
    "assert_iso_date",
    "assert_year_month",
    "assert_stripped_nonempty",
    "assert_positive_number",
    "ensure_day",
    "_compute_exit",
    "_ensure_day_defaults",
    "apply_set_professional_info",
    "apply_set_supervisor",
    "apply_set_day_metadata",
    "apply_set_day_item",
    "apply_set_day_items",
    "apply_delete_day_item",
    "apply_delete_day",
    "apply_move_item",
]

"""Servidor MCP stdio con los 12 tools del timesheet.

A partir de v0.3 el estado es multi-mes (no hay "mes activo" persistente).
Cada tool opera contra fechas ISO arbitrarias; el mes por defecto de las
operaciones que lo requieren (get_timesheet, calculate_hours, export_to_excel)
se resuelve en el momento de la llamada segun los parametros o, en su defecto,
el mes actual del sistema.
"""

from __future__ import annotations

import asyncio
import json
import sys
from datetime import date
from pathlib import Path
from typing import Any, Callable

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import ImageContent, TextContent, Tool

import timesheet.state as state_mod
from timesheet.dates import (
    month_of,
    months_present,
    per_day_totals,
    today_month,
)
from timesheet.renderer import render as do_render
from timesheet.state import ValidationError, load_state, save_state

server = Server("timesheet-mcp")

# ---------------------------------------------------------------------------
# Handlers puros (testeables sin MCP)
# ---------------------------------------------------------------------------

def _ok(payload: dict, *, is_error: bool = False) -> list[TextContent]:
    txt = json.dumps(payload, ensure_ascii=False, indent=2)
    return [TextContent(type="text", text=txt)]


def _err(message: str) -> list[TextContent]:
    return _ok({"error": message, "isError": True}, is_error=True)


def _handle(name: str, args: dict) -> list[TextContent]:
    try:
        handler: Callable[[dict], dict] = HANDLERS[name]
        result = handler(args or {})
        if "isError" not in result:
            result["isError"] = False
        return _ok(result)
    except ValidationError as exc:
        return _err(str(exc))
    except KeyError as exc:
        return _err(f"tool desconocido: {exc}")


def _resolve_year_month(args: dict, key_year: str = "year", key_month: str = "month") -> tuple[int, int, str]:
    """Devuelve (year, month, source). Source es 'explicit' o 'today'.

    Si year o month faltan (None), cae al mes actual del sistema.
    Si solo uno de los dos viene explicito, se exige el otro (ValidationError).
    """
    y_arg = args.get(key_year)
    m_arg = args.get(key_month)
    if y_arg is None and m_arg is None:
        y, m = today_month()
        return y, m, "today"
    if y_arg is None or m_arg is None:
        raise ValidationError(
            f"{key_year} y {key_month} deben proveerse juntos (o ninguno, "
            "en cuyo caso se usa el mes actual del sistema)"
        )
    if not isinstance(y_arg, int) or isinstance(y_arg, bool):
        raise ValidationError(f"{key_year} debe ser entero")
    if not isinstance(m_arg, int) or isinstance(m_arg, bool):
        raise ValidationError(f"{key_month} debe ser entero")
    state_mod.assert_year_month(y_arg, m_arg)
    return y_arg, m_arg, "explicit"


# ---- 5.1 get_timesheet ----------------------------------------------------

def _get_timesheet(args: dict) -> dict:
    s = load_state()
    year, month, source = _resolve_year_month(args)
    per_day = _per_day(s, year, month)
    return {
        "active_month": {"year": year, "month": month, "source": source},
        "state": json.loads(
            json.dumps(s.model_dump(mode="json", by_alias=True), ensure_ascii=False)
        ),
        "computed": {
            "per_day": per_day,
            "month_total_hours": _month_hours(s, year, month),
            "month_total_amount": _month_amount(s, year, month),
            "months_present": _months_present(s),
        },
    }


# ---- 5.2 set_professional_info ------------------------------------------

def _set_professional_info(args: dict) -> dict:
    s = load_state()
    out = state_mod.apply_set_professional_info(
        s,
        name=args.get("name", state_mod._UNSET),
        specialty=args.get("specialty", state_mod._UNSET),
        hourly_rate=args.get("hourly_rate", state_mod._UNSET),
    )
    save_state(s)
    return {
        "professional": s.professional.model_dump(by_alias=True),
        **out,
    }


# ---- 5.3 set_supervisor --------------------------------------------------

def _set_supervisor(args: dict) -> dict:
    s = load_state()
    out = state_mod.apply_set_supervisor(
        s,
        name=args.get("name", state_mod._UNSET),
        sup_date=args.get("date", state_mod._UNSET),
    )
    save_state(s)
    return {"supervisor": s.supervisor.model_dump(mode="json", by_alias=True)}


# ---- 5.4 set_day_metadata ------------------------------------------------

def _set_day_metadata(args: dict) -> dict:
    if "date" not in args:
        raise ValidationError("date es requerido")
    s = load_state()
    out = state_mod.apply_set_day_metadata(
        s,
        iso_date=args["date"],
        entry=args.get("entry", state_mod._UNSET),
        exit_=args.get("exit", state_mod._UNSET),
    )
    save_state(s)
    return out


# ---- 5.5 set_day_item ----------------------------------------------------

def _set_day_item(args: dict) -> dict:
    if "date" not in args:
        raise ValidationError("date es requerido")
    if "description" not in args:
        raise ValidationError("description es requerido")
    s = load_state()
    out = state_mod.apply_set_day_item(
        s,
        iso_date=args["date"],
        description=args["description"],
        hours=args.get("hours", state_mod._UNSET),
    )
    save_state(s)
    return out


# ---- 5.6 set_day_items ---------------------------------------------------

def _set_day_items(args: dict) -> dict:
    if "date" not in args:
        raise ValidationError("date es requerido")
    if "items" not in args:
        raise ValidationError("items es requerido")
    s = load_state()
    out = state_mod.apply_set_day_items(
        s,
        iso_date=args["date"],
        items=args["items"],
    )
    save_state(s)
    return out


# ---- 5.7 delete_day_item -------------------------------------------------

def _delete_day_item(args: dict) -> dict:
    if "date" not in args:
        raise ValidationError("date es requerido")
    if "description" not in args:
        raise ValidationError("description es requerido")
    s = load_state()
    out = state_mod.apply_delete_day_item(
        s, iso_date=args["date"], description=args["description"]
    )
    save_state(s)
    return out


# ---- 5.8 delete_day ------------------------------------------------------

def _delete_day(args: dict) -> dict:
    if "date" not in args:
        raise ValidationError("date es requerido")
    s = load_state()
    out = state_mod.apply_delete_day(s, iso_date=args["date"])
    save_state(s)
    return out


# ---- 5.9 move_item -------------------------------------------------------

def _move_item(args: dict) -> dict:
    for k in ("from_date", "to_date", "item_id"):
        if k not in args:
            raise ValidationError(f"{k} es requerido")
    s = load_state()
    out = state_mod.apply_move_item(
        s,
        from_date=args["from_date"],
        to_date=args["to_date"],
        item_id=args["item_id"],
    )
    save_state(s)
    return out


# ---- 5.10 calculate_hours ------------------------------------------------

def _calculate_hours(args: dict) -> dict:
    s = load_state()
    year, month, source = _resolve_year_month(args)
    per_day = _per_day(s, year, month)
    return {
        "active_month": {"year": year, "month": month, "source": source},
        "per_day": [
            {**entry, "items": _items_for(s, entry["date"])}
            for entry in per_day
        ],
        "month_total_hours": _month_hours(s, year, month),
        "month_total_amount": _month_amount(s, year, month),
        "months_present": _months_present(s),
    }


# ---- 5.11 export_to_excel ------------------------------------------------

def _export_to_excel(args: dict) -> dict:
    year, month, source = _resolve_year_month(args)
    s = load_state()
    from calendar import monthrange
    _, days_in_month = monthrange(year, month)
    prefix = f"{year:04d}-{month:02d}-"
    has_data = any(
        iso.startswith(prefix) and s.days[iso].items
        for iso in s.days.keys()
    )
    if not has_data:
        return {
            "warning": f"no hay entradas con items para {year:04d}-{month:02d}",
            "year": year,
            "month": month,
            "days_with_data": 0,
            "days_in_month": days_in_month,
            "path": None,
            "isError": False,
        }
    out = do_render(
        state=s,
        year=year,
        month=month,
        out_path=Path(args["path"]) if args.get("path") else None,
    )
    out["active_month_source"] = source
    return out


# ---- 5.12 list_months ----------------------------------------------------

def _list_months(args: dict) -> dict:
    s = load_state()
    return {
        "months_present": _months_present(s),
        "today_month": {"year": today_month()[0], "month": today_month()[1]},
        "isError": False,
    }


# ---------------------------------------------------------------------------
# Helpers locales (evitan import circular con renderer/dates)
# ---------------------------------------------------------------------------

def _per_day(s, year: int, month: int) -> list[dict]:
    return per_day_totals(s, year, month)


def _month_hours(s, year: int, month: int) -> float:
    total = 0.0
    for entry in per_day_totals(s, year, month):
        total += entry["items_total"]
    return round(total, 2)


def _month_amount(s, year: int, month: int) -> float:
    return round(_month_hours(s, year, month) * s.professional.hourly_rate, 2)


def _months_present(s) -> list[dict]:
    out = []
    for year, month in months_present(s):
        count = sum(
            1 for iso in s.days.keys() if month_of(iso) == (year, month)
        )
        out.append({"year": year, "month": month, "day_count": count})
    return out


def _items_for(s, iso: str) -> list[dict]:
    day = s.days.get(iso)
    if day is None:
        return []
    return [
        {"id": it.id, "description": it.description, "hours": it.hours}
        for it in day.items
    ]


HANDLERS: dict[str, Callable[[dict], dict]] = {
    "get_timesheet": _get_timesheet,
    "set_professional_info": _set_professional_info,
    "set_supervisor": _set_supervisor,
    "set_day_metadata": _set_day_metadata,
    "set_day_item": _set_day_item,
    "set_day_items": _set_day_items,
    "delete_day_item": _delete_day_item,
    "delete_day": _delete_day,
    "move_item": _move_item,
    "calculate_hours": _calculate_hours,
    "export_to_excel": _export_to_excel,
    "list_months": _list_months,
}


# ---------------------------------------------------------------------------
# Tool schemas (JSON Schema para inputSchema)
# ---------------------------------------------------------------------------

def _schema(props: dict, required: list[str]) -> dict:
    return {
        "type": "object",
        "properties": props,
        "required": required,
        "additionalProperties": False,
    }


def _string(description: str, *, enum: list[str] | None = None) -> dict:
    s = {"type": "string", "description": description}
    if enum is not None:
        s["enum"] = enum
    return s


def _integer(description: str, *, minimum: int | None = None, maximum: int | None = None) -> dict:
    s: dict = {"type": "integer", "description": description}
    if minimum is not None:
        s["minimum"] = minimum
    if maximum is not None:
        s["maximum"] = maximum
    return s


def _number(description: str, *, minimum: float | None = None) -> dict:
    s: dict = {"type": "number", "description": description}
    if minimum is not None:
        s["minimum"] = minimum
    return s


def _null_or(t: dict) -> dict:
    return {"anyOf": [t, {"type": "null"}]}


TOOLS: list[Tool] = [
    Tool(
        name="get_timesheet",
        description=(
            "Devuelve el estado completo del timesheet (professional, supervisor, "
            "todos los dias de todos los meses) mas un bloque computed con el "
            "mes activo (filtrado). Sin year/month usa el mes actual del sistema."
        ),
        inputSchema=_schema(
            {
                "year": _integer("Filtra al anio indicado (2000-2100). Opcional.", minimum=2000, maximum=2100),
                "month": _integer("Filtra al mes indicado (1-12). Opcional.", minimum=1, maximum=12),
            },
            [],
        ),
    ),
    Tool(
        name="set_professional_info",
        description=(
            "Edita nombre, especialidad o tarifa del profesional. "
            "Cambiar estos campos NUNCA borra entradas existentes: el estado "
            "es multi-mes y cada entrada vive por su fecha ISO."
        ),
        inputSchema=_schema(
            {
                "name": _string("Nombre completo del profesional"),
                "specialty": _string("Especialidad / puesto"),
                "hourly_rate": _number("Tarifa por hora (>= 0)", minimum=0),
            },
            [],
        ),
    ),
    Tool(
        name="set_supervisor",
        description="Edita el nombre y/o la fecha del supervisor.",
        inputSchema=_schema(
            {
                "name": _string("Nombre del supervisor"),
                "date": _null_or(_string("Fecha del supervisor (YYYY-MM-DD) o null")),
            },
            [],
        ),
    ),
    Tool(
        name="set_day_metadata",
        description=(
            "Edita la metadata de un dia (entry/exit en formato HH:MM). "
            "El parametro no provisto conserva su valor actual. Pasar null lo limpia. "
            "Acepta cualquier fecha ISO (multi-mes)."
        ),
        inputSchema=_schema(
            {
                "date": _string("Fecha YYYY-MM-DD"),
                "entry": _null_or(_string("Hora de entrada HH:MM")),
                "exit": _null_or(_string("Hora de salida HH:MM")),
            },
            ["date"],
        ),
    ),
    Tool(
        name="set_day_item",
        description=(
            "Upsert de un item en el dia: si la descripcion (strippeada) ya existe, "
            "actualiza las horas del primer match; si no, crea un item. "
            "Si hours se omite y la descripcion no existe, se calcula como "
            "max(0, 8 - sum(otros.hours)). Si hours se omite y la descripcion existe, "
            "se preserva el item y se devuelve warning. Acepta cualquier fecha ISO."
        ),
        inputSchema=_schema(
            {
                "date": _string("Fecha YYYY-MM-DD"),
                "description": _string("Descripcion del item (no vacia tras strip)"),
                "hours": _number("Horas trabajadas (>= 0). Opcional.", minimum=0),
            },
            ["date", "description"],
        ),
    ),
    Tool(
        name="set_day_items",
        description=(
            "Reemplaza la lista completa de items del dia. Conserva entry/exit. "
            "items=[] deja la lista vacia pero conserva entry/exit. "
            "Acepta cualquier fecha ISO."
        ),
        inputSchema=_schema(
            {
                "date": _string("Fecha YYYY-MM-DD"),
                "items": {
                    "type": "array",
                    "description": "Lista de items {id?, description, hours?}",
                    "items": {
                        "type": "object",
                        "properties": {
                            "id": _string("UUID4 del item (opcional, se genera si falta)"),
                            "description": _string("Descripcion no vacia"),
                            "hours": _number("Horas (>= 0)", minimum=0),
                        },
                        "required": ["description"],
                    },
                },
            },
            ["date", "items"],
        ),
    ),
    Tool(
        name="delete_day_item",
        description=(
            "Borra el primer item cuya descripcion (strippeada) coincida en el dia. "
            "Si no hay match, devuelve warning. Acepta cualquier fecha ISO."
        ),
        inputSchema=_schema(
            {
                "date": _string("Fecha YYYY-MM-DD"),
                "description": _string("Descripcion del item a borrar"),
            },
            ["date", "description"],
        ),
    ),
    Tool(
        name="delete_day",
        description=(
            "Borra el dia completo (items + entry + exit). Si no existe, warning. "
            "Acepta cualquier fecha ISO."
        ),
        inputSchema=_schema(
            {"date": _string("Fecha YYYY-MM-DD")},
            ["date"],
        ),
    ),
    Tool(
        name="move_item",
        description=(
            "Mueve el item identificado por item_id desde from_date a to_date, "
            "preservando descripcion y horas. Conserva entry/exit de from_date aunque "
            "la lista de items quede vacia. Acepta cualquier par de fechas ISO."
        ),
        inputSchema=_schema(
            {
                "from_date": _string("Fecha origen YYYY-MM-DD"),
                "to_date": _string("Fecha destino YYYY-MM-DD"),
                "item_id": _string("UUID4 del item a mover"),
            },
            ["from_date", "to_date", "item_id"],
        ),
    ),
    Tool(
        name="calculate_hours",
        description=(
            "Calcula totales por dia y por mes sin escribir nada. "
            "Sin year/month usa el mes actual del sistema. "
            "Devuelve months_present para saber que meses tienen datos."
        ),
        inputSchema=_schema(
            {
                "year": _integer("Filtra al anio indicado (2000-2100). Opcional.", minimum=2000, maximum=2100),
                "month": _integer("Filtra al mes indicado (1-12). Opcional.", minimum=1, maximum=12),
            },
            [],
        ),
    ),
    Tool(
        name="export_to_excel",
        description=(
            "Renderiza el mes indicado del estado a un .xlsx respetando la "
            "plantilla PRIS. year y month son REQUERIDOS: si no se pasan, "
            "se usa el mes actual del sistema. Si el mes no tiene entradas, "
            "devuelve warning sin escribir archivo."
        ),
        inputSchema=_schema(
            {
                "year": _integer("Anio a exportar (2000-2100). Opcional: si falta junto con month, usa el mes actual.", minimum=2000, maximum=2100),
                "month": _integer("Mes a exportar (1-12). Opcional: si falta junto con year, usa el mes actual.", minimum=1, maximum=12),
                "path": _string("Ruta absoluta o relativa del .xlsx de salida (opcional)"),
            },
            [],
        ),
    ),
    Tool(
        name="list_months",
        description=(
            "Lista los meses (year, month) que tienen al menos una entrada "
            "en state.days, ordenados cronologicamente. Tambien devuelve "
            "el mes actual del sistema."
        ),
        inputSchema=_schema({}, []),
    ),
]


# ---------------------------------------------------------------------------
# MCP server wiring
# ---------------------------------------------------------------------------

@server.list_tools()
async def list_tools() -> list[Tool]:
    return TOOLS


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent | ImageContent]:
    return _handle(name, arguments)


async def _main() -> None:
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


def main() -> None:
    asyncio.run(_main())


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--smoke":
        # Modo de import/smoke: carga el modulo, valida handlers, no inicia stdio.
        print(json.dumps({
            "tools": [t.name for t in TOOLS],
            "handlers": sorted(HANDLERS.keys()),
        }, ensure_ascii=False, indent=2))
    else:
        main()

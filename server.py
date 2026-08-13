"""Servidor MCP stdio con los 11 tools del timesheet."""

from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path
from typing import Any, Callable

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import ImageContent, TextContent, Tool

import timesheet.state as state_mod
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


# ---- 5.1 get_timesheet ----------------------------------------------------

def _get_timesheet(args: dict) -> dict:
    s = load_state()
    return {
        "state": json.loads(
            json.dumps(s.model_dump(mode="json", by_alias=True), ensure_ascii=False)
        ),
        "computed": {
            "per_day": _per_day(s),
            "month_total_hours": _month_hours(s),
            "month_total_amount": _month_amount(s),
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
        year=args.get("year", state_mod._UNSET),
        month=args.get("month", state_mod._UNSET),
    )
    save_state(s)
    return {
        "year": s.year,
        "month": s.month,
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
    per_day = []
    for d in state_mod.working_days(s.year, s.month):
        iso = d.isoformat()
        day = s.days.get(iso)
        per_day.append({
            "date": iso,
            "items_total": round(sum(it.hours for it in (day.items if day else [])), 2),
            "items_count": len(day.items) if day else 0,
            "entry": day.entry.strftime("%H:%M") if day and day.entry else None,
            "exit": day.exit.strftime("%H:%M") if day and day.exit else None,
            "items": [
                {"id": it.id, "description": it.description, "hours": it.hours}
                for it in (day.items if day else [])
            ],
        })
    return {
        "per_day": per_day,
        "month_total_hours": _month_hours(s),
        "month_total_amount": _month_amount(s),
    }


# ---- 5.11 export_to_excel ------------------------------------------------

def _export_to_excel(args: dict) -> dict:
    path = args.get("path")
    out = do_render(out_path=Path(path) if path else None)
    return out


# ---------------------------------------------------------------------------
# Helpers locales (evitan import circular con renderer/dates)
# ---------------------------------------------------------------------------

def _per_day(s) -> list[dict]:
    from timesheet.dates import working_days as wd_func
    out = []
    for d in wd_func(s.year, s.month):
        iso = d.isoformat()
        day = s.days.get(iso)
        items = day.items if day else []
        out.append({
            "date": iso,
            "items_total": round(sum(it.hours for it in items), 2),
            "items_count": len(items),
            "entry": day.entry.strftime("%H:%M") if day and day.entry else None,
            "exit": day.exit.strftime("%H:%M") if day and day.exit else None,
        })
    return out


def _month_hours(s) -> float:
    total = 0.0
    for day in s.days.values():
        total += sum(it.hours for it in day.items)
    return round(total, 2)


def _month_amount(s) -> float:
    return round(_month_hours(s) * s.professional.hourly_rate, 2)


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
            "Devuelve el estado completo del timesheet (professional, supervisor, days) "
            "mas un bloque computed con totales por dia y mes."
        ),
        inputSchema=_schema({}, []),
    ),
    Tool(
        name="set_professional_info",
        description=(
            "Edita nombre, especialidad, tarifa, anio o mes del profesional. "
            "Si cambia year o month, descarta TODOS los dias del mes anterior "
            "y los lista en discarded_days."
        ),
        inputSchema=_schema(
            {
                "name": _string("Nombre completo del profesional"),
                "specialty": _string("Especialidad / puesto"),
                "year": _integer("Anio (2000-2100)", minimum=2000, maximum=2100),
                "month": _integer("Mes (1-12)", minimum=1, maximum=12),
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
            "El parametro no provisto conserva su valor actual. Pasar null lo limpia."
        ),
        inputSchema=_schema(
            {
                "date": _string("Fecha YYYY-MM-DD (laborable del mes activo)"),
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
            "se preserva el item y se devuelve warning."
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
            "items=[] deja la lista vacia pero conserva entry/exit."
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
            "Si no hay match, devuelve warning."
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
        description="Borra el dia completo (items + entry + exit). Si no existe, warning.",
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
            "la lista de items quede vacia. No valida colisiones en to_date."
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
            "Incluye los items con id por dia."
        ),
        inputSchema=_schema({}, []),
    ),
    Tool(
        name="export_to_excel",
        description=(
            "Renderiza el estado a un .xlsx respetando la plantilla PRIS. "
            "Si no se pasa path, usa exports/timesheet-YYYY-MM.xlsx. "
            "Si el mes tiene > 22 laborables, inserta filas y parchea formulas."
        ),
        inputSchema=_schema(
            {"path": _string("Ruta absoluta o relativa del .xlsx de salida (opcional)")},
            [],
        ),
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

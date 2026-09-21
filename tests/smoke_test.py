"""Smoke test: ejecuta los casos del spec + extras de borde.

Adaptado a v0.3 (estado multi-mes, sin mes activo persistente).
"""

from __future__ import annotations

import json
import shutil
import sys
import uuid
import zipfile
from datetime import date
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import load_workbook  # noqa: E402

from timesheet.state import (  # noqa: E402
    DATA_DIR,
    EXPORTS_DIR,
    STATE_FILE,
    TEMPLATE_FILE,
    default_state,
    load_state,
    save_state,
)
import server as server_mod  # noqa: E402
from server import HANDLERS  # noqa: E402


# ---------------------------------------------------------------------------
# Mini test framework
# ---------------------------------------------------------------------------

class TestFailed(AssertionError):
    pass


_RESULTS: list[tuple[str, bool, str]] = []


def _reset() -> None:
    if STATE_FILE.exists():
        STATE_FILE.unlink()
    if EXPORTS_DIR.exists():
        for f in EXPORTS_DIR.glob("*.xlsx"):
            f.unlink()
    save_state(default_state())


def _check(name: str, cond: bool, detail: str = "") -> None:
    status = "OK " if cond else "FAIL"
    _RESULTS.append((name, cond, detail))
    print(f"  [{status}] {name}{(' :: ' + detail) if detail and not cond else ''}")
    if not cond:
        raise TestFailed(f"{name}: {detail}")


def _call(tool: str, **kwargs) -> dict:
    """Invoca un handler del MCP y devuelve el dict (ya parseado de JSON)."""
    contents = server_mod._handle(tool, kwargs)
    assert len(contents) == 1
    return json.loads(contents[0].text)


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

def test_01_export_clean_template() -> None:
    print("\n[01] export_to_excel(year=2026, month=6) con 1 entrada 0h -> xlsx con totales 0")
    _reset()
    # Placeholder para que el mes "tenga data" (necesario en v0.3 para que se
    # genere el archivo; un mes sin entradas devuelve warning sin crear xlsx).
    _call("set_day_item", date="2026-06-01", description="Placeholder", hours=0)
    out = _call("export_to_excel", year=2026, month=6,
                path=str(EXPORTS_DIR / "t01.xlsx"))
    _check("01.1 archivo existe", Path(out["path"]).exists())
    _check("01.2 isError False", out.get("isError") is False)
    _check("01.3 month_total_hours = 0", out["month_total_hours"] == 0)
    _check("01.4 inserted_rows = 0 (n=1 <= 22)",
           out["inserted_rows"] == 0)
    _check("01.5 days_with_data = 1", out["days_with_data"] == 1)
    _check("01.6 days_in_month = 30", out["days_in_month"] == 30)
    with zipfile.ZipFile(out["path"]) as zf:
        has_media = any(n.startswith("xl/media/") for n in zf.namelist())
    _check("01.7 imagen preservada en el zip (con o sin plan B)", has_media)


def test_02_set_day_item_default_hours() -> None:
    print("\n[02] set_day_item('2026-06-01', 'Tarea A') -> hours=8 con id UUID4")
    _reset()
    out = _call("set_day_item", date="2026-06-01", description="Tarea A")
    _check("02.1 sin warning", "warning" not in out)
    _check("02.2 hours = 8", out["item"]["hours"] == 8, str(out["item"]))
    uuid.UUID(out["item"]["id"])  # raises if invalid
    _check("02.3 id es UUID4", True)
    _check("02.4 items_total = 8", out["items_total"] == 8)
    s = load_state()
    d = s.days["2026-06-01"]
    _check("02.5 entry default persistido = 08:00",
           d.entry.strftime("%H:%M") == "08:00", str(d.entry))
    _check("02.6 exit default persistido = 16:00 (08:00 + 8h)",
           d.exit.strftime("%H:%M") == "16:00", str(d.exit))


def test_03_set_day_item_explicit_hours() -> None:
    print("\n[03] set_day_item('2026-06-01', 'Tarea B', hours=3) -> total=11h")
    _reset()
    _call("set_day_item", date="2026-06-01", description="Tarea A")
    out = _call("set_day_item", date="2026-06-01", description="Tarea B", hours=3)
    _check("03.1 hours=3", out["item"]["hours"] == 3)
    _check("03.2 total=11h", out["items_total"] == 11)


def test_04_set_day_item_update_first() -> None:
    print("\n[04] set_day_item('2026-06-01', 'Tarea A', hours=2) -> A=2h, total=5h")
    _reset()
    _call("set_day_item", date="2026-06-01", description="Tarea A")
    _call("set_day_item", date="2026-06-01", description="Tarea B", hours=3)
    out = _call("set_day_item", date="2026-06-01", description="Tarea A", hours=2)
    _check("04.1 A=2h (update)", out["item"]["hours"] == 2)
    _check("04.2 total=5h (2+3)", out["items_total"] == 5)
    s = load_state()
    _check("04.3 dos items en el dia", len(s.days["2026-06-01"].items) == 2)


def test_05_export_with_five_days() -> None:
    print("\n[05] export_to_excel(year=2026, month=6) con 5 dias cargados, total=40")
    _reset()
    for iso, hrs in [
        ("2026-06-01", 8),
        ("2026-06-02", 8),
        ("2026-06-03", 8),
        ("2026-06-04", 8),
        ("2026-06-05", 8),
    ]:
        _call("set_day_item", date=iso, description="Trabajo", hours=hrs)
    out = _call("export_to_excel", year=2026, month=6,
                path=str(EXPORTS_DIR / "t05.xlsx"))
    wb = load_workbook(out["path"], data_only=False)
    ws = wb["HOJA LABOR REALIZADA"]
    ws1 = wb["FACT-SERV PROF"]
    # n=5 -> inserted=0 (5 <= 22). Totales en n+15 = fila 20.
    _check("05.0 inserted_rows = 0 (n=5 <= 22)",
           out["inserted_rows"] == 0, str(out))
    f20 = ws.cell(row=20, column=6).value
    _check("05.1 F20 (totales en n+15) = =SUM(F10:F14)",
           f20 == "=SUM(F10:F14)", repr(f20))
    _check("05.1b D12 sheet1 parcheado a F20",
           ws1["D12"].value == "='HOJA LABOR REALIZADA'!F20")
    f10_to_f14 = [ws.cell(row=r, column=6).value for r in range(10, 15)]
    _check("05.2 F10..F14 = 8", all(v == 8 for v in f10_to_f14), str(f10_to_f14))
    _check("05.3 days_with_data = 5", out["days_with_data"] == 5)
    _check("05.4 days_in_month = 30", out["days_in_month"] == 30)
    _check("05.5 last_day = 2026-06-05", out["last_day"] == "2026-06-05")


def test_06_set_professional_info_does_not_wipe() -> None:
    print("\n[06] set_professional_info NUNCA borra entradas (estado multi-mes)")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A")
    _call("set_day_item", date="2026-06-02", description="B")
    _call("set_day_item", date="2026-08-15", description="C")  # otro mes
    out = _call("set_professional_info", name="Otro Nombre",
                specialty="Otra Especialidad", hourly_rate=10)
    _check("06.1 no hay discarded_days en respuesta",
           "discarded_days" not in out, str(out))
    _check("06.2 professional.name actualizado",
           out["professional"]["name"] == "Otro Nombre")
    _check("06.3 hourly_rate actualizado",
           out["professional"]["hourly_rate"] == 10)
    s = load_state()
    _check("06.4 los 3 dias siguen presentes",
           all(k in s.days for k in ["2026-06-01", "2026-06-02", "2026-08-15"]))
    _check("06.5 junio intacto (2 dias con items)",
           sum(len(s.days[k].items) for k in ["2026-06-01", "2026-06-02"]) == 2)
    _check("06.6 agosto intacto (1 dia con item)",
           len(s.days["2026-08-15"].items) == 1)


def test_07_export_reload_image_merges() -> None:
    print("\n[07] re-abrir .xlsx: imagen, merges y formulas preservadas")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A")
    out = _call("export_to_excel", year=2026, month=6,
                path=str(EXPORTS_DIR / "t07.xlsx"))
    wb = load_workbook(out["path"])
    ws = wb["HOJA LABOR REALIZADA"]
    merges = [str(m) for m in ws.merged_cells.ranges]
    # n=1 -> inserted=0 -> layout base intacto
    _check("07.1 merge A6:B6 presente", "A6:B6" in merges, str(merges[:6]))
    _check("07.2 merge B10:C10 presente", "B10:C10" in merges, str(merges))
    _check("07.3 merge A47:B47 presente (sin shift, n=1)",
           "A47:B47" in merges, str(merges))
    _check("07.3b supervisor.name en A26 (n+25=26, sin shift)",
           ws.cell(row=26, column=1).value is not None
           and "Olivero" in str(ws.cell(row=26, column=1).value),
           repr(ws.cell(row=26, column=1).value))
    with zipfile.ZipFile(out["path"]) as zf:
        has_media = any(n.startswith("xl/media/") for n in zf.namelist())
    _check("07.4 imagen del logo presente en el zip", has_media)


def test_08_existing_desc_no_hours_is_noop() -> None:
    print("\n[08] set_day_item desc existente y hours=None -> no-op + warning")
    _reset()
    _call("set_day_item", date="2026-06-01", description="Tarea A", hours=4)
    out = _call("set_day_item", date="2026-06-01", description="Tarea A")
    _check("08.1 warning presente", out.get("warning", "").startswith("horas omitidas"))
    _check("08.2 item preservado con hours=4", out["item"]["hours"] == 4)
    _check("08.3 isError=False", out["isError"] is False)


def test_09_move_item_by_id() -> None:
    print("\n[09] move_item por item_id, from_date conserva entry/exit")
    _reset()
    _call("set_day_metadata", date="2026-06-01", entry="08:00", exit="17:00")
    _call("set_day_item", date="2026-06-01", description="A")
    _call("set_day_item", date="2026-06-01", description="B", hours=3)
    state = load_state()
    item_b_id = next(it.id for it in state.days["2026-06-01"].items if it.description == "B")
    out = _call("move_item", from_date="2026-06-01", to_date="2026-06-02", item_id=item_b_id)
    _check("09.1 moved.description = B", out["moved"]["description"] == "B")
    _check("09.2 moved.hours = 3", out["moved"]["hours"] == 3)
    s2 = load_state()
    src = s2.days["2026-06-01"]
    _check("09.3 from_date conserva entry", src.entry.strftime("%H:%M") == "08:00")
    _check("09.4 from_date conserva exit", src.exit.strftime("%H:%M") == "17:00")
    _check("09.5 from_date tiene 1 item restante (A)", len(src.items) == 1)
    _check("09.6 to_date tiene B", any(it.id == item_b_id for it in s2.days["2026-06-02"].items))


def test_10_export_full_month_with_inserts() -> None:
    print("\n[10] export 2027-12 con 23 dias -> insert_rows=1, formulas parcheadas")
    _reset()
    for iso in [
        "2027-12-01", "2027-12-02", "2027-12-03", "2027-12-06", "2027-12-07",
        "2027-12-08", "2027-12-09", "2027-12-10", "2027-12-13", "2027-12-14",
        "2027-12-15", "2027-12-16", "2027-12-17", "2027-12-20", "2027-12-21",
        "2027-12-22", "2027-12-23", "2027-12-24", "2027-12-27", "2027-12-28",
        "2027-12-29", "2027-12-30", "2027-12-31",
    ]:
        _call("set_day_item", date=iso, description="T")
    out = _call("export_to_excel", year=2027, month=12,
                path=str(EXPORTS_DIR / "t10.xlsx"))
    # n=23 -> inserted=1 (23 - 22). Totales en n+15 = fila 38.
    _check("10.1 inserted_rows = 1 (n=23 - 22)",
           out["inserted_rows"] == 1, str(out))
    _check("10.2 warning presente", out["warning"] is not None)
    wb = load_workbook(out["path"], data_only=False)
    ws = wb["HOJA LABOR REALIZADA"]
    ws1 = wb["FACT-SERV PROF"]
    f38 = ws.cell(row=38, column=6).value
    _check("10.3 F38 (totales en n+15) = SUM(F10:F32)",
           f38 == "=SUM(F10:F32)", repr(f38))
    _check("10.4 D12 sheet1 parcheado a F38",
           ws1["D12"].value == "='HOJA LABOR REALIZADA'!F38")
    _check("10.5 A20 sheet1 parcheado (A44:B44)", "A44:B44" in (ws1["A20"].value or ""))
    _check("10.6 A26 sheet1 parcheado (A48:B48)", "A48:B48" in (ws1["A26"].value or ""))
    _check("10.7 days_with_data = 23", out["days_with_data"] == 23)
    _check("10.8 days_in_month = 31", out["days_in_month"] == 31)
    _check("10.9 last_day = 2027-12-31", out["last_day"] == "2027-12-31")


# ---------------------------------------------------------------------------
# Tests extras de borde
# ---------------------------------------------------------------------------

def test_11_hours_over_8() -> None:
    print("\n[11] set_day_item con hours=11 -> acepta y persiste")
    _reset()
    out = _call("set_day_item", date="2026-06-01", description="A", hours=11)
    _check("11.1 hours=11 aceptado", out["item"]["hours"] == 11)
    _check("11.2 total=11", out["items_total"] == 11)


def test_12_existing_desc_hours_explicit_updates_first() -> None:
    print("\n[12] desc duplicada + hours -> update del PRIMER match")
    _reset()
    _call("set_day_items", date="2026-06-01", items=[
        {"description": "A", "hours": 1},
        {"description": "A", "hours": 2},
    ])
    out = _call("set_day_item", date="2026-06-01", description="A", hours=9)
    _check("12.1 primer item ahora 9", out["item"]["hours"] == 9)
    s = load_state()
    a_items = [it for it in s.days["2026-06-01"].items if it.description == "A"]
    _check("12.2 sigue habiendo 2 items A", len(a_items) == 2)
    _check("12.3 primero=9, segundo=2", a_items[0].hours == 9 and a_items[1].hours == 2)
    _check("12.4 total=11", out["items_total"] == 11)


def test_13_default_hours_zero_when_sum_ge_8() -> None:
    print("\n[13] set_day_item sin hours en dia con suma>=8 -> nuevo item hours=0")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A", hours=8)
    out = _call("set_day_item", date="2026-06-01", description="B")
    _check("13.1 B=0", out["item"]["hours"] == 0)
    _check("13.2 total=8", out["items_total"] == 8)


def test_14_delete_day_nonexistent() -> None:
    print("\n[14] delete_day sobre dia inexistente -> warning")
    _reset()
    out = _call("delete_day", date="2026-06-01")
    _check("14.1 warning dia inexistente", "inexistente" in (out.get("warning") or ""))
    _check("14.2 isError=False", out["isError"] is False)


def test_15_migrate_on_load_assigns_ids() -> None:
    print("\n[15] migracion on-load: state.json con items sin id -> UUID4 al cargar")
    _reset()
    payload = default_state().model_dump(mode="json")
    payload["days"] = {
        "2026-06-01": {
            "entry": None,
            "exit": None,
            "items": [
                {"id": "", "description": "Viejo", "hours": 5},
                {"description": "Sin id", "hours": 3},
            ],
        }
    }
    STATE_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    s = load_state()
    items = s.days["2026-06-01"].items
    uuid.UUID(items[0].id)
    uuid.UUID(items[1].id)
    _check("15.1 dos items con id valido", all(it.id for it in items))
    raw = json.loads(STATE_FILE.read_text(encoding="utf-8"))
    raw_items = raw["days"]["2026-06-01"]["items"]
    _check("15.2 archivo reescrito con ids", all(i["id"] for i in raw_items))


def test_15b_migrate_drops_legacy_year_month() -> None:
    print("\n[15b] migracion: state.json con year/month legacy se carga y reescribe sin ellos")
    _reset()
    payload = default_state().model_dump(mode="json")
    payload["year"] = 2026
    payload["month"] = 8
    payload["days"] = {
        "2026-08-15": {
            "entry": "08:00",
            "exit": "17:00",
            "items": [{"id": str(uuid.uuid4()), "description": "X", "hours": 8}],
        }
    }
    STATE_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    s = load_state()
    _check("15b.1 year/month legacy no estan en State",
           not hasattr(s, "year") and not hasattr(s, "month"))
    _check("15b.2 dias preservados", "2026-08-15" in s.days)
    raw = json.loads(STATE_FILE.read_text(encoding="utf-8"))
    _check("15b.3 archivo reescrito sin year", "year" not in raw)
    _check("15b.4 archivo reescrito sin month", "month" not in raw)
    _check("15b.5 dias preservados en disco",
           "2026-08-15" in raw["days"])


def test_16_encoding_unicode_preserved() -> None:
    print("\n[16] encoding UTF-8: 'Raúl' se preserva sin escape")
    _reset()
    out = _call("set_supervisor", name="  Raúl D. Olivero  ", date="2026-06-30")
    _check("16.1 supervisor.name = 'Raúl D. Olivero'", out["supervisor"]["name"] == "Raúl D. Olivero")
    raw = STATE_FILE.read_text(encoding="utf-8")
    _check("16.2 archivo contiene 'Raúl' literal", "Raúl" in raw)
    _check("16.3 archivo NO contiene 'Ra\\u00fal'", "Ra\\u00fal" not in raw)


def test_17_delete_day_item_not_found() -> None:
    print("\n[17] delete_day_item sin match -> warning")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A")
    out = _call("delete_day_item", date="2026-06-01", description="NoExiste")
    _check("17.1 warning item no encontrado", "no encontrado" in (out.get("warning") or ""))
    s = load_state()
    _check("17.2 dia intacto", len(s.days["2026-06-01"].items) == 1)


def test_18_weekend_days_accepted() -> None:
    print("\n[18] sabado y domingo son dias validos (Lun-Dom)")
    _reset()
    # 2026-06-06 es sabado
    out = _call("set_day_item", date="2026-06-06", description="Sabado", hours=4)
    _check("18.1 sabado aceptado (no isError)", out.get("isError") is False, str(out))
    _check("18.2 sabado: item persistido con hours=4",
           out.get("item", {}).get("hours") == 4, str(out))
    # 2026-06-07 es domingo
    out2 = _call("set_day_item", date="2026-06-07", description="Domingo", hours=2)
    _check("18.3 domingo aceptado (no isError)", out2.get("isError") is False, str(out2))
    _check("18.4 domingo: item persistido con hours=2",
           out2.get("item", {}).get("hours") == 2, str(out2))
    s = load_state()
    _check("18.5 sabado guardado en state",
           any(it.description == "Sabado" for it in s.days["2026-06-06"].items))
    _check("18.6 domingo guardado en state",
           any(it.description == "Domingo" for it in s.days["2026-06-07"].items))
    out3 = _call("set_day_metadata", date="2026-06-06", entry="10:00", exit="14:00")
    _check("18.7 set_day_metadata en sabado aceptado",
           out3.get("isError") is False, str(out3))


def test_18b_weekend_in_per_day_totals() -> None:
    print("\n[18b] per_day(year=2026, month=6) incluye sabado y domingo con sus items")
    _reset()
    _call("set_day_item", date="2026-06-06", description="S", hours=4)
    _call("set_day_item", date="2026-06-07", description="D", hours=2)
    out = _call("calculate_hours", year=2026, month=6)
    iso_by_date = {e["date"]: e for e in out["per_day"]}
    _check("18b.1 sabado en per_day", "2026-06-06" in iso_by_date)
    _check("18b.2 domingo en per_day", "2026-06-07" in iso_by_date)
    _check("18b.3 sabado total=4", iso_by_date["2026-06-06"]["items_total"] == 4)
    _check("18b.4 domingo total=2", iso_by_date["2026-06-07"]["items_total"] == 2)
    _check("18b.5 per_day cubre 30 dias (junio completo)",
           len(out["per_day"]) == 30)
    _check("18b.6 month_total_hours = 6", out["month_total_hours"] == 6)


def test_19_get_timesheet_full() -> None:
    print("\n[19] get_timesheet devuelve state + computed (filtrado a mes activo)")
    _reset()
    _call("set_professional_info", hourly_rate=10)
    _call("set_day_item", date="2026-06-01", description="A", hours=8)
    out = _call("get_timesheet", year=2026, month=6)
    _check("19.1 tiene state", "state" in out)
    _check("19.2 tiene computed", "computed" in out)
    _check("19.3 month_total_hours = 8", out["computed"]["month_total_hours"] == 8)
    _check("19.4 month_total_amount = 80 (rate 10 * 8)",
           out["computed"]["month_total_amount"] == 80)
    _check("19.5 per_day cubre los 30 dias de junio 2026",
           len(out["computed"]["per_day"]) == 30)
    _check("19.6 active_month.year = 2026", out["active_month"]["year"] == 2026)
    _check("19.7 active_month.source = explicit",
           out["active_month"]["source"] == "explicit")


def test_19b_get_timesheet_defaults_to_today() -> None:
    print("\n[19b] get_timesheet sin year/month -> mes actual del sistema")
    _reset()
    _call("set_day_item", date="2026-09-15", description="Hoy", hours=4)
    out = _call("get_timesheet")
    _check("19b.1 active_month.source = today",
           out["active_month"]["source"] == "today")
    _check("19b.2 year = 2026", out["active_month"]["year"] == 2026)
    _check("19b.3 month = 9", out["active_month"]["month"] == 9)
    _check("19b.4 month_total_hours = 4", out["computed"]["month_total_hours"] == 4)


def test_20_set_day_items_replace() -> None:
    print("\n[20] set_day_items reemplaza lista conservando entry/exit")
    _reset()
    _call("set_day_metadata", date="2026-06-01", entry="08:00", exit="17:00")
    _call("set_day_item", date="2026-06-01", description="A", hours=8)
    out = _call("set_day_items", date="2026-06-01", items=[
        {"description": "X", "hours": 2},
        {"description": "Y", "hours": 3},
    ])
    _check("20.1 items_count = 2", out["items_count"] == 2)
    _check("20.2 items_total = 5", out["items_total"] == 5)
    s = load_state()
    d = s.days["2026-06-01"]
    _check("20.3 entry conservado", d.entry.strftime("%H:%M") == "08:00")
    _check("20.4 exit conservado", d.exit.strftime("%H:%M") == "17:00")
    _check("20.5 items son X e Y", [it.description for it in d.items] == ["X", "Y"])


def test_21_hhmm_normalization() -> None:
    print("\n[21] '8:00' se normaliza a '08:00' (cero-padding)")
    _reset()
    out = _call("set_day_metadata", date="2026-06-01", entry="8:00", exit="5:30")
    _check("21.1 entry normalizado", out["entry"] == "08:00", str(out))
    _check("21.2 exit normalizado", out["exit"] == "05:30", str(out))


def test_22_defaults_persisted_on_new_day() -> None:
    print("\n[22] set_day_item en dia nuevo persiste entry=08:00 + exit=entry+hours")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A", hours=5)
    s = load_state()
    d = s.days["2026-06-01"]
    _check("22.1 entry = 08:00", d.entry.strftime("%H:%M") == "08:00", str(d.entry))
    _check("22.2 exit = 13:00 (08:00 + 5h)",
           d.exit.strftime("%H:%M") == "13:00", str(d.exit))
    _call("set_day_item", date="2026-06-01", description="B", hours=3)
    s = load_state()
    d = s.days["2026-06-01"]
    _check("22.3 entry sigue siendo 08:00",
           d.entry.strftime("%H:%M") == "08:00", str(d.entry))
    _check("22.4 exit recalculado a 16:00 (08:00 + 8h, no estaba locked)",
           d.exit.strftime("%H:%M") == "16:00", str(d.exit))


def test_23_metadata_override_respected() -> None:
    print("\n[23] set_day_metadata explicito prevalece sobre defaults")
    _reset()
    _call("set_day_metadata", date="2026-06-01", entry="09:30")
    _call("set_day_item", date="2026-06-01", description="A", hours=4)
    s = load_state()
    d = s.days["2026-06-01"]
    _check("23.1 entry explicito se preserva (09:30)",
           d.entry.strftime("%H:%M") == "09:30", str(d.entry))
    _check("23.2 exit = 13:30 (09:30 + 4h)",
           d.exit.strftime("%H:%M") == "13:30", str(d.exit))
    _call("set_day_metadata", date="2026-06-01", exit="18:00")
    _call("set_day_item", date="2026-06-01", description="B", hours=2)
    s = load_state()
    d = s.days["2026-06-01"]
    _check("23.3 entry sigue siendo 09:30",
           d.entry.strftime("%H:%M") == "09:30", str(d.entry))
    _check("23.4 exit explicito 18:00 se preserva aunque se agreguen items",
           d.exit.strftime("%H:%M") == "18:00", str(d.exit))


def test_24_set_day_items_zero_hours_keeps_entry() -> None:
    print("\n[24] set_day_items con horas=0 -> exit = entry (sin diferencia)")
    _reset()
    _call("set_day_items", date="2026-06-01", items=[
        {"description": "A", "hours": 0},
    ])
    s = load_state()
    d = s.days["2026-06-01"]
    _check("24.1 entry = 08:00", d.entry.strftime("%H:%M") == "08:00", str(d.entry))
    _check("24.2 exit = 08:00 (entry + 0h)",
           d.exit.strftime("%H:%M") == "08:00", str(d.exit))


def test_25_renderer_applies_defaults_for_preexisting_day() -> None:
    print("\n[25] export con dia sin entry/exit guardados -> renderer aplica defaults")
    _reset()
    payload = default_state().model_dump(mode="json")
    payload["days"] = {
        "2026-06-01": {
            "entry": None,
            "exit": None,
            "items": [
                {"id": str(uuid.uuid4()), "description": "Trabajo", "hours": 7},
            ],
        }
    }
    STATE_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    out = _call("export_to_excel", year=2026, month=6,
                path=str(EXPORTS_DIR / "t25.xlsx"))
    wb = load_workbook(out["path"], data_only=False)
    ws = wb["HOJA LABOR REALIZADA"]
    d10 = ws.cell(row=10, column=4).value
    e10 = ws.cell(row=10, column=5).value
    _check("25.1 D10 (entry) = 08:00", d10.strftime("%H:%M") == "08:00", str(d10))
    _check("25.2 E10 (exit) = 15:00 (08:00 + 7h)",
           e10.strftime("%H:%M") == "15:00", str(e10))


def test_26_move_creates_to_date_without_defaults() -> None:
    print("\n[26] move_item a fecha nueva -> to_date sin defaults hasta nuevo set_day_item")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A", hours=4)
    state = load_state()
    item_id = state.days["2026-06-01"].items[0].id
    _call("move_item", from_date="2026-06-01", to_date="2026-06-02", item_id=item_id)
    s = load_state()
    src = s.days["2026-06-01"]
    dst = s.days["2026-06-02"]
    _check("26.1 src conserva entry default 08:00",
           src.entry.strftime("%H:%M") == "08:00", str(src.entry))
    _check("26.2 dst sin entry guardada (move no aplica defaults)",
           dst.entry is None, str(dst.entry))
    _check("26.3 dst sin exit guardada",
           dst.exit is None, str(dst.exit))
    _call("set_day_item", date="2026-06-02", description="C", hours=3)
    s = load_state()
    dst = s.days["2026-06-02"]
    _check("26.4 set_day_item posterior aplica defaults a dst",
           dst.entry.strftime("%H:%M") == "08:00"
           and dst.exit.strftime("%H:%M") == "15:00",
           f"entry={dst.entry} exit={dst.exit}")


# ---------------------------------------------------------------------------
# Tests nuevos del modelo multi-mes (v0.3)
# ---------------------------------------------------------------------------

def test_30_multi_month_persistence() -> None:
    print("\n[30] entradas en distintos meses conviven sin pisarse")
    _reset()
    _call("set_day_item", date="2026-06-01", description="Junio", hours=8)
    _call("set_day_item", date="2026-07-01", description="Julio", hours=8)
    _call("set_day_item", date="2026-08-15", description="Agosto", hours=8)
    s = load_state()
    _check("30.1 los 3 dias presentes",
           all(k in s.days for k in ["2026-06-01", "2026-07-01", "2026-08-15"]))
    _check("30.2 junio intacto",
           s.days["2026-06-01"].items[0].description == "Junio")
    _check("30.3 julio intacto",
           s.days["2026-07-01"].items[0].description == "Julio")
    _check("30.4 agosto intacto",
           s.days["2026-08-15"].items[0].description == "Agosto")


def test_31_get_timesheet_filters_by_month() -> None:
    print("\n[31] get_timesheet(year, month) filtra computed.per_day al mes pedido")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A", hours=8)
    _call("set_day_item", date="2026-07-15", description="B", hours=8)
    out = _call("get_timesheet", year=2026, month=6)
    _check("31.1 month_total_hours = 8 (solo junio)",
           out["computed"]["month_total_hours"] == 8)
    _check("31.2 per_day = 30 dias de junio",
           len(out["computed"]["per_day"]) == 30)
    # El state completo sigue teniendo los dos dias:
    _check("31.3 state.days tiene ambos dias",
           all(k in out["state"]["days"] for k in ["2026-06-01", "2026-07-15"]))
    _check("31.4 months_present = [(2026,6), (2026,7)]",
           [(m["year"], m["month"]) for m in out["computed"]["months_present"]]
           == [(2026, 6), (2026, 7)])


def test_32_list_months_tool() -> None:
    print("\n[32] list_months devuelve meses con datos + mes actual del sistema")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A")
    _call("set_day_item", date="2026-06-02", description="B")
    _call("set_day_item", date="2026-08-15", description="C")
    out = _call("list_months")
    _check("32.1 isError False", out["isError"] is False)
    _check("32.2 months_present = [(2026,6,2), (2026,8,1)]",
           [(m["year"], m["month"], m["day_count"]) for m in out["months_present"]]
           == [(2026, 6, 2), (2026, 8, 1)])
    _check("32.3 today_month es year/month del sistema",
           out["today_month"]["year"] == date.today().year
           and out["today_month"]["month"] == date.today().month)


def test_33_export_no_data_returns_warning() -> None:
    print("\n[33] export_to_excel de mes sin entradas -> warning, sin archivo")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A")  # junio tiene data
    out = _call("export_to_excel", year=2027, month=3,
                path=str(EXPORTS_DIR / "t33.xlsx"))
    _check("33.1 warning presente", out.get("warning") is not None)
    _check("33.2 isError False (es warning, no error)", out["isError"] is False)
    _check("33.3 year/month devueltos", out["year"] == 2027 and out["month"] == 3)
    _check("33.4 days_with_data = 0", out["days_with_data"] == 0)
    _check("33.5 days_in_month = 31 (marzo)", out["days_in_month"] == 31)
    _check("33.6 path = None", out["path"] is None)


def test_34_export_requires_year_month_explicit() -> None:
    print("\n[34] export_to_excel sin year/month usa mes actual del sistema")
    _reset()
    # Crear entrada en el mes actual del sistema
    today = date.today()
    iso = today.isoformat()
    _call("set_day_item", date=iso, description="Hoy", hours=8)
    out = _call("export_to_excel", path=str(EXPORTS_DIR / "t34.xlsx"))
    _check("34.1 exporto el mes actual", out["year"] == today.year and out["month"] == today.month)
    _check("34.2 active_month_source = today", out["active_month_source"] == "today")


def test_35_calculate_hours_explicit_month() -> None:
    print("\n[35] calculate_hours(year, month) filtra al mes pedido")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A", hours=8)
    _call("set_day_item", date="2026-06-02", description="B", hours=4)
    _call("set_day_item", date="2026-08-15", description="C", hours=8)
    out = _call("calculate_hours", year=2026, month=6)
    _check("35.1 month_total_hours = 12 (junio)", out["month_total_hours"] == 12)
    out2 = _call("calculate_hours", year=2026, month=8)
    _check("35.2 month_total_hours = 8 (agosto)", out2["month_total_hours"] == 8)
    out3 = _call("calculate_hours", year=2026, month=7)
    _check("35.3 month_total_hours = 0 (julio sin data)",
           out3["month_total_hours"] == 0)


def test_36_partial_year_month_args_rejected() -> None:
    print("\n[36] year sin month (o viceversa) -> error de validacion")
    _reset()
    out1 = _call("get_timesheet", year=2026)
    _check("36.1 year solo -> isError True", out1["isError"] is True)
    out2 = _call("get_timesheet", month=6)
    _check("36.2 month solo -> isError True", out2["isError"] is True)
    out3 = _call("export_to_excel", year=2026)
    _check("36.3 export con year solo -> isError True", out3["isError"] is True)


def test_37_set_professional_info_ignores_year_month() -> None:
    print("\n[37] set_professional_info ignora year/month (no pueden borrar data)")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A")
    _call("set_day_item", date="2026-08-15", description="B")
    out = _call("set_professional_info", name="Otro", year=2027, month=1)
    _check("37.1 sin error", out.get("isError") is False, str(out))
    _check("37.2 name actualizado",
           out["professional"]["name"] == "Otro")
    s = load_state()
    _check("37.3 junio intacto", "2026-06-01" in s.days)
    _check("37.4 agosto intacto", "2026-08-15" in s.days)


# ---------------------------------------------------------------------------
# Tests del comportamiento "solo dias con entradas" en el export (v0.3+)
# ---------------------------------------------------------------------------

def test_38_export_skips_days_without_entries() -> None:
    print("\n[38] export junio con 5 entradas dispersas -> solo esas 5 filas con datos")
    _reset()
    # Fechas NO consecutivas: 02, 07, 14, 21, 28 (saltando fines de semana)
    for iso in ["2026-06-02", "2026-06-07", "2026-06-14", "2026-06-21", "2026-06-28"]:
        _call("set_day_item", date=iso, description="Trabajo", hours=8)
    out = _call("export_to_excel", year=2026, month=6,
                path=str(EXPORTS_DIR / "t38.xlsx"))
    _check("38.1 days_with_data = 5", out["days_with_data"] == 5)
    _check("38.2 days_in_month = 30", out["days_in_month"] == 30)
    wb = load_workbook(out["path"], data_only=False)
    ws = wb["HOJA LABOR REALIZADA"]
    # Filas 10..14 deben tener las 5 fechas en orden cronologico.
    # openpyxl lee `date` como `datetime`; normalizamos a date para comparar.
    row_dates = [ws.cell(row=r, column=1).value.date()
                 if ws.cell(row=r, column=1).value else None
                 for r in range(10, 15)]
    expected = [
        __import__("datetime").date(2026, 6, 2),
        __import__("datetime").date(2026, 6, 7),
        __import__("datetime").date(2026, 6, 14),
        __import__("datetime").date(2026, 6, 21),
        __import__("datetime").date(2026, 6, 28),
    ]
    _check("38.3 filas 10..14 tienen las fechas no consecutivas correctas",
           row_dates == expected, f"got {row_dates}")
    # Las 5 filas de gap entre los datos y los totales (filas 15..19 con n=5)
    # deben estar vacias en columna A. El resto (totales/firma/supervisor)
    # intencionalmente tiene contenido.
    gap_rows = list(range(10 + 5, 10 + 5 + 5))  # 15..19
    non_empty_gap = [r for r in gap_rows
                     if ws.cell(row=r, column=1).value is not None]
    _check("38.4 filas 15..19 (gap) sin fecha en col A",
           non_empty_gap == [],
           f"non-empty gap rows: {non_empty_gap}")
    # Totales en n+15 = 20
    f20 = ws.cell(row=20, column=6).value
    _check("38.5 F20 = SUM(F10:F14)", f20 == "=SUM(F10:F14)", repr(f20))


def test_39_export_no_items_returns_no_file() -> None:
    print("\n[39] export de mes sin items -> warning, sin archivo, path=None")
    _reset()
    out = _call("export_to_excel", year=2026, month=11,
                path=str(EXPORTS_DIR / "t39.xlsx"))
    _check("39.1 path = None", out["path"] is None)
    _check("39.2 warning presente", out.get("warning") is not None)
    _check("39.3 days_with_data = 0", out["days_with_data"] == 0)
    _check("39.4 days_in_month = 30 (noviembre)", out["days_in_month"] == 30)
    _check("39.5 archivo no creado", not Path(str(EXPORTS_DIR / "t39.xlsx")).exists())


def test_40_export_many_days_triggers_insert() -> None:
    print("\n[40] export junio con 25 dias -> inserted_rows = 3")
    _reset()
    for day in range(1, 26):  # 2026-06-01 .. 2026-06-25
        iso = f"2026-06-{day:02d}"
        _call("set_day_item", date=iso, description="T", hours=8)
    out = _call("export_to_excel", year=2026, month=6,
                path=str(EXPORTS_DIR / "t40.xlsx"))
    # n=25 -> inserted=3 (25 - 22)
    _check("40.1 inserted_rows = 3", out["inserted_rows"] == 3, str(out))
    _check("40.2 days_with_data = 25", out["days_with_data"] == 25)
    wb = load_workbook(out["path"], data_only=False)
    ws = wb["HOJA LABOR REALIZADA"]
    # Totales en n+15 = 40
    f40 = ws.cell(row=40, column=6).value
    _check("40.3 F40 = SUM(F10:F34)", f40 == "=SUM(F10:F34)", repr(f40))
    # Ultima fila de datos: 10 + 25 - 1 = 34 (con 3 inserts, las 25 filas
    # de datos van de 10 a 34). La celda A34 debe tener 2026-06-25.
    last_data_a = ws.cell(row=34, column=1).value
    last_data = last_data_a.date() if last_data_a else None
    _check("40.4 Fila 34 (ultimo dia con data) = 2026-06-25",
           last_data == __import__("datetime").date(2026, 6, 25),
           repr(last_data))


def test_41_last_day_is_last_with_data() -> None:
    print("\n[41] last_day en el export = ultimo dia con entradas, no fin de mes")
    _reset()
    # El mes tiene 30 dias pero solo registramos hasta el dia 10
    for day in [3, 5, 7, 10]:
        iso = f"2026-06-{day:02d}"
        _call("set_day_item", date=iso, description="T", hours=8)
    out = _call("export_to_excel", year=2026, month=6,
                path=str(EXPORTS_DIR / "t41.xlsx"))
    _check("41.1 last_day en respuesta = 2026-06-10",
           out["last_day"] == "2026-06-10", str(out))
    wb = load_workbook(out["path"], data_only=False)
    ws = wb["HOJA LABOR REALIZADA"]
    # n=4 -> inserted=0. last_day en n+21 = 25.
    last_day_cell = ws.cell(row=25, column=5).value
    actual = last_day_cell.date() if last_day_cell else None
    _check("41.2 E25 (ultimo dia en Excel) = 2026-06-10",
           actual == __import__("datetime").date(2026, 6, 10),
           repr(actual))


def test_42_blank_rows_in_template_are_truly_empty() -> None:
    print("\n[42] gap rows entre datos y totales vacias (template limpiado)")
    _reset()
    # 3 dias en junio (n=3): filas 10..12 son data, 13..17 son gap (5 filas),
    # 18 totales, 19 tarifa, 20 monto, 24 firma, 28 supervisor.
    for iso in ["2026-06-01", "2026-06-15", "2026-06-30"]:
        _call("set_day_item", date=iso, description="T", hours=8)
    out = _call("export_to_excel", year=2026, month=6,
                path=str(EXPORTS_DIR / "t42.xlsx"))
    wb = load_workbook(out["path"])
    ws = wb["HOJA LABOR REALIZADA"]
    # Las 5 filas de gap entre los datos y los totales (13..17 con n=3)
    # deben estar vacias en A..F.
    n = 3
    gap_rows = list(range(10 + n, 10 + n + 5))
    cols_with_content = []
    for r in gap_rows:
        for c in range(1, 7):
            if ws.cell(row=r, column=c).value is not None:
                cols_with_content.append((r, c, ws.cell(row=r, column=c).value))
    _check("42.1 gap rows (13..17) vacias en A..F",
           cols_with_content == [],
           f"non-empty cells: {cols_with_content[:5]}")
    # Filas de data si tienen contenido
    for r in range(10, 10 + n):
        _check(f"42.2 fila {r} (data) tiene descripcion en B",
               ws.cell(row=r, column=2).value is not None,
               repr(ws.cell(row=r, column=2).value))


ALL_TESTS = [
    test_01_export_clean_template,
    test_02_set_day_item_default_hours,
    test_03_set_day_item_explicit_hours,
    test_04_set_day_item_update_first,
    test_05_export_with_five_days,
    test_06_set_professional_info_does_not_wipe,
    test_07_export_reload_image_merges,
    test_08_existing_desc_no_hours_is_noop,
    test_09_move_item_by_id,
    test_10_export_full_month_with_inserts,
    test_11_hours_over_8,
    test_12_existing_desc_hours_explicit_updates_first,
    test_13_default_hours_zero_when_sum_ge_8,
    test_14_delete_day_nonexistent,
    test_15_migrate_on_load_assigns_ids,
    test_15b_migrate_drops_legacy_year_month,
    test_16_encoding_unicode_preserved,
    test_17_delete_day_item_not_found,
    test_18_weekend_days_accepted,
    test_18b_weekend_in_per_day_totals,
    test_19_get_timesheet_full,
    test_19b_get_timesheet_defaults_to_today,
    test_20_set_day_items_replace,
    test_21_hhmm_normalization,
    test_22_defaults_persisted_on_new_day,
    test_23_metadata_override_respected,
    test_24_set_day_items_zero_hours_keeps_entry,
    test_25_renderer_applies_defaults_for_preexisting_day,
    test_26_move_creates_to_date_without_defaults,
    test_30_multi_month_persistence,
    test_31_get_timesheet_filters_by_month,
    test_32_list_months_tool,
    test_33_export_no_data_returns_warning,
    test_34_export_requires_year_month_explicit,
    test_35_calculate_hours_explicit_month,
    test_36_partial_year_month_args_rejected,
    test_37_set_professional_info_ignores_year_month,
    test_38_export_skips_days_without_entries,
    test_39_export_no_items_returns_no_file,
    test_40_export_many_days_triggers_insert,
    test_41_last_day_is_last_with_data,
    test_42_blank_rows_in_template_are_truly_empty,
]


def main() -> int:
    if not TEMPLATE_FILE.exists():
        print(f"FATAL: plantilla no encontrada en {TEMPLATE_FILE}")
        return 2

    backup: Path | None = None
    if STATE_FILE.exists():
        backup = STATE_FILE.with_suffix(".json.bak")
        shutil.copy2(STATE_FILE, backup)
    try:
        for t in ALL_TESTS:
            try:
                t()
            except TestFailed as exc:
                print(f"  FAILED: {exc}")
        print("\n" + "=" * 60)
        passed = sum(1 for _, ok, _ in _RESULTS if ok)
        total = len(_RESULTS)
        print(f"Resultado: {passed}/{total} checks OK")
        failed = [(n, d) for n, ok, d in _RESULTS if not ok]
        if failed:
            print("FALLAS:")
            for n, d in failed:
                print(f"  - {n}: {d}")
            return 1
        return 0
    finally:
        if backup and backup.exists():
            shutil.copy2(backup, STATE_FILE)
            backup.unlink()
        else:
            if STATE_FILE.exists():
                STATE_FILE.unlink()


if __name__ == "__main__":
    sys.exit(main())

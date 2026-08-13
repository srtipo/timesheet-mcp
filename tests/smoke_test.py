"""Smoke test: ejecuta los 10 casos del spec §10 + extras de borde.

Uso: python tests/smoke_test.py
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
    print("\n[01] export_to_excel con estado vacio")
    _reset()
    out = _call("export_to_excel", path=str(EXPORTS_DIR / "t01.xlsx"))
    _check("01.1 archivo existe", Path(out["path"]).exists())
    _check("01.2 isError False", out.get("isError") is False)
    _check("01.3 month_total_hours = 0", out["month_total_hours"] == 0)
    _check("01.4 inserted_rows = 0", out["inserted_rows"] == 0)
    with zipfile.ZipFile(out["path"]) as zf:
        has_media = any(n.startswith("xl/media/") for n in zf.namelist())
    _check("01.5 imagen preservada en el zip (con o sin plan B)", has_media)


def test_02_set_day_item_default_hours() -> None:
    print("\n[02] set_day_item('2026-06-01', 'Tarea A') -> hours=8 con id UUID4")
    _reset()
    out = _call("set_day_item", date="2026-06-01", description="Tarea A")
    _check("02.1 sin warning", "warning" not in out)
    _check("02.2 hours = 8", out["item"]["hours"] == 8, str(out["item"]))
    uuid.UUID(out["item"]["id"])  # raises if invalid
    _check("02.3 id es UUID4", True)
    _check("02.4 items_total = 8", out["items_total"] == 8)


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
    print("\n[05] export_to_excel con 5 dias cargados, F37 esperado = 40")
    _reset()
    for iso, hrs in [
        ("2026-06-01", 8),
        ("2026-06-02", 8),
        ("2026-06-03", 8),
        ("2026-06-04", 8),
        ("2026-06-05", 8),
    ]:
        _call("set_day_item", date=iso, description="Trabajo", hours=hrs)
    out = _call("export_to_excel", path=str(EXPORTS_DIR / "t05.xlsx"))
    wb = load_workbook(out["path"], data_only=False)
    ws = wb["HOJA LABOR REALIZADA"]
    f37 = ws.cell(row=37, column=6).value
    _check("05.1 F37 = =SUM(F10:F31)", f37 == "=SUM(F10:F31)", repr(f37))
    f10_to_f14 = [ws.cell(row=r, column=6).value for r in range(10, 15)]
    _check("05.2 F10..F14 = 8", all(v == 8 for v in f10_to_f14), str(f10_to_f14))


def test_06_change_month_discards_days() -> None:
    print("\n[06] set_professional_info(year=2026, month=7) -> junio descartado")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A")
    _call("set_day_item", date="2026-06-02", description="B")
    out = _call("set_professional_info", year=2026, month=7)
    _check("06.1 discarded_days contiene 2026-06-01", "2026-06-01" in out["discarded_days"])
    _check("06.2 discarded_days contiene 2026-06-02", "2026-06-02" in out["discarded_days"])
    _check("06.3 month = 7", out["month"] == 7)
    s = load_state()
    _check("06.4 days vacio", s.days == {})


def test_07_export_reload_image_merges() -> None:
    print("\n[07] re-abrir .xlsx: imagen, merges y formulas preservadas")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A")
    out = _call("export_to_excel", path=str(EXPORTS_DIR / "t07.xlsx"))
    wb = load_workbook(out["path"])
    ws = wb["HOJA LABOR REALIZADA"]
    merges = [str(m) for m in ws.merged_cells.ranges]
    _check("07.1 merge A6:B6 presente", "A6:B6" in merges, str(merges[:6]))
    _check("07.2 merge B10:C10 presente", "B10:C10" in merges, str(merges))
    _check("07.3 merge A47:B47 presente", "A47:B47" in merges, str(merges))
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


def test_10_export_23_working_days() -> None:
    print("\n[10] export de 2027-12 (23 laborables) -> insert_rows > 0, formulas parcheadas")
    _reset()
    _call("set_professional_info", year=2027, month=12)
    for iso in [
        "2027-12-01", "2027-12-02", "2027-12-03", "2027-12-06", "2027-12-07",
        "2027-12-08", "2027-12-09", "2027-12-10", "2027-12-13", "2027-12-14",
        "2027-12-15", "2027-12-16", "2027-12-17", "2027-12-20", "2027-12-21",
        "2027-12-22", "2027-12-23", "2027-12-24", "2027-12-27", "2027-12-28",
        "2027-12-29", "2027-12-30", "2027-12-31",
    ]:
        _call("set_day_item", date=iso, description="T")
    out = _call("export_to_excel", path=str(EXPORTS_DIR / "t10.xlsx"))
    _check("10.1 inserted_rows = 1", out["inserted_rows"] == 1, str(out))
    _check("10.2 warning presente", out["warning"] is not None)
    wb = load_workbook(out["path"], data_only=False)
    ws = wb["HOJA LABOR REALIZADA"]
    ws1 = wb["FACT-SERV PROF"]
    f38 = ws.cell(row=38, column=6).value
    _check("10.3 F38 (totales) = SUM(F10:F32)", f38 == "=SUM(F10:F32)", repr(f38))
    _check("10.4 D12 sheet1 parcheado", ws1["D12"].value == "='HOJA LABOR REALIZADA'!F38")
    _check("10.5 A20 sheet1 parcheado (A44:B44)", "A44:B44" in (ws1["A20"].value or ""))
    _check("10.6 A26 sheet1 parcheado (A48:B48)", "A48:B48" in (ws1["A26"].value or ""))


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


def test_18_validation_rejects_weekend() -> None:
    print("\n[18] validacion: sabado/domingo se rechaza")
    _reset()
    out = _call("set_day_item", date="2026-06-06", description="Sabado")
    _check("18.1 sabado rechazado (isError)", out.get("isError") is True, str(out))
    _check("18.1b mensaje de error contiene 'no es laborable'",
           "no es laborable" in (out.get("error") or "").lower(), str(out))
    # 2026-06-07 es domingo
    out2 = _call("set_day_metadata", date="2026-06-07", entry="08:00")
    _check("18.2 domingo rechazado (isError)", out2.get("isError") is True, str(out2))
    _check("18.2b mensaje de error contiene 'no es laborable'",
           "no es laborable" in (out2.get("error") or "").lower(), str(out2))


def test_19_get_timesheet_full() -> None:
    print("\n[19] get_timesheet devuelve state + computed")
    _reset()
    _call("set_day_item", date="2026-06-01", description="A", hours=8)
    out = _call("get_timesheet")
    _check("19.1 tiene state", "state" in out)
    _check("19.2 tiene computed", "computed" in out)
    _check("19.3 month_total_hours = 8", out["computed"]["month_total_hours"] == 8)
    _check("19.4 month_total_amount = 28", out["computed"]["month_total_amount"] == 28)
    _check("19.5 per_day tiene 22 dias laborables", len(out["computed"]["per_day"]) == 22)


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


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------

ALL_TESTS = [
    test_01_export_clean_template,
    test_02_set_day_item_default_hours,
    test_03_set_day_item_explicit_hours,
    test_04_set_day_item_update_first,
    test_05_export_with_five_days,
    test_06_change_month_discards_days,
    test_07_export_reload_image_merges,
    test_08_existing_desc_no_hours_is_noop,
    test_09_move_item_by_id,
    test_10_export_23_working_days,
    test_11_hours_over_8,
    test_12_existing_desc_hours_explicit_updates_first,
    test_13_default_hours_zero_when_sum_ge_8,
    test_14_delete_day_nonexistent,
    test_15_migrate_on_load_assigns_ids,
    test_16_encoding_unicode_preserved,
    test_17_delete_day_item_not_found,
    test_18_validation_rejects_weekend,
    test_19_get_timesheet_full,
    test_20_set_day_items_replace,
    test_21_hhmm_normalization,
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
